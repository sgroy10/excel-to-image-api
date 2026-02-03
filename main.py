from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
import subprocess
import tempfile
import os
from pathlib import Path
import io
import logging
import traceback

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Excel to Image API",
    description="Convert Excel BOM sheets to PNG images with full formatting",
    version="2.0.0"
)

# Allow CORS for Lovable and other frontends
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
def health_check():
    """Health check endpoint"""
    return {"status": "ok", "service": "Excel to Image API", "version": "2.0.0"}


@app.get("/health")
def detailed_health():
    """Detailed health check"""
    # Check if LibreOffice is available
    try:
        result = subprocess.run(['libreoffice', '--version'], capture_output=True, text=True, timeout=10)
        lo_version = result.stdout.strip() if result.returncode == 0 else "Not found"
    except Exception as e:
        lo_version = f"Error: {str(e)}"
    
    return {
        "status": "ok",
        "libreoffice": lo_version,
        "service": "Excel to Image API"
    }


@app.post("/convert")
async def convert_excel_to_image(
    file: UploadFile = File(...),
    dpi: int = 150
):
    """
    Convert Excel file to PNG image.
    """
    logger.info(f"Received file: {file.filename}, size: {file.size if hasattr(file, 'size') else 'unknown'}")
    
    # Validate file type
    if not file.filename.lower().endswith(('.xlsx', '.xls', '.xlsm')):
        raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls, .xlsm) are supported")
    
    temp_dir = None
    try:
        # Create temp directory
        temp_dir = tempfile.mkdtemp()
        logger.info(f"Created temp dir: {temp_dir}")
        
        # Save uploaded file
        input_path = Path(temp_dir) / file.filename
        content = await file.read()
        
        if len(content) == 0:
            raise HTTPException(status_code=400, detail="Empty file received")
        
        with open(input_path, "wb") as f:
            f.write(content)
        logger.info(f"Saved file: {input_path}, size: {len(content)} bytes")
        
        # Convert Excel to PDF using LibreOffice
        logger.info("Starting LibreOffice conversion...")
        result = subprocess.run([
            'libreoffice', 
            '--headless', 
            '--nofirststartwizard',
            '--norestore',
            '--convert-to', 'pdf',
            '--outdir', temp_dir, 
            str(input_path)
        ], capture_output=True, text=True, timeout=120)
        
        logger.info(f"LibreOffice stdout: {result.stdout}")
        logger.info(f"LibreOffice stderr: {result.stderr}")
        logger.info(f"LibreOffice return code: {result.returncode}")
        
        pdf_path = input_path.with_suffix('.pdf')
        if not pdf_path.exists():
            # Try alternative PDF name
            possible_pdfs = list(Path(temp_dir).glob("*.pdf"))
            if possible_pdfs:
                pdf_path = possible_pdfs[0]
            else:
                logger.error(f"PDF not created. Dir contents: {list(Path(temp_dir).iterdir())}")
                raise HTTPException(
                    status_code=500, 
                    detail=f"PDF conversion failed. LibreOffice error: {result.stderr or result.stdout or 'Unknown error'}"
                )
        
        logger.info(f"PDF created: {pdf_path}")
        
        # Convert PDF to PNG using pdftoppm (from poppler-utils)
        logger.info("Converting PDF to PNG...")
        png_path = Path(temp_dir) / "output"
        
        result = subprocess.run([
            'pdftoppm',
            '-png',
            '-r', str(dpi),
            '-singlefile',
            str(pdf_path),
            str(png_path)
        ], capture_output=True, text=True, timeout=60)
        
        logger.info(f"pdftoppm return code: {result.returncode}")
        if result.stderr:
            logger.info(f"pdftoppm stderr: {result.stderr}")
        
        # Find the output PNG
        final_png = Path(temp_dir) / "output.png"
        if not final_png.exists():
            possible_pngs = list(Path(temp_dir).glob("*.png"))
            if possible_pngs:
                final_png = possible_pngs[0]
            else:
                logger.error(f"PNG not created. Dir contents: {list(Path(temp_dir).iterdir())}")
                raise HTTPException(status_code=500, detail="PNG conversion failed")
        
        logger.info(f"PNG created: {final_png}")
        
        # Read and return the PNG
        with open(final_png, 'rb') as f:
            img_bytes = f.read()
        
        logger.info(f"Returning PNG, size: {len(img_bytes)} bytes")
        
        return Response(
            content=img_bytes,
            media_type="image/png",
            headers={
                "Content-Disposition": f"inline; filename={input_path.stem}.png"
            }
        )
            
    except subprocess.TimeoutExpired:
        logger.error("Conversion timed out")
        raise HTTPException(status_code=504, detail="Conversion timed out")
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=f"Conversion error: {str(e)}")
    finally:
        # Cleanup temp directory
        if temp_dir and os.path.exists(temp_dir):
            try:
                import shutil
                shutil.rmtree(temp_dir)
                logger.info(f"Cleaned up temp dir: {temp_dir}")
            except Exception as e:
                logger.warning(f"Failed to cleanup temp dir: {e}")
