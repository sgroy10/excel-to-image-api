from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
import subprocess
import tempfile
import os
from pathlib import Path
from pdf2image import convert_from_path
import io

app = FastAPI(
    title="Excel to Image API",
    description="Convert Excel BOM sheets to PNG images with full formatting",
    version="1.0.0"
)

# Allow CORS for Lovable and other frontends
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, replace with your Lovable domain
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
def health_check():
    """Health check endpoint"""
    return {"status": "ok", "service": "Excel to Image API"}


@app.post("/convert")
async def convert_excel_to_image(
    file: UploadFile = File(...),
    dpi: int = 200,
    page: int = 1
):
    """
    Convert Excel file to PNG image.
    
    - **file**: Excel file (.xlsx or .xls)
    - **dpi**: Image resolution (default: 200, use 300 for print quality)
    - **page**: Page number to convert (default: 1)
    
    Returns: PNG image
    """
    
    # Validate file type
    if not file.filename.endswith(('.xlsx', '.xls', '.xlsm')):
        raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls, .xlsm) are supported")
    
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            # Save uploaded file
            input_path = Path(temp_dir) / file.filename
            with open(input_path, "wb") as f:
                content = await file.read()
                f.write(content)
            
            # Convert Excel to PDF using LibreOffice
            result = subprocess.run([
                'libreoffice', '--headless', '--convert-to', 'pdf',
                '--outdir', temp_dir, str(input_path)
            ], capture_output=True, text=True, timeout=60)
            
            pdf_path = input_path.with_suffix('.pdf')
            if not pdf_path.exists():
                raise HTTPException(status_code=500, detail=f"PDF conversion failed: {result.stderr}")
            
            # Convert PDF to PNG
            images = convert_from_path(str(pdf_path), dpi=dpi)
            
            if page > len(images):
                raise HTTPException(status_code=400, detail=f"Page {page} not found. Document has {len(images)} page(s)")
            
            # Get requested page (1-indexed)
            img = images[page - 1]
            
            # Convert to bytes
            img_bytes = io.BytesIO()
            img.save(img_bytes, format='PNG', optimize=True)
            img_bytes.seek(0)
            
            return Response(
                content=img_bytes.getvalue(),
                media_type="image/png",
                headers={
                    "Content-Disposition": f"inline; filename={input_path.stem}_page{page}.png"
                }
            )
            
    except subprocess.TimeoutExpired:
        raise HTTPException(status_code=504, detail="Conversion timed out")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/convert-all")
async def convert_all_pages(
    file: UploadFile = File(...),
    dpi: int = 200
):
    """
    Convert all pages of Excel file to PNG images.
    Returns JSON with base64 encoded images.
    
    - **file**: Excel file (.xlsx or .xls)
    - **dpi**: Image resolution (default: 200)
    
    Returns: JSON with base64 images
    """
    import base64
    
    if not file.filename.endswith(('.xlsx', '.xls', '.xlsm')):
        raise HTTPException(status_code=400, detail="Only Excel files are supported")
    
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / file.filename
            with open(input_path, "wb") as f:
                content = await file.read()
                f.write(content)
            
            # Convert to PDF
            subprocess.run([
                'libreoffice', '--headless', '--convert-to', 'pdf',
                '--outdir', temp_dir, str(input_path)
            ], capture_output=True, timeout=60)
            
            pdf_path = input_path.with_suffix('.pdf')
            if not pdf_path.exists():
                raise HTTPException(status_code=500, detail="PDF conversion failed")
            
            # Convert all pages
            images = convert_from_path(str(pdf_path), dpi=dpi)
            
            result = []
            for i, img in enumerate(images, 1):
                img_bytes = io.BytesIO()
                img.save(img_bytes, format='PNG', optimize=True)
                img_bytes.seek(0)
                
                result.append({
                    "page": i,
                    "image": base64.b64encode(img_bytes.getvalue()).decode('utf-8')
                })
            
            return {
                "filename": file.filename,
                "total_pages": len(images),
                "images": result
            }
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
