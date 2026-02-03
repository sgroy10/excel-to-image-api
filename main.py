from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
import subprocess
import tempfile
import os
from pathlib import Path
import logging
import traceback
import zipfile
import re

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Excel to Image API",
    description="Convert Excel BOM sheets to PNG images with full formatting",
    version="3.0.0"
)

# Allow CORS for Lovable and other frontends
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def analyze_xlsx_worksheets(xlsx_path: str) -> dict:
    """
    Analyze Excel file to find the best worksheet (the real BOM sheet with product images).
    
    Scoring logic:
    - Sheet with MORE images scores higher (product images = extra images beyond logo)
    - Sheet with BOM keywords scores higher
    - The combination helps identify the real BOM sheet
    """
    worksheet_scores = {}
    
    try:
        with zipfile.ZipFile(xlsx_path, 'r') as zf:
            file_list = zf.namelist()
            
            # Find worksheets
            worksheets = [f for f in file_list if f.startswith('xl/worksheets/sheet') and f.endswith('.xml')]
            worksheets.sort()
            
            # Count images per drawing
            drawing_image_count = {}
            for f in file_list:
                if 'drawings/_rels/drawing' in f and f.endswith('.rels'):
                    try:
                        rel_content = zf.read(f).decode('utf-8')
                        # Count image references
                        image_count = rel_content.count('/image')
                        # Extract drawing number
                        match = re.search(r'drawing(\d+)\.xml\.rels', f)
                        if match:
                            drawing_num = match.group(1)
                            drawing_image_count[drawing_num] = image_count
                    except:
                        pass
            
            logger.info(f"Drawing image counts: {drawing_image_count}")
            
            # Map worksheets to drawings
            worksheet_to_drawing = {}
            for f in file_list:
                if 'worksheets/_rels/sheet' in f and f.endswith('.rels'):
                    try:
                        rel_content = zf.read(f).decode('utf-8')
                        match = re.search(r'sheet(\d+)\.xml\.rels', f)
                        if match:
                            sheet_num = match.group(1)
                            drawing_match = re.search(r'drawing(\d+)\.xml', rel_content)
                            if drawing_match:
                                worksheet_to_drawing[sheet_num] = drawing_match.group(1)
                    except:
                        pass
            
            logger.info(f"Worksheet to drawing mapping: {worksheet_to_drawing}")
            
            # Read shared strings for keyword search
            shared_strings = ""
            if 'xl/sharedStrings.xml' in file_list:
                try:
                    shared_strings = zf.read('xl/sharedStrings.xml').decode('utf-8').upper()
                except:
                    pass
            
            # Score each worksheet
            for ws in worksheets:
                match = re.search(r'sheet(\d+)\.xml', ws)
                if not match:
                    continue
                sheet_num = match.group(1)
                score = 0
                
                try:
                    # Read worksheet content
                    ws_content = zf.read(ws).decode('utf-8').upper()
                    
                    # Check for images - MORE images = higher score
                    # BOM sheet with product images typically has 3-4 images (logo + 2 product images)
                    # Template sheets typically have 1-2 images (just logo)
                    if sheet_num in worksheet_to_drawing:
                        drawing_num = worksheet_to_drawing[sheet_num]
                        img_count = drawing_image_count.get(drawing_num, 0)
                        
                        # Scoring based on image count
                        if img_count >= 4:
                            score += 100  # Definitely has product images
                        elif img_count >= 3:
                            score += 75   # Likely has product images
                        elif img_count >= 2:
                            score += 40   # Has logo + maybe 1 product image
                        elif img_count >= 1:
                            score += 20   # Has at least logo
                        
                        logger.info(f"Sheet {sheet_num}: {img_count} images -> +{score} points")
                    
                    # Check for filled data (not just template)
                    # Count cell values - real BOM has more filled cells
                    cell_values = ws_content.count('<V>')
                    if cell_values > 100:
                        score += 20
                    elif cell_values > 50:
                        score += 10
                    
                    # Check for specific STYLE NO value (not just header)
                    # Real BOM has actual style number like "DLER00268"
                    if re.search(r'DLER\d+|DER\d+|RING\d+|EAR\d+', ws_content):
                        score += 30  # Has actual style number filled in
                        logger.info(f"Sheet {sheet_num}: Has style number -> +30 points")
                    
                except Exception as e:
                    logger.warning(f"Error analyzing sheet {sheet_num}: {e}")
                
                worksheet_scores[int(sheet_num)] = score
                logger.info(f"Sheet {sheet_num} TOTAL SCORE: {score}")
            
    except Exception as e:
        logger.error(f"Error analyzing xlsx: {e}")
        return {'scores': {1: 0}, 'best_sheet': 1}
    
    # Find best sheet (highest score)
    if worksheet_scores:
        best_sheet = max(worksheet_scores, key=worksheet_scores.get)
        best_score = worksheet_scores[best_sheet]
        logger.info(f"Selected sheet: {best_sheet} with score {best_score}")
    else:
        best_sheet = 1
    
    return {
        'scores': worksheet_scores,
        'best_sheet': best_sheet
    }


@app.get("/")
def health_check():
    """Health check endpoint"""
    return {"status": "ok", "service": "Excel to Image API", "version": "3.0.0"}


@app.get("/health")
def detailed_health():
    """Detailed health check"""
    try:
        result = subprocess.run(['libreoffice', '--version'], capture_output=True, text=True, timeout=10)
        lo_version = result.stdout.strip() if result.returncode == 0 else "Not found"
    except Exception as e:
        lo_version = f"Error: {str(e)}"
    
    return {
        "status": "ok",
        "libreoffice": lo_version,
        "service": "Excel to Image API",
        "version": "3.0.0"
    }


@app.post("/convert")
async def convert_excel_to_image(
    file: UploadFile = File(...),
    dpi: int = 150,
    page: int = 0
):
    """
    Convert Excel file to PNG image.
    Automatically detects the best worksheet (BOM sheet with product images).
    
    - **file**: Excel file (.xlsx, .xls, .xlsm)
    - **dpi**: Image resolution (default: 150)
    - **page**: Force specific page (0 = auto-detect best sheet)
    """
    logger.info(f"Received file: {file.filename}")
    
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
        
        # Analyze worksheets to find the best one (BOM sheet with images)
        analysis = {'best_sheet': 1}
        if file.filename.lower().endswith('.xlsx'):
            analysis = analyze_xlsx_worksheets(str(input_path))
            logger.info(f"Worksheet analysis result: best_sheet={analysis['best_sheet']}")
        
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
        
        logger.info(f"LibreOffice return code: {result.returncode}")
        
        pdf_path = input_path.with_suffix('.pdf')
        if not pdf_path.exists():
            possible_pdfs = list(Path(temp_dir).glob("*.pdf"))
            if possible_pdfs:
                pdf_path = possible_pdfs[0]
            else:
                logger.error(f"PDF not created. Dir contents: {list(Path(temp_dir).iterdir())}")
                raise HTTPException(
                    status_code=500, 
                    detail=f"PDF conversion failed: {result.stderr or result.stdout or 'Unknown error'}"
                )
        
        logger.info(f"PDF created: {pdf_path}")
        
        # Get PDF page count
        pdfinfo_result = subprocess.run(['pdfinfo', str(pdf_path)], capture_output=True, text=True, timeout=30)
        total_pages = 1
        for line in pdfinfo_result.stdout.split('\n'):
            if 'Pages:' in line:
                total_pages = int(line.split(':')[1].strip())
                break
        logger.info(f"PDF has {total_pages} pages")
        
        # Determine which page to convert
        if page > 0:
            # User specified a page
            target_page = min(page, total_pages)
        else:
            # Auto-detect: use the best sheet from analysis
            target_page = analysis.get('best_sheet', 1)
            if target_page > total_pages:
                target_page = 1
        
        logger.info(f"Converting page {target_page} of {total_pages}")
        
        # Convert PDF page to PNG
        png_path = Path(temp_dir) / "output"
        
        result = subprocess.run([
            'pdftoppm',
            '-png',
            '-r', str(dpi),
            '-f', str(target_page),
            '-l', str(target_page),
            '-singlefile',
            str(pdf_path),
            str(png_path)
        ], capture_output=True, text=True, timeout=60)
        
        logger.info(f"pdftoppm return code: {result.returncode}")
        
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


@app.post("/analyze")
async def analyze_excel(file: UploadFile = File(...)):
    """
    Analyze Excel file and return worksheet scores.
    Useful for debugging which sheet will be selected.
    """
    if not file.filename.lower().endswith('.xlsx'):
        raise HTTPException(status_code=400, detail="Only .xlsx files can be analyzed")
    
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        input_path = Path(temp_dir) / file.filename
        content = await file.read()
        
        with open(input_path, "wb") as f:
            f.write(content)
        
        analysis = analyze_xlsx_worksheets(str(input_path))
        
        return {
            "filename": file.filename,
            "worksheet_scores": analysis['scores'],
            "selected_sheet": analysis['best_sheet'],
            "message": f"Sheet {analysis['best_sheet']} will be converted (highest score: {analysis['scores'].get(analysis['best_sheet'], 0)})"
        }
        
    finally:
        if temp_dir and os.path.exists(temp_dir):
            import shutil
            shutil.rmtree(temp_dir)
