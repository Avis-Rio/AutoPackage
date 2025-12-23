import sys
import os
import shutil
import logging
from pathlib import Path
from datetime import datetime
from typing import List

# 将父目录下的 AutoPackage 目录添加到路径，以便导入现有模块
current_dir = Path(__file__).resolve().parent
parent_dir = current_dir.parent
source_dir = parent_dir / "AutoPackage"

# 优先添加源码目录到 sys.path
if str(source_dir) not in sys.path:
    sys.path.insert(0, str(source_dir))

# 确保当前目录 (web_server) 在 sys.path 的最前面，
# 这样 uvicorn 加载 "main:app" 时会优先找到当前文件的 main 模块，
# 而不是 AutoPackage/main.py
if str(current_dir) not in sys.path:
    sys.path.insert(0, str(current_dir))

# 同时也添加父目录，以防万一
if str(parent_dir) not in sys.path:
    sys.path.append(str(parent_dir))

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks, Form
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# 导入现有业务逻辑
try:
    from excel_reader import AllocationTableReader, DetailTableReader, BoxSettingReader
    from data_transformer import DataTransformer
    from template_writer import TemplateWriter
    from delivery_note_generator import DeliveryNoteGenerator
    from assortment_generator import AssortmentGenerator
    from store_detail_writer import StoreDetailWriter
    from box_label_generator import BoxLabelGenerator
    from config import FileConfig, DeliveryNoteConfig, AssortmentConfig, StoreDetailConfig, AllocationConfig, BoxLabelConfig
except ImportError as e:
    print(f"Error importing core modules: {e}")
    print(f"Current sys.path: {sys.path}")
    sys.exit(1)

# Database imports
from sqlalchemy.orm import Session
from database import SessionLocal, engine, get_db
import models
from fastapi import Depends

# Create tables
models.Base.metadata.create_all(bind=engine)

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = FastAPI(title="AutoPackage V2", description="配分表自动转换工具 Web 版")

# 允许跨域
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# 目录配置
UPLOAD_DIR = parent_dir / "temp_uploads"
OUTPUT_DIR = parent_dir / "temp_outputs"
TEMPLATES_DIR = parent_dir / "templates"
STORAGE_DIR = parent_dir / "storage" / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
TEMPLATES_DIR.mkdir(exist_ok=True)
STORAGE_DIR.mkdir(parents=True, exist_ok=True)

# 挂载静态文件 (前端)
app.mount("/static", StaticFiles(directory=str(current_dir / "static")), name="static")

def cleanup_file(path: Path):
    """后台任务：清理临时文件"""
    try:
        if path.exists():
            os.remove(path)
            logger.info(f"Cleaned up file: {path}")
    except Exception as e:
        logger.error(f"Error cleaning up {path}: {e}")

@app.get("/")
async def read_root():
    return FileResponse(str(current_dir / "static" / "index.html"))

# --- Template Management APIs ---

@app.get("/api/templates")
async def list_templates():
    """List all templates in the library"""
    templates = []
    for f in TEMPLATES_DIR.glob("*"):
        if f.is_file() and not f.name.startswith("~"):
            stat = f.stat()
            templates.append({
                "name": f.name,
                "size": stat.st_size,
                "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            })
    return {"templates": templates}

@app.post("/api/templates")
async def upload_template(file: UploadFile = File(...)):
    """Upload a new template to the library"""
    try:
        file_path = TEMPLATES_DIR / file.filename
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        return {"status": "success", "message": f"Template {file.filename} uploaded"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.delete("/api/templates/{filename}")
async def delete_template(filename: str):
    """Delete a template from the library"""
    import unicodedata
    
    # Normalize filename to NFC
    filename = unicodedata.normalize('NFC', filename)
    
    system_templates = [
        FileConfig.DEFAULT_TEMPLATE_NAME,
        AllocationConfig.TEMPLATE_NAME,
        DeliveryNoteConfig.TEMPLATE_NAME,
        AssortmentConfig.TEMPLATE_NAME,
        StoreDetailConfig.TEMPLATE_NAME,
        "③受渡伝票_模板（上传系统资料）.xlsx",
        "③受渡伝票_模板（上传系统资料） .xlsx",
        "template.xlsx", # Legacy
        "template.xls"   # Legacy
    ]
    
    # Normalize system templates
    system_templates = [unicodedata.normalize('NFC', t) for t in system_templates]
    
    if filename in system_templates:
        raise HTTPException(status_code=403, detail="Cannot delete system template")

    file_path = TEMPLATES_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Template not found")
    try:
        os.remove(file_path)
        return {"status": "success", "message": f"Template {filename} deleted"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# --- End Template Management APIs ---

# --- History Management APIs ---

@app.get("/api/history")
async def get_history(
    limit: int = 20, 
    offset: int = 0, 
    db: Session = Depends(get_db)
):
    """Get conversion history with pagination"""
    total = db.query(models.ConversionHistory).count()
    history = db.query(models.ConversionHistory)\
        .order_by(models.ConversionHistory.created_at.desc())\
        .limit(limit)\
        .offset(offset)\
        .all()
    return {"items": history, "total": total}

class DeleteBatchRequest(BaseModel):
    ids: List[int]

@app.post("/api/history/delete_batch")
async def delete_history_batch(request: DeleteBatchRequest, db: Session = Depends(get_db)):
    """Batch delete history records"""
    records = db.query(models.ConversionHistory).filter(models.ConversionHistory.id.in_(request.ids)).all()
    count = 0
    for record in records:
        # Delete output file
        if record.file_path and os.path.exists(record.file_path):
            try:
                os.remove(record.file_path)
            except Exception as e:
                logger.warning(f"Failed to delete file {record.file_path}: {e}")
        
        # Delete source file
        if record.source_file_path and os.path.exists(record.source_file_path):
            try:
                os.remove(record.source_file_path)
            except Exception as e:
                logger.warning(f"Failed to delete source file {record.source_file_path}: {e}")

        db.delete(record)
        count += 1
    
    db.commit()
    return {"status": "success", "message": f"Deleted {count} records"}

@app.get("/api/history/{history_id}/preview")
async def preview_history_file(history_id: int, db: Session = Depends(get_db)):
    """Preview the first 20 rows of the output file"""
    record = db.query(models.ConversionHistory).filter(models.ConversionHistory.id == history_id).first()
    if not record or not record.file_path or not os.path.exists(record.file_path):
        raise HTTPException(status_code=404, detail="File not found")
    
    if record.file_path.lower().endswith('.zip'):
         raise HTTPException(status_code=400, detail="Cannot preview ZIP files")

    try:
        import pandas as pd
        import numpy as np
        # Read first 20 rows
        # header=None means we read the first row as data, preventing pandas from making up "Unnamed: X" headers
        # if the file has complex headers.
        # But if the file HAS headers, they will be row 0.
        # Let's read with header=None first, then try to detect if row 0 looks like a header.
        df = pd.read_excel(record.file_path, nrows=20, header=None)
        
        # Replace NaN, inf, -inf with None for valid JSON
        df = df.replace([np.inf, -np.inf], np.nan)
        df = df.where(pd.notnull(df), None)
        
        # Convert to serializable list (handle remaining edge cases)
        data = []
        for row in df.values.tolist():
            cleaned_row = []
            for cell in row:
                if cell is None:
                    cleaned_row.append(None)
                elif isinstance(cell, float) and (np.isnan(cell) or np.isinf(cell)):
                    cleaned_row.append(None)
                else:
                    cleaned_row.append(cell)
            data.append(cleaned_row)
            
        # If we read with header=None, columns are 0, 1, 2...
        # We can just return the data, and let frontend assume first row might be header
        # OR we can manually set columns to "Column 1", "Column 2" etc.
        # But the user issue is "Unnamed: 0". 
        # By using header=None, we avoid pandas generating "Unnamed".
        # Instead, the actual header row (if any) becomes the first row of 'data'.
        # We can pass empty columns list or generic ones.
        
        return {
            "columns": [f"Col {i+1}" for i in range(len(df.columns))], 
            "data": data,
            "has_header": False # Hint to frontend that we treated it as no-header
        }
    except Exception as e:
        logger.error(f"Preview failed: {e}")
        raise HTTPException(status_code=500, detail=f"Error reading file: {str(e)}")

@app.delete("/api/history/{history_id}")
async def delete_history(history_id: int, db: Session = Depends(get_db)):
    """Delete a history record"""
    record = db.query(models.ConversionHistory).filter(models.ConversionHistory.id == history_id).first()
    if not record:
        raise HTTPException(status_code=404, detail="History record not found")
    
    # Optional: Delete the physical file if it exists
    # Note: We might want to keep it if other records point to it, but here 1 record = 1 file usually
    if record.file_path and os.path.exists(record.file_path):
        try:
            os.remove(record.file_path)
        except Exception as e:
            logger.warning(f"Failed to delete file {record.file_path}: {e}")
            
    db.delete(record)
    db.commit()
    return {"status": "success", "message": "Record deleted"}

    db.delete(record)
    db.commit()
    return {"status": "success", "message": "Record deleted"}

from pydantic import BaseModel

class NoteUpdate(BaseModel):
    note: str

@app.patch("/api/history/{history_id}")
async def update_history_note(history_id: int, note_update: NoteUpdate, db: Session = Depends(get_db)):
    """Update history note"""
    record = db.query(models.ConversionHistory).filter(models.ConversionHistory.id == history_id).first()
    if not record:
        raise HTTPException(status_code=404, detail="History record not found")
    
    record.note = note_update.note
    db.commit()
    return {"status": "success", "message": "Note updated"}

@app.post("/api/history/{history_id}/rerun")
async def rerun_conversion(
    history_id: int,
    mode: str = Form("allocation"),
    template_name: str = Form(None),
    week_num: str = Form(None),
    db: Session = Depends(get_db)
):
    """Rerun a conversion using stored source file"""
    # 1. Find original record
    original_record = db.query(models.ConversionHistory).filter(models.ConversionHistory.id == history_id).first()
    if not original_record:
        raise HTTPException(status_code=404, detail="Original record not found")
    
    if not original_record.source_file_path or not os.path.exists(original_record.source_file_path):
        raise HTTPException(status_code=400, detail="Source file not found (maybe expired or deleted)")

    # 2. Prepare fake UploadFile from stored file
    # We need to copy the stored file to a temp input path as if it was uploaded
    # But wait, convert_file expects UploadFile. 
    # Refactoring convert_file to accept path is better, but to minimize changes, 
    # we can just call the internal logic or simulate the flow.
    # Let's extract the core logic of convert_file into a separate function `process_conversion`
    # that takes file paths instead of UploadFile objects.
    
    # However, for now, let's just copy the file to UPLOAD_DIR and call a shared internal function.
    # Or simpler: Just re-implement the setup part and call the generator.
    # Actually, calling `convert_file` directly is hard because of UploadFile.
    # Let's refactor `convert_file` slightly.
    
    # Alternative: We can construct an UploadFile object? No, that's for incoming requests.
    # Let's just duplicate the setup logic for now to avoid breaking existing API signature too much,
    # or better, extract a `_run_conversion` function.
    
    return await _run_conversion(
        db=db,
        input_filename=original_record.original_filename,
        input_file_path=Path(original_record.source_file_path),
        mode=mode,
        template_name=template_name,
        week_num=week_num,
        # Note: Detail file is tricky. If original mode was allocation, we need detail file.
        # We didn't store detail file path in DB model yet! 
        # Wait, the user requirement didn't mention detail file persistence.
        # But for allocation mode, detail file is required.
        # If we want to rerun allocation, we need the detail file too.
        # For now, let's assume we ONLY support rerun if we have what we need.
        # If detail file is missing, we fail.
        # We should probably store detail file too if we want full rerun support.
        # Let's check if we can support it. 
        # For this iteration, let's just support simple rerun (Delivery Note / Assortment) or Allocation if we assume detail file is not needed or we can't do it yet.
        # Actually, let's just try to run. If it fails, it fails.
        # But wait, `convert_file` logic requires detail_file for allocation.
        # Let's skip detail_file for now in rerun unless we stored it.
        # We can store it in the same folder with a suffix?
        # Let's update `convert_file` to store detail file as well if present.
        detail_file_path=None 
    )

async def _run_conversion(
    db: Session,
    input_filename: str,
    input_file_path: Path,
    mode: str,
    template_name: str = None,
    week_num: str = None,
    detail_file_path: Path = None
):
    # This function mimics convert_file but takes paths
    
    # Copy input to temp (so we don't modify storage)
    temp_input_path = UPLOAD_DIR / f"rerun_{int(datetime.now().timestamp())}_{input_filename}"
    shutil.copy2(input_file_path, temp_input_path)
    
    # Setup Template
    template_path = None
    real_template_path = None
    
    if template_name:
        template_path = TEMPLATES_DIR / template_name
        if not template_path.exists():
             # Fallback to default search if not found? No, explicit request.
             raise HTTPException(status_code=404, detail=f"Template {template_name} not found")
        # Copy to temp
        temp_copy = UPLOAD_DIR / f"temp_lib_tpl_{int(datetime.now().timestamp())}_{template_name}"
        shutil.copy2(template_path, temp_copy)
        real_template_path = temp_copy
        template_path = temp_copy
    else:
        # Search default
        if mode == "delivery_note":
             candidates = [
                 source_dir / DeliveryNoteConfig.TEMPLATE_NAME,
                 parent_dir / DeliveryNoteConfig.TEMPLATE_NAME,
                 TEMPLATES_DIR / DeliveryNoteConfig.TEMPLATE_NAME,
                 source_dir / (DeliveryNoteConfig.TEMPLATE_NAME + "x"),
                 parent_dir / (DeliveryNoteConfig.TEMPLATE_NAME + "x"),
                 TEMPLATES_DIR / (DeliveryNoteConfig.TEMPLATE_NAME + "x")
             ]
        elif mode == "assortment":
             candidates = [
                 source_dir / AssortmentConfig.TEMPLATE_NAME,
                 parent_dir / AssortmentConfig.TEMPLATE_NAME,
                 TEMPLATES_DIR / AssortmentConfig.TEMPLATE_NAME
             ]
        else:
             candidates = [
                 source_dir / AllocationConfig.TEMPLATE_NAME,
                 parent_dir / AllocationConfig.TEMPLATE_NAME,
                 TEMPLATES_DIR / AllocationConfig.TEMPLATE_NAME,
                 source_dir / "template.xlsx",
                 parent_dir / "template.xlsx",
                 TEMPLATES_DIR / "template.xlsx"
             ]
        
        for path in candidates:
            if path.exists():
                real_template_path = path
                break
    
    if not real_template_path or not real_template_path.exists():
         raise HTTPException(status_code=400, detail=f"Default template for {mode} not found.")

    # Detail file setup
    jan_map = {}
    # If we had a stored detail file, we would use it here.
    # For now, if mode is allocation and we don't have detail file, we might fail or produce partial results.
    # Let's assume for now we only support modes that don't strictly require new detail file OR we rely on what's available.
    
    # Output setup
    prefix = "DeliveryNote_" if mode == "delivery_note" else "Converted_"
    output_filename = f"{prefix}{input_filename}"
    if not output_filename.endswith('.xlsx'):
         output_filename = os.path.splitext(output_filename)[0] + '.xlsx'
    output_path = OUTPUT_DIR / output_filename
    
    response_stats = {"items_processed": 0, "generated_file": ""}
    transformer_logs = []
    
    # Create DB record for RERUN
    db_record = models.ConversionHistory(
        original_filename=input_filename,
        mode=mode,
        status="processing",
        note="Rerun"
    )
    db.add(db_record)
    db.commit()
    db.refresh(db_record)
    
    try:
        # EXECUTE LOGIC (Copied/Shared from convert_file)
        # Ideally we should refactor the logic class instantiation out.
        
        logger.info(f"Starting RERUN process for {input_filename}, Mode: {mode}")
        
        if mode == "assortment":
            generator = AssortmentGenerator(str(temp_input_path), str(real_template_path), str(output_path), week_num=week_num)
            generator.process()
            items_processed = len(generator.data_rows)
            response_stats["items_processed"] = items_processed
            transformer_logs = list(getattr(generator, "logs", []) or [])
            
            # Stats
            try:
                rows = generator.data_rows or []
                store_count = len({r.get("delivery_code") for r in rows if r.get("delivery_code") is not None})
                box_count = len({r.get("slip_no") for r in rows if r.get("slip_no") is not None})
                sku_count = len({r.get("manufacturer_code") for r in rows if r.get("manufacturer_code") is not None})
                total_qty = sum(int(r.get("qty") or 0) for r in rows)
                response_stats.update({"store_count": store_count, "box_count": box_count, "sku_count": sku_count, "total_qty": total_qty})
            except: pass

            # Rename logic
            try:
                kanri_no = generator.kanri_no if hasattr(generator, 'kanri_no') else ""
                final_week = generator.week_num if hasattr(generator, 'week_num') else (week_num or "")
                if kanri_no and final_week:
                    new_filename = f"{final_week}-{kanri_no}-アソート明細.xlsx"
                    new_output_path = OUTPUT_DIR / new_filename
                    if output_path.exists():
                        if new_output_path.exists(): os.remove(new_output_path)
                        os.rename(output_path, new_output_path)
                        output_filename = new_filename
                        output_path = new_output_path
            except: pass

        elif mode == "delivery_note":
            # Delivery Note Logic
            # ... (Simplified for rerun: assume template is compatible or handled)
            # We need to handle .xls conversion if needed, same as main.
            # For brevity, assuming template is valid or we use the same logic.
            # Let's just instantiate Generator.
            
            # Check template format
            if str(real_template_path).lower().endswith('.xls'):
                 # Try to find xlsx or convert
                 xlsx_candidate = Path(str(real_template_path) + "x")
                 if xlsx_candidate.exists():
                     real_template_path = xlsx_candidate
                 else:
                     # Conversion logic (simplified: try pandas)
                     try:
                        import pandas as pd
                        temp_xlsx = UPLOAD_DIR / f"temp_template_rerun_{int(datetime.now().timestamp())}.xlsx"
                        df = pd.read_excel(real_template_path, sheet_name=None)
                        with pd.ExcelWriter(temp_xlsx, engine='openpyxl') as writer:
                            for sheet_name, data in df.items():
                                data.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                        real_template_path = temp_xlsx
                     except: pass

            generator = DeliveryNoteGenerator(str(temp_input_path), str(real_template_path), str(output_path))
            generator.process()
            items_processed = len(generator.data_rows)
            response_stats["items_processed"] = items_processed
            # Stats...
            try:
                rows = generator.data_rows or []
                store_count = len({r.get("store_code") for r in rows if r.get("store_code") is not None})
                box_count = len({r.get("slip_no") for r in rows if r.get("slip_no") is not None})
                total_qty = sum(int(r.get("qty") or 0) for r in rows)
                response_stats.update({"store_count": store_count, "box_count": box_count, "total_qty": total_qty})
            except: pass

        else:
            # Allocation
            if not detail_file_path:
                 # If we don't have detail file, we can't fully run allocation mode as designed.
                 # But maybe the user just wants to try? 
                 # Or maybe we stored detail file?
                 # For now, let's error if allocation mode is requested in rerun without detail file support.
                 raise Exception("Rerun for Allocation mode not fully supported yet (requires detail file persistence).")
            
            # ... Allocation logic ...

        # Success
        response_stats["generated_file"] = output_filename
        db_record.status = "success"
        db_record.output_filename = output_filename
        db_record.file_path = str(output_path)
        db_record.stats = response_stats
        db.commit()
        
        return {
            "status": "success",
            "message": "Rerun successful",
            "download_url": f"/api/download/{output_filename}",
            "stats": response_stats
        }

    except Exception as e:
        logger.error(f"Rerun failed: {e}", exc_info=True)
        db_record.status = "failed"
        db_record.error_message = str(e)
        db.commit()
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Cleanup temp input
        if temp_input_path.exists():
            try: os.remove(temp_input_path)
            except: pass
        # Cleanup temp template
        if template_path and template_path.exists() and UPLOAD_DIR in template_path.parents:
             try: os.remove(template_path)
             except: pass

@app.post("/api/generate-labels")
async def generate_box_labels(
    history_id: int = Form(...),
    db: Session = Depends(get_db)
):
    """
    根据历史记录生成箱贴PDF
    """
    # 1. 查找历史记录
    record = db.query(models.ConversionHistory).filter(
        models.ConversionHistory.id == history_id
    ).first()
    
    if not record:
        raise HTTPException(status_code=404, detail="Record not found")
    
    # 2. 检查源文件
    if not record.source_file_path or not os.path.exists(record.source_file_path):
        raise HTTPException(status_code=400, detail="Source file not found")
        
    try:
        # 3. 重新读取和转换数据
        # 假设源文件是配分表格式
        reader = AllocationTableReader(record.source_file_path)
        allocation_data = reader.read()
        
        transformer = DataTransformer(allocation_data)
        transform_result = transformer.transform()
        
        # 4. 生成PDF
        timestamp = int(datetime.now().timestamp())
        # Clean up original filename for PDF name
        safe_filename = "".join([c for c in record.original_filename if c.isalnum() or c in (' ', '.', '-', '_')]).strip()
        pdf_filename = f"BoxLabels_{timestamp}_{safe_filename}.pdf"
        
        # 移除可能的 .xlsx后缀
        if pdf_filename.endswith(".xlsx.pdf"):
            pdf_filename = pdf_filename.replace(".xlsx.pdf", ".pdf")
        elif pdf_filename.endswith(".xls.pdf"):
            pdf_filename = pdf_filename.replace(".xls.pdf", ".pdf")
            
        pdf_path = OUTPUT_DIR / pdf_filename
        
        generator = BoxLabelGenerator(transform_result, str(pdf_path))
        generator.generate()
        
        return {
            "status": "success",
            "download_url": f"/api/download/{pdf_filename}",
            "filename": pdf_filename
        }
        
    except Exception as e:
        logger.error(f"Failed to generate labels: {e}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Label generation failed: {str(e)}")


@app.post("/api/generate-labels-from-file")
async def generate_labels_from_file(
    file: UploadFile = File(...),
    db: Session = Depends(get_db)
):
    """
    上传工厂返回的箱设定文件，生成箱贴PDF
    """
    try:
        # 1. Save uploaded file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"factory_box_setting_{timestamp}.xlsx"
        file_path = UPLOAD_DIR / filename
        
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        logger.info(f"Received factory file: {file.filename}, saved to {file_path}")
        
        # 2. Parse file
        reader = BoxSettingReader(str(file_path))
        boxes = reader.read()
        
        if not boxes:
            raise HTTPException(status_code=400, detail="未在文件中找到有效的箱设定数据 (请检查是否包含 PT 页)")
            
        # 3. Generate PDF
        pdf_filename = f"BoxLabels_{timestamp}.pdf"
        pdf_path = OUTPUT_DIR / pdf_filename
        
        generator = BoxLabelGenerator(boxes, str(pdf_path))
        output_path, stats = generator.generate()
        
        # 4. Create History Record
        history_record = models.ConversionHistory(
            original_filename=file.filename,
            mode="box_label",
            status="success",
            output_filename=pdf_filename,
            file_path=str(pdf_path),
            source_file_path=str(file_path),
            stats=stats
        )
        db.add(history_record)
        db.commit()
        db.refresh(history_record)
        
        # 5. Return response
        return {
            "status": "success",
            "message": f"成功生成 {len(boxes)} 张箱贴",
            "download_url": f"/api/download/{pdf_filename}",
            "filename": pdf_filename,
            "stats": stats
        }
        
    except Exception as e:
        logger.error(f"Error generating labels: {e}")
        import traceback
        traceback.print_exc()
        # Save failed history if possible? 
        # Since we might not have 'file_path' if error happens early, skip for simplicity or handle better.
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        pass

# --- End History Management APIs ---

@app.post("/api/convert")
async def convert_file(
    file: UploadFile = File(...),
    template: UploadFile = File(None),
    detail_file: UploadFile = File(None),
    template_name: str = Form(None), # Optional: name of template in library
    mode: str = Form("allocation"), # allocation, delivery_note, assortment
    week_num: str = Form(None), # Optional: week number for assortment
    db: Session = Depends(get_db)
):
    """
    核心转换接口
    mode: 'allocation' (default) for 配分表转换
          'delivery_note' for 受渡伝票生成
    """
    input_path = UPLOAD_DIR / f"input_{int(datetime.now().timestamp())}_{file.filename}"
    
    # Validation for allocation mode
    if mode == "allocation" and not detail_file:
         raise HTTPException(status_code=400, detail="请上传明细表 (Detail File is required for allocation mode)")

    # Determine template filename and path
    template_filename = f"template_{int(datetime.now().timestamp())}"
    if template:
        ext = os.path.splitext(template.filename)[1]
        if not ext:
            ext = ".xlsx"
        template_filename += ext
        template_path = UPLOAD_DIR / template_filename
    elif template_name:
        # Use template from library
        template_path = TEMPLATES_DIR / template_name
        if not template_path.exists():
            raise HTTPException(status_code=404, detail=f"Template {template_name} not found in library")
    else:
        # Default template extension depends on mode
        if mode == "delivery_note":
             template_filename += ".xls" # Original template is .xls
        elif mode == "assortment":
             template_filename += ".xlsx" # Assortment template is .xlsx
        else:
             template_filename += ".xlsx"
        template_path = UPLOAD_DIR / template_filename # Placeholder for default
    
    # Output filename
    prefix = "DeliveryNote_" if mode == "delivery_note" else "Converted_"
    output_filename = f"{prefix}{file.filename}"
    if not output_filename.endswith('.xlsx'):
         output_filename = os.path.splitext(output_filename)[0] + '.xlsx'
         
    output_path = OUTPUT_DIR / output_filename
    detail_path = None
    transformer_logs = []
    response_stats = {
        "items_processed": 0,
        "generated_file": ""
    }

    # Save source file to storage
    source_storage_path = STORAGE_DIR / f"{int(datetime.now().timestamp())}_{file.filename}"
    with open(source_storage_path, "wb") as buffer:
        # Reset file cursor just in case
        await file.seek(0)
        shutil.copyfileobj(file.file, buffer)
    
    # Also save detail file if present
    detail_storage_path = None
    if detail_file:
        detail_storage_path = STORAGE_DIR / f"{int(datetime.now().timestamp())}_detail_{detail_file.filename}"
        with open(detail_storage_path, "wb") as buffer:
            await detail_file.seek(0)
            shutil.copyfileobj(detail_file.file, buffer)
            await detail_file.seek(0) # Reset for later use

    # Create DB record
    db_record = models.ConversionHistory(
        original_filename=file.filename,
        mode=mode,
        status="processing",
        source_file_path=str(source_storage_path)
    )
    db.add(db_record)
    db.commit()
    db.refresh(db_record)
    
    # Reset file cursor for processing
    await file.seek(0)

    try:
        # 1. 保存上传的文件
        logger.info(f"Receiving file: {file.filename}, Mode: {mode}")
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        # 2. 处理模板
        real_template_path = None
        
        if template:
            # User uploaded template
            logger.info(f"Using uploaded template: {template.filename}")
            with open(template_path, "wb") as buffer:
                shutil.copyfileobj(template.file, buffer)
            real_template_path = template_path
        elif template_name:
             # Use library template
             logger.info(f"Using library template: {template_name}")
             # We should probably copy it to temp to avoid modification if code modifies it
             # But current code might modify real_template_path variable, not file content in place (except conversion)
             # Wait, conversion modifies real_template_path if it converts .xls to .xlsx
             # So we MUST copy it to temp
             temp_copy = UPLOAD_DIR / f"temp_lib_tpl_{int(datetime.now().timestamp())}_{template_name}"
             shutil.copy2(template_path, temp_copy)
             real_template_path = temp_copy
             # Mark as temporary so we clean it up? 
             # The finally block cleans 'template_path'. 
             # Here 'template_path' was set to library path. We shouldn't delete library path.
             # So we set template_path to temp_copy for cleanup purposes.
             template_path = temp_copy
        else:
            # Search for default template
            if mode == "delivery_note":
                # Check for xls first as user specified .xls template
                candidates = [
                    source_dir / DeliveryNoteConfig.TEMPLATE_NAME,
                    parent_dir / DeliveryNoteConfig.TEMPLATE_NAME,
                    TEMPLATES_DIR / DeliveryNoteConfig.TEMPLATE_NAME,
                    # Fallback to xlsx if converted
                    source_dir / (DeliveryNoteConfig.TEMPLATE_NAME + "x"),
                    parent_dir / (DeliveryNoteConfig.TEMPLATE_NAME + "x"),
                    TEMPLATES_DIR / (DeliveryNoteConfig.TEMPLATE_NAME + "x")
                ]
            elif mode == "assortment":
                candidates = [
                    source_dir / AssortmentConfig.TEMPLATE_NAME,
                    parent_dir / AssortmentConfig.TEMPLATE_NAME,
                    TEMPLATES_DIR / AssortmentConfig.TEMPLATE_NAME
                ]
            else:
                # Allocation mode
                candidates = [
                    source_dir / AllocationConfig.TEMPLATE_NAME,
                    parent_dir / AllocationConfig.TEMPLATE_NAME,
                    TEMPLATES_DIR / AllocationConfig.TEMPLATE_NAME,
                    # Fallback legacy
                    source_dir / "template.xlsx",
                    parent_dir / "template.xlsx",
                    TEMPLATES_DIR / "template.xlsx"
                ]
            
            for path in candidates:
                if path.exists():
                    real_template_path = path
                    break
        
        if not real_template_path or not real_template_path.exists():
             msg = f"Default template for {mode} not found."
             if mode == "delivery_note":
                 msg += f" Expected: {DeliveryNoteConfig.TEMPLATE_NAME}"
             raise HTTPException(status_code=400, detail=msg)

        # 2.5 处理明细表
        jan_map = {}
        if detail_file:
            detail_path = UPLOAD_DIR / f"detail_{int(datetime.now().timestamp())}_{detail_file.filename}"
            logger.info(f"Using detail file: {detail_file.filename}")
            with open(detail_path, "wb") as buffer:
                shutil.copyfileobj(detail_file.file, buffer)
            try:
                jan_map = DetailTableReader.read_jan_map(str(detail_path))
                logger.info(f"Loaded {len(jan_map)} JAN entries from detail file")
            except Exception as e:
                logger.error(f"Failed to read detail file: {e}")
                pass

        # 3. 执行核心逻辑
        logger.info("Starting process...")
        
        items_processed = 0
        
        if mode == "assortment":
            generator = AssortmentGenerator(str(input_path), str(real_template_path), str(output_path), week_num=week_num)
            generator.process()
            items_processed = len(generator.data_rows)
            response_stats["items_processed"] = items_processed
            transformer_logs = list(getattr(generator, "logs", []) or [])

            try:
                rows = generator.data_rows or []
                store_count = len({r.get("delivery_code") for r in rows if r.get("delivery_code") is not None})
                box_count = len({r.get("slip_no") for r in rows if r.get("slip_no") is not None})
                sku_count = len({r.get("manufacturer_code") for r in rows if r.get("manufacturer_code") is not None})
                total_qty = sum(int(r.get("qty") or 0) for r in rows)
                response_stats.update(
                    {
                        "store_count": store_count,
                        "box_count": box_count,
                        "sku_count": sku_count,
                        "total_qty": total_qty
                    }
                )
                transformer_logs.append(f"汇总: 店铺 {store_count}, 箱数 {box_count}, SKU {sku_count}, 总枚数 {total_qty}")
            except Exception as e:
                logger.warning(f"Failed to compute assortment summary stats: {e}")
            
            # Rename output file: {WeekNum}-{KanriNo}-アソート明細.xlsx
            try:
                kanri_no = generator.kanri_no if hasattr(generator, 'kanri_no') else ""
                final_week = generator.week_num if hasattr(generator, 'week_num') else (week_num or "")
                
                if kanri_no and final_week:
                    new_filename = f"{final_week}-{kanri_no}-アソート明細.xlsx"
                    new_output_path = OUTPUT_DIR / new_filename
                    
                    if output_path.exists():
                        if new_output_path.exists():
                            os.remove(new_output_path)
                        os.rename(output_path, new_output_path)
                        output_filename = new_filename
                        output_path = new_output_path
            except Exception as e:
                logger.warning(f"Failed to rename output file: {e}")

        elif mode == "delivery_note":
            # Delivery Note Generation
            # Check template format for writing (must be xlsx)
            if str(real_template_path).lower().endswith('.xls'):
                # Try to convert on the fly or fail
                # Ideally, we should have a pre-converted .xlsx template
                # For this task, let's assume we can use a converter or just rename if it's actually valid
                # But openpyxl cannot read .xls. 
                # Let's check if there is a .xlsx version available if we picked .xls
                xlsx_candidate = Path(str(real_template_path) + "x")
                if xlsx_candidate.exists():
                    real_template_path = xlsx_candidate
                else:
                    # Attempt to convert using win32com (Windows) to preserve styles
                    # Fallback to pandas if failed
                    try:
                        import win32com.client as win32
                        import pythoncom
                        
                        pythoncom.CoInitialize()
                        excel = win32.gencache.EnsureDispatch('Excel.Application')
                        excel.Visible = False
                        excel.DisplayAlerts = False
                        
                        try:
                            abs_template_path = str(real_template_path.resolve())
                            wb = excel.Workbooks.Open(abs_template_path)
                            temp_xlsx = UPLOAD_DIR / f"temp_template_{int(datetime.now().timestamp())}.xlsx"
                            # 51 = xlOpenXMLWorkbook (xlsx)
                            wb.SaveAs(str(temp_xlsx.resolve()), FileFormat=51)
                            wb.Close()
                            real_template_path = temp_xlsx
                            logger.info(f"Successfully converted .xls to .xlsx using win32com: {real_template_path}")
                        except Exception as e:
                            logger.error(f"win32com conversion failed: {e}")
                            raise e
                        finally:
                            # quit is dangerous if user has other excel open, but usually okay in server
                            # better to just keep it running or careful
                            # For safety in this environment, we might just close workbook.
                            # excel.Quit() 
                            # If we don't quit, the process remains. Let's Quit.
                            excel.Quit()
                            
                    except Exception as e_win32:
                        logger.warning(f"win32com conversion failed/not available, falling back to pandas: {e_win32}")
                        try:
                            import pandas as pd
                            temp_xlsx = UPLOAD_DIR / f"temp_template_{int(datetime.now().timestamp())}.xlsx"
                            logger.warning(f"Converting .xls template to .xlsx (styles will be lost): {real_template_path}")
                            df = pd.read_excel(real_template_path, sheet_name=None)
                            with pd.ExcelWriter(temp_xlsx, engine='openpyxl') as writer:
                                for sheet_name, data in df.items():
                                    data.to_excel(writer, sheet_name=sheet_name, index=False, header=False) # header=False to avoid Unnamed
                            real_template_path = temp_xlsx
                        except ImportError:
                            pass
            
            generator = DeliveryNoteGenerator(str(input_path), str(real_template_path), str(output_path))
            generator.process()
            # items_processed is hard to get without return, but let's assume success
            items_processed = len(generator.data_rows)
            response_stats["items_processed"] = items_processed
            try:
                rows = generator.data_rows or []
                store_count = len({r.get("store_code") for r in rows if r.get("store_code") is not None})
                box_count = len({r.get("slip_no") for r in rows if r.get("slip_no") is not None})
                sku_count = len(
                    {
                        (r.get("product_code"), r.get("color"), r.get("size"))
                        for r in rows
                        if r.get("product_code") is not None
                    }
                )
                total_qty = sum(int(r.get("qty") or 0) for r in rows)
                response_stats.update(
                    {
                        "store_count": store_count,
                        "box_count": box_count,
                        "sku_count": sku_count,
                        "total_qty": total_qty
                    }
                )
                transformer_logs.append(f"汇总: 店铺 {store_count}, 箱数 {box_count}, SKU {sku_count}, 总枚数 {total_qty}")
            except Exception as e:
                logger.warning(f"Failed to compute delivery note summary stats: {e}")
            
        else:
            # Standard Allocation Table Conversion
            # Step A: Read
            reader = AllocationTableReader(str(input_path))
            allocation_data = reader.read()
            items_processed = len(allocation_data.get('products', []))
            logger.info(f"Read {items_processed} products from allocation table.")
            response_stats["items_processed"] = items_processed
    
            # Step B: Transform
            transformer = DataTransformer(allocation_data, jan_map)
            transform_result = transformer.transform()
            
            # 收集详细日志
            transformer_logs = getattr(transformer, 'logs', [])
            for log in transformer_logs:
                logger.info(f"[Transformer] {log}")

            try:
                pt_groups = transform_result.get("pt_groups", []) or []
                sku_count = len(transform_result.get("skus", []) or [])
                pt_count = len(pt_groups)
                store_count = sum(len(g.get("stores", []) or []) for g in pt_groups)
                box_count = store_count
                total_qty = 0
                for g in pt_groups:
                    for s in g.get("stores", []) or []:
                        total_qty += int(s.get("total_qty") or 0)

                response_stats.update(
                    {
                        "sku_count": sku_count,
                        "pt_count": pt_count,
                        "store_count": store_count,
                        "box_count": box_count,
                        "total_qty": total_qty,
                        "jan_map_count": int(getattr(transformer, "jan_map_count", 0) or 0),
                        "jan_match_success": int(getattr(transformer, "jan_match_success", 0) or 0),
                        "jan_match_fail": int(getattr(transformer, "jan_match_fail", 0) or 0)
                    }
                )

                transformer_logs.append(
                    f"汇总: 店铺 {store_count}, 箱数 {box_count}, SKU {sku_count}, PT {pt_count}, 总枚数 {total_qty}"
                )
            except Exception as e:
                logger.warning(f"Failed to compute summary stats: {e}")
            
            # Step C: Write
            writer = TemplateWriter(str(real_template_path), str(output_path))
            writer.write(transform_result)
            
            # --- NEW: Step D: Write ④ Store Detail ---
            try:
                # Find template for ④
                sd_template_name = StoreDetailConfig.TEMPLATE_NAME
                sd_template_path = None
                for base in [source_dir, parent_dir, TEMPLATES_DIR]:
                    if (base / sd_template_name).exists():
                        sd_template_path = base / sd_template_name
                        break
                
                sd_output_path = None
                if sd_template_path:
                    sd_filename = f"StoreDetail_{file.filename}"
                    if not sd_filename.endswith('.xlsx'):
                        sd_filename = os.path.splitext(sd_filename)[0] + '.xlsx'
                    sd_output_path = OUTPUT_DIR / sd_filename
                    
                    sd_writer = StoreDetailWriter(str(sd_template_path), str(sd_output_path))
                    sd_writer.write(transform_result)
                    logger.info(f"Generated Store Detail: {sd_output_path}")
                else:
                    logger.warning(f"Store Detail template not found: {sd_template_name}")
                    transformer_logs.append(f"Warning: Store Detail template not found: {sd_template_name}")

                # --- Zip if multiple files ---
                if sd_output_path and sd_output_path.exists():
                    # Create zip
                    zip_filename = f"Package_{int(datetime.now().timestamp())}.zip"
                    zip_path = OUTPUT_DIR / zip_filename
                    
                    import zipfile
                    with zipfile.ZipFile(zip_path, 'w') as zf:
                        # Add ① (using simple name)
                        zf.write(output_path, arcname=output_path.name)
                        # Add ④
                        zf.write(sd_output_path, arcname=sd_output_path.name)
                    
                    # Update response to point to zip
                    output_path = zip_path # For download
                    output_filename = zip_filename
            except Exception as e:
                logger.error(f"Error generating Store Detail or Zip: {e}", exc_info=True)
                transformer_logs.append(f"Error generating Store Detail: {e}")
            
        logger.info("Process completed successfully.")
        response_stats["generated_file"] = output_filename

        # Update DB record success
        db_record.status = "success"
        db_record.output_filename = output_filename
        db_record.file_path = str(output_path)
        db_record.stats = response_stats
        db.commit()

        return {
            "status": "success",
            "message": "Conversion successful",
            "download_url": f"/api/download/{output_filename}",
            "stats": response_stats,
            "logs": transformer_logs
        }

    except Exception as e:
        logger.error(f"Conversion failed: {str(e)}", exc_info=True)
        
        # Update DB record failure
        if 'db_record' in locals():
            try:
                db_record.status = "failed"
                db_record.error_message = str(e)
                db.commit()
            except:
                pass

        return JSONResponse(
            status_code=500, 
            content={"status": "error", "message": str(e)}
        )
    finally:
        # 清理输入文件
        if input_path.exists():
            try:
                os.remove(input_path)
            except: pass
        
        # Only clean up template if it is in UPLOAD_DIR (temp)
        # We modified logic so that template_path is always a temp copy if we are using it
        if template_path and template_path.exists():
            # Check if it is inside UPLOAD_DIR to be safe
            if UPLOAD_DIR in template_path.parents:
                 try:
                    os.remove(template_path)
                 except: pass

        # 清理明细表
        if detail_path and detail_path.exists():
            try:
                os.remove(detail_path)
            except: pass

@app.get("/api/download/{filename}")
async def download_file(filename: str, background_tasks: BackgroundTasks):
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    
    suffix = file_path.suffix.lower()
    if suffix == ".zip":
        media_type = "application/zip"
    else:
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    # 下载后不立即删除，因为用户可能多次点击，或者设置定时任务清理
    return FileResponse(
        path=file_path, 
        filename=filename, 
        media_type=media_type
    )

if __name__ == "__main__":
    import uvicorn
    import socket
    
    # 获取本机IP的辅助函数
    def get_host_ip():
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(('8.8.8.8', 80))
            ip = s.getsockname()[0]
        except Exception:
            ip = '127.0.0.1'
        finally:
            s.close()
        return ip

    host_ip = get_host_ip()
    print(f"\n{'='*50}")
    print(f"AutoPackage Web 服务已启动")
    print(f"本机访问: http://127.0.0.1:8000")
    print(f"局域网访问: http://{host_ip}:8000 (请在其他电脑使用此地址)")
    print(f"{'='*50}\n")

    # 自动打开浏览器 (使用 localhost)
    import webbrowser
    webbrowser.open("http://127.0.0.1:8000")
    
    # 注意：当使用 reload=True 时，uvicorn 需要能够导入 main 模块
    # 如果直接运行 python main.py，uvicorn 会尝试从当前目录导入 main.py
    # 确保 app 对象在全局作用域中可用
    uvicorn.run("web_app:app", host="0.0.0.0", port=8000, reload=True)
