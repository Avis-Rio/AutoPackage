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

# 同时也添加父目录，以防万一
if str(parent_dir) not in sys.path:
    sys.path.append(str(parent_dir))

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks, Form
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

# 导入现有业务逻辑
try:
    from excel_reader import AllocationTableReader
    from data_transformer import DataTransformer
    from template_writer import TemplateWriter
    from delivery_note_generator import DeliveryNoteGenerator
    from config import FileConfig, DeliveryNoteConfig
except ImportError as e:
    print(f"Error importing core modules: {e}")
    print(f"Current sys.path: {sys.path}")
    sys.exit(1)

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
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
TEMPLATES_DIR.mkdir(exist_ok=True)

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
    file_path = TEMPLATES_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Template not found")
    try:
        os.remove(file_path)
        return {"status": "success", "message": f"Template {filename} deleted"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# --- End Template Management APIs ---

@app.post("/api/convert")
async def convert_file(
    file: UploadFile = File(...),
    template: UploadFile = File(None),
    template_name: str = Form(None), # Optional: name of template in library
    mode: str = Form("allocation") # allocation or delivery_note
):
    """
    核心转换接口
    mode: 'allocation' (default) for 配分表转换
          'delivery_note' for 受渡伝票生成
    """
    input_path = UPLOAD_DIR / f"input_{int(datetime.now().timestamp())}_{file.filename}"
    
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
        else:
             template_filename += ".xlsx"
        template_path = UPLOAD_DIR / template_filename # Placeholder for default
    
    # Output filename
    prefix = "DeliveryNote_" if mode == "delivery_note" else "Converted_"
    output_filename = f"{prefix}{file.filename}"
    if not output_filename.endswith('.xlsx'):
         output_filename = os.path.splitext(output_filename)[0] + '.xlsx'
         
    output_path = OUTPUT_DIR / output_filename

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
                    # Fallback to xlsx if converted
                    source_dir / (DeliveryNoteConfig.TEMPLATE_NAME + "x"),
                    parent_dir / (DeliveryNoteConfig.TEMPLATE_NAME + "x")
                ]
            else:
                candidates = [
                    source_dir / "template.xlsx",
                    source_dir / "template.xls",
                    parent_dir / "template.xlsx",
                    parent_dir / "template.xls"
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

        # 3. 执行核心逻辑
        logger.info("Starting process...")
        
        items_processed = 0
        
        if mode == "delivery_note":
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
            
        else:
            # Standard Allocation Table Conversion
            # Step A: Read
            reader = AllocationTableReader(str(input_path))
            allocation_data = reader.read()
            items_processed = len(allocation_data.get('products', []))
            logger.info(f"Read {items_processed} products from allocation table.")
    
            # Step B: Transform
            transformer = DataTransformer(allocation_data)
            transform_result = transformer.transform()
            
            # Step C: Write
            writer = TemplateWriter(str(real_template_path), str(output_path))
            writer.write(transform_result)
            
        logger.info("Process completed successfully.")

        return {
            "status": "success",
            "message": "Conversion successful",
            "download_url": f"/api/download/{output_filename}",
            "stats": {
                "items_processed": items_processed,
                "generated_file": output_filename
            }
        }

    except Exception as e:
        logger.error(f"Conversion failed: {str(e)}", exc_info=True)
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

@app.get("/api/download/{filename}")
async def download_file(filename: str, background_tasks: BackgroundTasks):
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    
    # 下载后不立即删除，因为用户可能多次点击，或者设置定时任务清理
    return FileResponse(
        path=file_path, 
        filename=filename, 
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == "__main__":
    import uvicorn
    # 自动打开浏览器
    import webbrowser
    webbrowser.open("http://127.0.0.1:8000")
    
    uvicorn.run(app, host="127.0.0.1", port=8000)
