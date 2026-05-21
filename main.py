import os
from fastapi import FastAPI, Request, File, UploadFile
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles

from core.parser import parse_financial_excel
from core.validation import FinancialValidationError

app = FastAPI(title="Financial Parser Dashboard")

# Ensure templates directory exists before mounting
os.makedirs("templates", exist_ok=True)
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def get_upload_page(request: Request):
    return templates.TemplateResponse("upload.html", {"request": request})

@app.post("/upload", response_class=HTMLResponse)
async def upload_file(request: Request, file: UploadFile = File(...)):
    try:
        # Read the file into memory
        contents = await file.read()
        
        # Parse the financial excel file
        report_data = parse_financial_excel(contents)
        
        # Generate downloadable excel report
        from core.excel_generator import generate_excel_report
        export_filename = generate_excel_report(report_data)
        
        return templates.TemplateResponse("report.html", {
            "request": request,
            "filename": file.filename,
            "export_filename": export_filename,
            "data": report_data
        })
        
    except FinancialValidationError as e:
        # If validation fails, return the error page
        return templates.TemplateResponse("error.html", {
            "request": request,
            "error_title": e.title,
            "error_message": str(e),
            "details": e.details
        })
    except Exception as e:
        # Catch-all for other errors
        return templates.TemplateResponse("error.html", {
            "request": request,
            "error_title": "Unexpected Error",
            "error_message": "An unexpected error occurred during processing.",
            "details": [str(e)]
        })

from fastapi.responses import FileResponse
import os

@app.get("/download/{filename}")
async def download_file(filename: str):
    file_path = os.path.join(os.getcwd(), "temp_reports", filename)
    if os.path.exists(file_path):
        return FileResponse(file_path, filename="Financial_Report.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return HTMLResponse("File not found", status_code=404)
