from fastapi import FastAPI, Request, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.middleware.cors import CORSMiddleware
from starlette.middleware.base import BaseHTTPMiddleware
import os
import json
import traceback
from dotenv import load_dotenv

from core.database import init_db, fetch_audit_data
from services.extraction import process_single_invoice, generate_excel_from_db

load_dotenv(override=True)
is_prod = os.getenv("ENVIRONMENT", "development") == "production"

app = FastAPI(
    title="Invoice Intelligence API",
    docs_url=None if is_prod else "/docs",
    redoc_url=None if is_prod else "/redoc",
    openapi_url=None if is_prod else "/openapi.json"
)

# --- Security & CORS ---
allowed_origins = os.getenv("ALLOWED_ORIGINS", "http://localhost:8000").split(",")
app.add_middleware(CORSMiddleware, allow_origins=allowed_origins, allow_credentials=True, allow_methods=["GET", "POST"], allow_headers=["*"])

class SecurityHeadersMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        response = await call_next(request)
        response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains; preload"
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["X-XSS-Protection"] = "1; mode=block"
        response.headers["Content-Security-Policy"] = "default-src 'self' blob:; script-src 'self' 'unsafe-inline' https://cdn.plot.ly; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:;"
        return response

app.add_middleware(SecurityHeadersMiddleware)

os.makedirs("static", exist_ok=True)
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.on_event("startup")
def startup_event():
    init_db()

# --- Page Routes ---
@app.get("/", response_class=HTMLResponse)
async def home(request: Request): return templates.TemplateResponse(request=request, name="extract.html")

@app.get("/analytics", response_class=HTMLResponse)
async def analytics(request: Request): return templates.TemplateResponse(request=request, name="analytics.html")

@app.get("/batch-qa", response_class=HTMLResponse)
async def batch_qa(request: Request): return templates.TemplateResponse(request=request, name="batch_qa.html")

# --- API Routes ---
@app.post("/api/extract-single")
async def extract_single_api(
    file: UploadFile = File(...), 
    batch_id: str = Form(...),
    max_pages: int = Form(15),
    dpi: int = Form(300),
    aliases: str = Form("{}"),
    custom_fields: str = Form("{}")
):
    try:
        file_bytes = await file.read()
        aliases_dict = json.loads(aliases)
        custom_fields_dict = json.loads(custom_fields)
        
        result = await process_single_invoice(
            file_bytes=file_bytes, filename=file.filename, batch_id=batch_id, 
            max_pages=max_pages, dpi=dpi, aliases=aliases_dict, custom_fields=custom_fields_dict
        )
        
        if "error" in result:
            # Send the explicit error back to the frontend
            raise HTTPException(status_code=500, detail=result["error"])
            
        return {"status": "success", "data": result}
        
    except Exception as e:
        error_msg = f"Exception: {str(e)}\n{traceback.format_exc()}"
        print(error_msg) # Logs to your terminal
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/generate-excel")
async def generate_excel_api(batch_id: str = Form(...), custom_fields: str = Form("[]")):
    try:
        custom_cols = json.loads(custom_fields)
        success = generate_excel_from_db(batch_id, custom_cols)
        if success: return {"status": "success"}
        raise HTTPException(status_code=404, detail="No data found for this batch")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/download-excel/{batch_id}")
async def download_excel(batch_id: str):
    file_path = os.path.join("audit", f"{batch_id}.xlsx")
    if os.path.exists(file_path):
        return FileResponse(path=file_path, filename=f"Invoice_Extraction_{batch_id}.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    raise HTTPException(status_code=404, detail="Excel file not found on server")

@app.get("/api/audit-data")
async def audit_data_api():
    return JSONResponse(content=fetch_audit_data().to_dict(orient="records"))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
