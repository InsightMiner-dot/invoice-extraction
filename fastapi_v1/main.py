from fastapi import FastAPI, Request, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.middleware.cors import CORSMiddleware
from starlette.middleware.base import BaseHTTPMiddleware
from typing import List
import os
from dotenv import load_dotenv

from core.database import init_db, fetch_audit_data
from services.extraction import process_batch_concurrently

load_dotenv(override=True)

is_prod = os.getenv("ENVIRONMENT", "development") == "production"
app = FastAPI(
    title="Invoice Intelligence API",
    docs_url=None if is_prod else "/docs",
    redoc_url=None if is_prod else "/redoc",
    openapi_url=None if is_prod else "/openapi.json"
)

allowed_origins = os.getenv("ALLOWED_ORIGINS", "http://localhost:8000").split(",")
app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins,
    allow_credentials=True,
    allow_methods=["GET", "POST"], 
    allow_headers=["*"],
)

class SecurityHeadersMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        response = await call_next(request)
        response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains; preload"
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["X-XSS-Protection"] = "1; mode=block"
        # Adjusted CSP to allow blob: for the local PDF viewer
        response.headers["Content-Security-Policy"] = "default-src 'self' blob:; script-src 'self' 'unsafe-inline' https://cdn.plot.ly; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:;"
        response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
        return response

app.add_middleware(SecurityHeadersMiddleware)

os.makedirs("static", exist_ok=True)
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.on_event("startup")
def startup_event():
    init_db()

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse(request=request, name="extract.html")

@app.get("/analytics", response_class=HTMLResponse)
async def analytics(request: Request):
    return templates.TemplateResponse(request=request, name="analytics.html")

@app.get("/batch-qa", response_class=HTMLResponse)
async def batch_qa(request: Request):
    return templates.TemplateResponse(request=request, name="batch_qa.html")

@app.post("/api/extract")
async def extract_api(files: List[UploadFile] = File(...)):
    try:
        file_bytes = [await f.read() for f in files]
        file_names = [f.filename for f in files]
        results = await process_batch_concurrently(file_bytes, file_names)
        return {"status": "success", "data": results}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/audit-data")
async def audit_data_api():
    df = fetch_audit_data()
    return JSONResponse(content=df.to_dict(orient="records"))
import json

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
        
        # Parse settings
        aliases_dict = json.loads(aliases)
        custom_fields_dict = json.loads(custom_fields)
        
        result = await process_single_invoice(
            file_bytes, file.filename, batch_id, max_pages, dpi, aliases_dict, custom_fields_dict
        )
        if "error" in result: raise HTTPException(status_code=500, detail=result["error"])
        return {"status": "success", "data": result}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/generate-excel")
async def generate_excel_api(batch_id: str = Form(...), custom_fields: str = Form("[]")):
    custom_cols = json.loads(custom_fields) # Need the keys to build Excel headers
    success = generate_excel_from_db(batch_id, custom_cols)
    if success: return {"status": "success"}
    raise HTTPException(status_code=404, detail="No data found for this batch")
    
# NEW ENDPOINT: Download the generated Excel file
@app.get("/api/download-excel/{batch_id}")
async def download_excel(batch_id: str):
    file_path = os.path.join("audit", f"{batch_id}.xlsx")
    if os.path.exists(file_path):
        return FileResponse(
            path=file_path, 
            filename=f"Invoice_Extraction_{batch_id}.xlsx", 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    raise HTTPException(status_code=404, detail="Excel file not found")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000)
