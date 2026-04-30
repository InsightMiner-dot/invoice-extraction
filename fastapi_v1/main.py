from fastapi import FastAPI, Request, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.middleware.cors import CORSMiddleware
from starlette.middleware.base import BaseHTTPMiddleware
from typing import List
import os
import io
import sqlite3
import pandas as pd
from dotenv import load_dotenv

from core.database import init_db, fetch_audit_data, DB_PATH
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
    CORSMiddleware, allow_origins=allowed_origins, allow_credentials=True,
    allow_methods=["GET", "POST"], allow_headers=["*"],
)

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
async def home(request: Request):
    return templates.TemplateResponse(request=request, name="extract.html")

@app.get("/viewer", response_class=HTMLResponse)
async def viewer(request: Request):
    return templates.TemplateResponse(request=request, name="viewer.html")

@app.get("/analytics", response_class=HTMLResponse)
async def analytics(request: Request):
    return templates.TemplateResponse(request=request, name="analytics.html")

@app.get("/batch-qa", response_class=HTMLResponse)
async def batch_qa(request: Request):
    return templates.TemplateResponse(request=request, name="batch_qa.html")

@app.get("/chat", response_class=HTMLResponse)
async def chat(request: Request):
    return templates.TemplateResponse(request=request, name="chat.html")

# --- API Routes ---
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

@app.get("/api/download-excel/{batch_id}")
async def download_excel(batch_id: str):
    try:
        conn = sqlite3.connect(DB_PATH)
        df = pd.read_sql_query(f"SELECT * FROM qc_audit WHERE batch_id='{batch_id}'", conn)
        conn.close()
        
        buffer = io.BytesIO()
        df.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)
        
        headers = {'Content-Disposition': f'attachment; filename="Invoice_Batch_{batch_id}.xlsx"'}
        return StreamingResponse(buffer, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers=headers)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000)
