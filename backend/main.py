from __future__ import annotations
import os
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, UploadFile, File, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from data_processor import parse_period, write_input_to_workbook
from aggregation import build_dashboard
from excel_reader import ExcelReader
from drive_service.drive_manager import DriveManager

load_dotenv()
app=FastAPI(title="Neraca Air API",version="2.0.0")
origins=[x.strip() for x in os.getenv("CORS_ORIGINS","http://localhost:5500,http://127.0.0.1:5500").split(",") if x.strip()]
app.add_middleware(CORSMiddleware,allow_origins=origins,allow_credentials=True,allow_methods=["*"],allow_headers=["*"])
drive=DriveManager(); reader=ExcelReader()

class DatasetEdit(BaseModel):
    input: dict

@app.get("/api/health")
def health(): return {"status":"ok","service":"neraca-air-backend"}

@app.get("/api/datasets")
def datasets():
    try: return {"datasets":drive.list_datasets()}
    except Exception as e: raise HTTPException(500,str(e))

@app.get("/api/datasets/{period}")
def detail(period:str):
    try:
        y,m=parse_period(period); f=drive.find_dataset(y,m)
        if not f: raise HTTPException(404,"Dataset tidak ditemukan")
        content=drive.download_file(f["id"])
        return {"period":f["period"],"name":f["name"],"input":reader.read_input(content)}
    except HTTPException: raise
    except ValueError as e: raise HTTPException(400,str(e))
    except Exception as e: raise HTTPException(500,str(e))

@app.get("/api/dashboard")
def dashboard(months:list[str]=Query(...,description="YYYY-MM; ulangi parameter untuk beberapa bulan")):
    try:
        selected=[]
        for period in sorted(set(months)):
            y,m=parse_period(period); f=drive.find_dataset(y,m)
            if not f: continue
            selected.append({"period":f["period"],"input":reader.read_input(drive.download_file(f["id"]))})
        if not selected: raise HTTPException(404,"Tidak ada dataset pada periode terpilih")
        return build_dashboard(selected)
    except HTTPException: raise
    except ValueError as e: raise HTTPException(400,str(e))
    except Exception as e: raise HTTPException(500,str(e))

@app.post("/api/datasets/{period}/upload")
async def upload(period:str,file:UploadFile=File(...)):
    try:
        y,m=parse_period(period); content=await file.read()
        reader.validate_excel(content)
        result=drive.upsert_dataset(y,m,content)
        return {"message":"Dataset disimpan","period":f"{y:04d}-{m:02d}",**result}
    except ValueError as e: raise HTTPException(400,str(e))
    except Exception as e: raise HTTPException(500,str(e))

@app.put("/api/datasets/{period}")
def save_detail(period:str,payload:DatasetEdit):
    """Simpan hasil edit Detail ke workbook bulan yang dipilih; sheet lain/formula dipertahankan."""
    try:
        y,m=parse_period(period); f=drive.find_dataset(y,m)
        if not f: raise HTTPException(404,"Dataset tidak ditemukan; upload workbook terlebih dahulu")
        original=drive.download_file(f["id"])
        updated=write_input_to_workbook(original,payload.input)
        reader.validate_excel(updated)
        result=drive.update_file(f["id"],updated)
        return {"message":"Dataset bulan berhasil diperbarui","period":f"{y:04d}-{m:02d}",**result}
    except HTTPException: raise
    except ValueError as e: raise HTTPException(400,str(e))
    except Exception as e: raise HTTPException(500,str(e))
