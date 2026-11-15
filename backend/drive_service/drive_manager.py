from __future__ import annotations
import io, os, re
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from .drive_client import create_drive_service
from data_processor import canonical_filename, parse_period

MIME_XLSX="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

class DriveManager:
    def __init__(self):
        self.service=create_drive_service()
        self.folder_id=os.getenv("GOOGLE_DRIVE_FOLDER_ID")
        if not self.folder_id: raise RuntimeError("GOOGLE_DRIVE_FOLDER_ID belum diatur")

    def list_datasets(self):
        q=f"'{self.folder_id}' in parents and trashed=false"
        token=None; found=[]
        while True:
            res=self.service.files().list(q=q,pageSize=1000,pageToken=token,
                fields="nextPageToken,files(id,name,mimeType,modifiedTime,size)").execute()
            for f in res.get("files",[]):
                if f.get("mimeType")=="application/vnd.google-apps.folder": continue
                name=f["name"]
                m=re.fullmatch(r"WB_([A-Za-z]{3})_(\\d{2}|\\d{4})\\.xlsx",name,re.I)
                if not m: continue
                try:
                    y,mo=parse_period(name[:-5])
                    found.append({"id":f["id"],"name":name,"period":f"{y:04d}-{mo:02d}",
                                  "modified_time":f.get("modifiedTime"),"size":f.get("size")})
                except ValueError: pass
            token=res.get("nextPageToken")
            if not token: break
        # Jika ada nama alias ganda untuk periode sama, ambil file terakhir diubah.
        by_period={}
        for f in found:
            if f["period"] not in by_period or (f.get("modified_time") or "")>(by_period[f["period"]].get("modified_time") or ""):
                by_period[f["period"]]=f
        return sorted(by_period.values(),key=lambda x:x["period"])

    def find_dataset(self,year,month):
        period=f"{year:04d}-{month:02d}"
        return next((x for x in self.list_datasets() if x["period"]==period),None)

    def download_file(self,file_id):
        req=self.service.files().get_media(fileId=file_id); buf=io.BytesIO()
        dl=MediaIoBaseDownload(buf,req); done=False
        while not done: _,done=dl.next_chunk()
        return buf.getvalue()

    def upsert_dataset(self,year,month,content):
        name=canonical_filename(year,month); old=self.find_dataset(year,month)
        media=MediaIoBaseUpload(io.BytesIO(content),mimetype=MIME_XLSX,resumable=True)
        if old:
            f=self.service.files().update(fileId=old["id"],body={"name":name},media_body=media,
                fields="id,name,modifiedTime").execute()
            action="updated"
        else:
            f=self.service.files().create(body={"name":name,"parents":[self.folder_id]},
                media_body=media,fields="id,name,modifiedTime").execute()
            action="created"
        return {"id":f["id"],"name":f["name"],"modified_time":f.get("modifiedTime"),"action":action}

    def update_file(self,file_id,content):
        media=MediaIoBaseUpload(io.BytesIO(content),mimetype=MIME_XLSX,resumable=True)
        f=self.service.files().update(fileId=file_id,media_body=media,fields="id,name,modifiedTime").execute()
        return {"id":f["id"],"name":f["name"],"modified_time":f.get("modifiedTime")}
