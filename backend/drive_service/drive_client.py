import os
from google.oauth2 import service_account
from googleapiclient.discovery import build

def create_drive_service():
    path=os.getenv("GOOGLE_APPLICATION_CREDENTIALS","service-account.json")
    if not os.path.isfile(path):
        raise RuntimeError("Service Account JSON tidak ditemukan; atur GOOGLE_APPLICATION_CREDENTIALS")
    creds=service_account.Credentials.from_service_account_file(
        path,scopes=["https://www.googleapis.com/auth/drive"])
    return build("drive","v3",credentials=creds,cache_discovery=False)
