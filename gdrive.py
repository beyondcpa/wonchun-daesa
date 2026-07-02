# -*- coding: utf-8 -*-
"""구글드라이브 폴더를 서비스계정으로 읽어 임시폴더에 내려받는다(하위폴더 구조 유지).
   앱에서 폴더 링크/ID만 주면 그 아래 모든 파일을 다운로드 → engine.scan_year 로 파싱."""
import io, os, re

def get_service(creds):
    import json
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    if isinstance(creds, str):
        creds = json.loads(creds)
    info = dict(creds)
    c = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive.readonly"])
    return build("drive", "v3", credentials=c, cache_discovery=False)

def extract_folder_id(url_or_id):
    s = (url_or_id or "").strip()
    m = re.search(r"/folders/([A-Za-z0-9_\-]+)", s)
    if m: return m.group(1)
    m = re.search(r"[?&]id=([A-Za-z0-9_\-]+)", s)
    if m: return m.group(1)
    return s

_FOLDER = "application/vnd.google-apps.folder"

def _walk(svc, folder_id, prefix=""):
    """(상대경로, 파일ID) 제너레이터 — 하위폴더 재귀."""
    page = None
    while True:
        resp = svc.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="nextPageToken, files(id,name,mimeType)",
            pageSize=1000, supportsAllDrives=True, includeItemsFromAllDrives=True,
            pageToken=page).execute()
        for f in resp.get("files", []):
            if f["mimeType"] == _FOLDER:
                yield from _walk(svc, f["id"], prefix + f["name"] + "/")
            else:
                yield (prefix + f["name"], f["id"])
        page = resp.get("nextPageToken")
        if not page:
            break

def download_folder(svc, folder_id, dest):
    """폴더 아래 모든 파일을 dest에 구조 유지하여 저장. 받은 파일 수 반환."""
    from googleapiclient.http import MediaIoBaseDownload
    n = 0
    for relpath, fid in _walk(svc, folder_id):
        try:
            outp = os.path.join(dest, relpath)
            os.makedirs(os.path.dirname(outp), exist_ok=True)
            req = svc.files().get_media(fileId=fid, supportsAllDrives=True)
            buf = io.BytesIO(); dl = MediaIoBaseDownload(buf, req)
            done = False
            while not done:
                _, done = dl.next_chunk()
            with open(outp, "wb") as out:
                out.write(buf.getvalue())
            n += 1
        except Exception:
            pass
    return n
