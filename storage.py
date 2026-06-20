"""storage.py — Charmigos CRM | Sheets (cloud) + Drive OAuth (lokal)"""

import io
import json
import threading
from pathlib import Path

import pandas as pd

# ── Google API Libraries ───────────────────────────────────────
try:
    import gspread
    from google.oauth2.service_account import Credentials as SACredentials
    from google.oauth2.credentials import Credentials as OAuthCredentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload
    import pickle
    _GOOGLE_LIBS_OK = True
except ImportError:
    _GOOGLE_LIBS_OK = False

# ── Deteksi environment: Streamlit Cloud atau lokal ───────────
def _is_cloud() -> bool:
    try:
        import streamlit as st
        _ = st.secrets["GOOGLE_CREDENTIALS"]
        return True
    except Exception:
        return False

# ══════════════════════════════════════════════════════════════
# KONFIGURASI
# ══════════════════════════════════════════════════════════════

SA_CREDS_FILE        = "credentials.json"
OAUTH_SECRET_FILE    = "token_secret.json"
OAUTH_TOKEN_FILE     = "oauth_token.pkl"

GDRIVE_FOLDER        = "1YlPRqG1Gh51-_YzXjrAVA4jbvA1KtD5H"
GDRIVE_HASIL_FOLDER  = "1uhkQXn7PK4JUXEHBCpJccuSibRhJdRqU"
GSHEET_ID            = "1A_1954pVqqjwX6-Eux0YHDOibJk0DiFTh__JTwh6DIo"
GSHEET_TAB           = "Transaksi"
GSHEET_GABUNGAN_TAB  = "Transaksi_Gabungan"

# Path lokal (hanya dipakai saat development)
BASE_DIR      = Path("/Users/fairuzaprasetyo/TA Olah Data")
PROCESSED_DIR = BASE_DIR / "Data" / "Data hasil pengolahan"
GABUNGAN_PATH = PROCESSED_DIR / "Transaksi_Gabungan.xlsx"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ══════════════════════════════════════════════════════════════
# AUTH — otomatis pilih cloud vs lokal
# ══════════════════════════════════════════════════════════════

def _get_sa_creds():
    """Load Service Account credentials — dari Secrets (cloud) atau file (lokal)."""
    if _is_cloud():
        import streamlit as st
        creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
        import tempfile, os
        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
        json.dump(creds_dict, tmp)
        tmp.close()
        creds = SACredentials.from_service_account_file(tmp.name, scopes=SCOPES)
        os.unlink(tmp.name)
        return creds
    return SACredentials.from_service_account_file(SA_CREDS_FILE, scopes=SCOPES)

def _get_oauth_creds():
    """OAuth2 — hanya untuk lokal (Drive upload). Di cloud skip."""
    creds = None
    token_path = Path(OAUTH_TOKEN_FILE)
    if token_path.exists():
        with open(token_path, "rb") as f:
            creds = pickle.load(f)
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    elif not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(OAUTH_SECRET_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        with open(token_path, "wb") as f:
            pickle.dump(creds, f)
    return creds

def _drive_service():
    """Drive pakai OAuth (lokal) atau Service Account (cloud)."""
    if _is_cloud():
        return build("drive", "v3", credentials=_get_sa_creds())
    return build("drive", "v3", credentials=_get_oauth_creds())

def _sheets_client():
    return gspread.authorize(_get_sa_creds())

# ══════════════════════════════════════════════════════════════
# HELPER DRIVE
# ══════════════════════════════════════════════════════════════

def _find_file_in_drive(service, filename: str, folder_id: str = None):
    fid = folder_id or GDRIVE_FOLDER
    q = f"name='{filename}' and '{fid}' in parents and trashed=false"
    results = service.files().list(q=q, fields="files(id)").execute()
    files = results.get("files", [])
    return files[0]["id"] if files else None

def _upload_or_replace(service, file_bytes: bytes, filename: str, folder_id: str = None):
    fid      = folder_id or GDRIVE_FOLDER
    mimetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    media    = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mimetype)
    existing = _find_file_in_drive(service, filename, fid)
    if existing:
        service.files().update(fileId=existing, media_body=media).execute()
    else:
        service.files().create(
            body={"name": filename, "parents": [fid]},
            media_body=media, fields="id"
        ).execute()

# ══════════════════════════════════════════════════════════════
# GOOGLE DRIVE — upload file mentah
# ══════════════════════════════════════════════════════════════

def save_raw_file(file_bytes: bytes, kanal: str, tahun: int, bulan: int):
    """Simpan file mentah Shopee/TikTok ke Google Drive."""
    if _is_cloud():
        print("[Drive] Skip upload di cloud — pakai Sheets sebagai storage utama.")
        return
    filename = f"{kanal}_{tahun}_{bulan:02d}.xlsx"
    _upload_or_replace(_drive_service(), file_bytes, filename)
    print(f"[Drive] Saved: {filename}")

# ══════════════════════════════════════════════════════════════
# GABUNGAN — simpan & load via Sheets (cloud) atau Drive (lokal)
# ══════════════════════════════════════════════════════════════

def save_gabungan(df: pd.DataFrame):
    """
    Cloud : simpan ke tab Transaksi_Gabungan di Google Sheets.
    Lokal : simpan ke Drive + backup lokal.
    """
    if _is_cloud():
        _sync_tab(df, GSHEET_GABUNGAN_TAB)
        print("[Sheets] Transaksi_Gabungan disimpan.")
        return

    # Lokal — simpan ke Drive
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Transaksi_Gabungan")
    buf.seek(0)
    file_bytes = buf.read()
    _upload_or_replace(_drive_service(), file_bytes, "Transaksi_Gabungan.xlsx")
    print("[Drive] Saved: Transaksi_Gabungan.xlsx")

    # Backup lokal
    try:
        PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
        GABUNGAN_PATH.write_bytes(file_bytes)
    except Exception as e:
        print(f"[Lokal] Gagal backup: {e}")

def load_gabungan_bytes():
    """
    Cloud : load dari tab Transaksi_Gabungan di Google Sheets → return bytes xlsx.
    Lokal : load dari Drive, fallback ke file lokal.
    """
    if _is_cloud():
        try:
            gc = _sheets_client()
            sh = gc.open_by_key(GSHEET_ID)
            ws = sh.worksheet(GSHEET_GABUNGAN_TAB)
            data = ws.get_all_values()
            if not data:
                return None
            df = pd.DataFrame(data[1:], columns=data[0])
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False)
            buf.seek(0)
            return buf.read()
        except Exception as e:
            print(f"[Sheets] Gagal load Transaksi_Gabungan: {e}")
            return None

    # Lokal
    try:
        service = _drive_service()
        file_id = _find_file_in_drive(service, "Transaksi_Gabungan.xlsx")
        if file_id:
            return service.files().get_media(fileId=file_id).execute()
    except Exception as e:
        print(f"[Drive] Gagal load: {e}")
    if GABUNGAN_PATH.exists():
        return GABUNGAN_PATH.read_bytes()
    return None

# ══════════════════════════════════════════════════════════════
# GOOGLE SHEETS
# ══════════════════════════════════════════════════════════════

def _sync_tab(df: pd.DataFrame, tab_name: str):
    """Replace isi tab Sheets dengan data terbaru."""
    gc = _sheets_client()
    sh = gc.open_by_key(GSHEET_ID)
    try:
        ws = sh.worksheet(tab_name)
        ws.clear()
    except Exception:
        ws = sh.add_worksheet(title=tab_name, rows="10000", cols="50")
    header = df.columns.tolist()
    rows   = df.fillna("").astype(str).values.tolist()
    ws.update([header] + rows, value_input_option="USER_ENTERED")

def gsheet_append_transaksi(df: pd.DataFrame):
    """Append transaksi baru ke tab Transaksi, skip duplikat."""
    gc = _sheets_client()
    sh = gc.open_by_key(GSHEET_ID)
    try:
        ws = sh.worksheet(GSHEET_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=GSHEET_TAB, rows="10000", cols="30")

    existing_data = ws.get_all_values()
    if not existing_data:
        header = df.columns.tolist()
        rows   = df.fillna("").astype(str).values.tolist()
        ws.append_row(header, value_input_option="USER_ENTERED")
        if rows:
            ws.append_rows(rows, value_input_option="USER_ENTERED")
        return

    header_row = existing_data[0]
    try:
        pesanan_col     = header_row.index("No. Pesanan")
        existing_orders = {row[pesanan_col] for row in existing_data[1:] if len(row) > pesanan_col}
        df_new          = df[~df["No. Pesanan"].astype(str).isin(existing_orders)]
    except ValueError:
        df_new = df

    if df_new.empty:
        return
    rows = df_new.fillna("").astype(str).values.tolist()
    ws.append_rows(rows, value_input_option="USER_ENTERED")
    print(f"[Sheets] {len(rows):,} baris baru ditambahkan.")

# ══════════════════════════════════════════════════════════════
# SYNC HASIL → SHEETS (cloud) + DRIVE (lokal)
# ══════════════════════════════════════════════════════════════

def sync_hasil_to_drive_and_sheet(df: pd.DataFrame, filename: str):
    """
    Cloud : update tab Sheets saja.
    Lokal : upload ke subfolder Drive + update tab Sheets.
    """
    tab_name = Path(filename).stem

    if not _is_cloud():
        # Upload ke Drive lokal
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False)
        buf.seek(0)
        _upload_or_replace(_drive_service(), buf.read(), filename, GDRIVE_HASIL_FOLDER)
        print(f"[Drive] Synced: {filename}")

    # Update tab Sheets (cloud & lokal)
    _sync_tab(df, tab_name)
    print(f"[Sheets] Tab '{tab_name}' updated.")

def sync_hasil_background(df: pd.DataFrame, filename: str):
    """Sync di background thread — app tidak lambat."""
    def _run():
        try: sync_hasil_to_drive_and_sheet(df.copy(), filename)
        except Exception as e: print(f"[Sync BG] Error {filename}: {e}")
    threading.Thread(target=_run, daemon=True).start()

# ══════════════════════════════════════════════════════════════
# STRATEGY DOCS
# ══════════════════════════════════════════════════════════════

def save_strategy_doc(content: str, filename: str):
    if _is_cloud():
        print("[Cloud] Strategy doc tidak disimpan ke Drive di cloud.")
        return
    file_bytes = content.encode("utf-8")
    service    = _drive_service()
    media      = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype="text/plain")
    existing   = _find_file_in_drive(service, filename)
    if existing:
        service.files().update(fileId=existing, media_body=media).execute()
    else:
        service.files().create(
            body={"name": filename, "parents": [GDRIVE_FOLDER]},
            media_body=media, fields="id"
        ).execute()

def list_strategy_docs() -> list:
    if _is_cloud():
        return []
    service = _drive_service()
    q = (f"'{GDRIVE_FOLDER}' in parents and trashed=false "
         "and (name contains '.txt' or name contains '.md')")
    results = service.files().list(q=q, fields="files(id,name,modifiedTime)").execute()
    return [{"id": f["id"], "name": f["name"], "modified": f.get("modifiedTime", "")}
            for f in results.get("files", [])]

# ══════════════════════════════════════════════════════════════
# CEK STATUS INTEGRASI
# ══════════════════════════════════════════════════════════════

def check_integrations() -> dict:
    result = {"gdrive": False, "gsheet": False, "error": None}
    if not _GOOGLE_LIBS_OK:
        result["error"] = "Library belum diinstall."
        return result

    if not _is_cloud():
        try:
            service = _drive_service()
            service.files().list(q=f"'{GDRIVE_FOLDER}' in parents",
                                 pageSize=1, fields="files(id)").execute()
            result["gdrive"] = True
        except Exception as e:
            result["error"] = f"Drive error: {str(e)[:120]}"

    try:
        gc = _sheets_client()
        gc.open_by_key(GSHEET_ID)
        result["gsheet"] = True
        if _is_cloud():
            result["gdrive"] = True  # di cloud, Sheets = storage utama
    except Exception as e:
        err = f"Sheets error: {str(e)[:120]}"
        result["error"] = (result["error"] + " | " + err) if result["error"] else err

    return result