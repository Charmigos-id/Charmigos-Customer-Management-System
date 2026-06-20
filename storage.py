"""storage.py — Charmigos CRM | Google Drive (OAuth) + Google Sheets (Service Account)"""

import io
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

# ══════════════════════════════════════════════════════════════
# KONFIGURASI
# ══════════════════════════════════════════════════════════════

# Untuk Google Drive (OAuth — pakai akun Gmail kamu)
OAUTH_SECRET_FILE = "token_secret.json"   # file yang baru didownload
OAUTH_TOKEN_FILE  = "oauth_token.pkl"     # dibuat otomatis setelah login pertama

# Untuk Google Sheets (Service Account)
SA_CREDS_FILE     = "credentials.json"

# ID folder Drive & Spreadsheet
GDRIVE_FOLDER        = "1YlPRqG1Gh51-_YzXjrAVA4jbvA1KtD5H"   # root folder
GDRIVE_HASIL_FOLDER  = "1uhkQXn7PK4JUXEHBCpJccuSibRhJdRqU"   # subfolder Data hasil pengolahan
GSHEET_ID            = "1A_1954pVqqjwX6-Eux0YHDOibJk0DiFTh__JTwh6DIo"
GSHEET_TAB           = "Transaksi"

# Path lokal (backup)
BASE_DIR      = Path("/Users/fairuzaprasetyo/TA Olah Data")
PROCESSED_DIR = BASE_DIR / "Data" / "Data hasil pengolahan"
GABUNGAN_PATH = PROCESSED_DIR / "Transaksi_Gabungan.xlsx"

# ══════════════════════════════════════════════════════════════
# AUTH
# ══════════════════════════════════════════════════════════════

DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]
SHEET_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def _get_oauth_creds():
    """
    OAuth2 untuk Google Drive.
    - Pertama kali: buka browser untuk login Gmail.
    - Selanjutnya: pakai token tersimpan (oauth_token.pkl).
    """
    creds = None
    token_path = Path(OAUTH_TOKEN_FILE)

    # Load token tersimpan jika ada
    if token_path.exists():
        with open(token_path, "rb") as f:
            creds = pickle.load(f)

    # Refresh token jika expired
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    elif not creds or not creds.valid:
        # Login via browser (hanya sekali)
        flow = InstalledAppFlow.from_client_secrets_file(OAUTH_SECRET_FILE, DRIVE_SCOPES)
        creds = flow.run_local_server(port=0)
        # Simpan token untuk sesi berikutnya
        with open(token_path, "wb") as f:
            pickle.dump(creds, f)

    return creds

def _get_sa_creds():
    """Service Account untuk Google Sheets."""
    return SACredentials.from_service_account_file(SA_CREDS_FILE, scopes=SHEET_SCOPES)

def _drive_service():
    return build("drive", "v3", credentials=_get_oauth_creds())

def _sheets_client():
    return gspread.authorize(_get_sa_creds())

# ══════════════════════════════════════════════════════════════
# HELPER DRIVE
# ══════════════════════════════════════════════════════════════

def _find_file_in_drive(service, filename: str):
    q = f"name='{filename}' and '{GDRIVE_FOLDER}' in parents and trashed=false"
    results = service.files().list(q=q, fields="files(id,name)").execute()
    files = results.get("files", [])
    return files[0]["id"] if files else None

def _upload_or_replace(service, file_bytes: bytes, filename: str):
    mimetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    media    = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mimetype)
    existing = _find_file_in_drive(service, filename)
    if existing:
        service.files().update(fileId=existing, media_body=media).execute()
    else:
        service.files().create(
            body={"name": filename, "parents": [GDRIVE_FOLDER]},
            media_body=media, fields="id"
        ).execute()

# ══════════════════════════════════════════════════════════════
# GOOGLE DRIVE — Fungsi utama
# ══════════════════════════════════════════════════════════════

def save_raw_file(file_bytes: bytes, kanal: str, tahun: int, bulan: int):
    """Simpan file mentah Shopee/TikTok ke Google Drive."""
    filename = f"{kanal}_{tahun}_{bulan:02d}.xlsx"
    _upload_or_replace(_drive_service(), file_bytes, filename)
    print(f"[Drive] Saved: {filename}")

def save_gabungan(df: pd.DataFrame):
    """Simpan Transaksi_Gabungan.xlsx ke Drive + backup lokal."""
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
        print(f"[Lokal] Backup: {GABUNGAN_PATH}")
    except Exception as e:
        print(f"[Lokal] Gagal backup: {e}")

def load_gabungan_bytes():
    """Load Transaksi_Gabungan.xlsx dari Drive, fallback ke lokal."""
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

def gsheet_append_transaksi(df: pd.DataFrame):
    """Append transaksi baru ke Google Sheets, skip duplikat."""
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
        print("[Sheets] Tidak ada baris baru.")
        return

    rows = df_new.fillna("").astype(str).values.tolist()
    ws.append_rows(rows, value_input_option="USER_ENTERED")
    print(f"[Sheets] {len(rows):,} baris baru ditambahkan.")

# ══════════════════════════════════════════════════════════════
# SYNC FILE HASIL → DRIVE + SHEETS
# ══════════════════════════════════════════════════════════════

def sync_hasil_to_drive_and_sheet(df: pd.DataFrame, filename: str):
    """Upload DataFrame ke subfolder Data hasil pengolahan di Drive + update tab Sheets."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    buf.seek(0)
    file_bytes = buf.read()

    # Upload ke subfolder Data hasil pengolahan
    mimetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    service  = _drive_service()
    media    = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mimetype)
    q = f"name='{filename}' and '{GDRIVE_HASIL_FOLDER}' in parents and trashed=false"
    results  = service.files().list(q=q, fields="files(id)").execute()
    files    = results.get("files", [])
    if files:
        service.files().update(fileId=files[0]["id"], media_body=media).execute()
    else:
        service.files().create(
            body={"name": filename, "parents": [GDRIVE_HASIL_FOLDER]},
            media_body=media, fields="id"
        ).execute()
    print(f"[Drive] Synced ke Data hasil pengolahan: {filename}")

    # Update tab Sheets
    tab_name = Path(filename).stem
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
    print(f"[Sheets] Tab '{tab_name}' updated: {len(rows):,} baris.")

# ══════════════════════════════════════════════════════════════
# BACKGROUND SYNC (tidak bikin app lambat)
# ══════════════════════════════════════════════════════════════

def sync_hasil_background(df: pd.DataFrame, filename: str):
    """Sync ke Drive & Sheets di background thread — app tidak perlu menunggu."""
    def _run():
        try: sync_hasil_to_drive_and_sheet(df.copy(), filename)
        except Exception as e: print(f"[Sync BG] Error {filename}: {e}")
    threading.Thread(target=_run, daemon=True).start()

# ══════════════════════════════════════════════════════════════
# STRATEGY DOCS
# ══════════════════════════════════════════════════════════════

def save_strategy_doc(content: str, filename: str):
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
    except Exception as e:
        err = f"Sheets error: {str(e)[:120]}"
        result["error"] = (result["error"] + " | " + err) if result["error"] else err

    return result