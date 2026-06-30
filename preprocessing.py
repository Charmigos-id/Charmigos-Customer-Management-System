"""
preprocessing.py — SESUAI PERSIS Shopee.ipynb + TikTok_2.ipynb + Merged_2.ipynb

Alur Shopee (Cell 0 → Cell 2 → Cell 1 → Cell 3):
  1. Baca semua raw file dengan dtype=str
  2. Bersihkan kolom harga: strip "Rp", hapus dot ribuan → int
  3. Rename "Jumlah Produk di Pesan"→"Jumlah", "Harga Awal Produk"→"Harga Awal"
  4. Total Harga Produk = Jumlah × Harga Awal
  5. Groupby No. Pesanan

Alur TikTok (Cell 0 → Cell 2 → Cell 3 → Cell 4):
  1. Baca raw file dengan dtype={"Order ID": str}
  2. Rename: Order ID→No. Pesanan, Created Time→Waktu Pesanan Dibuat, dll.
  3. Total Harga Produk = Jumlah × Harga Awal
  4. Groupby No. Pesanan
  5. Hapus null Waktu + Username, hapus Total == 0

Alur Merge (Merged_2.ipynb Cell 0 → Cell 1):
  1. Concat Shopee + TikTok
  2. Tambah kolom Kanal
  3. Waktu Pesanan Dibuat → Waktu Pemesanan (pd.to_datetime)
  4. Drop Waktu Pesanan Dibuat
  5. Clean: dropna, drop Total==0
  6. Sort by Waktu Pemesanan

NOTE: Notebook TIDAK filter Status Pesanan (file raw sudah berisi order selesai saja).
"""
import io, hashlib, glob
from pathlib import Path
import pandas as pd
import numpy as np


# ══════════════════════════════════════════════════════════════
# HELPER
# ══════════════════════════════════════════════════════════════
def _read_excel_all_str(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Baca Excel sebagai semua string (sesuai Shopee.ipynb Cell 0 dtype=str)."""
    fn = (filename or "").lower()
    try:
        if fn.endswith(".xls"):
            return pd.read_excel(io.BytesIO(file_bytes), engine="xlrd", dtype=str)
        if fn.endswith(".csv"):
            for enc in ["utf-8","utf-8-sig","latin1","cp1252"]:
                try: return pd.read_csv(io.BytesIO(file_bytes), encoding=enc, dtype=str)
                except Exception: continue
        return pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl", dtype=str)
    except Exception:
        try: return pd.read_excel(io.BytesIO(file_bytes), dtype=str)
        except Exception: return pd.DataFrame()


def _read_excel_order_str(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Baca Excel dengan Order ID sebagai string (sesuai TikTok_2.ipynb Cell 0)."""
    fn = (filename or "").lower()
    try:
        if fn.endswith(".csv"):
            for enc in ["utf-8","utf-8-sig","latin1","cp1252"]:
                try: return pd.read_csv(io.BytesIO(file_bytes), encoding=enc, dtype={"Order ID":str})
                except Exception: continue
        return pd.read_excel(io.BytesIO(file_bytes), dtype={"Order ID": str})
    except Exception:
        try: return pd.read_excel(io.BytesIO(file_bytes))
        except Exception: return pd.DataFrame()


def _get_1d(df: pd.DataFrame, col: str) -> pd.Series:
    """Ambil kolom, pastikan 1D (handle duplikat kolom)."""
    if col not in df.columns:
        return pd.Series([np.nan]*len(df), index=df.index)
    c = df[col]
    return c.iloc[:, 0] if isinstance(c, pd.DataFrame) else c


def _clean_harga(series: pd.Series) -> pd.Series:
    """
    Sesuai Shopee.ipynb Cell 0:
    'Rp 15.000' → 15000
    '15.000'    → 15000
    '15000'     → 15000
    """
    s = series.astype(str).str.strip()
    s = s.str.replace("Rp", "", regex=False).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)   # hapus trailing .0
    s = s.str.replace(".", "", regex=False)        # hapus titik ribuan
    s = s.str.replace(",", "", regex=False)        # hapus koma juga
    return pd.to_numeric(s, errors="coerce").fillna(0)


# Kolom harga Shopee yang harus dibersihkan (sesuai Shopee.ipynb Cell 0)
_SHOPEE_KOLOM_HARGA = [
    "Harga Awal", "Harga Setelah Diskon", "Total Harga Produk", "Total Diskon",
    "Diskon Dari Penjual", "Diskon Dari Shopee", "Voucher Ditanggung Penjual",
    "Cashback Koin", "Voucher Ditanggung Shopee",
    "Paket Diskon (Diskon dari Shopee)", "Paket Diskon (Diskon dari Penjual)",
    "Potongan Koin Shopee", "Diskon Kartu Kredit",
    "Ongkos Kirim Dibayar oleh Pembeli", "Estimasi Potongan Biaya Pengiriman",
    "Ongkos Kirim Pengembalian Barang", "Total Pembayaran", "Perkiraan Ongkos Kirim",
    # Nama alternatif yang mungkin muncul sebelum rename
    "Harga Awal Produk",
]


def _stable_uid(username_series: pd.Series, prefix: str) -> pd.Series:
    """User ID stabil: username sama → ID sama di semua file/sesi."""
    def _h(u):
        v = int(hashlib.md5(str(u).encode()).hexdigest(), 16) % 99999
        return f"{prefix}_{v+1:05d}"
    return username_series.fillna("UNKNOWN").astype(str).map(_h)


# ══════════════════════════════════════════════════════════════
# SHOPEE PREPROCESSING
# ══════════════════════════════════════════════════════════════
def preprocess_shopee(file_bytes: bytes, filename: str = "upload.xlsx") -> pd.DataFrame:
    """
    Sesuai Shopee.ipynb:
    Cell 0: baca dtype=str, bersihkan harga
    Cell 2: select + rename kolom
    Cell 3: Total = Jumlah × Harga Awal, groupby No. Pesanan
    """
    # ── Step 1: Baca sebagai string (Cell 0) ─────────────────
    df = _read_excel_all_str(file_bytes, filename)
    if df is None or df.empty:
        return pd.DataFrame()
    df.columns = df.columns.str.strip()
    df = df.loc[:, ~df.columns.duplicated()].copy()

    # ── Step 2: Bersihkan kolom harga (Cell 0) ───────────────
    for col in _SHOPEE_KOLOM_HARGA:
        if col in df.columns:
            df[col] = _clean_harga(_get_1d(df, col))

    # ── Step 2b: Filter pesanan dibatalkan/dikembalikan (Shopee.ipynb Cell 1) ──
    # Hapus baris yang punya status pembatalan/pengembalian atau total bayar = 0
    _cancel_mask = pd.Series([False] * len(df), index=df.index)
    for _cc in ["Status Pembatalan/ Pengembalian", "Status Pembatalan/Pengembalian",
                "Alasan Pembatalan", "Alasan Pengembalian"]:
        if _cc in df.columns:
            _cancel_mask = _cancel_mask | _get_1d(df, _cc).notna()
    if "Total Pembayaran" in df.columns:
        _tp = pd.to_numeric(
            _get_1d(df, "Total Pembayaran").astype(str)
            .str.replace("Rp","",regex=False).str.replace(".","",regex=False).str.strip(),
            errors="coerce").fillna(0)
        _cancel_mask = _cancel_mask | (_tp == 0)
    _before = len(df)
    df = df[~_cancel_mask].copy()
    if _before > len(df):
        pass  # cancelled rows removed silently; visible in cleaning_stats

    # ── Step 3: Rename kolom (Cell 1/Cell 2) ─────────────────
    rename_map = {}
    # Jumlah
    for alt in ["Jumlah Produk di Pesan", "Jumlah Produk Di Pesan",
                "Jumlah Produk Dipesan", "Qty"]:
        if alt in df.columns and "Jumlah" not in df.columns:
            rename_map[alt] = "Jumlah"; break
    # Harga Awal
    if "Harga Awal Produk" in df.columns and "Harga Awal" not in df.columns:
        rename_map["Harga Awal Produk"] = "Harga Awal"
    # Waktu
    for alt in ["Waktu Pesanan Dibuat", "Order Creation Time", "Waktu Transaksi"]:
        if alt in df.columns and "Waktu Pesanan Dibuat" not in df.columns:
            rename_map[alt] = "Waktu Pesanan Dibuat"; break
    # Username
    for alt in ["Pembeli", "Nama Pembeli", "Buyer Username", "Username"]:
        if alt in df.columns and "Username (Pembeli)" not in df.columns:
            rename_map[alt] = "Username (Pembeli)"; break
    # No. Pesanan
    for alt in ["Order ID", "No Pesanan", "Nomor Pesanan"]:
        if alt in df.columns and "No. Pesanan" not in df.columns:
            rename_map[alt] = "No. Pesanan"; break
    if rename_map:
        df = df.rename(columns=rename_map)

    # Validasi kolom wajib
    if "No. Pesanan" not in df.columns:
        raise ValueError(f"Kolom No. Pesanan tidak ditemukan di {filename}. "
                         f"Kolom tersedia: {list(df.columns[:10])}")
    if "Jumlah" not in df.columns or "Harga Awal" not in df.columns:
        raise ValueError(f"Kolom Jumlah/Harga Awal tidak ditemukan di {filename}.")

    # ── Step 4: Pastikan Jumlah numerik (Cell 3) ─────────────
    df["Jumlah"] = pd.to_numeric(_get_1d(df, "Jumlah"), errors="coerce").fillna(0)

    # ── Step 5: Total Harga Produk = Jumlah × Harga Awal ─────
    # Sesuai Shopee.ipynb Cell 3
    df["Total Harga Produk"] = df["Jumlah"] * _get_1d(df, "Harga Awal")

    # ── Step 6: User ID (Cell 2) ─────────────────────────────
    if "User ID" not in df.columns:
        df["User ID"] = _stable_uid(_get_1d(df, "Username (Pembeli)"), "Shope")

    # ── Step 7: Groupby No. Pesanan (Cell 3) ─────────────────
    df = df.loc[:, ~df.columns.duplicated()]
    agg = {
        "Waktu Pesanan Dibuat": "first",
        "Username (Pembeli)":   "first",
        "Total Harga Produk":   "sum",
        "User ID":              "first",
    }
    for c in ["Kota/Kabupaten", "Provinsi"]:
        if c in df.columns: agg[c] = "first"
    agg = {k: v for k, v in agg.items() if k in df.columns}
    df["No. Pesanan"] = _get_1d(df, "No. Pesanan").astype(str).str.strip()
    df = df.groupby("No. Pesanan", as_index=False).agg(agg)

    # ── Step 8: Clean (Cell 3/4) ─────────────────────────────
    df["Total Harga Produk"] = pd.to_numeric(
        _get_1d(df, "Total Harga Produk"), errors="coerce").fillna(0)
    df = df.dropna(subset=["Waktu Pesanan Dibuat", "Username (Pembeli)"])
    df = df[df["Total Harga Produk"] > 0]
    df["Kanal"] = "Shopee"

    return df.reset_index(drop=True)


# ══════════════════════════════════════════════════════════════
# TIKTOK PREPROCESSING
# ══════════════════════════════════════════════════════════════
def preprocess_tiktok(file_bytes: bytes, filename: str = "upload.xlsx") -> pd.DataFrame:
    """
    Sesuai TikTok_2.ipynb:
    Cell 0: baca dtype={"Order ID":str}, sort by Created Time
    Cell 2: filter kolom
    Cell 3: rename kolom
    Cell 4: Total = Jumlah × Harga Awal, groupby, clean null & 0
    """
    # ── Step 1: Baca (Cell 0) ─────────────────────────────────
    df = _read_excel_order_str(file_bytes, filename)
    if df is None or df.empty:
        return pd.DataFrame()
    df.columns = df.columns.str.strip()
    df = df.loc[:, ~df.columns.duplicated()].copy()

    # ── Step 2: Order ID as string (Cell 0) ───────────────────
    if "Order ID" in df.columns:
        df["Order ID"] = _get_1d(df, "Order ID").astype(str).str.strip()

    # ── Step 3: Rename (Cell 3) ───────────────────────────────
    rename_map = {
        "Order ID":                    "No. Pesanan",
        "Order Status":                "Status Pesanan",
        "Created Time":                "Waktu Pesanan Dibuat",
        "Variation":                   "Nama Variasi",
        "Quantity":                    "Jumlah",
        "SKU Unit Original Price":     "Harga Awal",
        "Buyer Username":              "Username (Pembeli)",
        "Province":                    "Provinsi",
        "Regency and City":            "Kota/Kabupaten",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    if "No. Pesanan" not in df.columns:
        raise ValueError(f"Kolom Order ID tidak ditemukan di {filename}.")

    df["No. Pesanan"] = _get_1d(df, "No. Pesanan").astype(str).str.strip()

    # ── Filter cancel/return (TikTok_2.ipynb Cell 3) ──────────
    # Hapus baris yang punya nilai di "Cancelation/Return Type"
    if "Cancelation/Return Type" in df.columns:
        _before_tt = len(df)
        df = df[_get_1d(df, "Cancelation/Return Type").isna()].copy()

    # ── Step 4: Total = Jumlah × Harga Awal (Cell 4) ─────────
    df["Jumlah"]     = pd.to_numeric(_get_1d(df, "Jumlah"),     errors="coerce").fillna(0)
    df["Harga Awal"] = pd.to_numeric(_get_1d(df, "Harga Awal"), errors="coerce").fillna(0)
    df["Total Harga Produk"] = df["Jumlah"] * df["Harga Awal"]

    # ── Step 5: User ID ────────────────────────────────────────
    # Sesuai Cell 2: gunakan kombinasi kolom identitas
    id_cols = ["Username (Pembeli)", "Provinsi", "Kota/Kabupaten"]
    id_cols_ada = [c for c in id_cols if c in df.columns]
    if id_cols_ada:
        df_temp = df[id_cols_ada].fillna("KOSONG").astype(str)
        nomor_urut = df_temp.groupby(id_cols_ada, sort=False).ngroup() + 1
        df["User ID"] = "TikTok_" + nomor_urut.astype(str).str.zfill(5)
    else:
        df["User ID"] = _stable_uid(_get_1d(df, "Username (Pembeli)"), "TikTok")

    # ── Step 6: Groupby No. Pesanan (Cell 4) ─────────────────
    df = df.loc[:, ~df.columns.duplicated()]
    agg = {
        "Waktu Pesanan Dibuat": "first",
        "Username (Pembeli)":   "first",
        "Total Harga Produk":   "sum",
        "User ID":              "first",
    }
    for c in ["Kota/Kabupaten", "Provinsi"]:
        if c in df.columns: agg[c] = "first"
    agg = {k: v for k, v in agg.items() if k in df.columns}
    df = df.groupby("No. Pesanan", as_index=False).agg(agg)

    # ── Step 7: Clean (Cell 4) ────────────────────────────────
    df["Total Harga Produk"] = pd.to_numeric(
        _get_1d(df, "Total Harga Produk"), errors="coerce").fillna(0)
    df = df.dropna(subset=["Waktu Pesanan Dibuat", "Username (Pembeli)"])
    df = df[df["Total Harga Produk"] > 0]
    df["Kanal"] = "TikTok"

    return df.reset_index(drop=True)


# ══════════════════════════════════════════════════════════════
# MERGE  (sesuai Merged_2.ipynb)
# ══════════════════════════════════════════════════════════════
_REQ = ["No. Pesanan", "Waktu Pemesanan", "Username (Pembeli)",
        "User ID", "Total Harga Produk", "Kanal"]


def _to_gabungan_format(df: pd.DataFrame) -> pd.DataFrame:
    """
    Sesuai Merged_2.ipynb Cell 0 & 1:
    - Tambah Kanal jika belum ada
    - Waktu Pesanan Dibuat → Waktu Pemesanan (pd.to_datetime)
    - Drop Waktu Pesanan Dibuat
    - No. Pesanan as string
    - dropna(Waktu Pemesanan, Username)
    - drop Total == 0
    """
    df = df.copy()
    # Konversi waktu
    if "Waktu Pesanan Dibuat" in df.columns and "Waktu Pemesanan" not in df.columns:
        df["Waktu Pemesanan"] = pd.to_datetime(
            _get_1d(df, "Waktu Pesanan Dibuat"), errors="coerce")
        df = df.drop(columns=["Waktu Pesanan Dibuat"], errors="ignore")
    elif "Waktu Pemesanan" in df.columns:
        df["Waktu Pemesanan"] = pd.to_datetime(
            _get_1d(df, "Waktu Pemesanan"), errors="coerce")

    df["No. Pesanan"] = _get_1d(df, "No. Pesanan").astype(str).str.strip()
    for c in _REQ:
        if c not in df.columns: df[c] = None
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.dropna(subset=["Waktu Pemesanan", "Username (Pembeli)"])
    df = df[_get_1d(df, "Total Harga Produk").fillna(0) != 0]
    return df


def merge_to_gabungan(df_new, df_existing=None):
    """Merge df_new ke existing, dedup by No. Pesanan."""
    if df_new is None or df_new.empty:
        return (df_existing if df_existing is not None else pd.DataFrame()), 0, 0
    df = _to_gabungan_format(df_new)
    n_dup = 0
    if df_existing is not None and not df_existing.empty:
        df_existing = df_existing.copy()
        df_existing["No. Pesanan"] = _get_1d(df_existing, "No. Pesanan").astype(str).str.strip()
        # Pastikan Waktu Pemesanan bertipe datetime (data dari Sheets bisa berupa string)
        if "Waktu Pemesanan" in df_existing.columns:
            df_existing["Waktu Pemesanan"] = pd.to_datetime(df_existing["Waktu Pemesanan"], errors="coerce")
        already = _get_1d(df, "No. Pesanan").isin(df_existing["No. Pesanan"])
        n_dup = int(already.sum())
        df_clean = df[~already]
        n_added = len(df_clean)
        merged = pd.concat([df_existing, df_clean[_REQ]], ignore_index=True)
    else:
        n_added = len(df)
        merged = df[_REQ].copy()
    merged["Waktu Pemesanan"] = pd.to_datetime(merged["Waktu Pemesanan"], errors="coerce")
    merged = merged.dropna(subset=["Waktu Pemesanan"])
    merged = merged.sort_values("Waktu Pemesanan").reset_index(drop=True)
    return merged, n_added, n_dup


# ══════════════════════════════════════════════════════════════
# SCAN & REBUILD SEMUA FILE RAW
# ══════════════════════════════════════════════════════════════
def scan_and_build_gabungan(shopee_dir: str, tiktok_dir: str,
                             progress_cb=None) -> tuple:
    """
    Baca SEMUA file dari shopee_dir + tiktok_dir, preprocess, merge.
    Sesuai alur Shopee.ipynb + TikTok_2.ipynb + Merged_2.ipynb.

    Returns: (df_gabungan, log)
    """
    log = []
    all_files = []

    for f in sorted(Path(shopee_dir).glob("*.xls*")):
        if not f.name.startswith("~$"): all_files.append(("Shopee", f))
    for f in sorted(Path(tiktok_dir).glob("*.xls*")):
        if not f.name.startswith("~$"): all_files.append(("TikTok", f))

    if not all_files:
        return pd.DataFrame(), [{"File":"—","Kanal":"—","Baris":0,
                                  "Status":"Tidak ada file di folder raw"}]

    frames = []
    for i, (kanal, fpath) in enumerate(all_files):
        if progress_cb: progress_cb(i+1, len(all_files), f"{kanal}: {fpath.name}")
        try:
            with open(fpath, "rb") as fb: fbytes = fb.read()
            df = preprocess_shopee(fbytes, fpath.name) if kanal == "Shopee" \
                 else preprocess_tiktok(fbytes, fpath.name)
            if df is not None and not df.empty:
                frames.append(df)
                log.append({"File": fpath.name, "Kanal": kanal,
                             "Baris": len(df), "Status": "✅ OK"})
            else:
                log.append({"File": fpath.name, "Kanal": kanal,
                             "Baris": 0, "Status": "⚠️ 0 baris"})
        except Exception as e:
            log.append({"File": fpath.name, "Kanal": kanal,
                         "Baris": 0, "Status": f"❌ {str(e)[:100]}"})

    if not frames:
        return pd.DataFrame(), log

    # Sesuai Merged_2.ipynb: concat → format tanggal → clean
    df_all = pd.concat(frames, ignore_index=True)

    # ── BEFORE cleaning metrics (sesuai Merged_2.ipynb Cell 1) ─
    df_tmp = df_all.copy()
    if "Waktu Pesanan Dibuat" in df_tmp.columns and "Waktu Pemesanan" not in df_tmp.columns:
        df_tmp["Waktu Pemesanan"] = pd.to_datetime(df_tmp["Waktu Pesanan Dibuat"], errors="coerce")
    before_rows     = len(df_tmp)
    before_null_w   = int(df_tmp["Waktu Pemesanan"].isna().sum()) if "Waktu Pemesanan" in df_tmp.columns else 0
    before_null_u   = int(df_tmp["Username (Pembeli)"].isna().sum()) if "Username (Pembeli)" in df_tmp.columns else 0
    before_zero_h   = int((df_tmp["Total Harga Produk"] == 0).sum()) if "Total Harga Produk" in df_tmp.columns else 0
    before_dup      = int(df_tmp.duplicated(subset=["No. Pesanan"]).sum()) if "No. Pesanan" in df_tmp.columns else 0

    # ── Cleaning ──────────────────────────────────────────────
    df_gabungan = _to_gabungan_format(df_all)
    df_gabungan = df_gabungan.drop_duplicates(subset=["No. Pesanan"], keep="last")
    df_gabungan = df_gabungan.sort_values("Waktu Pemesanan").reset_index(drop=True)

    # ── AFTER cleaning metrics ────────────────────────────────
    after_rows = len(df_gabungan)
    cleaning_stats = pd.DataFrame({
        "Metrik":  ["Total Baris","Null Waktu","Null Username","Harga = 0","Duplikat No. Pesanan"],
        "Sebelum": [before_rows, before_null_w, before_null_u, before_zero_h, before_dup],
        "Sesudah": [after_rows,  0,             0,             0,            0],
        "Dihapus": [before_rows - after_rows,
                    before_null_w, before_null_u, before_zero_h, before_dup],
    })
    # Tambahkan cleaning_stats ke log sebagai entry khusus
    log.append({"File":"[CLEANING SUMMARY]","Kanal":"—",
                 "Baris":after_rows,"Status":
                 f"Before:{before_rows} → After:{after_rows} "
                 f"(-{before_rows-after_rows} rows)"})

    return df_gabungan[_REQ], log, cleaning_stats