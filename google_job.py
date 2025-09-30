import io, os, time, tempfile, re
from datetime import datetime, timezone
import requests
from PIL import Image
import fitz  # PyMuPDF

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from googleapiclient.errors import HttpError

# ====== CONFIG (env with sane defaults) ======
SHEET_ID          = os.environ.get("SHEET_ID", "").strip()
SHEET_NAME        = os.environ.get("SHEET_NAME", "Dispatch Details").strip()
DEST_FOLDER_ID    = os.environ.get("DEST_FOLDER_ID", "").strip()

START_ROW         = int(os.environ.get("START_ROW", "2"))
MAX_ROWS_TO_CHECK = int(os.environ.get("MAX_ROWS_TO_CHECK", "10000"))

# Columns (1-based)
COL_URL      = 9   # I: source PDF URL (Drive sharable or HTTP)
COL_INVOICE  = 7   # G: invoice number (filename)
COL_OUTPUT   = 12  # L: compressed view URL
COL_FLAG     = 13  # M: status/log

# Compression targets
MAX_TARGET_BYTES   = int(os.environ.get("MAX_TARGET_BYTES", str(1*1024*1024)))  # 1MB
TARGET_WIDTH_PT    = int(os.environ.get("TARGET_WIDTH_PT", "595"))  # A4 width
TARGET_HEIGHT_PT   = int(os.environ.get("TARGET_HEIGHT_PT", "842")) # A4 height

START_DPI          = int(os.environ.get("START_DPI", "150"))
MIN_DPI            = int(os.environ.get("MIN_DPI", "72"))
START_JPEG_QUALITY = int(os.environ.get("START_JPEG_QUALITY", "85"))
MIN_JPEG_QUALITY   = int(os.environ.get("MIN_JPEG_QUALITY", "30"))
DPI_STEP           = int(os.environ.get("DPI_STEP", "10"))
QUALITY_STEP       = int(os.environ.get("QUALITY_STEP", "5"))

DOWNLOAD_TIMEOUT   = int(os.environ.get("DOWNLOAD_TIMEOUT", "60"))

# ====== Auth (OAuth refresh token) ======
CLIENT_ID      = os.environ["GOOGLE_CLIENT_ID"]
CLIENT_SECRET  = os.environ["GOOGLE_CLIENT_SECRET"]
REFRESH_TOKEN  = os.environ["GOOGLE_REFRESH_TOKEN"]

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

def make_creds():
    return Credentials(
        token=None,
        refresh_token=REFRESH_TOKEN,
        token_uri="https://oauth2.googleapis.com/token",
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET,
        scopes=SCOPES,
    )

def col_letter(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s

# ====== Google services ======
def build_services():
    creds = make_creds()
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)
    return drive, sheets

# ====== Sheet helpers ======
def read_column(sheets, col_idx: int, start_row: int):
    rng = f"{SHEET_NAME}!{col_letter(col_idx)}{start_row}:{col_letter(col_idx)}{start_row+MAX_ROWS_TO_CHECK}"
    resp = sheets.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=rng).execute()
    vals = resp.get("values", [])
    return [(r[0].strip() if r and r[0] else "") for r in vals]

def read_cell(sheets, row: int, col_idx: int):
    rng = f"{SHEET_NAME}!{col_letter(col_idx)}{row}"
    resp = sheets.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=rng).execute()
    vals = resp.get("values", [])
    return vals[0][0] if vals and vals[0] else ""

def write_cell(sheets, row: int, col_idx: int, value: str):
    rng = f"{SHEET_NAME}!{col_letter(col_idx)}{row}"
    body = {"range": rng, "values": [[value]]}
    sheets.spreadsheets().values().update(
        spreadsheetId=SHEET_ID, range=rng,
        valueInputOption="RAW", body=body
    ).execute()

# ====== Download helpers (Drive or HTTP) ======
def is_drive_share_url(url: str) -> bool:
    return bool(url and "drive.google.com" in url)

def extract_drive_file_id(url: str):
    if not url: return None
    m = re.search(r"/file/d/([a-zA-Z0-9_-]{10,})", url)
    if m: return m.group(1)
    m = re.search(r"[?&]id=([a-zA-Z0-9_-]{10,})", url)
    if m: return m.group(1)
    m = re.search(r"open\\?id=([a-zA-Z0-9_-]{10,})", url)
    if m: return m.group(1)
    return None

def download_drive_file_by_id(drive, file_id: str, out_path: str) -> int:
    request = drive.files().get_media(fileId=file_id)
    with io.FileIO(out_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    return os.path.getsize(out_path)

def download_url_to_file(drive, url: str, out_path: str, timeout=60) -> int:
    if is_drive_share_url(url):
        fid = extract_drive_file_id(url)
        if fid:
            try:
                return download_drive_file_by_id(drive, fid, out_path)
            except Exception as e:
                print("Drive API download failed, fallback to HTTP:", e)
    headers = {"User-Agent": "Mozilla/5.0 (compatible)"}
    with requests.get(url, stream=True, headers=headers, timeout=timeout) as r:
        r.raise_for_status()
        with open(out_path, "wb") as f:
            for chunk in r.iter_content(1024 * 64):
                if chunk:
                    f.write(chunk)
    return os.path.getsize(out_path)

# ====== Render & compose to A4 ======
def render_pages_to_images(input_pdf_path: str, dpi: int):
    doc = fitz.open(input_pdf_path)
    images = []
    try:
        for p in range(len(doc)):
            page = doc.load_page(p)
            mat = fitz.Matrix(dpi/72.0, dpi/72.0)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            images.append(img)
    finally:
        doc.close()
    return images

def compose_images_to_target_size(images, target_w_pt, target_h_pt, dpi, jpeg_quality):
    target_w_px = int(round(target_w_pt * dpi / 72.0))
    target_h_px = int(round(target_h_pt * dpi / 72.0))
    pages = []
    for img in images:
        iw, ih = img.size
        ratio = min(target_w_px/iw, target_h_px/ih)
        new_w, new_h = int(round(iw*ratio)), int(round(ih*ratio))
        resized = img.resize((new_w, new_h), resample=Image.LANCZOS)
        canvas = Image.new("RGB", (target_w_px, target_h_px), (255, 255, 255))
        left = (target_w_px - new_w)//2
        top  = (target_h_px - new_h)//2
        canvas.paste(resized, (left, top))
        pages.append(canvas)

    bio = io.BytesIO()
    pages[0].save(
        bio,
        save_all=True,
        append_images=pages[1:],
        format="PDF",
        quality=jpeg_quality,
        optimize=True,
    )
    bio.seek(0)
    return bio.getvalue()

def iterative_compress_to_limit(input_pdf_path: str):
    dpi = START_DPI
    quality = START_JPEG_QUALITY
    while True:
        imgs = render_pages_to_images(input_pdf_path, dpi)
        pdf_bytes = compose_images_to_target_size(imgs, TARGET_WIDTH_PT, TARGET_HEIGHT_PT, dpi, quality)
        size = len(pdf_bytes)
        print(f"  try dpi={dpi} q={quality} -> {size} bytes")
        if size <= MAX_TARGET_BYTES:
            return pdf_bytes, size, dpi, quality
        if quality - QUALITY_STEP >= MIN_JPEG_QUALITY:
            quality -= QUALITY_STEP
            continue
        if dpi - DPI_STEP >= MIN_DPI:
            dpi -= DPI_STEP
            quality = START_JPEG_QUALITY
            continue
        return pdf_bytes, size, dpi, quality  # best effort

# ====== Drive upload ======
def safe_filename(s: str) -> str:
    s = (s or "").strip()
    if not s:
        s = f"pdf_{int(time.time())}"
    s = re.sub(r'[\\/*?:"<>|]', "_", s)
    if not s.lower().endswith(".pdf"):
        s += ".pdf"
    return s

def upload_pdf_bytes_to_drive(drive, pdf_bytes: bytes, filename: str, folder_id: str):
    media = MediaIoBaseUpload(io.BytesIO(pdf_bytes), mimetype="application/pdf", resumable=True)
    body = {"name": filename, "parents": [folder_id]} if folder_id else {"name": filename}
    f = drive.files().create(body=body, media_body=media, fields="id,size").execute()
    return f.get("id"), int(f.get("size", 0) or 0)

def set_public_anyone(drive, file_id: str):
    try:
        drive.permissions().create(fileId=file_id, body={"role": "reader", "type": "anyone"}).execute()
    except Exception as e:
        print("Warning: set public failed:", e)

# ====== Main runner ======
def run():
    if not SHEET_ID or not DEST_FOLDER_ID:
        raise RuntimeError("SHEET_ID and DEST_FOLDER_ID must be set as repo Variables.")

    drive, sheets = build_services()

    urls     = read_column(sheets, COL_URL, START_ROW)
    invoices = read_column(sheets, COL_INVOICE, START_ROW)
    rows_count = max(len(urls), len(invoices))
    print(f"Found up to {rows_count} rows starting {START_ROW}")

    for idx in range(rows_count):
        row = START_ROW + idx
        url = (urls[idx] if idx < len(urls) else "").strip()
        inv = (invoices[idx] if idx < len(invoices) else "").strip()
        if not url:
            continue

        # Skip if L already filled
        try:
            if read_cell(sheets, row, COL_OUTPUT):
                continue
        except Exception as e:
            print("Warning reading existing L:", e)

        print(f"\nRow {row}: download -> {url} | invoice -> {inv or '[no name]'}")
        tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf"); tmp_in.close()

        try:
            size = download_url_to_file(drive, url, tmp_in.name, timeout=DOWNLOAD_TIMEOUT)
            print(f" Downloaded {size} bytes")

            print(f" Compressing to ≤ {MAX_TARGET_BYTES} bytes (A4)…")
            pdf_bytes, final_size, used_dpi, used_q = iterative_compress_to_limit(tmp_in.name)
            print(f" Result size={final_size} (dpi={used_dpi}, q={used_q})")

            filename = safe_filename(inv) if inv else safe_filename(os.path.basename(url.split('?')[0]))
            file_id, uploaded_size = upload_pdf_bytes_to_drive(drive, pdf_bytes, filename, DEST_FOLDER_ID)
            print(f" Uploaded id={file_id}, size={uploaded_size}")

            set_public_anyone(drive, file_id)
            view_url = f"https://drive.google.com/uc?export=view&id={file_id}"
            flag = "COMPRESSED" if uploaded_size <= MAX_TARGET_BYTES else "LARGE_FILE"

            write_cell(sheets, row, COL_OUTPUT, view_url)  # L
            write_cell(sheets, row, COL_FLAG, f"{flag} dpi={used_dpi} q={used_q} size={uploaded_size}")  # M
            print(f"Row {row}: done.")

        except Exception as e:
            print("Row error:", e)
            try:
                write_cell(sheets, row, COL_FLAG, f"ERROR: {str(e)[:250]}")
            except Exception as ee:
                print("Also failed to write error:", ee)
        finally:
            try:
                if os.path.exists(tmp_in.name):
                    os.remove(tmp_in.name)
            except:
                pass

    print("\nAll rows processed.")

if __name__ == "__main__":
    try:
        run()
    except HttpError as e:
        print("[FATAL] Google API error:", e)
        print("Ensure the SHEET and DEST_FOLDER are accessible by the SAME Google account used for OAuth.")
        raise
