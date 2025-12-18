import os
import json
import uuid
import re
from pathlib import Path
from flask import Flask, request, jsonify, render_template, send_from_directory, session
import requests
from dotenv import load_dotenv
from datetime import datetime
import shutil
import zipfile
import subprocess
import platform
import sqlite3

# ========== PDF GENERATION - MULTIPLE METHODS ==========
PDF_AVAILABLE = False
PDF_METHOD = None

# Try docx2pdf (Windows/macOS)
try:
    from docx2pdf import convert as docx_to_pdf
    PDF_AVAILABLE = True
    PDF_METHOD = "docx2pdf"
except ImportError:
    docx_to_pdf = None

# Try pypandoc (Alternative)
try:
    import pypandoc
    if not PDF_AVAILABLE:
        PDF_AVAILABLE = True
        PDF_METHOD = "pypandoc"
except ImportError:
    pypandoc = None

# LibreOffice detection - IMPROVED VERSION
def check_libreoffice():
    """Check if LibreOffice is available - Enhanced for Linux containers"""
    try:
        system = platform.system()
        print(f"🔍 Checking LibreOffice on {system}...")

        if system == "Windows":
            paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            ]
            for p in paths:
                if os.path.exists(p):
                    print(f"✅ Found LibreOffice at: {p}")
                    return p
        else:
            commands = ['libreoffice', 'soffice', '/usr/bin/libreoffice', '/usr/bin/soffice']
            for cmd in commands:
                try:
                    if os.path.exists(cmd) and os.access(cmd, os.X_OK):
                        print(f"✅ Found LibreOffice at: {cmd}")
                        return cmd

                    result = subprocess.run(
                        ['which', cmd],
                        capture_output=True,
                        text=True,
                        timeout=5
                    )
                    if result.returncode == 0 and result.stdout.strip():
                        path = result.stdout.strip()
                        print(f"✅ Found LibreOffice at: {path}")
                        return path
                except Exception:
                    continue

            for cmd in ['libreoffice', 'soffice']:
                try:
                    result = subprocess.run(
                        ['command', '-v', cmd],
                        capture_output=True,
                        text=True,
                        timeout=5,
                        shell=True
                    )
                    if result.returncode == 0 and result.stdout.strip():
                        path = result.stdout.strip()
                        print(f"✅ Found LibreOffice at: {path}")
                        return path
                except:
                    continue

        print("❌ LibreOffice not found")
    except Exception as e:
        print(f"❌ Error checking LibreOffice: {e}")

    return None

LIBREOFFICE_PATH = check_libreoffice()
if LIBREOFFICE_PATH and not PDF_AVAILABLE:
    PDF_AVAILABLE = True
    PDF_METHOD = "libreoffice"

load_dotenv()

OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "openai/gpt-4o-mini")

SERPER_API_KEY = os.getenv("SERPER_API_KEY")

BASE_DIR = Path(__file__).parent
FILES_DIR = BASE_DIR / "static" / "files"
TEMPLATE_FILE = BASE_DIR / "template_quotation.docx"
TEMP_DIR = BASE_DIR / "temp"
COUNTER_FILE = BASE_DIR / "counter.json"
DB_FILE = BASE_DIR / "chat_history.db"

FILES_DIR.mkdir(parents=True, exist_ok=True)
TEMP_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = "karya-limbah-2025"

conversations = {}

# =========================
# ✅ SIMPLE CORS (React Native fetch)
# =========================
@app.after_request
def add_cors_headers(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, DELETE, OPTIONS"
    return resp

# =========================
# ✅ DATABASE (SQLite) INIT + HELPERS
# =========================
def db_connect():
    conn = sqlite3.connect(str(DB_FILE))
    conn.row_factory = sqlite3.Row
    return conn

def _db_has_column(conn, table: str, column: str) -> bool:
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({table})")
    cols = [r[1] for r in cur.fetchall()]
    return column in cols

def init_db():
    conn = db_connect()
    cur = conn.cursor()

    # base table
    cur.execute("""
        CREATE TABLE IF NOT EXISTS chat_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            task_type TEXT NOT NULL,
            created_at TEXT NOT NULL,
            data_json TEXT NOT NULL,
            files_json TEXT NOT NULL
        )
    """)
    conn.commit()

    # ✅ MIGRATION: add messages_json + state_json if missing
    if not _db_has_column(conn, "chat_history", "messages_json"):
        cur.execute("ALTER TABLE chat_history ADD COLUMN messages_json TEXT NOT NULL DEFAULT '[]'")
    if not _db_has_column(conn, "chat_history", "state_json"):
        cur.execute("ALTER TABLE chat_history ADD COLUMN state_json TEXT NOT NULL DEFAULT '{}'")

    conn.commit()
    conn.close()

def db_insert_history(title: str, task_type: str, data: dict, files: list, messages: list = None, state: dict = None):
    conn = db_connect()
    cur = conn.cursor()
    created_at = datetime.now().isoformat()
    cur.execute("""
        INSERT INTO chat_history (title, task_type, created_at, data_json, files_json, messages_json, state_json)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        title,
        task_type,
        created_at,
        json.dumps(data or {}, ensure_ascii=False),
        json.dumps(files or [], ensure_ascii=False),
        json.dumps(messages or [], ensure_ascii=False),
        json.dumps(state or {}, ensure_ascii=False),
    ))
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return new_id

def db_list_histories(limit=200, q: str = None):
    conn = db_connect()
    cur = conn.cursor()
    if q:
        cur.execute("""
            SELECT id, title, task_type, created_at
            FROM chat_history
            WHERE title LIKE ?
            ORDER BY id DESC
            LIMIT ?
        """, (f"%{q}%", limit))
    else:
        cur.execute("""
            SELECT id, title, task_type, created_at
            FROM chat_history
            ORDER BY id DESC
            LIMIT ?
        """, (limit,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def db_get_history_detail(history_id: int):
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("""
        SELECT id, title, task_type, created_at, data_json, files_json, messages_json, state_json
        FROM chat_history
        WHERE id = ?
    """, (history_id,))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None

def db_update_title(history_id: int, new_title: str):
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("UPDATE chat_history SET title = ? WHERE id = ?", (new_title, history_id))
    conn.commit()
    changes = cur.rowcount
    conn.close()
    return changes > 0

def db_delete_history(history_id: int):
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("DELETE FROM chat_history WHERE id = ?", (history_id,))
    conn.commit()
    changes = cur.rowcount
    conn.close()
    return changes > 0

def db_append_message(history_id: int, sender: str, text: str, files: list = None):
    detail = db_get_history_detail(history_id)
    if not detail:
        return False

    try:
        messages = json.loads(detail.get("messages_json") or "[]")
    except:
        messages = []

    messages.append({
        "id": uuid.uuid4().hex[:12],
        "sender": sender,
        "text": text,
        "files": files or [],
        "timestamp": datetime.now().isoformat()
    })

    conn = db_connect()
    cur = conn.cursor()
    cur.execute("UPDATE chat_history SET messages_json = ? WHERE id = ?", (json.dumps(messages, ensure_ascii=False), history_id))
    conn.commit()
    ok = cur.rowcount > 0
    conn.close()
    return ok

def db_update_state(history_id: int, state: dict):
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("UPDATE chat_history SET state_json = ? WHERE id = ?", (json.dumps(state or {}, ensure_ascii=False), history_id))
    conn.commit()
    conn.close()

init_db()

# ===== COUNTER MANAGEMENT =====
def load_counter():
    """Load counter dari file"""
    if COUNTER_FILE.exists():
        try:
            with open(COUNTER_FILE, 'r') as f:
                data = json.load(f)
                return int(data.get('counter', 0))
        except:
            return 0
    return 0

def save_counter(counter):
    """Save counter ke file"""
    with open(COUNTER_FILE, 'w') as f:
        json.dump({'counter': int(counter)}, f)

def get_next_nomor():
    """Generate nomor depan otomatis (000..021 lalu reset ke 000)"""
    counter = load_counter()
    try:
        counter = int(counter)
    except:
        counter = 0

    if counter < 0:
        counter = 0
    if counter > 21:
        counter = 0

    nomor = str(counter).zfill(3)  # 000..021
    next_counter = (counter + 1) % 22
    save_counter(next_counter)
    return nomor

# ===== DATABASE LIMBAH B3 DARI PDF =====
LIMBAH_B3_DB = {
    "A102d": {"jenis": "Aki/baterai bekas", "satuan": "Kg", "karakteristik": "Beracun / Korosif"},
    "A103d": {"jenis": "Debu dan fiber asbes (crocidolite, amosite, janthrophyllite)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A106d": {"jenis": "Limbah dari laboratorium yang mengandung B3", "satuan": "Kg", "karakteristik": "Beracun"},
    "A107d": {"jenis": "Pelarut bekas lainnya yang belum dikodifikasi", "satuan": "Kg", "karakteristik": "Beracun"},
    "A108d": {"jenis": "Limbah terkontaminasi B3", "satuan": "Kg", "karakteristik": "Padatan Mudah Menyala"},
    "A111d": {"jenis": "Refrigerant bekas dari peralatan elektronik", "satuan": "Kg", "karakteristik": "Beracun"},
    "A303-2": {"jenis": "Residu proses produksi (Pestisida dan produk agrokimia)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A303-3": {"jenis": "Absorben dan filter bekas", "satuan": "Kg", "karakteristik": "Beracun"},
    "A303-6": {"jenis": "Sludge IPAL", "satuan": "Kg", "karakteristik": "Beracun"},
    "A304-1": {"jenis": "Resin adesif Fenolformaldehida (PF, UF, MF)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A304-2": {"jenis": "Lumpur encer mengandung adesif atau sealant", "satuan": "Kg", "karakteristik": "Beracun"},
    "A304-3": {"jenis": "Limbah minyak resin (terpentin)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A304-4": {"jenis": "Residu dari proses penyaringan produk (strainer)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A304-6": {"jenis": "Residu proses produksi atau kegiatan", "satuan": "Kg", "karakteristik": "Beracun"},
    "A305-1": {"jenis": "Monomer atau oligomer yang tidak bereaksi (Polimer)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A305-2": {"jenis": "Residu produksi atau reaksi pemurnian polimer", "satuan": "Kg", "karakteristik": "Beracun"},
    "A305-3": {"jenis": "Residu dari proses destilasi", "satuan": "Kg", "karakteristik": "Beracun"},
    "A306-1": {"jenis": "Sludge dari proses produksi minyak bumi/gas alam (Petrokimia)", "satuan": "Kg", "karakteristik": "Padatan Mudah Menyala"},
    "A307-1": {"jenis": "Sludge dari kilang minyak dan gas bumi", "satuan": "Kg", "karakteristik": "Padatan Mudah Menyala"},
    "A307-2": {"jenis": "Residu dasar tanki", "satuan": "Kg", "karakteristik": "Padatan Mudah Menyala"},
    "A307-3": {"jenis": "Slop padatan emulsi minyak dari penyulingan minyak bumi", "satuan": "Kg", "karakteristik": "Padatan Mudah Menyala"},
    "A309-1": {"jenis": "Fluxing agent bekas (Peleburan besi dan baja)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A309-3": {"jenis": "Spent pickle liquor", "satuan": "Kg", "karakteristik": "Beracun"},
    "A309-6": {"jenis": "Residu dari proses produksi kokas (tar)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A310-1": {"jenis": "Larutan asam, alkali bekas (Operasi penyempurnaan baja)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A310-5": {"jenis": "Sludge dari proses pengolahan residu", "satuan": "Kg", "karakteristik": "Beracun"},
    "A311-1": {"jenis": "Larutan asam bekas (Peleburan timah hitam Pb)", "satuan": "Kg", "karakteristik": "Korosif"},
    "A311-2": {"jenis": "Slag dari proses peleburan primer/sekunder", "satuan": "Kg", "karakteristik": "Korosif"},
    "A311-4": {"jenis": "Ash, dross, skimming dari peleburan primer/sekunder", "satuan": "Kg", "karakteristik": "Beracun"},
    "A312-4": {"jenis": "Sludge dari oil treatment (Peleburan tembaga Cu)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A313-4": {"jenis": "Sludge dari oil treatment (Peleburan alumunium)", "satuan": "Kg", "karakteristik": "Padatan Mudah Menyala"},
    "A314-2": {"jenis": "Sludge dari oil treatment (Peleburan seng Zn)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A319-1": {"jenis": "Sludge dari oil treatment (Peleburan timah putih Sn)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A322-1": {"jenis": "Pelarut bekas (cleaning) Tekstil", "satuan": "Kg", "karakteristik": "Beracun"},
    "A322-2": {"jenis": "Senyawa brom organik (fire retardant)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A322-3": {"jenis": "Dyestuffs dan pigment mengandung logam berat", "satuan": "Kg", "karakteristik": "Beracun"},
    "A323-1": {"jenis": "Pelarut bekas pencucian (Manufaktur kendaraan)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A323-2": {"jenis": "Sludge proses produksi manufacturing", "satuan": "Kg", "karakteristik": "Beracun"},
    "A323-3": {"jenis": "Residu proses produksi manufacturing", "satuan": "Kg", "karakteristik": "Beracun"},
    "A324-2": {"jenis": "Larutan bekas dari kegiatan pengolahan (Elektroplating)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A324-3": {"jenis": "Larutan asam (pickling)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A324-8": {"jenis": "Spent plating solutions (Cr, Pb, Ni, As, Cu, Zn, Cd, Fe, Sn)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A325-1": {"jenis": "Limbah cat dan varnish mengandung pelarut organik", "satuan": "Kg", "karakteristik": "Beracun"},
    "A325-2": {"jenis": "Sludge dari cat dan varnish", "satuan": "Kg", "karakteristik": "Beracun"},
    "A325-3": {"jenis": "Residu proses destilasi cat", "satuan": "Kg", "karakteristik": "Beracun"},
    "A325-4": {"jenis": "Cat anti korosi berbahan Pb dan Cr", "satuan": "Kg", "karakteristik": "Beracun"},
    "A325-5": {"jenis": "Debu/sludge dari unit pengendalian pencemaran udara", "satuan": "Kg", "karakteristik": "Beracun"},
    "A325-6": {"jenis": "Sludge proses depainting", "satuan": "Kg", "karakteristik": "Beracun"},
    "A325-7": {"jenis": "Sludge dari IPAL cat", "satuan": "Kg", "karakteristik": "Beracun"},
    "A330-1": {"jenis": "Residu dasar tangki minyak bumi", "satuan": "Kg", "karakteristik": "Padatan Mudah Menyala"},
    "A331-2": {"jenis": "Sludge dari oil treatment (Pertambangan)", "satuan": "Kg", "karakteristik": "Padatan Mudah Menyala"},
    "A332-1": {"jenis": "Sludge dari oil treatment (Industri listrik)", "satuan": "Kg", "karakteristik": "Padatan Mudah Menyala"},
    "A336-1": {"jenis": "Bahan/Produk farmasi tidak memenuhi spesifikasi", "satuan": "Kg", "karakteristik": "Beracun"},
    "A336-2": {"jenis": "Residu proses produksi dan formulasi farmasi", "satuan": "Kg", "karakteristik": "Beracun"},
    "A336-3": {"jenis": "Residu proses destilasi, evaporasi dan reaksi farmasi", "satuan": "Kg", "karakteristik": "Beracun"},
    "A336-4": {"jenis": "Reactor bottom wastes farmasi", "satuan": "Kg", "karakteristik": "Beracun"},
    "A336-5": {"jenis": "Sludge dari fasilitas produksi farmasi", "satuan": "Kg", "karakteristik": "Beracun"},
    "A337-1": {"jenis": "Limbah klinis memiliki karakteristik infeksius", "satuan": "Kg", "karakteristik": "Infeksius"},
    "A337-2": {"jenis": "Produk farmasi kedaluwarsa", "satuan": "Kg", "karakteristik": "Beracun"},
    "A337-3": {"jenis": "Bahan kimia kedaluwarsa rumah sakit", "satuan": "Kg", "karakteristik": "Beracun"},
    "A337-4": {"jenis": "Peralatan laboratorium terkontaminasi B3", "satuan": "Kg", "karakteristik": "Beracun"},
    "A338-1": {"jenis": "Bahan kimia kedaluwarsa laboratorium", "satuan": "Kg", "karakteristik": "Beracun"},
    "A338-2": {"jenis": "Peralatan laboratorium terkontaminasi B3", "satuan": "Kg", "karakteristik": "Beracun"},
    "A338-3": {"jenis": "Residu sampel Limbah B3", "satuan": "Kg", "karakteristik": "Beracun"},
    "A338-4": {"jenis": "Sludge IPAL laboratorium", "satuan": "Kg", "karakteristik": "Beracun"},
    "A339-1": {"jenis": "Larutan developer, fixer, bleach bekas (Fotografi)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A341-1": {"jenis": "Residu proses produksi dan konsentrat (Sabun deterjen, kosmetik)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A341-2": {"jenis": "Konsentrat tidak memenuhi spesifikasi teknis", "satuan": "Kg", "karakteristik": "Beracun"},
    "A341-3": {"jenis": "Heavy alkylated hydrocarbon", "satuan": "Kg", "karakteristik": "Beracun"},
    "A342-1": {"jenis": "Residu filtrasi (Pengolahan minyak hewani/nabati)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A342-2": {"jenis": "Residu proses destilasi minyak", "satuan": "Kg", "karakteristik": "Beracun"},
    "A343-1": {"jenis": "Glycerine pitch (Pengolahan oleokimia dasar)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A343-2": {"jenis": "Residu filtrasi oleokimia", "satuan": "Kg", "karakteristik": "Beracun"},
    "A345-1": {"jenis": "Emulsi minyak dari proses cutting dan pendingin", "satuan": "Kg", "karakteristik": "Beracun"},
    "A345-2": {"jenis": "Sludge logam (serbuk, gram) mengandung minyak", "satuan": "Kg", "karakteristik": "Beracun"},
    "A350-2": {"jenis": "Adhesive coating (Seal, Gasket, Packing)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A351-1": {"jenis": "Adesif atau perekat sisa dan kedaluwarsa (Pulp dan kertas)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A351-2": {"jenis": "Residu pencetakan (tinta/pewarna)", "satuan": "Kg", "karakteristik": "Beracun"},
    "A355-1": {"jenis": "Pelarut (cleaning, degreasing) Bengkel kendaraan", "satuan": "Kg", "karakteristik": "Beracun"},
    "B102d": {"jenis": "Debu dan fiber asbes putih (chrysotile)", "satuan": "Kg", "karakteristik": "Beracun"},
    "B103d": {"jenis": "Lead scrap", "satuan": "Kg", "karakteristik": "Korosif, Beracun"},
    "B104d": {"jenis": "Kemasan bekas B3", "satuan": "Kg", "karakteristik": "Beracun"},
    "B105d": {"jenis": "Minyak pelumas bekas (hidrolik, mesin, gear, lubrikasi)", "satuan": "Kg", "karakteristik": "Cairan Mudah Menyala"},
    "B106d": {"jenis": "Limbah resin atau penukar ion", "satuan": "Kg", "karakteristik": "Beracun"},
    "B107d": {"jenis": "Limbah elektronik (CRT, lampu TL, PCB, kawat logam)", "satuan": "Kg", "karakteristik": "Beracun"},
    "B108d": {"jenis": "Sludge IPAL dari fasilitas IPAL terpadu kawasan industri", "satuan": "Kg", "karakteristik": "Beracun"},
    "B109": {"jenis": "Filter bekas dari fasilitas pengendalian pencemaran udara", "satuan": "Kg", "karakteristik": "Beracun"},
    "B110d": {"jenis": "Kain majun bekas (used rags) dan yang sejenis", "satuan": "Kg", "karakteristik": "Padatan Mudah Menyala"},
}

def find_limbah_by_jenis(jenis_query):
    jenis_lower = jenis_query.lower()

    for kode, data in LIMBAH_B3_DB.items():
        if data['jenis'].lower() == jenis_lower:
            return kode, data

    for kode, data in LIMBAH_B3_DB.items():
        if jenis_lower in data['jenis'].lower() or data['jenis'].lower() in jenis_lower:
            return kode, data

    keywords = jenis_lower.split()
    for kode, data in LIMBAH_B3_DB.items():
        jenis_db_lower = data['jenis'].lower()
        match_count = sum(1 for kw in keywords if kw in jenis_db_lower)
        if match_count >= 2:
            return kode, data

    return None, None

def normalize_limbah_code(text):
    text_clean = text.strip().upper()
    strip_words = ['STRIP', 'MINUS', 'MIN', 'DASH', 'SAMPAI', 'HINGGA', 'GARIS']
    for word in strip_words:
        text_clean = re.sub(r'\b' + word + r'\b', '-', text_clean, flags=re.IGNORECASE)
    text_clean = re.sub(r'\s*-\s*', '-', text_clean)
    text_clean = re.sub(r'\s+', '', text_clean)
    return text_clean

def find_limbah_by_kode(kode_query):
    kode_normalized = normalize_limbah_code(kode_query)

    if kode_normalized in LIMBAH_B3_DB:
        return kode_normalized, LIMBAH_B3_DB[kode_normalized]

    kode_lower = kode_normalized.lower()
    if kode_lower in LIMBAH_B3_DB:
        return kode_lower, LIMBAH_B3_DB[kode_lower]

    if not kode_normalized.endswith('d') and not kode_normalized.endswith('D'):
        kode_with_d = kode_normalized + 'd'
        if kode_with_d in LIMBAH_B3_DB:
            return kode_with_d, LIMBAH_B3_DB[kode_with_d]

    return None, None

def angka_ke_romawi(bulan):
    romawi = {
        '1': 'I', '2': 'II', '3': 'III', '4': 'IV', '5': 'V', '6': 'VI',
        '7': 'VII', '8': 'VIII', '9': 'IX', '10': 'X', '11': 'XI', '12': 'XII',
        '01': 'I', '02': 'II', '03': 'III', '04': 'IV', '05': 'V', '06': 'VI',
        '07': 'VII', '08': 'VIII', '09': 'IX'
    }
    return romawi.get(str(bulan), 'I')

def angka_ke_terbilang(angka):
    try:
        n = int(angka)
    except:
        return 'empat belas'

    if n == 0:
        return 'nol'

    satuan = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh', 'delapan', 'sembilan']
    belasan = ['sepuluh', 'sebelas', 'dua belas', 'tiga belas', 'empat belas', 'lima belas',
               'enam belas', 'tujuh belas', 'delapan belas', 'sembilan belas']

    if n < 10:
        return satuan[n]
    elif n < 20:
        return belasan[n - 10]
    elif n < 100:
        puluhan = n // 10
        sisanya = n % 10
        if sisanya == 0:
            return satuan[puluhan] + ' puluh'
        else:
            return satuan[puluhan] + ' puluh ' + satuan[sisanya]
    elif n < 1000:
        ratusan = n // 100
        sisanya = n % 100
        if ratusan == 1:
            result = 'seratus'
        else:
            result = satuan[ratusan] + ' ratus'
        if sisanya > 0:
            result += ' ' + angka_ke_terbilang(sisanya)
        return result

    return str(n)

def convert_voice_to_number(text):
    text_lower = text.lower().strip()

    if re.match(r'^\d+$', text_lower.replace('.', '').replace(',', '')):
        return text_lower.replace('.', '').replace(',', '')

    kata_angka = {
        'nol': 0, 'kosong': 0,
        'satu': 1, 'se': 1,
        'dua': 2, 'tiga': 3, 'empat': 4,
        'lima': 5, 'enam': 6, 'tujuh': 7,
        'delapan': 8, 'sembilan': 9,
        'sepuluh': 10, 'sebelas': 11,
    }

    multipliers = {
        'belas': 10,
        'puluh': 10,
        'ratus': 100,
        'ribu': 1000,
        'juta': 1000000,
        'miliar': 1000000000,
        'milyar': 1000000000
    }

    words = text_lower.split()
    result = 0
    temp = 0

    for word in words:
        if word in kata_angka:
            temp += kata_angka[word]
        elif word == 'belas':
            temp += 10
        elif word == 'puluh':
            temp = (temp if temp > 0 else 1) * 10
        elif word == 'ratus':
            temp = (temp if temp > 0 else 1) * 100
        elif word == 'ribu':
            temp = (temp if temp > 0 else 1) * 1000
            result += temp
            temp = 0
        elif word in ['juta', 'miliar', 'milyar']:
            temp = (temp if temp > 0 else 1) * multipliers[word]
            result += temp
            temp = 0

    result += temp

    if result > 0:
        return str(result)

    return text

# ✅ khusus termin: ambil angka pertama saja (jadi input "14 200" -> 14)
def parse_termin_days(text: str, default: int = 14, min_days: int = 1, max_days: int = 365) -> str:
    s = (text or "").strip()
    if not s:
        return str(default)

    converted = convert_voice_to_number(s)
    m = re.search(r'\d+', str(converted).replace('.', '').replace(',', ''))
    if not m:
        return str(default)

    val = int(m.group(0))
    if val < min_days or val > max_days:
        return str(default)
    return str(val)

def call_ai(text, system_prompt=None):
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {"Authorization": f"Bearer {OPENROUTER_API_KEY}", "Content-Type": "application/json"}
    messages = []
    if system_prompt:
        messages.append({"role": "system", "content": system_prompt})
    messages.append({"role": "user", "content": text})

    resp = requests.post(url, headers=headers, json={
        "model": OPENROUTER_MODEL,
        "messages": messages,
        "temperature": 0.3,
        "max_tokens": 2000
    }, timeout=60)
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]

# =========================
# ✅ SEARCH ALAMAT VIA SERPER
# =========================
def _clean_address(addr: str) -> str:
    if not addr:
        return ""
    addr = re.sub(r'\s+', ' ', addr).strip()
    addr = addr.strip(' ,.-')
    return addr

def _extract_address_from_text(text: str) -> str:
    if not text:
        return ""
    t = text.replace('\n', ' ')
    t = re.sub(r'\s+', ' ', t)

    patterns = [
        r'(Jl\.?\s[^.,]{5,120}(?:No\.?\s?\d+[A-Za-z\/\-]?)?[^.]{0,120}(?:Jakarta|Bandung|Surabaya|Bekasi|Tangerang|Depok|Bogor|Medan|Semarang|Denpasar|Makassar)[^.,]{0,80})',
        r'(Rukan[^.,]{5,160}(?:Jakarta|Bekasi|Tangerang)[^.,]{0,80})',
        r'(Komplek[^.,]{5,160}(?:Jakarta|Bekasi|Tangerang)[^.,]{0,80})',
        r'(Kawasan[^.,]{5,160}(?:Jakarta|Bekasi|Tangerang)[^.,]{0,80})',
    ]
    for p in patterns:
        m = re.search(p, t, re.IGNORECASE)
        if m:
            return _clean_address(m.group(1))
    return ""

def search_company_address(company_name: str) -> str:
    name = (company_name or "").strip()
    if len(name) < 3:
        return ""

    if not SERPER_API_KEY:
        print("SERPER_API_KEY belum diset")
        return ""

    try:
        url = "https://google.serper.dev/search"
        headers = {"X-API-KEY": SERPER_API_KEY, "Content-Type": "application/json"}

        queries = [f"{name} alamat", f"\"{name}\" alamat kantor"]
        for q in queries:
            payload = {"q": q, "gl": "id", "hl": "id", "num": 7}
            r = requests.post(url, headers=headers, json=payload, timeout=30)
            r.raise_for_status()
            data = r.json()

            kg = data.get("knowledgeGraph") or {}
            for key in ["address", "formattedAddress"]:
                if isinstance(kg.get(key), str):
                    addr = _clean_address(kg.get(key))
                    if len(addr) >= 10:
                        return addr

            addr_obj = kg.get("address")
            if isinstance(addr_obj, dict):
                parts = []
                for k in ["streetAddress", "addressLocality", "addressRegion", "postalCode", "addressCountry"]:
                    v = addr_obj.get(k)
                    if isinstance(v, str) and v.strip():
                        parts.append(v.strip())
                addr = _clean_address(", ".join(parts))
                if len(addr) >= 10:
                    return addr

            places = data.get("places") or data.get("local") or []
            if isinstance(places, list) and places:
                p0 = places[0] or {}
                if isinstance(p0, dict):
                    addr = _clean_address(p0.get("address") or p0.get("formattedAddress") or "")
                    if len(addr) >= 10:
                        return addr

            organic = data.get("organic") or []
            for item in organic:
                if not isinstance(item, dict):
                    continue
                snippet = item.get("snippet") or ""
                title = item.get("title") or ""
                addr = _extract_address_from_text(snippet) or _extract_address_from_text(title)
                if len(addr) >= 10:
                    return addr

        return ""
    except Exception as e:
        print(f"Error searching address (Serper): {e}")
        return ""

def search_company_address_ai(company_name: str) -> str:
    try:
        name = (company_name or "").strip()
        if len(name) < 3 or not OPENROUTER_API_KEY:
            return ""

        system_prompt = (
            "Tulis alamat lengkap kantor pusat perusahaan di Indonesia jika Anda yakin.\n"
            "Jika tidak yakin, balas kosong.\n"
            "Jawaban hanya alamat satu baris."
        )
        out = call_ai(name, system_prompt=system_prompt).strip()
        out = _clean_address(out)
        return out if len(out) >= 10 else ""
    except Exception as e:
        print(f"AI address fallback error: {e}")
        return ""

def format_rupiah(angka_str):
    angka_clean = re.sub(r'[^\d]', '', str(angka_str))
    if not angka_clean:
        return angka_str
    try:
        angka_int = int(angka_clean)
        formatted = f"{angka_int:,}".replace(',', '.')
        return formatted
    except:
        return angka_str

def escape_xml(text):
    text = str(text)
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;')
    text = text.replace('>', '&gt;')
    text = text.replace('"', '&quot;')
    text = text.replace("'", '&apos;')
    return text

# =========================
# ✅ FIX TERMIN untuk template split <w:t>
# =========================
def replace_wt_text_in_context(xml: str, context_phrase: str, old_text: str, new_text: str, window: int = 3000, max_repl: int = 30) -> str:
    if not xml or not context_phrase or old_text is None:
        return xml

    ctx = context_phrase.lower()
    xml_lower = xml.lower()

    pattern = r'(<w:t[^>]*>)\s*' + re.escape(str(old_text)) + r'\s*(</w:t>)'
    rx = re.compile(pattern, re.IGNORECASE | re.DOTALL)

    out = []
    last = 0
    changed = 0

    for m in rx.finditer(xml):
        if changed >= max_repl:
            break

        start, end = m.start(), m.end()
        w_start = max(0, start - window)
        w_end = min(len(xml), end + window)

        if ctx in xml_lower[w_start:w_end]:
            out.append(xml[last:start])
            out.append(f"{m.group(1)}{new_text}{m.group(2)}")
            last = end
            changed += 1

    if changed > 0:
        out.append(xml[last:])
        return ''.join(out)

    return xml

def replace_split_digits_in_context(xml: str, context_phrase: str, digits: str, new_digits: str, window: int = 3000) -> str:
    if not xml or not digits or len(digits) != 2:
        return xml

    d1, d2 = digits[0], digits[1]
    ctx = context_phrase.lower()
    xml_lower = xml.lower()

    pattern = (
        r'(<w:t[^>]*>)\s*' + re.escape(d1) + r'\s*(</w:t>)'
        r'(\s*</w:r>\s*<w:r[^>]*>\s*(?:<w:rPr>.*?</w:rPr>)?\s*)'
        r'(<w:t[^>]*>)\s*' + re.escape(d2) + r'\s*(</w:t>)'
    )
    rx = re.compile(pattern, re.IGNORECASE | re.DOTALL)

    def _repl(m):
        start = m.start()
        end = m.end()
        w_start = max(0, start - window)
        w_end = min(len(xml), end + window)
        if ctx not in xml_lower[w_start:w_end]:
            return m.group(0)
        return f"{m.group(1)}{new_digits}{m.group(2)}"

    return rx.sub(_repl, xml)

def create_docx(data, filename):
    filepath = FILES_DIR / f"{filename}.docx"
    temp_extract = TEMP_DIR / f"extract_{uuid.uuid4().hex[:8]}"

    try:
        with zipfile.ZipFile(TEMPLATE_FILE, 'r') as zip_ref:
            zip_ref.extractall(temp_extract)

        now = datetime.now()
        bulan_romawi = angka_ke_romawi(now.strftime('%m'))
        bulan_id = {
            '01': 'Januari', '02': 'Februari', '03': 'Maret', '04': 'April',
            '05': 'Mei', '06': 'Juni', '07': 'Juli', '08': 'Agustus',
            '09': 'September', '10': 'Oktober', '11': 'November', '12': 'Desember'
        }

        nama_perusahaan = data['nama_perusahaan'].replace('\n', ' ').replace('\r', ' ')
        alamat_perusahaan = data['alamat_perusahaan'].replace('\n', ' ').replace('\r', ' ')
        harga_transportasi = format_rupiah(data.get('harga_transportasi', '1200000'))
        harga_mou = format_rupiah(data.get('harga_mou', '')) if data.get('harga_mou') else None

        doc_xml_path = temp_extract / "word" / "document.xml"
        with open(doc_xml_path, 'r', encoding='utf-8') as f:
            doc_content = f.read()

        doc_content = doc_content.replace('>027</w:t>', f'>{data.get("nomor_depan", "000")}</w:t>', 1)
        doc_content = doc_content.replace('>IX</w:t>', f'>{bulan_romawi}</w:t>', 1)
        doc_content = doc_content.replace('>PT Surgika Alkesindo, </w:t>', f'>{escape_xml(nama_perusahaan)}, </w:t>')
        doc_content = doc_content.replace('>PT. Surgika Alkesindo</w:t>', f'>{escape_xml(nama_perusahaan)}</w:t>')

        old_alamat = 'Jl Plumpang Semper No.6A RT.12/Rw.2, Tugu Utara, Kec. Koja, Jakarta Utara, DKI Jakarta 14260'
        doc_content = doc_content.replace(f'>{old_alamat}</w:t>', f'>{escape_xml(alamat_perusahaan)}</w:t>')
        doc_content = doc_content.replace('>28 November </w:t>', f'>{now.day} {bulan_id[now.strftime("%m")]} </w:t>', 1)

        # ✅ TERMIN FIX (pasti berubah walau split <w:t>)
        termin_hari = data.get('termin_hari', '14')
        termin_terbilang = angka_ke_terbilang(termin_hari)
        print("DEBUG_TERMIN_DATA:", termin_hari, termin_terbilang)

        doc_content = doc_content.replace('>14 (empat belas) Hari', f'>{termin_hari} ({termin_terbilang}) Hari')

        ctx_phrase = "Termin Pembayaran Paling Lambat"
        doc_content = replace_wt_text_in_context(doc_content, ctx_phrase, "14", termin_hari)
        doc_content = replace_wt_text_in_context(doc_content, ctx_phrase, "empat belas", termin_terbilang)
        doc_content = replace_split_digits_in_context(doc_content, ctx_phrase, "14", termin_hari)

        m = re.search(
            r'Termin Pembayaran Paling Lambat.*?<w:t[^>]*>\s*(\d+)\s*</w:t>.*?\(\s*</w:t>.*?<w:t[^>]*>\s*([^<]+?)\s*</w:t>',
            doc_content, re.IGNORECASE | re.DOTALL
        )
        print("DEBUG_TERMIN_AFTER_XML:", (m.group(1), m.group(2)) if m else "NOT_FOUND")

        table_start_pattern = r'<w:tbl>(.*?Jenis Limbah.*?)</w:tbl>'
        table_match = re.search(table_start_pattern, doc_content, re.DOTALL)

        if table_match:
            full_table = table_match.group(0)

            tblPr_match = re.search(r'(<w:tblPr>.*?</w:tblPr>)', full_table, re.DOTALL)
            tblGrid_match = re.search(r'(<w:tblGrid>.*?</w:tblGrid>)', full_table, re.DOTALL)

            tblPr = tblPr_match.group(1) if tblPr_match else ''
            tblGrid = tblGrid_match.group(1) if tblGrid_match else ''

            header_pattern = r'(<w:tr\b[^>]*>.*?Jenis Limbah.*?</w:tr>)'
            header_match = re.search(header_pattern, full_table, re.DOTALL)

            if header_match:
                header_row_xml = header_match.group(1)

                bold_border = '''<w:tcBorders>
                    <w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>
                    <w:left w:val="single" w:sz="12" w:space="0" w:color="000000"/>
                    <w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>
                    <w:right w:val="single" w:sz="12" w:space="0" w:color="000000"/>
                </w:tcBorders>'''

                new_rows_xml = ""
                items = data.get('items_limbah', [])
                for idx, item in enumerate(items, 1):
                    harga_formatted = format_rupiah(item.get('harga', ''))
                    jenis = escape_xml(item.get('jenis_limbah', ''))
                    kode = escape_xml(item.get('kode_limbah', ''))
                    satuan = escape_xml(item.get('satuan', 'Kg'))

                    new_rows_xml += f'''<w:tr>
                        <w:tc><w:tcPr><w:tcW w:w="850" w:type="dxa"/>{bold_border}</w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/></w:pPr>
                                <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>
                                    <w:t>{idx}</w:t>
                                </w:r>
                            </w:p>
                        </w:tc>
                        <w:tc><w:tcPr><w:tcW w:w="4536" w:type="dxa"/>{bold_border}<w:tcMar><w:left w:w="100" w:type="dxa"/></w:tcMar></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="left"/></w:pPr>
                                <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>
                                    <w:t>{jenis}</w:t>
                                </w:r>
                            </w:p>
                        </w:tc>
                        <w:tc><w:tcPr><w:tcW w:w="1701" w:type="dxa"/>{bold_border}</w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/></w:pPr>
                                <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>
                                    <w:t>{kode}</w:t>
                                </w:r>
                            </w:p>
                        </w:tc>
                        <w:tc><w:tcPr><w:tcW w:w="1701" w:type="dxa"/>{bold_border}</w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/></w:pPr>
                                <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>
                                    <w:t>{harga_formatted}</w:t>
                                </w:r>
                            </w:p>
                        </w:tc>
                        <w:tc><w:tcPr><w:tcW w:w="1134" w:type="dxa"/>{bold_border}</w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/></w:pPr>
                                <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>
                                    <w:t>{satuan}</w:t>
                                </w:r>
                            </w:p>
                        </w:tc>
                    </w:tr>'''

                new_rows_xml += f'''<w:tr>
                    <w:tc><w:tcPr><w:tcW w:w="7087" w:type="dxa"/><w:gridSpan w:val="3"/>{bold_border}</w:tcPr>
                        <w:p><w:pPr><w:jc w:val="center"/></w:pPr>
                            <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="22"/></w:rPr>
                                <w:t>Biaya Transportasi</w:t>
                            </w:r>
                        </w:p>
                    </w:tc>
                    <w:tc><w:tcPr><w:tcW w:w="1701" w:type="dxa"/>{bold_border}</w:tcPr>
                        <w:p><w:pPr><w:jc w:val="center"/></w:pPr>
                            <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>
                                <w:t>{harga_transportasi}</w:t>
                            </w:r>
                        </w:p>
                    </w:tc>
                    <w:tc><w:tcPr><w:tcW w:w="1134" w:type="dxa"/>{bold_border}</w:tcPr>
                        <w:p><w:pPr><w:jc w:val="center"/></w:pPr>
                            <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>
                                <w:t>ritase</w:t>
                            </w:r>
                        </w:p>
                    </w:tc>
                </w:tr>'''

                if harga_mou:
                    new_rows_xml += f'''<w:tr>
                        <w:tc><w:tcPr><w:tcW w:w="7087" w:type="dxa"/><w:gridSpan w:val="3"/>{bold_border}</w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/></w:pPr>
                                <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="22"/></w:rPr>
                                    <w:t>Biaya MoU</w:t>
                                </w:r>
                            </w:p>
                        </w:tc>
                        <w:tc><w:tcPr><w:tcW w:w="1701" w:type="dxa"/>{bold_border}</w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/></w:pPr>
                                <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>
                                    <w:t>{harga_mou}</w:t>
                                </w:r>
                            </w:p>
                        </w:tc>
                        <w:tc><w:tcPr><w:tcW w:w="1134" w:type="dxa"/>{bold_border}</w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/></w:pPr>
                                <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>
                                    <w:t>Tahun</w:t>
                                </w:r>
                            </w:p>
                        </w:tc>
                    </w:tr>'''

                new_table = f'<w:tbl>{tblPr}{tblGrid}{header_row_xml}{new_rows_xml}</w:tbl>'
                doc_content = doc_content.replace(full_table, new_table)

        with open(doc_xml_path, 'w', encoding='utf-8') as f:
            f.write(doc_content)

        word_dir = temp_extract / "word"
        header_replacements = {
            '>027</w:t>': f'>{data.get("nomor_depan", "000")}</w:t>',
            '>IX</w:t>': f'>{bulan_romawi}</w:t>',
            '>PT Surgika Alkesindo, </w:t>': f'>{escape_xml(nama_perusahaan)}, </w:t>',
            '>PT. Surgika Alkesindo</w:t>': f'>{escape_xml(nama_perusahaan)}</w:t>',
            f'>{old_alamat}</w:t>': f'>{escape_xml(alamat_perusahaan)}</w:t>',
            '>28 November </w:t>': f'>{now.day} {bulan_id[now.strftime("%m")]} </w:t>',
        }

        for xml_file in word_dir.glob("*.xml"):
            if xml_file.name.startswith(('header', 'footer')):
                with open(xml_file, 'r', encoding='utf-8') as f:
                    content = f.read()

                for old_text, new_text in header_replacements.items():
                    content = content.replace(old_text, new_text)

                content = content.replace('>14 (empat belas) Hari', f'>{termin_hari} ({termin_terbilang}) Hari')
                ctx_phrase = "Termin Pembayaran Paling Lambat"
                content = replace_wt_text_in_context(content, ctx_phrase, "14", termin_hari)
                content = replace_wt_text_in_context(content, ctx_phrase, "empat belas", termin_terbilang)
                content = replace_split_digits_in_context(content, ctx_phrase, "14", termin_hari)

                with open(xml_file, 'w', encoding='utf-8') as f:
                    f.write(content)

        with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as docx:
            for file_path in temp_extract.rglob('*'):
                if file_path.is_file():
                    arcname = file_path.relative_to(temp_extract)
                    docx.write(file_path, arcname)

        return f"{filename}.docx"

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        raise
    finally:
        if temp_extract.exists():
            shutil.rmtree(temp_extract)

# ========== PDF GENERATION ==========
def create_pdf_libreoffice(docx_path, pdf_path):
    try:
        print(f"  → Converting with LibreOffice...")
        print(f"     Path: {LIBREOFFICE_PATH}")

        if pdf_path.exists():
            pdf_path.unlink()
            print(f"     Removed existing PDF")

        FILES_DIR.mkdir(parents=True, exist_ok=True)

        cmd = [
            str(LIBREOFFICE_PATH),
            '--headless',
            '--invisible',
            '--nocrashreport',
            '--nodefault',
            '--nofirststartwizard',
            '--nolockcheck',
            '--nologo',
            '--norestore',
            '--convert-to', 'pdf:writer_pdf_Export',
            '--outdir', str(FILES_DIR),
            str(docx_path)
        ]

        print(f"     Running command: {' '.join(cmd)}")

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=90,
            check=False,
            env={**os.environ, 'HOME': str(TEMP_DIR)}
        )

        print(f"     Return code: {result.returncode}")
        if result.stdout:
            print(f"     STDOUT: {result.stdout}")
        if result.stderr:
            print(f"     STDERR: {result.stderr}")

        if pdf_path.exists():
            file_size = pdf_path.stat().st_size
            print(f"     ✓ PDF created: {file_size} bytes")
            if file_size > 1000:
                return True
            else:
                print(f"     ✗ PDF too small ({file_size} bytes) - likely corrupt")
                return False
        else:
            print(f"     ✗ PDF file not created at: {pdf_path}")
            return False

    except subprocess.TimeoutExpired:
        print(f"  ✗ LibreOffice conversion timeout (>90s)")
        return False
    except Exception as e:
        print(f"  ✗ LibreOffice error: {e}")
        import traceback
        traceback.print_exc()
        return False

def create_pdf_docx2pdf(docx_path, pdf_path):
    try:
        print(f"  → Trying docx2pdf...")
        docx_to_pdf(str(docx_path), str(pdf_path))

        if pdf_path.exists() and pdf_path.stat().st_size > 1000:
            print(f"  ✓ docx2pdf succeeded")
            return True
        else:
            print(f"  ✗ docx2pdf failed to create valid PDF")
            return False
    except Exception as e:
        print(f"  ✗ docx2pdf failed: {e}")
        return False

def create_pdf_pypandoc(docx_path, pdf_path):
    try:
        print(f"  → Trying pypandoc...")
        pypandoc.convert_file(
            str(docx_path),
            'pdf',
            outputfile=str(pdf_path),
            extra_args=['--pdf-engine=xelatex']
        )

        if pdf_path.exists() and pdf_path.stat().st_size > 1000:
            print(f"  ✓ pypandoc succeeded")
            return True
        else:
            print(f"  ✗ pypandoc failed to create valid PDF")
            return False
    except Exception as e:
        print(f"  ✗ pypandoc failed: {e}")
        return False

def create_pdf(filename):
    if not PDF_AVAILABLE:
        print("❌ PDF generation disabled - no converter available")
        return None

    docx_path = FILES_DIR / f"{filename}.docx"
    pdf_path = FILES_DIR / f"{filename}.pdf"

    if not docx_path.exists():
        print(f"❌ DOCX not found: {docx_path}")
        return None

    print(f"\n{'='*60}")
    print(f"🔄 Converting to PDF: {filename}.docx")
    print(f"{'='*60}")
    print(f"📄 DOCX size: {docx_path.stat().st_size:,} bytes")
    print(f"🔧 Primary method: {PDF_METHOD}")
    print(f"💻 Platform: {platform.system()}")
    print(f"{'='*60}")

    try:
        success = False

        if platform.system() == "Linux":
            if LIBREOFFICE_PATH:
                success = create_pdf_libreoffice(docx_path, pdf_path)
            if not success and pypandoc and PDF_METHOD != "libreoffice":
                success = create_pdf_pypandoc(docx_path, pdf_path)
        else:
            if PDF_METHOD == "docx2pdf" and docx_to_pdf:
                success = create_pdf_docx2pdf(docx_path, pdf_path)
            elif PDF_METHOD == "libreoffice" and LIBREOFFICE_PATH:
                success = create_pdf_libreoffice(docx_path, pdf_path)
            elif PDF_METHOD == "pypandoc" and pypandoc:
                success = create_pdf_pypandoc(docx_path, pdf_path)

            if not success:
                if not success and docx_to_pdf and PDF_METHOD != "docx2pdf":
                    success = create_pdf_docx2pdf(docx_path, pdf_path)
                if not success and LIBREOFFICE_PATH and PDF_METHOD != "libreoffice":
                    success = create_pdf_libreoffice(docx_path, pdf_path)
                if not success and pypandoc and PDF_METHOD != "pypandoc":
                    success = create_pdf_pypandoc(docx_path, pdf_path)

        if success and pdf_path.exists():
            file_size = pdf_path.stat().st_size
            if file_size > 1000:
                print(f"\n✅ PDF CREATED SUCCESSFULLY: {pdf_path.name} ({file_size} bytes)\n")
                return f"{filename}.pdf"
            else:
                print(f"\n❌ PDF too small ({file_size} bytes) - likely corrupt")
                return None
        else:
            print(f"\n❌ PDF CONVERSION FAILED")
            return None

    except Exception as e:
        print(f"\n❌ PDF CONVERSION ERROR: {e}")
        import traceback
        traceback.print_exc()
        return None

@app.route("/")
def index():
    return render_template("index.html")

# =========================
# ✅ HISTORY API
# =========================
@app.route("/api/history", methods=["GET"])
def api_history_list():
    try:
        q = (request.args.get("q") or "").strip()
        items = db_list_histories(limit=200, q=q if q else None)
        return jsonify({"items": items})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ✅ NEW: detail history + messages + files
@app.route("/api/history/<int:history_id>", methods=["GET"])
def api_history_detail(history_id):
    try:
        detail = db_get_history_detail(history_id)
        if not detail:
            return jsonify({"error": "history tidak ditemukan"}), 404

        return jsonify({
            "id": detail["id"],
            "title": detail["title"],
            "task_type": detail["task_type"],
            "created_at": detail["created_at"],
            "data": json.loads(detail.get("data_json") or "{}"),
            "files": json.loads(detail.get("files_json") or "[]"),
            "messages": json.loads(detail.get("messages_json") or "[]"),
            "state": json.loads(detail.get("state_json") or "{}"),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/history/<int:history_id>", methods=["PUT"])
def api_history_update(history_id):
    try:
        body = request.get_json() or {}
        new_title = (body.get("title") or "").strip()
        if not new_title:
            return jsonify({"error": "title wajib diisi"}), 400

        ok = db_update_title(history_id, new_title)
        if not ok:
            return jsonify({"error": "history tidak ditemukan"}), 404
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/history/<int:history_id>", methods=["DELETE"])
def api_history_delete(history_id):
    try:
        ok = db_delete_history(history_id)
        if not ok:
            return jsonify({"error": "history tidak ditemukan"}), 404
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ✅ NEW: Documents list (gabung semua files dari history)
@app.route("/api/documents", methods=["GET"])
def api_documents():
    try:
        q = (request.args.get("q") or "").strip().lower()
        items = db_list_histories(limit=500)

        docs = []
        for h in items:
            detail = db_get_history_detail(int(h["id"]))
            if not detail:
                continue
            try:
                files = json.loads(detail.get("files_json") or "[]")
            except:
                files = []
            for f in files:
                # f: {type, filename, url}
                filename = (f.get("filename") or "").strip()
                if not filename:
                    continue
                title = detail.get("title") or ""
                task_type = detail.get("task_type") or ""
                created_at = detail.get("created_at") or ""

                row = {
                    "history_id": int(detail["id"]),
                    "history_title": title,
                    "task_type": task_type,
                    "created_at": created_at,
                    "type": f.get("type"),
                    "filename": filename,
                    "url": f.get("url"),
                }

                if q:
                    hay = f"{title} {filename} {task_type}".lower()
                    if q not in hay:
                        continue

                docs.append(row)

        # terbaru di atas
        docs.sort(key=lambda x: x.get("created_at") or "", reverse=True)
        return jsonify({"items": docs})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# =========================
# ✅ CHAT
# =========================
@app.route("/api/chat", methods=["POST"])
def chat():
    try:
        data = request.get_json() or {}
        text = (data.get("message", "") or "").strip()
        history_id_in = data.get("history_id")  # ✅ NEW: agar chat bisa lanjut ke history yg sama

        if not text:
            return jsonify({"error": "Pesan kosong"}), 400

        sid = session.get('sid')
        if not sid:
            sid = str(uuid.uuid4())
            session['sid'] = sid

        state = conversations.get(sid, {'step': 'idle', 'data': {}})
        lower = text.lower()

        # ✅ kalau client kirim history_id, simpan pesan user ke history itu
        if history_id_in:
            try:
                db_append_message(int(history_id_in), "user", text, files=[])
                db_update_state(int(history_id_in), state)
            except:
                pass

        # Start conversation
        if 'quotation' in lower or 'penawaran' in lower or 'buat' in lower:
            nomor_depan = get_next_nomor()
            state['step'] = 'nama_perusahaan'
            now = datetime.now()
            state['data'] = {
                'nomor_depan': nomor_depan,
                'items_limbah': [],
                'bulan_romawi': now.strftime('%m')
            }
            conversations[sid] = state

            out_text = f"Baik, saya bantu buatkan quotation.<br><br>✅ Nomor Surat: <b>{nomor_depan}</b><br><br>❓ <b>1. Nama Perusahaan?</b>"

            # ✅ kalau belum ada history_id, buat history baru untuk chat ini
            history_id_created = None
            if not history_id_in:
                history_id_created = db_insert_history(
                    title="Chat Baru",
                    task_type=data.get("taskType") or "penawaran",
                    data={},
                    files=[],
                    messages=[
                        {"id": uuid.uuid4().hex[:12], "sender": "user", "text": text, "files": [], "timestamp": datetime.now().isoformat()},
                        {"id": uuid.uuid4().hex[:12], "sender": "assistant", "text": re.sub(r'<br\s*/?>', '\n', out_text), "files": [], "timestamp": datetime.now().isoformat()},
                    ],
                    state=state
                )
            else:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)

            return jsonify({"text": out_text, "history_id": history_id_created or history_id_in})

        # Step 1: Nama Perusahaan (alamat auto: Serper -> AI -> Di Tempat)
        if state['step'] == 'nama_perusahaan':
            state['data']['nama_perusahaan'] = text

            alamat = search_company_address(text).strip()
            if not alamat:
                alamat = search_company_address_ai(text).strip()
            if not alamat:
                alamat = "Di Tempat"

            state['data']['alamat_perusahaan'] = alamat
            state['step'] = 'jenis_kode_limbah'
            state['data']['current_item'] = {}
            conversations[sid] = state

            out_text = (
                f"✅ Nama: <b>{text}</b><br>✅ Alamat: <b>{alamat}</b><br><br>"
                f"📦 <b>Item #1</b><br>❓ <b>2. Sebutkan Jenis Limbah atau Kode Limbah</b><br>"
                f"<i>(Contoh: 'A102d' atau 'aki baterai bekas')</i>"
            )

            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)

            return jsonify({"text": out_text, "history_id": history_id_in})

        # Step 2: Jenis/Kode Limbah
        elif state['step'] == 'jenis_kode_limbah':
            kode, data_limbah = find_limbah_by_kode(text)

            if kode and data_limbah:
                state['data']['current_item']['kode_limbah'] = kode
                state['data']['current_item']['jenis_limbah'] = data_limbah['jenis']
                state['data']['current_item']['satuan'] = data_limbah['satuan']
                state['step'] = 'harga'
                conversations[sid] = state
                out_text = f"✅ Kode: <b>{kode}</b><br>✅ Jenis: <b>{data_limbah['jenis']}</b><br>✅ Satuan: <b>{data_limbah['satuan']}</b><br><br>❓ <b>3. Harga (Rp)?</b>"

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                    db_update_state(int(history_id_in), state)

                return jsonify({"text": out_text, "history_id": history_id_in})
            else:
                kode, data_limbah = find_limbah_by_jenis(text)

                if kode and data_limbah:
                    state['data']['current_item']['kode_limbah'] = kode
                    state['data']['current_item']['jenis_limbah'] = data_limbah['jenis']
                    state['data']['current_item']['satuan'] = data_limbah['satuan']
                    state['step'] = 'harga'
                    conversations[sid] = state
                    out_text = f"✅ Kode: <b>{kode}</b><br>✅ Jenis: <b>{data_limbah['jenis']}</b><br>✅ Satuan: <b>{data_limbah['satuan']}</b><br><br>❓ <b>3. Harga (Rp)?</b>"

                    if history_id_in:
                        db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                        db_update_state(int(history_id_in), state)

                    return jsonify({"text": out_text, "history_id": history_id_in})
                else:
                    out_text = (
                        f"❌ Maaf, limbah '<b>{text}</b>' tidak ditemukan dalam database.<br><br>"
                        "Silakan coba lagi dengan:<br>"
                        "• Kode limbah (contoh: A102d, B105d)<br>"
                        "• Nama jenis limbah (contoh: aki baterai bekas, minyak pelumas bekas)"
                    )

                    if history_id_in:
                        db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                        db_update_state(int(history_id_in), state)

                    return jsonify({"text": out_text, "history_id": history_id_in})

        # Step 3: Harga
        elif state['step'] == 'harga':
            harga_converted = convert_voice_to_number(text)
            state['data']['current_item']['harga'] = harga_converted

            state['data']['items_limbah'].append(state['data']['current_item'])
            num = len(state['data']['items_limbah'])
            state['step'] = 'tambah_item'
            conversations[sid] = state

            harga_formatted = format_rupiah(harga_converted)
            out_text = f"✅ Item #{num} tersimpan!<br>💰 Harga: <b>Rp {harga_formatted}</b><br><br>❓ <b>Tambah item lagi?</b> (ya/tidak)"

            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)

            return jsonify({"text": out_text, "history_id": history_id_in})

        # Step 4: Tambah Item?
        elif state['step'] == 'tambah_item':
            if 'ya' in lower or 'iya' in lower:
                num = len(state['data']['items_limbah'])
                state['step'] = 'jenis_kode_limbah'
                state['data']['current_item'] = {}
                conversations[sid] = state
                out_text = f"📦 <b>Item #{num+1}</b><br>❓ <b>2. Sebutkan Jenis Limbah atau Kode Limbah</b><br><i>(Contoh: 'A102d' atau 'aki baterai bekas')</i>"

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                    db_update_state(int(history_id_in), state)

                return jsonify({"text": out_text, "history_id": history_id_in})
            else:
                state['step'] = 'harga_transportasi'
                conversations[sid] = state
                out_text = f"✅ Total: <b>{len(state['data']['items_limbah'])} item</b><br><br>❓ <b>4. Biaya Transportasi (Rp)?</b><br><i>Satuan: ritase</i>"

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                    db_update_state(int(history_id_in), state)

                return jsonify({"text": out_text, "history_id": history_id_in})

        # Step 5: Harga Transportasi
        elif state['step'] == 'harga_transportasi':
            transportasi_converted = convert_voice_to_number(text)
            state['data']['harga_transportasi'] = transportasi_converted
            state['step'] = 'tanya_mou'
            conversations[sid] = state
            transportasi_formatted = format_rupiah(transportasi_converted)
            out_text = f"✅ Transportasi: <b>Rp {transportasi_formatted}/ritase</b><br><br>❓ <b>5. Tambah Biaya MoU?</b> (ya/tidak)"

            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)

            return jsonify({"text": out_text, "history_id": history_id_in})

        # Step 6: Tanya MoU
        elif state['step'] == 'tanya_mou':
            if 'ya' in lower or 'iya' in lower:
                state['step'] = 'harga_mou'
                conversations[sid] = state
                out_text = "❓ <b>Biaya MoU (Rp)?</b><br><i>Satuan: Tahun</i>"

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                    db_update_state(int(history_id_in), state)

                return jsonify({"text": out_text, "history_id": history_id_in})
            else:
                state['data']['harga_mou'] = None
                state['step'] = 'tanya_termin'
                conversations[sid] = state
                out_text = "❓ <b>6. Edit Termin Pembayaran?</b><br><i>Default: 14 hari</i><br>(ketik angka atau 'tidak' untuk default)"

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                    db_update_state(int(history_id_in), state)

                return jsonify({"text": out_text, "history_id": history_id_in})

        # Step 7: Harga MoU
        elif state['step'] == 'harga_mou':
            mou_converted = convert_voice_to_number(text)
            state['data']['harga_mou'] = mou_converted
            state['step'] = 'tanya_termin'
            conversations[sid] = state

            mou_formatted = format_rupiah(mou_converted)
            out_text = f"✅ MoU: <b>Rp {mou_formatted}/Tahun</b><br><br>❓ <b>6. Edit Termin Pembayaran?</b><br><i>Default: 14 hari</i><br>(ketik angka atau 'tidak' untuk default)"

            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)

            return jsonify({"text": out_text, "history_id": history_id_in})

        # Step 8: Tanya Termin
        elif state['step'] == 'tanya_termin':
            if 'tidak' in lower or 'skip' in lower or 'lewat' in lower:
                state['data']['termin_hari'] = '14'
            else:
                state['data']['termin_hari'] = parse_termin_days(text, default=14, min_days=1, max_days=365)

            fname = f"Quotation_{re.sub(r'[^A-Za-z0-9]+', '_', state['data']['nama_perusahaan'])}_{uuid.uuid4().hex[:6]}"

            print(f"\n{'='*60}")
            print(f"📝 Creating documents for: {state['data']['nama_perusahaan']}")
            print(f"{'='*60}")

            docx = create_docx(state['data'], fname)
            print(f"✅ DOCX created: {docx}")

            pdf = create_pdf(fname)
            if pdf:
                print(f"✅ PDF created: {pdf}")
            else:
                print(f"⚠️  PDF not created - continuing without PDF")

            print(f"{'='*60}\n")

            conversations[sid] = {'step': 'idle', 'data': {}}

            files = [{"type": "docx", "filename": docx, "url": f"/static/files/{docx}"}]
            if pdf:
                files.append({"type": "pdf", "filename": pdf, "url": f"/static/files/{pdf}"})

            nama_pt = state['data'].get('nama_perusahaan', '').strip()
            history_title = f"Penawaran {nama_pt}" if nama_pt else "Penawaran"
            history_task_type = "penawaran"

            # ✅ jika sudah ada history_id (chat sedang berjalan) -> update title + files + data, bukan insert baru
            if history_id_in:
                conn = db_connect()
                cur = conn.cursor()
                cur.execute("""
                    UPDATE chat_history
                    SET title = ?, task_type = ?, data_json = ?, files_json = ?
                    WHERE id = ?
                """, (
                    history_title,
                    history_task_type,
                    json.dumps(state['data'], ensure_ascii=False),
                    json.dumps(files, ensure_ascii=False),
                    int(history_id_in),
                ))
                conn.commit()
                conn.close()
                history_id = int(history_id_in)
            else:
                history_id = db_insert_history(
                    title=history_title,
                    task_type=history_task_type,
                    data=state['data'],
                    files=files,
                    messages=[],
                    state={}
                )

            termin_terbilang = angka_ke_terbilang(state['data']['termin_hari'])
            out_text = f"✅ Termin: <b>{state['data']['termin_hari']} ({termin_terbilang}) hari</b><br><br>🎉 <b>Quotation berhasil dibuat!</b>"

            # ✅ simpan pesan assistant + file card ke history
            db_append_message(history_id, "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=files)

            return jsonify({
                "text": out_text,
                "files": files,
                "history_id": history_id
            })

        # fallback AI
        ai_out = call_ai(text)
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", ai_out, files=[])
        return jsonify({"text": ai_out, "history_id": history_id_in})

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route("/download/<path:filename>")
def download(filename):
    return send_from_directory(str(FILES_DIR), filename, as_attachment=True)

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    debug_mode = os.getenv("FLASK_ENV") != "production"

    print("\n" + "="*60)
    print("🚀 QUOTATION GENERATOR - SMART LIMBAH B3 DETECTION")
    print("="*60)
    print(f"📁 Template: {TEMPLATE_FILE.exists() and '✅ Found' or '❌ Missing'}")
    print(f"🔑 OpenRouter Key: {OPENROUTER_API_KEY and '✅ Set' or '❌ Not Set'}")
    print(f"🔎 Serper Key: {SERPER_API_KEY and '✅ Set' or '❌ Not Set'}")
    print(f"📄 PDF Generation: {PDF_AVAILABLE and '✅ ENABLED' or '❌ DISABLED'}")
    if PDF_AVAILABLE:
        print(f"   Primary Method: {PDF_METHOD}")
        if PDF_METHOD == "docx2pdf":
            print(f"   Library: docx2pdf")
        elif PDF_METHOD == "libreoffice":
            print(f"   Path: {LIBREOFFICE_PATH}")
            print(f"   ⚠️  Optimized for Linux/Render.com")
        elif PDF_METHOD == "pypandoc":
            print(f"   Library: pypandoc")
    print(f"🗄️  Database: {len(LIMBAH_B3_DB)} jenis limbah B3")
    print(f"🔢 Current Counter: {load_counter()}")
    print(f"🌐 Port: {port}")
    print(f"🔧 Debug: {debug_mode}")
    print(f"💻 Platform: {platform.system()}")
    print("="*60 + "\n")

    app.run(host="0.0.0.0", port=port, debug=debug_mode)
