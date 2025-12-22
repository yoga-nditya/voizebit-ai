"""
Utils Module - Database, PDF, Document, Search Functions
"""
import os
import json
import uuid
import re
from pathlib import Path
from datetime import datetime
import shutil
import zipfile
import subprocess
import platform
import sqlite3
from html import unescape
import requests

from config_new import *
from limbah_database import format_rupiah, angka_ke_terbilang, angka_ke_romawi

PDF_AVAILABLE = False
PDF_METHOD = None

try:
    from docx2pdf import convert as docx_to_pdf
    PDF_AVAILABLE = True
    PDF_METHOD = "docx2pdf"
except ImportError:
    docx_to_pdf = None

try:
    import pypandoc
    if not PDF_AVAILABLE:
        PDF_AVAILABLE = True
        PDF_METHOD = "pypandoc"
except ImportError:
    pypandoc = None

def check_libreoffice():
    try:
        system = platform.system()
        
        if system == "Windows":
            paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            ]
            for p in paths:
                if os.path.exists(p):
                    return p
        else:
            commands = ['libreoffice', 'soffice', '/usr/bin/libreoffice', '/usr/bin/soffice']
            for cmd in commands:
                try:
                    if os.path.exists(cmd) and os.access(cmd, os.X_OK):
                        return cmd

                    result = subprocess.run(['which', cmd], capture_output=True, text=True, timeout=5)
                    if result.returncode == 0 and result.stdout.strip():
                        return result.stdout.strip()
                except Exception:
                    continue

            for cmd in ['libreoffice', 'soffice']:
                try:
                    result = subprocess.run(['command', '-v', cmd], capture_output=True, text=True, timeout=5, shell=True)
                    if result.returncode == 0 and result.stdout.strip():
                        return result.stdout.strip()
                except:
                    continue
    except Exception as e:
        print(f"Error checking LibreOffice: {e}")

    return None

LIBREOFFICE_PATH = check_libreoffice()
if LIBREOFFICE_PATH and not PDF_AVAILABLE:
    PDF_AVAILABLE = True
    PDF_METHOD = "libreoffice"


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

def plain(s: str) -> str:
    if not s:
        return ""
    s = str(s)
    s = unescape(s)
    s = re.sub(r'<br\s*/?>', '\n', s, flags=re.IGNORECASE)
    s = re.sub(r'<[^>]+>', '', s)
    s = s.replace('\r', '')
    s = re.sub(r'\n{3,}', '\n\n', s)
    s = re.sub(r'[ \t]+', ' ', s)
    return s.strip()

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

def load_counter():
    if COUNTER_FILE.exists():
        try:
            with open(COUNTER_FILE, 'r') as f:
                data = json.load(f)
                return int(data.get('counter', 0))
        except:
            return 0
    return 0

def save_counter(counter):
    with open(COUNTER_FILE, 'w') as f:
        json.dump({'counter': int(counter)}, f)

def get_next_nomor():
    counter = load_counter()
    try:
        counter = int(counter)
    except:
        counter = 0

    if counter < 0:
        counter = 0
    if counter > 21:
        counter = 0

    nomor = str(counter).zfill(3)
    next_counter = (counter + 1) % 22
    save_counter(next_counter)
    return nomor

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
    if len(name) < 3 or not SERPER_API_KEY:
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
        print(f"Error searching address: {e}")
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
        print(f"AI address error: {e}")
        return ""

def escape_xml(text):
    text = str(text)
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;')
    text = text.replace('>', '&gt;')
    text = text.replace('"', '&quot;')
    text = text.replace("'", '&apos;')
    return text

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
        bulan_id = BULAN_ID

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

        termin_hari = data.get('termin_hari', '14')
        termin_terbilang = angka_ke_terbilang(termin_hari)

        doc_content = doc_content.replace('>14 (empat belas) Hari', f'>{termin_hari} ({termin_terbilang}) Hari')

        ctx_phrase = "Termin Pembayaran Paling Lambat"
        doc_content = replace_wt_text_in_context(doc_content, ctx_phrase, "14", termin_hari)
        doc_content = replace_wt_text_in_context(doc_content, ctx_phrase, "empat belas", termin_terbilang)
        doc_content = replace_split_digits_in_context(doc_content, ctx_phrase, "14", termin_hari)

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

def create_pdf_libreoffice(docx_path, pdf_path):
    try:
        if pdf_path.exists():
            pdf_path.unlink()

        FILES_DIR.mkdir(parents=True, exist_ok=True)

        cmd = [
            str(LIBREOFFICE_PATH),
            '--headless', '--invisible', '--nocrashreport', '--nodefault',
            '--nofirststartwizard', '--nolockcheck', '--nologo', '--norestore',
            '--convert-to', 'pdf:writer_pdf_Export',
            '--outdir', str(FILES_DIR),
            str(docx_path)
        ]

        result = subprocess.run(cmd, capture_output=True, text=True, timeout=90, check=False, env={**os.environ, 'HOME': str(TEMP_DIR)})

        if pdf_path.exists() and pdf_path.stat().st_size > 1000:
            return True
        return False
    except:
        return False

def create_pdf_docx2pdf(docx_path, pdf_path):
    try:
        docx_to_pdf(str(docx_path), str(pdf_path))
        return pdf_path.exists() and pdf_path.stat().st_size > 1000
    except:
        return False

def create_pdf_pypandoc(docx_path, pdf_path):
    try:
        pypandoc.convert_file(str(docx_path), 'pdf', outputfile=str(pdf_path), extra_args=['--pdf-engine=xelatex'])
        return pdf_path.exists() and pdf_path.stat().st_size > 1000
    except:
        return False

def create_pdf(filename):
    if not PDF_AVAILABLE:
        return None

    docx_path = FILES_DIR / f"{filename}.docx"
    pdf_path = FILES_DIR / f"{filename}.pdf"

    if not docx_path.exists():
        return None

    try:
        success = False

        if platform.system() == "Linux":
            if LIBREOFFICE_PATH:
                success = create_pdf_libreoffice(docx_path, pdf_path)
            if not success and pypandoc:
                success = create_pdf_pypandoc(docx_path, pdf_path)
        else:
            if PDF_METHOD == "docx2pdf" and docx_to_pdf:
                success = create_pdf_docx2pdf(docx_path, pdf_path)
            elif PDF_METHOD == "libreoffice" and LIBREOFFICE_PATH:
                success = create_pdf_libreoffice(docx_path, pdf_path)
            elif PDF_METHOD == "pypandoc" and pypandoc:
                success = create_pdf_pypandoc(docx_path, pdf_path)

            if not success and docx_to_pdf:
                success = create_pdf_docx2pdf(docx_path, pdf_path)
            if not success and LIBREOFFICE_PATH:
                success = create_pdf_libreoffice(docx_path, pdf_path)
            if not success and pypandoc:
                success = create_pdf_pypandoc(docx_path, pdf_path)

        if success and pdf_path.exists() and pdf_path.stat().st_size > 1000:
            return f"{filename}.pdf"
        return None
    except:
        return None