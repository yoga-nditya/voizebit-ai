import os
import json
import re
from datetime import datetime

from utils import (
    search_company_address, search_company_address_ai,
)
from config_new import FILES_DIR


# =========================
# Helpers (as-is)
# =========================

def is_non_b3_input(text: str) -> bool:
    if not text:
        return False
    t = text.strip().lower()
    norm = re.sub(r'[\s\-_]+', '', t)
    return norm in ("nonb3", "nonbii3") or norm.startswith("nonb3")

def make_unique_filename_base(base_name: str) -> str:
    base_name = (base_name or "").strip()
    if not base_name:
        base_name = "Dokumen"
    try:
        folder = str(FILES_DIR)
    except Exception:
        folder = "static/files"

    def exists_any(name: str) -> bool:
        return (
            os.path.exists(os.path.join(folder, f"{name}.docx")) or
            os.path.exists(os.path.join(folder, f"{name}.pdf")) or
            os.path.exists(os.path.join(folder, f"{name}.xlsx")) or
            os.path.exists(os.path.join(folder, name))
        )

    if not exists_any(base_name):
        return base_name

    i = 2
    while True:
        candidate = f"{base_name} ({i})"
        if not exists_any(candidate):
            return candidate
        i += 1

def month_to_roman(m: int) -> str:
    rom = {
        1: "I", 2: "II", 3: "III", 4: "IV",
        5: "V", 6: "VI", 7: "VII", 8: "VIII",
        9: "IX", 10: "X", 11: "XI", 12: "XII"
    }
    return rom.get(m, "")

def company_to_code(name: str) -> str:
    if not name:
        return "XXX"
    t = re.sub(r'[^A-Za-z0-9 ]+', ' ', name).strip()
    t = re.sub(r'\s+', ' ', t)
    parts = [p for p in t.split() if p.lower() not in ("pt", "pt.", "persero", "tbk")]
    if not parts:
        return "XXX"
    if len(parts) == 1:
        return (parts[0][:3]).upper().ljust(3, "X")
    code = "".join([p[0] for p in parts[:3]]).upper()
    return code.ljust(3, "X")

def build_mou_nomor_surat(mou_data: dict) -> str:
    no_depan = (mou_data.get("nomor_depan") or "").strip()
    kode_p1 = company_to_code((mou_data.get("pihak_pertama") or "").strip())
    kode_p2 = (mou_data.get("pihak_kedua_kode") or "STBJ").strip().upper()
    kode_p3 = (mou_data.get("pihak_ketiga_kode") or "").strip().upper()
    now = datetime.now()
    romawi = month_to_roman(now.month)
    tahun = str(now.year)
    if not kode_p3:
        kode_p3 = "XXX"
    return f"{no_depan}/PKPLNB3/{kode_p1}-{kode_p2}-{kode_p3}/{romawi}/{tahun}"

def format_tanggal_indonesia(dt: datetime) -> str:
    hari_map = {0: "Senin", 1: "Selasa", 2: "Rabu", 3: "Kamis", 4: "Jumat", 5: "Sabtu", 6: "Minggu"}
    bulan_map = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
                 7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
    hari = hari_map.get(dt.weekday(), "")
    bulan = bulan_map.get(dt.month, "")
    return f"{hari}, tanggal {dt.day} {bulan} {dt.year}"


# =========================
# ✅ FIX: baca angka singkat (rb/jt/miliar/triliun/k)
# =========================
def parse_id_short_amount(text: str):
    """
    Contoh input yang didukung:
      - "1.5jt" / "1,5jt" / "1.5 juta" / "1,5 juta" -> 1500000
      - "250rb" / "250 ribu" -> 250000
      - "2m" / "2 miliar" -> 2000000000
      - "1.2t" / "1,2 triliun" -> 1200000000000
      - "10k" -> 10000
      - "1 koma 5 juta" -> 1500000  (opsional)
    Return:
      int atau None jika tidak terbaca.
    """
    if not text:
        return None

    raw = text.strip()
    low = raw.lower().strip()

    # ---- 1) format "koma" (opsional): "1 koma 5 juta"
    if "koma" in low:
        scale_words = {
            "ribu": 1_000,
            "rb": 1_000,
            "k": 1_000,
            "jt": 1_000_000,
            "juta": 1_000_000,
            "m": 1_000_000_000,
            "miliar": 1_000_000_000,
            "t": 1_000_000_000_000,
            "triliun": 1_000_000_000_000,
        }
        scale = None
        for k, v in scale_words.items():
            if re.search(rf"\b{k}\b", low):
                scale = v
                break

        parts = re.split(r"\bkoma\b", low, maxsplit=1)
        left = parts[0].strip()
        right = parts[1].strip() if len(parts) > 1 else ""

        left_tok = re.findall(r"[0-9]+", left)
        right_tok = re.findall(r"[0-9]+", right)

        if left_tok and right_tok:
            try:
                val = float(f"{left_tok[-1]}.{right_tok[0]}")
                if scale:
                    val *= scale
                return int(round(val))
            except Exception:
                pass

    # ---- 2) format compact: "1.5jt", "250rb", "2m", "10k"
    # normalize: buang spasi
    t = re.sub(r"\s+", "", low)

    m = re.match(r"^(\d+(?:[.,]\d+)?)(k|rb|ribu|jt|juta|m|miliar|t|triliun)\b", t, re.IGNORECASE)
    if not m:
        return None

    num_s = m.group(1).replace(",", ".")
    suf = m.group(2).lower()

    mult_map = {
        "k": 1_000,
        "rb": 1_000,
        "ribu": 1_000,
        "jt": 1_000_000,
        "juta": 1_000_000,
        "m": 1_000_000_000,
        "miliar": 1_000_000_000,
        "t": 1_000_000_000_000,
        "triliun": 1_000_000_000_000,
    }

    try:
        val = float(num_s)
    except Exception:
        return None

    mult = mult_map.get(suf)
    if not mult:
        return None

    return int(round(val * mult))


# =========================
# Address FIX
# =========================
def _sanitize_company_address(addr: str) -> str:
    a = (addr or "").strip()
    if not a:
        return "Di tempat"

    low = a.lower()

    bad_patterns = [
        r"tidak\s*dapat\s+menentukan",
        r"tidak\s*bisa\s+menentukan",
        r"tidak\s+dapat\s+menemukan",
        r"tidak\s+bisa\s+menemukan",
        r"tidak\s*ditemukan",
        r"tidak\s*ketemu",
        r"tidak\s+ada\s+informasi",
        r"tidak\s+memiliki\s+informasi",
        r"saya\s+tidak\s+memiliki",
        r"informasi\s+yang\s+cukup",
        r"tidak\s+cukup\s+informasi",
        r"untuk\s+menentukan",
        r"nama\s+tersebut\s+terlalu\s+umum",
        r"terlalu\s+umum",
        r"tidak\s+spesifik",
        r"placeholder",
        r"nama\s+contoh",
        r"banyak\s+perusahaan.*nama\s+serupa",
        r"mungkin\s+menggunakan\s+nama\s+serupa",
        r"maaf",
        r"gagal",
        r"cannot\s+find",
        r"not\s+found",
        r"no\s+information",
        r"no\s+result",
        r"unknown",
    ]

    for p in bad_patterns:
        if re.search(p, low):
            return "Di tempat"

    if len(a) > 120 and not re.search(r"\b(jl|jalan|rt|rw|kec|kel|kab|kota|no\.?|blok|desa)\b", low):
        return "Di tempat"

    if len(a) > 250:
        return "Di tempat"

    return a


def resolve_company_address(company_name: str) -> str:
    addr = ""
    try:
        addr = (search_company_address(company_name) or "").strip()
    except Exception:
        addr = ""
    addr = _sanitize_company_address(addr)
    if addr != "Di tempat":
        return addr

    try:
        addr2 = (search_company_address_ai(company_name) or "").strip()
    except Exception:
        addr2 = ""
    return _sanitize_company_address(addr2)


# =========================
# FIX: replace yang tahan run ter-split
# =========================
def replace_in_runs_keep_format(paragraph, old: str, new: str):
    if not old or not paragraph.text:
        return False
    if old not in paragraph.text:
        return False

    changed = False
    for run in paragraph.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)
            changed = True
    if changed:
        return True

    full = paragraph.text
    replaced = full.replace(old, new)
    if replaced == full:
        return False

    if paragraph.runs:
        paragraph.runs[0].text = replaced
        for r in paragraph.runs[1:]:
            r.text = ""
    else:
        paragraph.add_run(replaced)
    return True

def replace_in_cell_keep_format(cell, old: str, new: str):
    changed = False
    for p in cell.paragraphs:
        if replace_in_runs_keep_format(p, old, new):
            changed = True
    return changed

def replace_everywhere_keep_format(doc, old_list, new_value):
    if new_value is None:
        return
    for p in doc.paragraphs:
        for old in old_list:
            replace_in_runs_keep_format(p, old, new_value)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for old in old_list:
                    replace_in_cell_keep_format(cell, old, new_value)


# =========================
# MoU Counter (as-is)
# =========================

def _mou_counter_path() -> str:
    try:
        folder = str(FILES_DIR)
    except Exception:
        folder = "static/files"
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "mou_counter.json")

def load_mou_counter() -> int:
    path = _mou_counter_path()
    try:
        if not os.path.exists(path):
            return -1
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f) or {}
        return int(data.get("counter", -1))
    except Exception:
        return -1

def save_mou_counter(n: int) -> None:
    path = _mou_counter_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"counter": int(n)}, f)

def get_next_mou_no_depan() -> str:
    n = load_mou_counter() + 1
    save_mou_counter(n)
    return str(n).zfill(3)
