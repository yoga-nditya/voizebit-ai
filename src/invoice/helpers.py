import os
import json
import re
from datetime import datetime

from limbah_database import convert_voice_to_number
from utils import search_company_address, search_company_address_ai
from config_new import FILES_DIR


def is_non_b3_input(text: str) -> bool:
    if not text:
        return False
    t = text.strip().lower()
    norm = re.sub(r"[\s\-_]+", "", t)
    return norm in ("nonb3", "nonbii3") or norm.startswith("nonb3")


def normalize_id_number_text(text: str) -> str:
    if not text:
        return text
    t = text.strip()
    t = re.sub(r"(?<=\d)[\.,](?=\d{3}(\D|$))", "", t)
    t = re.sub(r"(?<=\d),(?=\d)", ".", t)
    return t


# =========================
# ✅ parse nominal singkat: 1.5jt, 250rb, 10k, 2m, 1.2t, dll
# =========================
def parse_id_short_amount(text: str):
    """
    Contoh:
      "1.5jt" / "1,5jt" / "1.5 juta" -> 1500000
      "250rb" / "250 ribu" -> 250000
      "10k" -> 10000
      "2m" / "2 miliar" -> 2000000000
      "1.2t" / "1,2 triliun" -> 1200000000000
      "Rp 1.5jt" -> 1500000
    Return int atau None.
    """
    if text is None:
        return None

    raw = str(text).strip().lower()
    if not raw:
        return None

    raw = raw.replace("rupiah", "").replace("idr", "").strip()
    raw = re.sub(r"^\s*rp\.?\s*", "", raw)

    t = re.sub(r"\s+", " ", raw).strip()

    m = re.match(
        r"^(\d+(?:[.,]\d+)?)\s*(k|rb|ribu|jt|juta|m|miliar|t|triliun)\b",
        t,
        flags=re.IGNORECASE
    )
    if not m:
        t2 = re.sub(r"\s+", "", raw)
        m = re.match(
            r"^(\d+(?:[.,]\d+)?)(k|rb|ribu|jt|juta|m|miliar|t|triliun)\b",
            t2,
            flags=re.IGNORECASE
        )
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

    mult = mult_map.get(suf)
    if not mult:
        return None

    try:
        val = float(num_s)
        return int(round(val * mult))
    except:
        return None


# =========================
# Validasi input nominal
# =========================
_ID_NUMBER_HINT_WORDS = set([
    "nol", "kosong",
    "satu", "se", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan",
    "sepuluh", "sebelas", "belas", "puluh", "ratus",
    "seratus", "seribu",
    "ribu", "juta", "miliar", "triliun",
    "koma",
    "rupiah", "rp", "idr",
    "jt", "rb", "k", "m", "t",
])

def _is_zero_like_input(text: str) -> bool:
    if not text:
        return False
    t = text.strip().lower()
    t = re.sub(r"\s+", " ", t)
    if t in ("0", "nol", "kosong"):
        return True
    if re.fullmatch(r"0+(?:[.,]0+)?", t):
        return True
    if re.fullmatch(r"(rp|idr)\s*0+(?:[.,]0+)?", t):
        return True
    if re.fullmatch(r"0+(?:[.,]0+)?\s*(rp|idr|rupiah)", t):
        return True
    if "nol" in t or "kosong" in t:
        return True
    return False

def _is_amount_like_input(text: str) -> bool:
    if not text:
        return False
    t = text.strip().lower()
    if re.search(r"\d", t):
        return True
    tokens = re.findall(r"[a-zA-Z]+", t)
    for w in tokens:
        if w in _ID_NUMBER_HINT_WORDS:
            return True
    return False


# =========================
# words to number ID
# =========================
_ID_SMALL = {
    "nol": 0, "kosong": 0,
    "satu": 1, "se": 1,
    "dua": 2, "tiga": 3, "empat": 4, "lima": 5,
    "enam": 6, "tujuh": 7, "delapan": 8, "sembilan": 9,
    "sepuluh": 10, "sebelas": 11,
}

_ID_SCALES = {
    "ribu": 1_000,
    "juta": 1_000_000,
    "miliar": 1_000_000_000,
    "triliun": 1_000_000_000_000,
}

def _tokenize_id_words(s: str):
    s = (s or "").lower().strip()
    s = re.sub(r"[-_]", " ", s)
    s = re.sub(r"[^a-z0-9\s,\.]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s.split()

def _parse_id_integer_words(tokens):
    total = 0
    current = 0
    i = 0
    while i < len(tokens):
        w = tokens[i]

        if re.fullmatch(r"\d+(?:\.\d+)?", w):
            try:
                current += int(float(w))
            except:
                pass
            i += 1
            continue

        if w == "seratus":
            current += 100
            i += 1
            continue

        if w == "seribu":
            total += (current if current else 1) * 1000
            current = 0
            i += 1
            continue

        if w in _ID_SMALL:
            val = _ID_SMALL[w]

            if i + 1 < len(tokens) and tokens[i + 1] == "belas":
                current += 10 + val
                i += 2
                continue

            if i + 1 < len(tokens) and tokens[i + 1] == "puluh":
                current += val * 10
                i += 2
                continue

            if i + 1 < len(tokens) and tokens[i + 1] == "ratus":
                current += val * 100
                i += 2
                continue

            current += val
            i += 1
            continue

        if w in _ID_SCALES:
            scale = _ID_SCALES[w]
            if current == 0:
                current = 1
            total += current * scale
            current = 0
            i += 1
            continue

        i += 1

    return total + current

def words_to_number_id(text: str):
    if not text:
        return None

    raw = text.strip().lower()

    short_val = parse_id_short_amount(raw)
    if short_val is not None:
        return float(short_val)

    norm = normalize_id_number_text(raw)
    if re.fullmatch(r"\d+(?:\.\d+)?", norm):
        try:
            return float(norm)
        except:
            pass

    tokens = _tokenize_id_words(raw)
    if not tokens:
        return None

    if "koma" in tokens:
        k = tokens.index("koma")
        left = tokens[:k]
        right = tokens[k + 1:]

        left_int = _parse_id_integer_words(left) if left else 0

        if not right:
            return float(left_int) + 0.5

        digits = []
        for w in right:
            if w in _ID_SMALL:
                digits.append(str(_ID_SMALL[w]))
                continue
            if re.fullmatch(r"\d+", w):
                digits.append(w)
                continue
            if w in _ID_SCALES:
                break

        frac_str = "".join(digits) if digits else "5"
        frac_val = float("0." + frac_str)

        scale = None
        for w in right:
            if w in _ID_SCALES:
                scale = _ID_SCALES[w]
                break

        val = float(left_int) + frac_val
        if scale:
            val *= scale
        return val

    return float(_parse_id_integer_words(tokens))


def parse_amount_id(text: str):
    """
    Return int jika valid; None jika invalid (contoh: 'p').
    """
    if text is None:
        return None

    raw = str(text).strip().lower()
    if raw == "":
        return None

    if not _is_amount_like_input(raw):
        return None

    if _is_zero_like_input(raw):
        return 0

    short_val = parse_id_short_amount(raw)
    if short_val is not None:
        return int(short_val)

    wv = words_to_number_id(raw)
    if wv is not None:
        try:
            return int(round(float(wv)))
        except:
            pass

    tnorm = normalize_id_number_text(str(text))
    val = convert_voice_to_number(tnorm)

    if val is None:
        digits = re.sub(r"\D+", "", tnorm)
        if digits:
            try:
                return int(digits)
            except:
                return None
        return None

    try:
        return int(round(float(val)))
    except:
        digits = re.sub(r"\D+", "", str(val))
        return int(digits) if digits else None


def parse_qty_id(text: str) -> float:
    if not text:
        return 0.0

    raw = text.strip().lower()
    wv = words_to_number_id(raw)
    if wv is not None:
        try:
            return float(wv)
        except:
            pass

    t = normalize_id_number_text(text)
    v = convert_voice_to_number(t)
    try:
        return float(v)
    except:
        m = re.findall(r"\d+(?:\.\d+)?", t)
        return float(m[0]) if m else 0.0


def normalize_voice_strip(text: str) -> str:
    if not text:
        return text
    t = text.strip()
    low = t.lower()
    if low == "strip":
        return "-"
    if re.fullmatch(r"(.*\b)?strip(\b.*)?", low):
        return "-"
    return text


# =========================
# Address sanitize
# =========================
def _sanitize_company_address(addr: str) -> str:
    a = (addr or "").strip()
    if not a:
        return "Di tempat"

    low = a.lower()

    bad_patterns = [
        r"tidak\s*dapat\s+menentukan",
        r"tidak\s*bisa\s+menentukan",
        r"tidak\s*dapat\s+menemukan",
        r"tidak\s*bisa\s+menemukan",
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
    except:
        addr = ""
    addr = _sanitize_company_address(addr)
    if addr != "Di tempat":
        return addr

    try:
        addr2 = (search_company_address_ai(company_name) or "").strip()
    except:
        addr2 = ""
    return _sanitize_company_address(addr2)


# =========================
# Filename unique
# =========================
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


# =========================
# Invoice counter
# =========================
def _invoice_counter_path() -> str:
    try:
        folder = str(FILES_DIR)
    except Exception:
        folder = "static/files"
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "invoice_counter.json")


def load_invoice_counter(prefix: str) -> int:
    path = _invoice_counter_path()
    try:
        if not os.path.exists(path):
            return 0
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f) or {}
        counters = data.get("counters") or {}
        return int(counters.get(prefix, 0))
    except:
        return 0


def save_invoice_counter(prefix: str, n: int) -> None:
    path = _invoice_counter_path()
    data = {}
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f) or {}
    except:
        data = {}

    counters = data.get("counters") or {}
    counters[prefix] = int(n)
    data["counters"] = counters

    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)


def get_next_invoice_no() -> str:
    now = datetime.now()
    prefix = now.strftime("%d%m")
    n = load_invoice_counter(prefix) + 1
    save_invoice_counter(prefix, n)
    return f"{prefix}{str(n).zfill(3)}"
