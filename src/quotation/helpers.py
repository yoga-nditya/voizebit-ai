import re

from limbah_database import convert_voice_to_number


# =========================
# Helpers
# =========================

def is_non_b3_input(text: str) -> bool:
    if not text:
        return False
    t = text.strip().lower()
    norm = re.sub(r'[\s\-_]+', '', t)
    return norm in ("nonb3", "nonbii3") or norm.startswith("nonb3")


def normalize_id_number_text(text: str) -> str:
    if not text:
        return text
    t = text.strip()
    # hapus pemisah ribuan ala ID: 1.500.000 atau 1,500,000
    t = re.sub(r'(?<=\d)[\.,](?=\d{3}(\D|$))', '', t)
    # ubah koma desimal jadi titik
    t = re.sub(r'(?<=\d),(?=\d)', '.', t)
    return t


# =========================
# ✅ FIX: parse_amount_id agar "1.5jt" -> 1.500.000 (bukan 5)
# =========================

def parse_amount_id(text: str) -> int:
    """
    Parse nominal Indonesia dari:
    - 1.5jt, 1,5jt, 1.5 juta, 1,5 juta -> 1_500_000
    - 250rb, 250 rb, 250ribu -> 250_000
    - 2m, 2 miliar -> 2_000_000_000
    - 1.2T, 1,2 triliun -> 1_200_000_000_000
    - 10k -> 10_000
    - "1 koma 5 juta" -> 1_500_000
    - angka biasa: "1.500.000" / "1500000" -> 1_500_000
    - teks via convert_voice_to_number tetap didukung
    """
    if not text:
        return 0

    raw = text.strip()
    lower = raw.lower().strip()

    # --- 1) Tangkap format compact: "1.5jt", "1,5jt", "250rb", "2m", "1.2t", "10k"
    compact_re = re.compile(
        r"^\s*(\d+(?:[.,]\d+)?)\s*(k|rb|ribu|jt|juta|m|miliar|t|triliun)\b",
        re.IGNORECASE
    )
    m = compact_re.search(lower)
    if m:
        num_str = m.group(1).replace(",", ".")
        suf = m.group(2).lower()

        scale_map = {
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
            base = float(num_str)
            return int(round(base * scale_map.get(suf, 1)))
        except Exception:
            pass

    # --- 2) Deteksi skala kata (untuk "15 juta", "1 koma 5 juta", dst.)
    scale_map_words = {
        "ribu": 1_000,
        "juta": 1_000_000,
        "miliar": 1_000_000_000,
        "triliun": 1_000_000_000_000,
    }
    scale = None
    for k, mul in scale_map_words.items():
        if re.search(rf"\b{k}\b", lower):
            scale = mul
            break

    # --- 3) Tangkap format "koma": "1 koma 5 juta" / "15 koma 2 ribu"
    if "koma" in lower:
        parts = re.split(r"\bkoma\b", lower, maxsplit=1)
        left_part = parts[0].strip()
        right_part = parts[1].strip() if len(parts) > 1 else ""

        left_tokens = re.findall(r"[a-zA-Z0-9]+", left_part)
        right_tokens = re.findall(r"[a-zA-Z0-9]+", right_part)

        digit_map = {
            "nol": 0, "kosong": 0,
            "satu": 1, "se": 1,
            "dua": 2, "tiga": 3, "empat": 4, "lima": 5,
            "enam": 6, "tujuh": 7, "delapan": 8, "sembilan": 9
        }

        def token_to_num_str(tok: str):
            tok = tok.lower().strip()
            if tok.isdigit():
                return tok
            if tok in digit_map:
                return str(digit_map[tok])
            return None

        left_num = token_to_num_str(left_tokens[-1]) if left_tokens else None
        right_num = token_to_num_str(right_tokens[0]) if right_tokens else None

        if left_num is not None and right_num is not None:
            try:
                val = float(f"{left_num}.{right_num}")
                if scale:
                    val *= scale
                return int(round(val))
            except Exception:
                pass

    # --- 4) Fallback: normalize angka lalu lempar ke convert_voice_to_number
    tnorm = normalize_id_number_text(raw)
    val = convert_voice_to_number(tnorm)
    if val is None:
        val = 0

    # kalau ada kata skala tapi val kecil, kalikan (mis. "15 juta" jadi 15)
    if scale:
        try:
            f = float(val)
            if f < scale:
                val = f * scale
        except Exception:
            pass

    try:
        return int(round(float(val)))
    except Exception:
        digits = re.sub(r"\D+", "", str(val))
        return int(digits) if digits else 0


# ✅ FIX: Anggap hasil pencarian alamat yang berupa penolakan/ketidakpastian AI sebagai "tidak ditemukan"
def normalize_found_address(addr: str) -> str:
    if not addr:
        return ""
    t = str(addr).strip()
    if not t:
        return ""

    lower = t.lower()

    not_found_patterns = [
        r"tidak\s*ditemukan",
        r"tidak\s*ketemu",
        r"nggak\s*ditemukan",
        r"gak\s*ditemukan",
        r"ga\s*ditemukan",
        r"tidak\s*ada",
        r"belum\s*ada",
        r"unknown",
        r"not\s*found",
        r"no\s*result",
        r"tidak\s*tersedia",
        r"data\s*tidak\s*tersedia",
    ]

    ai_refusal_patterns = [
        r"saya\s+tidak\s+memiliki\s+informasi",
        r"informasi\s+yang\s+cukup",
        r"tidak\s+cukup\s+informasi",
        r"tidak\s+dapat\s+menentukan",
        r"tidak\s+bisa\s+menentukan",
        r"nama\s+tersebut\s+terlalu\s+umum",
        r"terlalu\s+umum",
        r"tidak\s+spesifik",
        r"nama\s+contoh",
        r"placeholder",
        r"banyak\s+perusahaan.*nama\s+serupa",
        r"mungkin\s+menggunakan\s+nama\s+serupa",
    ]

    for p in not_found_patterns:
        if re.search(p, lower):
            return ""

    for p in ai_refusal_patterns:
        if re.search(p, lower):
            return ""

    if len(t) > 120 and ("jalan" not in lower and "jl" not in lower and "rt" not in lower and "rw" not in lower):
        return ""

    if len(t) <= 6 and lower in ("-", "none", "null", "n/a"):
        return ""

    return t
