"""
Kode Limbah Module
Berisi: Database Limbah B3, Search Functions, dan Text/Voice Converters
"""

import re

# ===== DATABASE LIMBAH B3 =====
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

# ===== NORMALISASI KODE LIMBAH (FIX STRIP/MINUS + ANGKA KATA) =====
_NUM_WORDS_ID = {
    "nol": "0", "kosong": "0",
    "satu": "1", "se": "1",
    "dua": "2", "tiga": "3", "empat": "4",
    "lima": "5", "enam": "6", "tujuh": "7",
    "delapan": "8", "sembilan": "9",
    "sepuluh": "10", "sebelas": "11",
}

_DASH_WORDS = [
    "strip", "strips", "minus", "min", "dash", "garis", "tanda minus", "tanda strip",
    "hyphen", "penghubung"
]

def normalize_limbah_code(text: str) -> str:
    """
    Contoh input yang harus lolos:
      - "A336-1"
      - "A336 1"
      - "A336 strip 1"
      - "A336 strip satu"
      - "a 336 minus dua"
      - "A303 2" -> "A303-2" (kalau ada di DB)
      - "A102" -> bisa match A102d (via find_limbah_by_kode)
    """
    if not text:
        return ""

    raw = str(text).strip()
    if not raw:
        return ""

    s = raw.lower()

    # 1) ubah kata strip/minus/dash -> "-"
    for w in _DASH_WORDS:
        s = re.sub(rf"\b{re.escape(w)}\b", "-", s, flags=re.IGNORECASE)

    # 2) rapikan spasi di sekitar dash
    s = re.sub(r"\s*-\s*", "-", s)

    # 3) ubah angka kata (khususnya setelah dash / atau berdiri sendiri)
    #    contoh: "A336-satu" -> "A336-1"
    for k, v in _NUM_WORDS_ID.items():
        s = re.sub(rf"\b{re.escape(k)}\b", v, s)

    # 4) buang karakter aneh selain huruf/angka/dash
    s = re.sub(r"[^a-z0-9\- ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()

    # 5) gabungkan token jadi bentuk kode
    #    - jika ada dash, biarkan dash
    #    - jika tidak ada dash tapi pola "A336 1" -> "A336-1"
    tokens = s.split()

    if not tokens:
        return ""

    joined = "".join(tokens)  # keep dash already in token? (dash tidak hilang karena bukan spasi)
    # kalau dash sempat hilang karena tokenization, pastikan dari s asli:
    joined = s.replace(" ", "")
    joined = re.sub(r"-+", "-", joined)

    # 6) Pola "A3361" (tanpa dash) tapi DB ada "A336-1"
    #    jika awal huruf + angka panjang + 1 digit terakhir => sisipkan dash sebelum digit terakhir
    m = re.match(r"^([a-z])(\d{3,})(\d)$", joined)
    if m:
        joined_candidate = f"{m.group(1)}{m.group(2)}-{m.group(3)}"
        joined = joined_candidate

    return joined.upper()

# ===== CARI LIMBAH BY KODE =====
def find_limbah_by_kode(kode_query):
    if not kode_query:
        return None, None

    kode_normalized = normalize_limbah_code(kode_query)

    # 1) exact
    if kode_normalized in LIMBAH_B3_DB:
        return kode_normalized, LIMBAH_B3_DB[kode_normalized]

    # 2) case-insensitive
    for db_key in LIMBAH_B3_DB.keys():
        if db_key.upper() == kode_normalized.upper():
            return db_key, LIMBAH_B3_DB[db_key]

    # 3) tanpa suffix 'd' -> coba match (misal input A102d atau A102)
    if kode_normalized.endswith("D"):
        kode_without_d = kode_normalized[:-1]
        for db_key in LIMBAH_B3_DB.keys():
            if db_key.upper() == kode_without_d.upper():
                return db_key, LIMBAH_B3_DB[db_key]
    else:
        kode_with_d = kode_normalized + "D"
        if kode_with_d in LIMBAH_B3_DB:
            return kode_with_d, LIMBAH_B3_DB[kode_with_d]
        for db_key in LIMBAH_B3_DB.keys():
            if db_key.upper() == kode_with_d.upper():
                return db_key, LIMBAH_B3_DB[db_key]

    # 4) ignoring dashes
    kode_no_dash = kode_normalized.replace("-", "")
    for db_key in LIMBAH_B3_DB.keys():
        if db_key.replace("-", "").upper() == kode_no_dash.upper():
            return db_key, LIMBAH_B3_DB[db_key]

    # 5) kalau user ngetik "A336 1" jadi "A3361", coba bentuk dash umum: A336-1
    m2 = re.match(r"^([A-Z])(\d{3,})(\d)$", kode_no_dash.upper())
    if m2:
        candidate = f"{m2.group(1)}{m2.group(2)}-{m2.group(3)}"
        if candidate in LIMBAH_B3_DB:
            return candidate, LIMBAH_B3_DB[candidate]

    return None, None

# ===== CARI LIMBAH BY JENIS =====
def find_limbah_by_jenis(jenis_query):
    if not jenis_query:
        return None, None

    jenis_lower = str(jenis_query).lower().strip()
    if not jenis_lower:
        return None, None

    # Exact match
    for kode, data in LIMBAH_B3_DB.items():
        if data["jenis"].lower() == jenis_lower:
            return kode, data

    # Contains match
    for kode, data in LIMBAH_B3_DB.items():
        if jenis_lower in data["jenis"].lower() or data["jenis"].lower() in jenis_lower:
            return kode, data

    # Keyword match (>=2)
    keywords = jenis_lower.split()
    for kode, data in LIMBAH_B3_DB.items():
        jenis_db_lower = data["jenis"].lower()
        match_count = sum(1 for kw in keywords if kw in jenis_db_lower)
        if match_count >= 2:
            return kode, data

    return None, None


# ===== VOICE TO NUMBER CONVERTER =====
def convert_voice_to_number(text):
    # (Pakai versi kamu, saya tidak ubah logic besarnya — hanya rapikan sedikit)
    if not text:
        return "0"

    text_original = str(text).strip()
    text_lower = text_original.lower()

    abbreviations = {
        r"\bjt\b": "juta",
        r"\bjuta\b": "juta",
        r"\brb\b": "ribu",
        r"\bk\b": "ribu",
        r"\bm\b": "miliar",
        r"\bmi?ly?ar\b": "miliar",
        r"\bt\b": "triliun",
    }
    for pattern, replacement in abbreviations.items():
        text_lower = re.sub(pattern, replacement, text_lower)

    kata_angka = {
        "nol": 0, "kosong": 0,
        "satu": 1, "se": 1,
        "dua": 2, "tiga": 3, "empat": 4,
        "lima": 5, "enam": 6, "tujuh": 7,
        "delapan": 8, "sembilan": 9,
        "sepuluh": 10, "sebelas": 11,
        "dua belas": 12, "tiga belas": 13,
        "empat belas": 14, "lima belas": 15,
        "enam belas": 16, "tujuh belas": 17,
        "delapan belas": 18, "sembilan belas": 19,
    }

    multipliers = {
        "belas": 10,
        "puluh": 10,
        "ratus": 100,
        "ribu": 1000,
        "juta": 1000000,
        "miliar": 1000000000,
        "milyar": 1000000000,
        "milyard": 1000000000,
        "triliun": 1000000000000,
    }

    # literal decimal (1,7 juta)
    literal_decimal_pattern = r"^(\d+)[,.](\d+)\s*(juta|ribu|miliar|milyar|triliun)?$"
    literal_match = re.match(literal_decimal_pattern, text_lower.strip())
    if literal_match:
        before_decimal = literal_match.group(1)
        after_decimal = literal_match.group(2)
        multiplier_word = literal_match.group(3)
        before_value = int(before_decimal)
        after_value = int(after_decimal)
        decimal_value = before_value + (after_value / (10 ** len(after_decimal)))
        if multiplier_word and multiplier_word in multipliers:
            final_value = int(decimal_value * multipliers[multiplier_word])
        else:
            final_value = int(decimal_value)
        return str(final_value)

    # no-space (500ribu)
    no_space_pattern = r"^(\d+)\s*(juta|ribu|miliar|milyar|triliun)$"
    no_space_match = re.match(no_space_pattern, text_lower.strip())
    if no_space_match:
        number = int(no_space_match.group(1))
        multiplier_word = no_space_match.group(2)
        if multiplier_word in multipliers:
            return str(number * multipliers[multiplier_word])

    # angka murni
    if re.match(r"^\d[\d\.,]*\d$", text_original) or re.match(r"^\d+$", text_original):
        cleaned = text_original.replace(".", "").replace(",", "")
        return cleaned

    # kata "koma/titik"
    koma_pattern = r"(.+?)\s*(?:koma|titik)\s*(.+?)(?:\s+(juta|ribu|miliar|milyar|triliun))?$"
    koma_match = re.search(koma_pattern, text_lower)
    if koma_match:
        before_koma = koma_match.group(1).strip()
        after_koma = koma_match.group(2).strip()
        multiplier_word = koma_match.group(3)

        before_value = int(before_koma) if before_koma.isdigit() else convert_voice_to_number_simple(before_koma)
        after_value = int(after_koma) if after_koma.isdigit() else convert_voice_to_number_simple(after_koma)

        decimal_value = before_value + (after_value / (10 ** len(str(after_value))))
        if multiplier_word and multiplier_word in multipliers:
            final_value = int(decimal_value * multipliers[multiplier_word])
        else:
            final_value = int(decimal_value)
        return str(final_value)

    # umum
    result = 0
    temp = 0

    for phrase, num in kata_angka.items():
        if " " in phrase:
            text_lower = text_lower.replace(phrase, str(num))

    words = text_lower.split()
    i = 0
    while i < len(words):
        word = words[i]

        if word.isdigit():
            temp += int(word)
            i += 1
            continue

        if word in kata_angka:
            temp += kata_angka[word]
            i += 1
            continue

        if word == "seratus":
            temp = 100
            i += 1
            continue

        if word == "belas":
            if temp == 0:
                temp = 1
            temp += 10
            i += 1
            continue

        if word == "puluh":
            if temp == 0:
                temp = 1
            temp *= 10
            i += 1
            continue

        if word == "ratus":
            if temp == 0:
                temp = 1
            temp *= 100
            i += 1
            continue

        if word in ["ribu", "juta", "miliar", "milyar", "triliun"]:
            if temp == 0:
                temp = 1
            temp *= multipliers[word]
            result += temp
            temp = 0
            i += 1
            continue

        i += 1

    result += temp
    if result > 0:
        return str(result)

    numbers = re.findall(r"\d+", text_original)
    if numbers:
        return "".join(numbers)

    return text_original


def convert_voice_to_number_simple(text):
    kata_angka = {
        "nol": 0, "kosong": 0,
        "satu": 1, "se": 1,
        "dua": 2, "tiga": 3, "empat": 4,
        "lima": 5, "enam": 6, "tujuh": 7,
        "delapan": 8, "sembilan": 9,
        "sepuluh": 10, "sebelas": 11,
    }
    t = str(text).lower().strip()
    if t.isdigit():
        return int(t)
    if t == "seratus":
        return 100
    if t == "seribu":
        return 1000
    if t in kata_angka:
        return kata_angka[t]
    numbers = re.findall(r"\d+", t)
    if numbers:
        return int(numbers[0])
    return 0


def parse_termin_days(text: str, default: int = 14, min_days: int = 1, max_days: int = 365) -> str:
    s = (text or "").strip()
    if not s:
        return str(default)

    converted = convert_voice_to_number(s)
    m = re.search(r"\d+", str(converted).replace(".", "").replace(",", ""))
    if not m:
        return str(default)

    val = int(m.group(0))
    if val < min_days or val > max_days:
        return str(default)
    return str(val)


def angka_ke_terbilang(angka):
    try:
        n = int(angka)
    except:
        return "empat belas"

    if n == 0:
        return "nol"

    satuan = ["", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan"]
    belasan = ["sepuluh", "sebelas", "dua belas", "tiga belas", "empat belas", "lima belas",
               "enam belas", "tujuh belas", "delapan belas", "sembilan belas"]

    if n < 10:
        return satuan[n]
    elif n < 20:
        return belasan[n - 10]
    elif n < 100:
        puluhan = n // 10
        sisanya = n % 10
        if sisanya == 0:
            return satuan[puluhan] + " puluh"
        else:
            return satuan[puluhan] + " puluh " + satuan[sisanya]
    elif n < 1000:
        ratusan = n // 100
        sisanya = n % 100
        if ratusan == 1:
            result = "seratus"
        else:
            result = satuan[ratusan] + " ratus"
        if sisanya > 0:
            result += " " + angka_ke_terbilang(sisanya)
        return result

    return str(n)


def format_rupiah(angka_str):
    angka_clean = re.sub(r"[^\d]", "", str(angka_str))
    if not angka_clean:
        return angka_str
    try:
        angka_int = int(angka_clean)
        return f"{angka_int:,}".replace(",", ".")
    except:
        return angka_str


def angka_ke_romawi(bulan):
    romawi = {
        "1": "I", "2": "II", "3": "III", "4": "IV", "5": "V", "6": "VI",
        "7": "VII", "8": "VIII", "9": "IX", "10": "X", "11": "XI", "12": "XII",
        "01": "I", "02": "II", "03": "III", "04": "IV", "05": "V", "06": "VI",
        "07": "VII", "08": "VIII", "09": "IX",
    }
    return romawi.get(str(bulan), "I")
