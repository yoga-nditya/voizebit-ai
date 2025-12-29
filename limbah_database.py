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


# ===== NORMALISASI KODE LIMBAH =====
def normalize_limbah_code(text):
    """Normalisasi kode limbah dari berbagai format input"""
    if not text:
        return ""
    
    text = str(text).strip()
    text_clean = re.sub(r'[\s\-_./\\]+', '', text)
    text_clean = text_clean.upper()
    
    strip_words = ['STRIP', 'MINUS', 'MIN', 'DASH', 'SAMPAI', 'HINGGA', 'GARIS']
    for word in strip_words:
        text = re.sub(r'\b' + word + r'\b', '-', text, flags=re.IGNORECASE)
    
    text_clean = re.sub(r'[\s]+', '', text)
    text_clean = re.sub(r'-+', '-', text_clean)
    text_clean = text_clean.upper()
    
    number_words = {
        'NOL': '0', 'KOSONG': '0',
        'SATU': '1', 'SE': '1',
        'DUA': '2', 'TIGA': '3', 'EMPAT': '4',
        'LIMA': '5', 'ENAM': '6', 'TUJUH': '7',
        'DELAPAN': '8', 'SEMBILAN': '9',
        'SEPULUH': '10', 'SEBELAS': '11',
        'SERATUS': '100', 'SERIBU': '1000'
    }
    
    words = text.upper().split()
    converted_parts = []
    skip_next = False
    
    for i, word in enumerate(words):
        if skip_next:
            skip_next = False
            continue
            
        if word in ['SERATUS', 'SEPULUH', 'SEBELAS'] and i + 1 < len(words):
            next_word = words[i + 1]
            if word == 'SERATUS' and next_word in number_words:
                base = 100
                converted_parts.append(str(base + int(number_words.get(next_word, '0'))))
                skip_next = True
                continue
            elif word in ['SEPULUH', 'SEBELAS']:
                converted_parts.append(number_words[word])
                continue
        
        if word in number_words:
            converted_parts.append(number_words[word])
        else:
            cleaned = re.sub(r'[^A-Z0-9\-]', '', word)
            if cleaned:
                converted_parts.append(cleaned)
    
    if converted_parts:
        text_converted = ''.join(converted_parts)
        text_converted = re.sub(r'-+', '-', text_converted)
        
        if re.match(r'^[A-Z]\d', text_converted):
            if not text_converted.endswith(('d', 'D')) and not re.search(r'-\d+$', text_converted):
                if re.match(r'^[A-Z]\d+$', text_converted):
                    text_clean = text_converted + 'd'
                else:
                    text_clean = text_converted
            else:
                text_clean = text_converted
    
    return text_clean


# ===== CARI LIMBAH BY KODE =====
def find_limbah_by_kode(kode_query):
    """Cari limbah berdasarkan kode dengan berbagai variasi"""
    if not kode_query:
        return None, None
    
    kode_normalized = normalize_limbah_code(kode_query)
    print(f"🔍 Mencari kode: '{kode_query}' -> normalized: '{kode_normalized}'")
    
    # 1. Exact match
    if kode_normalized in LIMBAH_B3_DB:
        print(f"✅ Found exact match: {kode_normalized}")
        return kode_normalized, LIMBAH_B3_DB[kode_normalized]
    
    # 2. Case insensitive
    for db_key in LIMBAH_B3_DB.keys():
        if db_key.upper() == kode_normalized.upper():
            print(f"✅ Found case-insensitive match: {db_key}")
            return db_key, LIMBAH_B3_DB[db_key]
    
    # 3. Tanpa suffix 'd'
    if kode_normalized.endswith('d') or kode_normalized.endswith('D'):
        kode_without_d = kode_normalized[:-1]
        for db_key in LIMBAH_B3_DB.keys():
            if db_key.upper() == kode_without_d.upper():
                print(f"✅ Found match without 'd': {db_key}")
                return db_key, LIMBAH_B3_DB[db_key]
    else:
        # 4. Dengan suffix 'd'
        kode_with_d = kode_normalized + 'd'
        if kode_with_d in LIMBAH_B3_DB:
            print(f"✅ Found match with 'd': {kode_with_d}")
            return kode_with_d, LIMBAH_B3_DB[kode_with_d]
        
        for db_key in LIMBAH_B3_DB.keys():
            if db_key.upper() == kode_with_d.upper():
                print(f"✅ Found case-insensitive match with 'd': {db_key}")
                return db_key, LIMBAH_B3_DB[db_key]
    
    # 5. Ignoring dashes
    kode_no_dash = kode_normalized.replace('-', '')
    for db_key in LIMBAH_B3_DB.keys():
        db_key_no_dash = db_key.replace('-', '')
        if db_key_no_dash.upper() == kode_no_dash.upper():
            print(f"✅ Found match ignoring dashes: {db_key}")
            return db_key, LIMBAH_B3_DB[db_key]
    
    print(f"❌ No match found for: {kode_query}")
    return None, None


# ===== CARI LIMBAH BY JENIS =====
def find_limbah_by_jenis(jenis_query):
    """Cari limbah berdasarkan jenis/nama"""
    jenis_lower = jenis_query.lower()

    # Exact match
    for kode, data in LIMBAH_B3_DB.items():
        if data['jenis'].lower() == jenis_lower:
            return kode, data

    # Contains match
    for kode, data in LIMBAH_B3_DB.items():
        if jenis_lower in data['jenis'].lower() or data['jenis'].lower() in jenis_lower:
            return kode, data

    # Keyword match (2+ keywords)
    keywords = jenis_lower.split()
    for kode, data in LIMBAH_B3_DB.items():
        jenis_db_lower = data['jenis'].lower()
        match_count = sum(1 for kw in keywords if kw in jenis_db_lower)
        if match_count >= 2:
            return kode, data

    return None, None


# ===== VOICE TO NUMBER CONVERTER =====
def convert_voice_to_number(text):
    """Konversi text/voice ke angka: 1,7 jt, seratus ribu, dll"""
    if not text:
        return "0"
    
    text_original = str(text).strip()
    text_lower = text_original.lower()
    
    print(f"💬 Converting: '{text_original}'")
    
    # Replace singkatan
    abbreviations = {
        r'\bjt\b': 'juta',
        r'\bjuta\b': 'juta',
        r'\brb\b': 'ribu',
        r'\bk\b': 'ribu',
        r'\bm\b': 'miliar',
        r'\bmi?ly?ar\b': 'miliar',
        r'\bt\b': 'triliun',
    }
    
    for pattern, replacement in abbreviations.items():
        text_lower = re.sub(pattern, replacement, text_lower)
    
    print(f"   After abbreviation: '{text_lower}'")
    
    kata_angka = {
        'nol': 0, 'kosong': 0,
        'satu': 1, 'se': 1,
        'dua': 2, 'tiga': 3, 'empat': 4,
        'lima': 5, 'enam': 6, 'tujuh': 7,
        'delapan': 8, 'sembilan': 9,
        'sepuluh': 10, 'sebelas': 11,
        'dua belas': 12, 'tiga belas': 13,
        'empat belas': 14, 'lima belas': 15,
        'enam belas': 16, 'tujuh belas': 17,
        'delapan belas': 18, 'sembilan belas': 19,
    }
    
    multipliers = {
        'belas': 10,
        'puluh': 10,
        'ratus': 100,
        'ribu': 1000,
        'juta': 1000000,
        'miliar': 1000000000,
        'milyar': 1000000000,
        'milyard': 1000000000,
        'triliun': 1000000000000,
    }
    
    # Handle karakter koma/titik literal (1,7 juta atau 1.7 juta)
    literal_decimal_pattern = r'^(\d+)[,.](\d+)\s*(juta|ribu|miliar|milyar|triliun)?$'
    literal_match = re.match(literal_decimal_pattern, text_lower.strip())
    
    if literal_match:
        before_decimal = literal_match.group(1)
        after_decimal = literal_match.group(2)
        multiplier_word = literal_match.group(3)
        
        print(f"🔢 Detected literal decimal: '{before_decimal}' . '{after_decimal}' {multiplier_word or ''}")
        
        before_value = int(before_decimal)
        after_value = int(after_decimal)
        
        decimal_value = before_value + (after_value / (10 ** len(after_decimal)))
        
        if multiplier_word and multiplier_word in multipliers:
            final_value = int(decimal_value * multipliers[multiplier_word])
        else:
            final_value = int(decimal_value)
        
        print(f"✅ Literal decimal result: {final_value}")
        return str(final_value)
    
    # Handle angka tanpa spasi (1juta, 500ribu)
    no_space_pattern = r'^(\d+)\s*(juta|ribu|miliar|milyar|triliun)$'
    no_space_match = re.match(no_space_pattern, text_lower.strip())
    
    if no_space_match:
        number = int(no_space_match.group(1))
        multiplier_word = no_space_match.group(2)
        
        print(f"🔢 Detected no-space format: '{number}' x '{multiplier_word}'")
        
        if multiplier_word in multipliers:
            final_value = number * multipliers[multiplier_word]
            print(f"✅ No-space result: {final_value}")
            return str(final_value)
    
    # Angka murni
    if re.match(r'^\d[\d\.,]*\d$', text_original) or re.match(r'^\d+$', text_original):
        cleaned = text_original.replace('.', '').replace(',', '')
        print(f"✅ Already number: {cleaned}")
        return cleaned
    
    # Handle kata "koma" atau "titik"
    koma_pattern = r'(.+?)\s*(?:koma|titik)\s*(.+?)(?:\s+(juta|ribu|miliar|milyar|triliun))?$'
    koma_match = re.search(koma_pattern, text_lower)
    
    if koma_match:
        before_koma = koma_match.group(1).strip()
        after_koma = koma_match.group(2).strip()
        multiplier_word = koma_match.group(3)
        
        print(f"🔢 Detected word decimal: '{before_koma}' . '{after_koma}' {multiplier_word or ''}")
        
        if before_koma.isdigit():
            before_value = int(before_koma)
        else:
            before_value = convert_voice_to_number_simple(before_koma)
        
        if after_koma.isdigit():
            after_value = int(after_koma)
        else:
            after_value = convert_voice_to_number_simple(after_koma)
        
        decimal_value = before_value + (after_value / (10 ** len(str(after_value))))
        
        if multiplier_word and multiplier_word in multipliers:
            final_value = int(decimal_value * multipliers[multiplier_word])
        else:
            final_value = int(decimal_value)
        
        print(f"✅ Word decimal result: {final_value}")
        return str(final_value)
    
    # Handle format "X juta Y ribu"
    result = 0
    temp = 0
    
    for phrase, num in kata_angka.items():
        if ' ' in phrase:
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
        
        # Handle "seratus" khusus
        if word == 'seratus':
            temp = 100
            i += 1
            continue
        
        if word == 'belas':
            if temp == 0:
                temp = 1
            temp += 10
            i += 1
            continue
        
        if word == 'puluh':
            if temp == 0:
                temp = 1
            temp *= 10
            i += 1
            continue
        
        if word == 'ratus':
            if temp == 0:
                temp = 1
            temp *= 100
            i += 1
            continue
        
        if word in ['ribu', 'juta', 'miliar', 'milyar', 'triliun']:
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
        print(f"✅ Voice conversion result: {result}")
        return str(result)
    
    numbers = re.findall(r'\d+', text_original)
    if numbers:
        result = ''.join(numbers)
        print(f"✅ Extracted numbers: {result}")
        return result
    
    print(f"⚠️  Could not convert: {text_original}")
    return text_original


def convert_voice_to_number_simple(text):
    """Helper untuk konversi sederhana"""
    kata_angka = {
        'nol': 0, 'kosong': 0,
        'satu': 1, 'se': 1,
        'dua': 2, 'tiga': 3, 'empat': 4,
        'lima': 5, 'enam': 6, 'tujuh': 7,
        'delapan': 8, 'sembilan': 9,
        'sepuluh': 10, 'sebelas': 11,
    }
    
    text = str(text).lower().strip()
    
    if text.isdigit():
        return int(text)
    
    if text == 'seratus':
        return 100
    
    if text == 'seribu':
        return 1000
    
    if text in kata_angka:
        return kata_angka[text]
    
    numbers = re.findall(r'\d+', text)
    if numbers:
        return int(numbers[0])
    
    return 0


def parse_termin_days(text: str, default: int = 14, min_days: int = 1, max_days: int = 365) -> str:
    """Parse termin days dari input"""
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


def angka_ke_terbilang(angka):
    """Convert angka ke terbilang Indonesia"""
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


def format_rupiah(angka_str):
    """Format angka jadi Rupiah dengan titik"""
    angka_clean = re.sub(r'[^\d]', '', str(angka_str))
    if not angka_clean:
        return angka_str
    try:
        angka_int = int(angka_clean)
        formatted = f"{angka_int:,}".replace(',', '.')
        return formatted
    except:
        return angka_str


def angka_ke_romawi(bulan):
    """Convert bulan ke romawi"""
    romawi = {
        '1': 'I', '2': 'II', '3': 'III', '4': 'IV', '5': 'V', '6': 'VI',
        '7': 'VII', '8': 'VIII', '9': 'IX', '10': 'X', '11': 'XI', '12': 'XII',
        '01': 'I', '02': 'II', '03': 'III', '04': 'IV', '05': 'V', '06': 'VI',
        '07': 'VII', '08': 'VIII', '09': 'IX'
    }
    return romawi.get(str(bulan), 'I')