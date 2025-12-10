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

try:
    from docx2pdf import convert as docx_to_pdf
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

load_dotenv()

OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "openai/gpt-4o-mini")

BASE_DIR = Path(__file__).parent
FILES_DIR = BASE_DIR / "static" / "files"
TEMPLATE_FILE = BASE_DIR / "template_quotation.docx"
TEMP_DIR = BASE_DIR / "temp"
FILES_DIR.mkdir(parents=True, exist_ok=True)
TEMP_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = "karya-limbah-2025"

conversations = {}

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
    "A341-1": {"jenis": "Residu produksi dan konsentrat (Sabun deterjen, kosmetik)", "satuan": "Kg", "karakteristik": "Beracun"},
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
    """Cari kode limbah berdasarkan nama/jenis limbah"""
    jenis_lower = jenis_query.lower()
    
    # Exact match
    for kode, data in LIMBAH_B3_DB.items():
        if data['jenis'].lower() == jenis_lower:
            return kode, data
    
    # Partial match
    for kode, data in LIMBAH_B3_DB.items():
        if jenis_lower in data['jenis'].lower() or data['jenis'].lower() in jenis_lower:
            return kode, data
    
    # Keyword match
    keywords = jenis_lower.split()
    for kode, data in LIMBAH_B3_DB.items():
        jenis_db_lower = data['jenis'].lower()
        match_count = sum(1 for kw in keywords if kw in jenis_db_lower)
        if match_count >= 2:  # At least 2 keywords match
            return kode, data
    
    return None, None

def find_limbah_by_kode(kode_query):
    """Cari jenis limbah berdasarkan kode"""
    kode_upper = kode_query.upper().strip()
    
    if kode_upper in LIMBAH_B3_DB:
        return kode_upper, LIMBAH_B3_DB[kode_upper]
    
    return None, None

def angka_ke_romawi(bulan):
    romawi = {
        '1': 'I', '2': 'II', '3': 'III', '4': 'IV', '5': 'V', '6': 'VI',
        '7': 'VII', '8': 'VIII', '9': 'IX', '10': 'X', '11': 'XI', '12': 'XII',
        '01': 'I', '02': 'II', '03': 'III', '04': 'IV', '05': 'V', '06': 'VI',
        '07': 'VII', '08': 'VIII', '09': 'IX'
    }
    return romawi.get(str(bulan), 'I')

def format_tanggal_indonesia():
    bulan_id = {
        '01': 'Januari', '02': 'Februari', '03': 'Maret', '04': 'April', 
        '05': 'Mei', '06': 'Juni', '07': 'Juli', '08': 'Agustus', 
        '09': 'September', '10': 'Oktober', '11': 'November', '12': 'Desember'
    }
    now = datetime.now()
    return f"Tangerang, {now.day} {bulan_id[now.strftime('%m')]} {now.year}"

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
        
        doc_content = doc_content.replace('>027</w:t>', f'>{data.get("nomor_depan", "002")}</w:t>', 1)
        doc_content = doc_content.replace('>IX</w:t>', f'>{bulan_romawi}</w:t>', 1)
        doc_content = doc_content.replace('>PT Surgika Alkesindo, </w:t>', f'>{escape_xml(nama_perusahaan)}, </w:t>')
        doc_content = doc_content.replace('>PT. Surgika Alkesindo</w:t>', f'>{escape_xml(nama_perusahaan)}</w:t>')
        
        old_alamat = 'Jl Plumpang Semper No.6A RT.12/Rw.2, Tugu Utara, Kec. Koja, Jakarta Utara, DKI Jakarta 14260'
        doc_content = doc_content.replace(f'>{old_alamat}</w:t>', f'>{escape_xml(alamat_perusahaan)}</w:t>')
        doc_content = doc_content.replace('>28 November </w:t>', f'>{now.day} {bulan_id[now.strftime("%m")]} </w:t>', 1)
        
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
            '>027</w:t>': f'>{data.get("nomor_depan", "002")}</w:t>',
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

def create_pdf(filename):
    if not PDF_AVAILABLE:
        return None
    
    docx_path = FILES_DIR / f"{filename}.docx"
    pdf_path = FILES_DIR / f"{filename}.pdf"
    
    try:
        docx_to_pdf(str(docx_path), str(pdf_path))
        
        if pdf_path.exists() and pdf_path.stat().st_size > 0:
            return f"{filename}.pdf"
        else:
            return None
    except Exception as e:
        print(f"PDF Error: {e}")
        return None

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/chat", methods=["POST"])
def chat():
    try:
        data = request.get_json()
        text = data.get("message", "").strip()
        
        if not text:
            return jsonify({"error": "Pesan kosong"}), 400
        
        sid = session.get('sid')
        if not sid:
            sid = str(uuid.uuid4())
            session['sid'] = sid
        
        state = conversations.get(sid, {'step': 'idle', 'data': {}})
        lower = text.lower()
        
        # Start conversation
        if 'quotation' in lower or 'penawaran' in lower or 'buat' in lower:
            state['step'] = 'nomor_depan'
            now = datetime.now()
            state['data'] = {
                'items_limbah': [],
                'bulan_romawi': now.strftime('%m')
            }
            conversations[sid] = state
            return jsonify({"text": "Baik, saya bantu buatkan quotation.<br><br>❓ <b>1. Nomor Depan Surat?</b> (contoh: 002)"})
        
        # Step 1: Nomor Depan
        if state['step'] == 'nomor_depan':
            state['data']['nomor_depan'] = text
            state['step'] = 'nama_perusahaan'
            conversations[sid] = state
            return jsonify({"text": f"✅ Nomor: <b>{text}</b><br><br>❓ <b>2. Nama Perusahaan?</b>"})
        
        # Step 2: Nama Perusahaan
        elif state['step'] == 'nama_perusahaan':
            state['data']['nama_perusahaan'] = text
            state['step'] = 'alamat_perusahaan'
            conversations[sid] = state
            return jsonify({"text": f"✅ Nama: <b>{text}</b><br><br>❓ <b>3. Alamat Perusahaan?</b>"})
        
        # Step 3: Alamat Perusahaan
        elif state['step'] == 'alamat_perusahaan':
            state['data']['alamat_perusahaan'] = text
            state['step'] = 'jenis_kode_limbah'
            state['data']['current_item'] = {}
            conversations[sid] = state
            return jsonify({"text": f"✅ Alamat: <b>{text}</b><br><br>📦 <b>Item #1</b><br>❓ <b>4. Sebutkan Jenis Limbah atau Kode Limbah</b><br><i>(Contoh: 'A102d' atau 'aki baterai bekas')</i>"})
        
        # Step 4: Jenis/Kode Limbah (NEW SMART DETECTION)
        elif state['step'] == 'jenis_kode_limbah':
            # Try to detect if it's a code
            kode, data_limbah = find_limbah_by_kode(text)
            
            if kode and data_limbah:
                # User provided a code
                state['data']['current_item']['kode_limbah'] = kode
                state['data']['current_item']['jenis_limbah'] = data_limbah['jenis']
                state['data']['current_item']['satuan'] = data_limbah['satuan']
                state['step'] = 'harga'
                conversations[sid] = state
                return jsonify({"text": f"✅ Kode: <b>{kode}</b><br>✅ Jenis: <b>{data_limbah['jenis']}</b><br>✅ Satuan: <b>{data_limbah['satuan']}</b><br><br>❓ <b>5. Harga (Rp)?</b>"})
            else:
                # Try to find by jenis
                kode, data_limbah = find_limbah_by_jenis(text)
                
                if kode and data_limbah:
                    # Found by jenis
                    state['data']['current_item']['kode_limbah'] = kode
                    state['data']['current_item']['jenis_limbah'] = data_limbah['jenis']
                    state['data']['current_item']['satuan'] = data_limbah['satuan']
                    state['step'] = 'harga'
                    conversations[sid] = state
                    return jsonify({"text": f"✅ Kode: <b>{kode}</b><br>✅ Jenis: <b>{data_limbah['jenis']}</b><br>✅ Satuan: <b>{data_limbah['satuan']}</b><br><br>❓ <b>5. Harga (Rp)?</b>"})
                else:
                    # Not found
                    return jsonify({"text": f"❌ Maaf, limbah '<b>{text}</b>' tidak ditemukan dalam database.<br><br>Silakan coba lagi dengan:<br>• Kode limbah (contoh: A102d, B105d)<br>• Nama jenis limbah (contoh: aki baterai bekas, minyak pelumas bekas)"})
        
        # Step 5: Harga
        elif state['step'] == 'harga':
            harga_converted = convert_voice_to_number(text)
            state['data']['current_item']['harga'] = harga_converted
            
            # Add item to list
            state['data']['items_limbah'].append(state['data']['current_item'])
            num = len(state['data']['items_limbah'])
            state['step'] = 'tambah_item'
            conversations[sid] = state
            
            harga_formatted = format_rupiah(harga_converted)
            return jsonify({"text": f"✅ Item #{num} tersimpan!<br>💰 Harga: <b>Rp {harga_formatted}</b><br><br>❓ <b>Tambah item lagi?</b> (ya/tidak)"})
        
        # Step 6: Tambah Item?
        elif state['step'] == 'tambah_item':
            if 'ya' in lower or 'iya' in lower:
                num = len(state['data']['items_limbah'])
                state['step'] = 'jenis_kode_limbah'
                state['data']['current_item'] = {}
                conversations[sid] = state
                return jsonify({"text": f"📦 <b>Item #{num+1}</b><br>❓ <b>4. Sebutkan Jenis Limbah atau Kode Limbah</b><br><i>(Contoh: 'A102d' atau 'aki baterai bekas')</i>"})
            else:
                state['step'] = 'harga_transportasi'
                conversations[sid] = state
                return jsonify({"text": f"✅ Total: <b>{len(state['data']['items_limbah'])} item</b><br><br>❓ <b>6. Biaya Transportasi (Rp)?</b><br><i>Satuan: ritase</i>"})
        
        # Step 7: Harga Transportasi
        elif state['step'] == 'harga_transportasi':
            transportasi_converted = convert_voice_to_number(text)
            state['data']['harga_transportasi'] = transportasi_converted
            state['step'] = 'tanya_mou'
            conversations[sid] = state
            transportasi_formatted = format_rupiah(transportasi_converted)
            return jsonify({"text": f"✅ Transportasi: <b>Rp {transportasi_formatted}/ritase</b><br><br>❓ <b>7. Tambah Biaya MoU?</b> (ya/tidak)"})
        
        # Step 8: Tanya MoU
        elif state['step'] == 'tanya_mou':
            if 'ya' in lower or 'iya' in lower:
                state['step'] = 'harga_mou'
                conversations[sid] = state
                return jsonify({"text": "❓ <b>Biaya MoU (Rp)?</b><br><i>Satuan: Tahun</i>"})
            else:
                # Generate document without MoU
                state['data']['harga_mou'] = None
                fname = f"Quotation_{re.sub(r'[^A-Za-z0-9]+', '_', state['data']['nama_perusahaan'])}_{uuid.uuid4().hex[:6]}"
                docx = create_docx(state['data'], fname)
                pdf = create_pdf(fname)
                
                conversations[sid] = {'step': 'idle', 'data': {}}
                
                files = [{"type": "docx", "filename": docx, "url": f"/static/files/{docx}"}]
                if pdf:
                    files.append({"type": "pdf", "filename": pdf, "url": f"/static/files/{pdf}"})
                
                return jsonify({"text": f"🎉 <b>Quotation berhasil dibuat!</b>", "files": files})
        
        # Step 9: Harga MoU
        elif state['step'] == 'harga_mou':
            mou_converted = convert_voice_to_number(text)
            state['data']['harga_mou'] = mou_converted
            
            # Generate document with MoU
            fname = f"Quotation_{re.sub(r'[^A-Za-z0-9]+', '_', state['data']['nama_perusahaan'])}_{uuid.uuid4().hex[:6]}"
            docx = create_docx(state['data'], fname)
            pdf = create_pdf(fname)
            
            conversations[sid] = {'step': 'idle', 'data': {}}
            
            files = [{"type": "docx", "filename": docx, "url": f"/static/files/{docx}"}]
            if pdf:
                files.append({"type": "pdf", "filename": pdf, "url": f"/static/files/{pdf}"})
            
            mou_formatted = format_rupiah(mou_converted)
            return jsonify({"text": f"✅ MoU: <b>Rp {mou_formatted}/Tahun</b><br><br>🎉 <b>Quotation berhasil dibuat!</b>", "files": files})
        
        # Fallback to AI
        return jsonify({"text": call_ai(text)})
        
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
    print(f"🔑 API Key: {OPENROUTER_API_KEY and '✅ Set' or '❌ Not Set'}")
    print(f"📄 PDF: {PDF_AVAILABLE and '✅ Enabled' or '❌ Disabled'}")
    print(f"🗄️  Database: {len(LIMBAH_B3_DB)} jenis limbah B3")
    print(f"🌐 Port: {port}")
    print(f"🔧 Debug: {debug_mode}")
    print("="*60 + "\n")
    
    app.run(host="0.0.0.0", port=port, debug=debug_mode)