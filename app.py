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


def convert_voice_to_unit(text):
    text_lower = text.lower().strip()

    satuan_mapping = {
        'kilogram': 'Kg',
        'kilo': 'Kg',
        'kg': 'Kg',
        'gram': 'Gram',
        'ton': 'Ton',
        'liter': 'L',
        'meter kubik': 'm³',
        'meter persegi': 'm²',
        'meter': 'm',
        'borong': 'Borong',
        'paket': 'Paket',
        'unit': 'Unit',
        'buah': 'Buah',
        'pcs': 'Pcs',
        'pieces': 'Pcs',
        'karung': 'Karung',
        'dus': 'Dus',
        'box': 'Box',
        'ritase': 'ritase',
        'trip': 'Trip',
        'tahun': 'Tahun',
    }

    if text_lower in satuan_mapping:
        return satuan_mapping[text_lower]

    return text.capitalize()


def convert_voice_to_waste_code(text):
    text_upper = text.upper().strip()

    if re.match(r'^[A-Z]\d+-\d+$', text_upper):
        return text_upper

    kata_ke_angka = {
        'NOL': '0', 'KOSONG': '0',
        'SATU': '1', 'SE': '1',
        'DUA': '2', 'TIGA': '3',
        'EMPAT': '4', 'LIMA': '5',
        'ENAM': '6', 'TUJUH': '7',
        'DELAPAN': '8', 'SEMBILAN': '9'
    }

    text_processed = text_upper.replace('STRIP', '|||').replace('MINUS', '|||')
    words = text_processed.split()

    result = []
    for word in words:
        if word == '|||':
            result.append('-')
        elif word in kata_ke_angka:
            result.append(kata_ke_angka[word])
        elif len(word) == 1 and word.isalpha():
            result.append(word)
        elif word.isdigit():
            result.append(word)
        elif re.match(r'^[A-Z]\d+$', word):
            result.append(word[0])
            result.append(word[1:])

    code = ''.join(result)

    if re.match(r'^[A-Z]\d+-\d+$', code):
        return code

    match = re.match(r'^([A-Z])(\d+)-(\d+)$', code)
    if match:
        return code

    match = re.match(r'^([A-Z])(\d+)(\d)$', code)
    if match:
        return f"{match.group(1)}{match.group(2)}-{match.group(3)}"

    return text


def call_ai(text, system_prompt=None):
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json"
    }

    messages = []
    if system_prompt:
        messages.append({"role": "system", "content": system_prompt})
    messages.append({"role": "user", "content": text})

    resp = requests.post(
        url,
        headers=headers,
        json={
            "model": OPENROUTER_MODEL,
            "messages": messages,
            "temperature": 0.3,
            "max_tokens": 2000
        },
        timeout=60
    )
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
    except Exception:
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
                    satuan = escape_xml(item.get('satuan', ''))

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

        if 'quotation' in lower or 'penawaran' in lower or 'buat' in lower:
            state['step'] = 'nomor_depan'
            now = datetime.now()
            state['data'] = {
                'items_limbah': [],
                'bulan_romawi': now.strftime('%m')
            }
            conversations[sid] = state
            return jsonify({"text": "Baik, saya bantu buatkan quotation.<br><br>❓ <b>1. Nomor Depan Surat?</b> (contoh: 002)"})

        if state['step'] == 'nomor_depan':
            state['data']['nomor_depan'] = text
            state['step'] = 'nama_perusahaan'
            conversations[sid] = state
            return jsonify({"text": f"✅ Nomor: <b>{text}</b><br><br>❓ <b>2. Nama Perusahaan?</b>"})

        elif state['step'] == 'nama_perusahaan':
            state['data']['nama_perusahaan'] = text
            state['step'] = 'alamat_perusahaan'
            conversations[sid] = state
            return jsonify({"text": f"✅ Nama: <b>{text}</b><br><br>❓ <b>3. Alamat Perusahaan?</b>"})

        elif state['step'] == 'alamat_perusahaan':
            state['data']['alamat_perusahaan'] = text
            state['step'] = 'jenis_limbah'
            state['data']['current_item'] = {}
            conversations[sid] = state
            return jsonify({"text": f"✅ Alamat: <b>{text}</b><br><br>📦 <b>Item #1</b><br>❓ <b>4. Jenis Limbah?</b>"})

        elif state['step'] == 'jenis_limbah':
            state['data']['current_item']['jenis_limbah'] = text
            state['step'] = 'kode_limbah'
            conversations[sid] = state
            return jsonify({"text": f"✅ Jenis: <b>{text}</b><br><br>❓ <b>5. Kode Limbah?</b> (contoh: A331-1)"})

        elif state['step'] == 'kode_limbah':
            kode_converted = convert_voice_to_waste_code(text)
            state['data']['current_item']['kode_limbah'] = kode_converted
            state['step'] = 'harga'
            conversations[sid] = state
            return jsonify({"text": f"✅ Kode: <b>{kode_converted}</b><br><br>❓ <b>6. Harga (Rp)?</b>"})

        elif state['step'] == 'harga':
            harga_converted = convert_voice_to_number(text)
            state['data']['current_item']['harga'] = harga_converted
            state['step'] = 'satuan'
            conversations[sid] = state
            harga_formatted = format_rupiah(harga_converted)
            return jsonify({"text": f"✅ Harga: <b>Rp {harga_formatted}</b><br><br>❓ <b>7. Satuan?</b>"})

        elif state['step'] == 'satuan':
            satuan_converted = convert_voice_to_unit(text)
            state['data']['current_item']['satuan'] = satuan_converted
            state['data']['items_limbah'].append(state['data']['current_item'])
            num = len(state['data']['items_limbah'])
            state['step'] = 'tambah_item'
            conversations[sid] = state
            return jsonify({"text": f"✅ Item #{num} tersimpan!<br><br>❓ <b>Tambah item lagi?</b> (ya/tidak)"})

        elif state['step'] == 'tambah_item':
            if 'ya' in lower or 'iya' in lower:
                num = len(state['data']['items_limbah'])
                state['step'] = 'jenis_limbah'
                state['data']['current_item'] = {}
                conversations[sid] = state
                return jsonify({"text": f"📦 <b>Item #{num+1}</b><br>❓ <b>4. Jenis Limbah?</b>"})
            else:
                state['step'] = 'harga_transportasi'
                conversations[sid] = state
                return jsonify({"text": f"✅ Total: <b>{len(state['data']['items_limbah'])} item</b><br><br>❓ <b>8. Biaya Transportasi (Rp)?</b>"})

        elif state['step'] == 'harga_transportasi':
            transportasi_converted = convert_voice_to_number(text)
            state['data']['harga_transportasi'] = transportasi_converted
            state['step'] = 'tanya_mou'
            conversations[sid] = state
            transportasi_formatted = format_rupiah(transportasi_converted)
            return jsonify({"text": f"✅ Transportasi: <b>Rp {transportasi_formatted}/ritase</b><br><br>❓ <b>9. Tambah Biaya MoU?</b> (ya/tidak)"})

        elif state['step'] == 'tanya_mou':
            if 'ya' in lower or 'iya' in lower:
                state['step'] = 'harga_mou'
                conversations[sid] = state
                return jsonify({"text": "❓ <b>Biaya MoU (Rp)?</b>"})
            else:
                state['data']['harga_mou'] = None
                fname = f"Quotation_{re.sub(r'[^A-Za-z0-9]+', '_', state['data']['nama_perusahaan'])}_{uuid.uuid4().hex[:6]}"
                docx = create_docx(state['data'], fname)
                pdf = create_pdf(fname)

                conversations[sid] = {'step': 'idle', 'data': {}}

                files = [{"type": "docx", "filename": docx, "url": f"/static/files/{docx}"}]
                if pdf:
                    files.append({"type": "pdf", "filename": pdf, "url": f"/static/files/{pdf}"})

                return jsonify({"text": "<b>Quotation berhasil dibuat!</b>", "files": files})

        elif state['step'] == 'harga_mou':
            mou_converted = convert_voice_to_number(text)
            state['data']['harga_mou'] = mou_converted

            fname = f"Quotation_{re.sub(r'[^A-Za-z0-9]+', '_', state['data']['nama_perusahaan'])}_{uuid.uuid4().hex[:6]}"
            docx = create_docx(state['data'], fname)
            pdf = create_pdf(fname)

            conversations[sid] = {'step': 'idle', 'data': {}}

            files = [{"type": "docx", "filename": docx, "url": f"/static/files/{docx}"}]
            if pdf:
                files.append({"type": "pdf", "filename": pdf, "url": f"/static/files/{pdf}"})

            return jsonify({"text": "<b>Quotation berhasil dibuat!</b>", "files": files})

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
    app.run(host="0.0.0.0", port=5000, debug=True)
