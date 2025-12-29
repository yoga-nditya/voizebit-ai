import os
import json
import uuid
import re
from flask import Flask, request, jsonify, render_template, send_from_directory, session
from datetime import datetime
import platform

from docx import Document  # ✅ untuk generate MoU dari template DOCX

from config_new import *
from limbah_database import (
    LIMBAH_B3_DB,
    find_limbah_by_kode,
    find_limbah_by_jenis,
    convert_voice_to_number,
    parse_termin_days,
    angka_ke_terbilang,
    format_rupiah
)
from utils import (
    init_db, load_counter,
    db_insert_history, db_list_histories, db_get_history_detail,
    db_update_title, db_delete_history, db_append_message, db_update_state,
    get_next_nomor, create_docx, create_pdf,
    search_company_address, search_company_address_ai, call_ai,
    PDF_AVAILABLE, PDF_METHOD, LIBREOFFICE_PATH
)

app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = FLASK_SECRET_KEY

conversations = {}

init_db()

@app.after_request
def add_cors_headers(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, DELETE, OPTIONS"
    return resp


# ✅ helper NON B3 (variasi penulisan)
def is_non_b3_input(text: str) -> bool:
    if not text:
        return False
    t = text.strip().lower()
    norm = re.sub(r'[\s\-_]+', '', t)
    return norm in ("nonb3", "nonbii3") or norm.startswith("nonb3")


# ✅ normalisasi angka Indonesia:
# - 3.000 / 3,000 => 3000
# - 3,5 => 3.5
def normalize_id_number_text(text: str) -> str:
    if not text:
        return text
    t = text.strip()
    t = re.sub(r'(?<=\d)[\.,](?=\d{3}(\D|$))', '', t)   # hapus pemisah ribuan
    t = re.sub(r'(?<=\d),(?=\d)', '.', t)              # koma desimal jadi titik
    return t


# ✅ parse angka voice + dukung "koma" + satuan ribu/juta/miliar/triliun
# Fix kasus: "tiga koma lima ribu" => 3500 (bukan 8000)
def parse_amount_id(text: str) -> int:
    if not text:
        return 0

    raw = text.strip()
    lower = raw.lower()

    digit_map = {
        "nol": 0, "kosong": 0,
        "satu": 1, "se": 1,
        "dua": 2,
        "tiga": 3,
        "empat": 4,
        "lima": 5,
        "enam": 6,
        "tujuh": 7,
        "delapan": 8,
        "sembilan": 9
    }

    def token_to_digit(tok: str):
        tok = tok.strip().lower()
        if tok.isdigit():
            return int(tok)
        return digit_map.get(tok, None)

    scale_map = {
        "ribu": 1_000,
        "juta": 1_000_000,
        "miliar": 1_000_000_000,
        "triliun": 1_000_000_000_000,
    }
    scale = None
    for k, m in scale_map.items():
        if k in lower:
            scale = m
            break

    # “tiga koma lima ribu” => 3.5 * 1000 = 3500
    if "koma" in lower:
        parts = re.split(r'\bkoma\b', lower, maxsplit=1)
        left_part = parts[0].strip()
        right_part = parts[1].strip() if len(parts) > 1 else ""

        left_tokens = re.findall(r'[a-zA-Z0-9]+', left_part)
        left_tok = left_tokens[-1] if left_tokens else ""
        left_digit = token_to_digit(left_tok)

        right_tokens = re.findall(r'[a-zA-Z0-9]+', right_part)
        right_tok = right_tokens[0] if right_tokens else ""
        right_digit = token_to_digit(right_tok)

        if left_digit is not None and right_digit is not None:
            val = float(f"{left_digit}.{right_digit}")
            if scale:
                val *= scale
            return int(round(val))

    tnorm = normalize_id_number_text(raw)
    val = convert_voice_to_number(tnorm)
    if val is None:
        val = 0

    try:
        f = float(val)
        if scale and f < scale:
            val = f * scale
    except:
        pass

    try:
        return int(round(float(val)))
    except:
        digits = re.sub(r'\D+', '', str(val))
        return int(digits) if digits else 0


# ✅ buat nama file unik
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


# ===========================
# ✅ COUNTER KHUSUS MOU (mulai dari 000)
# ===========================
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
    except:
        return -1


def save_mou_counter(n: int) -> None:
    path = _mou_counter_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"counter": int(n)}, f)


def get_next_mou_no_depan() -> str:
    n = load_mou_counter() + 1
    save_mou_counter(n)
    return str(n).zfill(3)  # 000, 001, 002, ...


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
    bulan_map = {
        1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
        5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
        9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }
    hari = hari_map.get(dt.weekday(), "")
    bulan = bulan_map.get(dt.month, "")
    return f"{hari}, tanggal {dt.day} {bulan} {dt.year}"


# ===========================
# ✅ DOCX HELPERS (anti rusak format)
# ===========================
def _iter_all_paragraphs(doc: Document):
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def replace_text_in_runs(doc: Document, old: str, new: str):
    """Replace text hanya di dalam RUN yang mengandung 'old' => format (bold/align) aman."""
    if not old:
        return
    for p in _iter_all_paragraphs(doc):
        for r in p.runs:
            if old in r.text:
                r.text = r.text.replace(old, new)


def replace_regex_in_runs(doc: Document, pattern: str, repl: str, flags=re.IGNORECASE):
    """Regex replacement per RUN (tidak clear semua runs) => format aman."""
    for p in _iter_all_paragraphs(doc):
        for r in p.runs:
            if re.search(pattern, r.text, flags=flags):
                r.text = re.sub(pattern, repl, r.text, flags=flags)


def clean_address(addr: str) -> str:
    addr = (addr or "").strip()
    addr = re.sub(r'\s+', ' ', addr).strip()
    # hindari address ending aneh seperti ");" atau ","
    addr = re.sub(r'[;,]+$', '', addr).strip()
    addr = re.sub(r'\)\s*;$', ')', addr).strip()
    return addr


# ===========================
# ✅ GENERATE MOU DOCX (FORMAT SAMA TEMPLATE)
# ===========================
def create_mou_docx(mou_data: dict, fname_base: str) -> str:
    # ✅ TEMPLATE DI ROOT (bukan folder templates)
    base_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(base_dir, "tamplate MoU.docx")
    if not os.path.exists(template_path):
        raise Exception("Template MoU tidak ditemukan. Pastikan ada di root: tamplate MoU.docx")

    doc = Document(template_path)

    pihak1 = (mou_data.get("pihak_pertama") or "").strip()
    pihak2 = (mou_data.get("pihak_kedua") or "").strip()
    pihak3 = (mou_data.get("pihak_ketiga") or "").strip()

    alamat1 = clean_address(mou_data.get("alamat_pihak_pertama") or "")
    alamat3 = clean_address(mou_data.get("alamat_pihak_ketiga") or "")

    ttd1 = (mou_data.get("ttd_pihak_pertama") or "").strip()
    ttd3 = (mou_data.get("ttd_pihak_ketiga") or "").strip()

    nomor_full = (mou_data.get("nomor_surat") or mou_data.get("nomor_depan") or "").strip()

    # ===========================
    # 1) Replace NAMA PIHAK (global) - aman
    #    Header butuh uppercase seperti template
    # ===========================
    # Pihak 1
    replace_text_in_runs(doc, "PT. PANPAN LUCKY INDONESIA", pihak1.upper() if pihak1 else "PT. PANPAN LUCKY INDONESIA")
    replace_text_in_runs(doc, "PT PANPAN LUCKY INDONESIA", pihak1.upper() if pihak1 else "PT PANPAN LUCKY INDONESIA")
    replace_text_in_runs(doc, "PT. Panpan Lucky Indonesia", pihak1 if pihak1 else "PT. Panpan Lucky Indonesia")
    replace_text_in_runs(doc, "PT Panpan Lucky Indonesia", pihak1 if pihak1 else "PT Panpan Lucky Indonesia")

    # Pihak 2 (tetap template)
    # Pihak 3
    replace_text_in_runs(doc, "PT. HARAPAN BARU SEJAHTERA PLASTIK", pihak3.upper() if pihak3 else "PT. HARAPAN BARU SEJAHTERA PLASTIK")
    replace_text_in_runs(doc, "PT HARAPAN BARU SEJAHTERA PLASTIK", pihak3.upper() if pihak3 else "PT HARAPAN BARU SEJAHTERA PLASTIK")
    replace_text_in_runs(doc, "PT. Harapan Baru Sejahtera Plastik", pihak3 if pihak3 else "PT. Harapan Baru Sejahtera Plastik")
    replace_text_in_runs(doc, "PT Harapan Baru Sejahtera Plastik", pihak3 if pihak3 else "PT Harapan Baru Sejahtera Plastik")

    # ===========================
    # 2) Replace NOMOR "No : ...." (tanpa rusak format)
    # ===========================
    if nomor_full:
        replace_regex_in_runs(doc, r'No\s*:\s*.*', f"No : {nomor_full}")

    # ===========================
    # 3) Replace TANGGAL (tanpa dobel dan format tetap)
    #    Template punya 2 run: (run0= tanggal) + (run1= “kami yang bertanda...”)
    # ===========================
    tanggal_text = format_tanggal_indonesia(datetime.now())
    for p in doc.paragraphs:
        if "Pada hari ini" in p.text and "bertanda tangan" in p.text and "tanggal" in p.text:
            if p.runs:
                p.runs[0].text = f"Pada hari ini {tanggal_text} "
                if len(p.runs) >= 2:
                    # biarkan run1 tetap template (jangan diganti kalau sudah benar)
                    if "kami yang bertanda tangan" not in p.runs[1].text:
                        p.runs[1].text = "kami yang bertanda tangan di bawah ini : "
                else:
                    p.add_run("kami yang bertanda tangan di bawah ini : ")
            break

    # ===========================
    # 4) Replace ALAMAT PIHAK 1 di bagian deskripsi (tanpa rusak bold)
    #    Template address pihak1 ada RUN khusus berisi “Jl. Raya Serang ...”
    # ===========================
    if alamat1:
        for p in doc.paragraphs:
            if "PIHAK PERTAMA" in p.text and "diwakili" in p.text:
                for r in p.runs:
                    if "Jl." in r.text and ("Serang" in r.text or "Tigaraksa" in r.text):
                        r.text = alamat1
                    # ganti nama perwakilan di paragraf pihak 1
                    if "Huang Feifang" in r.text and ttd1:
                        r.text = r.text.replace("Huang Feifang", ttd1)
                break

    # ===========================
    # 5) Replace ALAMAT PIHAK 3 di bagian deskripsi
    #    Di template, alamat pihak 3 ada dalam 1 RUN besar -> pakai regex segment
    # ===========================
    if alamat3:
        for p in doc.paragraphs:
            if "PIHAK KETIGA" in p.text and "Pengolah" in p.text:
                for r in p.runs:
                    if "berkedudukan dan berkantor di" in r.text:
                        r.text = re.sub(
                            r'(berkedudukan dan berkantor di\s+)(.*?)(,\s*yang)',
                            r'\1' + alamat3 + r'\3',
                            r.text,
                            flags=re.IGNORECASE
                        )
                break

    # ===========================
    # 6) Replace TTD BAWAH (table 3 kolom)
    # ===========================
    if len(doc.tables) >= 2:
        sig_table = doc.tables[1]
        # cell 0: pihak 1 - nama penandatangan template: "Huang Feifang"
        if ttd1:
            for p in sig_table.rows[0].cells[0].paragraphs:
                for r in p.runs:
                    if "Huang Feifang" in r.text:
                        r.text = r.text.replace("Huang Feifang", ttd1)

        # cell 2: pihak 3 - nama penandatangan template: "Yogi Permana"
        if ttd3:
            for p in sig_table.rows[0].cells[2].paragraphs:
                for r in p.runs:
                    if "Yogi Permana" in r.text:
                        r.text = r.text.replace("Yogi Permana", ttd3)

    # ===========================
    # 7) Isi TABLE LIMBAH (table pertama)
    # ===========================
    items = mou_data.get("items_limbah") or []
    target_table = None
    for t in doc.tables:
        header_text = " ".join([c.text.strip() for c in t.rows[0].cells]) if t.rows else ""
        if ("Jenis Limbah" in header_text) and ("Kode Limbah" in header_text):
            target_table = t
            break

    if target_table is not None:
        while len(target_table.rows) > 1:
            target_table._tbl.remove(target_table.rows[1]._tr)

        for i, it in enumerate(items, start=1):
            row = target_table.add_row()
            cells = row.cells
            if len(cells) >= 1:
                cells[0].text = str(i)
            if len(cells) >= 2:
                cells[1].text = (it.get("jenis_limbah") or "").strip()
            if len(cells) >= 3:
                cells[2].text = (it.get("kode_limbah") or "").strip()

    try:
        folder = str(FILES_DIR)
    except Exception:
        folder = "static/files"
    os.makedirs(folder, exist_ok=True)

    out_path = os.path.join(folder, f"{fname_base}.docx")
    doc.save(out_path)

    return f"{fname_base}.docx"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/history", methods=["GET"])
def api_history_list():
    try:
        q = (request.args.get("q") or "").strip()
        items = db_list_histories(limit=200, q=q if q else None)
        return jsonify({"items": items})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


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

        docs.sort(key=lambda x: x.get("created_at") or "", reverse=True)
        return jsonify({"items": docs})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/chat", methods=["POST"])
def chat():
    try:
        data = request.get_json() or {}
        text = (data.get("message", "") or "").strip()
        history_id_in = data.get("history_id")

        if not text:
            return jsonify({"error": "Pesan kosong"}), 400

        sid = request.headers.get("X-Session-ID") or session.get("sid")
        if not sid:
            sid = str(uuid.uuid4())
            session["sid"] = sid

        state = conversations.get(sid, {'step': 'idle', 'data': {}})
        lower = text.lower()

        if history_id_in:
            try:
                db_append_message(int(history_id_in), "user", text, files=[])
                db_update_state(int(history_id_in), state)
            except:
                pass

        # ============================================================
        # ✅ FITUR MOU TRIPARTIT (BARU)
        # Trigger: user ketik "mou"
        # ============================================================
        if ('mou' in lower) and (state.get('step') == 'idle'):
            nomor_depan = get_next_mou_no_depan()  # ✅ mulai dari 000
            state['step'] = 'mou_pihak_pertama'
            state['data'] = {
                'nomor_depan': nomor_depan,
                'nomor_surat': "",
                'items_limbah': [],
                'current_item': {},
                'pihak_kedua': "PT Sarana Trans Bersama Jaya",
                'pihak_kedua_kode': "STBJ",
                'pihak_pertama': "",
                'alamat_pihak_pertama': "",
                'pihak_ketiga': "",
                'alamat_pihak_ketiga': "",  # ✅ akan auto/set default
                'pihak_ketiga_kode': "",
                'ttd_pihak_pertama': "",     # ✅ ditanya terakhir
                'ttd_pihak_ketiga': "",      # ✅ ditanya terakhir
            }
            conversations[sid] = state

            out_text = (
                "Baik, saya bantu buatkan <b>MoU Tripartit</b>.<br><br>"
                f"✅ No Depan: <b>{nomor_depan}</b> (auto mulai 000)<br>"
                "✅ Tanggal: <b>otomatis hari ini</b><br><br>"
                "❓ <b>1. Nama Perusahaan (PIHAK PERTAMA / Penghasil Limbah)?</b>"
            )

            history_id_created = None
            if not history_id_in:
                history_id_created = db_insert_history(
                    title="Chat Baru",
                    task_type=data.get("taskType") or "mou",
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

        # Step MoU: pihak pertama
        if state.get('step') == 'mou_pihak_pertama':
            state['data']['pihak_pertama'] = text.strip()

            alamat = search_company_address(text).strip()
            if not alamat:
                alamat = search_company_address_ai(text).strip()
            if not alamat:
                alamat = "Di Tempat"
            state['data']['alamat_pihak_pertama'] = alamat

            state['step'] = 'mou_pilih_pihak_ketiga'
            conversations[sid] = state

            out_text = (
                f"✅ PIHAK PERTAMA: <b>{state['data']['pihak_pertama']}</b><br>"
                f"✅ Alamat: <b>{alamat}</b><br><br>"
                "❓ <b>2. Pilih PIHAK KETIGA (Pengelola Limbah):</b><br>"
                "1. HBSP<br>"
                "2. KJL<br>"
                "3. MBI<br>"
                "4. CGA<br><br>"
                "<i>(Ketik nomor 1-4 atau ketik langsung HBSP/KJL/MBI/CGA)</i>"
            )

            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)

            return jsonify({"text": out_text, "history_id": history_id_in})

        # Step MoU: pilih pihak ketiga
        if state.get('step') == 'mou_pilih_pihak_ketiga':
            pilihan = text.strip().upper()

            mapping = {
                "1": "HBSP",
                "2": "KJL",
                "3": "MBI",
                "4": "CGA",
                "HBSP": "HBSP",
                "KJL": "KJL",
                "MBI": "MBI",
                "CGA": "CGA",
            }
            kode = mapping.get(pilihan)
            if not kode:
                out_text = (
                    "⚠️ Pilihan tidak valid.<br><br>"
                    "Pilih PIHAK KETIGA:<br>"
                    "1. HBSP<br>2. KJL<br>3. MBI<br>4. CGA"
                )
                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                return jsonify({"text": out_text, "history_id": history_id_in})

            pihak3_map = {
                "HBSP": "PT Harapan Baru Sejahtera Plastik",
                "KJL": "PT KJL",
                "MBI": "PT MBI",
                "CGA": "PT CGA",
            }
            # alamat default (HBSP sesuai template; lainnya kosong supaya tidak salah)
            alamat3_map = {
                "HBSP": "Jl. Karawang – Bekasi KM. 1 Bojongsari, Kec. Kedungwaringin, Kab. Bekasi – Jawa Barat.",
                "KJL": "",
                "MBI": "",
                "CGA": "",
            }

            state['data']['pihak_ketiga'] = pihak3_map.get(kode, f"PT {kode}")
            state['data']['pihak_ketiga_kode'] = kode
            state['data']['alamat_pihak_ketiga'] = alamat3_map.get(kode, "")

            # nomor surat full
            state['data']['nomor_surat'] = build_mou_nomor_surat(state['data'])

            state['step'] = 'mou_jenis_kode_limbah'
            state['data']['current_item'] = {}
            conversations[sid] = state

            out_text = (
                f"✅ PIHAK KETIGA: <b>{state['data']['pihak_ketiga']}</b><br>"
                f"✅ Nomor MoU: <b>{state['data']['nomor_surat']}</b><br><br>"
                f"📦 <b>Item #1</b><br>"
                "❓ <b>3. Sebutkan Jenis Limbah atau Kode Limbah</b><br>"
                "<i>(Contoh: 'A102d' atau 'aki baterai bekas' | atau ketik <b>NON B3</b> untuk manual)</i>"
            )

            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)

            return jsonify({"text": out_text, "history_id": history_id_in})

        # Step MoU: input limbah (mirip quotation, tapi tanpa harga)
        if state.get('step') == 'mou_jenis_kode_limbah':
            if is_non_b3_input(text):
                state['data']['current_item']['kode_limbah'] = "NON B3"
                state['data']['current_item']['jenis_limbah'] = ""
                state['step'] = 'mou_manual_jenis_limbah'
                conversations[sid] = state

                out_text = (
                    "✅ Kode: <b>NON B3</b><br><br>"
                    "❓ <b>3A. Jenis Limbah (manual) apa?</b><br>"
                    "<i>(Contoh: 'plastik bekas', 'kertas bekas', dll)</i>"
                )
                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                    db_update_state(int(history_id_in), state)

                return jsonify({"text": out_text, "history_id": history_id_in})

            kode, data_limbah = find_limbah_by_kode(text)
            if not (kode and data_limbah):
                kode, data_limbah = find_limbah_by_jenis(text)

            if kode and data_limbah:
                state['data']['current_item']['kode_limbah'] = kode
                state['data']['current_item']['jenis_limbah'] = data_limbah['jenis']

                state['data']['items_limbah'].append(state['data']['current_item'])
                num = len(state['data']['items_limbah'])

                state['step'] = 'mou_tambah_item'
                state['data']['current_item'] = {}
                conversations[sid] = state

                out_text = (
                    f"✅ Item #{num} tersimpan!<br>"
                    f"• Jenis: <b>{data_limbah['jenis']}</b><br>"
                    f"• Kode: <b>{kode}</b><br><br>"
                    "❓ <b>Tambah item lagi?</b> (ya/tidak)"
                )

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                    db_update_state(int(history_id_in), state)

                return jsonify({"text": out_text, "history_id": history_id_in})

            out_text = (
                f"❌ Maaf, limbah '<b>{text}</b>' tidak ditemukan dalam database.<br><br>"
                "Silakan coba lagi dengan:<br>"
                "• Kode limbah (contoh: A102d, B105d)<br>"
                "• Nama jenis limbah (contoh: aki baterai bekas, minyak pelumas bekas)<br>"
                "• Atau ketik <b>NON B3</b> untuk input manual"
            )

            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)

            return jsonify({"text": out_text, "history_id": history_id_in})

        # Step MoU: manual jenis
        if state.get('step') == 'mou_manual_jenis_limbah':
            state['data']['current_item']['jenis_limbah'] = text.strip()

            state['data']['items_limbah'].append(state['data']['current_item'])
            num = len(state['data']['items_limbah'])

            state['step'] = 'mou_tambah_item'
            state['data']['current_item'] = {}
            conversations[sid] = state

            out_text = (
                f"✅ Item #{num} tersimpan!<br>"
                f"• Jenis (manual): <b>{state['data']['items_limbah'][-1]['jenis_limbah']}</b><br>"
                f"• Kode: <b>NON B3</b><br><br>"
                "❓ <b>Tambah item lagi?</b> (ya/tidak)"
            )

            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)

            return jsonify({"text": out_text, "history_id": history_id_in})

        # Step MoU: tambah item atau lanjut ke TTD (TTD HARUS TERAKHIR)
        if state.get('step') == 'mou_tambah_item':
            if re.match(r'^\d+', text.strip()):
                out_text = "⚠️ Mohon jawab dengan <b>'ya'</b> atau <b>'tidak'</b><br><br>❓ <b>Tambah item lagi?</b>"
                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                return jsonify({"text": out_text, "history_id": history_id_in})

            if ('ya' in lower) or ('iya' in lower):
                num = len(state['data']['items_limbah'])
                state['step'] = 'mou_jenis_kode_limbah'
                state['data']['current_item'] = {}
                conversations[sid] = state

                out_text = (
                    f"📦 <b>Item #{num+1}</b><br>"
                    "❓ <b>3. Sebutkan Jenis Limbah atau Kode Limbah</b><br>"
                    "<i>(Contoh: 'A102d' atau 'aki baterai bekas' | atau ketik <b>NON B3</b> untuk manual)</i>"
                )

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                    db_update_state(int(history_id_in), state)

                return jsonify({"text": out_text, "history_id": history_id_in})

            if ('tidak' in lower) or ('skip' in lower) or ('lewat' in lower) or ('gak' in lower) or ('nggak' in lower):
                # ✅ TANYA TTD TERAKHIR (STEP 1)
                state['step'] = 'mou_ttd_pihak_pertama'
                conversations[sid] = state

                out_text = (
                    "✅ Baik, item limbah sudah selesai.<br><br>"
                    "❓ <b>Terakhir, siapa nama penandatangan PIHAK PERTAMA?</b><br>"
                    "<i>(Nama yang akan muncul di bagian tanda tangan bawah)</i>"
                )

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                    db_update_state(int(history_id_in), state)

                return jsonify({"text": out_text, "history_id": history_id_in})

            out_text = "⚠️ Mohon jawab dengan <b>'ya'</b> atau <b>'tidak'</b><br><br>❓ <b>Tambah item lagi?</b>"
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            return jsonify({"text": out_text, "history_id": history_id_in})

        # ✅ TTD PIHAK PERTAMA (TERAKHIR - Q1)
        if state.get('step') == 'mou_ttd_pihak_pertama':
            state['data']['ttd_pihak_pertama'] = text.strip()
            state['step'] = 'mou_ttd_pihak_ketiga'
            conversations[sid] = state

            out_text = (
                "❓ <b>Terakhir, siapa nama penandatangan PIHAK KETIGA?</b><br>"
                "<i>(Nama yang akan muncul di bagian tanda tangan bawah)</i>"
            )

            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)

            return jsonify({"text": out_text, "history_id": history_id_in})

        # ✅ TTD PIHAK KETIGA (TERAKHIR - Q2) lalu GENERATE
        if state.get('step') == 'mou_ttd_pihak_ketiga':
            state['data']['ttd_pihak_ketiga'] = text.strip()

            # Generate MoU
            nama_pt_raw = state['data'].get('pihak_pertama', '').strip()
            safe_pt = re.sub(r'[^A-Za-z0-9 \-]+', '', nama_pt_raw).strip()
            safe_pt = re.sub(r'\s+', ' ', safe_pt).strip()

            base_fname = f"MoU - {safe_pt}" if safe_pt else "MoU - Perusahaan"
            fname_base = make_unique_filename_base(base_fname)

            if not state['data'].get("nomor_surat"):
                state['data']['nomor_surat'] = build_mou_nomor_surat(state['data'])

            docx = create_mou_docx(state['data'], fname_base)
            pdf = create_pdf(fname_base)

            conversations[sid] = {'step': 'idle', 'data': {}}

            files = [{"type": "docx", "filename": docx, "url": f"/static/files/{docx}"}]
            if pdf:
                files.append({"type": "pdf", "filename": pdf, "url": f"/static/files/{pdf}"})

            history_title = f"MoU {nama_pt_raw}" if nama_pt_raw else "MoU"
            history_task_type = "mou"

            if history_id_in:
                from utils import db_connect
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

            out_text = (
                "🎉 <b>MoU berhasil dibuat!</b><br><br>"
                f"✅ Nomor MoU: <b>{state['data'].get('nomor_surat')}</b><br>"
                f"✅ PIHAK PERTAMA: <b>{state['data'].get('pihak_pertama')}</b><br>"
                f"✅ PIHAK KEDUA: <b>{state['data'].get('pihak_kedua')}</b><br>"
                f"✅ PIHAK KETIGA: <b>{state['data'].get('pihak_ketiga')}</b><br>"
                f"✅ Total Limbah: <b>{len(state['data'].get('items_limbah') or [])} item</b><br>"
                f"✅ TTD PIHAK PERTAMA: <b>{state['data'].get('ttd_pihak_pertama')}</b><br>"
                f"✅ TTD PIHAK KETIGA: <b>{state['data'].get('ttd_pihak_ketiga')}</b>"
            )

            db_append_message(history_id, "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=files)

            return jsonify({
                "text": out_text,
                "files": files,
                "history_id": history_id
            })

        # ============================================================
        # ✅ FITUR QUOTATION (EXISTING) - MINOR FIX:
        # Biar "buat MoU" tidak salah masuk ke quotation
        # ============================================================
        if ('quotation' in lower or 'penawaran' in lower or ('buat' in lower and 'mou' not in lower)):
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

        # ======= FLOW QUOTATION EXISTING (kode Anda tetap) =======
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

        elif state['step'] == 'jenis_kode_limbah':
            if is_non_b3_input(text):
                state['data']['current_item']['kode_limbah'] = "NON B3"
                state['data']['current_item']['jenis_limbah'] = ""
                state['data']['current_item']['satuan'] = ""
                state['step'] = 'manual_jenis_limbah'
                conversations[sid] = state

                out_text = (
                    "✅ Kode: <b>NON B3</b><br><br>"
                    "❓ <b>2A. Jenis Limbah (manual) apa?</b><br>"
                    "<i>(Contoh: 'plastik bekas', 'kertas bekas', dll)</i>"
                )

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                    db_update_state(int(history_id_in), state)

                return jsonify({"text": out_text, "history_id": history_id_in})

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
                        "• Nama jenis limbah (contoh: aki baterai bekas, minyak pelumas bekas)<br>"
                        "• Atau ketik <b>NON B3</b> untuk input manual"
                    )

                    if history_id_in:
                        db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                        db_update_state(int(history_id_in), state)

                    return jsonify({"text": out_text, "history_id": history_id_in})

        elif state['step'] == 'manual_jenis_limbah':
            state['data']['current_item']['jenis_limbah'] = text
            state['step'] = 'manual_satuan'
            conversations[sid] = state

            out_text = (
                f"✅ Jenis (manual): <b>{text}</b><br><br>"
                "❓ <b>2B. Satuan (manual) apa?</b><br>"
                "<i>(Contoh: kg, liter, drum, pcs, dll)</i>"
            )

            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)

            return jsonify({"text": out_text, "history_id": history_id_in})

        elif state['step'] == 'manual_satuan':
            state['data']['current_item']['satuan'] = text
            state['step'] = 'harga'
            conversations[sid] = state

            out_text = (
                f"✅ Satuan (manual): <b>{text}</b><br><br>"
                "❓ <b>3. Harga (Rp)?</b>"
            )

            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)

            return jsonify({"text": out_text, "history_id": history_id_in})

        elif state['step'] == 'harga':
            harga_converted = parse_amount_id(text)
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

        # ... sisanya quotation kamu biarkan sama (lanjutkan seperti file kamu)

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
    port = FLASK_PORT
    debug_mode = FLASK_DEBUG

    print("\n" + "="*60)
    print("🚀 QUOTATION GENERATOR")
    print("="*60)
    print(f"📁 Template: {TEMPLATE_FILE.exists() and '✅ Found' or '❌ Missing'}")
    print(f"🔑 OpenRouter: {OPENROUTER_API_KEY and '✅' or '❌'}")
    print(f"🔎 Serper: {SERPER_API_KEY and '✅' or '❌'}")
    print(f"📄 PDF: {PDF_AVAILABLE and f'✅ {PDF_METHOD}' or '❌ Disabled'}")
    print(f"🗄️  Database: {len(LIMBAH_B3_DB)} jenis limbah")
    print(f"🔢 Counter: {load_counter()}")
    print(f"🔢 MoU Counter: {load_mou_counter()}")
    print(f"🌐 Port: {port}")
    print(f"💻 Platform: {platform.system()}")
    print("="*60 + "\n")

    app.run(host="0.0.0.0", port=port, debug=debug_mode)
