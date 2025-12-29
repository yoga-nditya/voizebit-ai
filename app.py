import os
import json
import uuid
import re
import platform
from datetime import datetime

from flask import Flask, request, jsonify, render_template, send_from_directory, session

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

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
    PDF_AVAILABLE, PDF_METHOD, LIBREOFFICE_PATH,
    FILES_DIR
)

app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = FLASK_SECRET_KEY

# in-memory session state (Render: ini akan reset kalau container restart)
conversations = {}

init_db()


# =========================
# CORS (Simple)
# =========================
@app.after_request
def add_cors_headers(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization, X-Session-ID"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, DELETE, OPTIONS"
    return resp


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
    # hapus pemisah ribuan 3.000 / 3,000
    t = re.sub(r'(?<=\d)[\.,](?=\d{3}(\D|$))', '', t)
    # koma desimal -> titik
    t = re.sub(r'(?<=\d),(?=\d)', '.', t)
    return t


def parse_amount_id(text: str) -> int:
    """
    Parse angka Indonesia untuk voice/text.
    Support: "3.000", "3,000", "3,5 ribu" => 3500
    """
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
        if re.search(rf'\b{k}\b', lower):
            scale = m
            break

    # case "tiga koma lima ribu" => 3.5 * 1000
    if "koma" in lower:
        parts = re.split(r'\bkoma\b', lower, maxsplit=1)
        left_part = parts[0].strip()
        right_part = parts[1].strip() if len(parts) > 1 else ""

        left_tokens = re.findall(r'[a-zA-Z0-9]+', left_part)
        right_tokens = re.findall(r'[a-zA-Z0-9]+', right_part)

        left_digit = token_to_digit(left_tokens[-1]) if left_tokens else None
        right_digit = token_to_digit(right_tokens[0]) if right_tokens else None

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


def make_unique_filename_base(base_name: str) -> str:
    base_name = (base_name or "").strip() or "Dokumen"
    folder = str(FILES_DIR) if FILES_DIR else "static/files"
    os.makedirs(folder, exist_ok=True)

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


# =========================
# MoU Counter (mulai 000)
# =========================
def _mou_counter_path() -> str:
    folder = str(FILES_DIR) if FILES_DIR else "static/files"
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
    return str(n).zfill(3)


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
    kode_p3 = (mou_data.get("pihak_ketiga_kode") or "XXX").strip().upper() or "XXX"

    now = datetime.now()
    romawi = month_to_roman(now.month)
    tahun = str(now.year)

    return f"{no_depan}/PKPLNB3/{kode_p1}-{kode_p2}-{kode_p3}/{romawi}/{tahun}"


def format_tanggal_indonesia(dt: datetime) -> str:
    hari_map = {
        0: "Senin", 1: "Selasa", 2: "Rabu", 3: "Kamis",
        4: "Jumat", 5: "Sabtu", 6: "Minggu",
    }
    bulan_map = {
        1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
        5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
        9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }
    hari = hari_map.get(dt.weekday(), "")
    bulan = bulan_map.get(dt.month, "")
    return f"{hari}, tanggal {dt.day} {bulan} {dt.year}"


# =========================
# DOCX Helpers (format aman)
# =========================
def set_run_font(run, font_name="Times New Roman", size=10, bold=None):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:ascii'), font_name)
    run._element.rPr.rFonts.set(qn('w:hAnsi'), font_name)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run._element.rPr.rFonts.set(qn('w:cs'), font_name)
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold


def replace_in_runs_keep_format(paragraph, old: str, new: str):
    if not old or not paragraph.text or old not in paragraph.text:
        return False
    changed = False
    for run in paragraph.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)
            changed = True
    return changed


def replace_in_cell_keep_format(cell, old: str, new: str):
    changed = False
    for p in cell.paragraphs:
        if replace_in_runs_keep_format(p, old, new):
            changed = True
    return changed


def replace_everywhere_keep_format(doc, old_list, new_value):
    if not new_value:
        return
    for p in doc.paragraphs:
        for old in old_list:
            replace_in_runs_keep_format(p, old, new_value)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for old in old_list:
                    replace_in_cell_keep_format(cell, old, new_value)


def style_cell_paragraph(cell, align="left", left_indent_pt=0, font="Times New Roman", size=10):
    if not cell.paragraphs:
        cell.add_paragraph("")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if align == "center" else WD_ALIGN_PARAGRAPH.LEFT
    if left_indent_pt and align == "left":
        p.paragraph_format.left_indent = Pt(left_indent_pt)
    for r in p.runs:
        set_run_font(r, font, size)


def create_mou_docx(mou_data: dict, fname_base: str) -> str:
    template_path = "tamplate MoU.docx"
    if not os.path.exists(template_path):
        raise Exception("Template MoU tidak ditemukan. Pastikan file 'tamplate MoU.docx' ada di root project.")

    doc = Document(template_path)

    pihak1_raw = (mou_data.get("pihak_pertama") or "").strip()
    pihak2_raw = (mou_data.get("pihak_kedua") or "").strip()
    pihak3_raw = (mou_data.get("pihak_ketiga") or "").strip()

    pihak1 = pihak1_raw.upper()
    pihak2 = pihak2_raw.upper()
    pihak3 = pihak3_raw.upper()

    alamat1 = (mou_data.get("alamat_pihak_pertama") or "").strip()
    alamat3 = (mou_data.get("alamat_pihak_ketiga") or "").strip()

    ttd1 = (mou_data.get("ttd_pihak_pertama") or "").strip()
    jab1 = (mou_data.get("jabatan_pihak_pertama") or "").strip()
    ttd3 = (mou_data.get("ttd_pihak_ketiga") or "").strip()
    jab3 = (mou_data.get("jabatan_pihak_ketiga") or "").strip()

    nomor_full = (mou_data.get("nomor_surat") or "").strip()
    tanggal_text = format_tanggal_indonesia(datetime.now())

    contoh_pihak1_candidates = [
        "PT. PANPAN LUCKY INDONESIA",
        "PT. Panpan Lucky Indonesia",
        "PT PANPAN LUCKY INDONESIA",
        "PT Panpan Lucky Indonesia",
    ]
    contoh_pihak2_candidates = [
        "PT. SARANA TRANS BERSAMA JAYA",
        "PT Sarana Trans Bersama Jaya",
        "PT SARANA TRANS BERSAMA JAYA",
        "PT Sarana Trans Bersama Jaya",
    ]
    contoh_pihak3_candidates = [
        "PT. HARAPAN BARU SEJAHTERA PLASTIK",
        "PT Harapan Baru Sejahtera Plastik",
        "PT HARAPAN BARU SEJAHTERA PLASTIK",
        "PT Harapan Baru Sejahtera Plastik",
    ]

    replace_everywhere_keep_format(doc, contoh_pihak1_candidates, pihak1)
    replace_everywhere_keep_format(doc, contoh_pihak2_candidates, pihak2)
    replace_everywhere_keep_format(doc, contoh_pihak3_candidates, pihak3)

    def replace_no_line(container_paragraphs):
        for p in container_paragraphs:
            if re.search(r'\bNo\s*:', p.text, flags=re.IGNORECASE):
                for run in p.runs:
                    if re.search(r'\bNo\s*:', run.text, flags=re.IGNORECASE):
                        run.text = re.sub(r'\bNo\s*:\s*.*', f"No : {nomor_full}", run.text, flags=re.IGNORECASE)
                        return True
                return True
        return False

    replace_no_line(doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                if replace_no_line(cell.paragraphs):
                    break

    kalimat_tanggal = f"Pada hari ini {tanggal_text} kami yang bertanda tangan di bawah ini :"

    def replace_pada_hari_ini(container_paragraphs):
        for p in container_paragraphs:
            if "Pada hari ini" in p.text and "bertanda tangan" in p.text:
                if p.runs:
                    p.runs[0].text = kalimat_tanggal
                    for r in p.runs[1:]:
                        r.text = ""
                else:
                    p.add_run(kalimat_tanggal)
                return True
        return False

    replace_pada_hari_ini(doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                if replace_pada_hari_ini(cell.paragraphs):
                    break

    if ttd1:
        replace_everywhere_keep_format(doc, ["Huang Feifang"], ttd1)
    if jab1:
        replace_everywhere_keep_format(doc, ["Direktur Utama"], jab1)
    if alamat1:
        replace_everywhere_keep_format(doc, [
            "Jl. Raya Serang KM. 22 No. 30, Desa Pasir Bolang, Kec Tigaraksa, Tangerang Banten",
            "Jl. Raya Serang KM. 22 No. 30, Desa Pasir Bolang, Kec. Tigaraksa, Tangerang Banten",
        ], alamat1)

    if ttd3:
        replace_everywhere_keep_format(doc, ["Yogi Aditya", "Yogi Permana", "Yogi"], ttd3)
    if jab3:
        replace_everywhere_keep_format(doc, ["Direktur", "Direktur Utama"], jab3)
    if alamat3:
        replace_everywhere_keep_format(doc, [
            "Jl. Karawang – Bekasi KM. 1 Bojongsari, Kec. Kedungwaringin, Kab. Bekasi – Jawa Barat",
            "Jl. Karawang - Bekasi KM. 1 Bojongsari, Kec. Kedungwaringin, Kab. Bekasi - Jawa Barat",
        ], alamat3)

    items = mou_data.get("items_limbah") or []
    target_table = None
    for t in doc.tables:
        if not t.rows:
            continue
        header_text = " ".join([c.text.strip() for c in t.rows[0].cells])
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
                style_cell_paragraph(cells[0], align="center")

            if len(cells) >= 2:
                cells[1].text = (it.get("jenis_limbah") or "").strip()
                style_cell_paragraph(cells[1], align="left", left_indent_pt=6)

            if len(cells) >= 3:
                cells[2].text = (it.get("kode_limbah") or "").strip()
                style_cell_paragraph(cells[2], align="center")

    folder = str(FILES_DIR) if FILES_DIR else "static/files"
    os.makedirs(folder, exist_ok=True)
    out_path = os.path.join(folder, f"{fname_base}.docx")
    doc.save(out_path)
    return f"{fname_base}.docx"


def get_session_id() -> str:
    sid = request.headers.get("X-Session-ID") or session.get("sid")
    if not sid:
        sid = str(uuid.uuid4())
        session["sid"] = sid
    return sid


def ensure_history_user_message(history_id_in, text):
    try:
        if history_id_in:
            db_append_message(int(history_id_in), "user", text, files=[])
    except:
        pass


def ensure_history_assistant_message(history_id_in, out_text, files=None):
    try:
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=files or [])
    except:
        pass


# =========================
# Routes
# =========================
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


# =========================
# CHAT ROUTE
# =========================
@app.route("/api/chat", methods=["POST"])
def chat():
    try:
        data = request.get_json() or {}
        text = (data.get("message", "") or "").strip()
        history_id_in = data.get("history_id")

        if not text:
            return jsonify({"error": "Pesan kosong"}), 400

        sid = get_session_id()
        state = conversations.get(sid, {"step": "idle", "data": {}})

        lower = text.lower()

        # routing utama dari React
        task_type_req = (data.get("taskType") or "").strip().lower()  # invoice | quotation | mou
        if task_type_req == "penawaran":
            task_type_req = "quotation"

        # simpan user msg ke history kalau ada
        ensure_history_user_message(history_id_in, text)

        # =========================================================
        # START FLOW: INVOICE
        # =========================================================
        if state.get("step") == "idle" and (
            task_type_req == "invoice" or
            any(k in lower for k in ["invoice", "faktur", "tagihan", "invois", "invoys", "invoyce"])
        ):
            state["step"] = "invoice_bill_to"
            state["data"] = {
                "date": datetime.now().strftime("%d-%b-%y"),
                "bill_to_name": "",
                "bill_to_address": "",
                "ship_to_name": "",
                "ship_to_address": "",
                "attn": "Accounting / Finance",
                "sales_person": "",
                "items": [],
                "current_item": {},
                "ppn_percent": 11,
                "transfer_default": True,
            }
            conversations[sid] = state

            out_text = (
                "Baik, saya bantu buatkan <b>Invoice</b>.<br><br>"
                "❓ <b>1. Bill To (Nama Perusahaan)?</b>"
            )

            if not history_id_in:
                history_id_created = db_insert_history(
                    title="Chat Baru",
                    task_type="invoice",
                    data={},
                    files=[],
                    messages=[
                        {"id": uuid.uuid4().hex[:12], "sender": "user", "text": text, "files": [], "timestamp": datetime.now().isoformat()},
                        {"id": uuid.uuid4().hex[:12], "sender": "assistant", "text": re.sub(r'<br\s*/?>', '\n', out_text), "files": [], "timestamp": datetime.now().isoformat()},
                    ],
                    state=state
                )
                return jsonify({"text": out_text, "history_id": history_id_created})

            ensure_history_assistant_message(history_id_in, out_text)
            db_update_state(int(history_id_in), state)
            return jsonify({"text": out_text, "history_id": history_id_in})

        # =========================================================
        # START FLOW: MOU
        # =========================================================
        if state.get("step") == "idle" and (task_type_req == "mou" or "mou" in lower):
            nomor_depan = get_next_mou_no_depan()

            state["step"] = "mou_pihak_pertama"
            state["data"] = {
                "nomor_depan": nomor_depan,
                "nomor_surat": "",
                "items_limbah": [],
                "current_item": {},
                "pihak_kedua": "PT Sarana Trans Bersama Jaya",
                "pihak_kedua_kode": "STBJ",
                "pihak_pertama": "",
                "alamat_pihak_pertama": "",
                "pihak_ketiga": "",
                "pihak_ketiga_kode": "",
                "alamat_pihak_ketiga": "",
                "ttd_pihak_pertama": "",
                "jabatan_pihak_pertama": "",
                "ttd_pihak_ketiga": "",
                "jabatan_pihak_ketiga": "",
            }
            conversations[sid] = state

            out_text = (
                "Baik, saya bantu buatkan <b>MoU Tripartit</b>.<br><br>"
                f"✅ No Depan: <b>{nomor_depan}</b> (auto mulai 000)<br>"
                "✅ Nomor lengkap otomatis mengikuti format template.<br>"
                "✅ Tanggal otomatis hari ini.<br><br>"
                "❓ <b>1. Nama Perusahaan (PIHAK PERTAMA / Penghasil Limbah)?</b>"
            )

            if not history_id_in:
                history_id_created = db_insert_history(
                    title="Chat Baru",
                    task_type="mou",
                    data={},
                    files=[],
                    messages=[
                        {"id": uuid.uuid4().hex[:12], "sender": "user", "text": text, "files": [], "timestamp": datetime.now().isoformat()},
                        {"id": uuid.uuid4().hex[:12], "sender": "assistant", "text": re.sub(r'<br\s*/?>', '\n', out_text), "files": [], "timestamp": datetime.now().isoformat()},
                    ],
                    state=state
                )
                return jsonify({"text": out_text, "history_id": history_id_created})

            ensure_history_assistant_message(history_id_in, out_text)
            db_update_state(int(history_id_in), state)
            return jsonify({"text": out_text, "history_id": history_id_in})

        # =========================================================
        # START FLOW: QUOTATION
        # =========================================================
        if state.get("step") == "idle" and (
            task_type_req == "quotation" or
            any(k in lower for k in ["quotation", "penawaran", "kuotasi"])
        ):
            nomor_depan = get_next_nomor()
            now = datetime.now()

            state["step"] = "nama_perusahaan"
            state["data"] = {
                "nomor_depan": nomor_depan,
                "items_limbah": [],
                "bulan_romawi": now.strftime("%m"),
            }
            conversations[sid] = state

            out_text = (
                "Baik, saya bantu buatkan quotation.<br><br>"
                f"✅ Nomor Surat: <b>{nomor_depan}</b><br><br>"
                "❓ <b>1. Nama Perusahaan?</b>"
            )

            if not history_id_in:
                history_id_created = db_insert_history(
                    title="Chat Baru",
                    task_type="quotation",
                    data={},
                    files=[],
                    messages=[
                        {"id": uuid.uuid4().hex[:12], "sender": "user", "text": text, "files": [], "timestamp": datetime.now().isoformat()},
                        {"id": uuid.uuid4().hex[:12], "sender": "assistant", "text": re.sub(r'<br\s*/?>', '\n', out_text), "files": [], "timestamp": datetime.now().isoformat()},
                    ],
                    state=state
                )
                return jsonify({"text": out_text, "history_id": history_id_created})

            ensure_history_assistant_message(history_id_in, out_text)
            db_update_state(int(history_id_in), state)
            return jsonify({"text": out_text, "history_id": history_id_in})

        # =========================================================
        # CONTINUE FLOW: MOU
        # =========================================================
        if state.get("step") == "mou_pihak_pertama":
            state["data"]["pihak_pertama"] = text.strip()

            alamat = search_company_address(text).strip()
            if not alamat:
                alamat = search_company_address_ai(text).strip()
            if not alamat:
                alamat = "Di Tempat"

            state["data"]["alamat_pihak_pertama"] = alamat
            state["step"] = "mou_pilih_pihak_ketiga"
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

            ensure_history_assistant_message(history_id_in, out_text)
            if history_id_in:
                db_update_state(int(history_id_in), state)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "mou_pilih_pihak_ketiga":
            pilihan = text.strip().upper()
            mapping = {"1": "HBSP", "2": "KJL", "3": "MBI", "4": "CGA", "HBSP": "HBSP", "KJL": "KJL", "MBI": "MBI", "CGA": "CGA"}
            kode = mapping.get(pilihan)

            if not kode:
                out_text = "⚠️ Pilihan tidak valid.<br><br>Pilih PIHAK KETIGA:<br>1. HBSP<br>2. KJL<br>3. MBI<br>4. CGA"
                ensure_history_assistant_message(history_id_in, out_text)
                return jsonify({"text": out_text, "history_id": history_id_in})

            pihak3_nama_map = {"HBSP": "PT Harapan Baru Sejahtera Plastik", "KJL": "KJL", "MBI": "MBI", "CGA": "CGA"}
            pihak3_alamat_map = {"HBSP": "Jl. Karawang – Bekasi KM. 1 Bojongsari, Kec. Kedungwaringin, Kab. Bekasi – Jawa Barat", "KJL": "", "MBI": "", "CGA": ""}

            state["data"]["pihak_ketiga"] = pihak3_nama_map.get(kode, kode)
            state["data"]["pihak_ketiga_kode"] = kode
            state["data"]["alamat_pihak_ketiga"] = pihak3_alamat_map.get(kode, "")
            state["data"]["nomor_surat"] = build_mou_nomor_surat(state["data"])

            state["step"] = "mou_jenis_kode_limbah"
            state["data"]["current_item"] = {}
            conversations[sid] = state

            out_text = (
                f"✅ PIHAK KETIGA: <b>{state['data']['pihak_ketiga']}</b><br>"
                f"✅ Nomor MoU: <b>{state['data']['nomor_surat']}</b><br><br>"
                "📦 <b>Item #1</b><br>"
                "❓ <b>3. Sebutkan Jenis Limbah atau Kode Limbah</b><br>"
                "<i>(Contoh: 'A102d' atau 'aki baterai bekas' | atau ketik <b>NON B3</b> untuk manual)</i>"
            )

            ensure_history_assistant_message(history_id_in, out_text)
            if history_id_in:
                db_update_state(int(history_id_in), state)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "mou_jenis_kode_limbah":
            if is_non_b3_input(text):
                state["data"]["current_item"] = {"kode_limbah": "NON B3", "jenis_limbah": ""}
                state["step"] = "mou_manual_jenis_limbah"
                conversations[sid] = state

                out_text = (
                    "✅ Kode: <b>NON B3</b><br><br>"
                    "❓ <b>3A. Jenis Limbah (manual) apa?</b><br>"
                    "<i>(Contoh: 'plastik bekas', 'kertas bekas', dll)</i>"
                )
                ensure_history_assistant_message(history_id_in, out_text)
                if history_id_in:
                    db_update_state(int(history_id_in), state)
                return jsonify({"text": out_text, "history_id": history_id_in})

            kode, data_limbah = find_limbah_by_kode(text)
            if not (kode and data_limbah):
                kode, data_limbah = find_limbah_by_jenis(text)

            if kode and data_limbah:
                item = {"kode_limbah": kode, "jenis_limbah": data_limbah["jenis"]}
                state["data"]["items_limbah"].append(item)
                num = len(state["data"]["items_limbah"])

                state["step"] = "mou_tambah_item"
                state["data"]["current_item"] = {}
                conversations[sid] = state

                out_text = (
                    f"✅ Item #{num} tersimpan!<br>"
                    f"• Jenis: <b>{data_limbah['jenis']}</b><br>"
                    f"• Kode: <b>{kode}</b><br><br>"
                    "❓ <b>Tambah item lagi?</b> (ya/tidak)"
                )
                ensure_history_assistant_message(history_id_in, out_text)
                if history_id_in:
                    db_update_state(int(history_id_in), state)
                return jsonify({"text": out_text, "history_id": history_id_in})

            out_text = (
                f"❌ Maaf, limbah '<b>{text}</b>' tidak ditemukan.<br><br>"
                "Coba:<br>"
                "• Kode (A102d, B105d)<br>"
                "• Nama (aki baterai bekas)<br>"
                "• Atau ketik <b>NON B3</b>"
            )
            ensure_history_assistant_message(history_id_in, out_text)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "mou_manual_jenis_limbah":
            jenis = text.strip()
            state["data"]["items_limbah"].append({"kode_limbah": "NON B3", "jenis_limbah": jenis})
            num = len(state["data"]["items_limbah"])

            state["step"] = "mou_tambah_item"
            state["data"]["current_item"] = {}
            conversations[sid] = state

            out_text = (
                f"✅ Item #{num} tersimpan!<br>"
                f"• Jenis (manual): <b>{jenis}</b><br>"
                f"• Kode: <b>NON B3</b><br><br>"
                "❓ <b>Tambah item lagi?</b> (ya/tidak)"
            )
            ensure_history_assistant_message(history_id_in, out_text)
            if history_id_in:
                db_update_state(int(history_id_in), state)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "mou_tambah_item":
            if ("ya" in lower) or ("iya" in lower):
                num = len(state["data"]["items_limbah"])
                state["step"] = "mou_jenis_kode_limbah"
                state["data"]["current_item"] = {}
                conversations[sid] = state

                out_text = (
                    f"📦 <b>Item #{num+1}</b><br>"
                    "❓ <b>3. Sebutkan Jenis Limbah atau Kode Limbah</b><br>"
                    "<i>(Contoh: 'A102d' atau 'aki baterai bekas' | atau ketik <b>NON B3</b>)</i>"
                )
                ensure_history_assistant_message(history_id_in, out_text)
                if history_id_in:
                    db_update_state(int(history_id_in), state)
                return jsonify({"text": out_text, "history_id": history_id_in})

            if any(k in lower for k in ["tidak", "skip", "lewat", "gak", "nggak"]):
                state["step"] = "mou_ttd_pihak_pertama"
                conversations[sid] = state

                out_text = (
                    "✅ Data limbah selesai.<br><br>"
                    "❓ <b>Terakhir, siapa nama penandatangan PIHAK PERTAMA?</b>"
                )
                ensure_history_assistant_message(history_id_in, out_text)
                if history_id_in:
                    db_update_state(int(history_id_in), state)
                return jsonify({"text": out_text, "history_id": history_id_in})

            out_text = "⚠️ Mohon jawab <b>ya</b> atau <b>tidak</b>.<br><br>❓ <b>Tambah item lagi?</b>"
            ensure_history_assistant_message(history_id_in, out_text)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "mou_ttd_pihak_pertama":
            state["data"]["ttd_pihak_pertama"] = text.strip()
            state["step"] = "mou_jabatan_pihak_pertama"
            conversations[sid] = state

            out_text = "❓ <b>Jabatan penandatangan PIHAK PERTAMA?</b>"
            ensure_history_assistant_message(history_id_in, out_text)
            if history_id_in:
                db_update_state(int(history_id_in), state)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "mou_jabatan_pihak_pertama":
            state["data"]["jabatan_pihak_pertama"] = text.strip()
            state["step"] = "mou_ttd_pihak_ketiga"
            conversations[sid] = state

            out_text = "❓ <b>Nama penandatangan PIHAK KETIGA?</b>"
            ensure_history_assistant_message(history_id_in, out_text)
            if history_id_in:
                db_update_state(int(history_id_in), state)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "mou_ttd_pihak_ketiga":
            state["data"]["ttd_pihak_ketiga"] = text.strip()
            state["step"] = "mou_jabatan_pihak_ketiga"
            conversations[sid] = state

            out_text = "❓ <b>Jabatan penandatangan PIHAK KETIGA?</b>"
            ensure_history_assistant_message(history_id_in, out_text)
            if history_id_in:
                db_update_state(int(history_id_in), state)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "mou_jabatan_pihak_ketiga":
            state["data"]["jabatan_pihak_ketiga"] = text.strip()

            nama_pt_raw = state["data"].get("pihak_pertama", "").strip()
            safe_pt = re.sub(r'[^A-Za-z0-9 \-]+', '', nama_pt_raw).strip()
            safe_pt = re.sub(r'\s+', ' ', safe_pt).strip()

            base_fname = f"MoU - {safe_pt}" if safe_pt else "MoU - Perusahaan"
            fname_base = make_unique_filename_base(base_fname)

            if not state["data"].get("nomor_surat"):
                state["data"]["nomor_surat"] = build_mou_nomor_surat(state["data"])

            docx = create_mou_docx(state["data"], fname_base)
            pdf = create_pdf(fname_base)

            conversations[sid] = {"step": "idle", "data": {}}

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
                    json.dumps(state["data"], ensure_ascii=False),
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
                    data=state["data"],
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
                f"✅ Total Limbah: <b>{len(state['data'].get('items_limbah') or [])} item</b>"
            )

            db_append_message(history_id, "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=files)
            return jsonify({"text": out_text, "files": files, "history_id": history_id})

        # =========================================================
        # CONTINUE FLOW: QUOTATION (PENAWARAN)
        # =========================================================
        if state.get("step") == "nama_perusahaan":
            state["data"]["nama_perusahaan"] = text

            alamat = search_company_address(text).strip()
            if not alamat:
                alamat = search_company_address_ai(text).strip()
            if not alamat:
                alamat = "Di Tempat"

            state["data"]["alamat_perusahaan"] = alamat
            state["step"] = "jenis_kode_limbah"
            state["data"]["current_item"] = {}
            conversations[sid] = state

            out_text = (
                f"✅ Nama: <b>{text}</b><br>"
                f"✅ Alamat: <b>{alamat}</b><br><br>"
                "📦 <b>Item #1</b><br>"
                "❓ <b>2. Sebutkan Jenis Limbah atau Kode Limbah</b><br>"
                "<i>(Contoh: 'A102d' atau 'aki baterai bekas')</i>"
            )

            ensure_history_assistant_message(history_id_in, out_text)
            if history_id_in:
                db_update_state(int(history_id_in), state)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "jenis_kode_limbah":
            if is_non_b3_input(text):
                state["data"]["current_item"] = {"kode_limbah": "NON B3", "jenis_limbah": "", "satuan": ""}
                state["step"] = "manual_jenis_limbah"
                conversations[sid] = state

                out_text = (
                    "✅ Kode: <b>NON B3</b><br><br>"
                    "❓ <b>2A. Jenis Limbah (manual)?</b>"
                )
                ensure_history_assistant_message(history_id_in, out_text)
                if history_id_in:
                    db_update_state(int(history_id_in), state)
                return jsonify({"text": out_text, "history_id": history_id_in})

            kode, data_limbah = find_limbah_by_kode(text)
            if not (kode and data_limbah):
                kode, data_limbah = find_limbah_by_jenis(text)

            if kode and data_limbah:
                state["data"]["current_item"] = {
                    "kode_limbah": kode,
                    "jenis_limbah": data_limbah["jenis"],
                    "satuan": data_limbah["satuan"],
                }
                state["step"] = "harga"
                conversations[sid] = state

                out_text = (
                    f"✅ Kode: <b>{kode}</b><br>"
                    f"✅ Jenis: <b>{data_limbah['jenis']}</b><br>"
                    f"✅ Satuan: <b>{data_limbah['satuan']}</b><br><br>"
                    "❓ <b>3. Harga (Rp)?</b>"
                )
                ensure_history_assistant_message(history_id_in, out_text)
                if history_id_in:
                    db_update_state(int(history_id_in), state)
                return jsonify({"text": out_text, "history_id": history_id_in})

            out_text = (
                f"❌ Limbah '<b>{text}</b>' tidak ditemukan.<br><br>"
                "Coba kode (A102d) / nama (aki baterai bekas) / atau ketik <b>NON B3</b>."
            )
            ensure_history_assistant_message(history_id_in, out_text)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "manual_jenis_limbah":
            state["data"]["current_item"]["jenis_limbah"] = text
            state["step"] = "manual_satuan"
            conversations[sid] = state

            out_text = "❓ <b>2B. Satuan (manual)?</b> (kg/liter/drum/pcs)"
            ensure_history_assistant_message(history_id_in, out_text)
            if history_id_in:
                db_update_state(int(history_id_in), state)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "manual_satuan":
            state["data"]["current_item"]["satuan"] = text
            state["step"] = "harga"
            conversations[sid] = state

            out_text = "❓ <b>3. Harga (Rp)?</b>"
            ensure_history_assistant_message(history_id_in, out_text)
            if history_id_in:
                db_update_state(int(history_id_in), state)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "harga":
            harga_converted = parse_amount_id(text)
            state["data"]["current_item"]["harga"] = harga_converted

            state["data"]["items_limbah"].append(state["data"]["current_item"])
            num = len(state["data"]["items_limbah"])

            state["step"] = "tambah_item"
            state["data"]["current_item"] = {}
            conversations[sid] = state

            out_text = (
                f"✅ Item #{num} tersimpan!<br>"
                f"💰 Harga: <b>Rp {format_rupiah(harga_converted)}</b><br><br>"
                "❓ <b>Tambah item lagi?</b> (ya/tidak)"
            )
            ensure_history_assistant_message(history_id_in, out_text)
            if history_id_in:
                db_update_state(int(history_id_in), state)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "tambah_item":
            if ("ya" in lower) or ("iya" in lower):
                num = len(state["data"]["items_limbah"])
                state["step"] = "jenis_kode_limbah"
                state["data"]["current_item"] = {}
                conversations[sid] = state

                out_text = (
                    f"📦 <b>Item #{num+1}</b><br>"
                    "❓ <b>2. Sebutkan Jenis Limbah atau Kode Limbah</b>"
                )
                ensure_history_assistant_message(history_id_in, out_text)
                if history_id_in:
                    db_update_state(int(history_id_in), state)
                return jsonify({"text": out_text, "history_id": history_id_in})

            if any(k in lower for k in ["tidak", "skip", "lewat", "gak", "nggak"]):
                state["step"] = "harga_transportasi"
                conversations[sid] = state

                out_text = (
                    f"✅ Total: <b>{len(state['data']['items_limbah'])} item</b><br><br>"
                    "❓ <b>4. Biaya Transportasi (Rp)?</b><br>"
                    "<i>Satuan: ritase</i>"
                )
                ensure_history_assistant_message(history_id_in, out_text)
                if history_id_in:
                    db_update_state(int(history_id_in), state)
                return jsonify({"text": out_text, "history_id": history_id_in})

            out_text = "⚠️ Mohon jawab <b>ya</b> atau <b>tidak</b>.<br><br>❓ <b>Tambah item lagi?</b>"
            ensure_history_assistant_message(history_id_in, out_text)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "harga_transportasi":
            transportasi = parse_amount_id(text)
            state["data"]["harga_transportasi"] = transportasi
            state["step"] = "tanya_mou"
            conversations[sid] = state

            out_text = (
                f"✅ Transportasi: <b>Rp {format_rupiah(transportasi)}/ritase</b><br><br>"
                "❓ <b>5. Tambah Biaya MoU?</b> (ya/tidak)"
            )
            ensure_history_assistant_message(history_id_in, out_text)
            if history_id_in:
                db_update_state(int(history_id_in), state)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "tanya_mou":
            if ("ya" in lower) or ("iya" in lower):
                state["step"] = "harga_mou"
                conversations[sid] = state
                out_text = "❓ <b>Biaya MoU (Rp)?</b> <i>(Satuan: Tahun)</i>"
                ensure_history_assistant_message(history_id_in, out_text)
                if history_id_in:
                    db_update_state(int(history_id_in), state)
                return jsonify({"text": out_text, "history_id": history_id_in})

            if any(k in lower for k in ["tidak", "skip", "lewat", "gak", "nggak"]):
                state["data"]["harga_mou"] = None
                state["step"] = "tanya_termin"
                conversations[sid] = state

                out_text = "❓ <b>6. Edit Termin Pembayaran?</b><br><i>Default: 14 hari</i>"
                ensure_history_assistant_message(history_id_in, out_text)
                if history_id_in:
                    db_update_state(int(history_id_in), state)
                return jsonify({"text": out_text, "history_id": history_id_in})

            out_text = "⚠️ Mohon jawab <b>ya</b> atau <b>tidak</b>."
            ensure_history_assistant_message(history_id_in, out_text)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "harga_mou":
            mou_cost = parse_amount_id(text)
            state["data"]["harga_mou"] = mou_cost
            state["step"] = "tanya_termin"
            conversations[sid] = state

            out_text = (
                f"✅ MoU: <b>Rp {format_rupiah(mou_cost)}/Tahun</b><br><br>"
                "❓ <b>6. Edit Termin Pembayaran?</b><br><i>Default: 14 hari</i>"
            )
            ensure_history_assistant_message(history_id_in, out_text)
            if history_id_in:
                db_update_state(int(history_id_in), state)
            return jsonify({"text": out_text, "history_id": history_id_in})

        if state.get("step") == "tanya_termin":
            if any(k in lower for k in ["tidak", "skip", "lewat"]):
                state["data"]["termin_hari"] = "14"
            else:
                state["data"]["termin_hari"] = parse_termin_days(text, default=14, min_days=1, max_days=365)

            nama_pt_raw = (state["data"].get("nama_perusahaan") or "").strip()
            safe_pt = re.sub(r'[^A-Za-z0-9 \-]+', '', nama_pt_raw).strip()
            safe_pt = re.sub(r'\s+', ' ', safe_pt).strip()

            base_fname = f"Quotation - {safe_pt}" if safe_pt else "Quotation - Penawaran"
            fname = make_unique_filename_base(base_fname)

            docx = create_docx(state["data"], fname)
            pdf = create_pdf(fname)

            conversations[sid] = {"step": "idle", "data": {}}

            files = [{"type": "docx", "filename": docx, "url": f"/static/files/{docx}"}]
            if pdf:
                files.append({"type": "pdf", "filename": pdf, "url": f"/static/files/{pdf}"})

            history_title = f"Penawaran {nama_pt_raw}" if nama_pt_raw else "Penawaran"
            history_task_type = "quotation"

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
                    json.dumps(state["data"], ensure_ascii=False),
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
                    data=state["data"],
                    files=files,
                    messages=[],
                    state={}
                )

            termin_terbilang = angka_ke_terbilang(state["data"]["termin_hari"])
            out_text = (
                f"✅ Termin: <b>{state['data']['termin_hari']} ({termin_terbilang}) hari</b><br><br>"
                "🎉 <b>Quotation berhasil dibuat!</b>"
            )

            db_append_message(history_id, "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=files)
            return jsonify({"text": out_text, "files": files, "history_id": history_id})

        # =========================================================
        # DEFAULT (AI)
        # =========================================================
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

    print("\n" + "=" * 60)
    print("🚀 DOCUMENT GENERATOR")
    print("=" * 60)
    print(f"🔑 OpenRouter: {OPENROUTER_API_KEY and '✅' or '❌'}")
    print(f"🔎 Serper: {SERPER_API_KEY and '✅' or '❌'}")
    print(f"📄 PDF: {PDF_AVAILABLE and f'✅ {PDF_METHOD}' or '❌ Disabled'}")
    print(f"🗄️  Database: {len(LIMBAH_B3_DB)} jenis limbah")
    print(f"🔢 Counter: {load_counter()}")
    print(f"🌐 Port: {port}")
    print(f"💻 Platform: {platform.system()}")
    print("=" * 60 + "\n")

    app.run(host="0.0.0.0", port=port, debug=debug_mode)
