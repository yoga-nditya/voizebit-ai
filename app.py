import os
import json
import uuid
import re
from flask import Flask, request, jsonify, render_template, send_from_directory, session
from datetime import datetime
import platform

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


# ✅ TAMBAHAN: helper untuk deteksi NON B3 (berbagai variasi penulisan)
def is_non_b3_input(text: str) -> bool:
    if not text:
        return False
    t = text.strip().lower()
    # normalisasi: hilangkan spasi/underscore/dash
    norm = re.sub(r'[\s\-_]+', '', t)
    return norm in ("nonb3", "nonbii3") or norm.startswith("nonb3")


# ✅ TAMBAHAN: normalisasi angka format Indonesia agar "3.000" tidak kebaca 3.0
def normalize_id_number_text(text: str) -> str:
    if not text:
        return text
    t = text.strip()

    # Hapus titik sebagai pemisah ribuan: 3.000 -> 3000, 12.345.678 -> 12345678
    # (hanya hapus titik yang diikuti tepat 3 digit)
    t = re.sub(r'(?<=\d)\.(?=\d{3}(\D|$))', '', t)

    # Ubah koma desimal jadi titik: 3,5 -> 3.5
    t = re.sub(r'(?<=\d),(?=\d)', '.', t)

    return t


# ✅ TAMBAHAN: buat nama file unik (Quotation - Nama PT, Quotation - Nama PT (2), dst)
def make_unique_filename_base(base_name: str) -> str:
    base_name = (base_name or "").strip()
    if not base_name:
        base_name = "Quotation - Penawaran"

    # FILES_DIR dari config_new (dipakai juga di route download)
    try:
        folder = str(FILES_DIR)
    except Exception:
        folder = "static/files"

    def exists_any(name: str) -> bool:
        # cek kemungkinan docx/pdf (dan kalau sistem Anda pernah pakai underscore, tetap aman)
        return (
            os.path.exists(os.path.join(folder, f"{name}.docx")) or
            os.path.exists(os.path.join(folder, f"{name}.pdf")) or
            os.path.exists(os.path.join(folder, name))  # kalau create_docx mengembalikan full name langsung
        )

    if not exists_any(base_name):
        return base_name

    i = 2
    while True:
        candidate = f"{base_name} ({i})"
        if not exists_any(candidate):
            return candidate
        i += 1


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

        if 'quotation' in lower or 'penawaran' in lower or 'buat' in lower:
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
            # ✅ TAMBAHAN: jika input adalah NON B3 -> masuk manual flow
            if is_non_b3_input(text):
                state['data']['current_item']['kode_limbah'] = "NON B3"
                state['data']['current_item']['jenis_limbah'] = ""   # akan diisi manual
                state['data']['current_item']['satuan'] = ""         # akan diisi manual
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

        # ✅ TAMBAHAN: Step manual untuk NON B3
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
            # ✅ REVISI: normalisasi "3.000" supaya tidak jadi 3.0
            harga_converted = convert_voice_to_number(normalize_id_number_text(text))
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

        elif state['step'] == 'tambah_item':
            if re.match(r'^\d+', text.strip()):
                out_text = "⚠️ Mohon jawab dengan <b>'ya'</b> atau <b>'tidak'</b><br><br>❓ <b>Tambah item lagi?</b>"

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])

                return jsonify({"text": out_text, "history_id": history_id_in})

            if 'ya' in lower or 'iya' in lower:
                num = len(state['data']['items_limbah'])
                state['step'] = 'jenis_kode_limbah'
                state['data']['current_item'] = {}
                conversations[sid] = state
                out_text = f"📦 <b>Item #{num+1}</b><br>❓ <b>2. Sebutkan Jenis Limbah atau Kode Limbah</b><br><i>(Contoh: 'A102d' atau 'aki baterai bekas')</i>"

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                    db_update_state(int(history_id_in), state)

                return jsonify({"text": out_text, "history_id": history_id_in})
            elif 'tidak' in lower or 'skip' in lower or 'lewat' in lower or 'gak' in lower or 'nggak' in lower:
                state['step'] = 'harga_transportasi'
                conversations[sid] = state
                out_text = f"✅ Total: <b>{len(state['data']['items_limbah'])} item</b><br><br>❓ <b>4. Biaya Transportasi (Rp)?</b><br><i>Satuan: ritase</i>"

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                    db_update_state(int(history_id_in), state)

                return jsonify({"text": out_text, "history_id": history_id_in})
            else:
                out_text = "⚠️ Mohon jawab dengan <b>'ya'</b> atau <b>'tidak'</b><br><br>❓ <b>Tambah item lagi?</b>"

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])

                return jsonify({"text": out_text, "history_id": history_id_in})

        elif state['step'] == 'harga_transportasi':
            # ✅ REVISI: normalisasi angka
            transportasi_converted = convert_voice_to_number(normalize_id_number_text(text))
            state['data']['harga_transportasi'] = transportasi_converted
            state['step'] = 'tanya_mou'
            conversations[sid] = state
            transportasi_formatted = format_rupiah(transportasi_converted)
            out_text = f"✅ Transportasi: <b>Rp {transportasi_formatted}/ritase</b><br><br>❓ <b>5. Tambah Biaya MoU?</b> (ya/tidak)"

            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)

            return jsonify({"text": out_text, "history_id": history_id_in})

        elif state['step'] == 'tanya_mou':
            if re.match(r'^\d+', text.strip()):
                out_text = "⚠️ Mohon jawab dengan <b>'ya'</b> atau <b>'tidak'</b><br><br>❓ <b>5. Tambah Biaya MoU?</b> (ya/tidak)"

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])

                return jsonify({"text": out_text, "history_id": history_id_in})

            if 'ya' in lower or 'iya' in lower:
                state['step'] = 'harga_mou'
                conversations[sid] = state
                out_text = "❓ <b>Biaya MoU (Rp)?</b><br><i>Satuan: Tahun</i>"

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                    db_update_state(int(history_id_in), state)

                return jsonify({"text": out_text, "history_id": history_id_in})
            elif 'tidak' in lower or 'skip' in lower or 'lewat' in lower or 'gak' in lower or 'nggak' in lower:
                state['data']['harga_mou'] = None
                state['step'] = 'tanya_termin'
                conversations[sid] = state
                out_text = "❓ <b>6. Edit Termin Pembayaran?</b><br><i>Default: 14 hari</i><br>(ketik angka atau 'tidak' untuk default)"

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                    db_update_state(int(history_id_in), state)

                return jsonify({"text": out_text, "history_id": history_id_in})
            else:
                out_text = "⚠️ Mohon jawab dengan <b>'ya'</b> atau <b>'tidak'</b><br><br>❓ <b>5. Tambah Biaya MoU?</b> (ya/tidak)"

                if history_id_in:
                    db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])

                return jsonify({"text": out_text, "history_id": history_id_in})

        elif state['step'] == 'harga_mou':
            # ✅ REVISI: normalisasi angka
            mou_converted = convert_voice_to_number(normalize_id_number_text(text))
            state['data']['harga_mou'] = mou_converted
            state['step'] = 'tanya_termin'
            conversations[sid] = state

            mou_formatted = format_rupiah(mou_converted)
            out_text = f"✅ MoU: <b>Rp {mou_formatted}/Tahun</b><br><br>❓ <b>6. Edit Termin Pembayaran?</b><br><i>Default: 14 hari</i><br>(ketik angka atau 'tidak' untuk default)"

            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)

            return jsonify({"text": out_text, "history_id": history_id_in})

        elif state['step'] == 'tanya_termin':
            if 'tidak' in lower or 'skip' in lower or 'lewat' in lower:
                state['data']['termin_hari'] = '14'
            else:
                state['data']['termin_hari'] = parse_termin_days(text, default=14, min_days=1, max_days=365)

            # ✅ REVISI: nama file jadi "Quotation - (Nama PT)" + unik (2), (3), dst
            nama_pt_raw = state['data'].get('nama_perusahaan', '').strip()
            safe_pt = re.sub(r'[^A-Za-z0-9 \-]+', '', nama_pt_raw).strip()
            safe_pt = re.sub(r'\s+', ' ', safe_pt).strip()
            base_fname = f"Quotation - {safe_pt}" if safe_pt else "Quotation - Penawaran"
            fname = make_unique_filename_base(base_fname)

            docx = create_docx(state['data'], fname)
            pdf = create_pdf(fname)

            conversations[sid] = {'step': 'idle', 'data': {}}

            files = [{"type": "docx", "filename": docx, "url": f"/static/files/{docx}"}]
            if pdf:
                files.append({"type": "pdf", "filename": pdf, "url": f"/static/files/{pdf}"})

            nama_pt = state['data'].get('nama_perusahaan', '').strip()
            history_title = f"Penawaran {nama_pt}" if nama_pt else "Penawaran"
            history_task_type = "penawaran"

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

            termin_terbilang = angka_ke_terbilang(state['data']['termin_hari'])
            out_text = f"✅ Termin: <b>{state['data']['termin_hari']} ({termin_terbilang}) hari</b><br><br>🎉 <b>Quotation berhasil dibuat!</b>"

            db_append_message(history_id, "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=files)

            return jsonify({
                "text": out_text,
                "files": files,
                "history_id": history_id
            })

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
    print(f"🌐 Port: {port}")
    print(f"💻 Platform: {platform.system()}")
    print("="*60 + "\n")

    app.run(host="0.0.0.0", port=port, debug=debug_mode)
