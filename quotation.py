import re
import json
from datetime import datetime

from limbah_database import (
    find_limbah_by_kode,
    find_limbah_by_jenis,
    convert_voice_to_number,
    parse_termin_days,
    angka_ke_terbilang,
    format_rupiah,
)
from utils import (
    db_insert_history, db_append_message, db_update_state,
    get_next_nomor, create_docx, create_pdf,
    search_company_address, search_company_address_ai,
)

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
    t = re.sub(r'(?<=\d)[\.,](?=\d{3}(\D|$))', '', t)
    t = re.sub(r'(?<=\d),(?=\d)', '.', t)
    return t

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
        if re.search(rf'\b{k}\b', lower):
            scale = m
            break

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


# =========================
# CHAT HANDLER QUOTATION
# =========================

def handle_quotation_flow(data: dict, text: str, lower: str, sid: str, state: dict, conversations: dict, history_id_in):
    """
    Return:
      - None jika bukan flow quotation
      - dict response jika handled
    """

    if ('quotation' in lower or 'penawaran' in lower or ('buat' in lower and 'mou' not in lower)) and (state.get("step") == "idle"):
        nomor_depan = get_next_nomor()
        state['step'] = 'nama_perusahaan'
        now = datetime.now()
        state['data'] = {'nomor_depan': nomor_depan, 'items_limbah': [], 'bulan_romawi': now.strftime('%m')}
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
                    {"id": __import__("uuid").uuid4().hex[:12], "sender": "user", "text": text, "files": [], "timestamp": datetime.now().isoformat()},
                    {"id": __import__("uuid").uuid4().hex[:12], "sender": "assistant", "text": re.sub(r'<br\s*/?>', '\n', out_text), "files": [], "timestamp": datetime.now().isoformat()},
                ],
                state=state
            )
        else:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_created or history_id_in}

    if state.get("step") == 'nama_perusahaan':
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
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == 'jenis_kode_limbah':
        if is_non_b3_input(text):
            state['data']['current_item']['kode_limbah'] = "NON B3"
            state['data']['current_item']['jenis_limbah'] = ""
            state['data']['current_item']['satuan'] = ""
            state['step'] = 'manual_jenis_limbah'
            conversations[sid] = state

            out_text = "✅ Kode: <b>NON B3</b><br><br>❓ <b>2A. Jenis Limbah (manual) apa?</b>"
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

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
            return {"text": out_text, "history_id": history_id_in}
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
                return {"text": out_text, "history_id": history_id_in}
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
                return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == 'manual_jenis_limbah':
        state['data']['current_item']['jenis_limbah'] = text
        state['step'] = 'manual_satuan'
        conversations[sid] = state
        out_text = f"✅ Jenis (manual): <b>{text}</b><br><br>❓ <b>2B. Satuan (manual) apa?</b>"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == 'manual_satuan':
        state['data']['current_item']['satuan'] = text
        state['step'] = 'harga'
        conversations[sid] = state
        out_text = f"✅ Satuan (manual): <b>{text}</b><br><br>❓ <b>3. Harga (Rp)?</b>"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == 'harga':
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
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == 'tambah_item':
        if re.match(r'^\d+', text.strip()):
            out_text = "⚠️ Mohon jawab dengan <b>'ya'</b> atau <b>'tidak'</b><br><br>❓ <b>Tambah item lagi?</b>"
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            return {"text": out_text, "history_id": history_id_in}

        if 'ya' in lower or 'iya' in lower:
            num = len(state['data']['items_limbah'])
            state['step'] = 'jenis_kode_limbah'
            state['data']['current_item'] = {}
            conversations[sid] = state
            out_text = f"📦 <b>Item #{num+1}</b><br>❓ <b>2. Sebutkan Jenis Limbah atau Kode Limbah</b><br><i>(Contoh: 'A102d' atau 'aki baterai bekas')</i>"
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

        if 'tidak' in lower or 'skip' in lower or 'lewat' in lower or 'gak' in lower or 'nggak' in lower:
            state['step'] = 'harga_transportasi'
            conversations[sid] = state
            out_text = f"✅ Total: <b>{len(state['data']['items_limbah'])} item</b><br><br>❓ <b>4. Biaya Transportasi (Rp)?</b><br><i>Satuan: ritase</i>"
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

        out_text = "⚠️ Mohon jawab dengan <b>'ya'</b> atau <b>'tidak'</b><br><br>❓ <b>Tambah item lagi?</b>"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == 'harga_transportasi':
        transportasi_converted = parse_amount_id(text)
        state['data']['harga_transportasi'] = transportasi_converted
        state['step'] = 'tanya_mou'
        conversations[sid] = state
        transportasi_formatted = format_rupiah(transportasi_converted)
        out_text = f"✅ Transportasi: <b>Rp {transportasi_formatted}/ritase</b><br><br>❓ <b>5. Tambah Biaya MoU?</b> (ya/tidak)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == 'tanya_mou':
        if re.match(r'^\d+', text.strip()):
            out_text = "⚠️ Mohon jawab dengan <b>'ya'</b> atau <b>'tidak'</b><br><br>❓ <b>5. Tambah Biaya MoU?</b> (ya/tidak)"
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            return {"text": out_text, "history_id": history_id_in}

        if 'ya' in lower or 'iya' in lower:
            state['step'] = 'harga_mou'
            conversations[sid] = state
            out_text = "❓ <b>Biaya MoU (Rp)?</b><br><i>Satuan: Tahun</i>"
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

        if 'tidak' in lower or 'skip' in lower or 'lewat' in lower or 'gak' in lower or 'nggak' in lower:
            state['data']['harga_mou'] = None
            state['step'] = 'tanya_termin'
            conversations[sid] = state
            out_text = "❓ <b>6. Edit Termin Pembayaran?</b><br><i>Default: 14 hari</i><br>(ketik angka atau 'tidak' untuk default)"
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

        out_text = "⚠️ Mohon jawab dengan <b>'ya'</b> atau <b>'tidak'</b><br><br>❓ <b>5. Tambah Biaya MoU?</b> (ya/tidak)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == 'harga_mou':
        mou_converted = parse_amount_id(text)
        state['data']['harga_mou'] = mou_converted
        state['step'] = 'tanya_termin'
        conversations[sid] = state
        mou_formatted = format_rupiah(mou_converted)
        out_text = f"✅ MoU: <b>Rp {mou_formatted}/Tahun</b><br><br>❓ <b>6. Edit Termin Pembayaran?</b><br><i>Default: 14 hari</i><br>(ketik angka atau 'tidak' untuk default)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == 'tanya_termin':
        if 'tidak' in lower or 'skip' in lower or 'lewat' in lower:
            state['data']['termin_hari'] = '14'
        else:
            state['data']['termin_hari'] = parse_termin_days(text, default=14, min_days=1, max_days=365)

        nama_pt_raw = state['data'].get('nama_perusahaan', '').strip()
        safe_pt = re.sub(r'[^A-Za-z0-9 \-]+', '', nama_pt_raw).strip()
        safe_pt = re.sub(r'\s+', ' ', safe_pt).strip()
        base_fname = f"Quotation - {safe_pt}" if safe_pt else "Quotation - Penawaran"
        fname = (lambda base: base if base else "Quotation - Penawaran")(base_fname)

        # pakai helper asli kamu: make_unique_filename_base ada di file awal,
        # tapi karena modul ini terpisah dan user minta 4 file, kita pakai pola nama tetap.
        # (Kalau kamu mau persis sama unique filename, tinggal copy fungsi itu ke sini)
        fname = base_fname

        docx = create_docx(state['data'], fname)
        pdf = create_pdf(fname)

        conversations[sid] = {'step': 'idle', 'data': {}}

        files = [{"type": "docx", "filename": docx, "url": f"/download/{docx}"}]
        if pdf:
            files.append({"type": "pdf", "filename": pdf, "url": f"/download/{pdf}"})

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

        return {"text": out_text, "files": files, "history_id": history_id}

    return None
