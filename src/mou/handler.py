import json
import re
from datetime import datetime

from limbah_database import (
    find_limbah_by_kode,
    find_limbah_by_jenis,
)

from utils import (
    db_insert_history, db_append_message, db_update_state,
    create_pdf,
)

from .helpers import (
    is_non_b3_input,
    make_unique_filename_base,
    build_mou_nomor_surat,
    resolve_company_address,
    get_next_mou_no_depan,
)

from .docx_builder import create_mou_docx


# =========================
# CHAT HANDLER MoU
# =========================

def handle_mou_flow(data: dict, text: str, lower: str, sid: str, state: dict, conversations: dict, history_id_in):
    """
    Return:
      - None jika bukan flow mou
      - dict response jika handled
    """

    if ('mou' in lower) and (state.get('step') == 'idle'):
        nomor_depan = get_next_mou_no_depan()
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
            'pihak_ketiga_kode': "",
            'alamat_pihak_ketiga': "",
            'ttd_pihak_pertama': "",
            'jabatan_pihak_pertama': "",
            # ✅ penanganan nama & jabatan pihak ketiga DIHAPUS
        }
        conversations[sid] = state

        out_text = (
            "Baik, saya akan membantu menyusun <b>MoU</b>.<br><br>"
            "Pertanyaan 1: <b>Nama perusahaan (Pihak Pertama / Penghasil Limbah)?</b>"
        )

        history_id_created = None
        if not history_id_in:
            history_id_created = db_insert_history(
                title="Chat Baru",
                task_type=data.get("taskType") or "mou",
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

    if state.get('step') == 'mou_pihak_pertama':
        state['data']['pihak_pertama'] = text.strip()

        alamat = resolve_company_address(text)
        state['data']['alamat_pihak_pertama'] = alamat

        state['step'] = 'mou_pilih_pihak_ketiga'
        conversations[sid] = state

        out_text = (
            f"Pihak Pertama: <b>{state['data']['pihak_pertama']}</b><br>"
            f"Alamat: <b>{alamat}</b><br><br>"
            "Pertanyaan 2: <b>Pilih Pihak Ketiga (Pengelola Limbah)</b><br>"
            "1. HBSP<br>2. KJL<br>3. MBI<br>4. CGA<br><br>"
            "<i>Ketik nomor 1-4 atau ketik HBSP/KJL/MBI/CGA</i>"
        )

        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)

        return {"text": out_text, "history_id": history_id_in}

    if state.get('step') == 'mou_pilih_pihak_ketiga':
        pilihan = text.strip().upper()
        mapping = {"1": "HBSP", "2": "KJL", "3": "MBI", "4": "CGA", "HBSP": "HBSP", "KJL": "KJL", "MBI": "MBI", "CGA": "CGA"}
        kode = mapping.get(pilihan)
        if not kode:
            out_text = (
                "Input tidak valid.<br><br>"
                "Pilih Pihak Ketiga:<br>"
                "1. HBSP<br>2. KJL<br>3. MBI<br>4. CGA"
            )
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            return {"text": out_text, "history_id": history_id_in}

        pihak3_nama_map = {"HBSP": "PT Harapan Baru Sejahtera Plastik", "KJL": "KJL", "MBI": "MBI", "CGA": "CGA"}
        pihak3_alamat_map = {
            "HBSP": "Jl. Karawang – Bekasi KM. 1 Bojongsari, Kec. Kedungwaringin, Kab. Bekasi – Jawa Barat",
            "KJL": "",
            "MBI": "",
            "CGA": "",
        }

        state['data']['pihak_ketiga'] = pihak3_nama_map.get(kode, kode)
        state['data']['pihak_ketiga_kode'] = kode
        state['data']['alamat_pihak_ketiga'] = pihak3_alamat_map.get(kode, "")

        state['data']['nomor_surat'] = build_mou_nomor_surat(state['data'])
        state['step'] = 'mou_jenis_kode_limbah'
        state['data']['current_item'] = {}
        conversations[sid] = state

        out_text = (
            f"Pihak Ketiga: <b>{state['data']['pihak_ketiga']}</b><br>"
            f"Nomor MoU: <b>{state['data']['nomor_surat']}</b><br><br>"
            "Item 1<br>"
            "Pertanyaan 3: <b>Sebutkan jenis limbah atau kode limbah</b><br>"
            "<i>Contoh: A102d atau aki baterai bekas. Atau ketik NON B3 untuk input manual.</i>"
        )

        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)

        return {"text": out_text, "history_id": history_id_in}

    if state.get('step') == 'mou_jenis_kode_limbah':
        if is_non_b3_input(text):
            state['data']['current_item']['kode_limbah'] = "NON B3"
            state['data']['current_item']['jenis_limbah'] = ""
            state['step'] = 'mou_manual_jenis_limbah'
            conversations[sid] = state
            out_text = (
                "Kode limbah: <b>NON B3</b><br><br>"
                "Pertanyaan 3A: <b>Jenis limbah (manual) apa?</b>"
            )
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

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
                f"Item {num} tersimpan.<br>"
                f"Jenis: <b>{data_limbah['jenis']}</b><br>"
                f"Kode: <b>{kode}</b><br><br>"
                "Pertanyaan: <b>Tambah item lagi?</b> (ya/tidak)"
            )
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

        out_text = (
            f"Maaf, limbah <b>{text}</b> tidak ditemukan.<br><br>"
            "Silakan ketik kode/jenis lain atau ketik <b>NON B3</b> untuk input manual."
        )
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get('step') == 'mou_manual_jenis_limbah':
        state['data']['current_item']['jenis_limbah'] = text.strip()
        state['data']['items_limbah'].append(state['data']['current_item'])
        num = len(state['data']['items_limbah'])
        state['step'] = 'mou_tambah_item'
        state['data']['current_item'] = {}
        conversations[sid] = state
        out_text = (
            f"Item {num} tersimpan.<br>"
            f"Jenis (manual): <b>{state['data']['items_limbah'][-1]['jenis_limbah']}</b><br>"
            "Kode: <b>NON B3</b><br><br>"
            "Pertanyaan: <b>Tambah item lagi?</b> (ya/tidak)"
        )
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get('step') == 'mou_tambah_item':
        if re.match(r'^\d+', text.strip()):
            out_text = (
                "Mohon jawab dengan <b>ya</b> atau <b>tidak</b>.<br><br>"
                "Pertanyaan: <b>Tambah item lagi?</b>"
            )
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            return {"text": out_text, "history_id": history_id_in}

        if ('ya' in lower) or ('iya' in lower):
            num = len(state['data']['items_limbah'])
            state['step'] = 'mou_jenis_kode_limbah'
            state['data']['current_item'] = {}
            conversations[sid] = state
            out_text = (
                f"Item {num+1}<br>"
                "Pertanyaan 3: <b>Sebutkan jenis limbah atau kode limbah</b><br>"
                "<i>Contoh: A102d. Atau ketik NON B3 untuk input manual.</i>"
            )
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

        if ('tidak' in lower) or ('skip' in lower) or ('lewat' in lower) or ('gak' in lower) or ('nggak' in lower):
            state['step'] = 'mou_ttd_pihak_pertama'
            conversations[sid] = state
            out_text = "Pertanyaan: <b>Nama penandatangan Pihak Pertama?</b>"
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

        out_text = (
            "Mohon jawab dengan <b>ya</b> atau <b>tidak</b>.<br><br>"
            "Pertanyaan: <b>Tambah item lagi?</b>"
        )
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
        return {"text": out_text, "history_id": history_id_in}

    if state.get('step') == 'mou_ttd_pihak_pertama':
        state['data']['ttd_pihak_pertama'] = text.strip()
        state['step'] = 'mou_jabatan_pihak_pertama'
        conversations[sid] = state
        out_text = "Pertanyaan: <b>Jabatan penandatangan Pihak Pertama?</b>"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get('step') == 'mou_jabatan_pihak_pertama':
        state['data']['jabatan_pihak_pertama'] = text.strip()

        # ✅ Penanganan pihak ketiga (nama/jabatan) dihapus -> langsung generate dokumen
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

        files = [{"type": "docx", "filename": docx, "url": f"/download/{docx}"}]
        if pdf:
            files.append({"type": "pdf", "filename": pdf, "url": f"/download/{pdf}"})

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
            "<b>MoU berhasil dibuat.</b><br><br>"
        )

        db_append_message(history_id, "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=files)
        return {"text": out_text, "files": files, "history_id": history_id}

    return None
