import json
import re
from datetime import datetime

from limbah_database import find_limbah_by_kode, find_limbah_by_jenis

from utils import db_insert_history, db_append_message, db_update_state

from .helpers import (
    normalize_voice_strip,
    resolve_company_address,
    parse_qty_id,
    parse_amount_id,
    is_non_b3_input,
    make_unique_filename_base,
    get_next_invoice_no,
)

from .excel_builder import create_invoice_xlsx
from .pdf_builder import create_invoice_pdf


def handle_invoice_flow(data: dict, text: str, lower: str, sid: str, state: dict, conversations: dict, history_id_in):
    text = normalize_voice_strip(text)
    lower = (text or "").strip().lower()

    if (("invoice" in lower) or ("faktur" in lower)) and (state.get("step") == "idle"):
        inv_no = get_next_invoice_no()
        state["step"] = "inv_billto_name"
        state["data"] = {
            "invoice_no": inv_no,
            "invoice_date": datetime.now().strftime("%d-%b-%y"),
            "bill_to": {"name": "", "address": "", "address2": ""},
            "ship_to": {"name": "", "address": "", "address2": ""},
            "phone": "",
            "fax": "",
            "attn": "Accounting / Finance",
            "sales_person": "Syaeful Bakri",
            "ref_no": "",
            "ship_via": "",
            "ship_date": "",
            "terms": "",
            "no_surat_jalan": "",
            "items": [],
            "current_item": {},
            "freight": 0,
            "ppn_rate": 0.11,
            "deposit": 0,
            "payment": {
                "beneficiary": "PT. Sarana Trans Bersama Jaya",
                "bank_name": "BCA",
                "branch": "Cibadak - Sukabumi",
                "idr_acct": "35212 26666",
            }
        }
        conversations[sid] = state

        out_text = (
            "Baik, saya akan membantu membuat <b>Invoice</b>.<br><br>"
            f"Nomor invoice: <b>{inv_no}</b><br>"
            f"Tanggal: <b>{state['data']['invoice_date']}</b><br><br>"
            "Pertanyaan 1: <b>Nama perusahaan untuk Bill To?</b>"
        )

        history_id_created = None
        if not history_id_in:
            history_id_created = db_insert_history(
                title="Chat Baru",
                task_type=data.get("taskType") or "invoice",
                data={},
                files=[],
                messages=[
                    {"id": __import__("uuid").uuid4().hex[:12], "sender": "user", "text": text, "files": [], "timestamp": datetime.now().isoformat()},
                    {"id": __import__("uuid").uuid4().hex[:12], "sender": "assistant", "text": re.sub(r"<br\s*/?>", "\n", out_text), "files": [], "timestamp": datetime.now().isoformat()},
                ],
                state=state
            )
        else:
            db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
            db_update_state(int(history_id_in), state)

        return {"text": out_text, "history_id": history_id_created or history_id_in}

    if state.get("step") == "inv_billto_name":
        state["data"]["bill_to"]["name"] = text.strip()

        # ✅ FIX: jika alamat kosong/tidak ditemukan -> "Di tempat"
        alamat = resolve_company_address(text)
        alamat = alamat.strip() if isinstance(alamat, str) else ""
        if not alamat:
            alamat = "Di tempat"

        state["data"]["bill_to"]["address"] = alamat
        state["step"] = "inv_shipto_same"
        conversations[sid] = state

        out_text = (
            f"Bill To: <b>{state['data']['bill_to']['name']}</b><br>"
            f"Alamat: <b>{alamat}</b><br><br>"
            "Pertanyaan 2: <b>Apakah Ship To sama dengan Bill To?</b> (ya/tidak)"
        )

        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_shipto_same":
        if ("ya" in lower) or ("iya" in lower):
            state["data"]["ship_to"] = dict(state["data"]["bill_to"])
            state["step"] = "inv_phone"
            conversations[sid] = state
            out_text = "Pertanyaan 3: <b>Nomor telepon?</b> (boleh kosong; sebut <b>strip</b> jika tidak ada)"
        elif ("tidak" in lower) or ("gak" in lower) or ("nggak" in lower):
            state["step"] = "inv_shipto_name"
            conversations[sid] = state
            out_text = "Pertanyaan 2A: <b>Nama perusahaan untuk Ship To?</b>"
        else:
            out_text = (
                "Mohon jawab dengan <b>ya</b> atau <b>tidak</b>.<br><br>"
                "Pertanyaan 2: <b>Apakah Ship To sama dengan Bill To?</b>"
            )
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_shipto_name":
        state["data"]["ship_to"]["name"] = text.strip()

        # ✅ FIX: jika alamat kosong/tidak ditemukan -> "Di tempat"
        alamat = resolve_company_address(text)
        alamat = alamat.strip() if isinstance(alamat, str) else ""
        if not alamat:
            alamat = "Di tempat"

        state["data"]["ship_to"]["address"] = alamat
        state["step"] = "inv_phone"
        conversations[sid] = state
        out_text = "Pertanyaan 3: <b>Nomor telepon?</b> (boleh kosong; sebut <b>strip</b> jika tidak ada)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_phone":
        state["data"]["phone"] = "" if text.strip() in ("-", "") else text.strip()
        state["step"] = "inv_fax"
        conversations[sid] = state
        out_text = "Pertanyaan 4: <b>Fax?</b> (boleh kosong; sebut <b>strip</b> jika tidak ada)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_fax":
        state["data"]["fax"] = "" if text.strip() in ("-", "") else text.strip()
        state["step"] = "inv_attn"
        conversations[sid] = state
        out_text = "Pertanyaan 5: <b>Attn?</b> (default: Accounting / Finance; sebut <b>strip</b> untuk default)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_attn":
        if text.strip() not in ("-", ""):
            state["data"]["attn"] = text.strip()

        state["step"] = "inv_item_qty"
        state["data"]["current_item"] = {}
        conversations[sid] = state
        out_text = (
            "Item 1<br>"
            "Pertanyaan 6: <b>Qty?</b> (contoh: 749 atau 2,5 atau 'dua koma lima')"
        )
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_item_qty":
        qty = parse_qty_id(text)
        state["data"]["current_item"]["qty"] = qty
        state["data"]["current_item"]["unit"] = "Kg"
        state["data"]["current_item"]["date"] = state["data"]["invoice_date"]

        state["step"] = "inv_item_desc"
        conversations[sid] = state
        out_text = (
            "Pertanyaan 6B: <b>Jenis limbah atau kode limbah?</b><br>"
            "<i>Contoh: A102d atau aki baterai bekas. Atau sebut <b>NON B3</b>.</i>"
        )
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_item_desc":
        if is_non_b3_input(text):
            state["data"]["current_item"]["description"] = ""
            state["step"] = "inv_item_desc_manual"
            conversations[sid] = state
            out_text = "Pertanyaan 6C: <b>Deskripsi (manual) apa?</b>"
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

        kode, data_limbah = find_limbah_by_kode(text)
        if not (kode and data_limbah):
            kode, data_limbah = find_limbah_by_jenis(text)

        if kode and data_limbah:
            state["data"]["current_item"]["description"] = data_limbah["jenis"]
            state["step"] = "inv_item_price"
            conversations[sid] = state
            out_text = (
                f"Deskripsi: <b>{data_limbah['jenis']}</b><br><br>"
                "Pertanyaan 6D: <b>Harga (Rp)?</b>"
            )
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

        out_text = (
            f"Maaf, limbah <b>{text}</b> tidak ditemukan.<br><br>"
            "Silakan sebutkan kode/jenis lain atau sebut <b>NON B3</b> untuk input manual."
        )
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_item_desc_manual":
        state["data"]["current_item"]["description"] = text.strip()
        state["step"] = "inv_item_price"
        conversations[sid] = state
        out_text = "Pertanyaan 6D: <b>Harga (Rp)?</b>"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_item_price":
        price = parse_amount_id(text)
        if price is None:
            out_text = (
                "Input <b>harga</b> tidak valid.<br>"
                "Mohon masukkan angka. Contoh: <b>1500000</b> atau <b>1.5jt</b> atau <b>250rb</b> atau <b>nol</b> / <b>0</b>.<br><br>"
                "Pertanyaan 6D: <b>Harga (Rp)?</b>"
            )
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
            return {"text": out_text, "history_id": history_id_in}

        state["data"]["current_item"]["price"] = int(price)
        state["data"]["items"].append(state["data"]["current_item"])
        state["data"]["current_item"] = {}
        state["step"] = "inv_add_more_item"
        conversations[sid] = state
        out_text = "Pertanyaan: <b>Tambah item lagi?</b> (ya/tidak)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_add_more_item":
        if ("ya" in lower) or ("iya" in lower):
            num = len(state["data"]["items"])
            state["step"] = "inv_item_qty"
            state["data"]["current_item"] = {}
            conversations[sid] = state
            out_text = f"Item {num+1}<br>Pertanyaan 6: <b>Qty?</b>"
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

        if ("tidak" in lower) or ("gak" in lower) or ("nggak" in lower) or ("skip" in lower) or ("lewat" in lower):
            state["step"] = "inv_freight"
            conversations[sid] = state
            out_text = "Pertanyaan 7: <b>Biaya transportasi/Freight (Rp)?</b> (isi 0 atau nol jika tidak ada)"
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

        out_text = "Mohon jawab <b>ya</b> atau <b>tidak</b>."
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_freight":
        freight_val = parse_amount_id(text)
        if freight_val is None:
            out_text = (
                "Input <b>freight</b> tidak valid.<br>"
                "Mohon masukkan angka. Contoh: <b>0</b> / <b>nol</b> / <b>1.5jt</b> jika ada.<br><br>"
                "Pertanyaan 7: <b>Biaya transportasi/Freight (Rp)?</b>"
            )
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
            return {"text": out_text, "history_id": history_id_in}

        state["data"]["freight"] = int(freight_val)
        state["step"] = "inv_deposit"
        conversations[sid] = state
        out_text = "Pertanyaan 8: <b>Deposit (Rp)?</b> (isi 0 atau nol jika tidak ada)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_deposit":
        dep_val = parse_amount_id(text)
        if dep_val is None:
            out_text = (
                "Input <b>deposit</b> tidak valid.<br>"
                "Mohon masukkan angka. Contoh: <b>0</b> / <b>nol</b> / <b>250rb</b> jika ada.<br><br>"
                "Pertanyaan 8: <b>Deposit (Rp)?</b>"
            )
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=[])
            return {"text": out_text, "history_id": history_id_in}

        state["data"]["deposit"] = int(dep_val)

        nama_pt_raw = (state["data"].get("bill_to") or {}).get("name", "").strip()
        safe_pt = re.sub(r"[^A-Za-z0-9 \-]+", "", nama_pt_raw).strip()
        safe_pt = re.sub(r"\s+", " ", safe_pt).strip()
        base_fname = f"Invoice - {safe_pt}" if safe_pt else "Invoice"
        fname_base = make_unique_filename_base(base_fname)

        xlsx = create_invoice_xlsx(state["data"], fname_base)
        pdf_preview = create_invoice_pdf(state["data"], fname_base)

        files = [
            {"type": "xlsx", "filename": xlsx, "url": f"/download/{xlsx}"},
            {"type": "pdf", "filename": pdf_preview, "url": f"/download/{pdf_preview}"},
        ]

        conversations[sid] = {"step": "idle", "data": {}}

        history_title = f"Invoice {nama_pt_raw}" if nama_pt_raw else "Invoice"
        history_task_type = "invoice"

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
            "<b>Invoice berhasil dibuat.</b><br><br>"
            f"Nomor invoice: <b>{state['data'].get('invoice_no')}</b><br>"
            f"Bill To: <b>{(state['data'].get('bill_to') or {}).get('name','')}</b><br>"
            f"Jumlah item: <b>{len(state['data'].get('items') or [])}</b><br><br>"
            "Dokumen tersedia dalam format PDF (preview) dan Excel (.xlsx)."
        )

        db_append_message(history_id, "assistant", re.sub(r"<br\s*/?>", "\n", out_text), files=files)
        return {"text": out_text, "files": files, "history_id": history_id}

    return None
