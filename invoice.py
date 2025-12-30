import os
import json
import re
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

# ✅ NEW: untuk ambil image1.jpeg dari docx template
import zipfile
from io import BytesIO

from limbah_database import (
    find_limbah_by_kode,
    find_limbah_by_jenis,
    convert_voice_to_number,
)
from utils import (
    db_insert_history, db_append_message, db_update_state,
    search_company_address, search_company_address_ai,
)
from config_new import FILES_DIR


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

def parse_qty_id(text: str) -> float:
    if not text:
        return 0.0
    t = normalize_id_number_text(text)
    v = convert_voice_to_number(t)
    try:
        return float(v)
    except:
        m = re.findall(r'\d+(?:\.\d+)?', t)
        return float(m[0]) if m else 0.0

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
            os.path.exists(os.path.join(folder, f"{name}.xlsx")) or
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
# Counter Invoice (as-is)
# =========================

def _invoice_counter_path() -> str:
    try:
        folder = str(FILES_DIR)
    except Exception:
        folder = "static/files"
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "invoice_counter.json")

def load_invoice_counter() -> int:
    path = _invoice_counter_path()
    try:
        if not os.path.exists(path):
            return 0
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f) or {}
        return int(data.get("counter", 0))
    except:
        return 0

def save_invoice_counter(n: int) -> None:
    path = _invoice_counter_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"counter": int(n)}, f)

def get_next_invoice_no() -> str:
    now = datetime.now()
    prefix = now.strftime("%y%m")
    n = load_invoice_counter() + 1
    save_invoice_counter(n)
    return f"{prefix}{str(n).zfill(3)}"


# =========================
# Generate Invoice XLSX (AS-IS, jangan diubah)
# =========================

def _thin_border():
    side = Side(style="thin", color="000000")
    return Border(left=side, right=side, top=side, bottom=side)

def _set_border(ws, r1, c1, r2, c2, border):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = border

def create_invoice_xlsx(inv: dict, fname_base: str) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins.left = 0.35
    ws.page_margins.right = 0.35
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35

    col_widths = {"A": 8, "B": 6, "C": 12, "D": 45, "E": 14, "F": 16}
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    border = _thin_border()
    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="top", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    def money(cell):
        cell.number_format = '#,##0'

    payment = inv.get("payment") or {}
    defaults = {
        "beneficiary": "PT. Sarana Trans Bersama Jaya",
        "bank_name": "BCA",
        "branch": "Cibadak - Sukabumi",
        "idr_acct": "35212 26666",
    }
    for k, v in defaults.items():
        if not payment.get(k):
            payment[k] = v

    invoice_no = inv.get("invoice_no") or get_next_invoice_no()
    inv_date = inv.get("invoice_date") or datetime.now().strftime("%d-%b-%y")

    bill_to = inv.get("bill_to") or {}
    ship_to = inv.get("ship_to") or {}
    attn = inv.get("attn") or "Accounting / Finance"
    phone = inv.get("phone") or ""
    fax = inv.get("fax") or ""

    sales_person = inv.get("sales_person") or "Syaeful Bakri"
    ref_no = inv.get("ref_no") or ""
    ship_via = inv.get("ship_via") or ""
    ship_date = inv.get("ship_date") or ""
    terms = inv.get("terms") or ""
    no_surat_jalan = inv.get("no_surat_jalan") or ""

    ws["A1"].value = "Bill To:"
    ws["A1"].font = bold
    ws.merge_cells("A1:C1")

    ws["D1"].value = "Ship To:"
    ws["D1"].font = bold
    ws.merge_cells("D1:F1")

    bill_lines = [(bill_to.get("name") or "").strip(), (bill_to.get("address") or "").strip(), (bill_to.get("address2") or "").strip()]
    ship_lines = [(ship_to.get("name") or "").strip(), (ship_to.get("address") or "").strip(), (ship_to.get("address2") or "").strip()]
    bill_text = "\n".join([x for x in bill_lines if x])
    ship_text = "\n".join([x for x in ship_lines if x])

    ws["A2"].value = bill_text
    ws.merge_cells("A2:C3")
    ws["A2"].alignment = left

    ws["D2"].value = ship_text
    ws.merge_cells("D2:F3")
    ws["D2"].alignment = left

    ws["A4"].value = "Phone:"
    ws["A4"].font = bold
    ws.merge_cells("A4:B4")
    ws["C4"].value = phone
    ws["C4"].alignment = left

    ws["D4"].value = "Fax:"
    ws["D4"].font = bold
    ws.merge_cells("D4:E4")
    ws["F4"].value = fax
    ws["F4"].alignment = left

    ws["A5"].value = "Attn :"
    ws["A5"].font = bold
    ws.merge_cells("A5:B5")
    ws["C5"].value = attn
    ws.merge_cells("C5:F5")
    ws["C5"].alignment = left

    ws["E6"].value = "Invoice"
    ws["E6"].font = bold
    ws["E6"].alignment = center
    ws["F6"].value = invoice_no
    ws["F6"].alignment = center

    ws["E7"].value = "Date"
    ws["E7"].font = bold
    ws["E7"].alignment = center
    ws["F7"].value = inv_date
    ws["F7"].alignment = center

    ws["E8"].value = "No. Surat Jalan"
    ws["E8"].font = bold
    ws["E8"].alignment = center
    ws["F8"].value = no_surat_jalan
    ws["F8"].alignment = center

    ws.merge_cells("A9:B9")
    ws["A9"].value = "Ref No."
    ws["A9"].font = bold
    ws["A9"].alignment = center

    ws.merge_cells("C9:D9")
    ws["C9"].value = "Sales Person"
    ws["C9"].font = bold
    ws["C9"].alignment = center

    ws["E9"].value = "Ship Via"
    ws["E9"].font = bold
    ws["E9"].alignment = center

    ws["F9"].value = "Ship Date"
    ws["F9"].font = bold
    ws["F9"].alignment = center

    ws.merge_cells("A10:B10")
    ws["A10"].value = ref_no
    ws["A10"].alignment = center

    ws.merge_cells("C10:D10")
    ws["C10"].value = sales_person
    ws["C10"].alignment = center

    ws["E10"].value = ship_via
    ws["E10"].alignment = center

    ws["F10"].value = ship_date
    ws["F10"].alignment = center

    ws.merge_cells("E11:F11")
    ws["E11"].value = "Terms"
    ws["E11"].font = bold
    ws["E11"].alignment = center

    ws.merge_cells("E12:F12")
    ws["E12"].value = terms
    ws["E12"].alignment = center

    _set_border(ws, 1, 1, 12, 6, border)

    ws["A14"].value = "Qty"
    ws["B14"].value = ""
    ws["C14"].value = "Date"
    ws["D14"].value = "Description"
    ws["E14"].value = "Price"
    ws["F14"].value = "Amount (IDR)"
    for c in "ABCDEF":
        ws[f"{c}14"].font = bold
        ws[f"{c}14"].alignment = center

    items = inv.get("items") or []
    r = 15
    subtotal = 0
    for it in items:
        qty = float(it.get("qty") or 0)
        unit = (it.get("unit") or "").strip()
        desc = (it.get("description") or "").strip()
        price = int(it.get("price") or 0)
        line_date = it.get("date") or inv_date
        amount = int(round(qty * price))
        subtotal += amount

        ws[f"A{r}"].value = qty if qty % 1 != 0 else int(qty)
        ws[f"A{r}"].alignment = center
        ws[f"B{r}"].value = unit
        ws[f"B{r}"].alignment = center
        ws[f"C{r}"].value = line_date
        ws[f"C{r}"].alignment = center
        ws[f"D{r}"].value = desc
        ws[f"D{r}"].alignment = left
        ws[f"E{r}"].value = price
        ws[f"E{r}"].alignment = right
        money(ws[f"E{r}"])
        ws[f"F{r}"].value = amount
        ws[f"F{r}"].alignment = right
        money(ws[f"F{r}"])
        r += 1

    min_last_row = 26
    if r < min_last_row:
        r = min_last_row

    _set_border(ws, 14, 1, r - 1, 6, border)

    freight = int(inv.get("freight") or 0)
    ppn_rate = float(inv.get("ppn_rate") or 0.11)
    deposit = int(inv.get("deposit") or 0)

    total_before_ppn = subtotal + freight
    ppn = int(round(total_before_ppn * ppn_rate))
    balance = total_before_ppn + ppn - deposit

    sum_row = r
    ws[f"E{sum_row}"].value = "Total"
    ws[f"E{sum_row}"].alignment = right
    ws[f"E{sum_row}"].font = bold
    ws[f"F{sum_row}"].value = subtotal
    ws[f"F{sum_row}"].alignment = right
    money(ws[f"F{sum_row}"])

    ws[f"E{sum_row+1}"].value = "Freight"
    ws[f"E{sum_row+1}"].alignment = right
    ws[f"F{sum_row+1}"].value = freight
    ws[f"F{sum_row+1}"].alignment = right
    money(ws[f"F{sum_row+1}"])

    ws[f"E{sum_row+2}"].value = "Total"
    ws[f"E{sum_row+2}"].alignment = right
    ws[f"E{sum_row+2}"].font = bold
    ws[f"F{sum_row+2}"].value = total_before_ppn
    ws[f"F{sum_row+2}"].alignment = right
    money(ws[f"F{sum_row+2}"])

    ws[f"E{sum_row+3}"].value = f"PPN {int(ppn_rate*100)}%"
    ws[f"E{sum_row+3}"].alignment = right
    ws[f"F{sum_row+3}"].value = ppn
    ws[f"F{sum_row+3}"].alignment = right
    money(ws[f"F{sum_row+3}"])

    ws[f"E{sum_row+4}"].value = "Less: Deposit"
    ws[f"E{sum_row+4}"].alignment = right
    ws[f"F{sum_row+4}"].value = deposit
    ws[f"F{sum_row+4}"].alignment = right
    money(ws[f"F{sum_row+4}"])

    ws[f"E{sum_row+5}"].value = "Balance Due"
    ws[f"E{sum_row+5}"].alignment = right
    ws[f"E{sum_row+5}"].font = bold
    ws[f"F{sum_row+5}"].value = balance
    ws[f"F{sum_row+5}"].alignment = right
    ws[f"F{sum_row+5}"].font = bold
    money(ws[f"F{sum_row+5}"])

    _set_border(ws, sum_row, 5, sum_row + 5, 6, border)

    pay_row = sum_row
    ws.merge_cells(f"A{pay_row}:D{pay_row}")
    ws[f"A{pay_row}"].value = "Please Transfer Full Amount to:"
    ws[f"A{pay_row}"].font = bold
    ws[f"A{pay_row}"].alignment = left

    ws[f"A{pay_row+1}"].value = "Beneficiary :"
    ws.merge_cells(f"B{pay_row+1}:D{pay_row+1}")
    ws[f"B{pay_row+1}"].value = payment["beneficiary"]

    ws[f"A{pay_row+2}"].value = "Bank Name :"
    ws.merge_cells(f"B{pay_row+2}:D{pay_row+2}")
    ws[f"B{pay_row+2}"].value = payment["bank_name"]

    ws[f"A{pay_row+3}"].value = "Branch :"
    ws.merge_cells(f"B{pay_row+3}:D{pay_row+3}")
    ws[f"B{pay_row+3}"].value = payment["branch"]

    ws[f"A{pay_row+4}"].value = "IDR Acct :"
    ws.merge_cells(f"B{pay_row+4}:D{pay_row+4}")
    ws[f"B{pay_row+4}"].value = payment["idr_acct"]

    _set_border(ws, pay_row, 1, pay_row + 4, 4, border)

    try:
        folder = str(FILES_DIR)
    except Exception:
        folder = "static/files"
    os.makedirs(folder, exist_ok=True)

    out_path = os.path.join(folder, f"{fname_base}.xlsx")
    wb.save(out_path)
    return f"{fname_base}.xlsx"


# =========================
# ✅ NEW: PDF pakai template DOCX (background image) + overlay text
# =========================

def _extract_invoice_bg_image_bytes(template_docx_path: str) -> bytes:
    """
    Template kamu berisi 1 gambar: word/media/image1.jpeg
    Kita extract bytes-nya lalu dipakai sebagai background PDF.
    """
    if not os.path.exists(template_docx_path):
        raise Exception(f"Template invoice tidak ditemukan: {template_docx_path}")

    with zipfile.ZipFile(template_docx_path, "r") as z:
        # prioritas image1.jpeg, fallback cari media image pertama
        target = None
        for name in z.namelist():
            if name.lower() == "word/media/image1.jpeg":
                target = name
                break
        if not target:
            # fallback: cari file gambar pertama
            for name in z.namelist():
                if name.lower().startswith("word/media/") and name.lower().endswith((".jpg", ".jpeg", ".png")):
                    target = name
                    break
        if not target:
            raise Exception("Tidak menemukan gambar background di template invoice.docx (word/media/*).")

        return z.read(target)

def _rupiah(n: int) -> str:
    try:
        n = int(n)
    except:
        n = 0
    return f"{n:,}".replace(",", ".")

def create_invoice_pdf_from_template(inv: dict, fname_base: str) -> str:
    """
    Output PDF A4:
    - background: image dari tamplate invoice.docx
    - overlay: field invoice
    """
    try:
        folder = str(FILES_DIR)
    except Exception:
        folder = "static/files"
    os.makedirs(folder, exist_ok=True)

    out_path = os.path.join(folder, f"{fname_base}.pdf")

    template_path = "tamplate invoice.docx"  # ✅ sesuai permintaan: ada di root
    bg_bytes = _extract_invoice_bg_image_bytes(template_path)
    bg_img = ImageReader(BytesIO(bg_bytes))

    c = canvas.Canvas(out_path, pagesize=A4)
    W, H = A4

    # draw background full page
    c.drawImage(bg_img, 0, 0, width=W, height=H, preserveAspectRatio=True, mask='auto')

    # --- ambil data invoice ---
    invoice_no = inv.get("invoice_no") or ""
    inv_date = inv.get("invoice_date") or ""
    bill_to = inv.get("bill_to") or {}
    ship_to = inv.get("ship_to") or {}
    phone = inv.get("phone") or ""
    fax = inv.get("fax") or ""
    attn = inv.get("attn") or ""
    sales_person = inv.get("sales_person") or "Syaeful Bakri"
    ref_no = inv.get("ref_no") or ""
    ship_via = inv.get("ship_via") or ""
    ship_date = inv.get("ship_date") or ""
    terms = inv.get("terms") or ""
    no_surat_jalan = inv.get("no_surat_jalan") or ""
    items = inv.get("items") or []
    freight = int(inv.get("freight") or 0)
    ppn_rate = float(inv.get("ppn_rate") or 0.11)
    deposit = int(inv.get("deposit") or 0)
    payment = inv.get("payment") or {}

    # hitung subtotal
    subtotal = 0
    for it in items:
        qty = float(it.get("qty") or 0)
        price = int(it.get("price") or 0)
        subtotal += int(round(qty * price))

    total_before_ppn = subtotal + freight
    ppn = int(round(total_before_ppn * ppn_rate))
    balance = total_before_ppn + ppn - deposit

    # =========================
    # ✅ POSISI TEXT OVERLAY (KALAU MAU NGE-PASIN, EDIT ANGKA DI SINI)
    # Koordinat: (x, y) dalam POINT (origin kiri bawah)
    # =========================
    TPL = {
        # Bill To block
        "bill_x": 55,
        "bill_y": 655,   # baris pertama
        "bill_lh": 13,   # line height

        # Ship To block
        "ship_x": 330,
        "ship_y": 655,
        "ship_lh": 13,

        # Phone/Fax/Attn
        "phone_x": 88,
        "phone_y": 556,
        "fax_x": 325,
        "fax_y": 556,
        "attn_x": 115,
        "attn_y": 525,

        # kanan atas: invoice header
        "invno_x": 505,
        "invno_y": 568,
        "date_x": 505,
        "date_y": 552,
        "sj_x": 505,
        "sj_y": 536,

        # baris Ref / Sales / ShipVia / ShipDate / Terms
        "ref_x": 85,
        "ref_y": 482,
        "sales_x": 210,
        "sales_y": 482,
        "shipvia_x": 360,
        "shipvia_y": 482,
        "shipdate_x": 468,
        "shipdate_y": 482,
        "terms_x": 548,
        "terms_y": 482,

        # items table start
        "row_y": 430,
        "row_lh": 16,
        "qty_x": 63,
        "unit_x": 110,
        "dateitem_x": 135,
        "desc_x": 205,
        "price_x": 480,
        "amt_x": 555,

        "max_rows": 10,

        # totals kanan bawah
        "subtotal_x": 555,
        "subtotal_y": 286,
        "freight_x": 555,
        "freight_y": 270,
        "total2_x": 555,
        "total2_y": 254,
        "ppn_x": 555,
        "ppn_y": 238,
        "deposit_x": 555,
        "deposit_y": 222,
        "balance_x": 555,
        "balance_y": 206,

        # payment kiri bawah
        "pay_x": 115,
        "pay_y": 286,
        "pay_lh": 13,
    }

    def draw(x, y, txt, size=9, bold=False, align_right=False):
        c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
        if txt is None:
            txt = ""
        s = str(txt)
        if align_right:
            c.drawRightString(x, y, s)
        else:
            c.drawString(x, y, s)

    # ---- Bill To / Ship To (multi-line) ----
    bill_lines = [
        (bill_to.get("name") or "").strip(),
        (bill_to.get("address") or "").strip(),
        (bill_to.get("address2") or "").strip(),
    ]
    bill_lines = [x for x in bill_lines if x]
    yy = TPL["bill_y"]
    for line in bill_lines[:4]:
        draw(TPL["bill_x"], yy, line, size=9, bold=False)
        yy -= TPL["bill_lh"]

    ship_lines = [
        (ship_to.get("name") or "").strip(),
        (ship_to.get("address") or "").strip(),
        (ship_to.get("address2") or "").strip(),
    ]
    ship_lines = [x for x in ship_lines if x]
    yy2 = TPL["ship_y"]
    for line in ship_lines[:4]:
        draw(TPL["ship_x"], yy2, line, size=9, bold=False)
        yy2 -= TPL["ship_lh"]

    # phone / fax / attn
    draw(TPL["phone_x"], TPL["phone_y"], phone, size=9)
    draw(TPL["fax_x"], TPL["fax_y"], fax, size=9)
    draw(TPL["attn_x"], TPL["attn_y"], attn, size=9)

    # kanan atas
    draw(TPL["invno_x"], TPL["invno_y"], invoice_no, size=9, bold=True)
    draw(TPL["date_x"], TPL["date_y"], inv_date, size=9, bold=True)
    draw(TPL["sj_x"], TPL["sj_y"], no_surat_jalan, size=9)

    # header row kecil
    draw(TPL["ref_x"], TPL["ref_y"], ref_no, size=9)
    draw(TPL["sales_x"], TPL["sales_y"], sales_person, size=9)
    draw(TPL["shipvia_x"], TPL["shipvia_y"], ship_via, size=9)
    draw(TPL["shipdate_x"], TPL["shipdate_y"], ship_date, size=9)
    draw(TPL["terms_x"], TPL["terms_y"], terms, size=9, align_right=True)

    # items (maks 10 row)
    yrow = TPL["row_y"]
    for idx in range(TPL["max_rows"]):
        if idx < len(items):
            it = items[idx]
            qty = it.get("qty") or 0
            unit = (it.get("unit") or "").strip()
            dt = it.get("date") or inv_date
            desc = (it.get("description") or "").strip()
            price = int(it.get("price") or 0)
            amount = int(round(float(qty) * price))

            # qty tampil rapih (tanpa .0)
            try:
                qv = float(qty)
                qty_txt = str(int(qv)) if qv.is_integer() else str(qv)
            except:
                qty_txt = str(qty)

            draw(TPL["qty_x"], yrow, qty_txt, size=9, align_right=True)
            draw(TPL["unit_x"], yrow, unit, size=9)
            draw(TPL["dateitem_x"], yrow, dt, size=9)
            draw(TPL["desc_x"], yrow, desc[:55], size=9)
            draw(TPL["price_x"], yrow, _rupiah(price), size=9, align_right=True)
            draw(TPL["amt_x"], yrow, _rupiah(amount), size=9, align_right=True)

        yrow -= TPL["row_lh"]

    # totals kanan
    draw(TPL["subtotal_x"], TPL["subtotal_y"], _rupiah(subtotal), size=9, align_right=True)
    draw(TPL["freight_x"], TPL["freight_y"], _rupiah(freight) if freight else "-", size=9, align_right=True)
    draw(TPL["total2_x"], TPL["total2_y"], _rupiah(total_before_ppn), size=9, align_right=True)
    draw(TPL["ppn_x"], TPL["ppn_y"], _rupiah(ppn), size=9, align_right=True)
    draw(TPL["deposit_x"], TPL["deposit_y"], _rupiah(deposit) if deposit else "-", size=9, align_right=True)
    draw(TPL["balance_x"], TPL["balance_y"], _rupiah(balance), size=10, bold=True, align_right=True)

    # payment kiri bawah (ambil dari inv.payment kalau ada)
    beneficiary = payment.get("beneficiary", "PT. Sarana Trans Bersama Jaya")
    bank_name = payment.get("bank_name", "BCA")
    branch = payment.get("branch", "Cibadak - Sukabumi")
    idr_acct = payment.get("idr_acct", "35212 26666")

    pay_lines = [
        f"{beneficiary}",
        f"{bank_name}",
        f"{branch}",
        f"{idr_acct}",
    ]
    py = TPL["pay_y"]
    for line in pay_lines:
        draw(TPL["pay_x"], py, line, size=9)
        py -= TPL["pay_lh"]

    c.showPage()
    c.save()
    return f"{fname_base}.pdf"


# =========================
# CHAT HANDLER INVOICE (FLOW AS-IS, hanya output PDF diganti template)
# =========================

def handle_invoice_flow(data: dict, text: str, lower: str, sid: str, state: dict, conversations: dict, history_id_in):
    """
    Return:
      - None jika bukan flow invoice
      - dict response jika handled
    """

    # trigger invoice (sama)
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
            "Baik, saya bantu buatkan <b>INVOICE</b>.<br><br>"
            f"✅ Invoice No: <b>{inv_no}</b><br>"
            f"✅ Date: <b>{state['data']['invoice_date']}</b><br><br>"
            "❓ <b>1. Bill To - Nama Perusahaan?</b>"
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
                    {"id": __import__("uuid").uuid4().hex[:12], "sender": "assistant", "text": re.sub(r'<br\s*/?>', '\n', out_text), "files": [], "timestamp": datetime.now().isoformat()},
                ],
                state=state
            )
        else:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)

        return {"text": out_text, "history_id": history_id_created or history_id_in}

    # step-step invoice (sama)
    if state.get("step") == "inv_billto_name":
        state["data"]["bill_to"]["name"] = text.strip()

        alamat = search_company_address(text).strip()
        if not alamat:
            alamat = search_company_address_ai(text).strip()
        if not alamat:
            alamat = "Di Tempat"

        state["data"]["bill_to"]["address"] = alamat
        state["step"] = "inv_shipto_same"
        conversations[sid] = state

        out_text = (
            f"✅ Bill To: <b>{state['data']['bill_to']['name']}</b><br>"
            f"✅ Alamat: <b>{alamat}</b><br><br>"
            "❓ <b>2. Ship To sama dengan Bill To?</b> (ya/tidak)"
        )

        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_shipto_same":
        if ("ya" in lower) or ("iya" in lower):
            state["data"]["ship_to"] = dict(state["data"]["bill_to"])
            state["step"] = "inv_phone"
            conversations[sid] = state
            out_text = "❓ <b>3. Phone?</b> (boleh kosong, ketik '-' jika tidak ada)"
        elif ("tidak" in lower) or ("gak" in lower) or ("nggak" in lower):
            state["step"] = "inv_shipto_name"
            conversations[sid] = state
            out_text = "❓ <b>2A. Ship To - Nama Perusahaan?</b>"
        else:
            out_text = "⚠️ Mohon jawab dengan <b>'ya'</b> atau <b>'tidak'</b><br><br>❓ <b>2. Ship To sama dengan Bill To?</b>"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_shipto_name":
        state["data"]["ship_to"]["name"] = text.strip()

        alamat = search_company_address(text).strip()
        if not alamat:
            alamat = search_company_address_ai(text).strip()
        if not alamat:
            alamat = "Di Tempat"
        state["data"]["ship_to"]["address"] = alamat

        state["step"] = "inv_phone"
        conversations[sid] = state
        out_text = "❓ <b>3. Phone?</b> (boleh kosong, ketik '-' jika tidak ada)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_phone":
        state["data"]["phone"] = "" if text.strip() in ("-", "") else text.strip()
        state["step"] = "inv_fax"
        conversations[sid] = state
        out_text = "❓ <b>4. Fax?</b> (boleh kosong, ketik '-' jika tidak ada)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_fax":
        state["data"]["fax"] = "" if text.strip() in ("-", "") else text.strip()
        state["step"] = "inv_attn"
        conversations[sid] = state
        out_text = "❓ <b>5. Attn?</b> (default: Accounting / Finance | ketik '-' untuk default)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_attn":
        if text.strip() not in ("-", ""):
            state["data"]["attn"] = text.strip()
        state["step"] = "inv_item_qty"
        state["data"]["current_item"] = {}
        conversations[sid] = state
        out_text = "📦 <b>Item #1</b><br>❓ <b>6. Qty?</b> (contoh: 749 atau 3,5)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_item_qty":
        qty = parse_qty_id(text)
        state["data"]["current_item"]["qty"] = qty
        state["step"] = "inv_item_unit"
        conversations[sid] = state
        out_text = "❓ <b>6A. Unit?</b> (contoh: Kg / Liter / Pcs)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_item_unit":
        state["data"]["current_item"]["unit"] = text.strip()
        state["data"]["current_item"]["date"] = state["data"]["invoice_date"]
        state["step"] = "inv_item_desc"
        conversations[sid] = state
        out_text = "❓ <b>6B. Jenis Limbah / Kode Limbah?</b><br><i>(Contoh: 'A102d' atau 'aki baterai bekas' | atau ketik <b>NON B3</b>)</i>"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_item_desc":
        if is_non_b3_input(text):
            state["data"]["current_item"]["description"] = ""
            state["step"] = "inv_item_desc_manual"
            conversations[sid] = state
            out_text = "❓ <b>6C. Deskripsi (manual) apa?</b>"
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

        kode, data_limbah = find_limbah_by_kode(text)
        if not (kode and data_limbah):
            kode, data_limbah = find_limbah_by_jenis(text)

        if kode and data_limbah:
            state["data"]["current_item"]["description"] = data_limbah["jenis"]
            state["step"] = "inv_item_price"
            conversations[sid] = state
            out_text = f"✅ Deskripsi: <b>{data_limbah['jenis']}</b><br><br>❓ <b>6D. Price (Rp)?</b>"
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

        out_text = f"❌ Maaf, limbah '<b>{text}</b>' tidak ditemukan.<br><br>Ketik kode/jenis lain atau <b>NON B3</b>."
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_item_desc_manual":
        state["data"]["current_item"]["description"] = text.strip()
        state["step"] = "inv_item_price"
        conversations[sid] = state
        out_text = "❓ <b>6D. Price (Rp)?</b>"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_item_price":
        price = parse_amount_id(text)
        state["data"]["current_item"]["price"] = price
        state["data"]["items"].append(state["data"]["current_item"])
        state["data"]["current_item"] = {}
        state["step"] = "inv_add_more_item"
        conversations[sid] = state
        out_text = "❓ <b>Tambah item lagi?</b> (ya/tidak)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_add_more_item":
        if ("ya" in lower) or ("iya" in lower):
            num = len(state["data"]["items"])
            state["step"] = "inv_item_qty"
            state["data"]["current_item"] = {}
            conversations[sid] = state
            out_text = f"📦 <b>Item #{num+1}</b><br>❓ <b>6. Qty?</b>"
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

        if ("tidak" in lower) or ("gak" in lower) or ("nggak" in lower) or ("skip" in lower) or ("lewat" in lower):
            state["step"] = "inv_freight"
            conversations[sid] = state
            out_text = "❓ <b>7. Biaya Transportasi/Freight (Rp)?</b> (0 jika tidak ada)"
            if history_id_in:
                db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
                db_update_state(int(history_id_in), state)
            return {"text": out_text, "history_id": history_id_in}

        out_text = "⚠️ Mohon jawab <b>ya</b> atau <b>tidak</b>."
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_freight":
        state["data"]["freight"] = parse_amount_id(text)
        state["step"] = "inv_deposit"
        conversations[sid] = state
        out_text = "❓ <b>8. Deposit (Rp)?</b> (0 jika tidak ada)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_deposit":
        state["data"]["deposit"] = parse_amount_id(text)

        nama_pt_raw = (state["data"].get("bill_to") or {}).get("name", "").strip()
        safe_pt = re.sub(r'[^A-Za-z0-9 \-]+', '', nama_pt_raw).strip()
        safe_pt = re.sub(r'\s+', ' ', safe_pt).strip()
        base_fname = f"Invoice - {safe_pt}" if safe_pt else "Invoice"
        fname_base = make_unique_filename_base(base_fname)

        # ✅ Excel tetap seperti sebelumnya
        xlsx = create_invoice_xlsx(state["data"], fname_base)

        # ✅ PDF sekarang pakai template invoice.docx (background)
        pdf_preview = create_invoice_pdf_from_template(state["data"], fname_base)

        files = [
            {"type": "pdf", "filename": pdf_preview, "url": f"/download/{pdf_preview}"},
            {"type": "xlsx", "filename": xlsx, "url": f"/download/{xlsx}"},
        ]

        conversations[sid] = {'step': 'idle', 'data': {}}

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
            "🎉 <b>Invoice berhasil dibuat!</b><br><br>"
            f"✅ Invoice No: <b>{state['data'].get('invoice_no')}</b><br>"
            f"✅ Bill To: <b>{(state['data'].get('bill_to') or {}).get('name','')}</b><br>"
            f"✅ Total Item: <b>{len(state['data'].get('items') or [])}</b><br><br>"
            "📌 Preview/Download: PDF (template)<br>"
            "📌 Download: Excel (.xlsx) (format existing)"
        )

        db_append_message(history_id, "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=files)
        return {"text": out_text, "files": files, "history_id": history_id}

    return None
