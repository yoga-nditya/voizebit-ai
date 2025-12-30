import os
import json
import re
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

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


def normalize_voice_strip(text: str) -> str:
    if not text:
        return text
    t = text.strip()
    low = t.lower()
    if low == "strip":
        return "-"
    if re.fullmatch(r"(.*\b)?strip(\b.*)?", low):
        return "-"
    return text


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


def _invoice_counter_path() -> str:
    try:
        folder = str(FILES_DIR)
    except Exception:
        folder = "static/files"
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "invoice_counter.json")


def load_invoice_counter(prefix: str) -> int:
    path = _invoice_counter_path()
    try:
        if not os.path.exists(path):
            return 0
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f) or {}
        counters = data.get("counters") or {}
        return int(counters.get(prefix, 0))
    except:
        return 0


def save_invoice_counter(prefix: str, n: int) -> None:
    path = _invoice_counter_path()
    data = {}
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f) or {}
    except:
        data = {}

    counters = data.get("counters") or {}
    counters[prefix] = int(n)
    data["counters"] = counters

    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)


def get_next_invoice_no() -> str:
    now = datetime.now()
    prefix = now.strftime("%d%m")
    n = load_invoice_counter(prefix) + 1
    save_invoice_counter(prefix, n)
    return f"{prefix}{str(n).zfill(3)}"


def _side(style="thin"):
    return Side(style=style, color="000000")


def _clear_borders(ws, r1, c1, r2, c2):
    empty = Border()
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = empty


def _set_outer_border(ws, r1, c1, r2, c2, style="medium"):
    s = _side(style)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            left = s if c == c1 else None
            right = s if c == c2 else None
            top = s if r == r1 else None
            bottom = s if r == r2 else None
            cell.border = Border(left=left, right=right, top=top, bottom=bottom)


# ==========================================
# ✅ UPDATED LAYOUT: match seperti gambar #2
# (hanya fungsi ini yang dirombak)
# ==========================================
def create_invoice_xlsx(inv: dict, fname_base: str) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    # Page setup (biar rapih di A4)
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = 0.35
    ws.page_margins.right = 0.35
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35

    # Matikan gridlines biar clean seperti template
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110

    # Pakai A..I (lebih rapat, tidak “melebar ke kanan”)
    col_widths = {
        "A": 7,   # Qty
        "B": 6,   # Unit
        "C": 12,  # Date
        "D": 18,  # Desc part (merge D-G)
        "E": 18,
        "F": 18,
        "G": 18,
        "H": 14,  # Price
        "I": 18,  # Amount
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="top", wrap_text=True)
    left_mid = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    def money(cell):
        cell.number_format = '#,##0'

    # Helper border table: outer medium + inner thin
    def set_table_borders(r1, c1, r2, c2, outer_style="medium", inner_style="thin"):
        outer = _side(outer_style)
        inner = _side(inner_style)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                left_b = outer if c == c1 else inner
                right_b = outer if c == c2 else inner
                top_b = outer if r == r1 else inner
                bot_b = outer if r == r2 else inner
                ws.cell(r, c).border = Border(left=left_b, right=right_b, top=top_b, bottom=bot_b)

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

    # ======================
    # Header Bill/Ship
    # ======================
    ws["A1"].value = "Bill To:"
    ws["A1"].font = bold
    ws.merge_cells("A1:D1")

    ws["F1"].value = "Ship To:"
    ws["F1"].font = bold
    ws.merge_cells("F1:I1")

    bill_lines = [
        (bill_to.get("name") or "").strip(),
        (bill_to.get("address") or "").strip(),
        (bill_to.get("address2") or "").strip()
    ]
    ship_lines = [
        (ship_to.get("name") or "").strip(),
        (ship_to.get("address") or "").strip(),
        (ship_to.get("address2") or "").strip()
    ]
    bill_text = "\n".join([x for x in bill_lines if x])
    ship_text = "\n".join([x for x in ship_lines if x])

    ws["A2"].value = bill_text
    ws.merge_cells("A2:D3")
    ws["A2"].alignment = left

    ws["F2"].value = ship_text
    ws.merge_cells("F2:I3")
    ws["F2"].alignment = left

    # Row height biar rapih
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 30

    # Phone / Fax
    ws["A5"].value = "Phone:"
    ws["A5"].font = bold
    ws["A5"].alignment = left_mid
    ws.merge_cells("B5:D5")
    ws["B5"].value = phone
    ws["B5"].alignment = left_mid

    ws["F5"].value = "Fax:"
    ws["F5"].font = bold
    ws["F5"].alignment = left_mid
    ws.merge_cells("G5:I5")
    ws["G5"].value = fax
    ws["G5"].alignment = left_mid

    # Attn
    ws["A7"].value = "Attn :"
    ws["A7"].font = bold
    ws["A7"].alignment = left_mid
    ws.merge_cells("B7:D7")
    ws["B7"].value = attn
    ws["B7"].alignment = left_mid

    # Invoice No / Date / Surat Jalan (kanan atas)
    ws["H6"].value = "Invoice"
    ws["H6"].font = bold
    ws["H6"].alignment = right
    ws["I6"].value = invoice_no
    ws["I6"].alignment = left_mid

    ws["H7"].value = "Date"
    ws["H7"].font = bold
    ws["H7"].alignment = right
    ws["I7"].value = inv_date
    ws["I7"].alignment = left_mid

    ws.merge_cells("G8:H8")
    ws["G8"].value = "No. Surat Jalan"
    ws["G8"].font = bold
    ws["G8"].alignment = right
    ws["I8"].value = no_surat_jalan
    ws["I8"].alignment = left_mid

    # ======================
    # Box Ref/Sales/Ship/Terms (seperti template)
    # ======================
    ws.merge_cells("A10:B10")
    ws["A10"].value = "Ref No."
    ws["A10"].font = bold
    ws["A10"].alignment = center

    ws.merge_cells("C10:E10")
    ws["C10"].value = "Sales Person"
    ws["C10"].font = bold
    ws["C10"].alignment = center

    ws["F10"].value = "Ship Via"
    ws["F10"].font = bold
    ws["F10"].alignment = center

    ws["G10"].value = "Ship Date"
    ws["G10"].font = bold
    ws["G10"].alignment = center

    ws.merge_cells("H10:I10")
    ws["H10"].value = "Terms"
    ws["H10"].font = bold
    ws["H10"].alignment = center

    ws.merge_cells("A11:B11")
    ws["A11"].value = ref_no
    ws["A11"].alignment = center

    ws.merge_cells("C11:E11")
    ws["C11"].value = sales_person
    ws["C11"].alignment = center

    ws["F11"].value = ship_via
    ws["F11"].alignment = center

    ws["G11"].value = ship_date
    ws["G11"].alignment = center

    ws.merge_cells("H11:I11")
    ws["H11"].value = terms
    ws["H11"].alignment = center

    # Border box (outer tebal, inner tipis)
    set_table_borders(10, 1, 11, 9)

    # ======================
    # Items table
    # ======================
    hdr_row = 13
    ws["A13"].value = "Qty"
    ws["B13"].value = "Unit"
    ws["C13"].value = "Date"
    ws.merge_cells("D13:G13")
    ws["D13"].value = "Description"
    ws["H13"].value = "Price"
    ws["I13"].value = "Amount (IDR)"

    for cell in ["A13", "B13", "C13", "D13", "H13", "I13"]:
        ws[cell].font = bold
        ws[cell].alignment = center

    items = inv.get("items") or []
    start_row = 14
    max_rows = max(10, len(items))  # minimal 10 baris kosong seperti template
    subtotal = 0

    for idx in range(max_rows):
        r = start_row + idx

        # default alignment
        ws[f"A{r}"].alignment = center
        ws[f"B{r}"].alignment = center
        ws[f"C{r}"].alignment = center
        ws[f"H{r}"].alignment = right
        ws[f"I{r}"].alignment = right

        # merge description D-G setiap baris
        ws.merge_cells(f"D{r}:G{r}")
        ws[f"D{r}"].alignment = left

        if idx < len(items):
            it = items[idx]
            qty = float(it.get("qty") or 0)
            unit = (it.get("unit") or "").strip()
            desc = (it.get("description") or "").strip()
            price = int(it.get("price") or 0)
            line_date = it.get("date") or inv_date
            amount = int(round(qty * price))
            subtotal += amount

            ws[f"A{r}"].value = qty if qty % 1 != 0 else int(qty)
            ws[f"B{r}"].value = unit
            ws[f"C{r}"].value = line_date
            ws[f"D{r}"].value = desc
            ws[f"H{r}"].value = price
            ws[f"I{r}"].value = amount

            money(ws[f"H{r}"])
            money(ws[f"I{r}"])

    last_table_row = start_row + max_rows - 1
    set_table_borders(hdr_row, 1, last_table_row, 9)

    # ======================
    # Totals (kanan bawah)
    # ======================
    freight = int(inv.get("freight") or 0)
    ppn_rate = float(inv.get("ppn_rate") or 0.11)
    deposit = int(inv.get("deposit") or 0)

    total_before_ppn = subtotal + freight
    ppn = int(round(total_before_ppn * ppn_rate))
    balance = total_before_ppn + ppn - deposit

    sum_row = last_table_row + 1
    labels = [
        ("Total", subtotal, True),
        ("Freight", freight, False),
        ("Total", total_before_ppn, True),
        (f"PPN {int(ppn_rate * 100)}%", ppn, False),
        ("Less: Deposit", deposit, False),
        ("Balance Due", balance, True),
    ]
    for i, (lab, val, is_bold) in enumerate(labels):
        rr = sum_row + i
        ws[f"H{rr}"].value = lab
        ws[f"H{rr}"].alignment = right
        ws[f"H{rr}"].font = Font(bold=is_bold)

        ws[f"I{rr}"].value = val
        ws[f"I{rr}"].alignment = right
        ws[f"I{rr}"].font = Font(bold=is_bold)
        money(ws[f"I{rr}"])

    set_table_borders(sum_row, 8, sum_row + len(labels) - 1, 9)

    # ======================
    # Payment info (kiri bawah)
    # ======================
    ws.merge_cells(f"A{sum_row}:G{sum_row}")
    ws[f"A{sum_row}"].value = "Please Transfer Full Amount to:"
    ws[f"A{sum_row}"].font = bold
    ws[f"A{sum_row}"].alignment = left_mid

    pay_lines = [
        f"Beneficiary : {payment['beneficiary']}",
        f"Bank Name   : {payment['bank_name']}",
        f"Branch      : {payment['branch']}",
        f"IDR Acct    : {payment['idr_acct']}",
    ]
    for i, line in enumerate(pay_lines, start=1):
        ws.merge_cells(f"A{sum_row + i}:G{sum_row + i}")
        ws[f"A{sum_row + i}"].value = line
        ws[f"A{sum_row + i}"].alignment = left_mid

    # ======================
    # Box tanda tangan (bawah tengah-kanan)
    # ======================
    box_top = sum_row + 7
    box_bottom = box_top + 6
    ws.merge_cells(f"F{box_top}:I{box_top}")
    ws[f"F{box_top}"].value = "PT. Sarana Trans Bersama Jaya"
    ws[f"F{box_top}"].alignment = center
    ws[f"F{box_top}"].font = bold

    set_table_borders(box_top, 6, box_bottom, 9)

    # Footer
    footer_row = box_bottom + 2
    ws.merge_cells(f"A{footer_row}:I{footer_row}")
    ws[f"A{footer_row}"].value = "Please kindly fax to our attention upon receipt"
    ws[f"A{footer_row}"].alignment = center

    try:
        folder = str(FILES_DIR)
    except Exception:
        folder = "static/files"
    os.makedirs(folder, exist_ok=True)

    out_path = os.path.join(folder, f"{fname_base}.xlsx")
    wb.save(out_path)
    return f"{fname_base}.xlsx"


def create_invoice_pdf(inv: dict, fname_base: str) -> str:
    try:
        folder = str(FILES_DIR)
    except Exception:
        folder = "static/files"
    os.makedirs(folder, exist_ok=True)

    out_path = os.path.join(folder, f"{fname_base}.pdf")
    c = canvas.Canvas(out_path, pagesize=A4)
    width, height = A4

    def draw_text(x, y, txt, size=10, bold=False):
        c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
        c.drawString(x, y, txt)

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

    y = height - 50
    draw_text(40, y, "INVOICE", 18, True)
    draw_text(380, y, f"Invoice: {invoice_no}", 10, True)
    y -= 14
    draw_text(380, y, f"Date: {inv_date}", 10, True)
    y -= 14
    draw_text(380, y, f"No. Surat Jalan: {no_surat_jalan}", 9, False)

    y -= 30
    draw_text(40, y, "Bill To:", 11, True)
    draw_text(300, y, "Ship To:", 11, True)

    y -= 14
    bt_lines = [bill_to.get("name", ""), bill_to.get("address", ""), bill_to.get("address2", "")]
    st_lines = [ship_to.get("name", ""), ship_to.get("address", ""), ship_to.get("address2", "")]
    bt_lines = [x for x in bt_lines if (x or "").strip()]
    st_lines = [x for x in st_lines if (x or "").strip()]

    yy = y
    for line in bt_lines[:4]:
        draw_text(40, yy, str(line), 9, False)
        yy -= 12
    yy2 = y
    for line in st_lines[:4]:
        draw_text(300, yy2, str(line), 9, False)
        yy2 -= 12

    y = min(yy, yy2) - 10
    draw_text(40, y, f"Phone: {phone}", 9, False)
    draw_text(300, y, f"Fax: {fax}", 9, False)
    y -= 12
    draw_text(40, y, f"Attn: {attn}", 9, False)

    y -= 18
    draw_text(40, y, f"Ref No.: {ref_no}", 9, False)
    draw_text(170, y, f"Sales Person: {sales_person}", 9, False)
    draw_text(380, y, f"Ship Via: {ship_via}", 9, False)
    y -= 12
    draw_text(380, y, f"Ship Date: {ship_date}", 9, False)
    draw_text(40, y, f"Terms: {terms}", 9, False)

    y -= 20
    table_top = y
    table_height = 220
    c.rect(40, table_top - table_height, width - 80, table_height, stroke=1, fill=0)

    header_y = table_top - 18
    draw_text(50, header_y, "Qty", 9, True)
    draw_text(90, header_y, "Unit", 9, True)
    draw_text(130, header_y, "Date", 9, True)
    draw_text(190, header_y, "Description", 9, True)
    draw_text(430, header_y, "Price", 9, True)
    draw_text(500, header_y, "Amount", 9, True)

    row_y = header_y - 20
    subtotal = 0
    max_rows = 10
    for idx in range(max_rows):
        if idx < len(items):
            it = items[idx]
            qty = it.get("qty") or 0
            unit = it.get("unit") or ""
            dt = it.get("date") or inv_date
            desc = it.get("description") or ""
            price = int(it.get("price") or 0)
            amount = int(round(float(qty) * price))
            subtotal += amount

            draw_text(50, row_y, str(qty), 9, False)
            draw_text(90, row_y, str(unit), 9, False)
            draw_text(130, row_y, str(dt), 9, False)
            draw_text(190, row_y, str(desc)[:55], 9, False)
            draw_text(430, row_y, f"{price:,}".replace(",", "."), 9, False)
            draw_text(500, row_y, f"{amount:,}".replace(",", "."), 9, False)
        row_y -= 18

    total_before_ppn = subtotal + freight
    ppn = int(round(total_before_ppn * ppn_rate))
    balance = total_before_ppn + ppn - deposit

    y2 = table_top - table_height - 20
    draw_text(380, y2, "Total:", 10, True)
    draw_text(500, y2, f"{subtotal:,}".replace(",", "."), 10, False)
    y2 -= 14
    draw_text(380, y2, "Freight:", 10, True)
    draw_text(500, y2, f"{freight:,}".replace(",", "."), 10, False)
    y2 -= 14
    draw_text(380, y2, "Total:", 10, True)
    draw_text(500, y2, f"{total_before_ppn:,}".replace(",", "."), 10, False)
    y2 -= 14
    draw_text(380, y2, f"PPN {int(ppn_rate*100)}%:", 10, True)
    draw_text(500, y2, f"{ppn:,}".replace(",", "."), 10, False)
    y2 -= 14
    draw_text(380, y2, "Less: Deposit:", 10, True)
    draw_text(500, y2, f"{deposit:,}".replace(",", "."), 10, False)
    y2 -= 16
    draw_text(380, y2, "Balance Due:", 11, True)
    draw_text(500, y2, f"{balance:,}".replace(",", "."), 11, True)

    y3 = 90
    draw_text(40, y3 + 40, "Please Transfer Full Amount to:", 10, True)
    draw_text(40, y3 + 24, f"Beneficiary : {payment.get('beneficiary', '')}", 9, False)
    draw_text(40, y3 + 12, f"Bank Name   : {payment.get('bank_name', '')}", 9, False)
    draw_text(40, y3 + 0, f"Branch      : {payment.get('branch', '')}", 9, False)
    draw_text(40, y3 - 12, f"IDR Acct    : {payment.get('idr_acct', '')}", 9, False)

    c.showPage()
    c.save()
    return f"{fname_base}.pdf"


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
            out_text = "❓ <b>3. Phone?</b> (boleh kosong, sebut <b>strip</b> jika tidak ada)"
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
        out_text = "❓ <b>3. Phone?</b> (boleh kosong, sebut <b>strip</b> jika tidak ada)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_phone":
        state["data"]["phone"] = "" if text.strip() in ("-", "") else text.strip()
        state["step"] = "inv_fax"
        conversations[sid] = state
        out_text = "❓ <b>4. Fax?</b> (boleh kosong, sebut <b>strip</b> jika tidak ada)"
        if history_id_in:
            db_append_message(int(history_id_in), "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=[])
            db_update_state(int(history_id_in), state)
        return {"text": out_text, "history_id": history_id_in}

    if state.get("step") == "inv_fax":
        state["data"]["fax"] = "" if text.strip() in ("-", "") else text.strip()
        state["step"] = "inv_attn"
        conversations[sid] = state
        out_text = "❓ <b>5. Attn?</b> (default: Accounting / Finance | sebut <b>strip</b> untuk default)"
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
        out_text = "❓ <b>6B. Jenis Limbah / Kode Limbah?</b><br><i>(Contoh: 'A102d' atau 'aki baterai bekas' | atau sebut <b>NON B3</b>)</i>"
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

        out_text = f"❌ Maaf, limbah '<b>{text}</b>' tidak ditemukan.<br><br>Ucapkan kode/jenis lain atau <b>NON B3</b>."
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

        xlsx = create_invoice_xlsx(state["data"], fname_base)
        pdf_preview = create_invoice_pdf(state["data"], fname_base)

        files = [
            {"type": "xlsx", "filename": xlsx, "url": f"/download/{xlsx}"},
            {"type": "pdf", "filename": pdf_preview, "url": f"/download/{pdf_preview}"},
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
            "📌 Preview: PDF<br>"
            "📌 Download: Excel (.xlsx) / PDF"
        )

        db_append_message(history_id, "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=files)
        return {"text": out_text, "files": files, "history_id": history_id}

    return None
