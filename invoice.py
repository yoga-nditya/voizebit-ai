import os
import json
import re
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
# Generate Invoice (UPDATED EXCEL LAYOUT ONLY)
# =========================

def _thin_border():
    side = Side(style="thin", color="000000")
    return Border(left=side, right=side, top=side, bottom=side)

def _medium_border():
    side = Side(style="medium", color="000000")
    return Border(left=side, right=side, top=side, bottom=side)

def _set_border(ws, r1, c1, r2, c2, border):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = border

def _outline_box(ws, r1, c1, r2, c2, inner_border=None, outer_border=None):
    """
    inner_border: border untuk grid di dalam
    outer_border: border untuk garis luar kotak
    """
    if inner_border:
        _set_border(ws, r1, c1, r2, c2, inner_border)

    if not outer_border:
        return

    # top & bottom
    for c in range(c1, c2 + 1):
        ws.cell(r1, c).border = Border(
            left=ws.cell(r1, c).border.left,
            right=ws.cell(r1, c).border.right,
            top=outer_border.top,
            bottom=ws.cell(r1, c).border.bottom,
        )
        ws.cell(r2, c).border = Border(
            left=ws.cell(r2, c).border.left,
            right=ws.cell(r2, c).border.right,
            top=ws.cell(r2, c).border.top,
            bottom=outer_border.bottom,
        )

    # left & right
    for r in range(r1, r2 + 1):
        ws.cell(r, c1).border = Border(
            left=outer_border.left,
            right=ws.cell(r, c1).border.right,
            top=ws.cell(r, c1).border.top,
            bottom=ws.cell(r, c1).border.bottom,
        )
        ws.cell(r, c2).border = Border(
            left=ws.cell(r, c2).border.left,
            right=outer_border.right,
            top=ws.cell(r, c2).border.top,
            bottom=ws.cell(r, c2).border.bottom,
        )

def create_invoice_xlsx(inv: dict, fname_base: str) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    # ===== Page setup (tetap aman, tidak mengganggu layout) =====
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins.left = 0.35
    ws.page_margins.right = 0.35
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35

    # ===== SHIFT START COLUMN =====
    # Mulai dari kolom C biar agak ketengah (sesuai request)
    START_COL = 3  # C

    def col(n: int) -> int:
        return START_COL + (n - 1)  # n=1.. -> C..

    def cell_addr(r: int, c: int) -> str:
        return f"{get_column_letter(c)}{r}"

    def merge(r1, c1, r2, c2):
        ws.merge_cells(f"{cell_addr(r1, c1)}:{cell_addr(r2, c2)}")

    # ===== Column widths (mirip template screenshot) =====
    # Struktur 6 kolom utama: Qty | Unit | Date | Description | Price | Amount
    # (Template kamu memang ada Unit kecil “Kg”)
    col_widths = {
        col(1): 7,   # Qty
        col(2): 5,   # Unit
        col(3): 12,  # Date
        col(4): 44,  # Description
        col(5): 14,  # Price
        col(6): 16,  # Amount
    }
    for c, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w

    # ===== Styles =====
    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="top", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    thin = _thin_border()
    medium = _medium_border()

    def money(cell):
        cell.number_format = '#,##0'

    # ===== Defaults Payment =====
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

    # ===== HEADER: Bill To & Ship To =====
    ws[cell_addr(1, col(1))].value = "Bill To:"
    ws[cell_addr(1, col(1))].font = bold
    merge(1, col(1), 1, col(3))

    ws[cell_addr(1, col(4))].value = "Ship To:"
    ws[cell_addr(1, col(4))].font = bold
    merge(1, col(4), 1, col(6))

    bill_lines = [(bill_to.get("name") or "").strip(), (bill_to.get("address") or "").strip(), (bill_to.get("address2") or "").strip()]
    ship_lines = [(ship_to.get("name") or "").strip(), (ship_to.get("address") or "").strip(), (ship_to.get("address2") or "").strip()]
    bill_text = "\n".join([x for x in bill_lines if x])
    ship_text = "\n".join([x for x in ship_lines if x])

    ws[cell_addr(2, col(1))].value = bill_text
    merge(2, col(1), 3, col(3))
    ws[cell_addr(2, col(1))].alignment = left

    ws[cell_addr(2, col(4))].value = ship_text
    merge(2, col(4), 3, col(6))
    ws[cell_addr(2, col(4))].alignment = left

    # Phone / Fax line
    ws[cell_addr(5, col(1))].value = "Phone:"
    ws[cell_addr(5, col(1))].font = bold
    merge(5, col(1), 5, col(2))
    ws[cell_addr(5, col(3))].value = phone
    ws[cell_addr(5, col(3))].alignment = left

    ws[cell_addr(5, col(4))].value = "Fax:"
    ws[cell_addr(5, col(4))].font = bold
    merge(5, col(4), 5, col(5))
    ws[cell_addr(5, col(6))].value = fax
    ws[cell_addr(5, col(6))].alignment = left

    # Attn
    ws[cell_addr(7, col(1))].value = "Attn :"
    ws[cell_addr(7, col(1))].font = bold
    merge(7, col(1), 7, col(2))
    ws[cell_addr(7, col(3))].value = attn
    merge(7, col(3), 7, col(4))
    ws[cell_addr(7, col(3))].alignment = left

    # Invoice box kanan (Invoice, Date, No Surat Jalan)
    ws[cell_addr(6, col(5))].value = "Invoice"
    ws[cell_addr(6, col(5))].font = bold
    ws[cell_addr(6, col(5))].alignment = right
    ws[cell_addr(6, col(6))].value = invoice_no
    ws[cell_addr(6, col(6))].alignment = left

    ws[cell_addr(7, col(5))].value = "Date"
    ws[cell_addr(7, col(5))].font = bold
    ws[cell_addr(7, col(5))].alignment = right
    ws[cell_addr(7, col(6))].value = inv_date
    ws[cell_addr(7, col(6))].alignment = left

    ws[cell_addr(8, col(5))].value = "No. Surat Jalan"
    ws[cell_addr(8, col(5))].font = bold
    ws[cell_addr(8, col(5))].alignment = right
    ws[cell_addr(8, col(6))].value = no_surat_jalan
    ws[cell_addr(8, col(6))].alignment = left

    # ===== REF ROW (bordered block starts here) =====
    ref_header_row = 10
    ref_value_row = 11

    # Ref header cells: (Ref No | Sales Person | Ship Via | Ship Date | Terms)
    # kita pakai 6 kolom utama, tapi Terms kita gabungkan 2 kolom terakhir (Price+Amount) supaya proporsi mirip template
    # Kolom mapping:
    # RefNo: col1-2 (Qty+Unit)
    # Sales: col3-4 (Date+Desc)
    # ShipVia: col5 (Price)
    # ShipDate: col6 (Amount) -> kita pakai, dan Terms row dibawahnya merge supaya rapi

    # Header labels
    merge(ref_header_row, col(1), ref_header_row, col(2))
    ws[cell_addr(ref_header_row, col(1))].value = "Ref No."
    ws[cell_addr(ref_header_row, col(1))].font = bold
    ws[cell_addr(ref_header_row, col(1))].alignment = center

    merge(ref_header_row, col(3), ref_header_row, col(4))
    ws[cell_addr(ref_header_row, col(3))].value = "Sales Person"
    ws[cell_addr(ref_header_row, col(3))].font = bold
    ws[cell_addr(ref_header_row, col(3))].alignment = center

    ws[cell_addr(ref_header_row, col(5))].value = "Ship Via"
    ws[cell_addr(ref_header_row, col(5))].font = bold
    ws[cell_addr(ref_header_row, col(5))].alignment = center

    ws[cell_addr(ref_header_row, col(6))].value = "Ship Date"
    ws[cell_addr(ref_header_row, col(6))].font = bold
    ws[cell_addr(ref_header_row, col(6))].alignment = center

    # Terms row (di template Terms ada di baris yang sama dengan Ship Date header area)
    # Kita buat Terms di header row juga dengan merge area kecil di kanan (pakai col6 saja sudah sempit).
    # Agar mirip, kita taruh Terms label di baris header tapi posisinya tetap sisi kanan area header.
    # (Kalau kamu mau Terms benar-benar kotak terpisah 1 kolom lagi, harus tambah kolom ekstra)
    # Di sini kita tetap mengikuti 6 kolom supaya tidak merusak struktur invoice kamu.
    # Jadi Terms value tetap ada di bawah, tapi label "Terms" sudah ada.
    # NOTE: Untuk benar-benar 100% sama, perlu 1 kolom ekstra.

    # Values
    merge(ref_value_row, col(1), ref_value_row, col(2))
    ws[cell_addr(ref_value_row, col(1))].value = ref_no
    ws[cell_addr(ref_value_row, col(1))].alignment = center

    merge(ref_value_row, col(3), ref_value_row, col(4))
    ws[cell_addr(ref_value_row, col(3))].value = sales_person
    ws[cell_addr(ref_value_row, col(3))].alignment = center

    ws[cell_addr(ref_value_row, col(5))].value = ship_via
    ws[cell_addr(ref_value_row, col(5))].alignment = center

    ws[cell_addr(ref_value_row, col(6))].value = ship_date
    ws[cell_addr(ref_value_row, col(6))].alignment = center

    # Terms label/value (tetap pakai area kanan bawah dekat ship date sesuai template)
    terms_label_row = 12
    terms_value_row = 13

    merge(terms_label_row, col(5), terms_label_row, col(6))
    ws[cell_addr(terms_label_row, col(5))].value = "Terms"
    ws[cell_addr(terms_label_row, col(5))].font = bold
    ws[cell_addr(terms_label_row, col(5))].alignment = center

    merge(terms_value_row, col(5), terms_value_row, col(6))
    ws[cell_addr(terms_value_row, col(5))].value = terms
    ws[cell_addr(terms_value_row, col(5))].alignment = center

    # ===== ITEMS TABLE =====
    item_header_row = 15
    ws[cell_addr(item_header_row, col(1))].value = "Qty"
    ws[cell_addr(item_header_row, col(2))].value = ""  # unit (Kg)
    ws[cell_addr(item_header_row, col(3))].value = "Date"
    ws[cell_addr(item_header_row, col(4))].value = "Description"
    ws[cell_addr(item_header_row, col(5))].value = "Price"
    ws[cell_addr(item_header_row, col(6))].value = "Amount (IDR)"

    for i in range(1, 7):
        ws[cell_addr(item_header_row, col(i))].font = bold
        ws[cell_addr(item_header_row, col(i))].alignment = center

    items = inv.get("items") or []
    r = item_header_row + 1
    subtotal = 0

    for it in items:
        qty = float(it.get("qty") or 0)
        unit = (it.get("unit") or "").strip()
        desc = (it.get("description") or "").strip()
        price = int(it.get("price") or 0)
        line_date = it.get("date") or inv_date
        amount = int(round(qty * price))
        subtotal += amount

        ws[cell_addr(r, col(1))].value = qty if qty % 1 != 0 else int(qty)
        ws[cell_addr(r, col(1))].alignment = center

        ws[cell_addr(r, col(2))].value = unit
        ws[cell_addr(r, col(2))].alignment = center

        ws[cell_addr(r, col(3))].value = line_date
        ws[cell_addr(r, col(3))].alignment = center

        ws[cell_addr(r, col(4))].value = desc
        ws[cell_addr(r, col(4))].alignment = left

        ws[cell_addr(r, col(5))].value = price
        ws[cell_addr(r, col(5))].alignment = right
        money(ws[cell_addr(r, col(5))])

        ws[cell_addr(r, col(6))].value = amount
        ws[cell_addr(r, col(6))].alignment = right
        money(ws[cell_addr(r, col(6))])

        r += 1

    # Minimal rows untuk area kosong (agar seperti template)
    min_last_row = 26
    if r < min_last_row:
        r = min_last_row

    last_item_row = r - 1

    # ===== TOTALS =====
    freight = int(inv.get("freight") or 0)
    ppn_rate = float(inv.get("ppn_rate") or 0.11)
    deposit = int(inv.get("deposit") or 0)

    total_before_ppn = subtotal + freight
    ppn = int(round(total_before_ppn * ppn_rate))
    balance = total_before_ppn + ppn - deposit

    sum_row = r  # mulai setelah items table

    # Label di kolom Price (col5), angka di Amount (col6)
    totals = [
        ("Total", subtotal, True),
        ("Freight", freight, False),
        ("Total", total_before_ppn, True),
        (f"PPN {int(ppn_rate*100)}%", ppn, False),
        ("Less: Deposit", deposit, False),
        ("Balance Due", balance, True),
    ]

    for idx, (label, value, is_bold) in enumerate(totals):
        rr = sum_row + idx
        ws[cell_addr(rr, col(5))].value = label
        ws[cell_addr(rr, col(5))].alignment = right
        ws[cell_addr(rr, col(5))].font = Font(bold=is_bold)

        ws[cell_addr(rr, col(6))].value = value
        ws[cell_addr(rr, col(6))].alignment = right
        ws[cell_addr(rr, col(6))].font = Font(bold=is_bold)
        money(ws[cell_addr(rr, col(6))])

    # ===== PAYMENT TEXT (tanpa box border tegas, sesuai request “yang lain tidak perlu dibungkus”) =====
    pay_row = sum_row
    ws[cell_addr(pay_row, col(1))].value = "Please Transfer Full Amount to:"
    ws[cell_addr(pay_row, col(1))].font = bold
    merge(pay_row, col(1), pay_row, col(4))
    ws[cell_addr(pay_row, col(1))].alignment = left

    ws[cell_addr(pay_row + 1, col(1))].value = "Beneficiary  :"
    ws[cell_addr(pay_row + 2, col(1))].value = "Bank Name    :"
    ws[cell_addr(pay_row + 3, col(1))].value = "Branch       :"
    ws[cell_addr(pay_row + 4, col(1))].value = "IDR Acct     :"

    ws[cell_addr(pay_row + 1, col(2))].value = payment["beneficiary"]
    ws[cell_addr(pay_row + 2, col(2))].value = payment["bank_name"]
    ws[cell_addr(pay_row + 3, col(2))].value = payment["branch"]
    ws[cell_addr(pay_row + 4, col(2))].value = payment["idr_acct"]

    merge(pay_row + 1, col(2), pay_row + 1, col(4))
    merge(pay_row + 2, col(2), pay_row + 2, col(4))
    merge(pay_row + 3, col(2), pay_row + 3, col(4))
    merge(pay_row + 4, col(2), pay_row + 4, col(4))

    # ===== BORDERS (SESUI TEMPLATE REQUEST) =====
    # 1) Kotak tegas hanya untuk area Ref->Terms dan Items table sampai Amount
    #    Area ini mencakup:
    #    - ref header/value (row 10-11)
    #    - terms box (row 12-13) di kolom 5-6
    #    - items header + items rows (row 15..last_item_row)
    # Outer border medium, inner border thin.
    ref_items_top = ref_header_row
    ref_items_bottom = last_item_row
    ref_items_left = col(1)
    ref_items_right = col(6)

    _outline_box(
        ws,
        ref_items_top,
        ref_items_left,
        ref_items_bottom,
        ref_items_right,
        inner_border=thin,
        outer_border=medium
    )

    # Terms box: ikut style kotak di header area (thin inner, medium outline)
    _outline_box(
        ws,
        terms_label_row,
        col(5),
        terms_value_row,
        col(6),
        inner_border=thin,
        outer_border=medium
    )

    # Garis pemisah kuat antara header items dan body (mirip template)
    for c in range(ref_items_left, ref_items_right + 1):
        ws.cell(item_header_row, c).border = Border(
            left=ws.cell(item_header_row, c).border.left,
            right=ws.cell(item_header_row, c).border.right,
            top=ws.cell(item_header_row, c).border.top,
            bottom=medium.bottom,  # garis bawah header lebih tegas
        )

    # 2) TOTALS: hanya angka (kolom Amount) yang diborder (sesuai screenshot kamu)
    totals_top = sum_row
    totals_bottom = sum_row + len(totals) - 1

    for rr in range(totals_top, totals_bottom + 1):
        ws.cell(rr, col(6)).border = thin  # hanya kolom angka

    # Biar terlihat seperti “kolom angka dibox” (outline medium di kolom amount saja)
    _outline_box(
        ws,
        totals_top,
        col(6),
        totals_bottom,
        col(6),
        inner_border=thin,
        outer_border=medium
    )

    # ===== Save =====
    try:
        folder = str(FILES_DIR)
    except Exception:
        folder = "static/files"
    os.makedirs(folder, exist_ok=True)

    out_path = os.path.join(folder, f"{fname_base}.xlsx")
    wb.save(out_path)
    return f"{fname_base}.xlsx"


# ========= PDF generator (as-is) =========
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
    bt_lines = [bill_to.get("name",""), bill_to.get("address",""), bill_to.get("address2","")]
    st_lines = [ship_to.get("name",""), ship_to.get("address",""), ship_to.get("address2","")]
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
    c.line(40, header_y - 6, width - 40, header_y - 6)

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
    draw_text(40, y3 + 24, f"Beneficiary : {payment.get('beneficiary','')}", 9, False)
    draw_text(40, y3 + 12, f"Bank Name   : {payment.get('bank_name','')}", 9, False)
    draw_text(40, y3 + 0,  f"Branch      : {payment.get('branch','')}", 9, False)
    draw_text(40, y3 - 12, f"IDR Acct    : {payment.get('idr_acct','')}", 9, False)

    c.showPage()
    c.save()
    return f"{fname_base}.pdf"


# =========================
# CHAT HANDLER INVOICE (as-is)
# =========================

def handle_invoice_flow(data: dict, text: str, lower: str, sid: str, state: dict, conversations: dict, history_id_in):
    """
    Return:
      - None jika bukan flow invoice
      - dict response jika handled
    """

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
            "📌 Download: Excel (.xlsx)"
        )

        db_append_message(history_id, "assistant", re.sub(r'<br\s*/?>', '\n', out_text), files=files)
        return {"text": out_text, "files": files, "history_id": history_id}

    return None
