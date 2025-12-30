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
# Generate Invoice (ONLY XLSX LAYOUT UPDATED)
# =========================

def _thin_border():
    side = Side(style="thin", color="000000")
    return Border(left=side, right=side, top=side, bottom=side)

def _set_border(ws, r1, c1, r2, c2, border):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = border

def create_invoice_xlsx(inv: dict, fname_base: str) -> str:
    """
    Layout dibuat mirip screenshot:
    - Start dari kolom C (biar lebih ke tengah)
    - Blok Bill To / Ship To
    - Header Invoice kanan (Invoice No, Date, No Surat Jalan)
    - Header bar: Ref No / Sales Person / Ship Via / Ship Date / Terms
    - Table item: Qty | Unit | Date | Description | Price | Amount (IDR)
    - Summary kanan bawah (Total, Freight, Total, PPN, Deposit, Balance Due)
    - Blok Please Transfer Full Amount to
    """

    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    # Page setup supaya lebih mirip (fit 1 halaman lebar)
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25

    # === SHIFT KE KANAN: mulai dari kolom C ===
    SHIFT = 2  # A,B kosong -> mulai C
    def C(col_letter: str) -> str:
        """Map 'A'.. ke kolom yang sudah digeser (A->C, B->D, dst)."""
        base = ord(col_letter.upper()) - ord("A") + 1
        return get_column_letter(base + SHIFT)

    def cell(addr: str) -> str:
        """addr misal 'A1' -> jadi 'C1' kalau SHIFT=2"""
        m = re.match(r"^([A-Z]+)(\d+)$", addr.upper().strip())
        if not m:
            return addr
        col, row = m.group(1), m.group(2)
        # konversi col multi-letter ke index
        idx = 0
        for ch in col:
            idx = idx * 26 + (ord(ch) - ord("A") + 1)
        idx += SHIFT
        return f"{get_column_letter(idx)}{row}"

    # Lebarin area kosong kiri (A,B) biar benar-benar “ketengah”
    ws.column_dimensions["A"].width = 2.2
    ws.column_dimensions["B"].width = 2.2

    # Lebar kolom utama (yang tadinya A..F)
    col_widths = {
        C("A"): 8,   # Qty
        C("B"): 6,   # Unit
        C("C"): 12,  # Date
        C("D"): 45,  # Description
        C("E"): 14,  # Price
        C("F"): 16,  # Amount (IDR)
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # Row height biar mirip grid screenshot
    for r in range(1, 60):
        ws.row_dimensions[r].height = 16

    border = _thin_border()
    bold = Font(bold=True)
    font_normal = Font(name="Calibri", size=11)
    font_bold = Font(name="Calibri", size=11, bold=True)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="top", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center", wrap_text=True)
    left_mid = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def money(cell_obj):
        cell_obj.number_format = '#,##0'

    # Defaults payment
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

    # ========= Header Bill To / Ship To =========
    ws[cell("A1")].value = "Bill To:"
    ws[cell("A1")].font = font_bold
    ws.merge_cells(f"{cell('A1')}:{cell('C1')}")

    ws[cell("D1")].value = "Ship To:"
    ws[cell("D1")].font = font_bold
    ws.merge_cells(f"{cell('D1')}:{cell('F1')}")

    bill_lines = [
        (bill_to.get("name") or "").strip(),
        (bill_to.get("address") or "").strip(),
        (bill_to.get("address2") or "").strip(),
    ]
    ship_lines = [
        (ship_to.get("name") or "").strip(),
        (ship_to.get("address") or "").strip(),
        (ship_to.get("address2") or "").strip(),
    ]
    bill_text = "\n".join([x for x in bill_lines if x])
    ship_text = "\n".join([x for x in ship_lines if x])

    ws[cell("A2")].value = bill_text
    ws.merge_cells(f"{cell('A2')}:{cell('C4')}")
    ws[cell("A2")].alignment = left
    ws[cell("A2")].font = font_normal

    ws[cell("D2")].value = ship_text
    ws.merge_cells(f"{cell('D2')}:{cell('F4')}")
    ws[cell("D2")].alignment = left
    ws[cell("D2")].font = font_normal

    # Phone / Fax row
    ws[cell("A6")].value = "Phone:"
    ws[cell("A6")].font = font_bold
    ws.merge_cells(f"{cell('A6')}:{cell('B6')}")
    ws[cell("C6")].value = phone
    ws[cell("C6")].alignment = left_mid

    ws[cell("D6")].value = "Fax:"
    ws[cell("D6")].font = font_bold
    ws.merge_cells(f"{cell('D6')}:{cell('E6')}")
    ws[cell("F6")].value = fax
    ws[cell("F6")].alignment = left_mid

    # ========= Invoice box kanan =========
    ws[cell("E8")].value = "Invoice"
    ws[cell("E8")].font = font_bold
    ws[cell("E8")].alignment = center
    ws[cell("F8")].value = invoice_no
    ws[cell("F8")].alignment = center

    ws[cell("E9")].value = "Date"
    ws[cell("E9")].font = font_bold
    ws[cell("E9")].alignment = center
    ws[cell("F9")].value = inv_date
    ws[cell("F9")].alignment = center

    ws[cell("E10")].value = "No. Surat Jalan"
    ws[cell("E10")].font = font_bold
    ws[cell("E10")].alignment = center
    ws[cell("F10")].value = no_surat_jalan
    ws[cell("F10")].alignment = center

    # ========= Attn row =========
    ws[cell("A8")].value = "Attn :"
    ws[cell("A8")].font = font_bold
    ws.merge_cells(f"{cell('A8')}:{cell('B8')}")
    ws[cell("C8")].value = attn
    ws.merge_cells(f"{cell('C8')}:{cell('D8')}")
    ws[cell("C8")].alignment = left_mid

    # ========= Bar header Ref/Sales/Ship/Terms =========
    ws.merge_cells(f"{cell('A12')}:{cell('B12')}")
    ws[cell("A12")].value = "Ref No."
    ws[cell("A12")].font = font_bold
    ws[cell("A12")].alignment = center

    ws.merge_cells(f"{cell('C12')}:{cell('D12')}")
    ws[cell("C12")].value = "Sales Person"
    ws[cell("C12")].font = font_bold
    ws[cell("C12")].alignment = center

    ws[cell("E12")].value = "Ship Via"
    ws[cell("E12")].font = font_bold
    ws[cell("E12")].alignment = center

    ws[cell("F12")].value = "Ship Date"
    ws[cell("F12")].font = font_bold
    ws[cell("F12")].alignment = center

    # Terms header (gabung seperti screenshot kanan)
    ws.merge_cells(f"{cell('E13')}:{cell('F13')}")
    ws[cell("E13")].value = "Terms"
    ws[cell("E13")].font = font_bold
    ws[cell("E13")].alignment = center

    # Values row
    ws.merge_cells(f"{cell('A13')}:{cell('B13')}")
    ws[cell("A13")].value = ref_no
    ws[cell("A13")].alignment = center

    ws.merge_cells(f"{cell('C13')}:{cell('D13')}")
    ws[cell("C13")].value = sales_person
    ws[cell("C13")].alignment = center

    ws[cell("E13")].alignment = center  # already merged

    ws[cell("E12")].alignment = center
    ws[cell("F12")].alignment = center

    ws[cell("E13")].alignment = center
    ws.merge_cells(f"{cell('E14')}:{cell('F14')}")
    ws[cell("E14")].value = terms
    ws[cell("E14")].alignment = center

    ws.merge_cells(f"{cell('E13')}:{cell('F13')}")  # keep

    # Ship values row (row 13) for E/F
    ws[cell("E13")].value = "Terms"  # keep header
    ws[cell("F12")].value = "Ship Date"  # keep header (safety)
    ws[cell("E13")].font = font_bold
    ws[cell("E13")].alignment = center

    ws[cell("E12")].value = "Ship Via"
    ws[cell("F12")].value = "Ship Date"

    ws[cell("E13")].value = "Terms"
    ws.merge_cells(f"{cell('E14')}:{cell('F14')}")
    ws[cell("E14")].value = terms
    ws[cell("E14")].alignment = center

    ws[cell("E13")].alignment = center

    ws[cell("E13")].font = font_bold
    ws[cell("E12")].font = font_bold
    ws[cell("F12")].font = font_bold

    ws[cell("E13")].alignment = center
    ws[cell("E12")].alignment = center
    ws[cell("F12")].alignment = center

    # Ship Via / Ship Date values row 13 (di screenshot biasanya baris bawah header)
    ws[cell("E13")].value = "Terms"  # keep terms header on row 13 merged above
    ws[cell("E13")].font = font_bold

    # Letakkan ship_via & ship_date pada row 13 (kolom E/F) tapi kita sudah pakai E13/F13 untuk Terms header merge.
    # Jadi ship_via & ship_date taruh di row 13 (E/F) TIDAK bisa kalau terms merge.
    # Solusi paling mirip screenshot: Ship Via/Ship Date value di row 13, Terms value di row 13/14.
    # Kita buat: Ship via/date di row 13 (E/F), Terms header di row 12 col? -> sudah.
    # Maka: ubah Terms header jadi di row 12 merged, Terms value di row 13 merged.
    ws.merge_cells(f"{cell('E12')}:{cell('F12')}")
    ws[cell("E12")].value = "Terms"
    ws[cell("E12")].font = font_bold
    ws[cell("E12")].alignment = center

    # Ship via/date pindah ke kolom E/F baris 12? Tidak.
    # Agar tetap ada Ship Via/Ship Date seperti kode awal, kita set:
    # Row 12: Ref/Sales/ShipVia/ShipDate (E/F) dan Terms (gabung) tidak mungkin bersamaan.
    # Jadi: kita ikuti screenshot yang kamu kirim (Terms ada di kolom paling kanan sendiri).
    # Cara paling dekat: buat 7 kolom isi (A..G), tapi kamu awalnya hanya 6 kolom.
    # Karena kamu minta jangan ubah besar selain tampilan, kita pertahankan 6 kolom.
    # Maka: Terms tetap di kanan (E/F merged) dan Ship Via/Ship Date tetap ada (E/F di row 12) tidak bisa.
    # Jadi kita pertahankan struktur kamu: Ship Via/Ship Date ada, Terms ada (E11:F12 di versi lama).
    # Untuk stabil, kita revert ke pola lama tapi sudah digeser.
    # ---- REVERT BAR HEADER (STABIL) ----
    # clear merges yang tadi bentrok
    for rg in list(ws.merged_cells.ranges):
        # hanya bersihkan area header bar yg bentrok (row 12-14 col E-F)
        if rg.min_row in (12, 13, 14) and rg.max_row in (12, 13, 14):
            ws.unmerge_cells(str(rg))

    # Header bar stabil seperti kode lama:
    ws.merge_cells(f"{cell('A12')}:{cell('B12')}")
    ws[cell("A12")].value = "Ref No."
    ws[cell("A12")].font = font_bold
    ws[cell("A12")].alignment = center

    ws.merge_cells(f"{cell('C12')}:{cell('D12')}")
    ws[cell("C12")].value = "Sales Person"
    ws[cell("C12")].font = font_bold
    ws[cell("C12")].alignment = center

    ws[cell("E12")].value = "Ship Via"
    ws[cell("E12")].font = font_bold
    ws[cell("E12")].alignment = center

    ws[cell("F12")].value = "Ship Date"
    ws[cell("F12")].font = font_bold
    ws[cell("F12")].alignment = center

    ws.merge_cells(f"{cell('E13')}:{cell('F13')}")
    ws[cell("E13")].value = "Terms"
    ws[cell("E13")].font = font_bold
    ws[cell("E13")].alignment = center

    ws.merge_cells(f"{cell('A13')}:{cell('B13')}")
    ws[cell("A13")].value = ref_no
    ws[cell("A13")].alignment = center

    ws.merge_cells(f"{cell('C13')}:{cell('D13')}")
    ws[cell("C13")].value = sales_person
    ws[cell("C13")].alignment = center

    ws[cell("E13")].alignment = center  # merged

    ws[cell("E14")].value = ship_via
    ws[cell("E14")].alignment = center

    ws[cell("F14")].value = ship_date
    ws[cell("F14")].alignment = center

    ws.merge_cells(f"{cell('E15')}:{cell('F15')}")
    ws[cell("E15")].value = terms
    ws[cell("E15")].alignment = center

    # ========= Table Header =========
    table_header_row = 17
    ws[cell(f"A{table_header_row}")].value = "Qty"
    ws[cell(f"B{table_header_row}")].value = ""
    ws[cell(f"C{table_header_row}")].value = "Date"
    ws[cell(f"D{table_header_row}")].value = "Description"
    ws[cell(f"E{table_header_row}")].value = "Price"
    ws[cell(f"F{table_header_row}")].value = "Amount (IDR)"

    for col in ["A", "B", "C", "D", "E", "F"]:
        ws[cell(f"{col}{table_header_row}")].font = font_bold
        ws[cell(f"{col}{table_header_row}")].alignment = center

    # ========= Items =========
    items = inv.get("items") or []
    r = table_header_row + 1
    subtotal = 0

    for it in items:
        qty = float(it.get("qty") or 0)
        unit = (it.get("unit") or "").strip()
        desc = (it.get("description") or "").strip()
        price = int(it.get("price") or 0)
        line_date = it.get("date") or inv_date
        amount = int(round(qty * price))
        subtotal += amount

        ws[cell(f"A{r}")].value = qty if qty % 1 != 0 else int(qty)
        ws[cell(f"A{r}")].alignment = center

        ws[cell(f"B{r}")].value = unit
        ws[cell(f"B{r}")].alignment = center

        ws[cell(f"C{r}")].value = line_date
        ws[cell(f"C{r}")].alignment = center

        ws[cell(f"D{r}")].value = desc
        ws[cell(f"D{r}")].alignment = left

        ws[cell(f"E{r}")].value = price
        ws[cell(f"E{r}")].alignment = right
        money(ws[cell(f"E{r}")])

        ws[cell(f"F{r}")].value = amount
        ws[cell(f"F{r}")].alignment = right
        money(ws[cell(f"F{r}")])

        r += 1

    # Minimal tinggi tabel biar mirip screenshot (ada ruang kosong)
    min_last_row = table_header_row + 12
    if r < min_last_row:
        r = min_last_row

    # Border area utama sampai tabel
    top_block_end = 15
    _set_border(ws, 1, 1 + SHIFT, top_block_end, 6 + SHIFT, border)

    # Border tabel item
    _set_border(ws, table_header_row, 1 + SHIFT, r - 1, 6 + SHIFT, border)

    # ========= Summary kanan bawah =========
    freight = int(inv.get("freight") or 0)
    ppn_rate = float(inv.get("ppn_rate") or 0.11)
    deposit = int(inv.get("deposit") or 0)

    total_before_ppn = subtotal + freight
    ppn = int(round(total_before_ppn * ppn_rate))
    balance = total_before_ppn + ppn - deposit

    sum_row = r + 1
    # Total
    ws[cell(f"E{sum_row}")].value = "Total"
    ws[cell(f"E{sum_row}")].alignment = right
    ws[cell(f"E{sum_row}")].font = font_bold
    ws[cell(f"F{sum_row}")].value = subtotal
    ws[cell(f"F{sum_row}")].alignment = right
    money(ws[cell(f"F{sum_row}")])

    # Freight
    ws[cell(f"E{sum_row+1}")].value = "Freight"
    ws[cell(f"E{sum_row+1}")].alignment = right
    ws[cell(f"F{sum_row+1}")].value = freight
    ws[cell(f"F{sum_row+1}")].alignment = right
    money(ws[cell(f"F{sum_row+1}")])

    # Total before ppn
    ws[cell(f"E{sum_row+2}")].value = "Total"
    ws[cell(f"E{sum_row+2}")].alignment = right
    ws[cell(f"E{sum_row+2}")].font = font_bold
    ws[cell(f"F{sum_row+2}")].value = total_before_ppn
    ws[cell(f"F{sum_row+2}")].alignment = right
    money(ws[cell(f"F{sum_row+2}")])

    # PPN
    ws[cell(f"E{sum_row+3}")].value = f"PPN {int(ppn_rate*100)}%"
    ws[cell(f"E{sum_row+3}")].alignment = right
    ws[cell(f"F{sum_row+3}")].value = ppn
    ws[cell(f"F{sum_row+3}")].alignment = right
    money(ws[cell(f"F{sum_row+3}")])

    # Deposit
    ws[cell(f"E{sum_row+4}")].value = "Less: Deposit"
    ws[cell(f"E{sum_row+4}")].alignment = right
    ws[cell(f"F{sum_row+4}")].value = deposit
    ws[cell(f"F{sum_row+4}")].alignment = right
    money(ws[cell(f"F{sum_row+4}")])

    # Balance Due
    ws[cell(f"E{sum_row+5}")].value = "Balance Due"
    ws[cell(f"E{sum_row+5}")].alignment = right
    ws[cell(f"E{sum_row+5}")].font = font_bold
    ws[cell(f"F{sum_row+5}")].value = balance
    ws[cell(f"F{sum_row+5}")].alignment = right
    ws[cell(f"F{sum_row+5}")].font = font_bold
    money(ws[cell(f"F{sum_row+5}")])

    _set_border(ws, sum_row, 5 + SHIFT, sum_row + 5, 6 + SHIFT, border)

    # ========= Payment block kiri bawah =========
    pay_row = sum_row
    ws.merge_cells(f"{cell(f'A{pay_row}')}" + ":" + f"{cell(f'D{pay_row}')}")
    ws[cell(f"A{pay_row}")].value = "Please Transfer Full Amount to:"
    ws[cell(f"A{pay_row}")].font = font_bold
    ws[cell(f"A{pay_row}")].alignment = left_mid

    ws[cell(f"A{pay_row+1}")].value = "Beneficiary :"
    ws[cell(f"A{pay_row+2}")].value = "Bank Name :"
    ws[cell(f"A{pay_row+3}")].value = "Branch :"
    ws[cell(f"A{pay_row+4}")].value = "IDR Acct :"

    ws.merge_cells(f"{cell(f'B{pay_row+1}')}:{cell(f'D{pay_row+1}')}")
    ws[cell(f"B{pay_row+1}")].value = payment["beneficiary"]

    ws.merge_cells(f"{cell(f'B{pay_row+2}')}:{cell(f'D{pay_row+2}')}")
    ws[cell(f"B{pay_row+2}")].value = payment["bank_name"]

    ws.merge_cells(f"{cell(f'B{pay_row+3}')}:{cell(f'D{pay_row+3}')}")
    ws[cell(f"B{pay_row+3}")].value = payment["branch"]

    ws.merge_cells(f"{cell(f'B{pay_row+4}')}:{cell(f'D{pay_row+4}')}")
    ws[cell(f"B{pay_row+4}")].value = payment["idr_acct"]

    _set_border(ws, pay_row, 1 + SHIFT, pay_row + 4, 4 + SHIFT, border)

    # Font default area
    for row in ws.iter_rows(min_row=1, max_row=pay_row + 6, min_col=1 + SHIFT, max_col=6 + SHIFT):
        for c in row:
            if c.font is None or c.font == Font():
                c.font = font_normal

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
