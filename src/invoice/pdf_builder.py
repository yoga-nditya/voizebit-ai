import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

from config_new import FILES_DIR


def create_invoice_pdf(inv: dict, fname_base: str) -> str:
    try:
        folder = str(FILES_DIR)
    except Exception:
        folder = "static/files"
    os.makedirs(folder, exist_ok=True)

    out_path = os.path.join(folder, f"{fname_base}.pdf")
    c = canvas.Canvas(out_path, pagesize=A4)
    width, height = A4

    def txt(x, y, s, size=9, bold=False):
        c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
        c.drawString(x, y, s or "")

    def rtxt(x, y, s, size=9, bold=False):
        c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
        c.drawRightString(x, y, s or "")

    def rect(x, y, w, h, lw=1):
        c.setLineWidth(lw)
        c.rect(x, y, w, h)

    def vline(x, y1, y2, lw=0.6):
        c.setLineWidth(lw)
        c.line(x, y1, x, y2)

    def fmt_id(n: int) -> str:
        try:
            return f"{int(n):,}".replace(",", ".")
        except:
            return str(n)

    invoice_no = inv.get("invoice_no") or ""
    inv_date = inv.get("invoice_date") or ""
    bill_to = inv.get("bill_to") or {}
    ship_to = inv.get("ship_to") or {}
    phone = inv.get("phone") or ""
    fax = inv.get("fax") or ""
    attn = inv.get("attn") or "Accounting / Finance"
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

    left_margin = 40
    table_x = left_margin
    table_w = width - 80

    w_qty = 45
    w_unit = 35
    w_date = 70
    w_desc = 220
    w_price = 70
    w_amt = table_w - (w_qty + w_unit + w_date + w_desc + w_price)

    x_qty = table_x
    x_unit = x_qty + w_qty
    x_date = x_unit + w_unit
    x_desc = x_date + w_date
    x_price = x_desc + w_desc
    x_amt = x_price + w_price
    x_end = table_x + table_w

    y = height - 50

    txt(table_x, y, "Bill To:", 10, True)
    txt(table_x + table_w * 0.55, y, "Ship To:", 10, True)
    y -= 14

    bt_lines = [bill_to.get("name", ""), bill_to.get("address", ""), bill_to.get("address2", "")]
    st_lines = [ship_to.get("name", ""), ship_to.get("address", ""), ship_to.get("address2", "")]
    bt_lines = [s for s in bt_lines if (s or "").strip()]
    st_lines = [s for s in st_lines if (s or "").strip()]

    yy = y
    for line in bt_lines[:3]:
        txt(table_x, yy, str(line), 9, False)
        yy -= 12

    yy2 = y
    for line in st_lines[:3]:
        txt(table_x + table_w * 0.55, yy2, str(line), 9, False)
        yy2 -= 12

    rtxt(x_end, height - 62, invoice_no, 9, False)
    txt(x_end - 120, height - 62, "Invoice", 9, True)
    rtxt(x_end, height - 76, inv_date, 9, False)
    txt(x_end - 120, height - 76, "Date", 9, True)
    rtxt(x_end, height - 90, no_surat_jalan, 9, False)
    txt(x_end - 120, height - 90, "No. Surat Jalan", 9, True)

    y = min(yy, yy2) - 8
    txt(table_x, y, "Phone:", 9, True)
    txt(table_x + 50, y, phone, 9, False)
    txt(table_x + table_w * 0.55, y, "Fax:", 9, True)
    txt(table_x + table_w * 0.55 + 35, y, fax, 9, False)
    y -= 14
    txt(table_x, y, "Attn :", 9, True)
    txt(table_x + 45, y, attn, 9, False)

    y -= 28
    ref_box_top = y
    ref_box_h = 40
    rect(table_x, ref_box_top - ref_box_h, table_w, ref_box_h, lw=1)

    vline(table_x + table_w * 0.25, ref_box_top - ref_box_h, ref_box_top, lw=0.6)
    vline(table_x + table_w * 0.55, ref_box_top - ref_box_h, ref_box_top, lw=0.6)
    vline(table_x + table_w * 0.78, ref_box_top - ref_box_h, ref_box_top, lw=0.6)

    txt(table_x + 10, ref_box_top - 14, "Ref No.", 9, True)
    txt(table_x + table_w * 0.25 + 10, ref_box_top - 14, "Sales Person", 9, True)
    txt(table_x + table_w * 0.55 + 10, ref_box_top - 14, "Ship Via", 9, True)
    txt(table_x + table_w * 0.78 + 10, ref_box_top - 14, "Ship Date", 9, True)

    txt(table_x + 10, ref_box_top - 30, ref_no, 9, False)
    txt(table_x + table_w * 0.25 + 10, ref_box_top - 30, sales_person, 9, False)
    txt(table_x + table_w * 0.55 + 10, ref_box_top - 30, ship_via, 9, False)
    txt(table_x + table_w * 0.78 + 10, ref_box_top - 30, ship_date, 9, False)

    txt(x_amt - 5, ref_box_top - ref_box_h - 14, "Terms", 9, True)
    rtxt(x_end, ref_box_top - ref_box_h - 14, terms, 9, False)

    y = ref_box_top - ref_box_h - 28
    table_top = y
    table_h = 220
    rect(table_x, table_top - table_h, table_w, table_h, lw=1)

    vline(x_unit, table_top - table_h, table_top, lw=0.6)
    vline(x_date, table_top - table_h, table_top, lw=0.6)
    vline(x_desc, table_top - table_h, table_top, lw=0.6)
    vline(x_price, table_top - table_h, table_top, lw=0.6)
    vline(x_amt, table_top - table_h, table_top, lw=0.6)

    header_y = table_top - 16
    txt(x_qty + 4, header_y, "Qty", 9, True)
    txt(x_date + 4, header_y, "Date", 9, True)
    txt(x_desc + 4, header_y, "Description", 9, True)
    txt(x_price + 4, header_y, "Price", 9, True)
    txt(x_amt + 4, header_y, "Amount (IDR)", 9, True)

    row_y = header_y - 18
    subtotal = 0
    max_rows = 10
    for idx in range(max_rows):
        if idx < len(items):
            it = items[idx]
            qty = it.get("qty") or 0
            unit = it.get("unit") or "Kg"
            dt = it.get("date") or inv_date
            desc = it.get("description") or ""
            price = int(it.get("price") or 0)
            amount = int(round(float(qty) * price))
            subtotal += amount

            txt(x_qty + 4, row_y, str(qty), 9, False)
            txt(x_unit + 4, row_y, str(unit), 9, False)
            txt(x_date + 4, row_y, str(dt), 9, False)
            txt(x_desc + 4, row_y, str(desc)[:45], 9, False)
            rtxt(x_price + w_price - 4, row_y, fmt_id(price), 9, False)
            rtxt(x_end - 4, row_y, fmt_id(amount), 9, False)
        row_y -= 16

    total_before_ppn = subtotal + freight
    ppn = int(round(total_before_ppn * ppn_rate))
    balance = total_before_ppn + ppn - deposit

    base_y = table_top - table_h - 20

    txt(table_x, base_y, "Please Transfer Full Amount to:", 9, True)
    txt(table_x, base_y - 14, f"Beneficiary : {payment.get('beneficiary','')}", 9, False)
    txt(table_x, base_y - 28, f"Bank Name   : {payment.get('bank_name','')}", 9, False)
    txt(table_x, base_y - 42, f"Branch      : {payment.get('branch','')}", 9, False)
    txt(table_x, base_y - 56, f"IDR Acct    : {payment.get('idr_acct','')}", 9, False)

    box_w = w_price + w_amt
    box_x = x_price
    box_y_top = base_y + 8
    line_h = 14
    labels = [
        ("Total", subtotal),
        ("Freight", freight),
        ("Total", total_before_ppn),
        (f"PPN {int(ppn_rate*100)}%", ppn),
        ("Less: Deposit", deposit),
        ("Balance Due", balance),
    ]
    box_h = line_h * len(labels) + 6
    rect(box_x, box_y_top - box_h, box_w, box_h, lw=1)

    yy = box_y_top - 16
    for (lab, val) in labels:
        txt(box_x + 6, yy, lab, 9, True if lab in ("Total", "Balance Due") else False)
        rtxt(box_x + box_w - 6, yy, fmt_id(val), 9, True if lab in ("Balance Due",) else False)
        yy -= line_h

    sig_top = box_y_top - box_h - 30
    sig_w = box_w
    sig_h = 80
    rect(box_x, sig_top - sig_h, sig_w, sig_h, lw=1)
    txt(box_x + 10, sig_top - 14, "PT. Sarana Trans Bersama Jaya", 9, True)

    txt(box_x + 10, sig_top - sig_h - 14, "Please kindly fax to our attention upon receipt", 9, False)

    c.showPage()
    c.save()
    return f"{fname_base}.pdf"
