from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, Image, HRFlowable, KeepTogether
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from io import BytesIO
from datetime import datetime
import os

STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")

# ── Palette ───────────────────────────────────────────────────────────────────
DARK      = colors.HexColor("#1A1A1A")
CHARCOAL  = colors.HexColor("#2C2C2C")
GOLD      = colors.HexColor("#F5C518")
GOLD_DARK = colors.HexColor("#D4A800")
RED       = colors.HexColor("#CC0000")
WHITE     = colors.white
GREY_LT   = colors.HexColor("#F7F7F7")
GREY_MID  = colors.HexColor("#EBEBEB")
GREY_BD   = colors.HexColor("#CCCCCC")
GREY_DARK = colors.HexColor("#888888")


def _fmt_date(iso_date: str) -> str:
    try:
        return datetime.strptime(iso_date, "%Y-%m-%d").strftime("%d %b %Y")
    except Exception:
        return iso_date


def _img(filename, w, h):
    path = os.path.join(STATIC_DIR, filename)
    if os.path.exists(path):
        return Image(path, width=w, height=h)
    return None


def _make_photo_cell(pic_ref: str, max_w: float, max_h: float):
    """Return an Image flowable if the photo exists, otherwise a placeholder paragraph."""
    if not pic_ref:
        return _p("", size=8)
    path = os.path.join(STATIC_DIR, "uploads", pic_ref)
    if not os.path.exists(path):
        return _p(pic_ref, size=7, color=GREY_DARK, align=TA_CENTER)
    try:
        img = Image(path)
        iw, ih = img.imageWidth, img.imageHeight
        ratio = min(max_w / iw, max_h / ih)
        return Image(path, width=iw * ratio, height=ih * ratio)
    except Exception:
        return _p(pic_ref, size=7, color=GREY_DARK, align=TA_CENTER)


def _p(text, font="Helvetica", size=9, color=DARK, align=TA_LEFT, bold=False):
    if bold:
        font = font.replace("Helvetica", "Helvetica-Bold")
    return Paragraph(
        text,
        ParagraphStyle(
            f"_p_{hash(text+font+str(size))}",
            fontName=font,
            fontSize=size,
            textColor=color,
            alignment=align,
            leading=size * 1.35,
            spaceAfter=0,
            spaceBefore=0,
        )
    )


def generate_pdf(transfer: dict) -> bytes:
    buf = BytesIO()
    W_page, H_page = A4
    margin = 1.5 * cm
    W = W_page - 2 * margin

    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=margin, rightMargin=margin,
        topMargin=1.2 * cm, bottomMargin=2 * cm,
        title=transfer["collect_no"],
    )

    elements = []

    # ── 1. HEADER ─────────────────────────────────────────────────────────────
    sacoor_img = _img("sacoor_logo.png", 2.6 * cm, 1.55 * cm)
    dhl_img    = _img("dhl_logo.png",    2.4 * cm, 1.1 * cm)

    sacoor_cell = sacoor_img or _p("SACOOR<br/>brothers", size=11, bold=True, align=TA_CENTER)
    dhl_cell    = dhl_img    or _p("DHL", size=18, bold=True, color=RED, align=TA_CENTER)

    # Collection No inner box
    ref_inner = Table(
        [
            [_p("Collection No", size=7, bold=True, color=CHARCOAL, align=TA_CENTER)],
            [_p(transfer["collect_no"], size=11, bold=True, color=RED, align=TA_CENTER)],
        ],
        colWidths=[3.8 * cm],
    )
    ref_inner.setStyle(TableStyle([
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        ("LINEBELOW",     (0, 0), (-1, 0),  0.5, GREY_BD),
    ]))

    header_data = [[
        sacoor_cell,
        dhl_cell,
        _p("COMMERCIAL RETURNS", size=17, bold=True, color=DARK, align=TA_CENTER),
        ref_inner,
    ]]
    col_w = [3.0*cm, 3.0*cm, W - 9.8*cm, 3.8*cm]
    header = Table(header_data, colWidths=col_w)
    header.setStyle(TableStyle([
        ("BOX",           (0, 0), (-1, -1), 1.2, DARK),
        ("LINEBEFORE",    (1, 0), (1, 0),   0.6, GREY_BD),
        ("LINEBEFORE",    (2, 0), (2, 0),   0.6, GREY_BD),
        ("LINEBEFORE",    (3, 0), (3, 0),   1.2, DARK),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING",    (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
    ]))
    elements.append(header)

    # Sub-note
    note = Table([[
        _p(
            "This form to be <b>printed</b> as acknowledgement of receipts "
            "of the returns from the Boutiques/Outlets",
            size=7, color=GREY_DARK, align=TA_CENTER
        )
    ]], colWidths=[W])
    note.setStyle(TableStyle([
        ("BOX",           (0, 0), (-1, -1), 0.5, GREY_BD),
        ("BACKGROUND",    (0, 0), (-1, -1), GREY_LT),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    elements.append(note)

    # ── 2. INFO GRID ──────────────────────────────────────────────────────────
    # Label row (dark background)
    label_s   = dict(size=7,  bold=True, color=WHITE, align=TA_CENTER)
    value_s   = dict(size=9,  color=RED,  align=TA_CENTER)
    value_bk  = dict(size=9,  color=DARK, align=TA_CENTER)

    cw = [W * p for p in [0.13, 0.22, 0.13, 0.25, 0.15, 0.06, 0.06]]

    labels = [
        _p("Collection Date<br/>(dd/mm/yy)", **label_s),
        _p("Collect FROM",                   **label_s),
        _p("Delivery Date<br/>(dd/mm/yy)",   **label_s),
        _p("Deliver TO",                     **label_s),
        _p("Total Railing",                  **label_s),
        _p("Total<br/>Pcs",                  **label_s),
        _p("Total<br/>Ctn",                  **label_s),
    ]
    rln = str(transfer.get("total_rln") or 0)
    values = [
        _p(_fmt_date(transfer["collection_date"]), **value_s),
        _p(transfer["from_store_name"],            **value_s),
        _p(_fmt_date(transfer["delivery_date"]),   **value_s),
        _p(transfer["to_store_name"],              **value_s),
        _p(rln if rln != "0" else "",              **value_bk),
        _p(str(transfer["total_pcs"]),             **value_s),
        _p(str(transfer["total_ctn"]),             **value_s),
    ]

    info_table = Table([labels, values], colWidths=cw)
    info_table.setStyle(TableStyle([
        ("BOX",           (0, 0), (-1, -1), 1,   DARK),
        ("INNERGRID",     (0, 0), (-1, -1), 0.4, GREY_BD),
        # Label row — dark
        ("BACKGROUND",    (0, 0), (-1, 0),  CHARCOAL),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        # Accent separators on last 3 cols
        ("LINEBEFORE",    (4, 0), (4, -1),  1.2, GREY_BD),
    ]))
    elements.append(info_table)

    # ── 3. DESTINATION ADDRESS ────────────────────────────────────────────────
    addr = (f"{transfer['to_store_name']}, {transfer['to_address']}, "
            f"{transfer['to_city']}, {transfer['to_country']}")
    addr_row = Table([[
        _p("DELIVER TO:", size=7, bold=True, color=GREY_DARK),
        _p(addr, size=8, color=DARK),
    ]], colWidths=[2.2*cm, W - 2.2*cm])
    addr_row.setStyle(TableStyle([
        ("BOX",           (0, 0), (-1, -1), 0.5, GREY_BD),
        ("BACKGROUND",    (0, 0), (-1, -1), GREY_LT),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBEFORE",    (1, 0), (1, 0),   0.5, GREY_BD),
    ]))
    elements.append(addr_row)
    elements.append(Spacer(1, 0.4 * cm))

    # ── 4. ITEMS TABLE ────────────────────────────────────────────────────────
    icw = [W * 0.28, W * 0.45, W * 0.13, W * 0.14]

    item_rows = [[
        _p("TG",          size=9, bold=True, color=WHITE, align=TA_CENTER),
        _p("DESCRIPTION", size=9, bold=True, color=WHITE, align=TA_CENTER),
        _p("UOM",         size=9, bold=True, color=WHITE, align=TA_CENTER),
        _p("QTY",         size=9, bold=True, color=WHITE, align=TA_CENTER),
    ]]

    lines = transfer.get("lines", [])
    for line in lines:
        item_rows.append([
            _p(line["tg_number"],                      size=9, color=DARK,  align=TA_CENTER),
            _p(line.get("description", "Stock Rotation by Email"), size=9, color=DARK, align=TA_CENTER),
            _p(line.get("uom", "Pcs"),                 size=9, color=RED,   align=TA_CENTER),
            _p(str(line["qty"]),                       size=9, color=DARK,  align=TA_CENTER),
        ])

    min_rows = max(12, len(lines) + 2)
    while len(item_rows) < min_rows + 1:  # +1 for header
        item_rows.append([_p("", size=9)] * 4)

    items_tbl = Table(item_rows, colWidths=icw, repeatRows=1)
    items_tbl.setStyle(TableStyle([
        # Outer border
        ("BOX",           (0, 0), (-1, -1), 1,   DARK),
        ("INNERGRID",     (0, 0), (-1, -1), 0.4, GREY_BD),
        # Header row — dark (same as info table)
        ("BACKGROUND",    (0, 0), (-1, 0),  CHARCOAL),
        ("LINEBELOW",     (0, 0), (-1, 0),  1.5, DARK),
        # Zebra rows
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [WHITE, GREY_LT]),
        # Padding
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elements.append(items_tbl)
    elements.append(Spacer(1, 0.6 * cm))

    # ── 5. SIGNATURE AREA ─────────────────────────────────────────────────────
    def sig_block(label):
        col_w = (W / 3) - 0.8 * cm
        return Table([
            [_p(label, size=8, bold=True, color=CHARCOAL)],
            [Spacer(1, 1.2 * cm)],
            [HRFlowable(width="90%", thickness=0.5, color=GREY_BD, spaceAfter=2)],
            [_p("Name / Signature", size=7, color=GREY_DARK)],
            [Spacer(1, 0.5 * cm)],
            [HRFlowable(width="60%", thickness=0.5, color=GREY_BD, spaceAfter=2)],
            [_p("Date", size=7, color=GREY_DARK)],
        ], colWidths=[col_w])

    sig_table = Table(
        [[sig_block("Sender:"), sig_block("Transport:"), sig_block("Receiver:")]],
        colWidths=[W / 3, W / 3, W / 3],
    )
    sig_table.setStyle(TableStyle([
        ("BOX",          (0, 0), (-1, -1), 0.8, GREY_BD),
        ("LINEBEFORE",   (1, 0), (1, 0),   0.8, GREY_BD),
        ("LINEBEFORE",   (2, 0), (2, 0),   0.8, GREY_BD),
        ("TOPPADDING",   (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 10),
        ("LEFTPADDING",  (0, 0), (-1, -1), 14),
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
    ]))
    elements.append(sig_table)

    # ── 6. FOOTER NOTE ────────────────────────────────────────────────────────
    elements.append(Spacer(1, 0.4 * cm))
    elements.append(HRFlowable(width="100%", thickness=0.4, color=GREY_BD))
    elements.append(Spacer(1, 0.15 * cm))
    footer_data = [[
        _p(f"Generated: {datetime.utcnow().strftime('%d %b %Y %H:%M')} UTC",
           size=6.5, color=GREY_DARK),
        _p(f"Ref: {transfer['collect_no']}",
           size=6.5, color=GREY_DARK, align=TA_RIGHT),
    ]]
    footer = Table(footer_data, colWidths=[W / 2, W / 2])
    footer.setStyle(TableStyle([
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    elements.append(footer)

    doc.build(elements)
    return buf.getvalue()


# ── DECORATION FORM ───────────────────────────────────────────────────────────

def generate_decoration_pdf(transfer: dict) -> bytes:
    """Same structure as commercial but items table has Description/UOM/QTY/Picture."""
    buf = BytesIO()
    W_page, H_page = A4
    margin = 1.5 * cm
    W = W_page - 2 * margin

    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=margin, rightMargin=margin,
        topMargin=1.2 * cm, bottomMargin=2 * cm,
        title=transfer["collect_no"],
    )
    elements = []

    # ── Header (identical to commercial) ─────────────────────────────────────
    sacoor_img = _img("sacoor_logo.png", 2.6 * cm, 1.55 * cm)
    dhl_img    = _img("dhl_logo.png",    2.4 * cm, 1.1 * cm)
    sacoor_cell = sacoor_img or _p("SACOOR<br/>brothers", size=11, bold=True, align=TA_CENTER)
    dhl_cell    = dhl_img    or _p("DHL", size=18, bold=True, color=RED, align=TA_CENTER)

    ref_inner = Table(
        [[_p("Collection No", size=7, bold=True, color=CHARCOAL, align=TA_CENTER)],
         [_p(transfer["collect_no"], size=11, bold=True, color=RED, align=TA_CENTER)]],
        colWidths=[3.8 * cm],
    )
    ref_inner.setStyle(TableStyle([
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        ("LINEBELOW",     (0, 0), (-1, 0),  0.5, GREY_BD),
    ]))

    # "DECORATION RETURNS" title
    header_data = [[
        sacoor_cell, dhl_cell,
        _p("DECORATION RETURNS", size=17, bold=True, color=DARK, align=TA_CENTER),
        ref_inner,
    ]]
    col_w = [3.0*cm, 3.0*cm, W - 9.8*cm, 3.8*cm]
    header = Table(header_data, colWidths=col_w)
    header.setStyle(TableStyle([
        ("BOX",           (0, 0), (-1, -1), 1.2, DARK),
        ("LINEBEFORE",    (1, 0), (1, 0),   0.6, GREY_BD),
        ("LINEBEFORE",    (2, 0), (2, 0),   0.6, GREY_BD),
        ("LINEBEFORE",    (3, 0), (3, 0),   1.2, DARK),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING",    (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
    ]))
    elements.append(header)

    note = Table([[_p(
        "This form to be <b>printed</b> as acknowledgement of receipts "
        "of the returns from the Boutiques/Outlets",
        size=7, color=GREY_DARK, align=TA_CENTER
    )]], colWidths=[W])
    note.setStyle(TableStyle([
        ("BOX",           (0, 0), (-1, -1), 0.5, GREY_BD),
        ("BACKGROUND",    (0, 0), (-1, -1), GREY_LT),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    elements.append(note)

    # ── Info grid ─────────────────────────────────────────────────────────────
    label_s = dict(size=7, bold=True, color=WHITE, align=TA_CENTER)
    value_s = dict(size=9, color=RED,  align=TA_CENTER)
    cw = [W * p for p in [0.13, 0.22, 0.13, 0.25, 0.15, 0.06, 0.06]]

    labels = [
        _p("Collection Date<br/>(dd/mm/yy)", **label_s),
        _p("Collect FROM",                   **label_s),
        _p("Delivery Date<br/>(dd/mm/yy)",   **label_s),
        _p("Deliver TO",                     **label_s),
        _p("Total Railing",                  **label_s),
        _p("Total<br/>Pcs",                  **label_s),
        _p("Total<br/>Ctn",                  **label_s),
    ]
    rln = str(transfer.get("total_rln") or 0)
    values = [
        _p(_fmt_date(transfer["collection_date"]), **value_s),
        _p(transfer["from_store_name"],            **value_s),
        _p(_fmt_date(transfer["delivery_date"]),   **value_s),
        _p(transfer["to_store_name"],              **value_s),
        _p(rln if rln != "0" else "",              **value_s),
        _p(str(transfer["total_pcs"]),             **value_s),
        _p(str(transfer["total_ctn"]),             **value_s),
    ]
    info_table = Table([labels, values], colWidths=cw)
    info_table.setStyle(TableStyle([
        ("BOX",           (0, 0), (-1, -1), 1,   DARK),
        ("INNERGRID",     (0, 0), (-1, -1), 0.4, GREY_BD),
        ("BACKGROUND",    (0, 0), (-1, 0),  CHARCOAL),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBEFORE",    (4, 0), (4, -1),  1.2, GREY_BD),
    ]))
    elements.append(info_table)

    # ── Address ───────────────────────────────────────────────────────────────
    addr = (f"{transfer['to_store_name']}, {transfer['to_address']}, {transfer['to_city']}, {transfer['to_country']}")
    addr_row = Table([[
        _p("DELIVER TO:", size=7, bold=True, color=GREY_DARK),
        _p(addr, size=8, color=DARK),
    ]], colWidths=[2.2*cm, W - 2.2*cm])
    addr_row.setStyle(TableStyle([
        ("BOX",           (0, 0), (-1, -1), 0.5, GREY_BD),
        ("BACKGROUND",    (0, 0), (-1, -1), GREY_LT),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBEFORE",    (1, 0), (1, 0),   0.5, GREY_BD),
    ]))
    elements.append(addr_row)
    elements.append(Spacer(1, 0.4 * cm))

    # ── Decoration items table (Description / UOM / QTY / Picture) ───────────
    icw = [W * 0.38, W * 0.12, W * 0.12, W * 0.38]

    item_rows = [[
        _p("DESCRIPTION", size=9, bold=True, color=WHITE, align=TA_CENTER),
        _p("UOM",         size=9, bold=True, color=WHITE, align=TA_CENTER),
        _p("QTY",         size=9, bold=True, color=WHITE, align=TA_CENTER),
        _p("PICTURE",     size=9, bold=True, color=WHITE, align=TA_CENTER),
    ]]

    lines = transfer.get("lines", [])
    for line in lines:
        pic_ref = line.get("picture_ref", "")
        pic_cell = _make_photo_cell(pic_ref, max_w=3.5*cm, max_h=2.8*cm)
        item_rows.append([
            _p(line.get("description", ""), size=9, color=DARK, align=TA_LEFT),
            _p(line.get("uom", "Pcs"),      size=9, color=RED,  align=TA_CENTER),
            _p(str(line["qty"]),            size=9, color=DARK, align=TA_CENTER),
            pic_cell,
        ])

    # Decoration items need taller rows for picture area
    min_rows = max(10, len(lines) + 2)
    while len(item_rows) < min_rows + 1:
        item_rows.append([_p("", size=9)] * 4)

    items_tbl = Table(item_rows, colWidths=icw, repeatRows=1,
                      rowHeights=[None] + [3.2*cm] * (len(item_rows) - 1))
    items_tbl.setStyle(TableStyle([
        ("BOX",           (0, 0), (-1, -1), 1,   DARK),
        ("INNERGRID",     (0, 0), (-1, -1), 0.4, GREY_BD),
        ("BACKGROUND",    (0, 0), (-1, 0),  CHARCOAL),
        ("LINEBELOW",     (0, 0), (-1, 0),  1.5, DARK),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [WHITE, GREY_LT]),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elements.append(items_tbl)
    elements.append(Spacer(1, 0.6 * cm))

    # ── Signature ─────────────────────────────────────────────────────────────
    def sig_block(label):
        return Table([
            [_p(label, size=8, bold=True, color=CHARCOAL)],
            [Spacer(1, 1.2 * cm)],
            [HRFlowable(width="90%", thickness=0.5, color=GREY_BD, spaceAfter=2)],
            [_p("Name / Signature", size=7, color=GREY_DARK)],
            [Spacer(1, 0.5 * cm)],
            [HRFlowable(width="60%", thickness=0.5, color=GREY_BD, spaceAfter=2)],
            [_p("Date", size=7, color=GREY_DARK)],
        ], colWidths=[(W / 2) - 1 * cm])

    sig_table = Table(
        [[sig_block("Prepared by:"), sig_block("Received by:")]],
        colWidths=[W / 2, W / 2],
    )
    sig_table.setStyle(TableStyle([
        ("BOX",          (0, 0), (-1, -1), 0.8, GREY_BD),
        ("LINEBEFORE",   (1, 0), (1, 0),   0.8, GREY_BD),
        ("TOPPADDING",   (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 10),
        ("LEFTPADDING",  (0, 0), (-1, -1), 14),
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
    ]))
    elements.append(sig_table)

    # ── Footer ────────────────────────────────────────────────────────────────
    elements.append(Spacer(1, 0.4 * cm))
    elements.append(HRFlowable(width="100%", thickness=0.4, color=GREY_BD))
    elements.append(Spacer(1, 0.15 * cm))
    footer = Table([[
        _p(f"Generated: {datetime.utcnow().strftime('%d %b %Y %H:%M')} UTC", size=6.5, color=GREY_DARK),
        _p(f"Ref: {transfer['collect_no']}", size=6.5, color=GREY_DARK, align=TA_RIGHT),
    ]], colWidths=[W / 2, W / 2])
    footer.setStyle(TableStyle([("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
    elements.append(footer)

    doc.build(elements)
    return buf.getvalue()
