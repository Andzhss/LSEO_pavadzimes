from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage, Flowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_RIGHT, TA_LEFT, TA_CENTER
import io
import os

# --- Krāsu definīcijas (sky blue theme) ---
THEME_DARK   = colors.HexColor("#1A6FA8")   # tumšzils — virsraksti, galvenes
THEME_MID    = colors.HexColor("#4DA6E8")   # vidējs zils — līnijas, apmales
THEME_LIGHT  = colors.HexColor("#D0EAF8")   # gaiši zils — rindu fons
THEME_WHITE  = colors.white
TEXT_COLOR   = colors.HexColor("#1A1A2E")   # gandrīz melns teksts

# --- Fontu ielāde ---
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
FONTS_DIR = os.path.join(CURRENT_DIR, "fonts")

try:
    pdfmetrics.registerFont(TTFont('Montserrat', os.path.join(FONTS_DIR, "Montserrat-Regular.ttf")))
    REGULAR_FONT = 'Montserrat'
except Exception as e:
    print(f"Neizdevās ielādēt Montserrat: {e}")
    REGULAR_FONT = 'Helvetica'

try:
    pdfmetrics.registerFont(TTFont('Montserrat-Bold', os.path.join(FONTS_DIR, "Montserrat-Bold.ttf")))
    BOLD_FONT = 'Montserrat-Bold'
except:
    BOLD_FONT = 'Helvetica-Bold'

try:
    pdfmetrics.registerFont(TTFont('Montserrat-Italic', os.path.join(FONTS_DIR, "Montserrat-Italic.ttf")))
    ITALIC_FONT = 'Montserrat-Italic'
except:
    ITALIC_FONT = 'Helvetica-Oblique'

try:
    pdfmetrics.registerFont(TTFont('Montserrat-BoldItalic', os.path.join(FONTS_DIR, "Montserrat-BoldItalic.ttf")))
    BOLD_ITALIC_FONT = 'Montserrat-BoldItalic'
except:
    BOLD_ITALIC_FONT = 'Helvetica-BoldOblique'

try:
    from reportlab.pdfbase.pdfmetrics import registerFontFamily
    registerFontFamily('Montserrat', normal='Montserrat', bold='Montserrat-Bold',
                       italic='Montserrat-Italic', boldItalic='Montserrat-BoldItalic')
except:
    pass

# --- Tulkojumi PDF dokumentam ---
PDF_TRANSLATIONS = {
    "lv": {
        "client":          "KLIENTS",
        "sender":          "PIEGĀDĀTĀJS",
        "receiver":        "Saņēmējs",
        "customer":        "Pasūtītājs",
        "sender_label":    "Nosūtītājs",
        "address":         "Adrese",
        "reg_no":          "Reģ. Nr.",
        "vat_no":          "PVN Nr.",
        "phone":           "Tālrunis",
        "bank":            "Banka",
        "swift":           "SWIFT/BIC",
        "account":         "Bankas konta numurs",
        "due_date":        "Apmaksāt līdz",
        "date":            "Datums",
        "col_name":        "NOSAUKUMS",
        "col_unit":        "Mērvienība",
        "col_qty":         "DAUDZUMS",
        "col_price":       "CENA (EUR)",
        "col_total":       "KOPĀ (EUR)",
        "total_label":     "KOPĀ",
        "total_no_vat":    "KOPĀ (bez PVN un atlaides)",
        "discount":        "Atlaides apjoms",
        "total_discount":  "Kopā ar atlaidi (bez PVN)",
        "vat":             "PVN",
        "grand_total":     "Kopumā",
        "advance_payable": "APMAKSĀJAMAIS AVANSS",
        "words_prefix":    "Vārdiem: ",
        "extra_info":      "Papildus informācija:",
        "prep_invoice":    "Pavadzīmi sagatavoja:",
        "prep_receipt":    "Rēķinu sagatavoja:",
        "prep_advance":    "Avansa rēķinu sagatavoja:",
        "recv_invoice":    "Pavadzīmi saņēma:",
        "recv_receipt":    "Rēķinu saņēma:",
        "recv_advance":    "Avansa rēķinu saņēma:",
    },
    "en": {
        "client":          "CLIENT",
        "sender":          "SUPPLIER",
        "receiver":        "Receiver",
        "customer":        "Customer",
        "sender_label":    "Sender",
        "address":         "Address",
        "reg_no":          "Reg. No.",
        "vat_no":          "VAT No.",
        "phone":           "Phone",
        "bank":            "Bank",
        "swift":           "SWIFT/BIC",
        "account":         "Bank account number",
        "due_date":        "Due date",
        "date":            "Date",
        "col_name":        "DESCRIPTION",
        "col_unit":        "Unit",
        "col_qty":         "QUANTITY",
        "col_price":       "PRICE (EUR)",
        "col_total":       "TOTAL (EUR)",
        "total_label":     "SUBTOTAL",
        "total_no_vat":    "SUBTOTAL (excl. VAT and discount)",
        "discount":        "Discount amount",
        "total_discount":  "Total with discount (excl. VAT)",
        "vat":             "VAT",
        "grand_total":     "Total",
        "advance_payable": "ADVANCE PAYABLE",
        "words_prefix":    "In words: ",
        "extra_info":      "Additional information:",
        "prep_invoice":    "Invoice prepared by:",
        "prep_receipt":    "Receipt prepared by:",
        "prep_advance":    "Advance invoice prepared by:",
        "recv_invoice":    "Invoice received by:",
        "recv_receipt":    "Receipt received by:",
        "recv_advance":    "Advance invoice received by:",
    }
}

# --- Horizontālā līnija (tumšzila) ---
class HorizontalLine(Flowable):
    def __init__(self, width=170*mm, color=THEME_DARK, thickness=1.0):
        Flowable.__init__(self)
        self.width = width
        self.color = color
        self.thickness = thickness

    def draw(self):
        self.canv.setStrokeColor(self.color)
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, 0, self.width, 0)

# --- Krāsains fona bloks galvenei ---
class ColorBar(Flowable):
    def __init__(self, width=170*mm, height=1.5*mm, color=THEME_MID):
        Flowable.__init__(self)
        self.width = width
        self.height = height
        self.color = color

    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.rect(0, 0, self.width, self.height, fill=1, stroke=0)

def fmt_curr(val):
    return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")

def generate_pdf(data):
    lang = data.get('lang', 'lv')
    tr = PDF_TRANSLATIONS.get(lang, PDF_TRANSLATIONS['lv'])

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=20*mm, leftMargin=20*mm,
                            topMargin=15*mm, bottomMargin=15*mm)

    styles = getSampleStyleSheet()

    # --- Stili ---
    style_normal = ParagraphStyle('CustomNormal', parent=styles['Normal'],
        fontName=REGULAR_FONT, fontSize=9.5, leading=13, textColor=TEXT_COLOR)
    style_bold = ParagraphStyle('CustomBold', parent=styles['Normal'],
        fontName=BOLD_FONT, fontSize=9.5, leading=13, textColor=TEXT_COLOR)
    style_italic = ParagraphStyle('CustomItalic', parent=styles['Normal'],
        fontName=ITALIC_FONT, fontSize=9.5, leading=13, textColor=TEXT_COLOR)
    style_small = ParagraphStyle('CustomSmall', parent=styles['Normal'],
        fontName=REGULAR_FONT, fontSize=8.5, leading=11, textColor=colors.HexColor("#555577"))

    style_header_title = ParagraphStyle('HeaderTitle', parent=styles['Normal'],
        fontName=BOLD_FONT, fontSize=18, alignment=TA_RIGHT,
        leading=22, textColor=THEME_DARK)
    style_header_subtitle = ParagraphStyle('HeaderSub', parent=styles['Normal'],
        fontName=REGULAR_FONT, fontSize=9.5, alignment=TA_RIGHT,
        leading=13, textColor=TEXT_COLOR)

    style_section_label = ParagraphStyle('SectionLabel', parent=styles['Normal'],
        fontName=BOLD_FONT, fontSize=8, leading=10,
        textColor=THEME_WHITE,
        backColor=THEME_DARK,
        borderPadding=(3, 6, 3, 6))

    style_table_header = ParagraphStyle('TableHeader', parent=styles['Normal'],
        fontName=BOLD_FONT, fontSize=9, alignment=TA_CENTER,
        textColor=colors.white, leading=11)

    style_cell_left   = ParagraphStyle('CellLeft',   parent=style_normal, alignment=TA_LEFT)
    style_cell_center = ParagraphStyle('CellCenter', parent=style_normal, alignment=TA_CENTER)
    style_cell_right  = ParagraphStyle('CellRight',  parent=style_normal, alignment=TA_RIGHT)

    elements = []

    # ==========================================
    # 1. VIRSRAKSTS / HEADER
    # ==========================================
    current_dir = os.path.dirname(os.path.abspath(__file__))
    logo_path = os.path.join(current_dir, "LatSEO logo black.png")

    if os.path.exists(logo_path):
        logo = RLImage(logo_path, width=40*mm, height=28*mm, kind='proportional')
    else:
        logo = Paragraph("<b>LatSEO</b>", style_bold)

    doc_type    = data.get('doc_type', 'Pavadzīme')
    doc_id      = data.get('doc_id', 'LSEO 0001')
    date_str    = data.get('date', '')
    due_str     = data.get('due_date', '')

    is_e_invoice = "e-rēķins" in doc_type.lower() or "e-invoice" in doc_type.lower()
    display_type = "Invoice" if is_e_invoice and lang == "en" else ("Rēķins" if is_e_invoice else doc_type)

    header_right = [
        Paragraph(f"{display_type} Nr. {doc_id}", style_header_title),
        Spacer(1, 3*mm),
        Paragraph(f"<b>{tr['date']}:</b> {date_str}", style_header_subtitle),
        Paragraph(f"<b>{tr['due_date']}:</b> {due_str}", style_header_subtitle),
    ]

    header_table = Table([[logo, header_right]], colWidths=[85*mm, 85*mm])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (0,0), 'LEFT'),
        ('ALIGN', (1,0), (1,0), 'RIGHT'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 4*mm))
    elements.append(ColorBar())
    elements.append(Spacer(1, 1*mm))
    elements.append(HorizontalLine(color=THEME_LIGHT, thickness=0.5))
    elements.append(Spacer(1, 5*mm))

    # ==========================================
    # 2. KLIENTS / E-RĒĶINA INFO (divkolonnu)
    # ==========================================
    if is_e_invoice:
        rec_name = data.get('receiver_name', '')
        rec_reg  = data.get('receiver_reg_no', '')
        rec_addr = data.get('receiver_address', '')
        cus_name = data.get('customer_name', '')
        cus_reg  = data.get('customer_reg_no', '')
        cus_addr = data.get('customer_address', '')

        rec_items = [
            Paragraph(f"<b>{tr['receiver']}</b>", style_bold),
            Spacer(1, 2*mm),
            Paragraph(f"<b>{rec_name}</b>", style_bold),
            Paragraph(f"<i>{tr['reg_no']}: {rec_reg}</i>", style_italic),
            Paragraph(f"<i>{tr['address']}: {rec_addr}</i>", style_italic),
        ]
        cus_items = [
            Paragraph(f"<b>{tr['customer']}</b>", style_bold),
            Spacer(1, 2*mm),
            Paragraph(f"<b>{cus_name}</b>", style_bold),
            Paragraph(f"<i>{tr['reg_no']}: {cus_reg}</i>", style_italic),
            Paragraph(f"<i>{tr['address']}: {cus_addr}</i>", style_italic),
        ]

        e_table = Table([[rec_items, cus_items]], colWidths=[83*mm, 83*mm])
        e_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('BACKGROUND', (0,0), (0,0), colors.HexColor("#EAF4FB")),
            ('BACKGROUND', (1,0), (1,0), colors.HexColor("#D0EAF8")),
            ('ROUNDEDCORNERS', [4, 4, 4, 4]),
            ('BOX', (0,0), (0,0), 0.5, THEME_MID),
            ('BOX', (1,0), (1,0), 0.5, THEME_MID),
            ('LEFTPADDING', (0,0), (-1,-1), 8),
            ('RIGHTPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ]))
        elements.append(e_table)
        elements.append(Spacer(1, 5*mm))
        elements.append(HorizontalLine(color=THEME_LIGHT, thickness=0.5))
        elements.append(Spacer(1, 5*mm))
        elements.append(Paragraph(f"<b>{tr['sender_label']}</b>", style_bold))
        elements.append(Spacer(1, 3*mm))
    else:
        # Klients un piegādātājs blakus
        client_items = [
            Paragraph(f"<b>{tr['client']}</b>", style_bold),
            Spacer(1, 2*mm),
            Paragraph(f"<b>{data.get('client_name', '')}</b>", style_bold),
            Paragraph(f"<i>{tr['address']}: {data.get('client_address', '')}</i>", style_italic),
            Paragraph(f"<i>{tr['reg_no']}: {data.get('client_reg_no', '')}</i>", style_italic),
            Paragraph(f"<i>{tr['vat_no']}: {data.get('client_vat_no', '')}</i>", style_italic),
        ]

        sender_items = [
            Paragraph(f"<b>{tr['sender']}</b>", style_bold),
            Spacer(1, 2*mm),
            Paragraph("<b>SIA Baltic SEO</b>", style_bold),
            Paragraph(f"<i>{tr['address']}: Ķekavas nov., Ķekavas pag., Odukalns, Kārklu iela 4, LV-2123</i>", style_italic),
            Paragraph(f"<i>{tr['reg_no']}: 40203749304</i>", style_italic),
            Paragraph(f"<i>{tr['phone']}: +371 24424434</i>", style_italic),
        ]

        two_col = Table([[client_items, sender_items]], colWidths=[83*mm, 83*mm])
        two_col.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('BACKGROUND', (0,0), (0,0), colors.HexColor("#EAF4FB")),
            ('BACKGROUND', (1,0), (1,0), colors.HexColor("#D0EAF8")),
            ('BOX', (0,0), (0,0), 0.5, THEME_MID),
            ('BOX', (1,0), (1,0), 0.5, THEME_MID),
            ('LEFTPADDING', (0,0), (-1,-1), 8),
            ('RIGHTPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ]))
        elements.append(two_col)
        elements.append(Spacer(1, 4*mm))

    # ==========================================
    # 3. BANKAS INFO (pilna platuma josla)
    # ==========================================
    bank_label_style = ParagraphStyle('BankLabel', parent=style_small,
        fontName=BOLD_FONT, fontSize=8, textColor=THEME_DARK)
    bank_val_style = ParagraphStyle('BankVal', parent=style_small,
        fontName=REGULAR_FONT, fontSize=8.5, textColor=TEXT_COLOR)

    bank_data_table = Table([
        [
            Paragraph("AS Swedbank", bank_label_style),
            Paragraph(f"{tr['swift']}: HABALV22", bank_val_style),
            Paragraph(f"{tr['account']}:", bank_val_style),
            Paragraph("<b>LV05HABA0551065262400</b>", bank_label_style),
        ]
    ], colWidths=[30*mm, 35*mm, 40*mm, 65*mm])
    bank_data_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#EAF4FB")),
        ('BOX', (0,0), (-1,-1), 0.5, THEME_MID),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    elements.append(bank_data_table)
    elements.append(Spacer(1, 7*mm))

    # ==========================================
    # 4. PREČU TABULA
    # ==========================================
    col_headers = [
        Paragraph(tr['col_name'],  style_table_header),
        Paragraph(tr['col_unit'],  style_table_header),
        Paragraph(tr['col_qty'],   style_table_header),
        Paragraph(tr['col_price'], style_table_header),
        Paragraph(tr['col_total'], style_table_header),
    ]

    table_data = [col_headers]
    items = data.get('items', [])

    if not items:
        for _ in range(3):
            items.append({'name': '', 'unit': '', 'qty': '', 'price': '', 'total': ''})

    for i, item in enumerate(items):
        seq_num  = item.get('seq', '')
        name_str = f"{seq_num}. {item['name']}" if seq_num else item['name']
        row_bg   = colors.HexColor("#EAF4FB") if i % 2 == 0 else colors.white
        table_data.append([
            Paragraph(name_str, style_cell_left),
            Paragraph(item['unit'], style_cell_center),
            Paragraph(str(item['qty']), style_cell_center),
            Paragraph(item['price'], style_cell_right),
            Paragraph(item['total'], style_cell_right),
        ])

    col_widths = [70*mm, 22*mm, 22*mm, 24*mm, 32*mm]
    t_table = Table(table_data, colWidths=col_widths)

    # Alternating row colours
    table_style_cmds = [
        ('BACKGROUND', (0,0), (-1,0), THEME_DARK),
        ('VALIGN', (0,0), (-1,0), 'MIDDLE'),
        ('ROWHEIGHT', (0,0), (-1,0), 18),
        ('LINEBEFORE', (1,0), (-1,0), 0.5, colors.white),
        ('VALIGN', (0,1), (-1,-1), 'TOP'),
        ('GRID', (0,0), (-1,-1), 0.4, THEME_MID),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
    ]
    for i in range(1, len(table_data)):
        bg = colors.HexColor("#EAF4FB") if i % 2 == 1 else colors.white
        table_style_cmds.append(('BACKGROUND', (0,i), (-1,i), bg))

    t_table.setStyle(TableStyle(table_style_cmds))
    elements.append(t_table)
    elements.append(Spacer(1, 3*mm))

    # ==========================================
    # 5. KOPSUMMAS (labā puse)
    # ==========================================
    subtotal_val     = data.get('subtotal', '0,00')
    vat_val          = data.get('vat', '0,00')
    total_val        = data.get('total', '0,00')
    raw_discount_eur = data.get('raw_discount_eur', 0.0)
    apply_vat        = data.get('apply_vat', True)

    style_sum_label = ParagraphStyle('SumLabel', parent=style_normal,
        fontName=REGULAR_FONT, alignment=TA_RIGHT, fontSize=9)
    style_sum_label_bold = ParagraphStyle('SumLabelBold', parent=style_normal,
        fontName=BOLD_FONT, alignment=TA_RIGHT, fontSize=9.5)
    style_sum_val = ParagraphStyle('SumVal', parent=style_normal,
        fontName=REGULAR_FONT, alignment=TA_RIGHT, fontSize=9)
    style_sum_val_bold = ParagraphStyle('SumValBold', parent=style_normal,
        fontName=BOLD_FONT, alignment=TA_RIGHT, fontSize=9.5, textColor=THEME_DARK)

    is_advance = "avansa" in doc_type.lower() or "advance" in doc_type.lower()

    totals_rows = []
    if raw_discount_eur > 0:
        totals_rows.append([Paragraph(tr['total_no_vat'], style_sum_label),
                            Paragraph(f"{subtotal_val} €", style_sum_val)])
        disc_str = f"{tr['discount']} ({data.get('discount_percent', 0):g}%)"
        totals_rows.append([Paragraph(disc_str, style_sum_label),
                            Paragraph(f"-{data.get('discount_eur', '0,00')} €", style_sum_val)])
        totals_rows.append([Paragraph(tr['total_discount'], style_sum_label),
                            Paragraph(f"{data.get('subtotal_after_discount', '0,00')} €", style_sum_val)])
        if apply_vat:
            totals_rows.append([Paragraph(tr['vat'], style_sum_label),
                                Paragraph(f"{vat_val} €", style_sum_val)])
        totals_rows.append([Paragraph(tr['grand_total'], style_sum_label_bold),
                            Paragraph(f"{total_val} €", style_sum_val_bold)])
    else:
        totals_rows.append([Paragraph(tr['total_label'], style_sum_label),
                            Paragraph(f"{subtotal_val} €", style_sum_val)])
        if apply_vat:
            totals_rows.append([Paragraph(tr['vat'], style_sum_label),
                                Paragraph(f"{vat_val} €", style_sum_val)])
        totals_rows.append([Paragraph(tr['grand_total'], style_sum_label_bold),
                            Paragraph(f"{total_val} €", style_sum_val_bold)])

    last_idx = len(totals_rows) - 1

    totals_table = Table(totals_rows, colWidths=[100*mm, 70*mm])
    totals_style = [
        ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LINEABOVE', (0, last_idx), (-1, last_idx), 1, THEME_MID),
        ('BACKGROUND', (0, last_idx), (-1, last_idx), colors.HexColor("#D0EAF8")),
    ]
    totals_table.setStyle(TableStyle(totals_style))
    elements.append(totals_table)

    # ==========================================
    # 6. AVANSS UN SUMMA VĀRDIEM
    # ==========================================
    elements.append(Spacer(1, 4*mm))

    if is_advance:
        raw_advance  = data.get('raw_advance', 0.0)
        fmt_advance  = fmt_curr(raw_advance)
        pct_val      = int(round(data.get('advance_percent', 0)))
        adv_text     = f'<font name="{BOLD_FONT}" color="#1A6FA8">{tr["advance_payable"]} ({pct_val}%): {fmt_advance} €</font>'
        elements.append(Paragraph(adv_text,
            ParagraphStyle('AdvRight', parent=style_bold, alignment=TA_RIGHT, fontSize=10)))
        elements.append(Spacer(1, 2*mm))

    amount_words = data.get('amount_words', '')
    elements.append(Paragraph(f"<i>{tr['words_prefix']}{amount_words}</i>",
        ParagraphStyle('WordsRight', parent=style_italic, alignment=TA_RIGHT, fontSize=9)))

    # ==========================================
    # 7. KOMENTĀRI
    # ==========================================
    comments = data.get('comments', '').strip()
    if comments:
        elements.append(Spacer(1, 7*mm))
        elements.append(HorizontalLine(color=THEME_LIGHT, thickness=0.5))
        elements.append(Spacer(1, 3*mm))
        elements.append(Paragraph(f"<b>{tr['extra_info']}</b>", style_bold))
        elements.append(Spacer(1, 2*mm))
        elements.append(Paragraph(comments.replace('\n', '<br/>'), style_normal))

    # ==========================================
    # 8. PARAKSTI
    # ==========================================
    elements.append(Spacer(1, 8*mm))
    elements.append(ColorBar(color=THEME_MID))
    elements.append(Spacer(1, 1*mm))
    elements.append(HorizontalLine(color=THEME_LIGHT, thickness=0.5))
    elements.append(Spacer(1, 4*mm))

    signatory = data.get('signatory', 'SIA Baltic SEO valdes loceklis Adrians Stankevičs')

    doc_lower = doc_type.lower()
    if "pavadzīme" in doc_lower or "invoice" in doc_lower:
        prep_text = tr['prep_invoice']
        recv_text = tr['recv_invoice']
    elif "avansa" in doc_lower or "advance" in doc_lower:
        prep_text = tr['prep_advance']
        recv_text = tr['recv_advance']
    else:
        prep_text = tr['prep_receipt']
        recv_text = tr['recv_receipt']

    sig_data = [
        [Paragraph(f"{prep_text} <i>{signatory}</i>", style_normal),
         Paragraph("__________________________", style_normal)],
        [Paragraph(recv_text, style_normal),
         Paragraph("__________________________", style_normal)],
    ]
    sig_table = Table(sig_data, colWidths=[105*mm, 65*mm])
    sig_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'BOTTOM'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ('ALIGN', (1,0), (1,1), 'RIGHT'),
    ]))
    elements.append(sig_table)

    doc.build(elements)
    buffer.seek(0)
    return buffer
