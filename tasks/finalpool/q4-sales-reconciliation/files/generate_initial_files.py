"""
Generate the initial workspace PDF file: Q4_2025_Regional_Targets.pdf

NOTE: After the snowflake->clickhouse fork the warehouse REGION values are
Russian. The PDF region names must therefore be russified through the SAME
central map (scripts/clickhouse_relabel_map.py REGIONS) so the agent can join
PDF targets to warehouse actuals by region. USD amounts are FROZEN.
"""
import os
import sys
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

# Import the central English->Russian relabel map (single source of truth).
_SCRIPTS = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "..", "..", "..", "..", "scripts"))
if _SCRIPTS not in sys.path:
    sys.path.insert(0, _SCRIPTS)
from clickhouse_relabel_map import REGIONS  # noqa: E402

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))


def _register_cyrillic_font():
    """Register a Cyrillic-capable TTF. Returns (regular, bold) font names.

    Falls back to Helvetica only if no Unicode font is found (should not happen
    in this environment; DejaVuSans ships with matplotlib).
    """
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    # Resolve DejaVuSans from matplotlib's bundled fonts dynamically (no hardcoded paths).
    try:
        import matplotlib
        _ttf = os.path.join(matplotlib.get_data_path(), "fonts", "ttf")
        dejavu = os.path.join(_ttf, "DejaVuSans.ttf")
        dejavu_bold = os.path.join(_ttf, "DejaVuSans-Bold.ttf")
    except Exception:
        dejavu = dejavu_bold = ""
    candidates = [
        ("DejaVuSans", "DejaVuSans-Bold", dejavu, dejavu_bold),
        ("ArialUni", "ArialUni",
         "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
         "/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),
    ]
    for reg_name, bold_name, reg_path, bold_path in candidates:
        if os.path.exists(reg_path) and os.path.exists(bold_path):
            try:
                pdfmetrics.registerFont(TTFont(reg_name, reg_path))
                pdfmetrics.registerFont(TTFont(bold_name, bold_path))
                return reg_name, bold_name
            except Exception:
                continue
    return "Helvetica", "Helvetica-Bold"


def create_targets_pdf():
    pdf_path = os.path.join(OUTPUT_DIR, "Q4_2025_Regional_Targets.pdf")
    FONT, FONT_BOLD = _register_cyrillic_font()
    doc = SimpleDocTemplate(pdf_path, pagesize=letter,
                            topMargin=1*inch, bottomMargin=1*inch)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'CustomTitle', parent=styles['Title'],
        fontName=FONT_BOLD, fontSize=20, alignment=TA_CENTER, spaceAfter=12
    )
    subtitle_style = ParagraphStyle(
        'CustomSubtitle', parent=styles['Normal'],
        fontName=FONT, fontSize=12, alignment=TA_CENTER, spaceAfter=30, textColor=colors.grey
    )
    footer_style = ParagraphStyle(
        'Footer', parent=styles['Normal'],
        fontName=FONT, fontSize=9, alignment=TA_CENTER, textColor=colors.grey, spaceBefore=40
    )

    elements = []

    # Title
    elements.append(Paragraph("Целевые показатели выручки по регионам, 4 квартал 2025", title_style))
    elements.append(Paragraph("Утверждено руководством — сентябрь 2025", subtitle_style))
    elements.append(Spacer(1, 20))

    # Table data — region names russified via the central map so the agent can
    # join these targets to the (Russian) warehouse regions. USD amounts FROZEN.
    data = [
        ["Регион", "Плановая выручка (USD)"],
        [REGIONS["Asia Pacific"], "$65,000.00"],
        [REGIONS["Europe"], "$60,000.00"],
        [REGIONS["Latin America"], "$55,000.00"],
        [REGIONS["Middle East"], "$50,000.00"],
        [REGIONS["North America"], "$55,000.00"],
        ["", ""],
        ["Итого", "$285,000.00"],
    ]

    table = Table(data, colWidths=[3*inch, 2.5*inch])
    table.setStyle(TableStyle([
        # Header
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2C3E50')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), FONT_BOLD),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        # Data rows
        ('FONTNAME', (0, 1), (-1, -1), FONT),
        ('FONTSIZE', (0, 1), (-1, -1), 11),
        ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        # Total row
        ('FONTNAME', (0, -1), (-1, -1), FONT_BOLD),
        ('LINEABOVE', (0, -1), (-1, -1), 1, colors.black),
        # Grid
        ('GRID', (0, 0), (-1, -3), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -3), [colors.white, colors.HexColor('#ECF0F1')]),
        # Padding
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LEFTPADDING', (0, 0), (-1, -1), 12),
        ('RIGHTPADDING', (0, 0), (-1, -1), 12),
    ]))

    elements.append(table)

    # Footer note
    elements.append(Paragraph(
        "Показатели приведены в USD. Результаты считаются только по доставленным заказам.",
        footer_style
    ))

    doc.build(elements)
    print(f"Created: {pdf_path}")
    return pdf_path


if __name__ == "__main__":
    create_targets_pdf()
