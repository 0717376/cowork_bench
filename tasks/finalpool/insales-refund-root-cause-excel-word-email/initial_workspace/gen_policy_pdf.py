"""
Reproducibly regenerate Quality_Policy.pdf with a Russian methodology.

Numeric thresholds ($50/$20, 3.0, 2 refunds) are FROZEN — they flow into the
eval and deliverables. English label literals are kept verbatim because they are
written into the Excel/Word deliverables and greped by the eval:
  - severity: Critical / Major / Minor
  - issue type: Quality Issue / Service Issue / No Reviews
  - root cause: Manufacturing Defect / Shipping Damage / Wrong Item /
    Customer Expectations / Other

Run standalone or import generate(out_path).
"""
import os


def _find_cyrillic_font():
    candidates = [
        "/Library/Fonts/Arial Unicode.ttf",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/Library/Fonts/Arial.ttf",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


CONTENT = [
    ("Политика контроля качества и расследования возвратов", "title"),

    ("1. Классификация серьёзности возвратов", "heading"),
    ("Все возвраты должны быть классифицированы по серьёзности на основе суммы возврата:", "body"),
    ("- Critical: сумма возврата больше $50 — требуется немедленная проверка.", "body"),
    ("- Major: сумма возврата от $20 до $50 (включая $20) — проверка в течение 48 часов.", "body"),
    ("- Minor: сумма возврата меньше $20 — стандартный цикл проверки.", "body"),

    ("2. Анализ корреляции с отзывами", "heading"),
    ("Для каждого товара, по которому был возврат, сопоставьте данные с отзывами на товар, "
     "чтобы определить тип проблемы:", "body"),
    ("- Если по товару есть возвраты И средний рейтинг отзывов ниже 3.0 — классифицируйте как Quality Issue.", "body"),
    ("- Если по товару есть возвраты, но средний рейтинг 3.0 или выше — классифицируйте как Service Issue.", "body"),
    ("- Если по товару есть возвраты, но отзывов нет — классифицируйте как No Reviews.", "body"),

    ("3. Категории первопричин", "heading"),
    ("При анализе причин возвратов и текстов отзывов отнесите первопричину к одной из "
     "следующих категорий:", "body"),
    ("- Manufacturing Defect", "body"),
    ("- Shipping Damage", "body"),
    ("- Wrong Item", "body"),
    ("- Customer Expectations", "body"),
    ("- Other", "body"),

    ("4. Пороги действий", "heading"),
    ("Товары с 2 и более возвратами требуют официальной проверки поставщика. Отдел контроля "
     "качества обязан инициировать процесс проверки поставщика и задокументировать находки.", "body"),
]


def generate(out_path):
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

    font_path = _find_cyrillic_font()
    font_name = "Helvetica"
    if font_path:
        pdfmetrics.registerFont(TTFont("CyrFont", font_path))
        font_name = "CyrFont"

    base = getSampleStyleSheet()
    styles = {
        "title": ParagraphStyle("t", parent=base["Title"], fontName=font_name, fontSize=16),
        "heading": ParagraphStyle("h", parent=base["Heading2"], fontName=font_name, fontSize=12),
        "body": ParagraphStyle("b", parent=base["Normal"], fontName=font_name, fontSize=10, leading=14),
    }

    doc = SimpleDocTemplate(out_path, pagesize=letter,
                            topMargin=0.7 * inch, bottomMargin=0.7 * inch)
    flow = []
    for text, style in CONTENT:
        flow.append(Paragraph(text, styles[style]))
        flow.append(Spacer(1, 4))
    doc.build(flow)


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    out = os.path.join(here, "Quality_Policy.pdf")
    generate(os.path.abspath(out))
    print(f"Generated {os.path.abspath(out)}")
