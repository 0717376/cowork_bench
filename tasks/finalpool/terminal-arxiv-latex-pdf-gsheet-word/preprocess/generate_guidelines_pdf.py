"""
Reproducibly regenerate review_guidelines.pdf with a Russian scoring rubric.

The five criterion names are kept bilingual (Russian + English field name) because
the English names map directly to Google Sheet columns the agent must produce.
The Accept/Revise/Reject recommendation literals are also kept in English.

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


# (russian_text, style) tuples. style in {title, subtitle, heading, body}
CONTENT = [
    ("Руководство по рецензированию статей конференции", "title"),
    ("Критерии оценки и балльная шкала", "subtitle"),
    ("Каждая статья оценивается по пяти измерениям, по шкале от 1 (минимум) до 5 (максимум):", "body"),

    ("1. Новизна (Novelty)", "heading"),
    ("Вводит ли статья действительно новые архитектуры, методы или идеи? Балл 5 за "
     "прорывные новые подходы, 4 за существенные расширения, 3 за инкрементальные "
     "улучшения, 2 за незначительные вариации, 1 за отсутствие новизны.", "body"),

    ("2. Строгость методологии (Methodology_Rigor)", "heading"),
    ("Хорошо ли определены методы с чёткими математическими формулировками? Балл 5 за "
     "строгое формальное изложение с доказательствами или выводами, 4 за солидный "
     "математический аппарат, 3 за адекватную формализацию, 2 за неформальные описания, "
     "1 за отсутствие деталей методологии.", "body"),

    ("3. Полнота экспериментов (Experimental_Completeness)", "heading"),
    ("Являются ли эксперименты тщательными и хорошо спланированными? Статьи с большим "
     "числом разделов и подробными экспериментальными постановками получают более "
     "высокий балл. Балл 5 за всесторонние эксперименты с абляциями, 4 за солидные "
     "эксперименты, 3 за адекватную оценку, 2 за ограниченные эксперименты, 1 за "
     "отсутствие экспериментов.", "body"),

    ("4. Ясность (Clarity)", "heading"),
    ("Хорошо ли написана статья и чёткая ли у неё структура? Статьи с хорошо "
     "организованными разделами и читабельными аннотациями получают более высокий балл. "
     "Балл 5 за образцовое изложение, 4 за чёткую подачу, 3 за адекватную ясность, "
     "2 за местами неясный текст, 1 за плохо написанный текст.", "body"),

    ("5. Значимость (Significance)", "heading"),
    ("Решает ли статья важную проблему с сильными результатами? Балл 5 за очень "
     "влиятельную работу с результатами уровня state-of-the-art, 4 за значимый вклад, "
     "3 за умеренное влияние, 2 за ограниченное влияние, 1 за пренебрежимо малую "
     "значимость.", "body"),

    ("Пороги рекомендаций (Recommendation)", "subtitle"),
    ("Итоговый балл (Total_Score) — сумма всех пяти критериев:", "body"),
    ("- 20-25: Accept (принять)", "body"),
    ("- 15-19: Revise (требуются крупные или мелкие правки)", "body"),
    ("- Менее 15: Reject (отклонить)", "body"),
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
        "subtitle": ParagraphStyle("s", parent=base["Heading2"], fontName=font_name, fontSize=13),
        "heading": ParagraphStyle("h", parent=base["Heading3"], fontName=font_name, fontSize=11),
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
    out = os.path.join(here, "..", "initial_workspace", "review_guidelines.pdf")
    generate(os.path.abspath(out))
    print(f"Generated {os.path.abspath(out)}")
