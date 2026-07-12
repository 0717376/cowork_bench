"""Generate Performance_Review_Criteria.pdf."""
import os
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Register a Unicode font so Cyrillic renders correctly. Fall back to Helvetica.
FONT, FONT_BOLD = "Helvetica", "Helvetica-Bold"
for cand, bold in [
    ("/System/Library/Fonts/Supplemental/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf"),
]:
    if os.path.exists(cand) and os.path.exists(bold):
        try:
            pdfmetrics.registerFont(TTFont("Body", cand))
            pdfmetrics.registerFont(TTFont("BodyBold", bold))
            FONT, FONT_BOLD = "Body", "BodyBold"
            break
        except Exception:
            pass


def make_pdf(path, title, lines):
    c = rl_canvas.Canvas(path, pagesize=A4)
    w, h = A4
    c.setFont(FONT_BOLD, 14)
    c.drawString(50, h - 50, title)
    c.setFont(FONT, 11)
    y = h - 80
    for line in lines:
        c.drawString(50, y, str(line))
        y -= 18
        if y < 80:
            c.showPage()
            c.setFont(FONT, 11)
            y = h - 50
    c.save()

lines = [
    "КРИТЕРИИ ОЦЕНКИ РЕЗУЛЬТАТИВНОСТИ",
    "",
    "1. Шкала оценок",
    "Сотрудники оцениваются по 5-балльной шкале результативности:",
    "  5 - Превосходно: стабильно превосходит ожидания по всем задачам",
    "  4 - Выше среднего: часто превосходит ожидания",
    "  3 - Соответствует ожиданиям: стабильно выполняет задачи",
    "  2 - Ниже ожиданий: частично выполняет задачи; требуется улучшение",
    "  1 - Неудовлетворительно: не выполняет основные задачи",
    "",
    "2. Классификация для совета по оценке",
    "Лучшие сотрудники (Top Performers): оценка 5 (Превосходно)",
    "  - Имеют право на повышение оклада и рассмотрение на повышение в должности",
    "  - Отмечаются в отчётах о результативности отделов",
    "",
    "Отстающие сотрудники (Underperformers): оценка 1 или 2",
    "  - Подлежат включению в план повышения результативности (PIP)",
    "  - Рассматриваются на Annual Performance Review Board Meeting",
    "",
    "3. График заседания совета",
    "Annual Performance Review Board Meeting проводится через 21 день после",
    "даты запуска цикла оценки. Присутствуют все руководители отделов и HR-руководство.",
    "",
    "4. Требования к данным",
    "Отчёты должны содержать разбивку по отделам:",
    "  - Количество лучших сотрудников и средняя зарплата",
    "  - Количество отстающих сотрудников и средняя зарплата",
    "  - Общая средняя оценка результативности по отделу",
    "",
    "5. Рассылка",
    "Отчёты направляются на executives@company.example.com",
    "с адреса hr@company.example.com",
]

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Performance_Review_Criteria.pdf")
make_pdf(out, "Performance Review Criteria", lines)
print(f"Created: {out}")
