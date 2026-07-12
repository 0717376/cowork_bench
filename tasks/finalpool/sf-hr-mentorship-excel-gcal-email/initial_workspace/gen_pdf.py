"""Generate Mentorship_Guidelines.pdf (русское руководство).

Числовые критерии (стаж/оценки) идентичны оригиналу. Email-адреса оставлены на
английском, так как eval/задача их сверяет дословно. Кириллица рендерится через
TTF-шрифт, зарегистрированный в reportlab.
"""
import os
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


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


FONT = "Body"
FONT_PATH = _find_cyrillic_font()
if FONT_PATH:
    pdfmetrics.registerFont(TTFont(FONT, FONT_PATH))
else:
    FONT = "Helvetica"


def make_pdf(path, title, lines):
    c = rl_canvas.Canvas(path, pagesize=A4)
    w, h = A4
    c.setFont(FONT, 14)
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
    "РУКОВОДСТВО ПО ПРОГРАММЕ НАСТАВНИЧЕСТВА",
    "",
    "Назначение",
    "Программа наставничества объединяет опытных сотрудников с высокими показателями",
    "и младших сотрудников для ускорения профессионального развития и передачи знаний.",
    "",
    "Критерии наставника",
    "- Не менее 10 лет опыта работы",
    "- Оценка эффективности 4 или выше (по 5-балльной шкале)",
    "- Готовность уделять не менее 2 часов в месяц наставнической деятельности",
    "",
    "Критерии подопечного",
    "- Не более 2 лет опыта работы",
    "- Оценка эффективности 3 или выше",
    "- Подтверждённая мотивация и стремление к обучению",
    "",
    "Процесс формирования пар",
    "Пары формируются последовательно по рангу эффективности. По возможности",
    "поощряются межведомственные пары для расширения кругозора.",
    "",
    "Продолжительность программы",
    "Каждый цикл наставничества длится 6 месяцев. Обе стороны обязуются проводить",
    "ежемесячные встречи и ежеквартальные обзоры прогресса.",
    "",
    "Установочная встреча",
    "Все пары посещают установочную встречу через 7 дней после запуска программы,",
    "чтобы согласовать цели, обсудить ожидания и составить начальный график встреч.",
    "",
    "Коммуникация",
    "Все обновления программы рассылаются по электронной почте на program@hr.example.com.",
    "Вопросы можно направлять на hr@company.example.com.",
]

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Mentorship_Guidelines.pdf")
make_pdf(out, "Руководство по программе наставничества", lines)
print(f"Created: {out}")
