"""Preprocess: regenerate the Russian quiz-guidelines PDF into the agent workspace.

Canvas data is read live (read-only); no database injection is needed. The only
local artifact the agent consumes is Quiz_Guidelines.pdf, which we rebuild in
Russian (keeping the numeric targets) so it is not a frozen English file. A
Cyrillic-capable TTF (DejaVuSans) is bundled under preprocess/assets so PDF
generation works regardless of the host fonts.
"""
import argparse
import os
import shutil


PREPROCESS_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(PREPROCESS_DIR, "assets")
FONT_PATH = os.path.join(ASSETS_DIR, "DejaVuSans.ttf")
PREBUILT_PDF = os.path.join(
    PREPROCESS_DIR, os.pardir, "initial_workspace", "Quiz_Guidelines.pdf"
)

# Russian headings/labels; numeric targets preserved from the original.
TITLE = "Методические указания по оценке успеваемости в тестах"
HEADER = ("Показатель", "Целевое значение")
ROWS = [
    ("Минимальный процент прохождения (Pass Rate)", "60%"),
    ("Целевой средний балл (Avg Score)", "70%"),
    ("Максимальное ограничение времени (Time Limit)", "60 минут"),
]
PURPOSE = (
    "Цель: оценить тесты курса по приведённым ниже целевым показателям. "
    "Pass_Rate_Pct = средний балл / максимальный балл * 100, округление до 1 знака."
)


def write_pdf(path):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.pdfgen import canvas
    except ImportError:
        # reportlab unavailable in this runtime: ship the pre-built Russian PDF
        # (identical content) so a missing optional dep never aborts the task.
        shutil.copyfile(PREBUILT_PDF, path)
        return

    font_name = "DejaVuSans"
    pdfmetrics.registerFont(TTFont(font_name, FONT_PATH))

    c = canvas.Canvas(path, pagesize=A4)
    width, height = A4
    y = height - 72

    c.setFont(font_name, 16)
    c.drawString(72, y, TITLE)
    y -= 36

    val_x = 430
    c.setFont(font_name, 12)
    c.drawString(72, y, HEADER[0])
    c.drawString(val_x, y, HEADER[1])
    y -= 8
    c.line(72, y, width - 72, y)
    y -= 24

    c.setFont(font_name, 10)
    for label, target in ROWS:
        c.drawString(72, y, label)
        c.drawString(val_x, y, target)
        y -= 22

    y -= 24
    c.setFont(font_name, 10)
    # Wrap the purpose line manually to fit the page width.
    words = PURPOSE.split()
    line = ""
    for w in words:
        trial = (line + " " + w).strip()
        if pdfmetrics.stringWidth(trial, font_name, 10) > (width - 144):
            c.drawString(72, y, line)
            y -= 16
            line = w
        else:
            line = trial
    if line:
        c.drawString(72, y, line)

    c.showPage()
    c.save()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    if not args.agent_workspace:
        print("No agent workspace provided; skipping PDF regeneration.")
        return

    os.makedirs(args.agent_workspace, exist_ok=True)
    out_pdf = os.path.join(args.agent_workspace, "Quiz_Guidelines.pdf")
    write_pdf(out_pdf)
    print(f"Wrote Russian quiz guidelines PDF: {out_pdf}")


if __name__ == "__main__":
    main()
