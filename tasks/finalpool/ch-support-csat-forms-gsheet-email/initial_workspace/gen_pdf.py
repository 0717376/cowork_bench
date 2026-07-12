"""Generate SLA_Policy_Reference.pdf."""
import os
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.pagesizes import A4

def make_pdf(path, title, lines):
    c = rl_canvas.Canvas(path, pagesize=A4)
    w, h = A4
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, h - 50, title)
    c.setFont("Helvetica", 11)
    y = h - 80
    for line in lines:
        c.drawString(50, y, str(line))
        y -= 18
        if y < 80:
            c.showPage()
            y = h - 50
    c.save()

lines = [
    "SPRAVOCHNIK PO POLITIKE SLA",
    "",
    "1. Obzor",
    "Soglasheniya ob urovne servisa (SLA) opredelyayut ozhidaemoe vremya",
    "otveta i resheniya po tiketam podderzhki v zavisimosti ot prioriteta.",
    "",
    "2. Urovni prioriteta i porogi SLA",
    "",
    "HIGH (vysokiy prioritet)",
    "  - SLA po vremeni otveta: 4 chasa",
    "  - SLA po vremeni resheniya: 24 chasa",
    "  - Opredelenie: kriticheskoe vliyanie na biznes, sboy sistemy ili risk poteri dannyh",
    "",
    "MEDIUM (sredniy prioritet)",
    "  - SLA po vremeni otveta: 8 chasov",
    "  - SLA po vremeni resheniya: 48 chasov",
    "  - Opredelenie: sushchestvennoe vliyanie na operacii, no est obhodnye resheniya",
    "",
    "LOW (nizkiy prioritet)",
    "  - SLA po vremeni otveta: 24 chasa",
    "  - SLA po vremeni resheniya: 72 chasa",
    "  - Opredelenie: melkie problemy ili pozhelaniya s minimalnym vliyaniem na biznes",
    "",
    "3. Izmerenie soblyudeniya SLA",
    "Soblyudenie SLA izmeryaetsya kak procent tiketov, po kotorym pervichnyy",
    "otvet byl dan v predelah ukazannogo okna vremeni SLA.",
    "",
    "4. Udovletvorennost klientov",
    "Posle resheniya tiketa klienty ocenivayut svoy opyt po shkale ot 1 do 5.",
    "Celevoy sredniy ball CSAT: 4.0 ili vyshe.",
    "",
    "5. Otchetnost",
    "Otchety po soblyudeniyu SLA otpravlyayutsya na support-management@company.example.com.",
    "Kontakt komandy analitiki: analytics@company.example.com.",
]

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "SLA_Policy_Reference.pdf")
make_pdf(out, "SLA Policy Reference", lines)
print(f"Created: {out}")
