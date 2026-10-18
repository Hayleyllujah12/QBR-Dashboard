# Builds SYNTHETIC Certiport-style "EXAM SCORE REPORT" PDFs for the MOS tests.
# Every name, email, ID and reference number here is invented. Never put real
# score reports in the repo (they hold minors' names, emails and addresses).
#   python tests/make-mos-fixture.py      -> tests/mos-fixture/...
# Layout mirrors the real export (positions in PDF points, letter page), including
# the quirks the parser must survive: the dropped "ffi" ligature ("O" / "ce 2019)"),
# words split into several text items, word gaps drawn as spacing, and the
# "Required Score" label missing from the FINAL SCORE box.
import os
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "mos-fixture")

EXCEL = ["Manage Worksheets and Workbooks", "Manage Data Cells and Ranges", "Manage Tables and Table Data",
         "Perform Operations by using Formulas and Functions", "Manage Charts"]
PPT = ["Manage Presentations", "Manage Slides", "Insert and Format Text, Shapes, and Images",
       "Insert Tables, Charts, SmartArt, 3D Models, and Media", "Apply Transitions and Animations"]


def split_text(c, x, y, parts, font="Helvetica", size=11, gap=0.0):
    """Draw consecutive text items (simulates one word split across items)."""
    for p in parts:
        c.setFont(font, size)
        c.drawString(x, y, p)
        x += stringWidth(p, font, size) + gap


def page(c, p):
    c.setFont("Helvetica-Bold", 13)
    if p.get("split_title"):
        split_text(c, 429, 666.8, ["EXAM S", "CORE REPORT"], "Helvetica-Bold", 13)
    else:
        c.drawString(429, 666.8, "EXAM SCORE REPORT")
    c.setFont("Helvetica-Bold", 12)
    c.drawString(40, 612, "CANDIDATE")
    c.drawString(315, 612, "EXAM")
    c.drawString(40, 587.7, p["name"])
    c.setFont("Helvetica", 11)
    c.drawString(40, 573.6, "123 Sample Street")          # address lines: must be ignored
    c.drawString(40, 558.2, "Example City 0000")
    c.drawString(40, 509.8, p["email"])
    c.setFont("Helvetica-Bold", 12)
    t1, t2 = p["title"]
    c.drawString(315, 587.7, t1)
    if t2:
        c.drawString(315, 569.9, t2)
    c.setFont("Helvetica", 11)
    if p.get("split_ref"):
        split_text(c, 315, 538.9, ["Exam refe", "rence #: " + p["ref"]])
    else:
        split_text(c, 315, 538.9, ["Exam reference #", ": " + p["ref"]])
    c.drawString(315, 524.4, "Date: " + p["date"])
    c.drawString(315, 509.8, "ID: " + p["cid"])
    c.setFont("Helvetica-Bold", 12)
    c.drawString(40, 467.8, "RESULTS")
    c.setFont("Helvetica", 11)
    for i, v in enumerate(range(100, 1100, 100)):
        c.drawString(210 + 38 * i, 467.7, str(v))
    c.drawString(40, 447.7, "Required Score")
    c.drawString(40, 422.6, "Your Score")
    c.setFont("Helvetica-Bold", 12)
    if p.get("gap_words"):
        split_text(c, 41, 378.7, ["SECTION", "ANALYSIS"], "Helvetica-Bold", 12, gap=3.6)
    else:
        c.drawString(41, 378.7, "SECTION ANALYSIS")
    c.drawString(415, 378.7, "FINAL SCORE")
    c.setFont("Helvetica", 11)
    y = 351.8
    for area, pct in zip(p["areas"], p["pcts"]):
        c.drawString(41.5, y, area)
        c.drawString(363.7, y, "%d%%" % pct)
        y -= 25.8
    if not p.get("no_req_label"):
        c.drawString(415.7, 352, "Required Score")
    c.drawString(546.2, 352, "700")
    c.drawString(415.7, 326, "Your Score")
    c.drawString(546.2, 326, str(p["score"]))
    c.setFont("Helvetica-Bold", 12)
    c.drawString(415, 291.2, "OUTCOME")
    c.setFont("Helvetica", 11)
    c.drawString(415, 256.7, "Pass" if p["score"] >= 700 else "Fail")
    c.setFont("Helvetica", 4.6)
    c.drawString(120, 58.9, "(c) 2026 Certiport sample footer (synthetic test fixture) - Certiport and the Certiport logo are registered trademarks")
    c.showPage()


def pdf(rel, pages, extra_text_pages=0):
    path = os.path.join(OUT, rel)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    c = canvas.Canvas(path, pagesize=(612, 792))
    c.setTitle("Synthetic MOS fixture")
    c.setAuthor("QBR tests")
    for p in pages:
        page(c, p)
    for i in range(extra_text_pages):
        c.setFont("Helvetica-Bold", 20)
        c.drawString(100, 600, "CERTIFICATE OF ACHIEVEMENT")
        c.setFont("Helvetica", 12)
        c.drawString(100, 560, "Sample Student %d" % (i + 1))
        c.showPage()
    c.save()
    print("wrote", os.path.relpath(path, HERE))


XL = ("Microsoft Excel (Microsoft 365 Apps and O", "ce 2019)")     # ligature dropped
XL365 = ("Microsoft Excel (Microsoft 365 Apps)", None)
PP = ("Microsoft PowerPoint (Microsoft 365 Apps", "and Oce 2019)")

A = dict(name="Alex Cruz Demo", email="alex.demo@example.com", cid="alexdemo01")
B = dict(name="Bea Lopez Sample", email="bea.sample@example.com", cid="beasample02")
C = dict(name="Carl Reyes Test", email="carl.test@example.com", cid="carltest03")
D = dict(name="Dana Uy Fixture", email="dana.fixture@example.com", cid="danafix04")


def att(who, **kw):
    d = dict(who); d.update(kw); return d


pdf("01 - Demo High School/Exam 9001 Results.pdf", [
    att(A, title=XL, ref="90000001", date="March 3, 2026", score=875, areas=EXCEL, pcts=[80, 100, 100, 75, 75]),
    att(B, title=XL, ref="90000002", date="March 3, 2026", score=600, areas=EXCEL, pcts=[40, 60, 45, 70, 80], split_title=True),
    att(C, title=PP, ref="90000003", date="March 3, 2026", score=920, areas=PPT, pcts=[90, 100, 88, 92, 95], split_ref=True, gap_words=True),
])
pdf("01 - Demo High School/Exam 9001.1 Retake Results.pdf", [
    att(B, title=XL, ref="90000004", date="April 14, 2026", score=760, areas=EXCEL, pcts=[70, 80, 75, 70, 85]),
    att(A, title=XL, ref="90000001", date="March 3, 2026", score=875, areas=EXCEL, pcts=[80, 100, 100, 75, 75]),  # duplicate page
])
pdf("02 - Sample Academy/Exam 9002 Results 12B.pdf", [
    att(D, title=XL365, ref="90000005", date="September 9, 2025", score=1000, areas=EXCEL, pcts=[100, 100, 100, 100, 100], no_req_label=True),
    att(dict(name="Eve Tan Mock", email="eve.mock@example.com", cid="evemock05"), title=XL365, ref="90000006", date="September 9, 2025", score=0, areas=EXCEL, pcts=[0, 0, 0, 0, 0]),
])
pdf("02 - Sample Academy/Exam 9002 Certificates.pdf", [], extra_text_pages=2)
