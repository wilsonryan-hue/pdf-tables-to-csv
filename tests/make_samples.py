"""Builds the sample PDFs used by the tests and the examples folder (reportlab)."""
from pathlib import Path
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet

HERE = Path(__file__).parent
STYLE = TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey)])
HEADER = ["Date", "Description", "Paid out", "Paid in", "Balance"]


def statement_rows():
    rows, bal = [], 1000.00
    for d in range(1, 61):
        out = 12.5 * (d % 4) if d % 3 else 0
        inn = 250.0 if d % 10 == 0 else 0
        bal = round(bal - out + inn, 2)
        rows.append([f"{(d - 1) % 28 + 1:02d}/0{(d - 1) // 28 + 1}/2026", f"Card payment {d}",
                     f"£{out:,.2f}" if out else "", f"£{inn:,.2f}" if inn else "", f"£{bal:,.2f}"])
    return rows


def build_statement(path):
    ss = getSampleStyleSheet()
    rows = statement_rows()
    story = [Paragraph("Example Bank plc - statement (sample data)", ss["Title"])]
    for i in range(0, len(rows), 30):
        t = Table([HEADER] + rows[i:i + 30])
        t.setStyle(STYLE)
        story += [t, PageBreak()]
    SimpleDocTemplate(str(path), pagesize=A4).build(story[:-1])
    return rows


def build_mixed(path):
    ss = getSampleStyleSheet()
    a = Table([["Product", "Unit price", "Qty"], ["Widget A", "£1,250.00", "2"],
               ["Widget B", "(15.00)", "1"], ["Service", "50.00 DR", "3"]])
    b = Table([["Region", "Customers"], ["North", "120"], ["South", "95"]])
    a.setStyle(STYLE); b.setStyle(STYLE)
    SimpleDocTemplate(str(path), pagesize=A4).build(
        [Paragraph("Price list", ss["Heading1"]), a, Spacer(1, 30),
         Paragraph("Customers", ss["Heading1"]), b])


if __name__ == "__main__":
    build_statement(HERE.parent / "examples" / "sample-statement.pdf")
    build_mixed(HERE.parent / "examples" / "sample-price-list.pdf")
    print("ok")
