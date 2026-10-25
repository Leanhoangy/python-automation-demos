"""Generate fake invoice PDFs (made-up companies) + ground-truth JSON for testing the extractor."""
import json, random
from datetime import date, timedelta
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer

random.seed(11)
OUT = Path(__file__).parent / "input"
OUT.mkdir(exist_ok=True)
vendors = ["Blue Harbor Supplies", "Northwind Office Co.", "Sunrise Packaging Ltd", "Greenleaf Stationery",
           "Delta Print Works", "Pioneer IT Trading", "Lotus Cleaning Services", "Orion Logistics"]
items = [("A4 Paper (box)", 4.5), ("Printer Ink Cartridge", 28.0), ("Desk Lamp", 17.9), ("Packing Tape (roll)", 1.8),
         ("Wireless Keyboard", 24.5), ("Monitor Cable", 6.2), ("Whiteboard Marker (set)", 5.4), ("Delivery Fee", 12.0)]
styles = getSampleStyleSheet()
truth = []
for i, v in enumerate(vendors, 1):
    inv_no = f"INV-2026-{1000 + i * 7}"
    d = date(2026, 9, 1) + timedelta(days=random.randint(0, 40))
    due = d + timedelta(days=30)
    lines = []
    for name, price in random.sample(items, random.randint(2, 5)):
        qty = random.randint(1, 20)
        lines.append((name, qty, price, round(qty * price, 2)))
    sub = round(sum(l[3] for l in lines), 2)
    vat = round(sub * 0.08, 2)
    total = round(sub + vat, 2)
    truth.append({"file": f"invoice_{i:02d}.pdf", "vendor": v, "invoice_no": inv_no, "date": d.isoformat(),
                  "due_date": due.isoformat(), "subtotal": sub, "vat": vat, "total": total, "n_lines": len(lines)})
    doc = SimpleDocTemplate(str(OUT / f"invoice_{i:02d}.pdf"), pagesize=A4, leftMargin=40, rightMargin=40, topMargin=40)
    el = [Paragraph(f"<b>{v}</b>", styles["Title"]), Paragraph("INVOICE", styles["Heading2"]), Spacer(1, 8),
          Paragraph(f"Invoice No: {inv_no}", styles["Normal"]),
          Paragraph(f"Invoice Date: {d.strftime('%d/%m/%Y')}", styles["Normal"]),
          Paragraph(f"Due Date: {due.strftime('%d/%m/%Y')}", styles["Normal"]),
          Paragraph("Bill To: Sample Customer Co., Ltd", styles["Normal"]), Spacer(1, 14)]
    data = [["Description", "Qty", "Unit Price", "Amount"]] + [[n, str(q), f"${p:,.2f}", f"${a:,.2f}"] for n, q, p, a in lines]
    t = Table(data, colWidths=[270, 50, 90, 90])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#14273B")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                           ("GRID", (0, 0), (-1, -1), 0.4, colors.grey), ("ALIGN", (1, 0), (-1, -1), "RIGHT")]))
    el += [t, Spacer(1, 14), Paragraph(f"Subtotal: ${sub:,.2f}", styles["Normal"]),
           Paragraph(f"VAT (8%): ${vat:,.2f}", styles["Normal"]), Paragraph(f"<b>Total Due: ${total:,.2f}</b>", styles["Normal"])]
    doc.build(el)
(OUT / "_ground_truth.json").write_text(json.dumps(truth, indent=1))
print(f"{len(truth)} invoices written to {OUT}")
