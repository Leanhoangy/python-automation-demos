"""Render before/after preview image from the real input/output files."""
import base64, html
from pathlib import Path
import pandas as pd
from playwright.sync_api import sync_playwright

here = Path(__file__).parent
img = base64.b64encode((here / "inv1-1.png").read_bytes()).decode()
rep = pd.read_excel(here / "output/invoices.xlsx", sheet_name=None)
inv = rep["Invoices"][["vendor", "invoice_no", "date", "due_date", "subtotal", "vat", "total", "check_ok"]].copy()
inv["date"] = inv["date"].astype(str).str[:10]
inv["due_date"] = inv["due_date"].astype(str).str[:10]
inv["check_ok"] = inv["check_ok"].map({True: "OK", False: "REVIEW"})
inv = inv.rename(columns={"invoice_no": "invoice", "due_date": "due"})
n_inv, n_lines, total = len(rep["Invoices"]), len(rep["Line Items"]), rep["Invoices"]["total"].sum()

def tbl(df):
    h = "".join(f"<th>{html.escape(c)}</th>" for c in df.columns)
    b = "".join("<tr>" + "".join(f"<td>{html.escape(str(v))}</td>" for v in r) + "</tr>" for r in df.itertuples(index=False))
    return f"<table><tr>{h}</tr>{b}</table>"

page = f"""<!doctype html><html><head><meta charset='utf-8'><style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{width:1600px;height:1200px;background:#0d1b2a;color:#f4f7fb;font-family:Inter,sans-serif;padding:60px 80px}}
h1{{font-size:56px;font-weight:800}} h1 span{{color:#34d3b4}} .sub{{color:#9db1c7;font-size:26px;margin:10px 0 36px}}
.cols{{display:flex;gap:40px}} .tag{{font-size:22px;letter-spacing:.15em;text-transform:uppercase;font-weight:700;margin-bottom:16px}}
.b .tag{{color:#ff8a7a}} .a .tag{{color:#34d3b4}} .b{{width:560px}} .a{{flex:1}}
.pdf{{background:#fff;border-radius:10px;padding:10px}} .pdf img{{width:100%;display:block}}
.f{{background:#14273b;border:1px solid #24405c;border-radius:14px;padding:20px}}
table{{border-collapse:collapse;width:100%;font-size:16px}} th{{text-align:left;color:#9db1c7;font-weight:600;padding:6px 8px;border-bottom:1px solid #24405c;white-space:nowrap}}
td{{padding:7px 8px;color:#dbe6f2;white-space:nowrap}}
.stat{{display:flex;gap:16px;margin-top:20px}} .s{{flex:1;background:#14273b;border:1px solid #24405c;border-radius:14px;padding:18px}}
.s b{{display:block;font-size:38px;color:#34d3b4;white-space:nowrap}} .s span{{font-size:18px;color:#9db1c7}}
.note{{color:#9db1c7;font-size:19px;margin-top:14px;line-height:1.4}}
</style></head><body>
<h1>{n_inv} PDF invoices <span>→ 1 Excel file</span></h1>
<div class='sub'>No more typing numbers by hand. Every invoice is checked: line items must add up to the subtotal</div>
<div class='cols'><div class='b'><div class='tag'>Before: PDF invoice</div><div class='pdf'><img src='data:image/png;base64,{img}'></div>
<div class='note'>{n_inv} files like this one, each with a different vendor, date and items</div></div>
<div class='a'><div class='tag'>After: invoices.xlsx</div><div class='f'>{tbl(inv)}</div>
<div class='stat'><div class='s'><b>{n_inv}</b><span>invoices read</span></div><div class='s'><b>{n_lines}</b><span>line items extracted</span></div>
<div class='s'><b>${total:,.2f}</b><span>total value</span></div></div>
<div class='note'>Second sheet "Line Items" lists every product, quantity and price. Anything that does not add up is flagged REVIEW.</div></div></div>
</body></html>"""
(here / "before_after.html").write_text(page)
with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page(viewport={'width': 1600, 'height': 1200})
    pg.goto(f'file://{here}/before_after.html'); pg.screenshot(path=str(here / 'before_after.png')); b.close()
