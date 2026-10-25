"""Render before/after preview image from the real input/output files."""
import html
from pathlib import Path
import pandas as pd
import merge_sales as m
from playwright.sync_api import sync_playwright

here = Path(__file__).parent
files = sorted(p for p in (here/"input").iterdir() if p.suffix in {".xlsx", ".csv"})
_, stats = m.clean(pd.concat([m.read_file(p) for p in files], ignore_index=True))
rep = pd.read_excel(here/"output/sales_report.xlsx", sheet_name=None)
raw = {"sales_hanoi.xlsx": pd.read_excel(here/"input/sales_hanoi.xlsx").head(4),
       "doanh_so_da_nang.xlsx": pd.read_excel(here/"input/doanh_so_da_nang.xlsx").head(4),
       "hcm_sales.csv": pd.read_csv(here/"input/hcm_sales.csv").head(4)}

def cell(v):
    return html.escape(v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else str(v))
def tbl(df, n=None):
    d = df if n is None else df.head(n)
    h = "".join(f"<th>{html.escape(str(c))}</th>" for c in d.columns)
    b = "".join("<tr>" + "".join(f"<td>{cell(v)}</td>" for v in r) + "</tr>" for r in d.itertuples(index=False))
    return f"<table><tr>{h}</tr>{b}</table>"

before = "".join(f"<div class='f'><div class='fn'>{k}</div>{tbl(v)}</div>" for k, v in raw.items())
page = f"""<!doctype html><html><head><meta charset='utf-8'><style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{width:1600px;height:1200px;background:#0d1b2a;color:#f4f7fb;font-family:Inter,sans-serif;padding:60px 80px}}
h1{{font-size:56px;font-weight:800}} h1 span{{color:#34d3b4}} .sub{{color:#9db1c7;font-size:26px;margin:10px 0 36px}}
.cols{{display:flex;gap:40px}} .col{{flex:1}} .tag{{font-size:22px;letter-spacing:.15em;text-transform:uppercase;font-weight:700;margin-bottom:16px}}
.b .tag{{color:#ff8a7a}} .a .tag{{color:#34d3b4}}
.f{{background:#14273b;border:1px solid #24405c;border-radius:14px;padding:16px;margin-bottom:16px}}
.fn{{font-size:20px;color:#7cc4ff;margin-bottom:8px;font-weight:600}}
table{{border-collapse:collapse;width:100%;font-size:17px}} th{{text-align:left;color:#9db1c7;font-weight:600;padding:5px 8px;border-bottom:1px solid #24405c;white-space:nowrap}}
td{{padding:5px 8px;color:#dbe6f2;white-space:nowrap}} .a .f{{padding:20px}}
.note{{color:#9db1c7;font-size:19px;margin-top:8px;line-height:1.4}}
.stat{{display:flex;gap:16px;margin-top:20px}} .s{{flex:1;background:#14273b;border:1px solid #24405c;border-radius:14px;padding:18px}}
.s b{{display:block;font-size:38px;color:#34d3b4;white-space:nowrap}} .s span{{font-size:18px;color:#9db1c7}}
</style></head><body>
<h1>5 messy files <span>→ 1 clean report</span></h1>
<div class='sub'>Different headers, date formats, duplicates and missing values, merged in seconds</div>
<div class='cols'><div class='col b'><div class='tag'>Before: 5 branch files</div>{before}
<div class='note'>Different column names, date formats (2026-08-01 / 01/08/2026 / 08-01-2026), prices stored as text ($29.90), stray spaces, UPPERCASE names, duplicate rows, missing values</div></div>
<div class='col a'><div class='tag'>After: sales_report.xlsx</div>
<div class='f'><div class='fn'>Sheet: By Branch</div>{tbl(rep['By Branch'].round(2))}</div>
<div class='f'><div class='fn'>Sheet: All Sales (first rows)</div>{tbl(rep['All Sales'], 5)}</div>
<div class='stat'><div class='s'><b>{stats['rows_in']} → {stats['rows_out']}</b><span>rows (clean)</span></div>
<div class='s'><b>{stats['duplicates_removed']}</b><span>duplicates removed</span></div>
<div class='s'><b>{stats['incomplete_removed']}</b><span>incomplete rows dropped</span></div></div>
</div></div></body></html>"""
(here/"before_after.html").write_text(page)
with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page(viewport={'width': 1600, 'height': 1200})
    pg.goto(f'file://{here}/before_after.html'); pg.screenshot(path=str(here/'before_after.png')); b.close()
print(stats)
