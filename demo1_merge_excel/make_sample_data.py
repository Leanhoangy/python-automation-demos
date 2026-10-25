"""Generate messy sample sales files (fake data) to demo the merge script."""
import random
from pathlib import Path
import pandas as pd

random.seed(7)
OUT = Path(__file__).parent / "input"
OUT.mkdir(exist_ok=True)
products = {"Laptop Stand": 29.9, "USB-C Hub": 39.5, "Wireless Mouse": 19.9,
            "Mechanical Keyboard": 79.0, "Webcam HD": 49.0, "Monitor Light": 34.5}
branches = ["Hanoi", "Da Nang", "Ho Chi Minh", "Can Tho", "Hai Phong"]
names = ["An", "Binh", "Chi", "Dung", "Giang", "Huy", "Khanh", "Lan"]

def rows(n, m):
    out = []
    for _ in range(n):
        p = random.choice(list(products))
        out.append((random.randint(1, 24), m, p, random.randint(1, 12), products[p],
                    random.choice(names)))
    return out

# Each branch uses different column names / formats on purpose
for i, b in enumerate(branches):
    r = rows(random.randint(35, 50), random.choice([8, 9, 10]))
    if i == 0:   # standard
        df = pd.DataFrame([(f"2026-{m:02d}-{d:02d}", p, q, u, s) for d, m, p, q, u, s in r],
                          columns=["Date", "Product", "Qty", "Unit Price", "Salesperson"])
        df["Product"] = df["Product"].where(df.index % 9 != 0, "  " + df["Product"] + " ")  # stray spaces
        df.to_excel(OUT / f"sales_{b.lower().replace(' ', '_')}.xlsx", index=False)
    elif i == 1: # different headers, dd/mm/yyyy
        df = pd.DataFrame([(f"{d:02d}/{m:02d}/2026", p, q, u, s) for d, m, p, q, u, s in r],
                          columns=["Ngay", "San pham", "So luong", "Don gia", "Nhan vien"])
        df.to_excel(OUT / "doanh_so_da_nang.xlsx", index=False)
    elif i == 2: # CSV, price as text with $ and commas
        df = pd.DataFrame([(f"{m:02d}-{d:02d}-2026", p.upper(), q, f"${u:,.2f}", s) for d, m, p, q, u, s in r],
                          columns=["date", "product", "quantity", "price", "seller"])
        df = pd.concat([df, df.iloc[:4]])            # duplicated rows
        df.to_csv(OUT / "hcm_sales.csv", index=False)
    elif i == 3: # missing values
        df = pd.DataFrame([(f"2026-{m:02d}-{d:02d}", p, q, u, s) for d, m, p, q, u, s in r],
                          columns=["Date", "Product", "Qty", "Unit Price", "Salesperson"])
        df.loc[[3, 11, 20], "Qty"] = None
        df.loc[[6], "Product"] = None
        df.to_excel(OUT / "can_tho_oct.xlsx", index=False)
    else:        # headers start on row 3
        df = pd.DataFrame([(f"2026-{m:02d}-{d:02d}", p, q, u, s) for d, m, p, q, u, s in r],
                          columns=["Date", "Product", "Qty", "Unit Price", "Salesperson"])
        with pd.ExcelWriter(OUT / "haiphong_report.xlsx", engine="openpyxl") as w:
            df.to_excel(w, index=False, startrow=2)
            w.sheets["Sheet1"].cell(1, 1, "Hai Phong branch - monthly sales")
print("Sample files written to", OUT)
