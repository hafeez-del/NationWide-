"""
QB Sales Merger - merges several QuickBooks "Sales by Customer Detail" monthly exports into one workbook.
"""
import sys
import os
import re
import glob
from datetime import datetime
from collections import defaultdict

import openpyxl
import xlsxwriter

# ============================================================
#  Only used when you run this file directly from the command line (not in Streamlit)
# ============================================================
FOLDER_PATH = r"E:\Downloads\Automate\nationwise\New folder"
# ============================================================

HEADERS = ["Transaction date", "Transaction type", "Num", "Product/Service full name",
           "Description", "Quantity", "Sales price", "Amount", "Balance"]
COMPANY = "Nation Wide Whole Sale LLC"
REPORT = "Sales by Customer Detail"
OUT_PREFIX = "merged_sales_"


def parse_date(v):
    if isinstance(v, datetime):
        return v
    s = str(v).strip()
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y", "%d-%b-%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    raise ValueError(f"Date samajh nahi aayi: {v!r}")


def num(v):
    if v is None or v == "":
        return 0
    if isinstance(v, (int, float)):
        return v
    try:
        s = str(v).replace(",", "").replace("$", "").strip()
        f = float(s)
        return int(f) if f.is_integer() else f
    except ValueError:
        return 0


def read_file(path):
    """Return (list of (customer, [row,...]), footer_text) from one QB file."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    hdr = None
    for i, r in enumerate(rows):
        if len(r) > 1 and str(r[1]).strip().lower() == "transaction date":
            hdr = i
            break
    if hdr is None:
        raise ValueError("Header row (Transaction date) nahi mila")

    out, footer, cust, grand = [], None, None, None
    for r in rows[hdr + 1:]:
        r = list(r) + [None] * (10 - len(r))
        a, b = r[0], r[1]
        if a is not None and str(a).strip():
            t = str(a).strip()
            if t.startswith("Total for "):
                continue
            if t == "TOTAL":                    # grand total row of monthly file
                grand = (num(r[6]), num(r[8]))
                continue
            if t.startswith("Accrual Basis") or t.startswith("Cash Basis"):
                footer = t
                continue
            if b is None:                       # customer heading
                cust = str(a).strip()
                continue
        if b is not None and str(b).strip() and cust is not None:
            out.append((cust, r[1:10]))         # B..J (J = old Balance, recomputed)
    # check: file ke apne TOTAL se rows ka sum match hona chahiye
    if grand is not None:
        q = sum(num(x[1][5]) for x in out)
        a = sum(num(x[1][7]) for x in out)
        if abs(q - grand[0]) > 0.01 or abs(a - grand[1]) > 0.01:
            print(f"   WARNING: {os.path.basename(path)} ka TOTAL match nahi hua "
                  f"(file: qty {grand[0]:.2f} amt {grand[1]:.2f} | rows: qty {q:.2f} amt {a:.2f})")
    return out, footer


def run(files, output_dir):
    """Merge the given QB monthly files (list of paths) into one workbook inside output_dir.

    Returns {"output": path, "customers": n, "rows": n, "skipped": [names]}.
    Raises ValueError if nothing could be read.
    """
    files = list(files)
    if not files:
        raise ValueError("No .xlsx file given")
    os.makedirs(output_dir, exist_ok=True)
    skipped = []

    data = defaultdict(list)        # customer -> [(date, seq, row)]
    footer, seq, total_rows = None, 0, 0
    for f in sorted(files):
        try:
            recs, ft = read_file(f)
        except Exception as e:
            print(f"SKIP {os.path.basename(f)}: {e}")
            skipped.append(os.path.basename(f))
            continue
        print(f"{os.path.basename(f)}: {len(recs)} rows")
        total_rows += len(recs)
        for cust, row in recs:
            seq += 1
            data[cust].append((parse_date(row[0]), seq, row))
        footer = ft or footer
    if not data:
        raise ValueError("No data found - none of the files could be read as a QB 'Sales by Customer Detail' export")

    # period title from actual dates
    dates = [d for v in data.values() for d, _, _ in v]
    d0, d1 = min(dates), max(dates)
    if (d0.year, d0.month) == (d1.year, d1.month):
        period = f"{d0:%B}, {d0.year}"
    elif d0.year == d1.year:
        period = f"{d0:%B}-{d1:%B}, {d0.year}"
    else:
        period = f"{d0:%B, %Y}-{d1:%B, %Y}"

    out_path = os.path.join(output_dir, f"{OUT_PREFIX}{d0:%b}-{d1:%b}_{d1.year}.xlsx")
    wb = xlsxwriter.Workbook(out_path, {"constant_memory": True})
    ws = wb.add_worksheet("Sheet1")

    def F(**k):
        base = {"font_name": "Arial", "font_size": 8, "text_wrap": True, "valign": "bottom"}
        base.update(k)
        return wb.add_format(base)

    f_t1 = F(bold=True, font_size=14, align="center")
    f_t2 = F(bold=True, font_size=12, align="center")
    f_t3 = F(bold=True, font_size=10, align="center")
    f_hdr = F(bold=True, font_size=9, align="center", bottom=1)
    f_cust = F(align="left")
    f_txt = F()
    f_num = F(num_format="#,##0.00")
    f_tot_l = F(bold=True, align="left")
    f_tot_q = F(bold=True, num_format="#,##0.00", top=1)
    f_tot_a = F(bold=True, num_format='"$"#,##0.00', top=1)
    f_foot = F(align="center")

    for c, w in zip("ABCDEFGHIJ", [53, 14.375, 13, 7.5, 86.625, 13, 16.125, 10.125, 20.375, 13]):
        ws.set_column(f"{c}:{c}", w)

    ws.merge_range("A1:I1", COMPANY, f_t1)
    ws.merge_range("A2:I2", REPORT, f_t2)
    ws.merge_range("A3:I3", period, f_t3)
    ws.write_row(4, 1, HEADERS, f_hdr)

    r = 5                                # 0-based row index
    for cust in sorted(data, key=lambda s: s.casefold()):
        ws.write(r, 0, cust, f_cust)
        r += 1
        first = r + 1                    # excel row no.
        bal = 0.0
        qty_sum = amt_sum = 0.0
        for d, _, row in sorted(data[cust], key=lambda x: (x[0], x[1])):
            date_s, typ, n, prod, desc, qty, price, amt = row[:8]
            bal += num(amt)
            qty_sum += num(qty)
            amt_sum += num(amt)
            ws.write(r, 1, date_s if isinstance(date_s, str) else d.strftime("%m/%d/%Y"), f_txt)
            ws.write(r, 2, typ, f_txt)
            ws.write(r, 3, n if n is None else str(n), f_txt)
            ws.write(r, 4, prod, f_txt)
            ws.write(r, 5, desc if desc is None else str(desc), f_txt)
            for col, v in ((6, qty), (7, price), (8, amt)):
                if v is None:
                    ws.write_blank(r, col, None, f_num)
                else:
                    ws.write_number(r, col, num(v), f_num)
            ws.write_number(r, 9, round(bal, 10), f_num)
            r += 1
        last = r
        ws.write(r, 0, f"Total for {cust}", f_tot_l)
        ws.write_formula(r, 6, f"=SUM(G{first}:G{last})", f_tot_q, round(qty_sum, 10))
        ws.write_formula(r, 8, f"=SUM(I{first}:I{last})", f_tot_a, round(amt_sum, 10))
        r += 1

    r += 3                               # 3 blank rows like original
    ws.merge_range(r, 0, r, 8, footer or f"Accrual Basis {datetime.now():%A, %B %d, %Y %I:%M %p}", f_foot)
    wb.close()
    types = defaultdict(int)
    for v in data.values():
        for _, _, row in v:
            types[row[1]] += 1
    print("Transaction types:", dict(types))
    print(f"\nDone: {len(data)} customers, {total_rows} rows")
    print("Output:", out_path)
    return {"output": out_path, "customers": len(data), "rows": total_rows, "skipped": skipped}


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else FOLDER_PATH
    if not os.path.isdir(folder):
        sys.exit(f"Folder not found: {folder}")
    files = [f for f in glob.glob(os.path.join(folder, "*.xlsx"))
             if not os.path.basename(f).startswith(("~$", OUT_PREFIX))]
    if not files:
        sys.exit("No .xlsx file found in the folder")
    try:
        run(sorted(files), folder)
    except ValueError as exc:
        sys.exit(str(exc))


if __name__ == "__main__":
    main()
