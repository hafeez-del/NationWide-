"""
QB Purchase Merger - merges several QuickBooks "Purchases by Vendor Detail" exports into one workbook
(same output layout as the QB export, so the demand-planning script reads it as the 6-month purchase file).

Typical use: old 6-month file + one new month  ->  one new combined file.
If two files cover the same dates (e.g. the new month starts before the old file ended), the file that was
EXPORTED LATER wins for those dates and the older rows for the same dates are dropped (no double counting).
"""
import sys
import os
import re
import glob
import calendar
from datetime import datetime, date, timedelta
from collections import defaultdict

import openpyxl
import xlsxwriter

# ============================================================
#  Only used when you run this file directly from the command line (not in Streamlit)
# ============================================================
FOLDER_PATH = r"E:\Downloads\Automate\nationwise\New folder"
# 0 / None = keep every month that is in the files.
# 6 = keep only the newest 6 calendar months (e.g. April-October files -> output starts 1 May), so the
#     "6 month" file stays 6 months long when you add a new month.
KEEP_LAST_MONTHS = None
# ============================================================

HEADERS = ["Date", "Transaction Type", "Num", "Product/Service", "Memo/Description", "Qty", "Rate", "Amount", "Balance"]
COMPANY = "Nation Wide Whole Sale LLC"
REPORT = "Purchases by  Vendor Detail"
OUT_PREFIX = "merged_purchases_"
SHEET = "Purchases by Vendor Detail"

MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_name) if m}
MONTHS.update({m.lower(): i for i, m in enumerate(calendar.month_abbr) if m})


def month_no(word):
    w = word.strip().strip(".,").lower()
    if w in MONTHS:
        return MONTHS[w]
    raise ValueError(word)


def parse_date(v):
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime.combine(v, datetime.min.time())
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
        f = float(str(v).replace(",", "").replace("$", "").strip())
        return int(f) if f.is_integer() else f
    except ValueError:
        return 0


def parse_period(label):
    """QB report period text -> (first day, last day) or None."""
    s = re.sub(r"\s+", " ", str(label or "")).strip()
    if not s:
        return None
    try:
        m = re.fullmatch(r"(\d{1,2}/\d{1,2}/\d{4}) ?[-–] ?(\d{1,2}/\d{1,2}/\d{4})", s)
        if m:
            return parse_date(m[1]).date(), parse_date(m[2]).date()
        m = re.fullmatch(r"([A-Za-z]+) (\d{1,2}),? (\d{4}) ?[-–] ?([A-Za-z]+) (\d{1,2}),? (\d{4})", s)
        if m:
            return date(int(m[3]), month_no(m[1]), int(m[2])), date(int(m[6]), month_no(m[4]), int(m[5]))
        m = re.fullmatch(r"([A-Za-z]+) (\d{1,2}) ?[-–] ?([A-Za-z]+) (\d{1,2}),? (\d{4})", s)
        if m:
            return date(int(m[5]), month_no(m[1]), int(m[2])), date(int(m[5]), month_no(m[3]), int(m[4]))
        m = re.fullmatch(r"([A-Za-z]+) (\d{1,2}) ?[-–] ?(\d{1,2}),? (\d{4})", s)
        if m:
            mo = month_no(m[1])
            return date(int(m[4]), mo, int(m[2])), date(int(m[4]), mo, int(m[3]))
        m = re.fullmatch(r"([A-Za-z]+) (\d{1,2}),? (\d{4})", s)
        if m:
            d = date(int(m[3]), month_no(m[1]), int(m[2]))
            return d, d
        m = re.fullmatch(r"([A-Za-z]+),? ?[-–] ?([A-Za-z]+),? (\d{4})", s)
        if m:
            y, a, b = int(m[3]), month_no(m[1]), month_no(m[2])
            return date(y, a, 1), date(y, b, calendar.monthrange(y, b)[1])
        m = re.fullmatch(r"([A-Za-z]+),? (\d{4})", s)
        if m:
            y, a = int(m[2]), month_no(m[1])
            return date(y, a, 1), date(y, a, calendar.monthrange(y, a)[1])
    except (ValueError, KeyError):
        return None
    return None


def parse_stamp(footer):
    """'Thursday, Oct 08, 2026 07:36:12 PM GMT-7 - Accrual Basis' -> datetime (local clock of the export)."""
    m = re.search(r"([A-Za-z]{3,9}) (\d{1,2}), (\d{4}) (\d{1,2}):(\d{2})(?::(\d{2}))? ?([AP]M)", footer or "")
    if not m:
        return None
    try:
        return datetime.strptime(f"{m[1][:3]} {m[2]} {m[3]} {m[4]}:{m[5]}:{m[6] or '00'} {m[7]}", "%b %d %Y %I:%M:%S %p")
    except ValueError:
        return None


def read_file(path):
    """Return dict(rows=[(vendor, [B..J]), ...], footer, period, title)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    hdr = None
    for i, r in enumerate(rows[:30]):
        if len(r) > 1 and str(r[1]).strip().lower() == "date":
            hdr = i
            break
    if hdr is None:
        raise ValueError("Header row (Date) nahi mila - ye 'Purchases by Vendor Detail' export nahi lagta")
    if not any("vendor" in str(r[0]).lower() for r in rows[:5] if r and r[0]):
        raise ValueError("Title 'Purchases by Vendor Detail' nahi mila (sales/other report to nahi?)")
    title = str(rows[2][0]).strip() if len(rows) > 2 and rows[2] and rows[2][0] else ""

    out, footer, vendor, grand = [], None, None, None
    for r in rows[hdr + 1:]:
        r = list(r) + [None] * (10 - len(r))
        a, b = r[0], r[1]
        if a is not None and str(a).strip():
            t = str(a).strip()
            if t.startswith("Total for "):
                continue
            if t == "TOTAL":
                grand = num(r[8])
                continue
            if re.search(r"(Accrual|Cash) Basis", t):
                footer = t
                continue
            if b is None:                       # vendor heading
                vendor = t
                continue
        if b is not None and str(b).strip() and vendor is not None:
            out.append((vendor, r[1:10]))       # B..J (J = old Balance, recomputed)
    if grand is not None:
        amt = sum(num(x[1][7]) for x in out)
        if abs(amt - grand) > 0.01:
            print(f"   WARNING: {os.path.basename(path)} ka TOTAL match nahi hua (file: {grand:,.2f} | rows: {amt:,.2f})")
    dates = []
    for _, row in out:
        try:
            dates.append(parse_date(row[0]).date())
        except ValueError:
            raise ValueError(f"Date samajh nahi aayi: {row[0]!r}")
    period = parse_period(title)
    lo, hi = (min(dates), max(dates)) if dates else (None, None)
    if period is None and lo is None:
        raise ValueError("Na report period mila na koi transaction date")
    cover = [x for x in (period[0] if period else None, lo) if x]
    cover_end = [x for x in (period[1] if period else None, hi) if x]
    return {"rows": out, "footer": footer, "title": title, "start": min(cover), "end": max(cover_end),
            "stamp": parse_stamp(footer) or datetime.fromtimestamp(os.path.getmtime(path)), "name": os.path.basename(path)}


def month_back(first_of_month, n):
    y, m = first_of_month.year, first_of_month.month - n
    while m <= 0:
        m += 12
        y -= 1
    return date(y, m, 1)


def run(files, output_dir, keep_last_months=None):
    """Merge the given QB purchase files (list of paths) into one workbook inside output_dir.

    Returns {"output": path, "vendors": n, "rows": n, "skipped": [names], "dropped": n}.
    Raises ValueError if nothing could be read.
    """
    files = list(files)
    if not files:
        raise ValueError("No .xlsx file given")
    keep_last_months = KEEP_LAST_MONTHS if keep_last_months is None else keep_last_months
    os.makedirs(output_dir, exist_ok=True)
    skipped, loaded = [], []
    for f in sorted(files):
        try:
            info = read_file(f)
        except Exception as e:
            print(f"SKIP {os.path.basename(f)}: {e}")
            skipped.append(os.path.basename(f))
            continue
        print(f"{info['name']}: {len(info['rows']):,} rows | period {info['start']:%d-%b-%Y} to {info['end']:%d-%b-%Y} | exported {info['stamp']:%d-%b-%Y %H:%M}")
        loaded.append(info)
    if not loaded:
        raise ValueError("No data found - none of the files could be read as a QB 'Purchases by Vendor Detail' export")

    # Newest export first; its dates are taken from it and older files lose their rows for those dates.
    loaded.sort(key=lambda x: (x["stamp"], x["name"]), reverse=True)
    covered, kept, dropped, footer = [], [], 0, loaded[0]["footer"]
    seq = 0
    for info in loaded:
        n_drop = 0
        for vendor, row in info["rows"]:
            d = parse_date(row[0])
            if any(s <= d.date() <= e for s, e in covered):
                n_drop += 1
                continue
            seq += 1
            kept.append((vendor, d, seq, row))
        covered.append((info["start"], info["end"]))
        if n_drop:
            print(f"   {info['name']}: {n_drop:,} rows dropped (same dates exist in a newer export)")
        dropped += n_drop
    kept.sort(key=lambda x: (x[1], x[2]))

    start = min(i["start"] for i in loaded)
    end = max(i["end"] for i in loaded)
    if keep_last_months and int(keep_last_months) > 0:
        cut = month_back(date(end.year, end.month, 1), int(keep_last_months) - 1)
        if cut > start:
            before = len(kept)
            kept = [k for k in kept if k[1].date() >= cut]
            print(f"   KEEP_LAST_MONTHS={keep_last_months}: rows before {cut:%d-%b-%Y} removed ({before - len(kept):,})")
            start = cut
    if not kept:
        raise ValueError("No transactions left after merging")

    data = defaultdict(list)
    for vendor, d, s, row in kept:
        data[vendor].append((d, s, row))
    if (start.year, start.month) == (end.year, end.month):
        period = f"{start:%B}, {start.year}"
    elif start.year == end.year:
        period = f"{start:%B} - {end:%B}, {start.year}"
    else:
        period = f"{start:%B, %Y} - {end:%B, %Y}"

    base = f"{OUT_PREFIX}{start:%b}-{end:%b}_{end.year}"
    out_path = os.path.join(output_dir, base + ".xlsx")
    try:
        wb = xlsxwriter.Workbook(out_path, {"constant_memory": True})
        ws = wb.add_worksheet(SHEET)
    except (PermissionError, OSError):
        out_path = os.path.join(output_dir, f"{base}_{datetime.now():%H%M%S}.xlsx")
        print(f"   (file open in Excel - saving as {os.path.basename(out_path)})")
        wb = xlsxwriter.Workbook(out_path, {"constant_memory": True})
        ws = wb.add_worksheet(SHEET)

    def F(**k):
        b = {"font_name": "Arial", "font_size": 8, "text_wrap": True, "valign": "bottom"}
        b.update(k)
        return wb.add_format(b)

    f_t1 = F(bold=True, font_size=14, align="center")
    f_t2 = F(bold=True, font_size=12, align="center")
    f_t3 = F(bold=True, font_size=10, align="center")
    f_hdr = F(bold=True, font_size=9, align="center", bottom=1)
    f_ven = F(align="left")
    f_txt = F()
    f_num = F(num_format="#,##0.00")
    f_int = F(num_format="0")
    f_tot_l = F(bold=True, align="left")
    f_tot_a = F(bold=True, num_format="#,##0.00", top=1)
    f_foot = F(align="center")

    for c, w in zip("ABCDEFGHIJ", [40, 12, 20, 10, 70, 20, 10, 10, 14, 16]):
        ws.set_column(f"{c}:{c}", w)
    ws.merge_range("A1:J1", COMPANY, f_t1)
    ws.merge_range("A2:J2", REPORT, f_t2)
    ws.merge_range("A3:J3", period, f_t3)
    ws.write_row(4, 1, HEADERS, f_hdr)

    r, grand = 5, 0.0
    for vendor in sorted(data, key=lambda s: s.casefold()):
        ws.write(r, 0, vendor, f_ven)
        r += 1
        first = r + 1
        bal = amt_sum = 0.0
        for d, _, row in sorted(data[vendor], key=lambda x: (x[0], x[1])):
            date_s, typ, n, prod, memo, qty, rate, amt = row[:8]
            bal += num(amt)
            amt_sum += num(amt)
            ws.write(r, 1, d.strftime("%m/%d/%Y"), f_txt)
            ws.write(r, 2, typ, f_txt)
            if isinstance(n, (int, float)) and float(n).is_integer():
                ws.write_number(r, 3, int(n), f_int)          # keep Bill numbers numeric, no '.0'
            elif n is None:
                ws.write_blank(r, 3, None, f_txt)
            else:
                ws.write_string(r, 3, str(n), f_txt)
            ws.write(r, 4, prod, f_txt)
            if isinstance(memo, float) and memo.is_integer():
                memo = str(int(memo))
            ws.write(r, 5, memo if memo is None else str(memo), f_txt)
            for col, v in ((6, qty), (7, rate), (8, amt)):
                if v is None:
                    ws.write_blank(r, col, None, f_num)
                else:
                    ws.write_number(r, col, num(v), f_num)
            ws.write_number(r, 9, round(bal, 10), f_num)
            r += 1
        last = r
        ws.write(r, 0, f"Total for {vendor}", f_tot_l)
        ws.write_formula(r, 8, f"=SUM(I{first}:I{last})", f_tot_a, round(amt_sum, 10))
        grand += amt_sum
        r += 1
    ws.write(r, 0, "TOTAL", f_tot_l)
    ws.write_formula(r, 8, '=SUMIF(A6:A%d,"Total for*",I6:I%d)' % (r, r), f_tot_a, round(grand, 10))
    r += 3
    ws.merge_range(r, 0, r, 9, footer or f"Accrual Basis {datetime.now():%A, %B %d, %Y %I:%M %p}", f_foot)
    wb.close()

    types = defaultdict(int)
    for v in data.values():
        for _, _, row in v:
            types[row[1]] += 1
    print("Transaction types:", dict(types))
    print(f"\nDone: {len(data)} vendors, {len(kept):,} rows (older duplicate dates dropped: {dropped:,}). Period: {period}")
    print("Output:", out_path)
    return {"output": out_path, "vendors": len(data), "rows": len(kept), "skipped": skipped, "dropped": dropped}


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else FOLDER_PATH
    if not os.path.isdir(folder):
        sys.exit(f"Folder not found: {folder}")
    files = [f for f in glob.glob(os.path.join(folder, "*.xlsx"))
             if not os.path.basename(f).startswith(("~$", OUT_PREFIX, "merged_sales_"))]
    if not files:
        sys.exit("No .xlsx file found in the folder")
    try:
        run(sorted(files), folder)
    except ValueError as exc:
        sys.exit(str(exc))


if __name__ == "__main__":
    main()
