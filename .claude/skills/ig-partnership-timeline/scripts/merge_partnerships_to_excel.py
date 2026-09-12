#!/usr/bin/env python3
"""
merge_partnerships_to_excel.py -- Add "Went Live Date(s)", "# Partnership
Posts", "Post URL(s)", "Caption / Transcript" columns to existing
creator-profile-metrics tabs in an xlsx, using scan JSON files produced by
scan_partnerships.py, and optionally append a clearly-separated block of
Ad-Library-only creators that the organic scan missed.

The source xlsx is expected to look like scrape_profiles.py's
"Profiles Overview" sheet: row 1 = title, row 2 = headers (including a
"Username" column with values like "@handle"), row 3+ = one row per
creator. Adjust --header-row / --username-col if a given sheet differs.

Usage:
    python merge_partnerships_to_excel.py <source.xlsx> <output.xlsx> \
        --map "SheetName1=scan_brand1.json" \
        --map "SheetName2=scan_brand2.json" \
        [--ad-library-extra "SheetName1=ad_library_extras_brand1.json"]

Each --map value points at a JSON file from scan_partnerships.py.

Each --ad-library-extra value (optional) points at a JSON list of objects
shaped like:
    {"name": "creator name", "count": 3, "launch_date": "2026-01-10",
     "note": "found via Meta Ad Library search, not in organic scan"}
These are creators found ONLY through a Meta Ad Library cross-check, where
no individual post URL or caption is available. NEVER invent a URL or
caption for these rows -- leave those cells explicitly marked as
unavailable rather than guessing.
"""
import re
import json
import argparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SURROGATE_RE = re.compile('[\ud800-\udfff]')


def clean(s):
    """Strip lone UTF-16 surrogates (e.g. from a JS .slice() cutting a
    4-byte emoji in half) that otherwise crash openpyxl on save with
    'UnicodeEncodeError: surrogates not allowed'."""
    return SURROGATE_RE.sub('', s) if isinstance(s, str) else s


def bf(size=10, color="000000"):
    return Font(name="Calibri", bold=True, size=size, color=color)


def nf(size=10, italic=False):
    return Font(name="Calibri", bold=False, size=size, italic=italic)


def lnk(size=10):
    return Font(name="Calibri", size=size, color="0563C1", underline="single")


def ctr(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)


def lft(wrap=True):
    return Alignment(horizontal="left", vertical="center", wrap_text=wrap)


def fill(hexcolor):
    return PatternFill("solid", fgColor=hexcolor)


THIN = Side(style="thin", color="D0D0D0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

NEW_COLS = ["Went Live Date(s)", "# Partnership Posts", "Post URL(s)", "Caption / Transcript"]


def add_partnership_columns(ws, scan_data, header_row, username_col):
    hits = scan_data.get("hits", [])
    lookup = {}
    for h in hits:
        for uname_lower, (method, disp) in h["partners"].items():
            lookup.setdefault(uname_lower, []).append((h["date"], h["url"], h.get("caption", "")))
    for k in lookup:
        lookup[k].sort()

    start_col = ws.max_column + 1
    for i, h in enumerate(NEW_COLS):
        col = start_col + i
        c = ws.cell(row=header_row, column=col, value=h)
        c.font = bf(color="FFFFFF")
        c.fill = fill("1F4E78")
        c.alignment = ctr(wrap=True)
        c.border = BORDER
    ws.column_dimensions[get_column_letter(start_col)].width = 22
    ws.column_dimensions[get_column_letter(start_col + 1)].width = 12
    ws.column_dimensions[get_column_letter(start_col + 2)].width = 42
    ws.column_dimensions[get_column_letter(start_col + 3)].width = 90
    if ws.row_dimensions[header_row].height is None or ws.row_dimensions[header_row].height < 30:
        ws.row_dimensions[header_row].height = 30

    matched, unmatched = 0, 0
    last_data_row = ws.max_row
    for row in ws.iter_rows(min_row=header_row + 1, max_row=last_data_row):
        uname_cell = row[username_col - 1]
        if not uname_cell.value:
            continue
        uname_key = str(uname_cell.value).lstrip("@").strip().lower()
        posts = lookup.get(uname_key, [])
        r = uname_cell.row
        if posts:
            matched += 1
            dates = "\n".join(p[0] for p in posts)
            urls = "\n".join(p[1] for p in posts)
            captions = "\n\n".join(
                (clean(p[2])[:300] if p[2] else "(no caption text)") for p in posts
            )
            count = len(posts)
        else:
            unmatched += 1
            dates, urls, captions, count = "Not found in scan", "", "", 0

        vals = [dates, count, urls, captions]
        n_lines = max(1, len(posts))
        for i, val in enumerate(vals):
            col = start_col + i
            c = ws.cell(row=r, column=col, value=val)
            c.border = BORDER
            if col == start_col + 2 and len(posts) == 1:
                c.font = lnk()
                c.alignment = lft(wrap=True)
                if posts[0][1]:
                    c.hyperlink = posts[0][1]
            elif col == start_col + 1:
                c.font = nf()
                c.alignment = ctr()
            else:
                c.font = nf()
                c.alignment = lft(wrap=True)
        ws.row_dimensions[r].height = max(20, 16 * n_lines)

    return matched, unmatched, start_col, last_data_row


def append_ad_library_extras(ws, extras, start_col, last_data_row):
    if not extras:
        return
    start = last_data_row + 2
    end_col_letter = get_column_letter(start_col + 3)
    ws.merge_cells(f"A{start}:{end_col_letter}{start}")
    ws[f"A{start}"] = (
        "Below: partners found via Meta Ad Library cross-check but ABSENT from the organic Instagram "
        "scan above. Only partner-level name/count/date is available for these rows (no individual post "
        "URL or caption) -- blank Post URL/Caption cells here mean genuinely unavailable data, not missing work."
    )
    ws[f"A{start}"].font = Font(name="Calibri", bold=True, italic=True, size=9, color="9C0006")
    ws[f"A{start}"].alignment = lft()
    ws.row_dimensions[start].height = 32

    header_row = start + 1
    headers = ["#", "Partner Name", "Post URL", "Went Live Date", "Note", "Live Ads (Ad Library)"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=header_row, column=col, value=h)
        c.font = bf()
        c.fill = fill("FCE4E4")
        c.alignment = ctr()
        c.border = BORDER
    ws.row_dimensions[header_row].height = 22

    for i, extra in enumerate(extras, 1):
        r = header_row + i
        vals = [
            i,
            extra.get("name", ""),
            "N/A -- not captured via organic scan",
            extra.get("launch_date", ""),
            extra.get("note", ""),
            extra.get("count", ""),
        ]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=r, column=col, value=val)
            c.border = BORDER
            if col == 2:
                c.font = bf()
                c.fill = fill("FFF8F0")
                c.alignment = lft(wrap=False)
            elif col == 3:
                c.font = nf(italic=True)
                c.alignment = lft(wrap=False)
            elif col == 5:
                c.font = nf(italic=True)
                c.alignment = lft(wrap=True)
            else:
                c.font = nf()
                c.alignment = ctr()
        ws.row_dimensions[r].height = 24


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("source_xlsx")
    ap.add_argument("output_xlsx")
    ap.add_argument("--map", action="append", default=[], help="SheetName=scan_result.json (repeatable)")
    ap.add_argument("--ad-library-extra", action="append", default=[],
                     help="SheetName=ad_library_extras.json (repeatable, optional)")
    ap.add_argument("--header-row", type=int, default=2, help="Row containing column headers (default 2)")
    ap.add_argument("--username-col", type=int, default=2, help="1-based column index of Username (default 2)")
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.source_xlsx)

    extras_map = {}
    for item in args.ad_library_extra:
        sheet, path = item.split("=", 1)
        with open(path, encoding="utf-8") as f:
            extras_map[sheet] = json.load(f)

    if not args.map:
        print("ERROR: pass at least one --map SheetName=scan_result.json")
        return

    for item in args.map:
        sheet_name, json_path = item.split("=", 1)
        if sheet_name not in wb.sheetnames:
            print(f"WARNING: sheet '{sheet_name}' not found in workbook (have: {wb.sheetnames}), skipping")
            continue
        ws = wb[sheet_name]
        with open(json_path, encoding="utf-8") as f:
            scan_data = json.load(f)
        matched, unmatched, start_col, last_row = add_partnership_columns(
            ws, scan_data, args.header_row, args.username_col
        )
        print(f"{sheet_name}: matched={matched} unmatched={unmatched}")
        append_ad_library_extras(ws, extras_map.get(sheet_name, []), start_col, ws.max_row)

    wb.save(args.output_xlsx)
    print(f"\nSaved: {args.output_xlsx}")


if __name__ == "__main__":
    main()
