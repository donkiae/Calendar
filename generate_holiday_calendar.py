r"""
generate_holiday_calendar.py
Generates APAC_Holiday_Dates.xlsx -- one sheet per year, 12-month grid.
Format matches the compact orange-header calendar style from the reference image.
Run: C:/anaconda3/python.exe generate_holiday_calendar.py
Deps: openpyxl (bundled with Anaconda)
"""

import re
import calendar
import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────────────────────
YEARS = [2025, 2026, 2027]

COUNTRIES = [
    ("SG", "Singapore",   "EF3340"),  # Singapore red
    ("MY", "Malaysia",    "003478"),  # Malaysia blue
    ("ID", "Indonesia",   "CE1126"),  # Indonesia red
    ("IN", "India",       "FF9933"),  # India saffron
    ("JP", "Japan",       "BC002D"),  # Japan crimson
    ("VN", "Vietnam",     "D4A017"),  # Vietnam gold
    ("TH", "Thailand",    "2D2A4A"),  # Thailand navy
    ("PH", "Philippines", "0038A8"),  # Philippines blue
]
COUNTRY_CODES  = [c[0] for c in COUNTRIES]
COUNTRY_COLORS = {c[0]: c[2] for c in COUNTRIES}  # code -> hex6

MONTH_NAMES = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December",
]

# ── Colours ──
COL_HEADER      = "C55A11"   # dark burnt-orange header
COL_MULTI_HOL   = "BFBFBF"   # gray when 2+ countries share a date
COL_WEEKEND_NUM = "CC0000"   # red numerals on Sat/Sun
COL_WHITE       = "FFFFFF"
COL_LIGHT_GRAY  = "F2F2F2"   # weekend cell background
COL_INK         = "000000"

TINT_AMOUNT = 0.22   # how much to blend country colour toward white (0=full colour, 1=white)

# ── Layout ──
MONTHS_PER_ROW = 3
DAY_COLS       = 7        # Mon → Sun
COL_GAP        = 2        # spacer columns between month blocks
HEADER_ROWS    = 2        # month-name row + day-label row
WEEK_ROWS      = 6        # max weeks per month
BLOCK_H        = HEADER_ROWS + WEEK_ROWS   # 8 rows per month block
ROW_GAP        = 2        # spacer rows between month-row groups
MARGIN_ROW     = 3        # rows above first month (title + blank)
MARGIN_COL     = 1        # columns left of first month (A = margin)


# ─────────────────────────────────────────────────────────────
#  STYLE HELPERS
# ─────────────────────────────────────────────────────────────
def fill(hex6: str) -> PatternFill:
    return PatternFill(fill_type="solid", fgColor=f"FF{hex6.upper()}")

def font(size=9, bold=False, color=COL_INK, name="Calibri") -> Font:
    return Font(name=name, size=size, bold=bold, color=f"FF{color}")

def border(color="D9D9D9") -> Border:
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def align(h="center", v="center", wrap=False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def tint(hex6: str, amount: float = TINT_AMOUNT) -> str:
    """Blend hex6 toward white. amount=0 → full colour, amount=1 → white."""
    r = int(hex6[0:2], 16)
    g = int(hex6[2:4], 16)
    b = int(hex6[4:6], 16)
    r = int(r * (1 - amount) + 255 * amount)
    g = int(g * (1 - amount) + 255 * amount)
    b = int(b * (1 - amount) + 255 * amount)
    return f"{r:02X}{g:02X}{b:02X}"


# ─────────────────────────────────────────────────────────────
#  PARSE holidays_data.js
# ─────────────────────────────────────────────────────────────
def load_holidays(path: str = "holidays_data.js") -> dict:
    text = Path(path).read_text(encoding="utf-8")
    result = {}
    valid = set(COUNTRY_CODES)

    for cm in re.finditer(r"\b([A-Z]{2})\s*:\s*\{", text):
        code = cm.group(1)
        if code not in valid:
            continue
        result[code] = {}
        depth, pos = 1, cm.end()
        while pos < len(text) and depth:
            if text[pos] == "{": depth += 1
            elif text[pos] == "}": depth -= 1
            pos += 1
        block = text[cm.end(): pos - 1]
        for ym in re.finditer(r"(\d{4})\s*:\s*\[(.*?)\]", block, re.DOTALL):
            yr = int(ym.group(1))
            entries = [
                {"date": m.group(1), "name": m.group(2)}
                for m in re.finditer(r'\{date:"([^"]+)",name:"([^"]+)"\}', ym.group(2))
            ]
            result[code][yr] = entries
    return result


def build_map(holidays: dict, year: int) -> dict:
    """'YYYY-MM-DD' -> list of country codes (preserves order) that have a holiday."""
    hmap: dict = {}
    for code in COUNTRY_CODES:
        for h in holidays.get(code, {}).get(year, []):
            if h["date"] not in hmap:
                hmap[h["date"]] = []
            if code not in hmap[h["date"]]:
                hmap[h["date"]].append(code)
    return hmap


def build_name_map(holidays: dict, year: int) -> dict:
    """'YYYY-MM-DD' -> list of (code, name) for cell comments."""
    nmap: dict = {}
    for code in COUNTRY_CODES:
        for h in holidays.get(code, {}).get(year, []):
            nmap.setdefault(h["date"], []).append((code, h["name"]))
    return nmap


def write_events_sheet(wb: Workbook, holidays: dict):
    ws = wb.create_sheet(title="Events")

    # Header
    headers = ["Date", "Day", "Country", "Holiday"]
    widths  = [12, 10, 16, 40]
    for ci, (hdr, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=1, column=ci, value=hdr)
        c.fill      = fill(COL_HEADER)
        c.font      = font(10, bold=True, color="FFFFFF")
        c.alignment = align("center", "center")
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 16

    row = 2
    for year in YEARS:
        # collect all holiday entries sorted by date then country order
        entries = []
        for code in COUNTRY_CODES:
            country_name = next(n for c2, n, _hex in COUNTRIES if c2 == code)
            for h in holidays.get(code, {}).get(year, []):
                entries.append((h["date"], code, country_name, h["name"]))
        entries.sort(key=lambda x: (x[0], COUNTRY_CODES.index(x[1])))

        # Year group header
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        yc = ws.cell(row=row, column=1, value=str(year))
        yc.fill      = fill("1A1A2E")
        yc.font      = font(10, bold=True, color="FFFFFF")
        yc.alignment = align("left", "center")
        ws.row_dimensions[row].height = 14
        row += 1

        for date_str, code, country_name, holiday_name in entries:
            d = datetime.date.fromisoformat(date_str)
            hex6 = COUNTRY_COLORS[code]

            ws.cell(row=row, column=1, value=d).number_format = "DD-MMM-YYYY"
            ws.cell(row=row, column=1).alignment = align("center", "center")
            ws.cell(row=row, column=1).font = font(9)

            ws.cell(row=row, column=2, value=d.strftime("%A")).alignment = align("center", "center")
            ws.cell(row=row, column=2).font = font(9)

            cc = ws.cell(row=row, column=3, value=country_name)
            cc.fill      = fill(tint(hex6, 0.55))
            cc.font      = font(9, bold=True, color=hex6)
            cc.alignment = align("center", "center")

            nc = ws.cell(row=row, column=4, value=holiday_name)
            nc.font      = font(9)
            nc.alignment = align("left", "center")

            bdr = border("D9D9D9")
            for ci in range(1, 5):
                ws.cell(row=row, column=ci).border = bdr
            ws.row_dimensions[row].height = 14
            row += 1

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"


# ─────────────────────────────────────────────────────────────
#  WRITE ONE YEAR SHEET
# ─────────────────────────────────────────────────────────────
def write_year(ws, year: int, holidays: dict):
    hmap  = build_map(holidays, year)
    nmap  = build_name_map(holidays, year)
    bdr  = border()
    bdr_none = Border()   # no border for spacers

    def block_col(mi):   # 0-based month-in-row index
        return MARGIN_COL + 1 + mi * (DAY_COLS + COL_GAP)

    def block_row(ri):   # 0-based row-of-months index
        return MARGIN_ROW + 1 + ri * (BLOCK_H + ROW_GAP)

    # ── column widths ──
    ws.column_dimensions["A"].width = 0.8
    for mi in range(MONTHS_PER_ROW):
        sc = block_col(mi)
        for ci in range(DAY_COLS):
            ws.column_dimensions[get_column_letter(sc + ci)].width = 4.2
        for gi in range(COL_GAP):
            ws.column_dimensions[get_column_letter(sc + DAY_COLS + gi)].width = 1.0

    # ── row heights ──
    ws.row_dimensions[1].height = 20  # title
    ws.row_dimensions[2].height = 6   # spacer
    for ri in range(4):
        sr = block_row(ri)
        ws.row_dimensions[sr    ].height = 16   # month header
        ws.row_dimensions[sr + 1].height = 13   # day labels
        for wr in range(WEEK_ROWS):
            ws.row_dimensions[sr + 2 + wr].height = 15

    # ── sheet title row ──
    title_end_col = block_col(MONTHS_PER_ROW - 1) + DAY_COLS - 1
    ws.merge_cells(start_row=1, start_column=2,
                   end_row=1,   end_column=title_end_col)
    tc = ws.cell(row=1, column=2)
    tc.value     = f"APAC Holiday Dates  {year}"
    tc.font      = font(size=13, bold=True, color="1A1A2E")
    tc.alignment = align("left", "center")

    # ── 12 month blocks ──
    for midx in range(12):
        ri = midx // MONTHS_PER_ROW
        mi = midx %  MONTHS_PER_ROW
        sr = block_row(ri)
        sc = block_col(mi)
        m  = midx + 1

        # abbreviated month-year label: "Jan-25"
        mo_abbr  = datetime.date(year, m, 1).strftime("%b")
        yr2      = str(year)[2:]
        hdr_text = f"{mo_abbr}-{yr2}"

        # ── month name header (merged 7 cols) ──
        ws.merge_cells(start_row=sr, start_column=sc,
                       end_row=sr,   end_column=sc + DAY_COLS - 1)
        hc = ws.cell(row=sr, column=sc)
        hc.value     = hdr_text
        hc.fill      = fill(COL_HEADER)
        hc.font      = font(size=10, bold=True, color="FFFFFF")
        hc.alignment = align("center", "center")
        for ci in range(1, DAY_COLS):
            ws.cell(row=sr, column=sc + ci).fill = fill(COL_HEADER)

        # ── day-of-week labels (same orange) ──
        day_labels = ["M", "T", "W", "T", "F", "S", "S"]
        for di, dlbl in enumerate(day_labels):
            c = ws.cell(row=sr + 1, column=sc + di)
            c.value     = dlbl
            c.fill      = fill(COL_HEADER)
            c.font      = font(size=8, bold=True, color="FFFFFF")
            c.alignment = align("center", "center")

        # ── date cells ──
        # monthcalendar: Monday=0 ... Sunday=6, 0 = outside this month
        weeks = calendar.monthcalendar(year, m)
        while len(weeks) < 6:
            weeks.append([0] * 7)

        for wi, week in enumerate(weeks):
            for di, day in enumerate(week):
                row = sr + 2 + wi
                col = sc + di
                c   = ws.cell(row=row, column=col)
                c.border    = bdr
                c.alignment = align("center", "center")

                if day == 0:
                    # out-of-month cell — blank, light fill
                    c.fill = fill("F7F7F7")
                    continue

                iso       = f"{year}-{m:02d}-{day:02d}"
                is_wkend  = di >= 5          # Sat=5, Sun=6
                codes     = hmap.get(iso, [])

                if codes:
                    if len(codes) == 1:
                        bg = tint(COUNTRY_COLORS[codes[0]])
                        fg = COUNTRY_COLORS[codes[0]]
                    else:
                        bg = COL_MULTI_HOL
                        fg = "444444"
                    c.fill  = fill(bg)
                    c.value = day
                    c.font  = font(size=9, bold=True, color=fg)
                    names = nmap.get(iso, [])
                    if names:
                        note = "\n".join(f"{cd}: {nm}" for cd, nm in names)
                        c.comment = Comment(note, "")
                elif is_wkend:
                    c.fill  = fill(COL_LIGHT_GRAY)
                    c.value = day
                    c.font  = font(size=9, bold=False, color=COL_WEEKEND_NUM)
                else:
                    c.fill  = fill(COL_WHITE)
                    c.value = day
                    c.font  = font(size=9, bold=False, color=COL_INK)

    # ── legend ──
    legend_row = block_row(3) + BLOCK_H + 1
    ws.row_dimensions[legend_row].height = 14

    col = block_col(0)
    lbl = ws.cell(row=legend_row, column=col)
    lbl.value     = "Legend:"
    lbl.font      = font(size=8, bold=True, color="595959")
    lbl.alignment = align("left", "center")
    col += 2

    # one swatch per country
    for code, name, hex6 in COUNTRIES:
        swatch = ws.cell(row=legend_row, column=col)
        swatch.fill   = fill(tint(hex6))
        swatch.border = border("AAAAAA")

        txt = ws.cell(row=legend_row, column=col + 1)
        txt.value = f"{code}"
        txt.font  = font(size=8, bold=True, color=hex6)
        txt.alignment = align("left", "center")
        col += 3

    # multi-country swatch
    swatch = ws.cell(row=legend_row, column=col)
    swatch.fill   = fill(COL_MULTI_HOL)
    swatch.border = border("AAAAAA")
    txt = ws.cell(row=legend_row, column=col + 1)
    txt.value = "Multi"
    txt.font  = font(size=8, bold=False, color="595959")
    txt.alignment = align("left", "center")
    col += 3

    # weekend swatch
    swatch = ws.cell(row=legend_row, column=col)
    swatch.fill   = fill(COL_LIGHT_GRAY)
    swatch.border = border("AAAAAA")
    txt = ws.cell(row=legend_row, column=col + 1)
    txt.value = "Weekend"
    txt.font  = font(size=8, bold=False, color="595959")
    txt.alignment = align("left", "center")

    ws.sheet_view.showGridLines = False


# ─────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────
def main():
    js_path = "holidays_data.js"
    if not Path(js_path).exists():
        print(f"ERROR: {js_path} not found. Run scrape_holidays.py first.")
        return

    print("Loading holidays_data.js ...")
    holidays = load_holidays(js_path)
    total = sum(len(v) for c in holidays.values() for v in c.values())
    print(f"  -> {total} holiday entries across {len(holidays)} countries")

    wb = Workbook()
    wb.remove(wb.active)

    for year in YEARS:
        print(f"Writing {year} ...")
        ws = wb.create_sheet(title=str(year))
        write_year(ws, year, holidays)

    print("Writing Events sheet ...")
    write_events_sheet(wb, holidays)

    # save -- auto-increment if file is open in Excel
    base = "APAC_Holiday_Dates"
    out  = f"{base}.xlsx"
    if Path(out).exists():
        for i in range(2, 99):
            candidate = f"{base}_{i}.xlsx"
            if not Path(candidate).exists():
                out = candidate
                break

    wb.save(out)
    print(f"\nSaved -> {out}")


if __name__ == "__main__":
    main()
