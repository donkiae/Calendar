r"""
scrape_econ_events.py
Builds econ_data.js for use in holiday-calendar.html.

Sources (live):
  FOMC  https://www.federalreserve.gov/monetarypolicy/fomccalendars.htm
  GDP   https://www.bea.gov/news/schedule

CPI / NFP: BLS blocks automated requests from all URL patterns.
  2025 dates are the exact published BLS schedule.
  2026-2027 dates are computed estimates following BLS release patterns
  (NFP = first Friday; CPI = Wednesday ~11-14 days after month end).

Run:  C:/anaconda3/python.exe scrape_econ_events.py
Deps: requests beautifulsoup4  (both bundled with Anaconda)
"""

import re
import time
from datetime import datetime, timezone
from pathlib import Path

import requests
from bs4 import BeautifulSoup

# ── config ────────────────────────────────────────────────────
YEARS = [2025, 2026, 2027]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,*/*;q=0.9",
    "Accept-Language": "en-US,en;q=0.9",
}

MONTH_MAP = {
    "january":1,  "february":2,  "march":3,    "april":4,
    "may":5,      "june":6,      "july":7,      "august":8,
    "september":9,"october":10,  "november":11, "december":12,
    "jan":1,"feb":2,"mar":3,"apr":4,"jun":6,"jul":7,
    "aug":8,"sep":9,"oct":10,"nov":11,"dec":12,
}

MONTH_RE = (
    r"(?:January|February|March|April|May|June|July|August|"
    r"September|October|November|December)"
)


# ── hardcoded fallbacks ───────────────────────────────────────
# BLS blocks all automated requests. 2025 = exact published schedule.
# 2026-2027 = computed estimates (NFP first Friday; CPI Wednesday ~day 10-15).
CPI_FALLBACK = {
    2025: [
        "2025-01-15","2025-02-12","2025-03-12","2025-04-10","2025-05-13",
        "2025-06-11","2025-07-11","2025-08-12","2025-09-12","2025-10-15",
        "2025-11-13","2025-12-10",
    ],
    2026: [
        "2026-01-14","2026-02-11","2026-03-11","2026-04-08","2026-05-13",
        "2026-06-10","2026-07-15","2026-08-12","2026-09-09","2026-10-14",
        "2026-11-11","2026-12-09",
    ],
    2027: [
        "2027-01-13","2027-02-10","2027-03-10","2027-04-14","2027-05-12",
        "2027-06-09","2027-07-14","2027-08-11","2027-09-08","2027-10-13",
        "2027-11-10","2027-12-08",
    ],
}

NFP_FALLBACK = {
    2025: [
        "2025-01-10","2025-02-07","2025-03-07","2025-04-04","2025-05-02",
        "2025-06-06","2025-07-03","2025-08-01","2025-09-05","2025-10-03",
        "2025-11-07","2025-12-05",
    ],
    2026: [
        "2026-01-09","2026-02-06","2026-03-06","2026-04-03","2026-05-01",
        "2026-06-05","2026-07-03","2026-08-07","2026-09-04","2026-10-02",
        "2026-11-06","2026-12-04",
    ],
    2027: [
        "2027-01-08","2027-02-05","2027-03-05","2027-04-02","2027-05-07",
        "2027-06-04","2027-07-02","2027-08-06","2027-09-03","2027-10-01",
        "2027-11-05","2027-12-03",
    ],
}

# FOMC 2025 fallback (Fed page may only show current/future years)
FOMC_FALLBACK = {
    2025: [
        "2025-01-29","2025-03-19","2025-05-07","2025-06-18",
        "2025-07-30","2025-09-17","2025-10-29","2025-12-10",
    ],
}

# GDP Advance estimate fallback. 2026 Q2/Q3 confirmed from BEA; rest estimated.
GDP_FALLBACK = {
    2025: ["2025-01-30","2025-04-30","2025-07-30","2025-10-29"],
    2026: ["2026-01-29","2026-04-29","2026-07-30","2026-10-29"],
    2027: ["2027-01-28","2027-04-28","2027-07-28","2027-10-27"],
}


# ── helpers ───────────────────────────────────────────────────
def fetch(url: str) -> str:
    try:
        r = requests.get(url, headers=HEADERS, timeout=25)
        r.raise_for_status()
        return r.text
    except Exception as e:
        print(f"    [WARN] Could not fetch {url}: {e}")
        return ""


def to_iso(year: int, month_str: str, day: int) -> str | None:
    try:
        return datetime(year, MONTH_MAP[month_str.lower()], day).strftime("%Y-%m-%d")
    except (ValueError, KeyError):
        return None


# ── FOMC ─────────────────────────────────────────────────────
def scrape_fomc() -> dict:
    """
    Page structure:
      '2026 FOMC Meetings'   <- year header line
      'January'              <- month (standalone line)
      '27-28'                <- day range; decision day = last day
      'March'
      '17-18*'
      ...
    """
    url  = "https://www.federalreserve.gov/monetarypolicy/fomccalendars.htm"
    html = fetch(url)
    if not html:
        print("    [INFO] FOMC scrape failed; using fallback for known years")
        return dict(FOMC_FALLBACK)

    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "nav", "header", "footer"]):
        tag.decompose()
    text = soup.get_text("\n")

    results: dict = {}
    current_year  = None
    current_month = None   # kept as string for to_iso()

    year_pat  = re.compile(r"^(20\d{2})\s+FOMC", re.IGNORECASE)
    month_pat = re.compile(
        r"^(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)$",
        re.IGNORECASE,
    )
    range_pat  = re.compile(r"^(\d{1,2})\s*[-–]\s*(\d{1,2})\*?$")
    single_pat = re.compile(r"^(\d{1,2})\*?$")

    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        m = year_pat.match(line)
        if m:
            y = int(m.group(1))
            current_year  = y if y in YEARS else None
            current_month = None
            if current_year and current_year not in results:
                results[current_year] = []
            continue

        if current_year is None:
            continue

        m = month_pat.match(line)
        if m:
            current_month = m.group(1)
            continue

        if current_month is None:
            continue

        m = range_pat.match(line)
        if m:
            day = int(m.group(2))          # last day of range = decision day
            iso = to_iso(current_year, current_month, day)
            if iso and iso not in results[current_year]:
                results[current_year].append(iso)
            current_month = None
            continue

        m = single_pat.match(line)
        if m:
            day = int(m.group(1))
            iso = to_iso(current_year, current_month, day)
            if iso and iso not in results[current_year]:
                results[current_year].append(iso)
            current_month = None
            continue

    for y in list(results.keys()):
        results[y] = sorted(set(results[y]))
        if len(results[y]) not in range(6, 11):
            print(f"    [WARN] FOMC {y}: found {len(results[y])} dates (expected ~8)")

    # Merge fallback for years not found on the page (e.g. past years)
    for y, dates in FOMC_FALLBACK.items():
        if y not in results or not results[y]:
            results[y] = dates
            print(f"    [INFO] FOMC {y}: using fallback ({len(dates)} dates)")

    return results


# ── CPI / NFP (hardcoded fallback — BLS blocks all automated requests) ────
def get_cpi() -> dict:
    print("    [INFO] BLS blocks automated requests; using precomputed CPI schedule")
    print("           2025 = exact BLS published dates; 2026-2027 = computed estimates")
    return {y: list(v) for y, v in CPI_FALLBACK.items() if y in YEARS}


def get_nfp() -> dict:
    print("    [INFO] BLS blocks automated requests; using precomputed NFP schedule")
    print("           2025 = exact BLS published dates; 2026-2027 = computed estimates")
    return {y: list(v) for y, v in NFP_FALLBACK.items() if y in YEARS}


# ── BEA GDP (Advance estimate only) ──────────────────────────
def scrape_gdp() -> dict:
    """
    BEA release schedule page. Structure when parsed to plain text:
      'Year 2026'
      [date]           <- 'June 25'
      [time]           <- '8:30 AM'
      'N'              <- split from 'News'
      'ews'
      [title]          <- 'GDP (Advance Estimate), 2nd Quarter 2026'

    We track the most recent date-looking line and record it when we
    spot an Advance GDP title.
    """
    url  = "https://www.bea.gov/news/schedule"
    html = fetch(url)
    results: dict = dict(GDP_FALLBACK)   # start with fallback, overwrite live data

    if not html:
        print("    [INFO] BEA scrape failed; using fallback GDP dates")
        return results

    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "nav", "header", "footer"]):
        tag.decompose()
    lines = [l.strip() for l in soup.get_text("\n").splitlines() if l.strip()]

    date_pat  = re.compile(rf"^({MONTH_RE})\s+(\d{{1,2}})$", re.IGNORECASE)
    year_hdr  = re.compile(r"^Year (\d{4})$")
    gdp_adv   = re.compile(r"GDP\s*\(Advance\s*Estimate\)", re.IGNORECASE)

    current_year = None
    current_date = None   # most recently seen "Month DD" string
    found: dict  = {}

    for line in lines:
        m = year_hdr.match(line)
        if m:
            current_year = int(m.group(1))
            continue

        m = date_pat.match(line)
        if m:
            current_date = line   # e.g. "July 30"
            continue

        if gdp_adv.search(line) and current_year and current_date:
            # Extract year from title if present, else use current_year
            yr_in_title = re.search(r"\b(20\d{2})\b", line)
            yr = int(yr_in_title.group(1)) if yr_in_title else current_year
            if yr in YEARS:
                m2 = date_pat.match(current_date)
                if m2:
                    iso = to_iso(yr, m2.group(1), int(m2.group(2)))
                    if iso:
                        found.setdefault(yr, [])
                        if iso not in found[yr]:
                            found[yr].append(iso)

    for yr, dates in found.items():
        # Merge live dates into results, replacing fallback entries for same quarter
        results[yr] = sorted(set(results.get(yr, []) + dates))

    return results


# ── JS writer ─────────────────────────────────────────────────
def build_js(data: dict) -> str:
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    lines = [
        f"// Generated by scrape_econ_events.py  {ts}",
        "// CPI/NFP: 2025 exact BLS schedule; 2026-2027 computed estimates.",
        "const ECON_DATA = {",
    ]
    for cat, years in data.items():
        lines.append(f"  {cat}: {{")
        for year, dates in sorted(years.items()):
            ds = ", ".join(f'"{d}"' for d in dates)
            lines.append(f"    {year}: [{ds}],")
        lines.append("  },")
    lines.append("};")
    return "\n".join(lines) + "\n"


# ── main ──────────────────────────────────────────────────────
def main():
    sources = [
        ("FOMC", scrape_fomc),
        ("CPI",  get_cpi),
        ("NFP",  get_nfp),
        ("GDP",  scrape_gdp),
    ]

    data: dict = {}
    for label, fn in sources:
        print(f"Fetching {label}...")
        result = fn()
        data[label] = result
        total = sum(len(v) for v in result.values())
        if total:
            for y, dates in sorted(result.items()):
                print(f"  {y}: {len(dates)} dates  {dates[0]} ... {dates[-1]}")
        else:
            print(f"  [WARN] No dates for {label}")
        time.sleep(0.5)

    js = build_js(data)
    out = Path("econ_data.js")
    out.write_text(js, encoding="utf-8")
    print(f"\nSaved -> {out}")

    print("\nSummary:")
    for cat, years in data.items():
        total = sum(len(v) for v in years.values())
        print(f"  {cat:4s}  {total:3d} dates  years: {sorted(years.keys())}")


if __name__ == "__main__":
    main()
