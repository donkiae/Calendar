"""
Scrapes public holiday dates from publicholidays.asia and writes holidays_data.js.
Run: python scrape_holidays.py
Deps: pip install requests beautifulsoup4
"""

import json
import time
import re
from datetime import datetime
import requests
from bs4 import BeautifulSoup

COUNTRIES = {
    "SG": "singapore",
    "MY": "malaysia",
    "ID": "indonesia",
    "IN": "india",
    "JP": "japan",
    "VN": "vietnam",
    "TH": "thailand",
    "PH": "philippines",
}

YEARS = [2025, 2026, 2027]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}


def scrape_country_year(country_slug: str, year: int) -> list[dict]:
    url = f"https://publicholidays.asia/{country_slug}/{year}-dates/"
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        r.raise_for_status()
    except requests.RequestException as e:
        print(f"  [ERROR] {url}: {e}")
        return []

    soup = BeautifulSoup(r.text, "html.parser")

    # The site renders a table with class "publicholidays"
    table = soup.find("table", class_=re.compile(r"publicholidays", re.I))
    if not table:
        # Fallback: any table inside the main content area
        table = soup.select_one(".content-area table, #main table, article table")
    if not table:
        print(f"  [WARN] No holiday table found for {country_slug} {year}")
        return []

    holidays = []
    for row in table.find_all("tr"):
        cells = [td.get_text(strip=True) for td in row.find_all("td")]
        if len(cells) < 3:
            continue

        date_text = cells[0]
        holiday_name = cells[2] if len(cells) > 2 else cells[1]

        # Parse date like "1 January 2025" or "January 1"
        parsed = None
        for fmt in (
            "%d %B %Y",   # "1 January 2025"
            "%B %d %Y",   # "January 1 2025"
            "%d %b %Y",   # "1 Jan 2025"
        ):
            try:
                # Some rows omit the year; append it if missing
                test = date_text if str(year) in date_text else f"{date_text} {year}"
                parsed = datetime.strptime(test.strip(), fmt)
                break
            except ValueError:
                pass

        if not parsed:
            continue

        iso = parsed.strftime("%Y-%m-%d")
        # Skip rows from adjacent years (sometimes listed)
        if parsed.year != year:
            continue

        holidays.append({"date": iso, "name": holiday_name})

    # Deduplicate by date+name
    seen = set()
    unique = []
    for h in holidays:
        key = (h["date"], h["name"])
        if key not in seen:
            seen.add(key)
            unique.append(h)

    return sorted(unique, key=lambda x: x["date"])


def build_js(data: dict) -> str:
    lines = ["const EMBEDDED_HOLIDAYS = {"]
    for code, years in data.items():
        lines.append(f"  {code}: {{")
        for year, holidays in years.items():
            lines.append(f"    {year}: [")
            for h in holidays:
                name = h["name"].replace("\\", "\\\\").replace('"', '\\"')
                lines.append(f'    {{date:"{h["date"]}",name:"{name}"}},')
            lines.append("    ],")
        lines.append("  },")
    lines.append("};")
    return "\n".join(lines) + "\n"


def main():
    data = {}
    for code, slug in COUNTRIES.items():
        data[code] = {}
        for year in YEARS:
            print(f"Scraping {code} ({slug}) {year}...")
            holidays = scrape_country_year(slug, year)
            data[code][year] = holidays
            print(f"  → {len(holidays)} holidays")
            time.sleep(0.8)  # be polite

    js_content = build_js(data)
    out_path = "holidays_data.js"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(js_content)
    print(f"\nWrote {out_path}")

    # Count totals
    total = sum(len(h) for c in data.values() for h in c.values())
    print(f"Total: {total} holiday entries across {len(COUNTRIES)} countries × {len(YEARS)} years")


if __name__ == "__main__":
    main()
