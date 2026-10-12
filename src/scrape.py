import requests
import json
import re
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote

import pandas as pd
from bs4 import BeautifulSoup

BASE_URL = "https://classes.oregonstate.edu/api/"
TERM = "202701"
CACHE_DIR = Path("cache")
DELAY = 0.5  # seconds between detail requests
USE_CACHE = False  # keep False for census runs so seat counts are fresh

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
}

# Shown first in the spreadsheet if present
FIRST_COLUMNS = [
    "code", "title", "crn", "no", "cap", "enrolled", "avail", "waitlist", "fetched_at",
]


def encode(payload: dict) -> str:
    return quote(json.dumps(payload, separators=(",", ":")))


def html_to_text(html: str) -> str:
    return BeautifulSoup(html or "", "html.parser").get_text(" ", strip=True)


def to_int(value):
    digits = re.sub(r"\D", "", html_to_text(str(value)) if value is not None else "")
    return int(digits) if digits else None


def parse_seats(d: dict) -> dict:
    return {
        "cap": to_int(d.get("max_enroll")),
        "enrolled": to_int(d.get("enrollment")),
        "avail": to_int(d.get("ssbsect_seats_avail")),
        "waitlist_cap": to_int(d.get("waitlist_capacity")),
        "waitlist": to_int(d.get("ssbsect_wait_count")),
        "waitlist_avail": to_int(d.get("ssbsect_wait_avail")),
        "xlist_cap": to_int(d.get("cross_list_max_enroll")),
        "xlist_avail": to_int(d.get("ssbxlst_seats_avail")),
    }


def search(session: requests.Session) -> list[dict]:
    payload = {
        "other": {"srcdb": TERM},
        "criteria": [
            {"field": "coursetype_01", "value": "Y"},
            {"field": "camp", "value": "C"},
        ],
    }
    res = session.post(
        BASE_URL,
        params={"page": "fose", "route": "search"},
        data=encode(payload),
        headers=HEADERS,
    )
    res.raise_for_status()
    data = res.json()
    if "results" not in data:
        raise RuntimeError(f"Unexpected response: {str(data)[:500]}")
    return data["results"]


def fetch_details(session: requests.Session, row: dict) -> dict:
    cache_file = CACHE_DIR / f"{row['crn']}.json"
    if USE_CACHE and cache_file.exists():
        return json.loads(cache_file.read_text())

    payload = {
        "group": f"code:{row['code']}",
        "key": f"crn:{row['crn']}",
        "srcdb": TERM,
        "matched": f"crn:{row['crn']}",
    }
    res = session.post(
        BASE_URL,
        params={"page": "fose", "route": "details"},
        data=encode(payload),
        headers=HEADERS,
    )
    res.raise_for_status()
    details = res.json()

    cache_file.write_text(json.dumps(details))
    time.sleep(DELAY)
    return details


def write_xlsx(path: str, rows: list[dict]) -> None:
    df = pd.DataFrame(rows)
    first = [c for c in FIRST_COLUMNS if c in df.columns]
    rest = [c for c in df.columns if c not in first]
    df = df[first + rest]

    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Sections")
        ws = xw.sheets["Sections"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            longest = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
            ws.column_dimensions[col[0].column_letter].width = min(max(longest + 2, 8), 50)


def main() -> None:
    CACHE_DIR.mkdir(exist_ok=True)
    session = requests.Session()
    session.get(
        "https://classes.oregonstate.edu/",
        headers={"User-Agent": HEADERS["User-Agent"]},
    )

    rows = search(session)
    print(f"Found {len(rows)} sections")

    merged = []
    for i, row in enumerate(rows, 1):
        try:
            d = fetch_details(session, row)
        except requests.RequestException as e:
            print(f"[{i}/{len(rows)}] CRN {row['crn']} failed: {e}")
            merged.append(row)
            continue

        merged.append({
            **row,
            "description": html_to_text(d.get("description")),
            "credits": html_to_text(d.get("hours_text") or d.get("hours")),
            "notes": html_to_text(d.get("clssnotes")),
            **parse_seats(d),
            "fetched_at": datetime.now(timezone.utc).isoformat(),
        })
        if i % 25 == 0:
            print(f"[{i}/{len(rows)}] done")

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out = f"osu_census_{TERM}_{stamp}.xlsx"
    write_xlsx(out, merged)
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
