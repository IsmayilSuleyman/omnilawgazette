#!/usr/bin/env python3
"""Extract recent legislative acts from e-qanun.az via its JSON API.

The public e-qanun.az site is a client-rendered Next.js app; its data comes from
a REST API at https://api.e-qanun.az/ that the front-end calls at runtime:

  GET /token                       -> a session token string (sent as Authorization)
  GET /all-type                    -> the document-type / section tree
  GET /getDetailSearch?...&array=N -> paginated list of acts in section N
  GET /framework/{id}              -> full metadata for one act (+ htmlUrl)

getDetailSearch requires an `array` (section id) filter and a trailing
`specialDate` flag; ordered by registerDate descending it yields the newest acts
first. We pull the newest pages per section and keep those whose acceptDate or
registerDate falls inside the reporting window.

Usage:
  python3 scripts/fetch_eqanun.py 2026-06-13 2026-06-20
"""
import sys
import json
import urllib.request
import urllib.error
from datetime import datetime, date

API = "https://api.e-qanun.az/"

# Section ids of interest (from /all-type), mapped to gazette-friendly labels.
SECTIONS = {
    30:  "Laws (Qanunlar)",
    28:  "Constitutional Laws",
    31:  "Presidential Decrees (Fərmanlar)",
    84:  "Presidential Orders (Sərəncamlar)",
    32:  "Cabinet of Ministers — Decisions (Qərarlar)",
    85:  "Cabinet of Ministers — Orders (Sərəncamlar)",
    83:  "Milli Majlis — Decisions (Qərarlar)",
    73:  "Constitutional Court — Decisions",
}


def get_token() -> str:
    req = urllib.request.Request(API + "token", headers={"Accept-Language": "AZ"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode().strip().strip('"')


def api_get(path: str, token: str):
    req = urllib.request.Request(
        API + path,
        headers={"Accept-Language": "AZ", "Authorization": token},
    )
    with urllib.request.urlopen(req, timeout=40) as r:
        return json.loads(r.read().decode())


def parse_date(s):
    """Parse a dd.mm.yyyy string into a date, or None."""
    if not s:
        return None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            continue
    return None


def fetch_section(section_id: int, token: str, pages: int = 2, page_len: int = 50):
    """Return the newest acts in a section (list of dicts)."""
    out = []
    for p in range(pages):
        start = p * page_len
        path = (
            f"getDetailSearch?start={start}&length={page_len}"
            f"&orderColumn=1&orderDirection=desc&specialDate=false&array={section_id}"
        )
        try:
            data = api_get(path, token).get("data", []) or []
        except urllib.error.HTTPError as e:
            print(f"  ! section {section_id} page {p}: HTTP {e.code}", file=sys.stderr)
            break
        if not data:
            break
        out.extend(data)
        if len(data) < page_len:
            break
    return out


def main():
    if len(sys.argv) >= 3:
        start = datetime.strptime(sys.argv[1], "%Y-%m-%d").date()
        end = datetime.strptime(sys.argv[2], "%Y-%m-%d").date()
    else:
        start, end = date(2026, 6, 13), date(2026, 6, 20)

    token = get_token()
    print(f"# e-Qanun extraction — window {start} .. {end}", file=sys.stderr)

    results = {}
    for sid, label in SECTIONS.items():
        rows = fetch_section(sid, token)
        hits = []
        for r in rows:
            ad = parse_date(r.get("acceptDate"))
            rd = parse_date(r.get("registerDate"))
            # an act is "this week" if it was accepted OR officially
            # registered/published within the window
            in_window = (ad and start <= ad <= end) or (rd and start <= rd <= end)
            if in_window:
                hits.append({
                    "id": r.get("id"),
                    "citation": r.get("citation") or "",
                    "title": r.get("title") or "",
                    "typeName": r.get("typeName") or "",
                    "acceptDate": r.get("acceptDate") or "",
                    "registerDate": r.get("registerDate") or "",
                    "status": r.get("statusName") or "",
                    "htmlUrl": f"https://e-qanun.az/framework/{r.get('id')}",
                })
        # Newest act in the section, for the "most recent overall" note.
        # List rows omit registerDate, so fall back to acceptDate.
        newest = None
        for r in rows:
            rd = parse_date(r.get("registerDate")) or parse_date(r.get("acceptDate"))
            if rd and (newest is None or rd > newest[0]):
                newest = (rd, r.get("citation") or "", r.get("title") or "")
        results[label] = {"hits": hits, "newest": newest}
        nd = newest[0].isoformat() if newest else "?"
        print(f"  {label}: {len(hits)} in-window  (latest overall: {nd} {newest[1] if newest else ''})",
              file=sys.stderr)

    # Emit machine-readable JSON to stdout
    serial = {}
    for label, v in results.items():
        serial[label] = {
            "hits": v["hits"],
            "newest": {
                "date": v["newest"][0].isoformat() if v["newest"] else None,
                "citation": v["newest"][1] if v["newest"] else None,
                "title": v["newest"][2] if v["newest"] else None,
            },
        }
    print(json.dumps(serial, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
