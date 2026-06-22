#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
e-Qanun.az register fetcher.

e-qanun.az is a Next.js SPA backed by a JSON API at https://api.e-qanun.az/.
The home page renders nothing without JavaScript, which is why naive HTML
scraping only sees a tracking pixel. This client talks to the API directly.

How it works (reverse-engineered from the site bundles):
  GET /token        -> opaque token (not actually required as a header)
  GET /sections     -> the 4 top-level act sections
  GET /all-type     -> the full type tree with IDs + live counts
  GET /getDetailSearch?... -> the document search (DataTables-style)

The crucial detail: getDetailSearch returns HTTP 500 for an *unfiltered*
query. It must be scoped to a type via array=<typeId>. Results are sorted by
registration recency (orderColumn=1, orderDirection=desc), and each row gives
the framework id, citation (doc number), title, acceptDate and status.

Registers of interest (typeId -> name):
  31 Presidential Decrees   84 Presidential Orders
  32 Cabinet Decisions      85 Cabinet Orders
  83 Milli Majlis Decrees   73 Constitutional Court decisions
  30 Laws                   28 Constitutional Laws        107 Codes

Usage:
  python3 scripts/fetch-eqanun.py --from 2026-06-15 --to 2026-06-22 \
      --out scripts/gazettes/eqanun-2026-06-22.json
"""

import argparse, json, ssl, sys, urllib.parse, urllib.request
from datetime import date

API = "https://api.e-qanun.az/"
HEADERS = {"Accept-Language": "AZ", "Referer": "https://e-qanun.az/"}
_CTX = ssl.create_default_context()
_CTX.check_hostname = False
_CTX.verify_mode = ssl.CERT_NONE

REGISTERS = {
    31: "Presidential Decrees", 84: "Presidential Orders",
    32: "Cabinet Decisions",    85: "Cabinet Orders",
    83: "Milli Majlis Decrees",  73: "Constitutional Court",
    30: "Laws",                  28: "Constitutional Laws", 107: "Codes",
}

def _get(path):
    req = urllib.request.Request(API + path, headers=HEADERS)
    with urllib.request.urlopen(req, context=_CTX, timeout=60) as r:
        return json.load(r)

def search_type(type_id, length=250):
    """Return rows for one register, newest registration first."""
    params = dict(start=0, length=length, orderColumn=1, orderDirection="desc",
                  array=type_id, title="true", codeType=1, dateType=1,
                  statusId=1, secondType=2, specialDate="false")
    return _get("getDetailSearch?" + urllib.parse.urlencode(params)).get("data", [])

def fetch_window(d_from, d_to, length=250):
    """Documents whose accept date (classCode, YYYY-MM-DD) is in [from, to]."""
    out = {}
    for tid, name in REGISTERS.items():
        rows = search_type(tid, length)
        cc = [r["classCode"] for r in rows if r.get("classCode")]
        covered = bool(cc) and min(cc) < d_from           # fetched past window start
        inwin = [r for r in rows if r.get("classCode") and d_from <= r["classCode"] <= d_to]
        out[name] = {
            "typeId": tid,
            "coverageComplete": covered,
            "newestRegistered": rows[0]["acceptDate"] if rows else None,
            "items": [{
                "id": r["id"],
                "citation": r.get("citation") or None,
                "title": r["title"],
                "acceptDate": r.get("acceptDate"),
                "classCode": r.get("classCode"),
                "status": r.get("statusName"),
                "url": f"https://e-qanun.az/framework/{r['id']}",
            } for r in sorted(inwin, key=lambda r: (r["classCode"], str(r.get("citation"))))],
        }
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="dfrom", default="2026-06-15")
    ap.add_argument("--to",   dest="dto",   default=date.today().isoformat())
    ap.add_argument("--length", type=int, default=250)
    ap.add_argument("--out", default=None)
    a = ap.parse_args()
    data = fetch_window(a.dfrom, a.dto, a.length)
    total = sum(len(v["items"]) for v in data.values())
    for name, v in data.items():
        flag = "" if v["coverageComplete"] else "  ⚠ coverage incomplete (raise --length)"
        print(f"{name:<22} in-window={len(v['items'])}{flag}", file=sys.stderr)
        for it in v["items"]:
            print(f"    {it['acceptDate']}  №{it['citation'] or '—'}  {it['url']}  {it['title'][:64]}",
                  file=sys.stderr)
    payload = {"from": a.dfrom, "to": a.dto, "source": "api.e-qanun.az", "registers": data}
    if a.out:
        json.dump(payload, open(a.out, "w"), ensure_ascii=False, indent=1)
        print(f"\n✅ {total} documents → {a.out}", file=sys.stderr)
    else:
        print(json.dumps(payload, ensure_ascii=False, indent=1))

if __name__ == "__main__":
    main()
