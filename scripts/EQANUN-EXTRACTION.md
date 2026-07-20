# How to extract e-qanun.az (the "it's not extractable" problem — solved)

> **TL;DR** e-qanun.az looks unscrapeable (empty Next.js SPA shell), but it is
> backed by a plain JSON REST API. Do **not** conclude "data not extractable"
> again. Use the recipe below. Verified working 2026-07-13.

## Why it *looks* impossible

`https://e-qanun.az/` is a client-rendered Next.js app. `curl` of any page
returns an empty `<div id="__next">` — no document data in the HTML. That is
what produced the old "renders content dynamically via JavaScript, cannot
extract" note in early gazette issues. It was wrong: the data is one API call
away.

## The hidden API

There are **two** backend hosts, one per legal database:

| Host | Database | Contains |
|------|----------|----------|
| `https://api.e-qanun.az/`  | **Republic of Azerbaijan** | Constitution, Laws, **Presidential Fərman/Sərəncam**, **Cabinet of Ministers Qərar/Sərəncam**, central-executive normative acts |
| `https://napi.e-qanun.az/` | **Nakhchivan Autonomous Republic** | Nakhchivan-level acts only |

⚠ For the Omni Law Gazette you almost always want **`api.e-qanun.az`** (the
Republic). `napi` returns only Nakhchivan documents and will make you think the
week was nearly empty. (Both hosts share the same route/parameter shape; only
the dataset differs.)

### Endpoint

```
GET https://api.e-qanun.az/getDetailSearch
```

Required headers (omit and you may get empty/500):

```
Accept-Language: AZ
Referer: https://e-qanun.az/
```

### Required query parameters

Every one of these must be present or the server returns **HTTP 500**:

```
start=0
length=500                # page size
orderColumn=1             # ⚠ MUST be 1. orderColumn=0 → HTTP 500.
orderDirection=desc
title=true
codeType=1
dateType=1
statusId=1
secondType=2
array=                    # empty = all registries; or CSV of type ids (see all-type)
specialDate=false
```

### Date-range filter (this is how you scope to one week)

```
beginDate=06.07.2026      # ⚠ format MUST be DD.MM.YYYY. ISO YYYY-MM-DD → HTTP 500.
endDate=11.07.2026
dateRange=true
```

`orderColumn=1` sorts by internal insertion id, **not** by date — and recent
bulk imports (e.g. a Nakhchivan batch on `napi`) can dominate the top of an
unfiltered sort. So for a weekly gazette **always use the date-range filter**;
do not just sort desc and slice.

### Working call (copy-paste)

```bash
curl -s "https://api.e-qanun.az/getDetailSearch?start=0&length=500&orderColumn=1&orderDirection=desc&title=true&codeType=1&dateType=1&statusId=1&secondType=2&array=&specialDate=false&beginDate=06.07.2026&endDate=11.07.2026&dateRange=true" \
  -H "Accept-Language: AZ" -H "Referer: https://e-qanun.az/"
```

### Response shape

```jsonc
{ "data": [
  {
    "id": 62149,                       // → full text at https://e-qanun.az/framework/62149
    "citation": "713",                 // official document number ("" if none assigned yet)
    "title": "…full Azerbaijani title…",
    "typeName": "AZƏRBAYCAN RESPUBLİKASI PREZİDENTİNİN FƏRMANLARI",
    "statusName": "Qüvvədədir",        // in force
    "acceptDate": "06.07.2026",        // DD.MM.YYYY
    "classCode": "2026-07-06",         // YYYY-MM-DD (handy for sorting)
    "registerCode": "", "registerDate": null, "effectDate": null
  }
] }
```

### Full document text / downloads

- Human page: `https://e-qanun.az/framework/<id>`
- API downloads (same host as the search): `downloadPdf/<id>`, `downloadWord/<id>`,
  `downloadDetailPdf/<id>`, `downloadImage/<id>`.

### Other useful endpoints (same hosts)

- `all-type` — registry list with live counts. On `api.e-qanun.az` this returns
  the Republic type ids, e.g.: `29` = Laws, `31` = Presidential Decrees
  (Fərmanlar), `32` = Cabinet of Ministers Decisions (Qərarlar), `33` =
  central-executive-authority acts. Put a CSV of these ids in `array=` to filter
  by document type.
- `detailDateCount?beginDate=…&endDate=…&dateType=1&dateRange=…` — count only.
- `detailNameCount?name=…`, `detailCodeCount?code=…`, `typeCount?id=…`,
  `getVersions?id=…` — facet counts / version history.

## How this API was discovered (repeat for other Next.js gov sites)

1. `curl https://e-qanun.az/sitemap.xml` and the page HTML → note it's Next.js
   (`/_next/static/...` chunks, `__NEXT_DATA__`).
2. Download the JS chunks under `/_next/static/chunks/` (incl. `pages/*`).
3. Grep them for the API base (`api.e-qanun.az`, `napi.e-qanun.az`) and for
   `.get("...".concat(base,"endpoint?params"))` call sites. The minified
   `getDetailSearch?start=...` template reveals every parameter.
4. Replay with `curl`; bisect params until 500s clear (that's how
   `orderColumn=1` and the `DD.MM.YYYY` format requirements were found).

## Environment gotcha (why we didn't just use a headless browser)

The pre-installed headless **Chromium cannot reach the internet through the
agent proxy** in this environment — its CONNECT tunnels reset even for
`example.com` (the proxy only relays specific tooling). `curl`/Node `fetch`
work fine because they honor `HTTPS_PROXY` + the CA bundle. So: **hit the JSON
API with curl/fetch; don't rely on Playwright here.**
