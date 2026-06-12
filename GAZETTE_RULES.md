# Omni Law Gazette — Production Rules

Fixed ruleset for generating each weekly issue. Derived from the approved design template
(`example01.docx`, Issue No. 1, 5–12 June 2026). Every future issue MUST follow these rules
unless the editor explicitly overrides them in the session.

---

## 1. Cadence, window, numbering

- **Frequency:** weekly, published on **Friday**.
- **Coverage window:** previous Friday through publication Friday, inclusive
  (e.g. `5–12 JUNE 2026`). En dash between days; month and year written once when shared.
- **Issue number:** sequential integer, `ISSUE No. N`, incremented by 1 each week.
  Check the latest published issue on the site (or `issues` table) before assigning.
- **Language:** English, **British spelling** (subsidises, centre, organisation).

## 2. Sources (scan ALL, every week)

| # | Source | URL | Role |
|---|--------|-----|------|
| 1 | e-qanun.az (State Register) | https://e-qanun.az/ | **Canonical record** — every instrument cited by its e-qanun framework link |
| 2 | President of Azerbaijan | https://president.az/en/documents | Decrees, orders, signed laws (English versions) |
| 3 | Cabinet of Ministers | https://nk.gov.az/az/senedler/hamisi | Resolutions and orders |
| 4 | Constitutional Court | https://www.constcourt.gov.az/az/decisions | Decisions/plenum rulings |
| 5a | Milli Majlis — Laws | https://meclis.gov.az/cat-qanun.php?cat=72&lang=az | Adopted laws |
| 5b | Milli Majlis — Decrees | https://meclis.gov.az/cat-qanun.php?cat=138&lang=az | Parliamentary decrees |
| 5c | Milli Majlis — Reports | https://meclis.gov.az/cat-cari.php?cat=89&lang=az | Session reports |
| 5d | Milli Majlis — Bills | https://meclis.gov.az/cat-layihe.php?cat=83&lang=az | Bills in progress → feeds the Bills Watch section |

**Dedupe rule:** the same instrument often appears on several sites. e-qanun.az is the
master record; other sites are used for cross-checking, English titles, and context.
Count each instrument once.

## 3. Issue structure (fixed order)

1. **Masthead** (single-column section)
2. **Featured stories** — numbered `01`, `02`, `03`… (three-column newspaper section)
3. **The Full Register** — complete list of ALL legislative updates of the week *(mandatory, never cut)*
4. **Milli Majlis Bills Watch** — bills under consideration *(mandatory, never cut)*
5. Closing line / footer (single-column section)

### 3.1 Masthead

- Navy banner (`#1F3864`) table with the wordmark: `omni` (72 pt) over `law gazette`
  (19 pt), Eloquia Text ExtLt, white on navy, thin navy cell dividers.
- Below it a strapline row closed by a navy hairline:
  - left: `WEEKLY LEGISLATIVE DIGEST · REPUBLIC OF AZERBAIJAN` — bold, navy, 9 pt
  - right: `ISSUE No. N  ·  D–D MONTH YYYY` — 9 pt, dates in caps

### 3.2 Featured stories (the front-page section)

- **6–10 items** per issue, ranked by legal/economic significance — not chronology.
  Item `01` is the week's defining act and gets the fullest treatment.
- Three-column flowing layout (A4, equal columns, 708-twip gutter).
- Each story block, top to bottom (all centred except body):
  | Element | Spec |
  |---|---|
  | Number (`01`) | Eloquia Text ExtLt, bold, 36 pt, colour `#3B3838`, centred |
  | Kicker | `ISSUER · TYPE · D MONTH YYYY` — Times New Roman, bold, 7 pt, caps, colour `#806000`, centred |
  | Headline | Times New Roman, bold, 12 pt, colour `#1F3864`, centred, sentence case, ends with a full stop |
  | Body | Eloquia Text ExtLt, 9 pt, justified |
  | Source line | `Source · e-qanun.az/framework/XXXXX` — Times New Roman, italic, 8–10 pt, colour `#8A6D1E` |
- **Body arc — exactly three paragraphs for the lead, two to three for the rest:**
  1. *What happened* — the instrument, who signed it, when, what it changes.
  2. *Practical effect* — who is affected and how, in plain terms.
  3. *Context* — where it fits in wider policy; cross-reference related items by number, e.g. "(item 02)".
- Mention the **total count of instruments** entered into the State Register during the
  review period somewhere in the lead story ("…the fourteen instruments entered into the
  official State Register during the review period…").
- Kicker issuer values: `PRESIDENT`, `MILLI MAJLIS`, `CABINET OF MINISTERS`,
  `CONSTITUTIONAL COURT`. Type values: `DECREE`, `ORDER`, `LAW`, `RESOLUTION`, `DECISION`.

### 3.3 The Full Register (mandatory)

Complete enumeration of **every** instrument registered/published in the window — including
those already featured. Compact list, one entry per instrument:

```
D MONTH · ISSUER · TYPE — Official title (shortened if needed) · e-qanun.az/framework/XXXXX
```

- Ordered chronologically, oldest first.
- Featured items carry their story number in brackets: `[see item 03]`.
- Nothing is omitted: minor amendments, ratifications, appointments and technical
  instruments all appear here even when not newsworthy.

### 3.4 Milli Majlis Bills Watch (mandatory)

Bills currently in the pipeline (source 5d, cross-checked with 5c session reports):

```
Bill title — stage (first/second/third reading or committee), date of last movement
```

- Include every bill that moved during the window; carry over pending bills with no
  movement only if they are significant.
- One short paragraph (≤2 sentences) of commentary is allowed for noteworthy bills.

## 4. Editorial style

- Newspaper register: declarative, factual, no hedging, no bullet-speak in story bodies.
- No invented facts: every claim traceable to a source document. If an instrument's full
  text is unavailable, say what is verifiable from the title/registry entry only.
- Dates in running text: `9 June` (no ordinal suffix, no year when within the window).
- Source URLs stripped of `https://www.` — e.g. `e-qanun.az/framework/62002`.
- Middle dot `·` is the universal separator (kickers, source lines, masthead).

## 5. Design tokens (do not deviate)

| Token | Value |
|---|---|
| Page | A4 portrait, 1.27 cm (0.5″) margins all round |
| Primary font | Eloquia Text ExtLt |
| Accent font | Times New Roman |
| Navy (masthead, headlines, rules) | `#1F3864` |
| Kicker gold | `#806000` |
| Source gold | `#8A6D1E` |
| Number grey | `#3B3838` |
| Body text | 9 pt justified |
| Columns | 3, equal width, 708-twip spacing |

## 6. Workflow checklist (every week)

1. Determine window (last Friday → today) and next issue number.
2. Fetch all eight sources; collect every instrument dated in the window.
3. Dedupe against e-qanun.az; resolve each item to its framework URL.
4. Rank; pick 6–10 featured stories; write per §3.2.
5. Build the Full Register (§3.3) and Bills Watch (§3.4) — these are never dropped,
   regardless of issue length.
6. Assemble the document to template spec; export PDF named
   `omni-law-gazette-issue-NN.pdf`.
7. Hand the PDF to the editor for publication via `/admin`
   (title: `Issue No. N — D–D Month YYYY`).
