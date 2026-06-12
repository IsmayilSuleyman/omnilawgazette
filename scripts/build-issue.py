#!/usr/bin/env python3
"""
Omni Law Gazette — DOCX issue builder.

Rebuilds word/document.xml (and footer1.xml) inside the firm's gazette
template (templates/OmniLawGazette_template.docx), keeping the masthead
header, fonts and styles intact, and repackages the result as
build/OmniLawGazette_Issue<N>.docx.

Run: python3 scripts/build-issue.py
"""

import os
import re
import shutil
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(ROOT, "templates", "OmniLawGazette_template.docx")
BUILD_DIR = os.path.join(ROOT, "build")
WORK = os.path.join(BUILD_DIR, "_docx")

ISSUE_NUMBER = 2
ISSUE_PERIOD = "5–12 JUNE 2026"

# ───────────────────────────── content ──────────────────────────────
# Issue No. 2 · review period 5–12 June 2026.
# Register links re-scraped from the official State Register API
# (api.e-qanun.az/framework/{id}) on 12 June 2026.

EQ = "https://e-qanun.az/framework/"

GLANCE = {
    "title": "Sixteen instruments enter the State Register",
    "counts": "5 presidential decrees   ·   8 Cabinet decisions   ·   3 presidential orders",
    "takeaways": [
        ("01", "Baku city administration restructured from the ground up"),
        ("02", "State credit guarantees and interest subsidies for business overhauled"),
        ("06", "A rare onshore oil exploration deal — SOCAR & Gran Tierra Energy — approved"),
    ],
}

LEAD = {
    "num": "01",
    "kicker": "PRESIDENT · DECREE · 10 JUNE 2026",
    "headline": "Baku city administration restructured from the ground up",
    "left": [
        ("lead",
         "The week’s defining act is a presidential decree of 10 June reorganising the structure and management of the Baku City Executive Authority.",
         " The decree redraws how the capital — home to roughly a quarter of the country’s population and the centre of its economic life — is administered, revising the authority’s apparatus, its subordinate bodies and the way the city’s districts are run."),
        ("body",
         "Structural decrees of this kind determine which departments exist, who answers to whom and where budgetary responsibility sits. By consolidating overlapping units and renewing management lines, the act aims to make the capital’s administration faster and more accountable in everyday matters — construction, transport, communal services and public amenities."),
    ],
    "right": [
        ("body",
         "The reorganisation is also a signal beyond Baku: the capital routinely serves as the model that other city and district executive authorities follow. Practitioners advising on permits, municipal contracts or property matters in Baku should verify which successor department now holds their counterparty’s functions."),
    ],
    "source": ("62002", EQ + "62002"),
}

STORIES = [
    {
        "num": "02",
        "kicker": "PRESIDENT · DECREE · 9 JUNE",
        "headline": "Credit guarantees and interest subsidies for entrepreneurs expanded",
        "hsize": 24,
        "paras": [
            "A presidential decree of 9 June refines the mechanisms through which the state guarantees bank loans extended to business entities and subsidises part of the interest they pay, amending several earlier presidential decrees in one stroke.",
            "The reform widens access to state-backed financing and lowers the effective cost of borrowing for private business. By easing the collateral burden on small and medium enterprises, it channels commercial credit toward the productive, non-oil economy — manufacturing, agriculture, services and export-oriented production.",
        ],
        "source": ("61987", EQ + "61987"),
    },
    {
        "num": "03",
        "kicker": "PRESIDENT · DECREE · 9 JUNE",
        "headline": "New support for non-oil-gas exports and transport costs",
        "hsize": 24,
        "paras": [
            "A second decree of 9 June introduces additional measures to stimulate the export of non-oil-gas goods and to defray the transport and logistics costs borne by exporters.",
            "The competitiveness of Azerbaijani goods abroad is often eroded by freight expense. By underwriting part of those shipping costs and reinforcing export-promotion mechanisms, the decree pairs with the credit-support reform signed the same day — two halves of a single push to grow the non-oil export base.",
        ],
        "source": ("61986", EQ + "61986"),
    },
    {
        "num": "04",
        "kicker": "PRESIDENT · DECREE · 8 JUNE",
        "headline": "Agrarian Credit and Development Agency overhauled",
        "hsize": 24,
        "paras": [
            "By decree of 8 June the President amended the Statute of the Agrarian Credit and Development Agency under the Ministry of Agriculture, revising several earlier decrees that together govern the agency.",
            "The agency is the main conduit for concessional credit and subsidies to farmers. Consolidating its legal basis modernises the administration of rural finance and tightens the link between the agency’s mandate and the government’s food-security agenda.",
        ],
        "source": ("61985", EQ + "61985"),
    },
    {
        "num": "05",
        "kicker": "PRESIDENT · DECREE · 10 JUNE",
        "headline": "Citizens’ appeals reform put into effect",
        "hsize": 24,
        "paras": [
            "A decree of 10 June provides for the implementation of the 26 May amendment to the Law on Citizens’ Appeals (No. 415-VIIQD), assigning the executive bodies responsible for applying the revised rules.",
            "The Law on Citizens’ Appeals is the workhorse statute through which individuals and companies petition state bodies and receive binding response deadlines. Implementation decrees like this one determine when and how the amended procedure starts to operate in practice.",
        ],
        "source": ("62000", EQ + "62000"),
    },
    {
        "num": "06",
        "kicker": "CABINET OF MINISTERS · DECISION NO. 174 · 9 JUNE",
        "headline": "SOCAR–Gran Tierra Energy onshore exploration agreement approved",
        "hsize": 21,
        "paras": [
            "Decision No. 174 approves the agreement between the State Oil Company (SOCAR) and Gran Tierra Energy (Azerbaijan) GmbH on oil exploration and production in onshore areas covering the Quba, Qusar and Shabran districts.",
            "Onshore exploration deals with foreign operators have been rare in recent years; this one signals renewed interest in the country’s under-explored land acreage. The Cabinet’s approval is the regulatory step that allows the contract regime — taxation, cost recovery, profit sharing — to take effect.",
        ],
        "source": ("62004", EQ + "62004"),
    },
    {
        "num": "07",
        "kicker": "CABINET OF MINISTERS · DECISION NO. 173 · 9 JUNE",
        "headline": "Gas-supply rules rewritten under the 2025 Gas Supply Law",
        "hsize": 21,
        "paras": [
            "Decision No. 173 brings the Cabinet’s subordinate regulations into line with the Gas Supply Law (No. 233-VIIQ of 8 July 2025), the framework statute that reorganised the legal regime for the gas sector.",
            "Folding the scattered legacy rules into the new statutory structure gives distributors, operators and large consumers a single, current rulebook and removes the risk of conflicting requirements left over from the pre-2025 regime.",
        ],
        "source": ("61994", EQ + "61994"),
    },
    {
        "num": "08",
        "kicker": "CABINET OF MINISTERS · DECISION NO. 171 · 9 JUNE",
        "headline": "Rules set for the use of public railway infrastructure",
        "hsize": 21,
        "paras": [
            "Decision No. 171 approves the Rules on the maintenance and use of public-use railway transport infrastructure, defining the responsibilities of the infrastructure manager and the terms on which carriers may obtain access to the network.",
            "Clear access conditions are a structural prerequisite for any future opening of the railway market — and for predictable operation of the freight corridors on which the country’s transit ambitions depend.",
        ],
        "source": ("61993", EQ + "61993"),
    },
    {
        "num": "09",
        "kicker": "CABINET OF MINISTERS · DECISION NO. 168 · 9 JUNE",
        "headline": "Energy-efficiency information system established",
        "hsize": 21,
        "paras": [
            "Decision No. 168 approves the Statute of the Energy Efficiency Information System — a national digital platform to collect, store and analyse data on how energy is consumed across the economy.",
            "The system provides the data backbone for monitoring efficiency targets, identifying waste and reporting progress under Azerbaijan’s climate and energy commitments.",
        ],
        "source": ("61995", EQ + "61995"),
    },
    {
        "num": "10",
        "kicker": "CABINET OF MINISTERS · DECISION NO. 169 · 9 JUNE",
        "headline": "State Security Service digital-analytics system implemented",
        "hsize": 21,
        "paras": [
            "Decision No. 169 amends several Cabinet decisions to give effect to the creation of the State Security Service’s Centralised Information and Digital Analytics System.",
            "The amendments align inter-agency data and reporting rules with the new system — the secondary legislation needed for it to operate. Entities that exchange data with state security bodies should note the revised channels.",
        ],
        "source": ("61988", EQ + "61988"),
    },
    {
        "num": "11",
        "kicker": "CABINET OF MINISTERS · DECISION NO. 170 · 9 JUNE",
        "headline": "Mandatory health-insurance services package amended",
        "hsize": 21,
        "paras": [
            "Decision No. 170 amends the Services Package provided under mandatory health insurance — the schedule of treatments and services covered by the state scheme.",
            "Adjustments to this package directly determine which medical services patients can obtain free of charge, making it one of the regulations most immediately felt by the public.",
        ],
        "source": ("61990", EQ + "61990"),
    },
    {
        "num": "12",
        "kicker": "CABINET OF MINISTERS · DECISION NO. 175 · 10 JUNE",
        "headline": "Ministry of Economy’s subordinate-bodies list revised",
        "hsize": 21,
        "paras": [
            "Decision No. 175 amends the list of institutions subordinate to the Ministry of Economy but outside its core structure, updating the ministry’s institutional perimeter.",
            "Such list changes typically accompany the creation, merger or transfer of agencies; counterparties of the affected bodies should confirm their current legal status and supervising authority.",
        ],
        "source": ("61999", EQ + "61999"),
    },
]

IN_BRIEF = {
    "kicker": "IN BRIEF · 5–12 JUNE",
    "headline": "Also entered on the register",
    "paras": [
        ("Presidential orders. ",
         "Three non-normative orders were registered: financial support to the Azerbaijan Minifootball Federation (5 June, doc. 61973); a state decoration for water-management and land-reclamation workers (5 June, doc. 61972); and the Honour Diploma of the President conferred on R.M. Fətəliyev (8 June, doc. 61984)."),
        ("Official parking. ",
         "Cabinet Decision No. 172 (9 June, doc. 62003) approved the list of state bodies entitled to designated service-parking areas and the number of spaces allocated to each."),
        ("Parliament. ",
         "No new parliamentary laws or decrees were entered on the register for the period. The Milli Majlis opened fifteen bills for public discussion on 5 and 12 June — among them the 2025 state-budget execution bill, the establishment of an embassy in Lisbon and amendments to the Land Code."),
        ("Not this week. ",
         "No Constitutional Court decisions were entered for the period; the most recent registered ruling dates from 25 May (doc. 61977)."),
        ("Methodology. ",
         "Compiled from the official State Register at e-qanun.az and cross-checked against president.az, nk.gov.az and meclis.gov.az. A complete, linked index of the week’s sixteen instruments follows on the final page."),
    ],
}

MECLIS = "https://meclis.gov.az/news-layihe.php?id="

HORIZON = {
    "kicker": "ON THE HORIZON · MİLLİ MƏCLİS",
    "headline": "Bills opened for public discussion",
    "intro": ("Fifteen bills were opened for public discussion during the review period — "
              "none yet law, all worth tracking. Submitted 12 June unless noted."),
    "left": [
        ("Execution of the 2025 state budget", "2649"),
        ("Land Code and related legislation — amendments", "2644"),
        ("Housing Code & IDP legislation — amendments", "2646"),
        ("Administrative Violations Code — amendments", "2638"),
        ("Administrative Violations Code & information laws", "2643"),
        ("Civil Service Law — amendments", "2642"),
        ("Law on Protection of Green Spaces — amendments", "2637"),
        ("Shusha Cultural Capital Law — amendments", "2648"),
    ],
    "right": [
        ("Embassy of Azerbaijan in Lisbon — establishment", "2645"),
        ("Anti-fouling Systems Convention — accession", "2647"),
        ("Criminal Procedure Code — amendments (5 June)", "2640"),
        ("Penal Enforcement Code & enforcement laws (5 June)", "2639"),
        ("Anti-Corruption Law — amendments (5 June)", "2641"),
        ("Diplomatic Service Law — amendments (5 June)", "2635"),
        ("Customs Code — amendments (5 June)", "2630"),
    ],
}

INDEX_TITLE = "All New Legislation · 5–12 June 2026"
INDEX_INTRO = ("Every act entered into the official State Register (e-qanun.az) "
               "during the review period. Click any link to open the full text.")

INDEX = [
    ("Presidential Decree — improvement of the structure and management of the Baku City Executive Authority", "10 June", "item 01", "62002"),
    ("Presidential Decree — improvement of credit-guarantee and interest-subsidy mechanisms for entrepreneurs", "9 June", "item 02", "61987"),
    ("Presidential Decree — additional measures to stimulate non-oil-gas exports and support transport costs", "9 June", "item 03", "61986"),
    ("Presidential Decree — amendments to the Statute of the Agrarian Credit and Development Agency", "8 June", "item 04", "61985"),
    ("Presidential Decree — implementation of the amendment to the Law on Citizens’ Appeals", "10 June", "item 05", "62000"),
    ("Cabinet Decision No. 174 — SOCAR–Gran Tierra Energy onshore oil exploration and production agreement", "9 June", "item 06", "62004"),
    ("Cabinet Decision No. 173 — gas-supply rules aligned with the 2025 Gas Supply Law", "9 June", "item 07", "61994"),
    ("Cabinet Decision No. 171 — Rules on the maintenance and use of public railway infrastructure", "9 June", "item 08", "61993"),
    ("Cabinet Decision No. 168 — Statute of the Energy Efficiency Information System", "9 June", "item 09", "61995"),
    ("Cabinet Decision No. 169 — implementation of the State Security Service digital-analytics system", "9 June", "item 10", "61988"),
    ("Cabinet Decision No. 170 — amendment to the mandatory health-insurance Services Package", "9 June", "item 11", "61990"),
    ("Cabinet Decision No. 175 — list of bodies subordinate to the Ministry of Economy revised", "10 June", "item 12", "61999"),
    ("Cabinet Decision No. 172 — list of state bodies with designated official parking areas", "9 June", "In Brief", "62003"),
    ("Presidential Order — financial support to the Azerbaijan Minifootball Federation", "5 June", "In Brief", "61973"),
    ("Presidential Order — state decoration for water-management and land-reclamation workers", "5 June", "In Brief", "61972"),
    ("Presidential Order — Honour Diploma of the President conferred on R.M. Fətəliyev", "8 June", "In Brief", "61984"),
]

# ───────────────────────── design tokens ────────────────────────────

NAVY, GOLD, BODY, GRAY = "033B74", "8A6D1E", "3A3A3A", "7A7A7A"
GOLD_BRIGHT = "D8B54A"        # gold accents on dark navy
IVORY = "F5F2E9"              # warm shaded boxes
IVORY_DEEP = "EFEAD9"         # zebra stripe
RULE = "DCD5C2"               # hairline rules between bands
NUM_BLUE = "B7C4D6"           # ghost story numbers
FONT = "Eloquia Text"
FONT_LT = "Eloquia Text ExtLt"

FULL_W = 10456                # content width in dxa

_hyperlinks: list[str] = []   # collected targets; rIds assigned from 100


def esc(t: str) -> str:
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def run(text, *, font=FONT, sz=17, color=BODY, b=False, i=False, u=False, track=None):
    rpr = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>'
    if b: rpr += "<w:b/>"
    if i: rpr += "<w:i/>"
    rpr += f'<w:color w:val="{color}"/>'
    if track: rpr += f'<w:spacing w:val="{track}"/>'
    rpr += f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    if u: rpr += '<w:u w:val="single"/>'
    rpr += '<w:lang w:val="en-US"/>'
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def hyperlink(url, text=None, **rkw):
    _hyperlinks.append(url)
    rid = f"rId{99 + len(_hyperlinks)}"
    return f'<w:hyperlink r:id="{rid}">{run(text or url, u=True, **rkw)}</w:hyperlink>'


def para(runs_xml, *, after=80, before=0, jc=None, tabs=None):
    ppr = ''
    if tabs:
        ppr += '<w:tabs>' + ''.join(f'<w:tab w:val="{v}" w:pos="{p}"/>' for v, p in tabs) + '</w:tabs>'
    ppr += f'<w:spacing w:before="{before}" w:after="{after}"/>'
    if jc: ppr += f'<w:jc w:val="{jc}"/>'
    return f'<w:p><w:pPr>{ppr}<w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr>{runs_xml}</w:p>'


def empty_para(sz=2):
    return (f'<w:p><w:pPr><w:spacing w:after="0" w:line="40" w:lineRule="exact"/>'
            f'<w:rPr><w:sz w:val="{sz}"/><w:lang w:val="en-US"/></w:rPr></w:pPr></w:p>')


def p_kicker(text):
    return para(run(text, sz=15, color=GOLD, b=True, track=30), after=80)


def p_headline(text, sz=24):
    return para(run(text, sz=sz, color=NAVY, b=True), after=40)


def p_lead(bold_part, rest):
    return para(run(bold_part, sz=18, b=True) + run(rest, sz=18), after=80, jc="both")


def p_body(text, sz=17, keep=False):
    keep_xml = '<w:keepNext/>' if keep else ''
    return (f'<w:p><w:pPr>{keep_xml}<w:spacing w:after="80"/><w:jc w:val="both"/>'
            f'<w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr>{run(text, sz=sz)}</w:p>')


_DOC_RE = re.compile(r'doc\. (\d{4,6})')


def p_brief(topic, rest):
    """In-Brief paragraph; every `doc. NNNNN` becomes a live register link."""
    runs = [run(topic, b=True)]
    pos = 0
    for m in _DOC_RE.finditer(rest):
        if m.start() > pos:
            runs.append(run(rest[pos:m.start()]))
        runs.append(hyperlink(EQ + m.group(1), f"doc. {m.group(1)}", sz=17, color=GOLD))
        pos = m.end()
    if pos < len(rest):
        runs.append(run(rest[pos:]))
    return para(''.join(runs), after=80, jc="both")


def p_source(url):
    short = url.replace("https://", "")
    return para(run("Source · ", sz=15, color=GOLD, i=True) +
                hyperlink(url, short, sz=15, color=GOLD, i=True), after=80)


def p_bignum(num):
    return (f'<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="{FONT_LT}" w:hAnsi="{FONT_LT}"/>'
            f'<w:color w:val="{NUM_BLUE}"/>'
            f'<w:sz w:val="56"/><w:szCs w:val="56"/><w:lang w:val="en-US"/></w:rPr>'
            f'<w:t>{num}</w:t></w:r></w:p>')


def borders_xml(borders):
    """borders: dict side -> (sz_eighths_pt, color)"""
    if not borders:
        return ''
    inner = ''.join(
        f'<w:{side} w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        for side, (sz, color) in borders.items())
    return f'<w:tcBorders>{inner}</w:tcBorders>'


def tc(width, content, *, span=None, valign=None, nowrap=False, vmerge=None,
       fill=None, borders=None, pad=None):
    pr = f'<w:tcW w:w="{width}" w:type="dxa"/>'
    if span: pr += f'<w:gridSpan w:val="{span}"/>'
    if vmerge: pr += ('<w:vMerge w:val="restart"/>' if vmerge == "restart" else '<w:vMerge/>')
    pr += borders_xml(borders)
    if fill: pr += f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
    if nowrap: pr += '<w:noWrap/>'
    if pad:
        pr += (f'<w:tcMar><w:top w:w="{pad}" w:type="dxa"/><w:left w:w="{pad}" w:type="dxa"/>'
               f'<w:bottom w:w="{pad}" w:type="dxa"/><w:right w:w="{pad}" w:type="dxa"/></w:tcMar>')
    if valign: pr += f'<w:vAlign w:val="{valign}"/>'
    return f'<w:tc><w:tcPr>{pr}</w:tcPr>{content or empty_para()}</w:tc>'


def tr(cells, *, cant_split=False, header=False, exact_h=None):
    trpr = ''
    if cant_split or header or exact_h:
        trpr = '<w:trPr>'
        if exact_h: trpr += f'<w:trHeight w:val="{exact_h}" w:hRule="exact"/>'
        if cant_split: trpr += '<w:cantSplit/>'
        if header: trpr += '<w:tblHeader/>'
        trpr += '</w:trPr>'
    return f'<w:tr>{trpr}{cells}</w:tr>'


GUTTER = tc(283, None)
COL_RULE = {"left": (4, RULE)}          # vertical rule between the two columns


def header_row(num_l, head_l, num_r, head_r):
    return tr(
        tc(1097, p_bignum(num_l) if num_l else None, valign="center", nowrap=True) +
        tc(3962, head_l) + GUTTER +
        tc(1104, p_bignum(num_r) if num_r else None, valign="center", nowrap=True,
           borders=COL_RULE) +
        tc(4010, head_r),
        cant_split=True
    )


def body_row(left_paras, right_paras):
    return tr(
        tc(5059, left_paras, span=2) + GUTTER +
        tc(5114, right_paras, span=2, borders=COL_RULE)
    )


def band_rule_row():
    """Slim spacer row whose top edge draws a hairline across the page."""
    return tr(
        tc(FULL_W, empty_para(), span=5, borders={"top": (6, RULE)}),
        exact_h=200
    )


def spacer_row(h=120):
    return tr(tc(FULL_W, empty_para(), span=5), exact_h=h)


# ───────────────────────── page-one furniture ───────────────────────

def issue_strip_row():
    """Folio line under the masthead: digest name left, issue + dates right."""
    content = para(
        run("WEEKLY LEGISLATIVE DIGEST · REPUBLIC OF AZERBAIJAN",
            sz=15, color=NAVY, b=True, track=30) +
        f'<w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:tab/></w:r>' +
        run(f"ISSUE No. {ISSUE_NUMBER}", sz=15, color=GOLD, b=True, track=30) +
        run(f"  ·  {ISSUE_PERIOD}", sz=15, color=NAVY, b=True, track=30),
        after=60, tabs=[("right", 10240)])
    return tr(
        tc(FULL_W, content, span=5, borders={"bottom": (12, NAVY)}),
        cant_split=True
    )


def glance_box_row():
    """Navy 'This Week' briefing box."""
    inner = para(run("THIS WEEK", sz=15, color=GOLD_BRIGHT, b=True, track=40), after=60)
    inner += para(run(GLANCE["title"], sz=26, color="FFFFFF", b=True), after=50)
    inner += para(run(GLANCE["counts"], sz=17, color="D9E2EF"), after=90)
    for num, text in GLANCE["takeaways"]:
        inner += para(
            run(num, sz=17, color=GOLD_BRIGHT, b=True) +
            run("   " + text, sz=17, color="FFFFFF"),
            after=40)
    return tr(
        tc(FULL_W, inner, span=5, fill=NAVY, pad=220),
        cant_split=True
    )


# ───────────────────────── document assembly ────────────────────────

def story_header(st):
    return p_kicker(st["kicker"]) + p_headline(st["headline"], st.get("hsize", 21))


def story_body(st):
    paras = st["paras"]
    out = "".join(p_body(p) for p in paras[:-1])
    out += p_body(paras[-1], keep=True)      # last para stays with its source line
    out += p_source(st["source"][1])
    return out


def build_document_xml():
    rows = []

    # Page-one furniture
    rows.append(issue_strip_row())
    rows.append(spacer_row(140))
    rows.append(glance_box_row())
    rows.append(spacer_row(200))

    # Band 1 — lead story (left, big headline) | story 02 (right)
    lead_header = p_kicker(LEAD["kicker"]) + p_headline(LEAD["headline"], 30)
    lead_body = "".join(p_lead(p[1], p[2]) if p[0] == "lead" else p_body(p[1])
                        for p in LEAD["left"] + LEAD["right"])
    lead_body += p_source(LEAD["source"][1])
    R = STORIES[0]
    rows.append(header_row(LEAD["num"], lead_header, R["num"], story_header(R)))
    rows.append(body_row(lead_body, story_body(R)))

    # Bands 2..6 — story pairs 03|04, 05|06, 07|08, 09|10, 11|12
    pairs = [(1, 2), (3, 4), (5, 6), (7, 8), (9, 10)]
    for li, ri in pairs:
        rows.append(band_rule_row())
        L, Rr = STORIES[li], STORIES[ri]
        rows.append(header_row(L["num"], story_header(L), Rr["num"], story_header(Rr)))
        rows.append(body_row(story_body(L), story_body(Rr)))

    # IN BRIEF — shaded full-width box with a gold top rule
    rows.append(spacer_row(140))
    brief = p_kicker(IN_BRIEF["kicker"]) + p_headline(IN_BRIEF["headline"], 21)
    brief += "".join(p_brief(t, r) for t, r in IN_BRIEF["paras"])
    rows.append(tr(tc(FULL_W, brief, span=5, fill=IVORY,
                      borders={"top": (12, GOLD)}, pad=200)))

    # ON THE HORIZON — bills under discussion, two columns of trackers
    def bill_line(title, bill_id):
        return para(
            run("›  ", sz=17, color=GOLD, b=True) +
            run(title + "  ", sz=17, color=NAVY, b=True) +
            hyperlink(MECLIS + bill_id + "&lang=az", f"meclis.gov.az · {bill_id}",
                      sz=14, color=GOLD),
            after=60)

    rows.append(spacer_row(140))
    horizon_head = (p_kicker(HORIZON["kicker"]) +
                    p_headline(HORIZON["headline"], 21) +
                    p_body(HORIZON["intro"]))
    rows.append(tr(tc(FULL_W, horizon_head, span=5,
                      borders={"top": (12, NAVY)}, pad=120), cant_split=True))
    rows.append(tr(
        tc(5059, "".join(bill_line(t, b) for t, b in HORIZON["left"]), span=2, pad=120) +
        GUTTER +
        tc(5114, "".join(bill_line(t, b) for t, b in HORIZON["right"]), span=2,
           borders=COL_RULE, pad=120)
    ))

    tblpr = ('<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>'
             '<w:tblBorders>'
             '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
             '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
             '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
             '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
             '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
             '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
             '</w:tblBorders>'
             '<w:tblLayout w:type="fixed"/>'
             '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>')
    grid = ('<w:tblGrid><w:gridCol w:w="1097"/><w:gridCol w:w="3962"/>'
            '<w:gridCol w:w="283"/><w:gridCol w:w="1104"/><w:gridCol w:w="4010"/></w:tblGrid>')
    table = f'<w:tbl>{tblpr}{grid}{"".join(rows)}</w:tbl>'

    # ── Index page: the week's register as a proper table ──
    index = ['<w:p><w:pPr><w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr>'
             '<w:r><w:br w:type="page"/></w:r></w:p>']
    index.append(para(run(INDEX_TITLE, sz=34, color=NAVY, b=True), after=100))
    index.append(para(run(INDEX_INTRO, sz=18), after=200))

    cols = (560, 6100, 1080, 2716)            # № | instrument | date | source
    idx_rows = [tr(
        tc(cols[0], para(run("№", sz=15, color="FFFFFF", b=True, track=20), after=0),
           fill=NAVY, pad=90) +
        tc(cols[1], para(run("INSTRUMENT", sz=15, color="FFFFFF", b=True, track=20), after=0),
           fill=NAVY, pad=90) +
        tc(cols[2], para(run("DATE", sz=15, color="FFFFFF", b=True, track=20), after=0),
           fill=NAVY, pad=90) +
        tc(cols[3], para(run("SOURCE", sz=15, color="FFFFFF", b=True, track=20), after=0),
           fill=NAVY, pad=90),
        cant_split=True, header=True)]

    for n, (title, date, ref, doc_id) in enumerate(INDEX, 1):
        fill = IVORY_DEEP if n % 2 == 0 else None
        sep = {"bottom": (4, RULE)}
        idx_rows.append(tr(
            tc(cols[0], para(run(str(n), sz=16, color=GRAY, b=True), after=0),
               fill=fill, borders=sep, pad=90, valign="center") +
            tc(cols[1], para(run(title, sz=17, color=NAVY, b=True) +
                             run(f"   ({ref})", sz=15, color=GRAY), after=0),
               fill=fill, borders=sep, pad=90, valign="center") +
            tc(cols[2], para(run(date, sz=16, color=BODY), after=0),
               fill=fill, borders=sep, pad=90, valign="center") +
            tc(cols[3], para(hyperlink(EQ + doc_id, f"framework/{doc_id}",
                                       sz=15, color=GOLD), after=0),
               fill=fill, borders=sep, pad=90, valign="center"),
            cant_split=True))

    idx_grid = '<w:tblGrid>' + ''.join(f'<w:gridCol w:w="{w}"/>' for w in cols) + '</w:tblGrid>'
    index.append(f'<w:tbl>{tblpr}{idx_grid}{"".join(idx_rows)}</w:tbl>')
    index.append(empty_para())

    sectpr = ('<w:sectPr><w:headerReference w:type="default" r:id="rId31"/>'
              '<w:footerReference w:type="default" r:id="rId32"/>'
              '<w:pgSz w:w="11906" w:h="16838"/>'
              '<w:pgMar w:top="2450" w:right="720" w:bottom="850" w:left="720" w:header="708" w:footer="420" w:gutter="0"/>'
              '<w:pgNumType w:start="1"/><w:cols w:space="708"/>'
              '<w:docGrid w:linePitch="360"/></w:sectPr>')

    ns = ('xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
          'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
          'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
          'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
          'xmlns:w10="urn:schemas-microsoft-com:office:word" '
          'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
          'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
          'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
          'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
          'mc:Ignorable="w14 w15 wp14"')

    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<w:document {ns}><w:body>{table}{"".join(index)}{sectpr}</w:body></w:document>')


def build_footer_xml():
    """Branded folio: issue identity left, 'Page X of Y' right."""
    ns = ('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
          'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
          'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
          'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
          'mc:Ignorable="w14"')

    def fld(instr):
        return (f'<w:fldSimple w:instr=" {instr} \\* MERGEFORMAT ">'
                f'{run("1", sz=14, color=GRAY)}</w:fldSimple>')

    p = ('<w:p><w:pPr>'
         '<w:pBdr><w:top w:val="single" w:sz="4" w:space="8" w:color="C9C2B2"/></w:pBdr>'
         '<w:tabs><w:tab w:val="right" w:pos="10456"/></w:tabs>'
         '<w:spacing w:before="0" w:after="0"/>'
         '<w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr>'
         + run(f"OMNI LAW GAZETTE · ISSUE No. {ISSUE_NUMBER} · {ISSUE_PERIOD}",
               sz=14, color=GRAY, track=30)
         + '<w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:tab/></w:r>'
         + run("Page ", sz=14, color=GRAY) + fld("PAGE")
         + run(" of ", sz=14, color=GRAY) + fld("NUMPAGES")
         + '</w:p>')

    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<w:ftr {ns}>{p}</w:ftr>')


def build_rels_xml():
    base = [
        ('rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'styles.xml'),
        ('rId2', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'settings.xml'),
        ('rId3', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings', 'webSettings.xml'),
        ('rId4', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes', 'footnotes.xml'),
        ('rId5', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes', 'endnotes.xml'),
        ('rId6', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'media/image1.jpeg'),
        ('rId31', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header', 'header1.xml'),
        ('rId32', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer', 'footer1.xml'),
        ('rId33', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable', 'fontTable.xml'),
        ('rId34', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', 'theme/theme1.xml'),
    ]
    rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for rid, typ, target in base:
        rels.append(f'<Relationship Id="{rid}" Type="{typ}" Target="{target}"/>')
    for n, url in enumerate(_hyperlinks):
        rels.append(f'<Relationship Id="rId{100 + n}" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                    f'Target="{esc(url)}" TargetMode="External"/>')
    rels.append('</Relationships>')
    return ''.join(rels)


def main():
    if not os.path.exists(TEMPLATE):
        sys.exit(f"Template not found: {TEMPLATE}")

    if os.path.exists(WORK):
        shutil.rmtree(WORK)
    os.makedirs(WORK, exist_ok=True)

    with zipfile.ZipFile(TEMPLATE) as z:
        z.extractall(WORK)

    doc_xml = build_document_xml()          # populates _hyperlinks
    rels_xml = build_rels_xml()
    footer_xml = build_footer_xml()

    with open(os.path.join(WORK, "word", "document.xml"), "w", encoding="utf-8") as f:
        f.write(doc_xml)
    with open(os.path.join(WORK, "word", "_rels", "document.xml.rels"), "w", encoding="utf-8") as f:
        f.write(rels_xml)
    with open(os.path.join(WORK, "word", "footer1.xml"), "w", encoding="utf-8") as f:
        f.write(footer_xml)

    out = os.path.join(BUILD_DIR, f"OmniLawGazette_Issue{ISSUE_NUMBER}.docx")
    if os.path.exists(out):
        os.remove(out)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(WORK):
            for fn in files:
                full = os.path.join(root, fn)
                arc = os.path.relpath(full, WORK)
                z.write(full, arc)

    print(f"✓  {out}  ({os.path.getsize(out)/1024:.1f} KB, {len(_hyperlinks)} links)")


if __name__ == "__main__":
    main()
