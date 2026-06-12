#!/usr/bin/env python3
"""
Omni Law Gazette — DOCX issue builder.

Rebuilds word/document.xml inside the firm's gazette template
(templates/OmniLawGazette_template.docx), keeping the masthead header,
footer, fonts, styles and hero image intact, and repackages the result
as build/OmniLawGazette_Issue<N>.docx.

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

# ───────────────────────────── content ──────────────────────────────
# Issue No. 2 · review period 5–12 June 2026.
# Register links re-scraped from the official State Register API
# (api.e-qanun.az/framework/{id}) on 12 June 2026.

EQ = "https://e-qanun.az/framework/"

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

# Stories are laid out in pairs: (left story, right story) per grid band.
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

INDEX_TITLE = "All New Legislation · 5–12 June 2026"
INDEX_INTRO = ("Every act entered into the official State Register (e-qanun.az) "
               "during the review period. Click any link to open the full text.")

INDEX = [
    ("Presidential Decree — improvement of the structure and management of the Baku City Executive Authority", "10 June · Gazette item 01", "62002"),
    ("Presidential Decree — improvement of credit-guarantee and interest-subsidy mechanisms for entrepreneurs", "9 June · Gazette item 02", "61987"),
    ("Presidential Decree — additional measures to stimulate non-oil-gas exports and support transport costs", "9 June · Gazette item 03", "61986"),
    ("Presidential Decree — amendments to the Statute of the Agrarian Credit and Development Agency", "8 June · Gazette item 04", "61985"),
    ("Presidential Decree — implementation of the amendment to the Law on Citizens’ Appeals", "10 June · Gazette item 05", "62000"),
    ("Cabinet Decision No. 174 — SOCAR–Gran Tierra Energy onshore oil exploration and production agreement", "9 June · Gazette item 06", "62004"),
    ("Cabinet Decision No. 173 — gas-supply rules aligned with the 2025 Gas Supply Law", "9 June · Gazette item 07", "61994"),
    ("Cabinet Decision No. 171 — Rules on the maintenance and use of public railway infrastructure", "9 June · Gazette item 08", "61993"),
    ("Cabinet Decision No. 168 — Statute of the Energy Efficiency Information System", "9 June · Gazette item 09", "61995"),
    ("Cabinet Decision No. 169 — implementation of the State Security Service digital-analytics system", "9 June · Gazette item 10", "61988"),
    ("Cabinet Decision No. 170 — amendment to the mandatory health-insurance Services Package", "9 June · Gazette item 11", "61990"),
    ("Cabinet Decision No. 175 — list of bodies subordinate to the Ministry of Economy revised", "10 June · Gazette item 12", "61999"),
    ("Cabinet Decision No. 172 — list of state bodies with designated official parking areas", "9 June · In Brief", "62003"),
    ("Presidential Order — financial support to the Azerbaijan Minifootball Federation", "5 June · In Brief", "61973"),
    ("Presidential Order — state decoration for water-management and land-reclamation workers", "5 June · In Brief", "61972"),
    ("Presidential Order — Honour Diploma of the President conferred on R.M. Fətəliyev", "8 June · In Brief", "61984"),
]

# ───────────────────────── XML helpers ──────────────────────────────

NAVY, GOLD, BODY, GRAY = "033B74", "8A6D1E", "3A3A3A", "7A7A7A"
FONT = "Eloquia Text"
FONT_LT = "Eloquia Text ExtLt"

_hyperlinks: list[str] = []  # collected targets; rIds assigned from 100


def esc(t: str) -> str:
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def run(text, *, font=FONT, sz=17, color=BODY, b=False, i=False, u=False):
    rpr = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>'
    if b: rpr += "<w:b/>"
    if i: rpr += "<w:i/>"
    rpr += f'<w:color w:val="{color}"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    if u: rpr += '<w:u w:val="single"/>'
    rpr += '<w:lang w:val="en-US"/>'
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def hyperlink(url, text=None, **rkw):
    _hyperlinks.append(url)
    rid = f"rId{99 + len(_hyperlinks)}"
    return f'<w:hyperlink r:id="{rid}">{run(text or url, u=True, **rkw)}</w:hyperlink>'


def para(runs_xml, *, after=80, jc=None):
    ppr = f'<w:spacing w:after="{after}"/>'
    if jc: ppr += f'<w:jc w:val="{jc}"/>'
    return f'<w:p><w:pPr>{ppr}<w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr>{runs_xml}</w:p>'


def empty_para():
    return '<w:p><w:pPr><w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr></w:p>'


def p_kicker(text):
    return para(run(text, sz=15, color=GOLD, b=True), after=80)


def p_headline(text, sz=24):
    return para(run(text, sz=sz, color=NAVY, b=True), after=40)


def p_lead(bold_part, rest):
    return para(run(bold_part, sz=18, b=True) + run(rest, sz=18), after=80, jc="both")


def p_body(text, sz=17):
    return para(run(text, sz=sz), after=80, jc="both")


def p_brief(topic, rest):
    return para(run(topic, b=True) + run(rest), after=80, jc="both")


def p_source(url):
    return para(run("Source · ", sz=15, color=GOLD, i=True) +
                hyperlink(url, sz=15, color=GOLD, i=True), after=80)


def p_bignum(num):
    return (f'<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="{FONT_LT}" w:hAnsi="{FONT_LT}"/>'
            f'<w:color w:val="B8B8B8"/>'
            f'<w:sz w:val="56"/><w:szCs w:val="56"/><w:lang w:val="en-US"/></w:rPr>'
            f'<w:t>{num}</w:t></w:r></w:p>')


def tc(width, content, *, span=None, valign=None, nowrap=False, vmerge=None):
    pr = f'<w:tcW w:w="{width}" w:type="dxa"/>'
    if span: pr += f'<w:gridSpan w:val="{span}"/>'
    if vmerge: pr += ('<w:vMerge w:val="restart"/>' if vmerge == "restart" else '<w:vMerge/>')
    if nowrap: pr += '<w:noWrap/>'
    if valign: pr += f'<w:vAlign w:val="{valign}"/>'
    return f'<w:tc><w:tcPr>{pr}</w:tcPr>{content or empty_para()}</w:tc>'


def tr(cells, *, cant_split=False):
    trpr = '<w:trPr><w:cantSplit/></w:trPr>' if cant_split else ''
    return f'<w:tr>{trpr}{cells}</w:tr>'


GUTTER = tc(283, None)


def header_row(num_l, head_l, num_r, head_r):
    """5-cell band header row: [num|kicker+headline] · gutter · [num|kicker+headline]."""
    return tr(
        tc(1097, p_bignum(num_l) if num_l else None, valign="center", nowrap=True) +
        tc(3962, head_l) + GUTTER +
        tc(1104, p_bignum(num_r) if num_r else None, valign="center", nowrap=True) +
        tc(4010, head_r),
        cant_split=True
    )


def body_row(left_paras, right_paras):
    return tr(
        tc(5059, left_paras, span=2) + GUTTER +
        tc(5114, right_paras, span=2)
    )


# ───────────────────────── document assembly ────────────────────────

def story_header(st):
    return p_kicker(st["kicker"]) + p_headline(st["headline"], st.get("hsize", 21))


def story_body(st):
    out = "".join(p_body(p) for p in st["paras"])
    out += p_source(st["source"][1])
    return out


def build_document_xml():
    rows = []

    # Band 1 — lead story (left) | story 02 (right)
    lead_header = p_kicker(LEAD["kicker"]) + p_headline(LEAD["headline"], 24)
    lead_body = "".join(p_lead(p[1], p[2]) if p[0] == "lead" else p_body(p[1])
                        for p in LEAD["left"] + LEAD["right"])
    lead_body += p_source(LEAD["source"][1])
    R = STORIES[0]
    rows.append(header_row(LEAD["num"], lead_header, R["num"], story_header(R)))
    rows.append(body_row(lead_body, story_body(R)))

    # Bands 2..6 — story pairs 03|04, 05|06, 07|08, 09|10, 11|12
    pairs = [(1, 2), (3, 4), (5, 6), (7, 8), (9, 10)]
    for li, ri in pairs:
        L, R = STORIES[li], STORIES[ri]
        rows.append(header_row(L["num"], story_header(L), R["num"], story_header(R)))
        rows.append(body_row(story_body(L), story_body(R)))

    # Final band — IN BRIEF across the full page width
    brief_head = p_kicker(IN_BRIEF["kicker"]) + p_headline(IN_BRIEF["headline"], 21)
    brief_body = "".join(p_brief(t, r) for t, r in IN_BRIEF["paras"])
    rows.append(tr(tc(10456, brief_head, span=5), cant_split=True))
    rows.append(tr(tc(10456, brief_body, span=5)))

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

    # Index (final page)
    index = ['<w:p><w:pPr><w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr>'
             '<w:r><w:br w:type="page"/></w:r></w:p>']
    index.append(para(run(INDEX_TITLE, sz=34, color=NAVY, b=True), after=120))
    index.append(para(run(INDEX_INTRO, sz=18), after=160))
    for n, (title, note, doc_id) in enumerate(INDEX, 1):
        index.append(para(
            run(f"{n}. {title}", sz=18, color=NAVY, b=True) +
            run(f"  ({note})", sz=17, color=GRAY),
            after=10))
        index.append(para(hyperlink(EQ + doc_id, sz=17, color=GOLD), after=100))

    sectpr = ('<w:sectPr><w:headerReference w:type="default" r:id="rId31"/>'
              '<w:footerReference w:type="default" r:id="rId32"/>'
              '<w:pgSz w:w="11906" w:h="16838"/>'
              '<w:pgMar w:top="2450" w:right="720" w:bottom="720" w:left="720" w:header="708" w:footer="708" w:gutter="0"/>'
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


def build_rels_xml():
    base = [
        ('rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'styles.xml', None),
        ('rId2', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'settings.xml', None),
        ('rId3', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings', 'webSettings.xml', None),
        ('rId4', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes', 'footnotes.xml', None),
        ('rId5', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes', 'endnotes.xml', None),
        ('rId6', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'media/image1.jpeg', None),
        ('rId31', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header', 'header1.xml', None),
        ('rId32', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer', 'footer1.xml', None),
        ('rId33', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable', 'fontTable.xml', None),
        ('rId34', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', 'theme/theme1.xml', None),
    ]
    rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for rid, typ, target, _ in base:
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

    with open(os.path.join(WORK, "word", "document.xml"), "w", encoding="utf-8") as f:
        f.write(doc_xml)
    with open(os.path.join(WORK, "word", "_rels", "document.xml.rels"), "w", encoding="utf-8") as f:
        f.write(rels_xml)

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
