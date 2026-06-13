#!/usr/bin/env python3
"""Build the Omni Law Gazette in the house newspaper-digest format.

Reproduces the layout/typography of the firm's example edition exactly:
A4 / 0.5in margins; single-column masthead + lead, then a three-column
body of numbered narrative articles, then a single-column footer.

Typography (from the example):
  item number  36pt  Eloquia Text ExtLt  #3B3838
  kicker        7pt  Times New Roman bold  #806000  (uppercase)
  headline     12pt  Times New Roman bold  #1F3864
  body          9pt  Eloquia Text ExtLt   justified
  source        8pt  Times New Roman italic #8A6D1E
  masthead "omni" 72pt / "law gazette" 19pt  Eloquia Text ExtLt
  banner        9pt  Eloquia Text ExtLt   (left bold navy)
"""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

# ── Palette ──────────────────────────────────────────────────────────────────
NAVY  = RGBColor(0x1F, 0x38, 0x64)
GREY  = RGBColor(0x3B, 0x38, 0x38)
GOLD  = RGBColor(0x80, 0x60, 0x00)
GOLD2 = RGBColor(0x8A, 0x6D, 0x1E)
BLACK = RGBColor(0x00, 0x00, 0x00)
BODY_FONT = 'Eloquia Text ExtLt'
SER_FONT  = 'Times New Roman'

# ── Articles (editorially ranked; 14 instruments, 5–12 June 2026) ────────────
LEAD = {
    'n': '01',
    'kicker': 'PRESIDENT · DECREE · 9 JUNE 2026',
    'headline': 'President expands credit guarantees and interest subsidies for entrepreneurs.',
    'body': [
        "The week’s defining act is a presidential decree, signed 9 June, that overhauls the "
        "state’s system of credit support for the private sector. The decree refines the mechanisms "
        "through which the government guarantees bank loans extended to entrepreneurs and subsidises the "
        "interest payable on them, and it amends a series of earlier presidential decrees so that the whole "
        "framework operates under a single, updated set of rules.",
        "In practice, the reform widens access to state-backed financing and lowers the effective cost of "
        "borrowing for private business. By reducing the collateral burden and the interest expense that "
        "small and medium enterprises face, it is designed to channel commercial credit into productive, "
        "non-oil sectors of the economy — manufacturing, agriculture, services and export-oriented production.",
        "It is the most consequential of the fourteen instruments entered into the official State Register "
        "during the review period, and it sits at the centre of a wider economic-policy package. The same day, "
        "the President signed a complementary decree on non-oil-gas exports (item 02), and on 8 June overhauled "
        "the Agrarian Credit and Development Agency (item 03) — three measures that together point firmly "
        "toward diversification away from hydrocarbon revenue.",
    ],
    'source': 'Source · e-qanun.az/framework/62002',
}

ARTICLES = [
    {
        'n': '02',
        'kicker': 'PRESIDENT · DECREE · 9 JUNE 2026',
        'headline': 'A parallel push to reward non-oil exporters.',
        'body': [
            "Signed the same day as the credit-guarantee reform, a second presidential decree introduces "
            "additional measures to stimulate the export of non-oil-and-gas goods and to defray the transport "
            "and logistics costs that weigh on exporters. The instrument extends targeted support to producers "
            "competing in foreign markets, where freight expense often erodes the margin on Azerbaijani goods.",
            "Read together with item 01, the decree confirms a deliberate tilt of fiscal policy toward the "
            "tradable, non-hydrocarbon economy. For manufacturers, agro-processors and logistics operators, the "
            "practical question will be the eligibility thresholds and the per-shipment ceilings, which fall to "
            "be set in the implementing rules.",
        ],
        'source': 'Source · president.az/az/articles/view/72713',
    },
    {
        'n': '03',
        'kicker': 'PRESIDENT · DECREE · 8 JUNE 2026',
        'headline': 'The Agrarian Credit and Development Agency is recast.',
        'body': [
            "On 8 June the President amended the founding statute of the Agrarian Credit and Development Agency "
            "under the Ministry of Agriculture — the 2015 decree (No. 571) that defines the body — "
            "together with related instruments governing state support to farming. The changes refit the agency "
            "that channels concessional finance into the rural economy.",
            "Coming a day before the broader entrepreneur-credit reform, the measure signals that agriculture is "
            "to be a primary beneficiary of the diversification drive. Agribusiness clients should expect revised "
            "lending criteria and, potentially, a wider mandate for the agency.",
        ],
        'source': 'Source · president.az/az/articles/view/72706',
    },
    {
        'n': '04',
        'kicker': 'PRESIDENT · DECREE · 12 JUNE 2026',
        'headline': 'A single rulebook for the country’s airports.',
        'body': [
            "The President approved, on 12 June, a comprehensive set of Rules on the Management of Airports "
            "(Aerodromes), issued under the constitutional power to regulate (Article 109.32) and framed as part "
            "of measures to improve management in civil aviation. The Rules establish the governance model, "
            "operational standards and responsibilities that will bind airport operators.",
            "For parties to concession, ground-handling, leasing or infrastructure agreements at the country’s "
            "airports, the decree introduces a new compliance baseline. Existing contracts should be reviewed "
            "against the Rules, and any transitional provisions noted.",
        ],
        'source': 'Source · president.az/az/articles/view/72735',
    },
    {
        'n': '05',
        'kicker': 'PRESIDENT · DECREE · 10 JUNE 2026',
        'headline': 'The capital’s administration is restructured.',
        'body': [
            "A decree of 10 June reorganises the structure and management of the Baku City Executive Authority, "
            "the body that administers the capital. The reform adjusts the internal architecture and lines of "
            "management of the country’s largest municipal administration.",
            "Governance reforms of this kind tend to ripple outward — into procurement, permitting and the "
            "handling of municipal matters generally. Clients with dealings before the Baku authority should watch "
            "for consequential changes to competences and points of contact.",
        ],
        'source': 'Source · president.az/az/articles/view/72724',
    },
    {
        'n': '06',
        'kicker': 'PRESIDENT · DECREE · 10 JUNE 2026',
        'headline': 'The new citizens’-appeals law takes effect.',
        'body': [
            "The President brought into force the 26 May amendments to the Law on Citizens’ Appeals "
            "(No. 415-VIIQD) and, in the same decree, amended the 2016 Rules (approved by Decree No. 950) on the "
            "handling of citizens’ appeals across state and municipal bodies, state-controlled legal entities "
            "and budget organisations.",
            "The measure tightens the procedural framework that governs how public bodies receive and answer "
            "citizens. Administrative-law practitioners should note the revised record-keeping and response "
            "obligations, which take effect immediately.",
        ],
        'source': 'Source · president.az/az/articles/view/72722',
    },
    {
        'n': '07',
        'kicker': 'CABINET · DECISION No. 171 · 9 JUNE 2026',
        'headline': 'Rules adopted for the national rail network.',
        'body': [
            "The Cabinet of Ministers approved Rules on the maintenance and use of general-use railway transport "
            "infrastructure — a foundational regulation for the rail sector. The Rules govern how the publicly "
            "accessible network is kept up and accessed by operators.",
            "For freight forwarders, infrastructure contractors and any party to a rail-access arrangement, the "
            "decision sets the terms of engagement with the network. It is the most structurally significant of "
            "the Cabinet’s eleven decisions this week.",
        ],
        'source': 'Source · nk.gov.az · Qərar No. 171',
    },
    {
        'n': '08',
        'kicker': 'CABINET · DECISION No. 173 · 9 JUNE 2026',
        'headline': 'The gas-supply rulebook is rewritten.',
        'body': [
            "Implementing the 2025 Law on Gas Supply (No. 233-VIIQ), the Cabinet amended a series of its own "
            "decisions and repealed three legacy instruments — rules dating to 1999 and 2023 on gas "
            "distribution and on the construction and operation of gas installations. The effect is to consolidate "
            "the sector under the current statutory regime.",
            "Distributors and large industrial consumers should map the repealed rules against the new framework "
            "to confirm that internal compliance procedures remain valid.",
        ],
        'source': 'Source · nk.gov.az · Qərar No. 173',
    },
    {
        'n': '09',
        'kicker': 'CABINET · DECISION No. 174 · 10 JUNE 2026',
        'headline': 'State guarantees underpin a new upstream PSA.',
        'body': [
            "The Cabinet approved the Government’s guarantees and obligations under the production-sharing "
            "agreement between SOCAR and Gran Tierra Energy (Azerbaijan) GmbH, covering exploration, development "
            "and production across parts of the Quba, Qusar, Shabran, Siyazan and Khizi districts. The decision "
            "gives the sovereign backing that international upstream investors require.",
            "The award signals continued foreign-investor appetite for Azerbaijani acreage and provides a "
            "template for the contractual architecture of future production-sharing agreements.",
        ],
        'source': 'Source · nk.gov.az · Qərar No. 174',
    },
    {
        'n': '10',
        'kicker': 'CABINET · DECISIONS Nos. 176–178 · 12 JUNE 2026',
        'headline': 'Three new sites earmarked for renewable energy.',
        'body': [
            "In a single sitting on 12 June the Cabinet designated two state-owned parcels in the Hacıqabul "
            "district as renewable-energy sites (Decisions Nos. 177 and 178) and amended its 2024 decision on a "
            "comparable site in Jabrayil (No. 176). The instruments enlarge the land bank available for solar and "
            "wind development.",
            "The Jabrayil amendment, revisiting a 2024 designation, shows the liberated-territories energy "
            "programme being refined in practice. Developers and lenders active in these districts should track "
            "the evolving site parameters.",
        ],
        'source': 'Source · nk.gov.az · Qərarlar No. 176–178',
    },
    {
        'n': '11',
        'kicker': 'CABINET · DECISION No. 169 · 9 JUNE 2026',
        'headline': 'The security service gains a central analytics system.',
        'body': [
            "The Cabinet amended several of its decisions to give effect to a 2025 presidential decree (No. 533) "
            "establishing the Centralised Information and Digital Analytics System of the State Security Service. "
            "The decision aligns subordinate regulation with the new architecture for state-security data.",
            "While operationally sensitive, the measure forms part of a broader digitisation of the security "
            "apparatus and may bear on data-sharing obligations imposed on regulated entities.",
        ],
        'source': 'Source · nk.gov.az · Qərar No. 169',
    },
    {
        'n': '12',
        'kicker': 'CABINET · DECISION No. 170 · 9 JUNE 2026',
        'headline': 'The mandatory-insurance benefit package is adjusted.',
        'body': [
            "The Cabinet amended the Mandatory Health Insurance Services Package — the schedule of services "
            "covered under the compulsory scheme, first approved in 2020. Changes to the package alter the scope "
            "of care that insured persons may claim and the corresponding obligations on providers.",
            "Employers and healthcare providers should confirm how the revised schedule affects coverage and "
            "reimbursement before the change takes effect.",
        ],
        'source': 'Source · nk.gov.az · Qərar No. 170',
    },
    {
        'n': '13',
        'kicker': 'CABINET · DECISION No. 168 · 9 JUNE 2026',
        'headline': 'A framework for energy-efficiency data.',
        'body': [
            "The Cabinet approved the Regulation on the Energy Efficiency Information System, establishing the "
            "legal basis for collecting, monitoring and reporting energy-consumption data across the economy. The "
            "system is intended to underpin efficiency policy with reliable measurement.",
            "Energy-intensive enterprises should anticipate reporting duties once the system becomes operational.",
        ],
        'source': 'Source · nk.gov.az · Qərar No. 168',
    },
    {
        'n': '14',
        'kicker': 'CABINET · ORDERS · 9–10 JUNE 2026',
        'headline': 'Four orders put recent presidential decrees into effect.',
        'body': [
            "Alongside its decisions, the Cabinet issued four implementing orders. Two of 10 June give effect to "
            "presidential decrees on the ASAN citizen-service agency (No. 678) and on Finance Ministry "
            "representation on the boards of state-owned entities (No. 675). Two of 9 June implement the renaming "
            "of territorial units in the Quba district (under Law No. 404-VIIQ) and changes to the State Agency "
            "for the Protection of Strategic Facilities.",
            "Individually routine, together they show the machinery by which presidential policy is translated "
            "into administrative practice within days of signature.",
        ],
        'source': 'Source · nk.gov.az · Sərəncamlar No. 389s, 400s, 402s, 403s',
    },
]

# ── Build ────────────────────────────────────────────────────────────────────
doc = Document()
st = doc.styles['Normal']
st.font.name = BODY_FONT
st.font.size = Pt(9)
st.paragraph_format.space_after = Pt(0)
st.paragraph_format.line_spacing = 1.0


def set_cols(section, num):
    sectPr = section._sectPr
    cols = sectPr.find(qn('w:cols'))
    if cols is None:
        cols = OxmlElement('w:cols'); sectPr.append(cols)
    cols.set(qn('w:num'), str(num))
    cols.set(qn('w:space'), '708')


def run(p, text, *, font=BODY_FONT, size=9, bold=False, italic=False, color=None):
    r = p.add_run(text)
    r.font.name = font
    # ensure east-asian/complex fallback uses same font name
    rPr = r._r.get_or_add_rPr(); rfonts = rPr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts'); rPr.insert(0, rfonts)
    for a in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rfonts.set(qn(a), font)
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    if color is not None:
        r.font.color.rgb = color
    return r


def para(align=None, before=0, after=0, line=1.0):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after); pf.line_spacing = line
    return p


def bottom_border(el, color='1F3864', sz='4'):
    """Apply a bottom border to a paragraph element."""
    pPr = el._p.get_or_add_pPr()
    pbdr = pPr.find(qn('w:pBdr'))
    if pbdr is None:
        pbdr = OxmlElement('w:pBdr'); pPr.append(pbdr)
    b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), sz)
    b.set(qn('w:space'), '4'); b.set(qn('w:color'), color)
    pbdr.append(b)


# ── Masthead ────────────────────────────────────────────────────────────────
p = para(align=WD_ALIGN_PARAGRAPH.LEFT, after=0)
run(p, 'omni', font=BODY_FONT, size=54, color=BLACK)
run(p, '  law ', font=BODY_FONT, size=17, color=BLACK)
run(p, 'gazette', font=BODY_FONT, size=17, color=NAVY)
bottom_border(p, sz='6')

# Banner row (two columns simulated with a tab)
p = para(after=8)
run(p, 'WEEKLY LEGISLATIVE DIGEST · REPUBLIC OF AZERBAIJAN', font=BODY_FONT, size=9, bold=True, color=NAVY)
# right-aligned issue label via right tab stop
from docx.enum.text import WD_TAB_ALIGNMENT
p.paragraph_format.tab_stops.add_tab_stop(Inches(7.27), WD_TAB_ALIGNMENT.RIGHT)
run(p, '\t', size=9)
run(p, 'ISSUE No. 1   ·   5–12 JUNE 2026', font=BODY_FONT, size=9, color=GREY)
bottom_border(p, sz='4')
para(after=4)  # spacer


def render_article(item, lead=False):
    # number
    p = para(align=WD_ALIGN_PARAGRAPH.CENTER, before=(12 if not lead else 8), after=0)
    run(p, item['n'], font=BODY_FONT, size=(36 if lead else 28), bold=True, color=GREY)
    # kicker
    p = para(align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=1)
    run(p, item['kicker'], font=SER_FONT, size=7, bold=True, color=GOLD)
    # headline
    p = para(align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=4)
    run(p, item['headline'], font=SER_FONT, size=(13 if lead else 12), bold=True, color=NAVY)
    # body
    for i, par in enumerate(item['body']):
        p = para(align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=0, after=4, line=1.05)
        run(p, par, font=BODY_FONT, size=9, color=BLACK)
    # source
    p = para(align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=0, after=6)
    run(p, item['source'], font=SER_FONT, size=8, italic=True, color=GOLD2)


# ── Section 0: lead article (single column) ─────────────────────────────────
render_article(LEAD, lead=True)

# close section 0 → make it 1 column
doc.add_section(WD_SECTION.CONTINUOUS)
set_cols(doc.sections[0], 1)

# ── Section 1: items 02–14 (three columns) ──────────────────────────────────
for item in ARTICLES:
    render_article(item, lead=False)

# close section 1 → make it 3 columns
doc.add_section(WD_SECTION.CONTINUOUS)
set_cols(doc.sections[1], 3)

# ── Section 2: footer (single column) ───────────────────────────────────────
p = para(before=6, after=2)
bottom_border(p, sz='4')
p = para(align=WD_ALIGN_PARAGRAPH.JUSTIFY, after=2)
run(p, 'Omni Law Gazette', font=BODY_FONT, size=7.5, bold=True, color=NAVY)
run(p, '  ·  Internal weekly digest of the legislative activity of the Republic of Azerbaijan, '
       'compiled by Omni Law Firm for informational purposes only. It does not constitute legal advice; '
       'consult the consolidated texts on e-qanun.az before relying on any item. '
       'Sources: president.az · nk.gov.az · meclis.gov.az · e-qanun.az · constcourt.gov.az',
    font=BODY_FONT, size=7.5, color=GREY)
p = para(after=0)
run(p, 'Issue No. 1  ·  Reporting period 5–12 June 2026  ·  © 2026 Omni Law Firm',
    font=BODY_FONT, size=7.5, bold=True, color=GREY)
set_cols(doc.sections[2], 1)

# A4 + 0.5in margins on all sections
for s in doc.sections:
    s.page_width = Inches(8.27); s.page_height = Inches(11.69)
    s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.5)

out = os.path.join(os.path.dirname(__file__), 'gazettes',
                   'Omni-Law-Gazette-Issue-1-2026-06-12.docx')
doc.save(out)
print('DOCX written:', out)
print('sections/cols:',
      [(i, s._sectPr.find(qn('w:cols')).get(qn('w:num')) if s._sectPr.find(qn('w:cols')) is not None else '1')
       for i, s in enumerate(doc.sections)])


# ── HTML twin (for high-fidelity PDF rendering via Chromium) ─────────────────
import html as _html


def _article_html(item, lead=False):
    paras = ''.join(f'<p class="body">{_html.escape(b)}</p>' for b in item['body'])
    return f'''<article class="{'lead' if lead else 'item'}">
      <div class="num">{item['n']}</div>
      <div class="kicker">{_html.escape(item['kicker'])}</div>
      <h2 class="headline">{_html.escape(item['headline'])}</h2>
      {paras}
      <p class="source">{_html.escape(item['source'])}</p>
    </article>'''


html_doc = f'''<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<style>
  @page {{ size: A4; margin: 12mm 10mm; }}
  * {{ box-sizing: border-box; }}
  body {{ margin: 0; color: #000;
    font-family: "Eloquia Text ExtLt", "Segoe UI", "Helvetica Neue", Arial, sans-serif; }}
  .masthead {{ display: flex; align-items: baseline; gap: 10px;
    border-bottom: 1.5px solid #1F3864; padding-bottom: 4px; }}
  .masthead .omni {{ font-size: 54pt; letter-spacing: -1px; }}
  .masthead .lawg {{ font-size: 17pt; }}
  .masthead .lawg b {{ color: #1F3864; }}
  .banner {{ display: flex; justify-content: space-between; align-items: baseline;
    border-bottom: 0.75px solid #1F3864; padding: 4px 0 5px; margin-bottom: 10px; }}
  .banner .left {{ font-size: 9pt; font-weight: 700; color: #1F3864; letter-spacing: 0.3px; }}
  .banner .right {{ font-size: 9pt; color: #3B3838; }}
  .num {{ text-align: center; font-weight: 700; color: #3B3838; line-height: 1; }}
  .item .num {{ font-size: 26pt; margin-top: 10px; }}
  .lead .num {{ font-size: 36pt; margin-top: 4px; }}
  .kicker {{ text-align: center; font-family: "Times New Roman", Georgia, serif;
    font-size: 7pt; font-weight: 700; color: #806000; letter-spacing: 0.4px;
    text-transform: uppercase; margin-top: 2px; }}
  .headline {{ text-align: center; font-family: "Times New Roman", Georgia, serif;
    font-weight: 700; color: #1F3864; margin: 2px 0 6px; }}
  .item .headline {{ font-size: 12pt; line-height: 1.2; }}
  .lead .headline {{ font-size: 14pt; }}
  .body {{ text-align: justify; font-size: 9pt; line-height: 1.45; margin: 0 0 6px; }}
  .source {{ text-align: justify; font-family: "Times New Roman", Georgia, serif;
    font-size: 8pt; font-style: italic; color: #8A6D1E; margin: 0 0 4px; }}
  .lead {{ margin-bottom: 6px; }}
  .columns {{ column-count: 3; column-gap: 14px; column-rule: 0.5px solid #d9d9d9; }}
  .columns .item {{ break-inside: avoid-column; -webkit-column-break-inside: avoid; }}
  .footer {{ border-top: 0.75px solid #1F3864; margin-top: 10px; padding-top: 5px; }}
  .footer .d {{ text-align: justify; font-size: 7.5pt; color: #3B3838; }}
  .footer .d b {{ color: #1F3864; }}
  .footer .c {{ font-size: 7.5pt; font-weight: 700; color: #3B3838; margin-top: 3px; }}
</style></head><body>
  <div class="masthead">
    <span class="omni">omni</span>
    <span class="lawg">law <b>gazette</b></span>
  </div>
  <div class="banner">
    <span class="left">WEEKLY LEGISLATIVE DIGEST · REPUBLIC OF AZERBAIJAN</span>
    <span class="right">ISSUE No. 1 &nbsp;·&nbsp; 5–12 JUNE 2026</span>
  </div>
  {_article_html(LEAD, lead=True)}
  <div class="columns">
    {''.join(_article_html(a) for a in ARTICLES)}
  </div>
  <div class="footer">
    <p class="d"><b>Omni Law Gazette</b> · Internal weekly digest of the legislative activity of the
      Republic of Azerbaijan, compiled by Omni Law Firm for informational purposes only. It does not
      constitute legal advice; consult the consolidated texts on e-qanun.az before relying on any item.
      Sources: president.az · nk.gov.az · meclis.gov.az · e-qanun.az · constcourt.gov.az</p>
    <p class="c">Issue No. 1 &nbsp;·&nbsp; Reporting period 5–12 June 2026 &nbsp;·&nbsp; © 2026 Omni Law Firm</p>
  </div>
</body></html>'''

html_out = out.replace('.docx', '.html')
with open(html_out, 'w', encoding='utf-8') as f:
    f.write(html_doc)
print('HTML written:', html_out)
