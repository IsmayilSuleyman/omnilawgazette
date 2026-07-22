#!/usr/bin/env python3
"""Build the Omni Law Gazette probe issue (1-12 June 2026) from the approved template.

v2 — review fixes:
- kicker/headline/source paragraphs now byte-match the template's formatting
  (kicker: spacing 0 + trailing break + 1pt par mark; headline: default spacing;
   source: 'Source · ' 10pt + URL 8pt; stories separated by an empty 9pt paragraph)
- every source citation is a full, HTTP-verified URL
- e-qanun.az/framework/62002 replaced by the verified president.az link
  (e-qanun is a SPA that returns 200 for any ID — unverifiable server-side)
"""
import shutil, zipfile, html, os, subprocess

TPL = '/root/.claude/uploads/f3ae65b9-b060-523c-abdf-aff24efe5d35/3d1a24d4-example01.docx'
OUT = '/tmp/omni-law-gazette-issue-01-probe-1-12-june-2026.docx'
SRC = '/tmp/docx_extract/word/document.xml'

xml = open(SRC, encoding='utf-8').read()

def esc(t): return html.escape(t, quote=False)

EL = '<w:rFonts w:ascii="Eloquia Text ExtLt" w:hAnsi="Eloquia Text ExtLt"/>'
ELCS = '<w:rFonts w:ascii="Eloquia Text ExtLt" w:hAnsi="Eloquia Text ExtLt" w:cs="Times New Roman"/>'
TNR = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
LANG = '<w:lang w:val="en-US"/>'

def p_num(n):
    rpr = f'{ELCS}<w:b/><w:color w:val="3B3838"/><w:sz w:val="72"/><w:szCs w:val="72"/>{LANG}'
    return (f'<w:p><w:pPr><w:spacing w:after="0"/><w:jc w:val="center"/><w:rPr>{rpr}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:t>{n}</w:t></w:r></w:p>')

def p_kicker(t):
    # template: spacing 0, centred, trailing <w:br/>, 1pt paragraph mark
    rpr = f'{TNR}<w:b/><w:color w:val="806000"/><w:sz w:val="14"/><w:szCs w:val="14"/>{LANG}'
    mark = f'{TNR}<w:b/><w:color w:val="806000"/><w:sz w:val="2"/><w:szCs w:val="2"/>{LANG}'
    return (f'<w:p><w:pPr><w:spacing w:after="0"/><w:jc w:val="center"/><w:rPr>{mark}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:t>{esc(t)}</w:t></w:r>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:br/></w:r></w:p>')

def p_headline(t):
    # template: default spacing, centred
    rpr = f'{TNR}<w:b/><w:bCs/><w:color w:val="1F3864"/><w:sz w:val="24"/><w:szCs w:val="24"/>{LANG}'
    return (f'<w:p><w:pPr><w:jc w:val="center"/><w:rPr>{rpr}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:t>{esc(t)}</w:t></w:r></w:p>')

def p_body(t):
    rpr = f'{EL}<w:sz w:val="18"/><w:szCs w:val="18"/>{LANG}'
    return (f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{rpr}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc(t)}</w:t></w:r></w:p>')

def p_source(url):
    # template: 'Source · ' at 10pt, URL at 8pt, italic gold, default spacing
    r1 = f'{TNR}<w:i/><w:color w:val="8A6D1E"/><w:sz w:val="20"/><w:szCs w:val="20"/>{LANG}'
    r2 = f'{TNR}<w:i/><w:color w:val="8A6D1E"/><w:sz w:val="16"/><w:szCs w:val="16"/>{LANG}'
    return (f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{TNR}<w:sz w:val="20"/><w:szCs w:val="20"/>{LANG}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r1}</w:rPr><w:t xml:space="preserve">Source · </w:t></w:r>'
            f'<w:r><w:rPr>{r2}</w:rPr><w:t>{esc(url)}</w:t></w:r></w:p>')

def p_gap():
    # empty 9pt paragraph: the template's story separator
    rpr = f'{EL}<w:sz w:val="18"/><w:szCs w:val="18"/>{LANG}'
    return f'<w:p><w:pPr><w:rPr>{rpr}</w:rPr></w:pPr></w:p>'

def story(num, kicker, headline, paras, source):
    return (p_gap() + p_num(num) + p_kicker(kicker) + p_headline(headline)
            + ''.join(p_body(p) for p in paras) + p_source(source))

def p_section_head(title, sub):
    return (f'<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="4" w:color="1F3864"/></w:pBdr>'
            f'<w:spacing w:after="0"/></w:pPr>'
            f'<w:r><w:rPr>{TNR}<w:b/><w:bCs/><w:color w:val="1F3864"/><w:sz w:val="28"/><w:szCs w:val="28"/>{LANG}</w:rPr>'
            f'<w:t>{esc(title)}</w:t></w:r></w:p>'
            f'<w:p><w:pPr><w:spacing w:after="160"/></w:pPr>'
            f'<w:r><w:rPr>{TNR}<w:b/><w:color w:val="806000"/><w:sz w:val="14"/><w:szCs w:val="14"/>{LANG}</w:rPr>'
            f'<w:t>{esc(sub)}</w:t></w:r></w:p>')

def p_entry(lead, title, src, tag=None):
    r = (f'<w:p><w:pPr><w:spacing w:after="80"/><w:jc w:val="both"/></w:pPr>'
         f'<w:r><w:rPr>{EL}<w:b/><w:bCs/><w:color w:val="1F3864"/><w:sz w:val="18"/><w:szCs w:val="18"/>{LANG}</w:rPr>'
         f'<w:t xml:space="preserve">{esc(lead)} — </w:t></w:r>'
         f'<w:r><w:rPr>{EL}<w:sz w:val="18"/><w:szCs w:val="18"/>{LANG}</w:rPr>'
         f'<w:t xml:space="preserve">{esc(title)}</w:t></w:r>')
    if tag:
        r += (f'<w:r><w:rPr>{EL}<w:b/><w:color w:val="806000"/><w:sz w:val="18"/><w:szCs w:val="18"/>{LANG}</w:rPr>'
              f'<w:t xml:space="preserve">  [{esc(tag)}]</w:t></w:r>')
    r += (f'<w:r><w:rPr>{TNR}<w:i/><w:color w:val="8A6D1E"/><w:sz w:val="16"/><w:szCs w:val="16"/>{LANG}</w:rPr>'
          f'<w:t xml:space="preserve">  · {esc(src)}</w:t></w:r></w:p>')
    return r

def p_note(t):
    return (f'<w:p><w:pPr><w:spacing w:after="160"/><w:jc w:val="both"/></w:pPr>'
            f'<w:r><w:rPr>{EL}<w:i/><w:sz w:val="18"/><w:szCs w:val="18"/>{LANG}</w:rPr>'
            f'<w:t xml:space="preserve">{esc(t)}</w:t></w:r></w:p>')

# ---------------- masthead + lead-story tweaks ----------------
# Wrap-proof the masthead: with a substituted font (no Eloquia installed) the
# wordmark cell breaks "omni" mid-word and the issue date wraps. <w:noWrap/>
# lets the auto-layout table widen the cell instead of wrapping.
xml = xml.replace('<w:vAlign w:val="center"/></w:tcPr>',
                  '<w:noWrap/><w:vAlign w:val="center"/></w:tcPr>')
xml = xml.replace('<w:vAlign w:val="bottom"/></w:tcPr>',
                  '<w:noWrap/><w:vAlign w:val="bottom"/></w:tcPr>')
xml = xml.replace('<w:tcW w:w="2801" w:type="dxa"/></w:tcPr>',
                  '<w:tcW w:w="2801" w:type="dxa"/><w:noWrap/></w:tcPr>')
# drop the strapline cell's trailing tabs/spaces so the date cell can take the width
xml = xml.replace('<w:r><w:rPr><w:rFonts w:ascii="Eloquia Text ExtLt" w:hAnsi="Eloquia Text ExtLt"/>'
                  '<w:sz w:val="18"/><w:szCs w:val="18"/><w:lang w:val="en-US"/></w:rPr><w:tab/></w:r>', '')
xml = xml.replace('<w:t xml:space="preserve">           </w:t>', '<w:t xml:space="preserve"></w:t>')
# Rebalance column widths: the tables span the full page, so noWrap alone cannot
# expand a cell. Move width from the empty banner cell into the wordmark cell
# (their shared boundary with the "law gazette" cell stays put), and from the
# strapline cell into the issue-date cell.
for old, new in [('w:w="2880"', 'w:w="2200"'), ('w:w="3045"', 'w:w="2200"'),
                 ('w:w="3377"', 'w:w="4057"'), ('w:w="3046"', 'w:w="4057"'),
                 ('w:w="7655"', 'w:w="6855"'), ('w:w="2801"', 'w:w="3601"')]:
    assert old in xml, old
    xml = xml.replace(old, new)

xml = xml.replace('  5–12 JUNE 2026', '  1–12 JUNE 2026')
xml = xml.replace('fourteen instruments', 'nineteen instruments')
# e-qanun is a SPA returning 200 for any ID; cite the verified president.az page instead
xml = xml.replace('e-qanun.az/framework/62002', 'president.az/az/articles/view/72714')

P = 'president.az/az/articles/view/'
NK_402 = 'nk.gov.az/az/senedler/serencamlar/dovlet-mulkiyyetinde-olan-ve-paylarinin-sehmlerini-9548'
NK_403 = 'nk.gov.az/az/senedler/serencamlar/azerbaycan-respublikasinin-prezidenti-yaninda-vete-9547'
NK_175 = 'nk.gov.az/az/senedler/qerarlar/azerbaycan-respublikasi-nazirler-kabinetinin-2010-9546'
def MB(i): return f'meclis.gov.az/news-layihe.php?id={i}&lang=az'

# ---------------- featured stories 02-07 ----------------
stories = ''
stories += story('02', 'PRESIDENT · DECREE · 9 JUNE 2026',
    'Non-oil exporters win transport-cost support in twin stimulus decree.',
    ["Signed the same day as the credit-support decree (item 01), a second presidential decree targets the other side of the diversification ledger: getting Azerbaijani non-oil-and-gas goods to foreign markets. The instrument provides for the stimulation of non-oil-gas exports and for state support of the transport costs that exporters incur in carrying goods abroad.",
     "For exporters of manufactured goods and agricultural produce, the measure addresses one of the most frequently cited barriers to competitiveness — the cost of carriage from a land-locked market to distant buyers. Read together with item 01, the two 9 June decrees subsidise both the financing of production and the logistics of selling it abroad."],
    P + '72713')
stories += story('03', 'PRESIDENT · DECREE · 8 JUNE 2026',
    'Agrarian credit machinery reorganised by amendments to four decrees.',
    ["A day earlier, on 8 June, the President amended four of his earlier decrees governing agricultural support: the 2015 statute of the Agrarian Credit and Development Agency, the 2018 decree on state support to agriculture and agrarian leasing, the 2019 decree creating the new agrarian subsidy mechanism, and the 2019 statute of the Electronic Agriculture information system.",
     "The decree completes the week's economic triptych: agrarian credit (this item), entrepreneurial credit (item 01) and export logistics (item 02) were all overhauled within forty-eight hours, a coordinated push toward the non-oil economy."],
    P + '72706')
stories += story('04', 'PRESIDENT · DECREE · 2 JUNE 2026',
    'National Cybersecurity Agency established with its own charter.',
    ["On 2 June the President signed a decree on the improvement of governance in the field of cybersecurity — amending several earlier decrees and orders and revoking one — and by a companion instrument approved the charter of a National Cybersecurity Agency. Together the two acts consolidate state responsibility for cyber defence in a single dedicated body.",
     "The reform will matter to any organisation that operates critical information infrastructure or handles regulated data: supervisory competence, incident reporting lines and certification expectations now converge on the new agency. Implementing acts defining its powers in detail should be expected in the coming weeks.",
     "A third decree signed the same day amended the founding decree of the State Agency for Citizen Service and Social Innovations (ASAN); the Cabinet of Ministers issued its implementing order on 10 June."],
    P + '72657')
stories += story('05', 'PRESIDENT · DECREE · 10 JUNE 2026',
    'Baku city government restructured.',
    ["A 10 June decree improves the structure and management of the Baku City Executive Authority, reorganising the administration of the capital.",
     "Restructuring of the capital's executive authority touches permitting, municipal services and local administrative practice — practitioners with matters pending before city structures should verify the competence of their counterpart bodies once the new structure is published."],
    P + '72724')
stories += story('06', 'MILLI MAJLIS · LAW · 10 JUNE 2026',
    "Citizens' appeals law amended; application decree signed same day.",
    ["The law amending the Law on Citizens' Appeals (No. 415-VIIQD), adopted by the Milli Majlis on 26 May, was signed and published on 10 June together with the presidential decree on its application.",
     "The Law on Citizens' Appeals governs how state bodies must receive, register and answer applications and complaints from individuals and legal entities — the amendments therefore bear directly on administrative practice and on the deadlines public bodies must observe."],
    P + '72721')
stories += story('07', 'CABINET OF MINISTERS · ORDER · 10 JUNE 2026',
    'Finance Ministry takes seats on state-owned company boards.',
    ["By Order No. 402s of 10 June, the Cabinet of Ministers regulated the representation of the Ministry of Finance on the supervisory boards (boards of trustees or directors) of legal entities in state ownership, of entities whose controlling shareholding belongs to the state, and of public legal entities created on behalf of the state.",
     "The order strengthens central fiscal oversight of the state-owned enterprise sector — a corporate-governance development relevant to anyone transacting with, financing or auditing state-owned entities."],
    NK_402)

# ---------------- full register ----------------
reg = p_section_head('THE FULL REGISTER', 'EVERY INSTRUMENT RECORDED · 1–12 JUNE 2026 · CHRONOLOGICAL')
entries = [
 ('1 JUNE · MILLI MAJLIS · LAW', 'On renaming certain territorial units of the Quba district (No. 404-VIIQ, adopted 12 May), signed and published.', P + '72633', None),
 ('1 JUNE · PRESIDENT · ORDER', 'On the application of Law No. 404-VIIQ on the Quba district territorial units.', P + '72632', None),
 ('1 JUNE · PRESIDENT · ORDER', 'On awarding Q.S. Qurbanov the "Dostluq" Order.', P + '72631', None),
 ('2 JUNE · PRESIDENT · DECREE', 'On the improvement of governance in the field of cybersecurity.', P + '72657', 'see item 04'),
 ('2 JUNE · PRESIDENT · DECREE', 'Approving the Charter of the National Cybersecurity Agency.', P + '72658', 'see item 04'),
 ('2 JUNE · PRESIDENT · DECREE', 'No. 678 — amending Decree No. 706 of 5 September 2012 on the State Agency for Citizen Service and Social Innovations (ASAN).', P + '72656', None),
 ('3 JUNE · PRESIDENT · DECREE', 'Amending the rules approved by Decree No. 725 of 13 January 2016.', P + '72672', None),
 ('5 JUNE · PRESIDENT · ORDER', 'On awarding employees of the water-management and melioration sector.', P + '72697', None),
 ('5 JUNE · PRESIDENT · ORDER', 'On financial support to the Azerbaijan Minifootball Federation.', P + '72696', None),
 ('8 JUNE · PRESIDENT · DECREE', 'Amending four presidential decrees on the Agrarian Credit and Development Agency, agrarian subsidies and agricultural leasing.', P + '72706', 'see item 03'),
 ('8 JUNE · PRESIDENT · ORDER', 'On awarding R.M. Fataliyev the Honorary Diploma of the President.', P + '72705', None),
 ('9 JUNE · PRESIDENT · DECREE', 'On improving the guarantee and interest-subsidy mechanisms for credits extended to entrepreneurs.', P + '72714', 'see item 01'),
 ('9 JUNE · PRESIDENT · DECREE', 'On stimulating the export of non-oil-gas goods and supporting transport costs.', P + '72713', 'see item 02'),
 ('10 JUNE · MILLI MAJLIS · LAW', "Amending the Law on Citizens' Appeals (No. 415-VIIQD, adopted 26 May), signed and published.", P + '72721', 'see item 06'),
 ('10 JUNE · PRESIDENT · DECREE', "On the application of Law No. 415-VIIQD amending the Law on Citizens' Appeals.", P + '72722', 'see item 06'),
 ('10 JUNE · PRESIDENT · DECREE', 'On improving the structure and management of the Baku City Executive Authority.', P + '72724', 'see item 05'),
 ('10 JUNE · CABINET · RESOLUTION', 'No. 175 — amending the list of bodies subordinate to the Ministry of Economy (Resolution No. 228 of 30 November 2010).', NK_175, None),
 ('10 JUNE · CABINET · ORDER', 'No. 402s — on representation of the Ministry of Finance on supervisory boards of state-owned legal entities and public legal entities.', NK_402, 'see item 07'),
 ('10 JUNE · CABINET · ORDER', 'No. 403s — implementing Presidential Decree No. 678 of 2 June 2026 (ASAN agency).', NK_403, None),
]
for e in entries: reg += p_entry(*e)
reg += p_note('Constitutional Court: no decisions could be retrieved for the review period — the Court’s website (constcourt.gov.az) was unavailable at the time of compilation. Any decisions published in the window will be carried in the next issue.')

# ---------------- bills watch ----------------
bills = p_section_head('MILLI MAJLIS · BILLS WATCH', 'LEGISLATION IN PROGRESS · DOCKETED 5–12 JUNE 2026')
bill_entries = [
 ('12 JUNE', 'Report on the execution of the 2025 state budget — placed before Parliament.', MB(2649)),
 ('12 JUNE', 'Accession to the International Convention on the Control of Harmful Anti-Fouling Systems on Ships.', MB(2647)),
 ('12 JUNE', 'Establishment of an embassy of Azerbaijan in Lisbon, Portugal.', MB(2645)),
 ('12 JUNE', 'Amendments to the Land Code and related laws.', MB(2644)),
 ('12 JUNE', 'Amendment to the law designating Shusha the cultural capital.', MB(2648)),
 ('12 JUNE', 'Amendments to forest-protection legislation.', MB(2637)),
 ('12 JUNE', 'Amendments to the Code of Administrative Violations.', MB(2638)),
 ('12 JUNE', 'Amendments to the Law on Civil Service.', MB(2642)),
 ('12 JUNE', 'Amendments to administrative-procedure and information-protection legislation.', MB(2643)),
 ('12 JUNE', 'Amendments to the Housing Code and legislation on internally displaced persons.', MB(2646)),
 ('5 JUNE', 'Amendments to the Law on Diplomatic Service.', MB(2635)),
 ('5 JUNE', 'Amendments to the Law on Combating Corruption.', MB(2641)),
 ('5 JUNE', 'Amendments to the Criminal Procedure Code.', MB(2640)),
 ('5 JUNE', 'Amendments to the Criminal Procedure Code and sentence-execution legislation.', MB(2639)),
 ('5 JUNE', 'Amendments to the Customs Code.', MB(2630)),
 ('CARRIED OVER · 26 MAY', 'New Law on Financial Leasing — a full replacement statute for the leasing market, still in committee.', MB(2624)),
]
for d, t, s in bill_entries: bills += p_entry(d, t, s)
bills += p_note('Watch points: the 2025 budget-execution report opens the annual fiscal-review cycle, and the draft Financial Leasing law would replace the current statute outright — both worth tracking to adoption.')

footer = (f'<w:p><w:pPr><w:spacing w:before="240"/><w:jc w:val="center"/></w:pPr>'
          f'<w:r><w:rPr>{TNR}<w:b/><w:color w:val="1F3864"/><w:sz w:val="14"/><w:szCs w:val="14"/>{LANG}</w:rPr>'
          f'<w:t>OMNI LAW GAZETTE · COMPILED FOR INTERNAL CIRCULATION · SOURCES: E-QANUN.AZ, PRESIDENT.AZ, NK.GOV.AZ, MECLIS.GOV.AZ, CONSTCOURT.GOV.AZ</w:t></w:r></w:p>')

# ---------------- splice ----------------
src_pos = xml.find('articles/view/72714')          # lead story source (replaced above)
after_lead = xml.find('</w:p>', src_pos) + len('</w:p>')
sect1 = xml.find('<w:sectPr', 20000)
sect1_para_start = xml.rfind('<w:p ', 0, sect1)
sect1_para_end = xml.find('</w:p>', sect1) + len('</w:p>')
final_sect = xml.find('<w:sectPr', sect1_para_end)

new_xml = (xml[:after_lead]
           + stories
           + xml[sect1_para_start:sect1_para_end]
           + reg + bills + footer
           + xml[final_sect:])

open(SRC, 'w', encoding='utf-8').write(new_xml)

if os.path.exists(OUT): os.remove(OUT)
shutil.copy(TPL, OUT)
os.chdir('/tmp/docx_extract')
subprocess.run(['zip', '-q', OUT, 'word/document.xml'], check=True)
print('written', OUT, os.path.getsize(OUT))
