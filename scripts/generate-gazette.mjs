/**
 * Omni Law Gazette — Issue generator
 * Generates a newspaper-style PDF for the weekly legislative digest.
 * Run: node scripts/generate-gazette.mjs
 */

import PDFDocument from "pdfkit";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.join(__dirname, "../gazette-issue.pdf");

// ── Palette ──────────────────────────────────────────────────────────────────
const BLACK      = "#0A0A0A";
const NAVY       = "#0C1A33";
const CRIMSON    = "#9B1C1C";
const GOLD       = "#B8860B";
const RULE_GRAY  = "#CCCCCC";
const TEXT_GRAY  = "#333333";
const LIGHT_GRAY = "#F5F5F0";
const WHITE      = "#FFFFFF";
const SILVER     = "#666666";

// ── Gazette data ─────────────────────────────────────────────────────────────
const ISSUE_NUMBER = 2;
const WEEK_START   = "5 June 2026";
const WEEK_END     = "12 June 2026";
const PUBLISHED    = "12 June 2026";

const GAZETTE_TITLE   = "OMNI LAW GAZETTE";
const GAZETTE_SUB     = "Weekly Digest of Legislative Developments · Republic of Azerbaijan";

const data = {
  presidential_decrees: [
    {
      title_az: "Bakı Şəhər İcra Hakimiyyətinin strukturunun və idarə edilməsinin təkmilləşdirilməsi haqqında",
      title_en: "On improving the structure and management of the Baku City Executive Authority",
      type: "Fərman (Decree)",
      date: "10 June 2026",
      source: "president.az",
    },
    {
      title_az: "Sahibkarlıq subyektlərinə verilən kreditlər üzrə təminat və faiz subsidiyası mexanizmlərinin təkmilləşdirilməsi haqqında",
      title_en: "On improving guarantee and interest-subsidy mechanisms for loans extended to business entities",
      type: "Fərman (Decree)",
      date: "9 June 2026",
      source: "president.az",
    },
    {
      title_az: "Qeyri-neft-qaz mallarının ixracının stimullaşdırılması və daşıma xərclərinin dəstəklənməsi haqqında",
      title_en: "On stimulating the export of non-oil-and-gas goods and supporting transportation costs",
      type: "Fərman (Decree)",
      date: "9 June 2026",
      source: "president.az",
    },
    {
      title_az: "\"Vətəndaşların müraciətləri haqqında\" Azərbaycan Respublikasının Qanununda dəyişiklik edilməsinin tətbiqi barədə",
      title_en: "On implementation of the amendment to the Law on Citizens' Appeals",
      type: "Fərman (Decree)",
      date: "10 June 2026",
      source: "president.az",
    },
  ],

  presidential_address: {
    title_en: "Address on the occasion of World Environment Day 2026",
    date: "5 June 2026",
    note: "Delivered as Azerbaijan hosts the World Environment Day celebrations. The President extended congratulations to all on the occasion.",
    source: "president.az",
  },

  cabinet_resolutions: [
    {
      number: "Q-175",
      date: "10 June 2026",
      title_az: "Azərbaycan Respublikası İqtisadiyyat Nazirliyinin strukturuna daxil olmayan tabeliyindəki qurumların siyahısında dəyişiklik barədə",
      title_en: "Amendment to the list of subordinate institutions not within the structure of the Ministry of Economy",
    },
    {
      number: "Q-174",
      date: "10 June 2026",
      title_az: "Azərbaycan Respublikasının Dövlət Neft Şirkəti və \"Gran Tierra Energy (Azerbaijan) GmbH\" şirkəti arasında neft kəşfiyyatı və hasilatı müqaviləsinə dair",
      title_en: "On the oil exploration and production agreement between SOCAR and Gran Tierra Energy (Azerbaijan) GmbH",
    },
    {
      number: "Q-173",
      date: "9 June 2026",
      title_az: "\"Qaz təchizatı haqqında\" 2025-ci il 8 iyul tarixli 233-VIIQ nömrəli Qanununun icrasının təmin edilməsi haqqında",
      title_en: "On ensuring implementation of the Gas Supply Law No. 233-VIIQ of 8 July 2025",
    },
    {
      number: "Q-172",
      date: "10 June 2026",
      title_az: "Xidməti parklanma yerləri təşkil olunan dövlət orqanlarının (qurumlarının) siyahısının təsdiq edilməsi haqqında",
      title_en: "On approval of the list of state bodies (institutions) for which official parking areas shall be established",
    },
    {
      number: "Q-171",
      date: "9 June 2026",
      title_az: "Ümumi istifadədə olan dəmiryol nəqliyyatının infrastrukturunun saxlanması xərclərinin maliyyələşdirilməsi haqqında",
      title_en: "On financing the maintenance costs of the public railway transport infrastructure",
    },
    {
      number: "Q-170",
      date: "9 June 2026",
      title_az: "İcbari tibbi sığorta üzrə Xidmətlər Zərfinin dəyişdirilməsi barədə",
      title_en: "Amendment to the Benefits Package under Mandatory Health Insurance",
    },
    {
      number: "Q-169",
      date: "9 June 2026",
      title_az: "Azərbaycan Respublikası Dövlət Təhlükəsizliyi Xidmətinin Mərkəzləşdirilmiş İnformasiya Sisteminə dair Əsasnamənin təsdiq edilməsi haqqında",
      title_en: "On approval of the Regulation on the Centralised Information System of the State Security Service",
    },
    {
      number: "Q-168",
      date: "9 June 2026",
      title_az: "\"Enerji effektivliyi informasiya sisteminin Əsasnaməsi\"nin təsdiq edilməsi barədə",
      title_en: "On approval of the Regulation on the Energy Efficiency Information System",
    },
  ],

  cabinet_orders: [
    {
      number: "S-403",
      date: "10 June 2026",
      title_az: "Azərbaycan Respublikasının Prezidenti yanında Vətəndaşlara Xidmət və Sosial İnnovasiyalar üzrə Dövlət Agentliyinin fəaliyyətinin təmin edilməsi haqqında Prezident Fərmanının (678, 2 iyun 2026) icrasının təmin edilməsi haqqında",
      title_en: "On implementation of Presidential Decree No. 678 of 2 June 2026 on the State Agency for Public Service and Social Innovations under the President",
    },
    {
      number: "S-402",
      date: "10 June 2026",
      title_az: "Dövlət mülkiyyətində olan hüquqi şəxslərin nəzarət şuralarında dövlət maraqlarının təmsil edilməsi haqqında",
      title_en: "On the representation of state interests in the supervisory boards of state-owned legal entities",
    },
  ],

  bills_june12: [
    { title_az: "Azərbaycan Respublikasının 2025-ci il dövlət büdcəsinin icrası haqqında", title_en: "On the execution of the 2025 State Budget of the Republic of Azerbaijan" },
    { title_az: "\"Gəmilərdə zərərli çirklənmə əleyhinə sistemlərə nəzarət haqqında\" Beynəlxalq Konvensiyaya qoşulmaq barədə", title_en: "On accession to the International Convention on the Control of Harmful Anti-fouling Systems on Ships" },
    { title_az: "Azərbaycan Respublikasının Portuqaliya Respublikasında (Lissabon şəhərində) Səfirliyinin təsis edilməsi haqqında", title_en: "On establishing the Embassy of the Republic of Azerbaijan in Portugal (Lisbon)" },
    { title_az: "Torpaq Məcəlləsi və əlaqəli qanunlarda dəyişikliklər barədə", title_en: "On amendments to the Land Code and related legislation" },
    { title_az: "\"Azərbaycanın mədəniyyət paytaxtı – Şuşa şəhəri haqqında\" Qanununda dəyişiklik", title_en: "Amendment to the Law on Shusha as the Cultural Capital of Azerbaijan" },
    { title_az: "\"Yaşıllıqların mühafizəsi haqqında\" Qanununda dəyişiklik", title_en: "Amendment to the Law on Protection of Green Spaces" },
    { title_az: "İnzibati Xətalar Məcəlləsində dəyişiklik", title_en: "Amendment to the Administrative Violations Code" },
    { title_az: "\"Dövlət qulluğu haqqında\" Qanununda dəyişiklik", title_en: "Amendment to the Civil Service Law" },
    { title_az: "İnzibati Xətalar Məcəlləsi və informasiya qanunlarında dəyişikliklər", title_en: "Amendments to the Administrative Violations Code and information legislation" },
    { title_az: "Mənzil Məcəlləsi və məcburi köçkünlərə dair qanunlarda dəyişikliklər", title_en: "Amendments to the Housing Code and legislation on internally displaced persons" },
  ],

  bills_june5: [
    { title_az: "\"Diplomatik xidmət haqqında\" Qanununda dəyişiklik", title_en: "Amendment to the Law on Diplomatic Service" },
    { title_az: "\"Korrupsiyaya qarşı mübarizə haqqında\" Qanununda dəyişiklik", title_en: "Amendment to the Anti-Corruption Law" },
    { title_az: "Cinayət-Prosessual Məcəlləsində dəyişiklik", title_en: "Amendment to the Criminal Procedure Code" },
    { title_az: "Cinayət-Prosessual Məcəllə, Cəzaların İcrası Məcəllə və İcra Qanununda dəyişikliklər", title_en: "Amendments to the Criminal Procedure Code, Penal Enforcement Code, and Enforcement Law" },
    { title_az: "Gömrük Məcəlləsində dəyişiklik", title_en: "Amendment to the Customs Code" },
  ],
};

// ── Helpers ───────────────────────────────────────────────────────────────────

function hline(doc, y, { color = RULE_GRAY, width = 0.5, mx = 50 } = {}) {
  doc
    .save()
    .moveTo(mx, y)
    .lineTo(doc.page.width - mx, y)
    .lineWidth(width)
    .strokeColor(color)
    .stroke()
    .restore();
}

function thick_hline(doc, y, { color = NAVY, width = 3, mx = 50 } = {}) {
  hline(doc, y, { color, width, mx });
}

function section_header(doc, title, y) {
  const mx = 50;
  const w  = doc.page.width - mx * 2;
  doc
    .rect(mx, y, w, 22)
    .fill(NAVY);
  doc
    .fontSize(9)
    .font("Helvetica-Bold")
    .fillColor(WHITE)
    .text(title.toUpperCase(), mx + 8, y + 6, { width: w - 16, lineBreak: false });
  return y + 22;
}

function pill(doc, label, x, y) {
  const pad = 5;
  const tw  = doc.widthOfString(label) + pad * 2;
  doc
    .rect(x, y - 1, tw, 13)
    .fill(CRIMSON);
  doc
    .fontSize(6.5)
    .font("Helvetica-Bold")
    .fillColor(WHITE)
    .text(label, x + pad, y + 1, { lineBreak: false });
  return x + tw + 6;
}

// ── Build PDF ─────────────────────────────────────────────────────────────────

const doc = new PDFDocument({
  size: "A4",
  margins: { top: 50, bottom: 50, left: 50, right: 50 },
  info: {
    Title:    `Omni Law Gazette — Issue No. ${ISSUE_NUMBER}`,
    Author:   "Omni Law Firm · Legal Desk",
    Subject:  "Weekly Digest of Legislative Developments, Republic of Azerbaijan",
    Keywords: "Azerbaijan, law, gazette, decree, legislation",
    Creator:  "Omni Law Gazette Generator",
  },
  bufferPages: true,
});

const stream = fs.createWriteStream(OUT);
doc.pipe(stream);

const PAGE_W = doc.page.width;
const PAGE_H = doc.page.height;
const MX     = 50;  // margin x
const CW     = PAGE_W - MX * 2;  // content width

let y = 50;

// ════════════════════════════════════════════════════════════════════════════
// MASTHEAD
// ════════════════════════════════════════════════════════════════════════════

// Top triple rule
thick_hline(doc, y, { color: CRIMSON, width: 1 });
y += 3;
thick_hline(doc, y, { color: NAVY,    width: 6 });
y += 8;
thick_hline(doc, y, { color: CRIMSON, width: 1 });
y += 10;

// Eyebrow text
doc
  .fontSize(7.5)
  .font("Helvetica")
  .fillColor(SILVER)
  .text(`OMNI LAW FIRM  ·  INTERNAL PUBLICATION  ·  CONFIDENTIAL`, MX, y, { width: CW, align: "center" });
y += 14;

// Title
doc
  .fontSize(42)
  .font("Helvetica-Bold")
  .fillColor(NAVY)
  .text(GAZETTE_TITLE, MX, y, { width: CW, align: "center", lineBreak: false });
y += 52;

// Sub-title
doc
  .fontSize(9)
  .font("Helvetica-Oblique")
  .fillColor(SILVER)
  .text(GAZETTE_SUB, MX, y, { width: CW, align: "center" });
y += 16;

// Double rule
hline(doc, y, { color: NAVY, width: 2 });
y += 4;
hline(doc, y, { color: NAVY, width: 0.5 });
y += 10;

// Meta bar  (issue · week · published)
const META_COL = CW / 3;
doc
  .fontSize(7.5)
  .font("Helvetica-Bold")
  .fillColor(NAVY)
  .text(`ISSUE NO. ${ISSUE_NUMBER}`, MX, y, { width: META_COL, align: "left", lineBreak: false });
doc
  .fontSize(7.5)
  .font("Helvetica")
  .fillColor(TEXT_GRAY)
  .text(`Week of ${WEEK_START} – ${WEEK_END}`, MX + META_COL, y, { width: META_COL, align: "center", lineBreak: false });
doc
  .fontSize(7.5)
  .font("Helvetica")
  .fillColor(TEXT_GRAY)
  .text(`Published ${PUBLISHED}`, MX + META_COL * 2, y, { width: META_COL, align: "right", lineBreak: false });
y += 14;

hline(doc, y, { color: NAVY, width: 2 });
y += 16;

// ════════════════════════════════════════════════════════════════════════════
// INTRODUCTION / SUMMARY BOX
// ════════════════════════════════════════════════════════════════════════════

const BOX_H = 60;
doc.rect(MX, y, CW, BOX_H).fill(LIGHT_GRAY);
hline(doc, y,        { color: GOLD, width: 2, mx: MX });
hline(doc, y+BOX_H, { color: GOLD, width: 2, mx: MX });

doc
  .fontSize(8)
  .font("Helvetica-Bold")
  .fillColor(NAVY)
  .text("EDITOR'S NOTE", MX + 10, y + 8);
doc
  .fontSize(8)
  .font("Helvetica")
  .fillColor(TEXT_GRAY)
  .text(
    "This issue covers the period 5–12 June 2026 and was compiled from official sources: president.az, nk.gov.az (Cabinet of Ministers), and meclis.gov.az (Milli Majlis). " +
    "This week saw four Presidential Decrees — including landmark measures on business credit subsidies and non-oil export stimulation — " +
    "alongside ten Cabinet resolutions and orders. Fifteen new bills are under Milli Majlis discussion, including the 2025 Budget Execution Report " +
    "and the establishment of an Embassy in Lisbon. The Constitutional Court website was unavailable (HTTP 503) at time of compilation.",
    MX + 10, y + 20,
    { width: CW - 20, lineGap: 1.5 }
  );
y += BOX_H + 16;

// ════════════════════════════════════════════════════════════════════════════
// SECTION 1 — PRESIDENTIAL DECREES
// ════════════════════════════════════════════════════════════════════════════

y = section_header(doc, "I.  Presidential Decrees  (Prezidentin Fərmanları)", y);
y += 10;

data.presidential_decrees.forEach((item, i) => {
  if (y > PAGE_H - 120) { doc.addPage(); y = 50; }

  // Serial number pill + date
  let xCur = MX;
  xCur = pill(doc, `FƏRMAN ${i + 1}`, xCur, y);
  doc
    .fontSize(7.5)
    .font("Helvetica")
    .fillColor(SILVER)
    .text(`${item.date}  ·  Source: ${item.source}`, xCur, y + 1, { lineBreak: false });
  y += 16;

  // English title (bold)
  doc
    .fontSize(9.5)
    .font("Helvetica-Bold")
    .fillColor(BLACK)
    .text(item.title_en, MX + 10, y, { width: CW - 10 });
  y = doc.y + 2;

  // Azerbaijani title (italic)
  doc
    .fontSize(8)
    .font("Helvetica-Oblique")
    .fillColor(SILVER)
    .text(item.title_az, MX + 10, y, { width: CW - 10 });
  y = doc.y + 8;

  if (i < data.presidential_decrees.length - 1) {
    hline(doc, y, { color: RULE_GRAY, width: 0.3, mx: MX + 10 });
    y += 8;
  }
});

y += 4;
hline(doc, y, { color: RULE_GRAY, width: 0.5 });
y += 12;

// ════════════════════════════════════════════════════════════════════════════
// SECTION 1b — PRESIDENTIAL ADDRESS
// ════════════════════════════════════════════════════════════════════════════

y = section_header(doc, "II.  Presidential Address", y);
y += 10;

if (y > PAGE_H - 80) { doc.addPage(); y = 50; }

let xCur2 = MX;
xCur2 = pill(doc, "ADDRESS", xCur2, y);
doc.fontSize(7.5).font("Helvetica").fillColor(SILVER)
   .text(`${data.presidential_address.date}  ·  Source: ${data.presidential_address.source}`, xCur2, y + 1, { lineBreak: false });
y += 16;

doc.fontSize(9.5).font("Helvetica-Bold").fillColor(BLACK)
   .text(data.presidential_address.title_en, MX + 10, y, { width: CW - 10 });
y = doc.y + 4;

doc.fontSize(8).font("Helvetica").fillColor(TEXT_GRAY)
   .text(data.presidential_address.note, MX + 10, y, { width: CW - 10 });
y = doc.y + 12;

hline(doc, y, { color: RULE_GRAY, width: 0.5 });
y += 12;

// ════════════════════════════════════════════════════════════════════════════
// SECTION 2 — CABINET RESOLUTIONS
// ════════════════════════════════════════════════════════════════════════════

if (y > PAGE_H - 140) { doc.addPage(); y = 50; }

y = section_header(doc, "III.  Cabinet of Ministers — Resolutions  (Nazirlər Kabineti Qərarları)", y);
y += 10;

data.cabinet_resolutions.forEach((item, i) => {
  if (y > PAGE_H - 100) { doc.addPage(); y = 50; }

  let xCur = MX;
  xCur = pill(doc, item.number, xCur, y);
  doc.fontSize(7.5).font("Helvetica").fillColor(SILVER)
     .text(item.date, xCur, y + 1, { lineBreak: false });
  y += 16;

  doc.fontSize(9.5).font("Helvetica-Bold").fillColor(BLACK)
     .text(item.title_en, MX + 10, y, { width: CW - 10 });
  y = doc.y + 2;

  doc.fontSize(8).font("Helvetica-Oblique").fillColor(SILVER)
     .text(item.title_az, MX + 10, y, { width: CW - 10 });
  y = doc.y + 8;

  if (i < data.cabinet_resolutions.length - 1) {
    hline(doc, y, { color: RULE_GRAY, width: 0.3, mx: MX + 10 });
    y += 8;
  }
});

y += 4;
hline(doc, y, { color: RULE_GRAY, width: 0.5 });
y += 12;

// ════════════════════════════════════════════════════════════════════════════
// SECTION 3 — CABINET ORDERS
// ════════════════════════════════════════════════════════════════════════════

if (y > PAGE_H - 100) { doc.addPage(); y = 50; }

y = section_header(doc, "IV.  Cabinet of Ministers — Orders  (Nazirlər Kabineti Sərəncamları)", y);
y += 10;

data.cabinet_orders.forEach((item, i) => {
  if (y > PAGE_H - 100) { doc.addPage(); y = 50; }

  let xCur = MX;
  xCur = pill(doc, item.number, xCur, y);
  doc.fontSize(7.5).font("Helvetica").fillColor(SILVER)
     .text(item.date, xCur, y + 1, { lineBreak: false });
  y += 16;

  doc.fontSize(9.5).font("Helvetica-Bold").fillColor(BLACK)
     .text(item.title_en, MX + 10, y, { width: CW - 10 });
  y = doc.y + 2;

  doc.fontSize(8).font("Helvetica-Oblique").fillColor(SILVER)
     .text(item.title_az, MX + 10, y, { width: CW - 10 });
  y = doc.y + 8;

  if (i < data.cabinet_orders.length - 1) {
    hline(doc, y, { color: RULE_GRAY, width: 0.3, mx: MX + 10 });
    y += 8;
  }
});

y += 4;
hline(doc, y, { color: RULE_GRAY, width: 0.5 });
y += 12;

// ════════════════════════════════════════════════════════════════════════════
// SECTION 4 — MILLI MAJLIS BILLS
// ════════════════════════════════════════════════════════════════════════════

if (y > PAGE_H - 140) { doc.addPage(); y = 50; }

y = section_header(doc, "V.  Milli Majlis — Bills Under Discussion  (Qanun Layihələri)", y);
y += 10;

// Sub-header: June 12
if (y > PAGE_H - 60) { doc.addPage(); y = 50; }
doc.rect(MX, y, CW, 16).fill("#E8EEF7");
doc.fontSize(7.5).font("Helvetica-Bold").fillColor(NAVY)
   .text("SUBMITTED  12 JUNE 2026", MX + 8, y + 4, { lineBreak: false });
y += 22;

data.bills_june12.forEach((item, i) => {
  if (y > PAGE_H - 80) { doc.addPage(); y = 50; }

  let xCur = MX;
  pill(doc, `BILL ${i + 1}`, xCur, y);
  y += 14;

  doc.fontSize(9).font("Helvetica-Bold").fillColor(BLACK)
     .text(item.title_en, MX + 10, y, { width: CW - 10 });
  y = doc.y + 2;

  doc.fontSize(7.5).font("Helvetica-Oblique").fillColor(SILVER)
     .text(item.title_az, MX + 10, y, { width: CW - 10 });
  y = doc.y + 7;

  if (i < data.bills_june12.length - 1) {
    hline(doc, y, { color: RULE_GRAY, width: 0.3, mx: MX + 10 });
    y += 7;
  }
});

y += 8;

// Sub-header: June 5
if (y > PAGE_H - 60) { doc.addPage(); y = 50; }
doc.rect(MX, y, CW, 16).fill("#E8EEF7");
doc.fontSize(7.5).font("Helvetica-Bold").fillColor(NAVY)
   .text("SUBMITTED  5 JUNE 2026", MX + 8, y + 4, { lineBreak: false });
y += 22;

data.bills_june5.forEach((item, i) => {
  if (y > PAGE_H - 80) { doc.addPage(); y = 50; }

  let xCur = MX;
  pill(doc, `BILL ${data.bills_june12.length + i + 1}`, xCur, y);
  y += 14;

  doc.fontSize(9).font("Helvetica-Bold").fillColor(BLACK)
     .text(item.title_en, MX + 10, y, { width: CW - 10 });
  y = doc.y + 2;

  doc.fontSize(7.5).font("Helvetica-Oblique").fillColor(SILVER)
     .text(item.title_az, MX + 10, y, { width: CW - 10 });
  y = doc.y + 7;

  if (i < data.bills_june5.length - 1) {
    hline(doc, y, { color: RULE_GRAY, width: 0.3, mx: MX + 10 });
    y += 7;
  }
});

y += 8;
hline(doc, y, { color: RULE_GRAY, width: 0.5 });
y += 12;

// ════════════════════════════════════════════════════════════════════════════
// SECTION 5 — NOTICE: Constitutional Court
// ════════════════════════════════════════════════════════════════════════════

if (y > PAGE_H - 80) { doc.addPage(); y = 50; }

y = section_header(doc, "VI.  Constitutional Court  (Konstitusiya Məhkəməsi)", y);
y += 10;

doc
  .rect(MX, y, CW, 36)
  .fill("#FFF8E1");
doc
  .rect(MX, y, 3, 36)
  .fill(GOLD);
doc
  .fontSize(8)
  .font("Helvetica-Bold")
  .fillColor("#7A5C00")
  .text("NOTICE — Source Unavailable", MX + 10, y + 6);
doc
  .fontSize(8)
  .font("Helvetica")
  .fillColor(TEXT_GRAY)
  .text(
    "The Constitutional Court website (constcourt.gov.az) returned HTTP 503 at time of compilation. " +
    "No decisions could be retrieved for this week. Readers should check the website directly for any rulings issued between 5–12 June 2026.",
    MX + 10, y + 18,
    { width: CW - 20 }
  );
y += 50;

hline(doc, y, { color: RULE_GRAY, width: 0.5 });
y += 14;

// ════════════════════════════════════════════════════════════════════════════
// SOURCE TABLE
// ════════════════════════════════════════════════════════════════════════════

if (y > PAGE_H - 120) { doc.addPage(); y = 50; }

doc.fontSize(8).font("Helvetica-Bold").fillColor(NAVY)
   .text("SOURCES", MX, y);
y += 12;

const sources = [
  ["Office of the President", "president.az/en/documents", "Decrees, Orders, Addresses"],
  ["Cabinet of Ministers", "nk.gov.az/az/senedler/hamisi", "Resolutions, Orders"],
  ["Milli Majlis — Laws", "meclis.gov.az/cat-qanun.php?cat=72&lang=az", "Enacted laws"],
  ["Milli Majlis — Decrees", "meclis.gov.az/cat-qanun.php?cat=138&lang=az", "Parliament resolutions"],
  ["Milli Majlis — Bills", "meclis.gov.az/cat-layihe.php?cat=83&lang=az", "Draft legislation"],
  ["Milli Majlis — Reports", "meclis.gov.az/cat-cari.php?cat=89&lang=az", "Committee reports"],
  ["E-Qanun", "e-qanun.az", "Consolidated legislation (partial load)"],
  ["Constitutional Court", "constcourt.gov.az/az/decisions", "UNAVAILABLE this issue (HTTP 503)"],
];

const COL = [160, 200, CW - 360];
const ROW_H = 14;

// Table header
doc.rect(MX, y, CW, ROW_H).fill(NAVY);
["Institution", "URL", "Coverage"].forEach((h, i) => {
  const xOff = i === 0 ? 0 : COL.slice(0, i).reduce((a, b) => a + b, 0);
  doc.fontSize(7).font("Helvetica-Bold").fillColor(WHITE)
     .text(h, MX + xOff + 4, y + 3, { width: COL[i] - 8, lineBreak: false });
});
y += ROW_H;

sources.forEach((row, ri) => {
  const bg = ri % 2 === 0 ? WHITE : LIGHT_GRAY;
  doc.rect(MX, y, CW, ROW_H).fill(bg);
  row.forEach((cell, ci) => {
    const xOff = ci === 0 ? 0 : COL.slice(0, ci).reduce((a, b) => a + b, 0);
    doc.fontSize(7).font("Helvetica").fillColor(ci === 1 ? "#1A56A0" : TEXT_GRAY)
       .text(cell, MX + xOff + 4, y + 3, { width: COL[ci] - 8, lineBreak: false });
  });
  y += ROW_H;
});

y += 14;

// ════════════════════════════════════════════════════════════════════════════
// FOOTER  (on every page via bufferPages)
// ════════════════════════════════════════════════════════════════════════════

const totalPages = doc.bufferedPageRange().count;
for (let p = 0; p < totalPages; p++) {
  doc.switchToPage(p);
  const fy = PAGE_H - 38;
  hline(doc, fy, { color: NAVY, width: 1 });
  doc.fontSize(6.5).font("Helvetica").fillColor(SILVER)
     .text(
       `Omni Law Gazette  ·  Issue No. ${ISSUE_NUMBER}  ·  ${PUBLISHED}  ·  Compiled by: Omni Law Firm Legal Desk  ·  CONFIDENTIAL — FOR INTERNAL USE ONLY`,
       MX, fy + 6, { width: CW - 60, align: "left" }
     );
  doc.fontSize(6.5).font("Helvetica").fillColor(SILVER)
     .text(`${p + 1} / ${totalPages}`, PAGE_W - 80, fy + 6, { width: 30, align: "right" });
}

// ════════════════════════════════════════════════════════════════════════════

doc.end();

stream.on("finish", () => {
  const size = fs.statSync(OUT).size;
  console.log(`✓  PDF written to: ${OUT}`);
  console.log(`   Size: ${(size / 1024).toFixed(1)} KB`);
});

stream.on("error", (err) => {
  console.error("Stream error:", err);
  process.exit(1);
});
