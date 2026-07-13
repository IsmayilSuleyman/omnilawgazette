#!/usr/bin/env node
/**
 * Omni Law Gazette — Weekly Scraper
 *
 * Fetches legislative updates from Azerbaijani government portals and
 * writes a dated HTML gazette file to scripts/gazettes/.
 *
 * Usage:
 *   node scripts/scrape-gazette.mjs
 *   node scripts/scrape-gazette.mjs --from 2026-06-06 --to 2026-06-13
 *
 * Sources (Tier 1):
 *   1. e-qanun.az              — normative acts register (hidden JSON API; see EQANUN-EXTRACTION.md)
 *   2. president.az/en/documents — presidential decrees, orders, letters
 *   3. nk.gov.az/az/senedler  — Cabinet of Ministers
 *   4. constcourt.gov.az/az/decisions — Constitutional Court
 *   5. meclis.gov.az           — Milli Majlis (laws, decrees, reports, bills)
 */

import { writeFileSync, mkdirSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';
import { parseArgs } from 'util';

const __dirname = dirname(fileURLToPath(import.meta.url));

// ── CLI args ──────────────────────────────────────────────────────────────────
const { values } = parseArgs({
  options: {
    from: { type: 'string' },
    to:   { type: 'string' },
    out:  { type: 'string' },
  },
  strict: false,
});

const toDate   = values.to   ? new Date(values.to)   : new Date();
const fromDate = values.from ? new Date(values.from)  : new Date(toDate.getTime() - 7 * 86400_000);
const outDir   = values.out  ?? join(__dirname, 'gazettes');
mkdirSync(outDir, { recursive: true });

const fmt = (d) => d.toISOString().slice(0, 10);
const fmtDisplay = (d) => d.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });

console.log(`\n📰 Omni Law Gazette Scraper`);
console.log(`   Period : ${fmt(fromDate)} → ${fmt(toDate)}`);

// ── HTTP fetch helper ─────────────────────────────────────────────────────────
async function fetchPage(url, label) {
  console.log(`   Fetching ${label} …`);
  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 15_000);
    const res = await fetch(url, {
      signal: controller.signal,
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; OmniLawGazette/1.0)',
        'Accept-Language': 'az,en;q=0.9',
        'Accept': 'text/html,application/xhtml+xml',
      },
    });
    clearTimeout(timeout);
    if (!res.ok) return { ok: false, status: res.status, text: '' };
    return { ok: true, status: res.status, text: await res.text() };
  } catch (e) {
    return { ok: false, status: 0, error: e.message, text: '' };
  }
}

// ── Simple HTML text extractor ────────────────────────────────────────────────
function stripTags(html) {
  return html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

// ── Date range filter helper ──────────────────────────────────────────────────
function inRange(dateStr) {
  const d = new Date(dateStr);
  return d >= fromDate && d <= toDate;
}

// ── Source scrapers ───────────────────────────────────────────────────────────

async function scrapePresidentLetters() {
  const { ok, text, status } = await fetchPage(
    'https://president.az/en/documents', 'president.az/letters'
  );
  if (!ok) return { ok: false, status, items: [] };

  // Extract list items with dates — simple regex approach for this static layout
  const items = [];
  const dateRe = /(\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4})/gi;
  const rows = text.split(/<li|<article|<div class="document/i).slice(1);
  for (const row of rows) {
    const plain = stripTags(row);
    const match = plain.match(dateRe);
    if (match && plain.length > 20) {
      items.push({ date: match[0], text: plain.slice(0, 200) });
    }
  }
  return { ok: true, status, items };
}

async function scrapeCabinet() {
  const { ok, text, status } = await fetchPage(
    'https://nk.gov.az/az/senedler/hamisi', 'nk.gov.az/Cabinet'
  );
  if (!ok) return { ok: false, status, items: [] };

  const items = [];
  // Cabinet page uses a table/list format — extract rows with document numbers and titles
  const rowRe = /(?:Qərar|Sərəncam)\s+(?:No\.|№)\s*(\d+\w*)/gi;
  const dateRe = /(\d{2}\.\d{2}\.\d{4})/g;
  let m;
  while ((m = rowRe.exec(text)) !== null) {
    const snippet = text.slice(Math.max(0, m.index - 50), m.index + 300);
    const dates = snippet.match(dateRe);
    items.push({
      type: text.slice(m.index, m.index + 7).trim(),
      number: m[1],
      date: dates?.[0] ?? '—',
      snippet: stripTags(snippet).slice(0, 180),
    });
  }
  return { ok: true, status, items };
}

async function scrapeMajlisBills() {
  const { ok, text, status } = await fetchPage(
    'https://meclis.gov.az/cat-layihe.php?cat=83&lang=az', 'meclis.gov.az/bills'
  );
  if (!ok) return { ok: false, status, items: [] };

  const rows = text.split(/<tr|<li|<div class="row/i).slice(1);
  const items = rows
    .map(r => stripTags(r).replace(/\s+/g, ' ').trim())
    .filter(r => r.length > 30)
    .slice(0, 30);
  return { ok: true, status, items };
}

async function scrapeMajlisLaws() {
  const { ok, text, status } = await fetchPage(
    'https://meclis.gov.az/cat-qanun.php?cat=72&lang=az', 'meclis.gov.az/laws'
  );
  if (!ok) return { ok: false, status, items: [] };

  const rows = text.split(/<tr|<li/i).slice(1);
  const items = rows
    .map(r => stripTags(r).replace(/\s+/g, ' ').trim())
    .filter(r => r.length > 30)
    .slice(0, 20);
  return { ok: true, status, items };
}

async function scrapeConstCourt() {
  const { ok, text, status } = await fetchPage(
    'https://www.constcourt.gov.az/az/decisions', 'constcourt.gov.az'
  );
  if (!ok) return { ok: false, status, items: [] };

  const rows = text.split(/<tr|<li|<article/i).slice(1);
  const items = rows
    .map(r => stripTags(r).replace(/\s+/g, ' ').trim())
    .filter(r => r.length > 20)
    .slice(0, 15);
  return { ok: true, status, items };
}

// ── e-Qanun: the official consolidated register via its hidden JSON API ─────────
//
// e-qanun.az is a Next.js SPA whose HTML is empty — this used to be logged as
// "data not extractable". It is NOT: the site is backed by a plain JSON REST
// API. Full recipe (hosts, params, gotchas) in scripts/EQANUN-EXTRACTION.md.
//
// Key facts, so nobody has to rediscover them:
//   • Host api.e-qanun.az  = Republic of Azerbaijan DB (Laws, Presidential
//     Fərman/Sərəncam, Cabinet Qərar/Sərəncam). napi.e-qanun.az = Nakhchivan.
//   • orderColumn MUST be 1 (0 → HTTP 500). Dates MUST be DD.MM.YYYY (ISO → 500).
//   • All the params below are required together or the server 500s.
async function scrapeEQanun() {
  const dmy = (d) => d.toLocaleDateString('en-GB').replaceAll('/', '.'); // DD.MM.YYYY
  const params = new URLSearchParams({
    start: '0', length: '500', orderColumn: '1', orderDirection: 'desc',
    title: 'true', codeType: '1', dateType: '1', statusId: '1', secondType: '2',
    array: '', specialDate: 'false',
    beginDate: dmy(fromDate), endDate: dmy(toDate), dateRange: 'true',
  });
  const url = `https://api.e-qanun.az/getDetailSearch?${params}`;
  console.log(`   Fetching e-qanun.az/Republic (JSON API) …`);
  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 30_000);
    const res = await fetch(url, {
      signal: controller.signal,
      headers: { 'Accept-Language': 'AZ', 'Referer': 'https://e-qanun.az/' },
    });
    clearTimeout(timeout);
    if (!res.ok) return { ok: false, status: res.status, items: [] };
    const json = await res.json();
    const items = (json.data ?? []).map((r) => ({
      id: r.id,
      number: r.citation || '—',
      type: r.typeName,
      status: r.statusName,
      date: r.acceptDate,           // DD.MM.YYYY
      title: r.title,
      url: `https://e-qanun.az/framework/${r.id}`,
    }));
    return { ok: true, status: res.status, items };
  } catch (e) {
    return { ok: false, status: 0, error: e.message, items: [] };
  }
}

// ── Main ──────────────────────────────────────────────────────────────────────
async function main() {
  const [letters, cabinet, bills, laws, court, eqanun] = await Promise.all([
    scrapePresidentLetters(),
    scrapeCabinet(),
    scrapeMajlisBills(),
    scrapeMajlisLaws(),
    scrapeConstCourt(),
    scrapeEQanun(),
  ]);

  const results = { letters, cabinet, bills, laws, court, eqanun };

  // Write raw JSON for debugging
  const jsonPath = join(outDir, `scrape-${fmt(toDate)}.json`);
  writeFileSync(jsonPath, JSON.stringify(results, null, 2));
  console.log(`\n✅ Raw scrape saved: ${jsonPath}`);
  console.log(`\nSummary:`);
  console.log(`   President letters : ${letters.ok ? letters.items.length : `ERROR ${letters.status}`}`);
  console.log(`   Cabinet items     : ${cabinet.ok ? cabinet.items.length : `ERROR ${cabinet.status}`}`);
  console.log(`   Bills             : ${bills.ok   ? bills.items.length   : `ERROR ${bills.status}`}`);
  console.log(`   Laws              : ${laws.ok    ? laws.items.length    : `ERROR ${laws.status}`}`);
  console.log(`   Const. Court      : ${court.ok   ? court.items.length   : `ERROR ${court.status}`}`);
  console.log(`   e-Qanun (Republic): ${eqanun.ok  ? eqanun.items.length  : `ERROR ${eqanun.status}`}`);
  console.log(`\n💡 Use the scraped JSON to update the HTML gazette template in scripts/gazettes/`);
  console.log(`   Or run with --from YYYY-MM-DD --to YYYY-MM-DD to customise the date range.\n`);
}

main().catch(e => { console.error(e); process.exit(1); });
