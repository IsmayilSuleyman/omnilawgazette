#!/usr/bin/env node
/**
 * Omni Law Gazette — Publish Script
 *
 * Converts a gazette HTML file to PDF using Playwright, then uploads it to
 * Supabase storage and inserts an issues record so it appears in the web app.
 *
 * Usage:
 *   node scripts/publish-gazette.mjs --html scripts/gazettes/gazette-2026-06-27.html \
 *     --issue 2 --date 2026-06-27 \
 *     --title "Intellectual Customs, Health Insurance Reform & Gas Sector Reset" \
 *     --tags "customs,health insurance,gas,SME,decrees" \
 *     --summary "..."
 */

import { readFileSync, mkdirSync } from 'fs';
import { join, dirname, basename } from 'path';
import { fileURLToPath } from 'url';
import { parseArgs } from 'util';
import { createClient } from '@supabase/supabase-js';
import { chromium } from 'playwright';

const __dirname = dirname(fileURLToPath(import.meta.url));

// ── CLI args ──────────────────────────────────────────────────────────────────
const { values } = parseArgs({
  options: {
    html:    { type: 'string' },
    issue:   { type: 'string' },
    date:    { type: 'string' },
    title:   { type: 'string' },
    summary: { type: 'string' },
    tags:    { type: 'string' },
    out:     { type: 'string' },
    'dry-run': { type: 'boolean' },
  },
  strict: false,
});

const htmlPath  = values.html ?? join(__dirname, 'gazettes', 'gazette-2026-06-27.html');
const issueNum  = Number(values.issue ?? 2);
const pubDate   = values.date ?? '2026-06-27';
const title     = values.title ?? 'Intellectual Customs, Health Insurance Reform & Gas Sector Reset';
const summary   = values.summary ?? 'A landmark legislative week for Azerbaijan: President Aliyev issues six decrees establishing the new State Medical Insurance Agency, launching the Intellectual Customs digital system, and ratifying an international workplace accident protocol. The Cabinet of Ministers delivers ten decisions and four orders spanning gas supply reform, horizontal tax monitoring, SME support, NGO funding, and tourism investment. One new law enacted: accession to the IMO AFS Convention on ship coatings.';
const rawTags   = values.tags ?? 'customs,health insurance,gas supply,SME,tax monitoring,decrees,orders';
const tags      = [...new Set(rawTags.split(',').map(t => t.trim()).filter(Boolean))].slice(0, 10);
const outDir    = values.out ?? join(__dirname, 'gazettes');
const dryRun    = values['dry-run'] ?? false;

const SUPABASE_URL = process.env.NEXT_PUBLIC_SUPABASE_URL ?? 'https://wqgjqljupslpzklmrwqm.supabase.co';
const SUPABASE_KEY = process.env.SUPABASE_SERVICE_KEY ?? process.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY ?? 'sb_publishable_eglpkiylAZ5O1d-YZSzDcQ_dShoQtws';
const BUCKET = 'gazette';

console.log('\n📰  Omni Law Gazette — Publish Script');
console.log(`    HTML   : ${htmlPath}`);
console.log(`    Issue  : No. ${issueNum}`);
console.log(`    Date   : ${pubDate}`);
console.log(`    Title  : ${title}`);
console.log(`    Tags   : ${tags.join(', ')}`);
if (dryRun) console.log('    MODE   : DRY RUN (no upload)');

// ── Step 1: HTML → PDF via Playwright ────────────────────────────────────────
mkdirSync(outDir, { recursive: true });
const pdfLocalPath = join(outDir, `gazette-${pubDate}.pdf`);

console.log('\n[1/3] Rendering HTML → PDF …');
const htmlContent = readFileSync(htmlPath, 'utf8');

const browser = await chromium.launch({
  executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
  args: ['--no-sandbox', '--disable-setuid-sandbox'],
});
const page = await browser.newPage();
await page.setContent(htmlContent, { waitUntil: 'networkidle' });

const pdfBuffer = await page.pdf({
  format: 'A4',
  printBackground: true,
  margin: { top: '10mm', right: '10mm', bottom: '10mm', left: '10mm' },
  displayHeaderFooter: false,
});
await browser.close();

import { writeFileSync } from 'fs';
writeFileSync(pdfLocalPath, pdfBuffer);
console.log(`    Saved  : ${pdfLocalPath} (${Math.round(pdfBuffer.length / 1024)} KB)`);

if (dryRun) {
  console.log('\n✅  Dry run complete — PDF generated, Supabase upload skipped.');
  process.exit(0);
}

// ── Step 2: Upload PDF to Supabase Storage ────────────────────────────────────
console.log('\n[2/3] Uploading PDF to Supabase Storage …');
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY, {
  auth: { persistSession: false, autoRefreshToken: false },
});

const rand = Math.random().toString(36).slice(2, 10);
const pdfPath = `pdfs/issue-${issueNum}-${rand}.pdf`;

const { error: uploadErr } = await supabase.storage
  .from(BUCKET)
  .upload(pdfPath, pdfBuffer, {
    contentType: 'application/pdf',
    cacheControl: '31536000',
    upsert: false,
  });

if (uploadErr) {
  console.error(`    ✗ Upload failed: ${uploadErr.message}`);
  process.exit(1);
}
console.log(`    Stored : ${pdfPath}`);

// ── Step 3: Insert issue record ────────────────────────────────────────────────
console.log('\n[3/3] Inserting issue record in database …');

// Rough page count estimation: PDF is generated so let's use a placeholder.
// The actual page count can be inferred from the rendered PDF size.
const estimatedPages = Math.max(1, Math.round(pdfBuffer.length / 40000));

const { error: insertErr } = await supabase.from('issues').insert({
  issue_number:    issueNum,
  title:           title,
  summary:         summary,
  published_at:    pubDate,
  tags:            tags,
  pdf_path:        pdfPath,
  cover_path:      null,
  page_count:      estimatedPages,
  file_size_bytes: pdfBuffer.length,
});

if (insertErr) {
  // Roll back PDF upload on failure
  await supabase.storage.from(BUCKET).remove([pdfPath]).catch(() => {});
  console.error(`    ✗ Insert failed: ${insertErr.message}`);
  if (insertErr.code === '23505') {
    console.error(`    Issue No. ${issueNum} already exists — use a different number.`);
  }
  process.exit(1);
}

console.log(`    Record : issue_number=${issueNum} inserted`);
console.log('\n✅  Issue No.', issueNum, 'is live!');
console.log(`    View at: /issues/${issueNum}`);
