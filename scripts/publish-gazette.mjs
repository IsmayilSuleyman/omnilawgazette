/**
 * Uploads gazette-issue.pdf to Supabase storage and inserts the issue record.
 * Run: node scripts/publish-gazette.mjs
 */

import { createClient } from "@supabase/supabase-js";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import crypto from "crypto";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const SUPABASE_URL = "https://wqgjqljupslpzklmrwqm.supabase.co";
const SUPABASE_KEY = "sb_publishable_eglpkiylAZ5O1d-YZSzDcQ_dShoQtws";
const BUCKET = "gazette";

const PDF_PATH = path.join(__dirname, "../gazette-issue.pdf");

const ISSUE = {
  issue_number: 2,
  title: "Baku Restructuring, Business Credit Reform & Non-Oil Export Drive",
  summary:
    "Four Presidential Decrees this week — highlighted by structural reform of the Baku City Executive Authority, new credit-guarantee and interest-subsidy mechanisms for businesses, and fresh incentives for non-oil export and transport costs. The Cabinet of Ministers issued eight resolutions and two orders, covering SOCAR's new oil deal with Gran Tierra Energy, gas-supply law implementation, railway infrastructure financing, mandatory health insurance amendments, and a new State Security Service IT regulation. Milli Majlis has 15 bills under discussion, including the 2025 Budget Execution Report, an Embassy establishment in Lisbon, and Land Code amendments. Constitutional Court unavailable (HTTP 503) this week.",
  published_at: "2026-06-12",
  tags: ["presidential-decree", "cabinet", "business", "energy", "parliament", "export"],
};

async function main() {
  const supabase = createClient(SUPABASE_URL, SUPABASE_KEY, {
    auth: { persistSession: false, autoRefreshToken: false, detectSessionInUrl: false },
  });

  const pdfBytes = fs.readFileSync(PDF_PATH);
  const rand = crypto.randomUUID().slice(0, 8);
  const pdfStoragePath = `pdfs/issue-${ISSUE.issue_number}-${rand}.pdf`;

  console.log(`Uploading PDF (${(pdfBytes.length / 1024).toFixed(1)} KB) → ${pdfStoragePath} …`);

  const { error: upErr } = await supabase.storage
    .from(BUCKET)
    .upload(pdfStoragePath, pdfBytes, {
      contentType: "application/pdf",
      cacheControl: "31536000",
    });

  if (upErr) {
    console.error("Upload failed:", upErr.message);
    process.exit(1);
  }
  console.log("✓  PDF uploaded.");

  console.log("Inserting issue record…");
  const { error: insErr } = await supabase.from("issues").insert({
    ...ISSUE,
    pdf_path: pdfStoragePath,
    cover_path: null,
    page_count: null,
    file_size_bytes: pdfBytes.length,
  });

  if (insErr) {
    console.error("Insert failed:", insErr.message);
    // Clean up the uploaded file
    await supabase.storage.from(BUCKET).remove([pdfStoragePath]);
    process.exit(1);
  }

  console.log(`✓  Issue No. ${ISSUE.issue_number} published successfully.`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
