import JSZip from "jszip";
import { saveAs } from "file-saver";

/**
 * NOTE ABOUT THE TEMPLATE:
 * - Place the bundled template at: `public/assets/template.pptx`
 * - This app also allows uploading a PPTX, but the template must have either:
 *   (A) a well-documented placeholder token like {{DATE}} in slide 1, OR
 *   (B) a "Date:" label followed by date fragments (as in the provided PPTX).
 *
 * This implementation uses a robust fallback:
 * - If "{{DATE}}" exists anywhere in slide1.xml, replace that token only.
 * - Otherwise, locate the paragraph that contains a run with "Date" and then
 *   rewrite only the subsequent runs that currently form the date value.
 *
 * By editing ONLY `ppt/slides/slide1.xml` and leaving all other zip entries
 * untouched, we keep the last slide unchanged.
 */

const SLIDE1_PATH = "ppt/slides/slide1.xml";
const PLACEHOLDER_TOKEN = "{{DATE}}";

/**
 * PUBLIC_INTERFACE
 * Fetches the bundled template PPTX from /public/assets/template.pptx.
 * @returns {Promise<ArrayBuffer>} PPTX bytes
 */
export async function fetchBundledTemplatePptx() {
  const res = await fetch("/assets/template.pptx");
  if (!res.ok) {
    throw new Error(
      `Failed to fetch bundled template from /assets/template.pptx (HTTP ${res.status}).`
    );
  }
  return await res.arrayBuffer();
}

/**
 * PUBLIC_INTERFACE
 * Reads a user-selected file into ArrayBuffer.
 * @param {File} file
 * @returns {Promise<ArrayBuffer>}
 */
export function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Failed to read file."));
    reader.onload = () => resolve(reader.result);
    reader.readAsArrayBuffer(file);
  });
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function formatDateForTemplate(dateInput) {
  // Template’s visible format is like: "24 Dec 2025"
  // We render day without leading zero.
  const d = new Date(dateInput);
  if (Number.isNaN(d.getTime())) return "";
  const day = d.getDate();
  const month = d.toLocaleString("en-GB", { month: "short" });
  const year = d.getFullYear();
  return `${day} ${month} ${year}`;
}

function escapeXmlText(text) {
  return text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

/**
 * Finds the <a:p> paragraph range (start,end) that contains "Date" label.
 * This is designed to match the provided PPTX where "Date" is a run in a single paragraph.
 */
function findDateParagraphBounds(slide1Xml) {
  const idx = slide1Xml.indexOf("Date</a:t>");
  if (idx < 0) return null;

  const pStart = slide1Xml.lastIndexOf("<a:p", idx);
  const pEnd = slide1Xml.indexOf("</a:p>", idx);
  if (pStart < 0 || pEnd < 0) return null;

  return { pStart, pEnd: pEnd + "</a:p>".length };
}

/**
 * Updates date in a paragraph by replacing the sequence of <a:t> runs that represent the date value.
 *
 * Strategy for the provided template (observed pattern):
 *   runs: [" ", ":", "\u00a0 24\u00a0", "Dec", " ", "202", "5"]
 * We keep the "Date : " label runs intact and only replace the following date runs with:
 *   runs: ["\u00a0 {day}\u00a0", "Mon", " ", "YYYY"]
 * with year potentially split; we will place full year in a single run to keep it simple.
 *
 * NOTE:
 * This does not modify other shapes/textboxes, only this paragraph’s date runs.
 */
function updateDateRunsInParagraph(paragraphXml, formattedDate) {
  // Prefer explicit placeholder if present.
  if (paragraphXml.includes(PLACEHOLDER_TOKEN)) {
    return paragraphXml.replaceAll(PLACEHOLDER_TOKEN, escapeXmlText(formattedDate));
  }

  // Extract runs: capture entire <a:r ...>...</a:r> blocks so we can preserve styling.
  const runRegex = /<a:r>([\s\S]*?)<\/a:r>/g;
  const runs = [];
  let match;
  while ((match = runRegex.exec(paragraphXml)) !== null) {
    runs.push({ full: match[0], inner: match[1], start: match.index, end: runRegex.lastIndex });
  }

  // Find the run index containing <a:t>Date</a:t>
  const dateLabelIdx = runs.findIndex((r) => /<a:t>\s*Date\s*<\/a:t>/.test(r.full));
  if (dateLabelIdx < 0) {
    // As a fallback, if we cannot reliably find label in runs, do not mutate.
    return paragraphXml;
  }

  // After "Date", there are typically runs for space, ":" etc. We want to start replacement
  // at the first run after ":" (or after "Date" if ":" not present).
  let replaceFrom = dateLabelIdx + 1;
  const colonIdx = runs.findIndex((r, i) => i > dateLabelIdx && /<a:t>\s*:\s*<\/a:t>/.test(r.full));
  if (colonIdx >= 0) replaceFrom = colonIdx + 1;

  // Replace until end of paragraph runs (but keep the paragraph structure).
  // Create new runs using the first date-value run’s <a:rPr> (style) if available.
  const styleSourceRun = runs[replaceFrom] || runs[colonIdx] || runs[dateLabelIdx];
  const rPrMatch = styleSourceRun.full.match(/<a:rPr[\s\S]*?<\/a:rPr>/);
  const rPr = rPrMatch ? rPrMatch[0] : null;

  const [day, mon, year] = formattedDate.split(" ");
  const newTexts = [
    // Keep non-breaking spaces around day to match template spacing.
    `\u00a0 ${day}\u00a0`,
    mon,
    " ",
    year,
  ];

  const newRunBlocks = newTexts.map((t) => {
    const safeText = escapeXmlText(t);
    const rPrXml = rPr ? rPr : "";
    // If we include rPr, ensure we keep it inside <a:r> in the same order as typical PPTX: rPr then t.
    return `<a:r>${rPrXml}<a:t>${safeText}</a:t></a:r>`;
  });

  // Rebuild paragraph: keep everything up to replaceFrom runs, then append new runs.
  const prefix = runs.slice(0, replaceFrom).map((r) => r.full).join("");
  const suffix = ""; // replace until end (we only want one editable field)
  // But we must preserve any content after date inside same paragraph if any; in this template there isn't.
  // If there is, this would be a future enhancement: detect end of date runs more precisely.
  const beforeRunsStart = paragraphXml.indexOf(runs[0]?.full ?? "");
  const afterRunsEnd =
    runs.length > 0 ? paragraphXml.lastIndexOf(runs[runs.length - 1].full) + runs[runs.length - 1].full.length : 0;

  const paragraphHead = paragraphXml.slice(0, beforeRunsStart);
  const paragraphTail = paragraphXml.slice(afterRunsEnd);

  return `${paragraphHead}${prefix}${newRunBlocks.join("")}${suffix}${paragraphTail}`;
}

/**
 * PUBLIC_INTERFACE
 * Updates ONLY the date field on slide 1, keeping all other slides and files unchanged.
 *
 * Returns:
 * - updatedPptxBytes: Uint8Array
 * - detected: info about detection method (placeholder vs date-paragraph)
 *
 * @param {ArrayBuffer} pptxArrayBuffer
 * @param {string} dateISO - value from <input type="date">
 */
export async function updatePptxDateOnly(pptxArrayBuffer, dateISO) {
  const formattedDate = formatDateForTemplate(dateISO);
  if (!formattedDate) {
    throw new Error("Invalid date.");
  }

  const zip = await JSZip.loadAsync(pptxArrayBuffer);
  const slide1File = zip.file(SLIDE1_PATH);
  if (!slide1File) {
    throw new Error(`Template missing expected file: ${SLIDE1_PATH}`);
  }

  const slide1Xml = await slide1File.async("string");

  // If template supports explicit placeholder, this is safest and most deterministic.
  if (slide1Xml.includes(PLACEHOLDER_TOKEN)) {
    const updatedXml = slide1Xml.replaceAll(PLACEHOLDER_TOKEN, escapeXmlText(formattedDate));
    zip.file(SLIDE1_PATH, updatedXml);
    const out = await zip.generateAsync({ type: "uint8array" });
    return { updatedPptxBytes: out, detected: { mode: "token", token: PLACEHOLDER_TOKEN } };
  }

  // Fallback: locate the paragraph containing "Date" and rewrite date runs.
  const bounds = findDateParagraphBounds(slide1Xml);
  if (!bounds) {
    throw new Error(
      'Could not detect date placeholder on slide 1. Add a "{{DATE}}" token in the template or ensure slide 1 contains a "Date" label.'
    );
  }

  const paragraphXml = slide1Xml.slice(bounds.pStart, bounds.pEnd);
  const updatedParagraphXml = updateDateRunsInParagraph(paragraphXml, formattedDate);

  const updatedSlide1Xml =
    slide1Xml.slice(0, bounds.pStart) + updatedParagraphXml + slide1Xml.slice(bounds.pEnd);

  zip.file(SLIDE1_PATH, updatedSlide1Xml);

  // IMPORTANT: all other slides are left untouched in the ZIP,
  // ensuring the last slide remains exactly as-is.
  const out = await zip.generateAsync({ type: "uint8array" });
  return { updatedPptxBytes: out, detected: { mode: "date-paragraph", slidePath: SLIDE1_PATH } };
}

/**
 * PUBLIC_INTERFACE
 * Triggers a browser download of the PPTX bytes.
 * @param {Uint8Array} pptxBytes
 * @param {string} filename
 */
export function downloadPptxBytes(pptxBytes, filename) {
  const blob = new Blob([pptxBytes], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
  saveAs(blob, filename);
}

/**
 * PUBLIC_INTERFACE
 * Creates an object URL suitable for embedding an Office Online preview iframe.
 * Note: Some browsers may block local blob URLs in certain iframe contexts.
 * @param {Uint8Array} pptxBytes
 */
export function createPptxObjectUrl(pptxBytes) {
  const blob = new Blob([pptxBytes], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
  return URL.createObjectURL(blob);
}

/**
 * PUBLIC_INTERFACE
 * Formats today's date for <input type="date"> default value (yyyy-mm-dd).
 */
export function todayIsoDate() {
  const d = new Date();
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}
