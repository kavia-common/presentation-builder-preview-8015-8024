import JSZip from "jszip";
import { saveAs } from "file-saver";

/**
 * STRICT TEMPLATE RULES (user requirements):
 * - The app ships with a built-in PPTX: `public/assets/template.pptx`.
 * - The last slide must be preserved exactly: we therefore do not modify any files
 *   except `ppt/slides/slide1.xml`, and inside that file we only mutate the exact
 *   date text nodes (no reformatting, no reflow, no styling changes).
 * - On slide 1, ONLY the date field is editable, and its visible formatting (font,
 *   size, color, spacing, position, and locale/format) must remain exactly as in
 *   the original file.
 *
 * Implementation approach:
 * - We load `ppt/slides/slide1.xml` and find the paragraph containing the "Date"
 *   label run, then locate the exact sequence of date runs:
 *     ["\u00a0 24\u00a0", "Dec", " ", "202", "5"]
 * - We replace text INSIDE those existing <a:t> nodes only, keeping the number of
 *   runs and their <a:rPr> styling identical.
 * - We intentionally do not support uploads or generic templates in this strict mode.
 */

const SLIDE1_PATH = "ppt/slides/slide1.xml";

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

function pad2(n) {
  return String(n).padStart(2, "0");
}

/**
 * The template’s visible date format is: "24 Dec 2025" (en-GB short month).
 * We must keep that exact locale/format.
 */
function formatDateForTemplate(dateInput) {
  const d = new Date(dateInput);
  if (Number.isNaN(d.getTime())) return "";
  const day = d.getDate(); // no leading zero
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
 * Finds the <a:p> paragraph range (start,end) that contains the run "Date".
 * This matches the shipped template.
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
 * Updates ONLY the exact date runs in the paragraph:
 *   "\u00a0 24\u00a0", "Dec", " ", "202", "5"
 * by replacing the text content inside those existing <a:t> nodes.
 *
 * This preserves:
 * - number of runs
 * - run properties (<a:rPr>)
 * - spacing and position behavior
 *
 * @param {string} paragraphXml - the full <a:p>...</a:p> block
 * @param {string} formattedDate - e.g. "24 Dec 2025"
 */
function updateExactTemplateDateRuns(paragraphXml, formattedDate) {
  // Collect all <a:r> blocks and their first <a:t> contents.
  const runRegex = /<a:r\b[^>]*>[\s\S]*?<\/a:r>/g;
  const runs = [];
  let m;
  while ((m = runRegex.exec(paragraphXml)) !== null) {
    const full = m[0];
    const tMatch = full.match(/<a:t[^>]*>([\s\S]*?)<\/a:t>/);
    // Not all runs have <a:t>, but in this paragraph they do; still be defensive.
    runs.push({
      full,
      tText: tMatch ? tMatch[1] : null,
      hasText: !!tMatch,
    });
  }

  // We expect a specific sequence in the shipped template:
  // ["Date", " ", ":", "\u00a0 24\u00a0", "Dec", " ", "202", "5"]
  const dateLabelIdx = runs.findIndex(
    (r) => r.hasText && r.tText === "Date"
  );
  if (dateLabelIdx < 0) {
    throw new Error('Strict template mismatch: could not find "Date" run.');
  }

  // Validate the subsequent run texts so we only ever touch the intended field.
  const expected = [
    " ",
    ":",
    "\u00a0 24\u00a0",
    "Dec",
    " ",
    "202",
    "5",
  ];
  const actual = runs
    .slice(dateLabelIdx + 1, dateLabelIdx + 1 + expected.length)
    .map((r) => (r.hasText ? r.tText : null));

  const matchesExpected =
    actual.length === expected.length &&
    actual.every((v, i) => v === expected[i]);

  if (!matchesExpected) {
    throw new Error(
      "Strict template mismatch: date field runs are not in the expected shape; refusing to modify."
    );
  }

  const [day, mon, year] = formattedDate.split(" ");
  const y = String(year);
  const yHead = y.slice(0, 3);
  const yTail = y.slice(3);

  // Replace ONLY the <a:t> contents for the 5 date value runs, preserving the run blocks.
  const replacements = [
    `\u00a0 ${day}\u00a0`, // keep NBSP and spacing pattern exactly
    mon,
    " ",
    yHead,
    yTail,
  ].map(escapeXmlText);

  const replaceRunIndexes = [
    dateLabelIdx + 3, // "\u00a0 24\u00a0"
    dateLabelIdx + 4, // "Dec"
    dateLabelIdx + 5, // " "
    dateLabelIdx + 6, // "202"
    dateLabelIdx + 7, // "5"
  ];

  const updatedRuns = runs.map((r) => r.full);

  for (let i = 0; i < replaceRunIndexes.length; i += 1) {
    const runIdx = replaceRunIndexes[i];
    const newText = replacements[i];

    // Replace only the inner text of the first <a:t> in this run block.
    updatedRuns[runIdx] = updatedRuns[runIdx].replace(
      /(<a:t[^>]*>)([\s\S]*?)(<\/a:t>)/,
      `$1${newText}$3`
    );
  }

  // Rebuild paragraph by replacing the original concatenated runs with updated ones.
  // Preserve paragraph wrapper and any other elements (pPr/endParaRPr) by
  // replacing only the run blocks region.
  const firstRunIdx = paragraphXml.search(/<a:r\b/);
  const lastRunEnd = paragraphXml.lastIndexOf("</a:r>");
  if (firstRunIdx < 0 || lastRunEnd < 0) {
    throw new Error("Strict template mismatch: could not parse paragraph runs.");
  }

  const runRegionEnd = lastRunEnd + "</a:r>".length;
  const head = paragraphXml.slice(0, firstRunIdx);
  const tail = paragraphXml.slice(runRegionEnd);

  return `${head}${updatedRuns.join("")}${tail}`;
}

/**
 * PUBLIC_INTERFACE
 * Updates ONLY the date field on slide 1 for the shipped default template,
 * keeping all other slides/files untouched (including the last slide).
 *
 * @param {ArrayBuffer} pptxArrayBuffer
 * @param {string} dateISO - value from <input type="date">
 * @returns {Promise<{updatedPptxBytes: Uint8Array, detected: {mode: string, slidePath: string}}>}
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

  const bounds = findDateParagraphBounds(slide1Xml);
  if (!bounds) {
    throw new Error(
      'Strict template mismatch: could not find the Slide 1 "Date" paragraph; refusing to modify.'
    );
  }

  const paragraphXml = slide1Xml.slice(bounds.pStart, bounds.pEnd);
  const updatedParagraphXml = updateExactTemplateDateRuns(
    paragraphXml,
    formattedDate
  );

  const updatedSlide1Xml =
    slide1Xml.slice(0, bounds.pStart) +
    updatedParagraphXml +
    slide1Xml.slice(bounds.pEnd);

  // This is the only mutation in the whole PPTX.
  zip.file(SLIDE1_PATH, updatedSlide1Xml);

  // IMPORTANT: We do not touch any other ZIP entries (slides, rels, media, etc.).
  // This preserves the last slide and all other content exactly.
  const out = await zip.generateAsync({ type: "uint8array" });

  return {
    updatedPptxBytes: out,
    detected: { mode: "strict-template", slidePath: SLIDE1_PATH },
  };
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
 * Creates an object URL suitable for embedding an Office preview iframe.
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
