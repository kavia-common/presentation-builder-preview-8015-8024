import JSZip from "jszip";
import { saveAs } from "file-saver";

/**
 * STRICT TEMPLATE RULES (user requirements):
 * - The app ships with a built-in PPTX: `public/assets/template.pptx`.
 * - Slide 1 must remain byte-for-byte identical to the template EXCEPT for the
 *   date text content.
 * - The date text remains editable, but its formatting, runs, position, and
 *   layout must remain unchanged.
 * - Do not modify any other elements or slide XML parts.
 * - Preserve the last slide exactly with no processing: we therefore do not
 *   modify any files except `ppt/slides/slide1.xml`.
 *
 * Implementation approach:
 * - Read `ppt/slides/slide1.xml` as a string (do not parse/re-serialize XML).
 * - Select the date text shape strictly by:
 *    1) containing "Date</a:t>" label
 *    2) (if multiple candidates) having the exact expected run sequence
 * - Update only the existing <a:t> node values within the 5 date runs:
 *     "\u00a0 {day}\u00a0", "{mon}", " ", "{yearHead3}", "{yearTail1}"
 *   leaving all <a:rPr> and all XML untouched.
 * - Verify slide1.xml changes occur only within those 5 <a:t> inner text ranges.
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
 * Returns an array of absolute ranges for <a:t>...</a:t> nodes within an XML fragment.
 * Each entry is { start, end, innerStart, innerEnd } where:
 * - start/end cover the entire <a:t ...>...</a:t> element
 * - innerStart/innerEnd cover the inner text only
 */
function collectATextNodeRanges(xml) {
  const ranges = [];
  const re = /<a:t\b[^>]*>[\s\S]*?<\/a:t>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    const fullStart = m.index;
    const fullEnd = fullStart + m[0].length;

    const openEnd = xml.indexOf(">", fullStart);
    const closeStart = xml.lastIndexOf("</a:t>", fullEnd);
    if (openEnd < 0 || closeStart < 0) continue;

    const innerStart = openEnd + 1;
    const innerEnd = closeStart;

    ranges.push({ start: fullStart, end: fullEnd, innerStart, innerEnd });
  }
  return ranges;
}

/**
 * Extracts the inner text for the first <a:t> inside a run (<a:r>...</a:r>).
 * Returns null if no <a:t> present.
 */
function getFirstATextFromRun(runXml) {
  const tMatch = runXml.match(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/);
  return tMatch ? tMatch[1] : null;
}

/**
 * Collects <a:r> blocks (as raw XML strings) inside a paragraph (<a:p>...</a:p>).
 */
function collectRunsFromParagraph(paragraphXml) {
  const runRegex = /<a:r\b[^>]*>[\s\S]*?<\/a:r>/g;
  const runs = [];
  let m;
  while ((m = runRegex.exec(paragraphXml)) !== null) {
    runs.push(m[0]);
  }
  return runs;
}

/**
 * Finds all <p:sp> blocks from slide XML (raw string).
 * This is intentionally regex-based to avoid re-serialization that could change bytes.
 */
function collectShapeBlocks(slide1Xml) {
  const shapes = [];
  const re = /<p:sp\b[\s\S]*?<\/p:sp>/g;
  let m;
  while ((m = re.exec(slide1Xml)) !== null) {
    shapes.push({
      xml: m[0],
      start: m.index,
      end: m.index + m[0].length,
    });
  }
  return shapes;
}

/**
 * Extract shape "name" from cNvPr.
 */
function getShapeName(shapeXml) {
  const m = shapeXml.match(/<p:cNvPr\b[^>]*\bname="([^"]*)"/);
  return m ? m[1] : null;
}

/**
 * Extract shape "id" from cNvPr.
 */
function getShapeId(shapeXml) {
  const m = shapeXml.match(/<p:cNvPr\b[^>]*\bid="([^"]*)"/);
  return m ? m[1] : null;
}

/**
 * Collect all paragraph blocks (<a:p>...</a:p>) within a shape text body.
 */
function collectParagraphsFromShape(shapeXml) {
  const bodyMatch = shapeXml.match(/<p:txBody\b[\s\S]*?<\/p:txBody>/);
  const scope = bodyMatch ? bodyMatch[0] : shapeXml;

  const paragraphs = [];
  const re = /<a:p\b[\s\S]*?<\/a:p>/g;
  let m;
  while ((m = re.exec(scope)) !== null) {
    paragraphs.push({
      xml: m[0],
      startInScope: m.index,
      endInScope: m.index + m[0].length,
    });
  }

  return {
    scopeXml: scope,
    scopeOffset: bodyMatch ? shapeXml.indexOf(scope) : 0,
    paragraphs,
  };
}

/**
 * PUBLIC_INTERFACE
 * Updates ONLY the date field on slide 1 for the shipped default template,
 * keeping all other slides/files untouched (including the last slide).
 *
 * Guarantees:
 * - slide1.xml is identical except for the targeted date <a:t> nodes.
 * - Does not alter any <a:rPr>, paragraph properties, shape geometry, or layout.
 *
 * @param {ArrayBuffer} pptxArrayBuffer
 * @param {string} dateISO - value from <input type="date">
 * @returns {Promise<{updatedPptxBytes: Uint8Array, detected: {mode: string, slidePath: string, shapeName: string, shapeId: string}}>}
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

  const shapes = collectShapeBlocks(slide1Xml);
  const dateLabelCandidates = shapes.filter((s) => s.xml.includes("Date</a:t>"));

  if (dateLabelCandidates.length < 1) {
    throw new Error(
      'Strict template mismatch: could not find any shape containing the "Date" label.'
    );
  }

  let shape = null;
  if (dateLabelCandidates.length === 1) {
    shape = dateLabelCandidates[0];
  } else {
    const expectedRunTexts = ["Date", " ", ":", "\u00a0 24\u00a0", "Dec", " ", "202", "5"];

    const matchesExpectedRuns = (shapeXml) => {
      const { paragraphs } = collectParagraphsFromShape(shapeXml);
      const p = paragraphs.find((x) => x.xml.includes("Date</a:t>"));
      if (!p) return false;

      const runs = collectRunsFromParagraph(p.xml).map((runXml) => ({
        tText: getFirstATextFromRun(runXml),
        // CRITICAL FIX: use `<a:t\b` (word boundary), not `<a:t\\b` (literal backslash+b)
        hasText: /<a:t\b/.test(runXml),
      }));

      const actualRunTexts = runs.map((r) => (r.hasText ? r.tText : null));
      return (
        actualRunTexts.length === expectedRunTexts.length &&
        actualRunTexts.every((v, i) => v === expectedRunTexts[i])
      );
    };

    shape = dateLabelCandidates.find((c) => matchesExpectedRuns(c.xml)) ?? null;
  }

  if (!shape) {
    const names = shapes
      .map((s) => getShapeName(s.xml))
      .filter(Boolean)
      .slice(0, 15)
      .join(", ");
    throw new Error(
      `Strict template mismatch: multiple "Date" label candidates and none matched the expected run pattern. Available names (first 15): ${names}`
    );
  }

  const shapeName = getShapeName(shape.xml) ?? "unknown";
  const shapeId = getShapeId(shape.xml) ?? "unknown";

  const { scopeXml, scopeOffset, paragraphs } = collectParagraphsFromShape(shape.xml);
  if (!paragraphs.length) {
    throw new Error("Strict template mismatch: date shape contains no paragraphs.");
  }

  const paragraphIndex = paragraphs.findIndex((p) => p.xml.includes("Date</a:t>"));
  if (paragraphIndex < 0) {
    throw new Error('Strict template mismatch: could not find "Date" paragraph inside date shape.');
  }

  const paragraphXml = paragraphs[paragraphIndex].xml;
  const runs = collectRunsFromParagraph(paragraphXml).map((runXml) => ({
    xml: runXml,
    tText: getFirstATextFromRun(runXml),
    hasText: /<a:t\b/.test(runXml),
  }));

  const expectedRunTexts = ["Date", " ", ":", "\u00a0 24\u00a0", "Dec", " ", "202", "5"];
  const actualRunTexts = runs.map((r) => (r.hasText ? r.tText : null));

  const sameLength = actualRunTexts.length === expectedRunTexts.length;
  const matches = sameLength && actualRunTexts.every((v, i) => v === expectedRunTexts[i]);

  if (!matches) {
    throw new Error(
      "Strict template mismatch: date paragraph runs differ from expected template; refusing to modify."
    );
  }

  const [day, mon, year] = formattedDate.split(" ");
  const y = String(year);
  const yHead = y.slice(0, 3);
  const yTail = y.slice(3);

  const replacements = [`\u00a0 ${day}\u00a0`, mon, " ", yHead, yTail].map(escapeXmlText);

  const dateRunIndexes = [3, 4, 5, 6, 7];
  const updatedRunsXml = runs.map((r) => r.xml);

  for (let i = 0; i < dateRunIndexes.length; i += 1) {
    const idx = dateRunIndexes[i];
    const newText = replacements[i];

    updatedRunsXml[idx] = updatedRunsXml[idx].replace(
      /(<a:t\b[^>]*>)([\s\S]*?)(<\/a:t>)/,
      `$1${newText}$3`
    );
  }

  const firstRunIdx = paragraphXml.search(/<a:r\b/);
  const lastRunEnd = paragraphXml.lastIndexOf("</a:r>");
  if (firstRunIdx < 0 || lastRunEnd < 0) {
    throw new Error("Strict template mismatch: could not parse date paragraph runs.");
  }
  const runRegionEnd = lastRunEnd + "</a:r>".length;
  const pHead = paragraphXml.slice(0, firstRunIdx);
  const pTail = paragraphXml.slice(runRegionEnd);
  const updatedParagraphXml = `${pHead}${updatedRunsXml.join("")}${pTail}`;

  const paraStartInScope = paragraphs[paragraphIndex].startInScope;
  const paraEndInScope = paragraphs[paragraphIndex].endInScope;

  const updatedScopeXml =
    scopeXml.slice(0, paraStartInScope) + updatedParagraphXml + scopeXml.slice(paraEndInScope);

  const updatedShapeXml =
    shape.xml.slice(0, scopeOffset) +
    updatedScopeXml +
    shape.xml.slice(scopeOffset + scopeXml.length);

  const updatedSlide1Xml =
    slide1Xml.slice(0, shape.start) + updatedShapeXml + slide1Xml.slice(shape.end);

  // IMPORTANT:
  // `innerStartWithinParagraphShifted` / `innerEndWithinParagraphShifted` are already offsets
  // within the *full paragraph XML string* (not within the <a:r> run only).
  //
  // To convert them into absolute positions within slide1.xml we need:
  //   shape.start (absolute start of the <p:sp> block in slide1.xml)
  // + scopeOffset (offset from shape start to the <p:txBody> scope we used)
  // + paraStartInScope (offset from the scope start to the <a:p> paragraph start)
  // + innerStartWithinParagraphShifted (offset from paragraph start to the <a:t> inner text)
  //
  // The previous implementation accidentally double-added `shape.start`, making the guard
  // think legitimate date edits were outside the allowed ranges.
  verifyOnlyAllowedSlide1Diffs({
    originalSlide1Xml: slide1Xml,
    updatedSlide1Xml,
    allowedChangedATextInnerRangesInUpdated: computeAllowedATextInnerRangesInUpdatedParagraph(
      updatedParagraphXml,
      dateRunIndexes
    ).map((r) => ({
      start: shape.start + scopeOffset + paraStartInScope + r.innerStartWithinParagraphShifted,
      end: shape.start + scopeOffset + paraStartInScope + r.innerEndWithinParagraphShifted,
    })),
  });

  zip.file(SLIDE1_PATH, updatedSlide1Xml);

  const out = await zip.generateAsync({ type: "uint8array" });

  return {
    updatedPptxBytes: out,
    detected: {
      mode: "strict-template/date-shape-by-label+diff-guard",
      slidePath: SLIDE1_PATH,
      shapeName,
      shapeId,
    },
  };
}

/**
 * Computes which <a:t> inner text ranges (in the UPDATED paragraph XML) are allowed to differ.
 * We only allow differences in the first <a:t> of the 5 targeted date runs.
 *
 * Returns array of objects:
 *  { innerStartWithinParagraphShifted, innerEndWithinParagraphShifted }
 * where the numbers are offsets within the full paragraph XML string.
 */
function computeAllowedATextInnerRangesInUpdatedParagraph(updatedParagraphXml, dateRunIndexes) {
  const runRegex = /<a:r\b[^>]*>[\s\S]*?<\/a:r>/g;

  const runMatches = [];
  let m;
  while ((m = runRegex.exec(updatedParagraphXml)) !== null) {
    runMatches.push({ runXml: m[0], start: m.index, end: m.index + m[0].length });
  }

  if (runMatches.length < Math.max(...dateRunIndexes) + 1) {
    throw new Error("Internal error: updated paragraph run count changed unexpectedly.");
  }

  const allowed = [];

  for (const idx of dateRunIndexes) {
    const run = runMatches[idx];
    const runTextRanges = collectATextNodeRanges(run.runXml);
    if (!runTextRanges.length) {
      throw new Error("Strict template mismatch: expected <a:t> in targeted date run.");
    }
    const first = runTextRanges[0];

    allowed.push({
      innerStartWithinParagraphShifted: run.start + first.innerStart,
      innerEndWithinParagraphShifted: run.start + first.innerEnd,
    });
  }

  return allowed;
}

/**
 * Verifies that differences between originalSlide1Xml and updatedSlide1Xml
 * occur ONLY inside the allowed <a:t> inner text ranges in the UPDATED XML.
 *
 * If any other difference is detected, throws an Error and refuses to output.
 */
function verifyOnlyAllowedSlide1Diffs({
  originalSlide1Xml,
  updatedSlide1Xml,
  allowedChangedATextInnerRangesInUpdated,
}) {
  if (originalSlide1Xml === updatedSlide1Xml) return;

  const maxLen = Math.max(originalSlide1Xml.length, updatedSlide1Xml.length);

  const isAllowedPos = (pos) =>
    allowedChangedATextInnerRangesInUpdated.some((r) => pos >= r.start && pos < r.end);

  let i = 0;
  while (i < maxLen) {
    const a = originalSlide1Xml[i];
    const b = updatedSlide1Xml[i];

    if (a === b) {
      i += 1;
      continue;
    }

    if (i >= originalSlide1Xml.length || i >= updatedSlide1Xml.length) {
      if (!isAllowedPos(i)) {
        throw new Error(
          "Safety check failed: slide1.xml length/content changed outside the date <a:t> nodes. Aborting."
        );
      }
      i += 1;
      continue;
    }

    if (!isAllowedPos(i)) {
      const ctxStart = Math.max(0, i - 40);
      const ctxEnd = Math.min(updatedSlide1Xml.length, i + 80);
      const ctx = updatedSlide1Xml.slice(ctxStart, ctxEnd);
      throw new Error(
        `Safety check failed: unintended slide1.xml modification detected at position ${i} outside allowed date text nodes. Context: ${ctx}`
      );
    }

    i += 1;
  }
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
 * Creates an object URL suitable for embedding an Office preview iframe/object.
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

/**
 * PUBLIC_INTERFACE
 * Verifies the strict invariant: the last slide XML must remain byte-for-byte identical
 * between the original template PPTX and the updated PPTX output.
 *
 * NOTE: We compare the raw bytes of the ZIP entry `ppt/slides/slide14.xml` because
 * the shipped template currently has 14 slides. If the template changes, update
 * this constant accordingly.
 *
 * @param {ArrayBuffer} originalPptxArrayBuffer
 * @param {Uint8Array} updatedPptxBytes
 * @returns {Promise<boolean>} true if unchanged, else throws Error
 */
export async function assertLastSlideUnchanged(originalPptxArrayBuffer, updatedPptxBytes) {
  const LAST_SLIDE_PATH = "ppt/slides/slide14.xml";

  const originalZip = await JSZip.loadAsync(originalPptxArrayBuffer);
  const updatedZip = await JSZip.loadAsync(updatedPptxBytes);

  const orig = originalZip.file(LAST_SLIDE_PATH);
  const next = updatedZip.file(LAST_SLIDE_PATH);

  if (!orig || !next) {
    throw new Error(
      `Invariant check failed: missing ${LAST_SLIDE_PATH} in ${!orig ? "original" : "updated"} PPTX.`
    );
  }

  const [origBytes, nextBytes] = await Promise.all([
    orig.async("uint8array"),
    next.async("uint8array"),
  ]);

  if (origBytes.length !== nextBytes.length) {
    throw new Error("Invariant check failed: last slide byte length changed.");
  }

  for (let i = 0; i < origBytes.length; i += 1) {
    if (origBytes[i] !== nextBytes[i]) {
      throw new Error(`Invariant check failed: last slide bytes differ at offset ${i}.`);
    }
  }

  return true;
}
