import JSZip from "jszip";
import { saveAs } from "file-saver";

/**
 * STRICT TEMPLATE RULES (user requirements):
 * - The app ships with a built-in PPTX: `public/assets/template.pptx`.
 * - Slide 1 must remain byte-for-byte identical to the template EXCEPT for:
 *    (a) the date text content (editable), AND
 *    (b) replacing the existing text/empty placeholder directly above the Name
 *        with the non-editable label "TATA ELXSI" (Slide 1 only).
 * - Do NOT add a new line for the label and do NOT overlap the Name:
 *   we must use the existing placeholder/shape that sits above the Name and set
 *   its text to exactly "TATA ELXSI".
 * - Preserve the last slide exactly with no processing: we therefore do not
 *   modify any files except `ppt/slides/slide1.xml`.
 *
 * Implementation approach (byte-preserving):
 * - Read `ppt/slides/slide1.xml` as a string (do not parse/re-serialize XML).
 * - Ensure the "TATA ELXSI" label exists by targeting the existing shape
 *   above Name in the bundled template (shape cNvPr id="8", name="object 8").
 * - Update only:
 *    - the <a:t> value(s) within that label placeholder shape (slide 1 only),
 *      without changing any geometry or adding paragraphs.
 *    - the existing <a:t> node values within the 5 date runs.
 *
 * NOTE:
 * - We intentionally keep the label non-editable: it is enforced to the same
 *   string on every generation.
 */

const SLIDE1_PATH = "ppt/slides/slide1.xml";
const STATIC_LABEL_TEXT = "TATA ELXSI";

/**
 * Template-specific: the placeholder directly above the "Name" row (verified in the bundled template).
 *
 * IMPORTANT:
 * - We must NOT move shapes or change geometry; only replace text in the existing placeholder.
 * - This ensures alignment/spacing exactly matches the template (and the latest screenshot reference).
 */
const LABEL_SHAPE_ID = "8";
const LABEL_SHAPE_NAME = "object 8";

/**
 * PUBLIC_INTERFACE
 * Fetches the bundled template PPTX from /public/assets/template.pptx.
 * @returns {Promise<ArrayBuffer>} PPTX bytes
 */
export async function fetchBundledTemplatePptx() {
  /** This is a public function. */
  const res = await fetch("/assets/template.pptx", { cache: "no-store" });
  if (!res.ok) {
    throw new Error(
      `Failed to fetch bundled template from /assets/template.pptx (HTTP ${res.status}).`
    );
  }

  const buf = await res.arrayBuffer();

  // Lightweight validity check: PPTX is a ZIP => first 2 bytes should be 'PK'.
  const header = new Uint8Array(buf.slice(0, 2));
  if (!(header[0] === 0x50 && header[1] === 0x4b)) {
    throw new Error(
      "Bundled template does not appear to be a valid PPTX (ZIP header missing). Check /public/assets/template.pptx."
    );
  }

  return buf;
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
  return String(text ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
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
 * Extracts the full first <a:t ...>...</a:t> node from a run XML.
 * Returns null if no <a:t>.
 */
function getFirstATextNodeFromRun(runXml) {
  const m = runXml.match(/<a:t\b[^>]*>[\s\S]*?<\/a:t>/);
  return m ? m[0] : null;
}

/**
 * Verifies that differences between originalSlide1Xml and updatedSlide1Xml
 * occur ONLY within the inner text of the intended date <a:t> nodes on slide 1.
 *
 * This avoids brittle positional diffs (day can change 2 digits -> 1 digit, shifting
 * subsequent characters and causing false positives).
 */
function verifyOnlyAllowedSlide1DiffsByATextNodes({
  originalSlide1Xml,
  updatedSlide1Xml,
  dateParagraphXmlOriginal,
  allowedDateRunIndexes,
}) {
  if (originalSlide1Xml === updatedSlide1Xml) return;

  const aTextRe = /<a:t\b[^>]*>[\s\S]*?<\/a:t>/g;

  // 1) Structure check: same number of <a:t> nodes.
  const origNodes = originalSlide1Xml.match(aTextRe) ?? [];
  const nextNodes = updatedSlide1Xml.match(aTextRe) ?? [];
  if (origNodes.length !== nextNodes.length) {
    throw new Error(
      "Safety check failed: slide1.xml <a:t> node count changed. Only date text nodes may change."
    );
  }

  // 2) Non-<a:t> parts must match exactly.
  const origParts = originalSlide1Xml.split(aTextRe);
  const nextParts = updatedSlide1Xml.split(aTextRe);
  if (origParts.length !== nextParts.length) {
    throw new Error(
      "Safety check failed: slide1.xml structure changed (unexpected <a:t> segmentation)."
    );
  }
  for (let i = 0; i < origParts.length; i += 1) {
    if (origParts[i] !== nextParts[i]) {
      throw new Error(
        "Safety check failed: slide1.xml changed outside <a:t> nodes. Only date text may change."
      );
    }
  }

  // 3) Determine which global <a:t> nodes correspond to the intended date runs.
  const origParaPos = originalSlide1Xml.indexOf(dateParagraphXmlOriginal);
  if (origParaPos < 0) {
    throw new Error(
      "Safety check failed: could not locate the expected date paragraph in original slide1.xml."
    );
  }

  const paraRuns = collectRunsFromParagraph(dateParagraphXmlOriginal);
  const allowedLocalATextNodes = allowedDateRunIndexes.map((runIdx) => {
    const run = paraRuns[runIdx];
    const node = run ? getFirstATextNodeFromRun(run) : null;
    if (!node) {
      throw new Error(
        "Safety check failed: expected <a:t> node in one of the targeted date runs."
      );
    }
    return node;
  });

  // Map node-string occurrences deterministically to global indexes (handle duplicates).
  const queues = new Map();
  origNodes.forEach((node, idx) => {
    const q = queues.get(node) ?? [];
    q.push(idx);
    queues.set(node, q);
  });

  const allowedGlobalIndexes = new Set();
  for (const node of allowedLocalATextNodes) {
    const q = queues.get(node) ?? [];
    if (!q.length) {
      throw new Error(
        "Safety check failed: could not map date <a:t> node to global index in slide1.xml."
      );
    }
    allowedGlobalIndexes.add(q.shift());
  }

  // 4) All <a:t> nodes except the allowed ones must be identical byte-for-byte.
  for (let i = 0; i < origNodes.length; i += 1) {
    if (allowedGlobalIndexes.has(i)) continue;
    if (origNodes[i] !== nextNodes[i]) {
      throw new Error(
        "Safety check failed: slide1.xml modified in a non-date <a:t> node. Only date text may change."
      );
    }
  }
}

/**
 * Sets the static label text into the existing placeholder shape above the Name.
 *
 * IMPORTANT:
 * - This MUST preserve styling + precise placement from the PPT template (per latest image).
 *   Therefore we ONLY replace text inside an existing <a:t> node of the existing shape.
 * - We do NOT add paragraphs, do NOT add new runs, and do NOT change geometry.
 * - If the placeholder has no existing <a:t>, we insert exactly ONE styled run
 *   (user-approved) into the existing first paragraph, without adding paragraphs
 *   and without changing any geometry.
 */
function setStaticLabelInExistingPlaceholder({ slide1Xml }) {
  const shapes = collectShapeBlocks(slide1Xml);

  const labelCandidates = shapes.filter((s) => {
    const sid = getShapeId(s.xml);
    const sname = getShapeName(s.xml);
    return sid === LABEL_SHAPE_ID || sname === LABEL_SHAPE_NAME;
  });

  if (labelCandidates.length !== 1) {
    throw new Error(
      `Strict template mismatch: expected exactly 1 label placeholder shape (id=${LABEL_SHAPE_ID} or name="${LABEL_SHAPE_NAME}"), found ${labelCandidates.length}.`
    );
  }

  const shape = labelCandidates[0];

  // If label already exists exactly, do nothing.
  if (shape.xml.includes(`>${STATIC_LABEL_TEXT}<`)) {
    return { updatedSlide1Xml: slide1Xml, changed: false };
  }

  const { scopeXml, scopeOffset, paragraphs } = collectParagraphsFromShape(shape.xml);
  if (!paragraphs.length) {
    throw new Error("Strict template mismatch: label placeholder contains no paragraphs.");
  }

  // Use the first paragraph only; we must not introduce extra lines (no new <a:p>).
  const para = paragraphs[0];
  const paraXml = para.xml;

  const runs = collectRunsFromParagraph(paraXml);
  const runIndexWithText = runs.findIndex((r) => /<a:t\b/.test(r));

  // Path A: Preferred — replace text in an existing <a:t> to preserve styling/layout.
  if (runIndexWithText >= 0) {
    const targetRun = runs[runIndexWithText];

    const replacedRun = targetRun.replace(
      /(<a:t\b[^>]*>)([\s\S]*?)(<\/a:t>)/,
      `$1${escapeXmlText(STATIC_LABEL_TEXT)}$3`
    );

    const updatedParaXml = paraXml.replace(targetRun, replacedRun);

    const updatedScopeXml =
      scopeXml.slice(0, para.startInScope) +
      updatedParaXml +
      scopeXml.slice(para.endInScope);

    const updatedShapeXml =
      shape.xml.slice(0, scopeOffset) +
      updatedScopeXml +
      shape.xml.slice(scopeOffset + scopeXml.length);

    const updatedSlide1Xml =
      slide1Xml.slice(0, shape.start) + updatedShapeXml + slide1Xml.slice(shape.end);

    return { updatedSlide1Xml, changed: true };
  }

  /**
   * Path B: User-approved exception — the label placeholder has no <a:t> runs.
   * We must insert EXACTLY ONE styled run into the existing first paragraph:
   * - Do NOT add paragraphs.
   * - Do NOT add multiple runs.
   * - Preserve styling context by cloning the first existing <a:r> (with <a:rPr>)
   *   and replacing its content with a single <a:t>TATA ELXSI</a:t>.
   *
   * This keeps geometry untouched (no <a:xfrm> changes) and minimizes layout risk.
   */
  const anyRunIndex = runs.findIndex((r) => /<a:r\b/.test(r));
  if (anyRunIndex < 0) {
    throw new Error(
      "Strict template mismatch: label placeholder paragraph contains no <a:r> runs to clone for styling."
    );
  }

  const styleSeedRun = runs[anyRunIndex];

  // Extract <a:rPr ...>...</a:rPr> if present to preserve styling.
  const rPrXml = styleSeedRun.match(/<a:rPr\b[\s\S]*?<\/a:rPr>/)?.[0] ?? "";

  // Construct exactly one run with exactly one <a:t>.
  const insertedRun = `<a:r>${rPrXml}<a:t>${escapeXmlText(STATIC_LABEL_TEXT)}</a:t></a:r>`;

  // Insert the run into the existing paragraph WITHOUT creating new paragraphs.
  // We insert it just before </a:p> so we don't disturb existing leading nodes (e.g., <a:pPr>).
  const updatedParaXml = paraXml.replace(/<\/a:p>$/, `${insertedRun}</a:p>`);
  if (updatedParaXml === paraXml) {
    throw new Error(
      "Strict template mismatch: failed to insert the single label run (could not locate </a:p>)."
    );
  }

  const updatedScopeXml =
    scopeXml.slice(0, para.startInScope) +
    updatedParaXml +
    scopeXml.slice(para.endInScope);

  const updatedShapeXml =
    shape.xml.slice(0, scopeOffset) +
    updatedScopeXml +
    shape.xml.slice(scopeOffset + scopeXml.length);

  const updatedSlide1Xml =
    slide1Xml.slice(0, shape.start) + updatedShapeXml + slide1Xml.slice(shape.end);

  return { updatedSlide1Xml, changed: true };
}

/**
 * PUBLIC_INTERFACE
 * Updates ONLY the date field on slide 1 for the shipped default template,
 * while also ensuring the label placeholder above Name reads "TATA ELXSI" (Slide 1).
 *
 * Guarantees:
 * - slide1.xml is identical except for:
 *    - the label placeholder's text runs (enforced to "TATA ELXSI"), and
 *    - the targeted date <a:t> nodes.
 * - Does not alter any <a:rPr>, paragraph properties, shape geometry, or layout.
 * - Does not change any other files; last slide stays byte-identical.
 *
 * @param {ArrayBuffer} pptxArrayBuffer
 * @param {string} dateISO - value from <input type="date">
 * @returns {Promise<{updatedPptxBytes: Uint8Array, detected: {mode: string, slidePath: string, shapeName: string, shapeId: string}}>}
 */
export async function updatePptxDateOnly(pptxArrayBuffer, dateISO) {
  /** This is a public function. */
  const formattedDate = formatDateForTemplate(dateISO);
  if (!formattedDate) {
    throw new Error("Invalid date.");
  }

  const zip = await JSZip.loadAsync(pptxArrayBuffer);

  const slide1File = zip.file(SLIDE1_PATH);
  if (!slide1File) {
    throw new Error(`Template missing expected file: ${SLIDE1_PATH}`);
  }

  const slide1XmlOriginal = await slide1File.async("string");

  // Step 0: ensure the static label is placed in the existing placeholder above Name.
  const labelResult = setStaticLabelInExistingPlaceholder({ slide1Xml: slide1XmlOriginal });
  const slide1XmlWithLabel = labelResult.updatedSlide1Xml;

  // From here on, all date edits operate on slide1XmlWithLabel.
  const shapes = collectShapeBlocks(slide1XmlWithLabel);
  const dateLabelCandidates = shapes.filter((s) => s.xml.includes("Date</a:t>"));

  if (dateLabelCandidates.length < 1) {
    throw new Error(
      'Strict template mismatch: could not find any shape containing the "Date" label.'
    );
  }

  // The template uses real NBSP characters, not literal "\\u00a0".
  const expectedRunTexts = ["Date", " ", ":", "\u00a0 24\u00a0", "Dec", " ", "202", "5"];

  let shape = null;
  if (dateLabelCandidates.length === 1) {
    shape = dateLabelCandidates[0];
  } else {
    const matchesExpectedRuns = (shapeXml) => {
      const { paragraphs } = collectParagraphsFromShape(shapeXml);
      const p = paragraphs.find((x) => x.xml.includes("Date</a:t>"));
      if (!p) return false;

      const runs = collectRunsFromParagraph(p.xml).map((runXml) => ({
        tText: getFirstATextFromRun(runXml),
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
    throw new Error(
      'Strict template mismatch: could not find "Date" paragraph inside date shape.'
    );
  }

  const paragraphXml = paragraphs[paragraphIndex].xml;
  const runs = collectRunsFromParagraph(paragraphXml).map((runXml) => ({
    xml: runXml,
    tText: getFirstATextFromRun(runXml),
    hasText: /<a:t\b/.test(runXml),
  }));

  const actualRunTexts = runs.map((r) => (r.hasText ? r.tText : null));
  const sameLength = actualRunTexts.length === expectedRunTexts.length;
  const matches =
    sameLength && actualRunTexts.every((v, i) => v === expectedRunTexts[i]);

  if (!matches) {
    throw new Error(
      "Strict template mismatch: date paragraph runs differ from expected template; refusing to modify."
    );
  }

  const [day, mon, year] = formattedDate.split(" ");
  const y = String(year);
  const yHead = y.slice(0, 3);
  const yTail = y.slice(3);

  // Use real NBSP characters to preserve spacing semantics.
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
    scopeXml.slice(0, paraStartInScope) +
    updatedParagraphXml +
    scopeXml.slice(paraEndInScope);

  const updatedShapeXml =
    shape.xml.slice(0, scopeOffset) +
    updatedScopeXml +
    shape.xml.slice(scopeOffset + scopeXml.length);

  const updatedSlide1Xml =
    slide1XmlWithLabel.slice(0, shape.start) +
    updatedShapeXml +
    slide1XmlWithLabel.slice(shape.end);

  // Guard: allow only the 5 intended date <a:t> nodes to differ (relative to slide1XmlWithLabel).
  verifyOnlyAllowedSlide1DiffsByATextNodes({
    originalSlide1Xml: slide1XmlWithLabel,
    updatedSlide1Xml,
    dateParagraphXmlOriginal: paragraphXml,
    allowedDateRunIndexes: dateRunIndexes,
  });

  zip.file(SLIDE1_PATH, updatedSlide1Xml);

  const out = await zip.generateAsync({ type: "uint8array" });

  return {
    updatedPptxBytes: out,
    detected: {
      mode: labelResult.changed
        ? "strict-template/set-existing-label-placeholder+date-only-node-guard"
        : "strict-template/date-only-node-guard",
      slidePath: SLIDE1_PATH,
      shapeName,
      shapeId,
    },
  };
}

/**
 * PUBLIC_INTERFACE
 * Triggers a browser download of the PPTX bytes.
 * @param {Uint8Array} pptxBytes
 * @param {string} filename
 */
export function downloadPptxBytes(pptxBytes, filename) {
  /** This is a public function. */
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
  /** This is a public function. */
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
  /** This is a public function. */
  const d = new Date();
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

/**
 * PUBLIC_INTERFACE
 * Verifies the strict invariant: the last slide XML must remain byte-for-byte identical
 * between the original template PPTX and the updated PPTX output.
 *
 * @param {ArrayBuffer} originalPptxArrayBuffer
 * @param {Uint8Array} updatedPptxBytes
 * @returns {Promise<boolean>} true if unchanged, else throws Error
 */
export async function assertLastSlideUnchanged(originalPptxArrayBuffer, updatedPptxBytes) {
  /** This is a public function. */
  const originalZip = await JSZip.loadAsync(originalPptxArrayBuffer);
  const updatedZip = await JSZip.loadAsync(updatedPptxBytes);

  // Determine the "last slide" path from the original template at runtime.
  const slidePaths = originalZip
    .file(/^ppt\/slides\/slide\d+\.xml$/)
    .map((f) => f.name)
    .sort((a, b) => {
      const na = Number(a.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
      const nb = Number(b.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
      return na - nb;
    });

  if (!slidePaths.length) {
    throw new Error("Invariant check failed: original PPTX contains no slide XML parts.");
  }

  const LAST_SLIDE_PATH = slidePaths[slidePaths.length - 1];

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
