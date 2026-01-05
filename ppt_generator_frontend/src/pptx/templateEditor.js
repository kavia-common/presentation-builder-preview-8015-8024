import JSZip from "jszip";
import { saveAs } from "file-saver";

/**
 * STRICT TEMPLATE RULES (user requirements):
 * - The app ships with a built-in PPTX: `public/assets/template.pptx`.
 * - Slide 1 has fixed template layout; we may update ONLY:
 *    (a) the Slide 1 date text runs (editable).
 * - Do NOT add new runs/paragraphs/shapes anywhere.
 * - Do NOT alter any other slides, and preserve the last slide byte-for-byte.
 *
 * Implementation approach (byte-preserving):
 * - Read `ppt/slides/slide1.xml` as a string (do not parse/re-serialize XML).
 * - Update only the inner text of existing <a:t> nodes corresponding to the date runs.
 * - Do not modify any other file in the PPTX zip.
 *
 * IMPORTANT:
 * This module must never modify Slide 1 labels/names or any other text besides the date.
 */

const SLIDE1_PATH = "ppt/slides/slide1.xml";
const PRESENTATION_XML_PATH = "ppt/presentation.xml";
const PRESENTATION_RELS_XML_PATH = "ppt/_rels/presentation.xml.rels";

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

  // 2) Non-<a:t> parts must match exactly (no paragraph/shape/layout edits).
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
        "Safety check failed: slide1.xml modified in a non-date <a:t> node. Only date is editable."
      );
    }
  }
}

/**
 * PUBLIC_INTERFACE
 * Updates ONLY the date field on slide 1 for the shipped default template.
 *
 * Guarantees:
 * - slide1.xml is identical except for the targeted date <a:t> nodes.
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

  // Slide 1 edits operate only on slide1.xml.
  // 1) Update Date runs (editable) in-place.
  const shapesAfterLabel = collectShapeBlocks(slide1XmlOriginal);
  const dateLabelCandidates = shapesAfterLabel.filter((s) =>
    s.xml.includes("Date</a:t>")
  );

  if (dateLabelCandidates.length < 1) {
    throw new Error(
      'Strict template mismatch: could not find any shape containing the "Date" label.'
    );
  }

  // The template uses real NBSP characters, not literal "\\u00a0".
  const expectedRunTexts = [
    "Date",
    " ",
    ":",
    "\u00a0 24\u00a0",
    "Dec",
    " ",
    "202",
    "5",
  ];

  let shape = null;
  if (dateLabelCandidates.length === 1) {
    shape = dateLabelCandidates[0];
  } else {
    const matchesExpectedRuns = (shapeXml) => {
      const { paragraphs } = collectParagraphsFromShape(shapeXml);
      const p = paragraphs.find((x) => x.xml.includes("Date</a:t>"));
      if (!p) return false;

      const runsInner = collectRunsFromParagraph(p.xml).map((runXml) => ({
        tText: getFirstATextFromRun(runXml),
        hasText: /<a:t\b/.test(runXml),
      }));

      const actualRunTexts = runsInner.map((r) => (r.hasText ? r.tText : null));
      return (
        actualRunTexts.length === expectedRunTexts.length &&
        actualRunTexts.every((v, i) => v === expectedRunTexts[i])
      );
    };

    shape = dateLabelCandidates.find((c) => matchesExpectedRuns(c.xml)) ?? null;
  }

  if (!shape) {
    const names = shapesAfterLabel
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
    slide1XmlOriginal.slice(0, shape.start) +
    updatedShapeXml +
    slide1XmlOriginal.slice(shape.end);

  // Guard: allow only the 5 intended date <a:t> nodes.
  // Everything else in slide1.xml must remain byte-for-byte identical (structure preserved).
  verifyOnlyAllowedSlide1DiffsByATextNodes({
    originalSlide1Xml: slide1XmlOriginal,
    updatedSlide1Xml,
    dateParagraphXmlOriginal: paragraphXml,
    allowedDateRunIndexes: dateRunIndexes,
  });

  zip.file(SLIDE1_PATH, updatedSlide1Xml);

  const out = await zip.generateAsync({ type: "uint8array" });

  return {
    updatedPptxBytes: out,
    detected: {
      mode: "strict-template/date-only",
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
      `Invariant check failed: missing ${LAST_SLIDE_PATH} in ${
        !orig ? "original" : "updated"
      } PPTX.`
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

/**
 * Finds slide XML part names and returns sorted slide numbers and paths.
 */
function listSlideXmlPaths(zip) {
  const slidePaths = zip.file(/^ppt\/slides\/slide\d+\.xml$/).map((f) => f.name);
  const parsed = slidePaths
    .map((p) => ({ path: p, n: Number(p.match(/slide(\d+)\.xml$/)?.[1] ?? 0) }))
    .filter((x) => Number.isFinite(x.n) && x.n > 0)
    .sort((a, b) => a.n - b.n);

  return parsed;
}

function parseRelationships(relsXml) {
  const relRe =
    /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bType="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/>/g;
  const rels = [];
  let m;
  while ((m = relRe.exec(relsXml)) !== null) {
    rels.push({ id: m[1], type: m[2], target: m[3], raw: m[0] });
  }
  return rels;
}

function replaceOnce(haystack, needle, replacement) {
  const idx = haystack.indexOf(needle);
  if (idx < 0) return null;
  return haystack.slice(0, idx) + replacement + haystack.slice(idx + needle.length);
}

/**
 * Extracts rId list from <p:sldIdLst> and returns:
 * - sldIdLstXml: full XML of the list node (string)
 * - entries: [{ id: number, rId: string, raw: string }]
 */
function parseSldIdList(presentationXml) {
  const listMatch = presentationXml.match(/<p:sldIdLst\b[\s\S]*?<\/p:sldIdLst>/);
  if (!listMatch) {
    throw new Error("presentation.xml missing <p:sldIdLst>.");
  }

  const sldIdLstXml = listMatch[0];
  const entryRe = /<p:sldId\b[^>]*\/>/g;
  const entries = [];
  let m;
  while ((m = entryRe.exec(sldIdLstXml)) !== null) {
    const raw = m[0];
    const id = Number(raw.match(/\bid="(\d+)"/)?.[1] ?? 0);
    const rId = raw.match(/\br:id="([^"]+)"/)?.[1] ?? "";
    if (Number.isFinite(id) && id > 0 && rId) {
      entries.push({ id, rId, raw });
    }
  }

  return { sldIdLstXml, entries };
}

/**
 * PUBLIC_INTERFACE
 * Prunes the PPTX to only include Slide 1 and the template's last slide (by deck order),
 * while keeping the *last slide XML bytes* untouched.
 *
 * Why this exists:
 * - The product requirement is a two-slide default deck (Slide 1 + last slide).
 * - Slide 1 is still date-only editable (handled elsewhere).
 * - The last slide must remain byte-for-byte identical to the template:
 *   we therefore do NOT touch its slideN.xml part at all.
 *
 * What this function changes:
 * - Updates ppt/presentation.xml <p:sldIdLst> to only keep the first and last entries.
 * - Updates ppt/_rels/presentation.xml.rels to remove unreferenced slide relationships.
 *
 * What this function intentionally does NOT do:
 * - It does NOT delete intermediate slide parts from the zip; they are simply no longer referenced.
 *   This is safer for invariants and avoids touching unrelated bytes.
 *
 * @param {Uint8Array} pptxBytes - PPTX bytes after date-only edit
 * @returns {Promise<{ updatedPptxBytes: Uint8Array, kept: { firstSlideNumber:number, lastSlideNumber:number } }>}
 */
export async function prunePptxToFirstAndLastSlides(pptxBytes) {
  /** This is a public function. */
  if (!pptxBytes || !pptxBytes.length) throw new Error("No PPTX bytes provided.");

  const zip = await JSZip.loadAsync(pptxBytes);

  const presFile = zip.file(PRESENTATION_XML_PATH);
  const presRelsFile = zip.file(PRESENTATION_RELS_XML_PATH);
  if (!presFile) throw new Error(`Missing required PPTX part: ${PRESENTATION_XML_PATH}`);
  if (!presRelsFile) throw new Error(`Missing required PPTX part: ${PRESENTATION_RELS_XML_PATH}`);

  const presXmlOriginal = await presFile.async("string");
  const { sldIdLstXml, entries } = parseSldIdList(presXmlOriginal);

  if (entries.length < 2) {
    // Nothing to prune.
    return {
      updatedPptxBytes: new Uint8Array(pptxBytes),
      kept: { firstSlideNumber: 1, lastSlideNumber: entries.length === 1 ? 1 : 0 },
    };
  }

  const firstEntry = entries[0];
  const lastEntry = entries[entries.length - 1];

  // Best-effort slide numbers (used only for debug/reporting, not correctness).
  // Map rId -> slideN by reading presentation.xml.rels targets.
  const presRelsXmlOriginal = await presRelsFile.async("string");
  const rels = parseRelationships(presRelsXmlOriginal);
  const ridToTarget = new Map(rels.map((r) => [r.id, r.target]));

  const toSlideNumber = (rid) => {
    const tgt = ridToTarget.get(rid) || "";
    const m = tgt.match(/slides\/slide(\d+)\.xml$/);
    const n = Number(m?.[1] ?? 0);
    return Number.isFinite(n) ? n : 0;
  };

  const firstSlideNumber = toSlideNumber(firstEntry.rId) || 1;
  const lastSlideNumber = toSlideNumber(lastEntry.rId) || 0;

  // 1) Update presentation.xml: keep only first+last nodes.
  const newSldIdLstXml = (() => {
    // Preserve original wrapper (<p:sldIdLst ...> ... </p:sldIdLst>) and inject only the two nodes.
    const openTag = sldIdLstXml.match(/^<p:sldIdLst\b[\s\S]*?>/)?.[0];
    const closeTag = "</p:sldIdLst>";
    if (!openTag || !sldIdLstXml.endsWith(closeTag)) {
      throw new Error("presentation.xml <p:sldIdLst> parse failed.");
    }
    return `${openTag}${firstEntry.raw}${lastEntry.raw}${closeTag}`;
  })();

  const presXmlUpdated = (() => {
    const replaced = replaceOnce(presXmlOriginal, sldIdLstXml, newSldIdLstXml);
    if (!replaced) throw new Error("Failed to update presentation.xml <p:sldIdLst>.");
    return replaced;
  })();

  zip.file(PRESENTATION_XML_PATH, presXmlUpdated);

  // 2) Update presentation.xml.rels: remove slide relationships not referenced by the kept rIds.
  // Keep *all other* relationship types unchanged.
  const keptRids = new Set([firstEntry.rId, lastEntry.rId]);
  const filteredRelsXml = (() => {
    // Remove any <Relationship ...Type=".../slide"...> whose Id is not in keptRids.
    // Do not attempt full reformatting; just strip the nodes.
    return presRelsXmlOriginal.replace(
      /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bType="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/slide"[^>]*\/>/g,
      (full, rid) => (keptRids.has(rid) ? full : "")
    );
  })();

  zip.file(PRESENTATION_RELS_XML_PATH, filteredRelsXml);

  const out = await zip.generateAsync({ type: "uint8array" });

  return {
    updatedPptxBytes: out,
    kept: { firstSlideNumber, lastSlideNumber },
  };
}
