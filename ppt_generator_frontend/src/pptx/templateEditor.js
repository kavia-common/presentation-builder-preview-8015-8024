import JSZip from "jszip";
import { saveAs } from "file-saver";

/**
 * STRICT TEMPLATE RULES (user requirements):
 * - The app ships with a built-in PPTX: `public/assets/template.pptx`.
 * - Slide 1 has fixed template layout; we may update:
 *    (a) the Slide 1 date text runs (editable), and
 *    (b) the existing Slide 1 Name VALUE text to "TATA ELXSI" (locked/fixed).
 * - Do NOT add new runs/paragraphs/shapes anywhere.
 * - Do NOT alter any other slides, and preserve the last slide byte-for-byte.
 *
 * Implementation approach (byte-preserving):
 * - Read `ppt/slides/slide1.xml` as a string (do not parse/re-serialize XML).
 * - Update only the inner text of existing <a:t> nodes:
 *    - 5 date runs (day/month/year splits)
 *    - the existing Name value runs (previously "Subrata B") inside the Name paragraph
 * - Do not modify any other file in the PPTX zip.
 *
 * IMPORTANT:
 * The "only date editable" constraint is preserved in the UI. The Name value is
 * forcibly set to a fixed string on generation to match the template requirement,
 * but we still keep the mutation extremely narrow: only specific existing <a:t>
 * nodes are changed, and no new XML nodes are introduced.
 */

const SLIDE1_PATH = "ppt/slides/slide1.xml";

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
 * Replaces the Slide 1 "Name" value runs in-place.
 *
 * Template structure (bundled template):
 * Name paragraph contains runs: "Name", " ", ":", " ", "Subrata", " ", "B"
 *
 * Requirements:
 * - Do not add runs/paragraphs/shapes.
 * - Replace the existing value text content directly.
 * - Keep the label "Name :" unchanged (no new label above it).
 *
 * Returns:
 * - updatedShapeXml: shape.xml with updated paragraph content
 * - allowedATextNodePairs: [{ originalNode, updatedNode }, ...] for safety guard
 */
function replaceSlide1NameValueInShape(shapeXml, targetValue) {
  const { scopeXml, scopeOffset, paragraphs } = collectParagraphsFromShape(shapeXml);
  if (!paragraphs.length) {
    throw new Error("Strict template mismatch: name shape contains no paragraphs.");
  }

  const paragraphIndex = paragraphs.findIndex(
    (p) => p.xml.includes("Name</a:t>") && p.xml.includes("<a:t>Subrata</a:t>")
  );
  if (paragraphIndex < 0) {
    throw new Error(
      'Strict template mismatch: could not find the expected "Name" paragraph with the original value in slide 1.'
    );
  }

  const paragraphXml = paragraphs[paragraphIndex].xml;
  const runs = collectRunsFromParagraph(paragraphXml).map((runXml) => ({
    xml: runXml,
    tText: getFirstATextFromRun(runXml),
    hasText: /<a:t\b/.test(runXml),
  }));

  const expectedRunTexts = ["Name", " ", ":", " ", "Subrata", " ", "B"];
  const actualRunTexts = runs.map((r) => (r.hasText ? r.tText : null));
  const matches =
    actualRunTexts.length === expectedRunTexts.length &&
    actualRunTexts.every((v, i) => v === expectedRunTexts[i]);

  if (!matches) {
    throw new Error(
      "Strict template mismatch: Name paragraph runs differ from expected template; refusing to modify."
    );
  }

  // We must not add runs. Encode the full target value into the existing "Subrata" run
  // and blank out the trailing "B" run to avoid leftover characters.
  const updatedRunsXml = runs.map((r) => r.xml);

  const originalNodeSubrata = getFirstATextNodeFromRun(updatedRunsXml[4]);
  const originalNodeB = getFirstATextNodeFromRun(updatedRunsXml[6]);
  if (!originalNodeSubrata || !originalNodeB) {
    throw new Error(
      "Strict template mismatch: could not locate <a:t> nodes for Name value runs."
    );
  }

  const newValueEscaped = escapeXmlText(targetValue);
  const updatedNodeSubrata = originalNodeSubrata.replace(
    /(<a:t\b[^>]*>)([\s\S]*?)(<\/a:t>)/,
    `$1${newValueEscaped}$3`
  );
  const updatedNodeB = originalNodeB.replace(
    /(<a:t\b[^>]*>)([\s\S]*?)(<\/a:t>)/,
    `$1$3`
  );

  updatedRunsXml[4] = updatedRunsXml[4].replace(
    originalNodeSubrata,
    updatedNodeSubrata
  );
  updatedRunsXml[6] = updatedRunsXml[6].replace(originalNodeB, updatedNodeB);

  // Rebuild paragraph preserving exact non-run content around the run region.
  const firstRunIdx = paragraphXml.search(/<a:r\b/);
  const lastRunEnd = paragraphXml.lastIndexOf("</a:r>");
  if (firstRunIdx < 0 || lastRunEnd < 0) {
    throw new Error("Strict template mismatch: could not parse Name paragraph runs.");
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
    shapeXml.slice(0, scopeOffset) +
    updatedScopeXml +
    shapeXml.slice(scopeOffset + scopeXml.length);

  return {
    updatedShapeXml,
    allowedATextNodePairs: [
      { originalNode: originalNodeSubrata, updatedNode: updatedNodeSubrata },
      { originalNode: originalNodeB, updatedNode: updatedNodeB },
    ],
  };
}

/**
 * Verifies that differences between originalSlide1Xml and updatedSlide1Xml
 * occur ONLY within the inner text of the intended date <a:t> nodes on slide 1
 * plus any explicitly allowed additional <a:t> node replacements (e.g., fixed Name value).
 *
 * This avoids brittle positional diffs (day can change 2 digits -> 1 digit, shifting
 * subsequent characters and causing false positives).
 */
function verifyOnlyAllowedSlide1DiffsByATextNodes({
  originalSlide1Xml,
  updatedSlide1Xml,
  dateParagraphXmlOriginal,
  allowedDateRunIndexes,
  allowedAdditionalATextNodePairs = [],
}) {
  if (originalSlide1Xml === updatedSlide1Xml) return;

  const aTextRe = /<a:t\b[^>]*>[\s\S]*?<\/a:t>/g;

  // 1) Structure check: same number of <a:t> nodes.
  const origNodes = originalSlide1Xml.match(aTextRe) ?? [];
  const nextNodes = updatedSlide1Xml.match(aTextRe) ?? [];
  if (origNodes.length !== nextNodes.length) {
    throw new Error(
      "Safety check failed: slide1.xml <a:t> node count changed. Only date/allowed text nodes may change."
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
        "Safety check failed: slide1.xml changed outside <a:t> nodes. Only date/allowed text may change."
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

  // Allow: the date nodes we are editing.
  for (const node of allowedLocalATextNodes) {
    const q = queues.get(node) ?? [];
    if (!q.length) {
      throw new Error(
        "Safety check failed: could not map date <a:t> node to global index in slide1.xml."
      );
    }
    allowedGlobalIndexes.add(q.shift());
  }

  // Allow: additional <a:t> nodes explicitly expected to change.
  for (const pair of allowedAdditionalATextNodePairs) {
    const q = queues.get(pair.originalNode) ?? [];
    if (!q.length) {
      throw new Error(
        "Safety check failed: could not map allowed <a:t> node to global index in slide1.xml."
      );
    }
    const globalIdx = q.shift();
    allowedGlobalIndexes.add(globalIdx);

    if (nextNodes[globalIdx] !== pair.updatedNode) {
      throw new Error(
        "Safety check failed: allowed <a:t> node did not match expected updated value."
      );
    }
  }

  // 4) All <a:t> nodes except the allowed ones must be identical byte-for-byte.
  for (let i = 0; i < origNodes.length; i += 1) {
    if (allowedGlobalIndexes.has(i)) continue;
    if (origNodes[i] !== nextNodes[i]) {
      throw new Error(
        "Safety check failed: slide1.xml modified in a non-date/non-allowed <a:t> node. Only date is editable; Name is fixed to a constant."
      );
    }
  }
}

/**
 * PUBLIC_INTERFACE
 * Updates ONLY the date field on slide 1 for the shipped default template,
 * and also forces the existing Slide 1 Name value to "TATA ELXSI" (in-place).
 *
 * Guarantees:
 * - slide1.xml is identical except for:
 *   - the targeted date <a:t> nodes, and
 *   - the existing Name value <a:t> nodes ("Subrata" and "B") updated in-place.
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
  // 1) Fix Name value to "TATA ELXSI" in-place (no new runs/paragraphs/shapes).
  const shapes = collectShapeBlocks(slide1XmlOriginal);
  const nameCandidates = shapes.filter(
    (s) => s.xml.includes("Name</a:t>") && s.xml.includes("<a:t>Subrata</a:t>")
  );
  if (nameCandidates.length !== 1) {
    throw new Error(
      'Strict template mismatch: could not uniquely locate the Slide 1 Name field/value shape.'
    );
  }

  const nameShape = nameCandidates[0];
  const {
    updatedShapeXml: updatedNameShapeXml,
    allowedATextNodePairs: allowedNameATextNodePairs,
  } = replaceSlide1NameValueInShape(nameShape.xml, "TATA ELXSI");

  const slide1XmlAfterName =
    slide1XmlOriginal.slice(0, nameShape.start) +
    updatedNameShapeXml +
    slide1XmlOriginal.slice(nameShape.end);

  // 2) Update Date runs (editable) in-place.
  const shapesAfterName = collectShapeBlocks(slide1XmlAfterName);
  const dateLabelCandidates = shapesAfterName.filter((s) =>
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
    const names = shapesAfterName
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
  const replacements = [`\u00a0 ${day}\u00a0`, mon, " ", yHead, yTail].map(
    escapeXmlText
  );

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
    slide1XmlAfterName.slice(0, shape.start) +
    updatedShapeXml +
    slide1XmlAfterName.slice(shape.end);

  // Guard: allow only the 5 intended date <a:t> nodes AND the two Name-value nodes.
  // Everything else in slide1.xml must remain byte-for-byte identical (structure preserved).
  verifyOnlyAllowedSlide1DiffsByATextNodes({
    originalSlide1Xml: slide1XmlOriginal,
    updatedSlide1Xml,
    dateParagraphXmlOriginal: paragraphXml,
    allowedDateRunIndexes: dateRunIndexes,
    allowedAdditionalATextNodePairs: allowedNameATextNodePairs,
  });

  zip.file(SLIDE1_PATH, updatedSlide1Xml);

  const out = await zip.generateAsync({ type: "uint8array" });

  return {
    updatedPptxBytes: out,
    detected: {
      mode: "strict-template/date-only + fixed-name-node-guard",
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
export async function assertLastSlideUnchanged(
  originalPptxArrayBuffer,
  updatedPptxBytes
) {
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
    throw new Error(
      "Invariant check failed: original PPTX contains no slide XML parts."
    );
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
      throw new Error(
        `Invariant check failed: last slide bytes differ at offset ${i}.`
      );
    }
  }

  return true;
}
