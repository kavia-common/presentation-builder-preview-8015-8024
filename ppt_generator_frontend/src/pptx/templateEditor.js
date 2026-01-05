import JSZip from "jszip";
import { saveAs } from "file-saver";

/**
 * STRICT TEMPLATE RULES (user requirements):
 * - The app ships with a built-in PPTX: `public/assets/template.pptx`.
 * - Slide 1 has fixed template layout; we may update ONLY:
 *    (a) the Slide 1 date text runs (editable), and
 *    (b) the label directly above the Name on Slide 1 to "TATA ELXSI" (locked/fixed).
 * - Do NOT add new runs/paragraphs/shapes anywhere.
 * - Do NOT alter any other slides, and preserve the last slide byte-for-byte.
 *
 * Implementation approach (byte-preserving):
 * - Read `ppt/slides/slide1.xml` as a string (do not parse/re-serialize XML).
 * - Update only the inner text of existing <a:t> nodes:
 *    - 5 date runs (day/month/year splits)
 *    - the existing label text-runs above Name (previously "TATA") updated in-place
 * - Update ONLY paragraph properties for that existing label paragraph to match the
 *   screenshot (centered alignment). No new shapes/paragraphs/runs are created.
 * - Do not modify any other file in the PPTX zip.
 *
 * IMPORTANT:
 * The "only date editable" constraint is preserved in the UI. The label update is
 * forced on generation to match the template requirement, but the mutation remains
 * extremely narrow: only specific existing <a:t> nodes are changed, and no new XML
 * nodes are introduced.
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
 * Replaces the Slide 1 label directly above the Name (in-place).
 *
 * Template structure (bundled template):
 * - There is a single paragraph with a single <a:r> whose <a:t> is "TATA"
 *   (shown above the "Name : Subrata B" line in the screenshot).
 *
 * Requirements:
 * - Replace only existing <a:t> inner text (no new runs/paragraphs/shapes).
 * - Copy style/spacing from template: by editing text only, we preserve font,
 *   weight, size, color, alignment, and positioning.
 *
 * Returns:
 * - updatedShapeXml
 * - allowedATextNodePairs for safety guard
 */
function replaceSlide1LabelAboveNameInShape(shapeXml, targetValue) {
  const { scopeXml, scopeOffset, paragraphs } = collectParagraphsFromShape(shapeXml);
  if (!paragraphs.length) {
    throw new Error("Strict template mismatch: label shape contains no paragraphs.");
  }

  // Strictly find the paragraph that contains exactly the "TATA" <a:t>.
  // We keep this conservative to avoid accidentally touching other text nodes.
  const paragraphIndex = paragraphs.findIndex((p) =>
    p.xml.includes("<a:t>TATA</a:t>")
  );
  if (paragraphIndex < 0) {
    throw new Error(
      'Strict template mismatch: could not find the expected "TATA" label paragraph in slide 1.'
    );
  }

  const paragraphXmlOriginal = paragraphs[paragraphIndex].xml;

  // Ensure the label paragraph is centered like the screenshot (TATA ELXSI centered above Name).
  // IMPORTANT: We do not create a new paragraph; we only edit the existing <a:pPr> (or inject one
  // if missing) to set algn="ctr" and keep everything else byte-identical outside <a:t> and this
  // label paragraph’s <a:pPr>.
  const enforceCenteredParagraph = (pXml) => {
    const hasPPr = /<a:pPr\b/.test(pXml);

    if (hasPPr) {
      // Modify existing a:pPr tag in-place, preserving other attributes/children.
      // If algn is already present, replace its value; otherwise inject algn="ctr" into the tag.
      return pXml.replace(/<a:pPr\b([^>]*)>/, (full, attrs) => {
        if (/\balgn="/.test(attrs)) {
          const nextAttrs = attrs.replace(/\balgn="[^"]*"/, 'algn="ctr"');
          return `<a:pPr${nextAttrs}>`;
        }
        return `<a:pPr${attrs} algn="ctr">`;
      });
    }

    // No a:pPr: inject a minimal one immediately after <a:p ...>
    // This does not create a new paragraph; it only adds paragraph properties to the existing one.
    return pXml.replace(/<a:p\b([^>]*)>/, `<a:p$1><a:pPr algn="ctr"/>`);
  };

  const paragraphXml = enforceCenteredParagraph(paragraphXmlOriginal);

  const runs = collectRunsFromParagraph(paragraphXml).map((runXml) => ({
    xml: runXml,
    tText: getFirstATextFromRun(runXml),
    hasText: /<a:t\b/.test(runXml),
  }));

  // Expect exactly one run with a:t="TATA" for this label paragraph.
  const textRuns = runs.filter((r) => r.hasText);
  const isStrictLabel =
    textRuns.length === 1 && textRuns[0].tText === "TATA" && runs.length >= 1;

  if (!isStrictLabel) {
    throw new Error(
      'Strict template mismatch: label paragraph runs differ from expected template; refusing to modify.'
    );
  }

  const updatedRunsXml = runs.map((r) => r.xml);
  const originalNode = getFirstATextNodeFromRun(textRuns[0].xml);
  if (!originalNode) {
    throw new Error(
      "Strict template mismatch: could not locate <a:t> node for label run."
    );
  }

  const newValueEscaped = escapeXmlText(targetValue);
  const updatedNode = originalNode.replace(
    /(<a:t\b[^>]*>)([\s\S]*?)(<\/a:t>)/,
    `$1${newValueEscaped}$3`
  );

  // Replace within the matching run.
  const runIdx = runs.findIndex((r) => r.xml === textRuns[0].xml);
  updatedRunsXml[runIdx] = updatedRunsXml[runIdx].replace(originalNode, updatedNode);

  // Rebuild paragraph preserving exact non-run content around the run region.
  const firstRunIdx = paragraphXml.search(/<a:r\b/);
  const lastRunEnd = paragraphXml.lastIndexOf("</a:r>");
  if (firstRunIdx < 0 || lastRunEnd < 0) {
    throw new Error(
      "Strict template mismatch: could not parse label paragraph runs."
    );
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
    allowedATextNodePairs: [{ originalNode, updatedNode }],
  };
}

/**
 * Verifies that differences between originalSlide1Xml and updatedSlide1Xml
 * occur ONLY within the inner text of the intended date <a:t> nodes on slide 1
 * plus any explicitly allowed additional <a:t> node replacements.
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

  // 2) Non-<a:t> parts must match exactly, EXCEPT we allow a very specific
  // label-paragraph <a:pPr> alignment change (to center) for the "TATA" label paragraph.
  //
  // We implement this by normalizing both XML strings to remove/standardize:
  // - the <a:pPr ...> inside the paragraph that contains "<a:t>TATA</a:t>"
  // Then we compare the remaining structure parts byte-for-byte.
  const normalizeForAllowedLabelPPr = (xml) => {
    // Find the paragraph containing the label node by anchoring on the (pre-replacement) text.
    // This ensures we only relax matching around the intended label paragraph properties.
    const paraRe = /<a:p\b[\s\S]*?<\/a:p>/g;
    const paras = xml.match(paraRe) ?? [];
    const labelPara = paras.find((p) => p.includes("<a:t>TATA</a:t>")) ?? null;
    if (!labelPara) return xml;

    const stripped = labelPara.replace(/<a:pPr\b[\s\S]*?<\/a:pPr>/g, "");
    return xml.replace(labelPara, stripped);
  };

  const normOrig = normalizeForAllowedLabelPPr(originalSlide1Xml);
  const normNext = normalizeForAllowedLabelPPr(updatedSlide1Xml);

  const origParts = normOrig.split(aTextRe);
  const nextParts = normNext.split(aTextRe);
  if (origParts.length !== nextParts.length) {
    throw new Error(
      "Safety check failed: slide1.xml structure changed (unexpected <a:t> segmentation)."
    );
  }
  for (let i = 0; i < origParts.length; i += 1) {
    if (origParts[i] !== nextParts[i]) {
      throw new Error(
        "Safety check failed: slide1.xml changed outside <a:t> nodes (and outside the allowed label paragraph <a:pPr>). Only date/allowed text may change."
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

  // Allow: additional <a:t> nodes explicitly expected to change (label above Name).
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
        "Safety check failed: slide1.xml modified in a non-date/non-allowed <a:t> node. Only date is editable; label above Name is fixed to a constant."
      );
    }
  }
}

/**
 * PUBLIC_INTERFACE
 * Updates ONLY the date field on slide 1 for the shipped default template,
 * and also forces the label above Name to "TATA ELXSI" (in-place).
 *
 * Guarantees:
 * - slide1.xml is identical except for:
 *   - the targeted date <a:t> nodes, and
 *   - the label-above-Name <a:t> node updated in-place ("TATA" -> "TATA ELXSI").
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
  // 1) Fix label above Name to "TATA ELXSI" in-place (no new runs/paragraphs/shapes).
  const shapes = collectShapeBlocks(slide1XmlOriginal);
  const labelCandidates = shapes.filter((s) => s.xml.includes("<a:t>TATA</a:t>"));
  if (labelCandidates.length !== 1) {
    throw new Error(
      'Strict template mismatch: could not uniquely locate the Slide 1 label above Name (expected a single "TATA" label node).'
    );
  }

  const labelShape = labelCandidates[0];
  const {
    updatedShapeXml: updatedLabelShapeXml,
    allowedATextNodePairs: allowedLabelATextNodePairs,
  } = replaceSlide1LabelAboveNameInShape(labelShape.xml, "TATA ELXSI");

  const slide1XmlAfterLabel =
    slide1XmlOriginal.slice(0, labelShape.start) +
    updatedLabelShapeXml +
    slide1XmlOriginal.slice(labelShape.end);

  // 2) Update Date runs (editable) in-place.
  const shapesAfterLabel = collectShapeBlocks(slide1XmlAfterLabel);
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
    slide1XmlAfterLabel.slice(0, shape.start) +
    updatedShapeXml +
    slide1XmlAfterLabel.slice(shape.end);

  // Guard: allow only the 5 intended date <a:t> nodes AND the label node.
  // Everything else in slide1.xml must remain byte-for-byte identical (structure preserved).
  verifyOnlyAllowedSlide1DiffsByATextNodes({
    originalSlide1Xml: slide1XmlOriginal,
    updatedSlide1Xml,
    dateParagraphXmlOriginal: paragraphXml,
    allowedDateRunIndexes: dateRunIndexes,
    allowedAdditionalATextNodePairs: allowedLabelATextNodePairs,
  });

  zip.file(SLIDE1_PATH, updatedSlide1Xml);

  const out = await zip.generateAsync({ type: "uint8array" });

  return {
    updatedPptxBytes: out,
    detected: {
      mode: "strict-template/date-only + fixed-label-above-name-node-guard",
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
