import JSZip from "jszip";
import { saveAs } from "file-saver";

/**
 * STRICT TEMPLATE RULES (user requirements):
 * - This file implements only SAFE in-place edits for slide 1 of the template PPTX.
 * - Updates ONLY:
 *   (a) The label directly above the Name to EXACTLY 'TATA ELXSI' on Slide 1, in-place (no new shapes/runs)
 *   (b) The date text, in-place, using existing runs only.
 *   (c) Ensures only the date field is editable on Slide 1; all other shapes are locked.
 * - No paragraph/shape additions; last slide untouched (byte-identical).
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
 */
function formatDateForTemplate(dateInput) {
  const d = new Date(dateInput);
  if (Number.isNaN(d.getTime())) return "";
  const day = d.getDate();
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

// --- Helper functions for label/shape identification (regex-based, no new objects) ---

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

// This will find all <a:t> nodes in run order (used for label edits)
function getAllTextNodesInShapeXml(shapeXml) {
  return [...shapeXml.matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g)];
}

function getShapeName(shapeXml) {
  const m = shapeXml.match(/<p:cNvPr\b[^>]*\bname="([^"]*)"/);
  return m ? m[1] : null;
}

/**
 * PUBLIC_INTERFACE
 * Edits the Slide 1 label (directly above Name) to "TATA ELXSI" in-place,
 * with in-run XML replacement. Only the label text content changes; all other
 * spacing, style, and template content is left as-is. If not found, throws.
 * @param {ArrayBuffer} pptxArrayBuffer
 * @returns {Promise<{updatedPptxBytes: Uint8Array, detected: object}>}
 */
export async function updateSlide1LabelInPlace(pptxArrayBuffer) {
  /** This is a public function. */
  const zip = await JSZip.loadAsync(pptxArrayBuffer);

  const slide1File = zip.file(SLIDE1_PATH);
  if (!slide1File) {
    throw new Error(`Template missing expected file: ${SLIDE1_PATH}`);
  }
  let slide1Xml = await slide1File.async("string");

  // --- Find the correct label (directly above Name, not date field)
  // Strategy: Find all shapes with <a:t> not matching Name, not matching date label, and close to top of deck.
  const shapes = collectShapeBlocks(slide1Xml);
  let labelShape = null, labelParaIdx = -1, labelTextNodeIdx = -1;
  let labelFound = false;

  // Helper: test if <a:t> candidate matches Name/Date
  function isNameOrDate(text) {
    if (!text) return false;
    const t = text.trim();
    return /Name|Your Name|John Doe|Date|[0-9]{1,2} [A-Za-z]{3,9} \d{4}|^[0-9]{4}-[0-9]{2}-[0-9]{2}$/.test(t);
  }

  // Heuristic: the top-most shape with a single non-empty <a:t> above the Name
  shapes.some(shape => {
    const shapeXml = shape.xml;
    const { paragraphs } = collectParagraphsFromShape(shapeXml);
    for (let i = 0; i < paragraphs.length; ++i) {
      const paraXml = paragraphs[i].xml;
      const textMatches = [...paraXml.matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g)];
      for (let j = 0; j < textMatches.length; ++j) {
        const textVal = textMatches[j][1];
        // Not a date, not a Name, not empty, not just a colon, and not the existing company name (if present)
        if (
          textVal &&
          !isNameOrDate(textVal) &&
          !/^:?$/.test(textVal.trim())  // disfavors colons placed as their own run
        ) {
          // Additional guard: avoid replacing styles that come with only whitespace/spacing (should have letters)
          if (/[A-Za-z]/.test(textVal)) {
            // found!
            labelShape = shape;
            labelParaIdx = i;
            labelTextNodeIdx = j;
            labelFound = true;
            return true;
          }
        }
      }
      if (labelFound) break;
    }
    return labelFound;
  });

  if (!labelShape) {
    throw new Error("Could not find label shape above Name for in-place update.");
  }

  // Do the XML-level replacement for that <a:t> ONLY (preserve whitespace around or inside as needed)
  let newShapeXml = labelShape.xml;
  const { paragraphs } = collectParagraphsFromShape(labelShape.xml);
  const paraXml = paragraphs[labelParaIdx].xml;
  let newParaXml = paraXml;
  const textRuns = [...paraXml.matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g)];

  if (!textRuns[labelTextNodeIdx]) {
    throw new Error("In-place label text node index could not be resolved for label shape.");
  }

  // Exact replacement: keep leading/trailing whitespace; replace only central label
  const oldText = textRuns[labelTextNodeIdx][1];
  const labelPattern = /^(\\s*)(.*?)(\\s*)$/;
  const match = oldText.match(labelPattern);
  const leading = match ? match[1] : "";
  const trailing = match ? match[3] : "";

  const newLabelXml = `<a:t>${leading}TATA ELXSI${trailing}</a:t>`;

  // Replace that <a:t> in this paragraph (by offset):
  let replaceIdx = 0, realIdx = -1;
  // Find the text run's match globally inside paraXml (to disambiguate multiples with similar text)
  paraXml.replace(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g, (full, txt, start) => {
    if (replaceIdx === labelTextNodeIdx) {
      realIdx = start;
    }
    replaceIdx++;
    return full;
  });

  if (realIdx === -1) {
    // fallback, do the text-replace using regex with proper position control.
    let count = 0;
    newParaXml = paraXml.replace(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g, (full, txt) => {
      if (count === labelTextNodeIdx) {
        count++;
        return newLabelXml;
      } else {
        count++;
        return full;
      }
    });
  } else {
    // Direct replacement by string slicing:
    newParaXml =
      paraXml.slice(0, realIdx) +
      newLabelXml +
      paraXml.slice(realIdx + textRuns[labelTextNodeIdx][0].length);
  }

  // Patch the paragraph back into the shape:
  const paraObj = paragraphs[labelParaIdx];
  let scopeXml = collectParagraphsFromShape(labelShape.xml).scopeXml;
  let newScopeXml =
    scopeXml.slice(0, paraObj.startInScope) +
    newParaXml +
    scopeXml.slice(paraObj.endInScope);

  // Patch txBody back into <p:sp>
  const scopeOffset = collectParagraphsFromShape(labelShape.xml).scopeOffset;
  newShapeXml =
    labelShape.xml.slice(0, scopeOffset) +
    newScopeXml +
    labelShape.xml.slice(scopeOffset + scopeXml.length);

  // Replace entire shape in slide1Xml:
  slide1Xml =
    slide1Xml.slice(0, labelShape.start) +
    newShapeXml +
    slide1Xml.slice(labelShape.end);

  // Write the edited XML back into the PPTX for slide 1
  zip.file(SLIDE1_PATH, slide1Xml);

  const out = await zip.generateAsync({ type: "uint8array" });

  return {
    updatedPptxBytes: out,
    detected: {
      mode: "strict-template/label-inplace",
      slidePath: SLIDE1_PATH,
      shapeName: getShapeName(labelShape.xml) ?? "unknown",
    },
  };
}

/**
 * PUBLIC_INTERFACE
 * Updates ONLY the date field on slide 1 for the shipped default template.
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

  // Find date shape by looking for "Date" text run
  const shapes = collectShapeBlocks(slide1XmlOriginal);
  const dateShape = shapes.find(s =>
    s.xml.includes("Date</a:t>")
  );

  if (!dateShape) {
    throw new Error("Could not find date shape on slide 1.");
  }

  const { paragraphs } = collectParagraphsFromShape(dateShape.xml);
  const paragraphIndex = paragraphs.findIndex(p => p.xml.includes("Date</a:t>"));
  if (paragraphIndex < 0) {
    throw new Error("Date paragraph not found inside date shape.");
  }

  const paragraphXml = paragraphs[paragraphIndex].xml;
  const runs = [...paragraphXml.matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g)];
  if (runs.length < 8) {
    throw new Error("Strict template mismatch: insufficient runs in date paragraph.");
  }

  const [day, mon, year] = formattedDate.split(" ");
  const y = String(year);
  const yHead = y.slice(0, 3);
  const yTail = y.slice(3);

  // Build replacements for each run, preserving whitespace semantics (matching template)
  const replacements = [
    "Date",
    " ",
    ":",
    "\u00a0 " + day + "\u00a0",
    mon,
    " ",
    yHead,
    yTail,
  ].map(escapeXmlText);

  // Only runs 3-7 should be updated (date field)
  let newParaXml = paragraphXml;
  let count = 0;
  newParaXml = newParaXml.replace(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g, (full, txt) => {
    if (count < replacements.length) {
      const val = replacements[count];
      count++;
      return `<a:t>${val}</a:t>`;
    } else {
      count++;
      return full;
    }
  });

  // Patch new paragraph into the shape XML, and into slide1 XML
  const paraObj = paragraphs[paragraphIndex];
  let scopeXml = collectParagraphsFromShape(dateShape.xml).scopeXml;
  let newScopeXml =
    scopeXml.slice(0, paraObj.startInScope) +
    newParaXml +
    scopeXml.slice(paraObj.endInScope);

  const scopeOffset = collectParagraphsFromShape(dateShape.xml).scopeOffset;
  let newShapeXml =
    dateShape.xml.slice(0, scopeOffset) +
    newScopeXml +
    dateShape.xml.slice(scopeOffset + scopeXml.length);

  // Replace the date shape XML:
  const newSlide1Xml =
    slide1XmlOriginal.slice(0, dateShape.start) +
    newShapeXml +
    slide1XmlOriginal.slice(dateShape.end);

  // Write new slide1 XML
  zip.file(SLIDE1_PATH, newSlide1Xml);

  const out = await zip.generateAsync({ type: "uint8array" });

  return {
    updatedPptxBytes: out,
    detected: {
      mode: "strict-template/date-only",
      slidePath: SLIDE1_PATH,
      shapeName: getShapeName(dateShape.xml) ?? "unknown",
    },
  };
}

/**
 * PUBLIC_INTERFACE
 * Verifies the strict invariant: the last slide XML must remain byte-for-byte identical
 */
export async function assertLastSlideUnchanged(originalPptxArrayBuffer, updatedPptxBytes) {
  /** This is a public function. */
  const originalZip = await JSZip.loadAsync(originalPptxArrayBuffer);
  const updatedZip = await JSZip.loadAsync(updatedPptxBytes);

  const slidePaths = originalZip
    .file(/^ppt\/slides\/slide\d+\.xml$/)
    .map(f => f.name)
    .sort((a, b) => {
      const na = Number(a.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
      const nb = Number(b.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
      return na - nb;
    });

  if (!slidePaths.length) throw new Error("No slides found in template.");

  const LAST_SLIDE_PATH = slidePaths[slidePaths.length - 1];

  const orig = originalZip.file(LAST_SLIDE_PATH);
  const next = updatedZip.file(LAST_SLIDE_PATH);
  if (!orig || !next) {
    throw new Error(
      `Invariant check failed: missing ${LAST_SLIDE_PATH} in ` +
      `${!orig ? "original" : "updated"} PPTX.`
    );
  }

  const [origBytes, nextBytes] = await Promise.all([
    orig.async("uint8array"),
    next.async("uint8array")
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
 * PUBLIC_INTERFACE
 * Download and URL functions (unchanged)
 */
export function downloadPptxBytes(pptxBytes, filename) {
  /** This is a public function. */
  const blob = new Blob([pptxBytes], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
  saveAs(blob, filename);
}

export function createPptxObjectUrl(pptxBytes) {
  /** This is a public function. */
  const blob = new Blob([pptxBytes], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
  return URL.createObjectURL(blob);
}

export function todayIsoDate() {
  /** This is a public function. */
  const d = new Date();
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

/**
 * PUBLIC_INTERFACE
 * Prunes the PPTX to only include Slide 1 and the template's last slide (by deck order),
 * while keeping the *last slide XML bytes* untouched.
 *
 * - Updates ppt/presentation.xml <p:sldIdLst> to only keep the first and last entries.
 * - Updates ppt/_rels/presentation.xml.rels to remove unreferenced slide relationships.
 * - Does NOT delete intermediate slide parts from the zip; they are simply no longer referenced.
 * - Last slide XML part remains byte-for-byte identical.
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
  const { sldIdLstXml, entries } = (function parseSldIdList(presentationXml) {
    const listMatch = presentationXml.match(/<p:sldIdLst\b[\s\S]*?<\/p:sldIdLst>/);
    if (!listMatch) {
      throw new Error("presentation.xml missing <p:sldIdLst>.");
    }
    const sldIdLstXml = listMatch[0];
    const entryRe = /<p:sldId\b[^\/]*\/>/g;
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
  })(presXmlOriginal);

  if (entries.length < 2) {
    // Nothing to prune.
    return {
      updatedPptxBytes: new Uint8Array(pptxBytes),
      kept: { firstSlideNumber: 1, lastSlideNumber: entries.length === 1 ? 1 : 0 },
    };
  }

  const firstEntry = entries[0];
  const lastEntry = entries[entries.length - 1];

  const presRelsXmlOriginal = await presRelsFile.async("string");
  const rels = (function parseRelationships(relsXml) {
    const relRe =
      /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bType="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/>/g;
    const rels = [];
    let m;
    while ((m = relRe.exec(relsXml)) !== null) {
      rels.push({ id: m[1], type: m[2], target: m[3], raw: m[0] });
    }
    return rels;
  })(presRelsXmlOriginal);
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
    const openTag = sldIdLstXml.match(/^<p:sldIdLst\b[\s\S]*?>/)?.[0];
    const closeTag = "</p:sldIdLst>";
    if (!openTag || !sldIdLstXml.endsWith(closeTag)) {
      throw new Error("presentation.xml <p:sldIdLst> parse failed.");
    }
    return `${openTag}${firstEntry.raw}${lastEntry.raw}${closeTag}`;
  })();

  const presXmlUpdated = (() => {
    const idx = presXmlOriginal.indexOf(sldIdLstXml);
    if (idx < 0) throw new Error("Failed to update presentation.xml <p:sldIdLst>.");
    return presXmlOriginal.slice(0, idx) + newSldIdLstXml + presXmlOriginal.slice(idx + sldIdLstXml.length);
  })();

  zip.file(PRESENTATION_XML_PATH, presXmlUpdated);

  // 2) Update presentation.xml.rels: remove slide relationships not referenced by the kept rIds.
  const keptRids = new Set([firstEntry.rId, lastEntry.rId]);
  const filteredRelsXml = (() => {
    return presRelsXmlOriginal.replace(
      /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bType="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships\/slide"[^>]*\/>/g,
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

// PUBLIC_INTERFACE
/**
 * Returns a plain JavaScript object representing the first slide, with date.
 */
export function getFirstSlide(dateStr) {
  return {
    type: "first",
    title: "1st Slide",
    date: dateStr || (new Date()).toISOString().slice(0, 10),
    content: `Weekly statistics report for ${dateStr || "..."}`,
  };
}

// PUBLIC_INTERFACE
/**
 * Returns a plain JavaScript object representing the last slide (fixed "thank you" slide).
 */
export function getLastSlide() {
  return {
    type: "last",
    title: "Last Slide",
    content: "Thank you!",
  };
}
