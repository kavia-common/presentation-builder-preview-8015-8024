import JSZip from "jszip";
import {
  applySkillFactoryPlaceholderSlides,
  applySkillFactorySlide1Content,
} from "./skillFactorySlide1";

/**
 * Skill Factory slide scaffolding utilities.
 *
 * IMPORTANT INVARIANTS (project requirements):
 * - The original deck's Slide 1 date editing remains the only editable content.
 * - The original last slide must remain byte-identical to the template.
 * - Skill Factory slides are inserted AFTER existing slides but BEFORE the last slide.
 *
 * This module is intentionally conservative:
 * - It creates new slide parts by COPYING existing slide parts (byte-preserving approach).
 * - It does not modify the template's last slide part.
 * - It only modifies the newly inserted (copied) Skill Factory slides.
 *
 * Future iterations can introduce richer slide XML generation, but must keep the invariant
 * that the template's last slide bytes remain unchanged.
 */

const SLIDES_DIR = "ppt/slides";
const SLIDES_RELS_DIR = "ppt/slides/_rels";
const PRESENTATION_XML = "ppt/presentation.xml";
const PRESENTATION_RELS_XML = "ppt/_rels/presentation.xml.rels";
const CONTENT_TYPES_XML = "[Content_Types].xml";

/**
 * PUBLIC_INTERFACE
 * Returns the list of slide indices present in the zip (sorted ascending).
 *
 * @param {JSZip} zip
 * @returns {number[]}
 */
export function listSlideIndexesFromZip(zip) {
  /** This is a public function. */
  const slidePaths = zip
    .file(new RegExp(`^${SLIDES_DIR}/slide\\d+\\.xml$`))
    .map((f) => f.name);

  return slidePaths
    .map((p) => Number(p.match(/slide(\d+)\.xml$/)?.[1] ?? 0))
    .filter((n) => Number.isFinite(n) && n > 0)
    .sort((a, b) => a - b);
}

function mustGetFileString(zip, path) {
  const f = zip.file(path);
  if (!f) throw new Error(`Missing required PPTX part: ${path}`);
  return f.async("string");
}

function mustGetFileBytes(zip, path) {
  const f = zip.file(path);
  if (!f) throw new Error(`Missing required PPTX part: ${path}`);
  return f.async("uint8array");
}

function getLastSlideIndex(indexes) {
  if (!indexes.length) throw new Error("No slides found in PPTX.");
  return indexes[indexes.length - 1];
}

function getSlidePath(n) {
  return `${SLIDES_DIR}/slide${n}.xml`;
}

function getSlideRelsPath(n) {
  return `${SLIDES_RELS_DIR}/slide${n}.xml.rels`;
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

function replaceOnce(haystack, needle, replacement) {
  const idx = haystack.indexOf(needle);
  if (idx < 0) return null;
  return haystack.slice(0, idx) + replacement + haystack.slice(idx + needle.length);
}

function parseRelationships(relsXml) {
  // Minimal parser: keep original formatting as much as possible by only appending new Relationship nodes.
  const relRe =
    /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bType="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/>/g;
  const rels = [];
  let m;
  while ((m = relRe.exec(relsXml)) !== null) {
    rels.push({ id: m[1], type: m[2], target: m[3], raw: m[0] });
  }
  return rels;
}

function nextRelationshipId(relsXml) {
  const rels = parseRelationships(relsXml);
  const max = rels
    .map((r) => Number(r.id.replace(/^rId/, "")))
    .filter((n) => Number.isFinite(n))
    .reduce((a, b) => Math.max(a, b), 0);
  return `rId${max + 1}`;
}

function ensureContentTypeForSlide(contentTypesXml) {
  // Slides are usually already covered by:
  // <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  // We must add overrides for any newly added slideN.xml if not present.
  return contentTypesXml;
}

function hasSlideOverride(contentTypesXml, slideIndex) {
  return contentTypesXml.includes(`PartName="/ppt/slides/slide${slideIndex}.xml"`);
}

function addSlideOverride(contentTypesXml, slideIndex) {
  if (hasSlideOverride(contentTypesXml, slideIndex)) return contentTypesXml;

  const override =
    `<Override PartName="/ppt/slides/slide${slideIndex}.xml" ` +
    `ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;

  // Insert before closing </Types>
  const close = "</Types>";
  if (!contentTypesXml.includes(close)) {
    throw new Error("[Content_Types].xml missing </Types>.");
  }
  return contentTypesXml.replace(close, `${override}${close}`);
}

function maxNumericIdInSldIdLst(entries) {
  return entries.reduce((m, e) => Math.max(m, Number(e.id) || 0), 0);
}

/**
 * Creates 4 new slide parts by copying an existing "template slide".
 *
 * For this iteration:
 * - Slide 1 of factory is implemented by copying a richer slide (defaults to slide1 as a visual stand-in).
 * - Slides 2-4 are placeholders (copied from the same template slide).
 *
 * NOTE: Copying keeps everything byte/relationship consistent for the duplicated slide itself.
 * We will generate new slideN.xml / slideN.xml.rels parts as byte-identical copies of the source.
 *
 * @param {JSZip} zip
 * @param {number} sourceSlideIndex
 * @param {number[]} newSlideIndexes
 */
async function copySlideParts(zip, sourceSlideIndex, newSlideIndexes) {
  const srcSlidePath = getSlidePath(sourceSlideIndex);
  const srcRelsPath = getSlideRelsPath(sourceSlideIndex);

  const [srcSlideBytes, srcRelsBytes] = await Promise.all([
    mustGetFileBytes(zip, srcSlidePath),
    // Some slides might not have rels; handle gracefully by creating an empty rels part if absent.
    zip.file(srcRelsPath) ? mustGetFileBytes(zip, srcRelsPath) : Promise.resolve(null),
  ]);

  for (const idx of newSlideIndexes) {
    zip.file(getSlidePath(idx), srcSlideBytes);
    if (srcRelsBytes) {
      zip.file(getSlideRelsPath(idx), srcRelsBytes);
    }
  }
}

/**
 * PUBLIC_INTERFACE
 * Adds a Skill Factory (4 slides) before the last slide in the PPTX.
 *
 * @param {Uint8Array} pptxBytes - existing PPTX bytes (already date-only edited)
 * @param {{ kind: "java" | "dataEngineering", label: string, slide1?: any }} factory
 * @returns {Promise<{ updatedPptxBytes: Uint8Array, insertedSlideIndexes: number[] }>}
 */
export async function addSkillFactoryToPptx(pptxBytes, factory) {
  /** This is a public function. */
  if (!pptxBytes || !pptxBytes.length) throw new Error("No PPTX bytes provided.");
  if (!factory || !factory.kind) throw new Error("Factory is required.");

  const zip = await JSZip.loadAsync(pptxBytes);

  const slideIndexes = listSlideIndexesFromZip(zip);
  const lastSlideIndex = getLastSlideIndex(slideIndexes);

  // We will append 4 slides after the highest existing slide index, and then insert them before the last slide in ordering.
  const maxExisting = lastSlideIndex;
  const newSlideIndexes = [
    maxExisting + 1,
    maxExisting + 2,
    maxExisting + 3,
    maxExisting + 4,
  ];

  // Choose a source slide to copy from.
  // For now we copy slide 1 as the base, then mutate the copied slide (factory slide 1) to match the screenshot.
  const sourceSlideIndex = 1;

  // 1) Copy slide parts (xml and rels) to new slide indices.
  await copySlideParts(zip, sourceSlideIndex, newSlideIndexes);

  // 1b) Apply Skill Factory content on the newly inserted slides only.
  // Slide 1 is "fully implemented"; slides 2-4 are placeholders for now.
  await applySkillFactorySlide1Content(zip, newSlideIndexes[0], factory.slide1 || {});
  await applySkillFactoryPlaceholderSlides(zip, newSlideIndexes.slice(1), factory.label);

  // 2) Update [Content_Types].xml to include Override for each new slide part.
  const ctXmlOriginal = await mustGetFileString(zip, CONTENT_TYPES_XML);
  let ctXml = ensureContentTypeForSlide(ctXmlOriginal);
  for (const idx of newSlideIndexes) {
    ctXml = addSlideOverride(ctXml, idx);
  }
  zip.file(CONTENT_TYPES_XML, ctXml);

  // 3) Update ppt/_rels/presentation.xml.rels with new slide relationships.
  const presRelsXmlOriginal = await mustGetFileString(zip, PRESENTATION_RELS_XML);
  let presRelsXml = presRelsXmlOriginal;

  // Find insertion point before </Relationships>
  if (!presRelsXml.includes("</Relationships>")) {
    throw new Error("presentation.xml.rels missing </Relationships>.");
  }

  const slideRelType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

  const newRIds = [];
  for (const idx of newSlideIndexes) {
    const newRid = nextRelationshipId(presRelsXml);
    newRIds.push(newRid);

    const relNode = `<Relationship Id="${newRid}" Type="${slideRelType}" Target="slides/slide${idx}.xml"/>`;
    presRelsXml = presRelsXml.replace("</Relationships>", `${relNode}</Relationships>`);
  }

  zip.file(PRESENTATION_RELS_XML, presRelsXml);

  // 4) Update ppt/presentation.xml slide order: insert new <p:sldId ... r:id="..."/> before the last slide entry.
  const presXmlOriginal = await mustGetFileString(zip, PRESENTATION_XML);
  const { sldIdLstXml, entries } = parseSldIdList(presXmlOriginal);

  if (!entries.length) throw new Error("presentation.xml has empty <p:sldIdLst>.");

  const lastEntry = entries[entries.length - 1];
  const lastEntryRaw = lastEntry.raw;

  // Compute new unique ids for p:sldId (these are not slide numbers; they are internal ids).
  let nextSldId = maxNumericIdInSldIdLst(entries) + 1;

  const insertedNodes = newRIds.map((rid) => {
    const node = `<p:sldId id="${nextSldId}" r:id="${rid}"/>`;
    nextSldId += 1;
    return node;
  });

  const updatedSldIdLstXml = (() => {
    const replacement = `${insertedNodes.join("")}${lastEntryRaw}`;
    const replaced = replaceOnce(sldIdLstXml, lastEntryRaw, replacement);
    if (!replaced) {
      throw new Error("Failed to insert Skill Factory slides into <p:sldIdLst>.");
    }
    return replaced;
  })();

  const presXmlUpdated = (() => {
    const replaced = replaceOnce(presXmlOriginal, sldIdLstXml, updatedSldIdLstXml);
    if (!replaced) throw new Error("Failed to update presentation.xml <p:sldIdLst>.");
    return replaced;
  })();

  zip.file(PRESENTATION_XML, presXmlUpdated);

  // NOTE: We intentionally do NOT modify the original last slide's XML part.
  // It remains the same path (slide{lastSlideIndex}.xml) and bytes, satisfying the invariant.

  const out = await zip.generateAsync({ type: "uint8array" });

  return {
    updatedPptxBytes: out,
    insertedSlideIndexes: newSlideIndexes,
  };
}

/**
 * PUBLIC_INTERFACE
 * Provides a standard 4-slide scaffold metadata for a given Skill Factory kind.
 *
 * @param {"java"|"dataEngineering"} kind
 * @returns {{ kind:string, label:string, slides: { title:string, kind:string, indexInFactory:number }[] }}
 */
export function getSkillFactoryScaffold(kind) {
  /** This is a public function. */
  const label =
    kind === "java" ? "Java" : kind === "dataEngineering" ? "Data Engineering" : String(kind);

  // Titles are placeholders for now; slide 1 is considered "implemented" layout-wise by copy.
  return {
    kind,
    label,
    slides: [
      { kind, indexInFactory: 1, title: `${label} Skill Factory — Overview` },
      { kind, indexInFactory: 2, title: `${label} Skill Factory — Placeholder 2` },
      { kind, indexInFactory: 3, title: `${label} Skill Factory — Placeholder 3` },
      { kind, indexInFactory: 4, title: `${label} Skill Factory — Placeholder 4` },
    ],
  };
}
