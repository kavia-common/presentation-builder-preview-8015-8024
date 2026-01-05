import JSZip from "jszip";

/**
 * Utilities for creating a reliable "thumbnail-like" preview for PPTX slides
 * without depending on browser PPTX rendering.
 *
 * IMPORTANT:
 * - This is a read-only operation over the generated PPTX bytes.
 * - We do NOT modify the PPTX in any way; therefore it cannot violate the
 *   "last slide byte-identical" requirement.
 */

/**
 * PUBLIC_INTERFACE
 * Extracts a text-only representation of a slide by reading ppt/slides/slideN.xml
 * and collecting <a:t> text nodes in order.
 *
 * This is intentionally lightweight and avoids full XML parsing to reduce the
 * chance of inconsistencies. It is "good enough" for a visible preview panel.
 *
 * @param {string} slideXml Slide XML content as string.
 * @returns {{ lines: string[], rawText: string }}
 */
export function extractSlideText(slideXml) {
  /** This is a public function. */
  const aTextRe = /<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g;
  const texts = [];
  let m;
  while ((m = aTextRe.exec(slideXml)) !== null) {
    // Basic XML entity unescape for readability in preview.
    const t = String(m[1] ?? "")
      .replaceAll("&lt;", "<")
      .replaceAll("&gt;", ">")
      .replaceAll("&quot;", '"')
      .replaceAll("&apos;", "'")
      .replaceAll("&amp;", "&");
    texts.push(t);
  }

  // Normalize whitespace and NBSP variants for display (visual only).
  const normalized = texts
    .join(" ")
    .replaceAll("\u00a0", " ")
    .replace(/\s+/g, " ")
    .trim();

  // Split into compact lines for rendering.
  const lines = [];
  if (normalized) {
    const maxLen = 44;
    let i = 0;
    while (i < normalized.length && lines.length < 10) {
      lines.push(normalized.slice(i, i + maxLen));
      i += maxLen;
    }
  }

  return { lines, rawText: normalized };
}

/**
 * Picks the last slide path by highest slide number present.
 */
function findLastSlidePathFromZip(zip) {
  const slidePaths = zip
    .file(/^ppt\/slides\/slide\d+\.xml$/)
    .map((f) => f.name)
    .sort((a, b) => {
      const na = Number(a.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
      const nb = Number(b.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
      return na - nb;
    });

  return slidePaths.length ? slidePaths[slidePaths.length - 1] : null;
}

/**
 * PUBLIC_INTERFACE
 * Produces "thumbnail cards" data for Slide 1 and the last slide.
 *
 * @param {Uint8Array} pptxBytes Generated PPTX bytes.
 * @returns {Promise<{ slide1: { index:number, path:string, lines:string[] }, last: { index:number, path:string, lines:string[] } }>}
 */
export async function getSlide1AndLastTextThumbs(pptxBytes) {
  /** This is a public function. */
  if (!pptxBytes || !pptxBytes.length) {
    throw new Error("No PPTX bytes provided.");
  }

  const zip = await JSZip.loadAsync(pptxBytes);

  const slide1Path = "ppt/slides/slide1.xml";
  const lastPath = findLastSlidePathFromZip(zip);

  const slide1File = zip.file(slide1Path);
  if (!slide1File) {
    throw new Error(`Missing ${slide1Path} in PPTX.`);
  }
  if (!lastPath) {
    throw new Error("Could not determine last slide path from PPTX.");
  }
  const lastFile = zip.file(lastPath);
  if (!lastFile) {
    throw new Error(`Missing ${lastPath} in PPTX.`);
  }

  const [slide1Xml, lastXml] = await Promise.all([
    slide1File.async("string"),
    lastFile.async("string"),
  ]);

  const s1 = extractSlideText(slide1Xml);
  const sl = extractSlideText(lastXml);

  const lastIndex = Number(lastPath.match(/slide(\d+)\.xml$/)?.[1] ?? 0);

  return {
    slide1: { index: 1, path: slide1Path, lines: s1.lines },
    last: { index: lastIndex || 0, path: lastPath, lines: sl.lines },
  };
}
