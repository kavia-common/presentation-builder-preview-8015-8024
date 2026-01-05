import JSZip from "jszip";

/**
 * PPTX slide rendering utilities (read-only).
 *
 * This renderer is intentionally conservative:
 * - It NEVER modifies PPTX bytes.
 * - It renders a single slide to an SVG data URL by placing the slide's picture
 *   fills (images) at their exact coordinates. This is sufficient for the
 *   default template where slide visuals are picture-based.
 *
 * Notes:
 * - Coordinates in PPTX are EMUs. We map them to a fixed pixel viewport while
 *   preserving the slide aspect ratio.
 * - Text rendering is intentionally omitted to avoid mis-positioning and to
 *   ensure slide 1 matches the template screenshot via its embedded pictures.
 */

// PPTX default unit: EMU
const DEFAULT_SLIDE_SIZE_EMU = { cx: 12192000, cy: 6858000 }; // widescreen 16:9

function safeB64FromBytes(bytes) {
  // Convert small/medium binary to base64 without stack overflow.
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

function parseRelTargetsById(relsXml) {
  // Minimal relationship parser (avoid external XML parser dependency).
  // Example: <Relationship Id="rId2" Type=".../image" Target="../media/image5.png"/>
  const relRe = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/>/g;
  const map = new Map();
  let m;
  while ((m = relRe.exec(relsXml)) !== null) {
    map.set(m[1], m[2]);
  }
  return map;
}

function resolveSlideRelTargetToZipPath(target) {
  // slide rels live at ppt/slides/_rels/slideN.xml.rels
  // image targets are commonly "../media/imageX.png"
  if (!target) return null;
  if (target.startsWith("../")) {
    return `ppt/${target.slice(3)}`;
  }
  if (target.startsWith("/")) {
    // Uncommon in PPTX rels; treat as ppt-relative
    return `ppt${target}`;
  }
  // Relative without ../ : treat relative to ppt/slides/
  return `ppt/slides/${target}`;
}

function extractPicsFromSlideXml(slideXml) {
  // Collect <p:pic> blocks, then find:
  // - blip embed rId: <a:blip r:embed="rIdX" .../>
  // - xfrm off/ext (EMU) from <a:off x=".." y=".."/><a:ext cx=".." cy=".."/>
  const pics = [];
  const picRe = /<p:pic\b[\s\S]*?<\/p:pic>/g;
  let m;
  while ((m = picRe.exec(slideXml)) !== null) {
    const picXml = m[0];
    const embed = picXml.match(/<a:blip\b[^>]*\br:embed="([^"]+)"/)?.[1] ?? null;

    const offMatch = picXml.match(/<a:off\b[^>]*\bx="(\d+)"\s+y="(\d+)"/);
    const extMatch = picXml.match(/<a:ext\b[^>]*\bcx="(\d+)"\s+cy="(\d+)"/);

    const x = offMatch ? Number(offMatch[1]) : 0;
    const y = offMatch ? Number(offMatch[2]) : 0;
    const cx = extMatch ? Number(extMatch[1]) : 0;
    const cy = extMatch ? Number(extMatch[2]) : 0;

    // Keep order as in XML (z-order approximated).
    pics.push({ embed, x, y, cx, cy });
  }
  return pics;
}

function getPresentationSlideSizeEmu(presentationXml) {
  const m = presentationXml.match(/<p:sldSz\b[^>]*\bcx="(\d+)"\s+cy="(\d+)"/);
  if (!m) return DEFAULT_SLIDE_SIZE_EMU;
  const cx = Number(m[1]);
  const cy = Number(m[2]);
  if (!Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) {
    return DEFAULT_SLIDE_SIZE_EMU;
  }
  return { cx, cy };
}

/**
 * PUBLIC_INTERFACE
 * Returns slide indices present in the PPTX (sorted).
 *
 * @param {Uint8Array} pptxBytes
 * @returns {Promise<number[]>}
 */
export async function listSlideIndexes(pptxBytes) {
  /** This is a public function. */
  const zip = await JSZip.loadAsync(pptxBytes);
  const slidePaths = zip.file(/^ppt\/slides\/slide\d+\.xml$/).map((f) => f.name);
  const indexes = slidePaths
    .map((p) => Number(p.match(/slide(\d+)\.xml$/)?.[1] ?? 0))
    .filter((n) => Number.isFinite(n) && n > 0)
    .sort((a, b) => a - b);
  return indexes;
}

/**
 * PUBLIC_INTERFACE
 * Reads the slide size (EMU) from ppt/presentation.xml. Falls back to widescreen 16:9.
 *
 * @param {Uint8Array} pptxBytes
 * @returns {Promise<{cx:number, cy:number}>}
 */
export async function getSlideSizeEmu(pptxBytes) {
  /** This is a public function. */
  const zip = await JSZip.loadAsync(pptxBytes);
  const pres = zip.file("ppt/presentation.xml");
  if (!pres) return DEFAULT_SLIDE_SIZE_EMU;
  const xml = await pres.async("string");
  return getPresentationSlideSizeEmu(xml);
}

/**
 * PUBLIC_INTERFACE
 * Renders a slide to an SVG data URL with embedded images placed at their PPTX coordinates.
 *
 * @param {Uint8Array} pptxBytes
 * @param {number} slideIndex 1-based slide number
 * @param {{widthPx?: number}} options
 * @returns {Promise<{ dataUrl: string, widthPx: number, heightPx: number }>}
 */
export async function renderSlideToSvgDataUrl(pptxBytes, slideIndex, options = {}) {
  /** This is a public function. */
  if (!pptxBytes || !pptxBytes.length) {
    throw new Error("No PPTX bytes provided.");
  }
  if (!Number.isFinite(slideIndex) || slideIndex < 1) {
    throw new Error("Invalid slide index.");
  }

  const zip = await JSZip.loadAsync(pptxBytes);

  const slidePath = `ppt/slides/slide${slideIndex}.xml`;
  const relsPath = `ppt/slides/_rels/slide${slideIndex}.xml.rels`;

  const slideFile = zip.file(slidePath);
  if (!slideFile) throw new Error(`Missing ${slidePath} in PPTX.`);
  const slideXml = await slideFile.async("string");

  const relsFile = zip.file(relsPath);
  const relsXml = relsFile ? await relsFile.async("string") : "";
  const relsById = parseRelTargetsById(relsXml);

  const slideSize = await (async () => {
    const pres = zip.file("ppt/presentation.xml");
    if (!pres) return DEFAULT_SLIDE_SIZE_EMU;
    return getPresentationSlideSizeEmu(await pres.async("string"));
  })();

  const widthPx = Math.max(320, Math.floor(options.widthPx ?? 960));
  const heightPx = Math.floor((widthPx * slideSize.cy) / slideSize.cx);

  const pics = extractPicsFromSlideXml(slideXml);

  // Build SVG elements for images.
  const imageEls = [];
  for (const pic of pics) {
    if (!pic.embed) continue;
    const target = relsById.get(pic.embed);
    const zipPath = resolveSlideRelTargetToZipPath(target);
    if (!zipPath) continue;

    const imgFile = zip.file(zipPath);
    if (!imgFile) continue;

    const bytes = await imgFile.async("uint8array");
    const ext = zipPath.split(".").pop()?.toLowerCase() ?? "";
    const mime =
      ext === "png"
        ? "image/png"
        : ext === "jpg" || ext === "jpeg"
          ? "image/jpeg"
          : ext === "gif"
            ? "image/gif"
            : ext === "svg"
              ? "image/svg+xml"
              : "application/octet-stream";

    const b64 = safeB64FromBytes(bytes);

    // Convert EMU to pixel in the chosen viewport.
    const x = (pic.x / slideSize.cx) * widthPx;
    const y = (pic.y / slideSize.cy) * heightPx;
    const w = (pic.cx / slideSize.cx) * widthPx;
    const h = (pic.cy / slideSize.cy) * heightPx;

    // Preserve order: earlier elements render behind later ones.
    imageEls.push(
      `<image x="${x}" y="${y}" width="${w}" height="${h}" href="data:${mime};base64,${b64}" preserveAspectRatio="none" />`
    );
  }

  // If we couldn't find any pictures, still return a valid SVG.
  const bg = imageEls.length
    ? imageEls.join("\n")
    : `<rect x="0" y="0" width="100%" height="100%" fill="#ffffff" />`;

  const svg = [
    `<?xml version="1.0" encoding="UTF-8"?>`,
    `<svg xmlns="http://www.w3.org/2000/svg" width="${widthPx}" height="${heightPx}" viewBox="0 0 ${widthPx} ${heightPx}">`,
    bg,
    `</svg>`,
  ].join("\n");

  const svgB64 = btoa(unescape(encodeURIComponent(svg)));
  return {
    dataUrl: `data:image/svg+xml;base64,${svgB64}`,
    widthPx,
    heightPx,
  };
}
