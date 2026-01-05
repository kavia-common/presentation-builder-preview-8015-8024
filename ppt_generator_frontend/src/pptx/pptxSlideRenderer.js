import JSZip from "jszip";

/**
 * PPTX slide rendering utilities (read-only).
 *
 * This renderer is intentionally conservative:
 * - It NEVER modifies PPTX bytes.
 * - It renders a single slide to an SVG data URL.
 *
 * Rendering strategy:
 * 1) Render solid background fills for shapes (when present).
 * 2) Render picture shapes using slide relationships (<p:pic> -> r:embed target).
 * 3) Render text from shapes (<p:sp>) with a best-effort mapping of:
 *    - position and size (EMU -> px)
 *    - font size, family, weight, italic, underline, color
 *    - alignment (left/center/right)
 *
 * Notes:
 * - This is not a full PowerPoint renderer; it is a deterministic, read-only
 *   preview aimed at matching the bundled default template, including the last
 *   slide shown in the user screenshot.
 * - We keep Slide 1 date-only editing logic elsewhere; this module does not
 *   inspect/alter slide 1 beyond read-only rendering.
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
  const relRe =
    /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/>/g;
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

function getPresentationSlideSizeEmu(presentationXml) {
  const m = presentationXml.match(
    /<p:sldSz\b[^>]*\bcx="(\d+)"\s+cy="(\d+)"/
  );
  if (!m) return DEFAULT_SLIDE_SIZE_EMU;
  const cx = Number(m[1]);
  const cy = Number(m[2]);
  if (!Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) {
    return DEFAULT_SLIDE_SIZE_EMU;
  }
  return { cx, cy };
}

function emuToPxX(xEmu, slideSize, widthPx) {
  return (Number(xEmu || 0) / slideSize.cx) * widthPx;
}
function emuToPxY(yEmu, slideSize, heightPx) {
  return (Number(yEmu || 0) / slideSize.cy) * heightPx;
}

function decodeXmlText(text) {
  return String(text ?? "")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function escapeXmlAttr(text) {
  return String(text ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function parseColorFromSrbgClr(spPrXml) {
  // Most common: <a:srgbClr val="RRGGBB"/>
  const m = spPrXml.match(/<a:srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/);
  return m ? `#${m[1]}` : null;
}

function extractShapeTransformEmu(shapeXml) {
  // <a:xfrm><a:off x=".." y=".."/><a:ext cx=".." cy=".."/></a:xfrm>
  const xfrmMatch = shapeXml.match(/<a:xfrm\b[\s\S]*?<\/a:xfrm>/);
  const xfrm = xfrmMatch ? xfrmMatch[0] : "";
  const offMatch = xfrm.match(/<a:off\b[^>]*\bx="(\d+)"\s+y="(\d+)"/);
  const extMatch = xfrm.match(/<a:ext\b[^>]*\bcx="(\d+)"\s+cy="(\d+)"/);
  return {
    x: offMatch ? Number(offMatch[1]) : 0,
    y: offMatch ? Number(offMatch[2]) : 0,
    cx: extMatch ? Number(extMatch[1]) : 0,
    cy: extMatch ? Number(extMatch[2]) : 0,
  };
}

function extractPicsFromSlideXml(slideXml) {
  // Collect <p:pic> blocks, then find:
  // - blip embed rId: <a:blip r:embed="rIdX" .../>
  // - xfrm off/ext (EMU)
  const pics = [];
  const picRe = /<p:pic\b[\s\S]*?<\/p:pic>/g;
  let m;
  while ((m = picRe.exec(slideXml)) !== null) {
    const picXml = m[0];
    const embed =
      picXml.match(/<a:blip\b[^>]*\br:embed="([^"]+)"/)?.[1] ?? null;

    const { x, y, cx, cy } = extractShapeTransformEmu(picXml);

    // Keep order as in XML (z-order approximated).
    pics.push({ embed, x, y, cx, cy });
  }
  return pics;
}

function collectShapeBlocks(slideXml) {
  const shapes = [];
  const re = /<p:sp\b[\s\S]*?<\/p:sp>/g;
  let m;
  while ((m = re.exec(slideXml)) !== null) {
    shapes.push(m[0]);
  }
  return shapes;
}

function extractShapeFillColor(shapeXml) {
  // Best-effort solid fill:
  // <p:spPr>...<a:solidFill><a:srgbClr val="..."/></a:solidFill>...</p:spPr>
  const spPr = shapeXml.match(/<p:spPr\b[\s\S]*?<\/p:spPr>/)?.[0] ?? "";
  if (!spPr) return null;
  if (!spPr.includes("<a:solidFill")) return null;
  const c = parseColorFromSrbgClr(spPr);
  return c;
}

function extractTextRunsFromShape(shapeXml) {
  // Extract paragraph(s) from <p:txBody>. Preserve order.
  // Return [{ text, rPrXml, pPrXml, idxInParagraph, paragraphIndex }]
  const txBody = shapeXml.match(/<p:txBody\b[\s\S]*?<\/p:txBody>/)?.[0] ?? "";
  if (!txBody) return [];

  const paragraphs = txBody.match(/<a:p\b[\s\S]*?<\/a:p>/g) ?? [];
  const out = [];

  for (let pIndex = 0; pIndex < paragraphs.length; pIndex += 1) {
    const pXml = paragraphs[pIndex];
    const pPrXml = pXml.match(/<a:pPr\b[\s\S]*?<\/a:pPr>/)?.[0] ?? "";

    // Include <a:r> runs and also handle <a:fld> (fields) similarly if present.
    const runMatches = pXml.match(/<(a:r|a:fld)\b[\s\S]*?<\/(a:r|a:fld)>/g) ?? [];
    let runIdx = 0;
    for (const runXml of runMatches) {
      const tNodes = runXml.match(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g) ?? [];
      if (!tNodes.length) {
        runIdx += 1;
        continue;
      }

      // Concatenate all <a:t> nodes within the run.
      const text = tNodes
        .map((n) => n.match(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/)?.[1] ?? "")
        .map(decodeXmlText)
        .join("");

      const rPrXml = runXml.match(/<a:rPr\b[\s\S]*?<\/a:rPr>/)?.[0] ?? "";
      out.push({
        text,
        rPrXml,
        pPrXml,
        idxInParagraph: runIdx,
        paragraphIndex: pIndex,
      });

      runIdx += 1;
    }
  }

  return out;
}

function parseTextStyle({ rPrXml, pPrXml }) {
  // Defaults to something reasonable if style is missing.
  const fontSizeHalfPoints = Number(rPrXml.match(/\bsz="(\d+)"/)?.[1] ?? 0);
  const fontSizePx = fontSizeHalfPoints ? (fontSizeHalfPoints / 100) * 1.333 : 18; // rough: pt -> px

  const isBold = /\bb="1"/.test(rPrXml);
  const isItalic = /\bi="1"/.test(rPrXml);
  const isUnderline = /\bu="(sng|dbl)"/.test(rPrXml);

  const latin = rPrXml.match(/<a:latin\b[^>]*\btypeface="([^"]+)"/)?.[1] ?? "";
  const fontFamily = latin || "Arial, sans-serif";

  const colorHex = (() => {
    // Prefer run color
    const runClr = rPrXml.match(/<a:srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/)?.[1];
    if (runClr) return `#${runClr}`;
    // Could also be theme-based; omit theme parsing for now.
    return "#111827";
  })();

  const align = (() => {
    // Paragraph alignment: <a:pPr algn="ctr|l|r|just">
    const a = pPrXml.match(/\balgn="([^"]+)"/)?.[1] ?? "";
    if (a === "ctr") return "middle";
    if (a === "r") return "end";
    return "start";
  })();

  return {
    fontSizePx,
    fontFamily,
    fontWeight: isBold ? 700 : 400,
    fontStyle: isItalic ? "italic" : "normal",
    textDecoration: isUnderline ? "underline" : "none",
    fill: colorHex,
    textAnchor: align,
  };
}

function parseLeftInsetEmu(pPrXml) {
  // <a:pPr marL="..."> in EMU for margin-left. Optional.
  const m = pPrXml.match(/\bmarL="(\d+)"/);
  return m ? Number(m[1]) : 0;
}

function extractShapeTextBodyInsetsEmu(shapeXml) {
  // <a:bodyPr lIns=".." tIns=".." rIns=".." bIns=".."/>
  const bodyPr = shapeXml.match(/<a:bodyPr\b[^>]*\/>/)?.[0] ?? "";
  const lIns = Number(bodyPr.match(/\blIns="(\d+)"/)?.[1] ?? 0);
  const tIns = Number(bodyPr.match(/\btIns="(\d+)"/)?.[1] ?? 0);
  const rIns = Number(bodyPr.match(/\brIns="(\d+)"/)?.[1] ?? 0);
  const bIns = Number(bodyPr.match(/\bbIns="(\d+)"/)?.[1] ?? 0);
  return { lIns, tIns, rIns, bIns };
}

function normalizeWhitespaceForSvg(text) {
  // Keep intentional spaces while avoiding SVG collapsing runs too aggressively.
  // PowerPoint often uses NBSP; map to space and preserve sequences via CSS.
  return String(text ?? "").replaceAll("\u00a0", " ");
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
 * Renders a slide to an SVG data URL with shapes/images/text placed at PPTX coordinates.
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

  const widthPx = Math.max(320, Math.floor(options.widthPx ?? 1040));
  const heightPx = Math.floor((widthPx * slideSize.cy) / slideSize.cx);

  // 1) Shapes (background fills behind content)
  const shapeEls = [];
  const shapeBlocks = collectShapeBlocks(slideXml);
  for (const shapeXml of shapeBlocks) {
    const fill = extractShapeFillColor(shapeXml);
    if (!fill) continue;

    const { x, y, cx, cy } = extractShapeTransformEmu(shapeXml);
    // Skip zero-sized shapes.
    if (!cx || !cy) continue;

    const rx = emuToPxX(x, slideSize, widthPx);
    const ry = emuToPxY(y, slideSize, heightPx);
    const rw = emuToPxX(cx, slideSize, widthPx);
    const rh = emuToPxY(cy, slideSize, heightPx);

    shapeEls.push(
      `<rect x="${rx}" y="${ry}" width="${rw}" height="${rh}" fill="${escapeXmlAttr(fill)}" />`
    );
  }

  // 2) Pictures
  const pics = extractPicsFromSlideXml(slideXml);
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

    const x = emuToPxX(pic.x, slideSize, widthPx);
    const y = emuToPxY(pic.y, slideSize, heightPx);
    const w = emuToPxX(pic.cx, slideSize, widthPx);
    const h = emuToPxY(pic.cy, slideSize, heightPx);

    imageEls.push(
      `<image x="${x}" y="${y}" width="${w}" height="${h}" href="data:${mime};base64,${b64}" preserveAspectRatio="none" />`
    );
  }

  // 3) Text (best effort; needed for the THANK YOU last slide screenshot)
  const textEls = [];
  for (const shapeXml of shapeBlocks) {
    // Only render if it has a txBody.
    if (!shapeXml.includes("<p:txBody")) continue;

    const { x, y, cx, cy } = extractShapeTransformEmu(shapeXml);
    if (!cx || !cy) continue;

    const { lIns, tIns } = extractShapeTextBodyInsetsEmu(shapeXml);

    const boxX = emuToPxX(x + lIns, slideSize, widthPx);
    const boxY = emuToPxY(y + tIns, slideSize, heightPx);
    const boxW = emuToPxX(Math.max(0, cx - lIns), slideSize, widthPx);
    const boxH = emuToPxY(Math.max(0, cy - tIns), slideSize, heightPx);

    const runs = extractTextRunsFromShape(shapeXml);
    if (!runs.length) continue;

    // PowerPoint text layout is complex; we approximate line breaks by paragraph index.
    // For each paragraph, join runs and render as a single <text> element.
    const paragraphs = new Map();
    for (const r of runs) {
      const key = r.paragraphIndex;
      const current = paragraphs.get(key) ?? { pieces: [], styleSeed: r, pPrXml: r.pPrXml };
      current.pieces.push(r.text);
      // Keep first run as style seed.
      paragraphs.set(key, current);
    }

    const sortedKeys = [...paragraphs.keys()].sort((a, b) => a - b);

    // Line height heuristic: 1.2 * fontSize
    let cursorY = boxY;
    for (const pIdx of sortedKeys) {
      const p = paragraphs.get(pIdx);
      if (!p) continue;

      const seed = p.styleSeed;
      const style = parseTextStyle({ rPrXml: seed.rPrXml, pPrXml: seed.pPrXml });
      const marL = parseLeftInsetEmu(seed.pPrXml);
      const insetX = emuToPxX(marL, slideSize, widthPx);

      const lineHeight = Math.max(1, style.fontSizePx * 1.2);
      cursorY += lineHeight;

      // anchor in X depends on alignment; use boxW to compute.
      let xPos = boxX + insetX;
      if (style.textAnchor === "middle") xPos = boxX + boxW / 2;
      if (style.textAnchor === "end") xPos = boxX + boxW;

      // SVG text y is baseline; this is a heuristic.
      const yPos = Math.min(boxY + boxH, cursorY);

      const text = normalizeWhitespaceForSvg(p.pieces.join(""));
      if (!text.trim()) continue;

      textEls.push(
        [
          `<text x="${xPos}" y="${yPos}"`,
          ` font-family="${escapeXmlAttr(style.fontFamily)}"`,
          ` font-size="${style.fontSizePx}"`,
          ` font-weight="${style.fontWeight}"`,
          ` font-style="${escapeXmlAttr(style.fontStyle)}"`,
          ` text-decoration="${escapeXmlAttr(style.textDecoration)}"`,
          ` fill="${escapeXmlAttr(style.fill)}"`,
          ` text-anchor="${escapeXmlAttr(style.textAnchor)}"`,
          ` style="white-space: pre;"`,
          `>`,
          `${escapeXmlAttr(text)}`,
          `</text>`,
        ].join("")
      );
    }
  }

  // Base slide background: white (template uses white background in screenshot).
  // Then shape fills, then images, then text.
  const svgLayers = [
    `<rect x="0" y="0" width="100%" height="100%" fill="#ffffff" />`,
    ...shapeEls,
    ...imageEls,
    ...textEls,
  ];

  const svg = [
    `<?xml version="1.0" encoding="UTF-8"?>`,
    `<svg xmlns="http://www.w3.org/2000/svg" width="${widthPx}" height="${heightPx}" viewBox="0 0 ${widthPx} ${heightPx}">`,
    svgLayers.join("\n"),
    `</svg>`,
  ].join("\n");

  const svgB64 = btoa(unescape(encodeURIComponent(svg)));
  return {
    dataUrl: `data:image/svg+xml;base64,${svgB64}`,
    widthPx,
    heightPx,
  };
}
