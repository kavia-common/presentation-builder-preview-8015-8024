import JSZip from "jszip";

/**
 * PPTX slide rendering utilities (read-only).
 *
 * This renderer is intentionally conservative:
 * - It NEVER modifies PPTX bytes.
 * - It renders a single slide to an SVG data URL.
 *
 * Rendering strategy (best-effort, deterministic):
 * 1) Render a base slide background fill.
 * 2) Render shape fills (solid) behind content (when present).
 * 3) Render picture shapes using slide relationships (<p:pic> -> r:embed target),
 *    honoring common crop rectangles to better match PowerPoint's view.
 * 4) Render text from shapes (<p:sp>) with a best-effort mapping of:
 *    - position and size (EMU -> px)
 *    - font size, family, weight, italic, underline, color
 *    - letter spacing (rPr spc) to match template fidelity (THANK YOU slide)
 *    - alignment (left/center/right)
 *
 * Notes:
 * - This is not a full PowerPoint renderer. It is tuned for the bundled template
 *   and especially the final "THANK YOU" slide fidelity.
 * - Slide 1 date-only editing logic lives elsewhere; this module is purely read-only.
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

function parseColorFromSrbgClr(scopeXml) {
  // Most common: <a:srgbClr val="RRGGBB"/>
  const m = scopeXml.match(/<a:srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/);
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

function collectSpTreeChildren(slideXml) {
  /**
   * Extracts the immediate children of <p:spTree> in order as raw xml fragments.
   * This preserves z-order significantly better than rendering by type groups.
   */
  const tree =
    slideXml.match(/<p:spTree\b[\s\S]*?<\/p:spTree>/)?.[0] ?? "";
  if (!tree) return [];
  // Remove the wrapper tags and keep inner xml.
  const inner = tree
    .replace(/^<p:spTree\b[\s\S]*?>/, "")
    .replace(/<\/p:spTree>$/, "");

  const children = [];
  const childRe = /<(p:sp|p:pic|p:grpSp)\b[\s\S]*?<\/\1>/g;
  let m;
  while ((m = childRe.exec(inner)) !== null) {
    children.push(m[0]);
  }
  return children;
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
  return parseColorFromSrbgClr(spPr);
}

function extractTextRunsFromShape(shapeXml) {
  // Extract paragraph(s) from <p:txBody>. Preserve order.
  // Return [{ text, rPrXml, pPrXml, idxInParagraph, paragraphIndex }]
  const txBody =
    shapeXml.match(/<p:txBody\b[\s\S]*?<\/p:txBody>/)?.[0] ?? "";
  if (!txBody) return [];

  const paragraphs = txBody.match(/<a:p\b[\s\S]*?<\/a:p>/g) ?? [];
  const out = [];

  for (let pIndex = 0; pIndex < paragraphs.length; pIndex += 1) {
    const pXml = paragraphs[pIndex];
    const pPrXml = pXml.match(/<a:pPr\b[\s\S]*?<\/a:pPr>/)?.[0] ?? "";

    // Include <a:r> runs and also handle <a:fld> similarly.
    const runMatches =
      pXml.match(/<(a:r|a:fld)\b[\s\S]*?<\/(a:r|a:fld)>/g) ?? [];
    let runIdx = 0;
    for (const runXml of runMatches) {
      const tNodes = runXml.match(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g) ?? [];
      if (!tNodes.length) {
        runIdx += 1;
        continue;
      }

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
  const fontSizeHundredthPoints = Number(
    rPrXml.match(/\bsz="(\d+)"/)?.[1] ?? 0
  );

  // PPTX sz is in 1/100 points. Convert: points -> px (1pt ≈ 1.333px)
  const fontSizePx = fontSizeHundredthPoints
    ? (fontSizeHundredthPoints / 100) * 1.333
    : 18;

  const isBold = /\bb="1"/.test(rPrXml);
  const isItalic = /\bi="1"/.test(rPrXml);
  const isUnderline = /\bu="(sng|dbl)"/.test(rPrXml);

  const latin =
    rPrXml.match(/<a:latin\b[^>]*\btypeface="([^"]+)"/)?.[1] ?? "";
  const fontFamily = latin || "Arial, sans-serif";

  const colorHex = (() => {
    // Prefer run color
    const runClr =
      rPrXml.match(/<a:srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/)?.[1];
    if (runClr) return `#${runClr}`;

    // Many template runs omit explicit color; for this project we bias towards
    // a strong, near-black used on the last slide ("THANK YOU" is #0D0D0D).
    return "#0D0D0D";
  })();

  const align = (() => {
    // Paragraph alignment: <a:pPr algn="ctr|l|r|just">
    const a = pPrXml.match(/\balgn="([^"]+)"/)?.[1] ?? "";
    if (a === "ctr") return "middle";
    if (a === "r") return "end";
    return "start";
  })();

  const letterSpacingPx = (() => {
    // rPr spc is in 1/1000 em (per ECMA-376). For SVG we convert to px:
    // letterSpacingPx ~= (spc/1000) * fontSizePx
    const spc = rPrXml.match(/\bspc="(-?\d+)"/)?.[1];
    if (!spc) return 0;
    const v = Number(spc);
    if (!Number.isFinite(v)) return 0;
    return (v / 1000) * fontSizePx;
  })();

  return {
    fontSizePx,
    fontFamily,
    fontWeight: isBold ? 700 : 400,
    fontStyle: isItalic ? "italic" : "normal",
    textDecoration: isUnderline ? "underline" : "none",
    fill: colorHex,
    textAnchor: align,
    letterSpacingPx,
  };
}

function parseLeftInsetEmu(pPrXml) {
  // <a:pPr marL="..."> in EMU for margin-left. Optional.
  const m = pPrXml.match(/\bmarL="(\d+)"/);
  return m ? Number(m[1]) : 0;
}

function extractShapeTextBodyInsetsEmu(shapeXml) {
  // <a:bodyPr lIns=".." tIns=".." rIns=".." bIns=".."/>
  const bodyPr = shapeXml.match(/<a:bodyPr\b[^>]*>/)?.[0] ?? "";
  const lIns = Number(bodyPr.match(/\blIns="(\d+)"/)?.[1] ?? 0);
  const tIns = Number(bodyPr.match(/\btIns="(\d+)"/)?.[1] ?? 0);
  const rIns = Number(bodyPr.match(/\brIns="(\d+)"/)?.[1] ?? 0);
  const bIns = Number(bodyPr.match(/\bbIns="(\d+)"/)?.[1] ?? 0);
  return { lIns, tIns, rIns, bIns };
}

function normalizeWhitespaceForSvg(text) {
  // Keep intentional spaces while avoiding SVG collapsing runs too aggressively.
  // PowerPoint often uses NBSP; map to space and preserve sequences via CSS.
  return String(text ?? "")
    .replaceAll("\u00a0", " ")
    // PPTX can include literal tab runs (as seen in THANK\tYOU); map to spaces.
    .replaceAll("\t", "    ");
}

function extractPicCropFractions(picXml) {
  /**
   * Extract crop fractions from <a:srcRect l="" t="" r="" b=""/> if present.
   * Values are in 1/1000 percent (0..100000). We convert to [0..1] fractions.
   */
  const srcRect = picXml.match(/<a:srcRect\b[^>]*\/>/)?.[0] ?? "";
  if (!srcRect) return null;

  const l = Number(srcRect.match(/\bl="(\d+)"/)?.[1] ?? 0);
  const t = Number(srcRect.match(/\bt="(\d+)"/)?.[1] ?? 0);
  const r = Number(srcRect.match(/\br="(\d+)"/)?.[1] ?? 0);
  const b = Number(srcRect.match(/\bb="(\d+)"/)?.[1] ?? 0);

  const toFrac = (v) => {
    const n = Number(v);
    if (!Number.isFinite(n) || n < 0) return 0;
    return Math.min(1, n / 100000);
  };

  return { l: toFrac(l), t: toFrac(t), r: toFrac(r), b: toFrac(b) };
}

function extractPicFromPicXml(picXml) {
  // Find:
  // - blip embed rId: <a:blip r:embed="rIdX" .../>
  // - xfrm off/ext (EMU)
  const embed =
    picXml.match(/<a:blip\b[^>]*\br:embed="([^"]+)"/)?.[1] ?? null;

  const { x, y, cx, cy } = extractShapeTransformEmu(picXml);
  const crop = extractPicCropFractions(picXml);

  return { embed, x, y, cx, cy, crop };
}

function mimeFromZipPath(zipPath) {
  const ext = zipPath.split(".").pop()?.toLowerCase() ?? "";
  return ext === "png"
    ? "image/png"
    : ext === "jpg" || ext === "jpeg"
      ? "image/jpeg"
      : ext === "gif"
        ? "image/gif"
        : ext === "svg"
          ? "image/svg+xml"
          : "application/octet-stream";
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

  // Helpers for EMU->px.
  const pxX = (emu) => emuToPxX(emu, slideSize, widthPx);
  const pxY = (emu) => emuToPxY(emu, slideSize, heightPx);

  // Preserve render order by iterating spTree children when possible.
  const spTreeChildren = collectSpTreeChildren(slideXml);
  const fallbackShapeBlocks = collectShapeBlocks(slideXml);
  const drawList = spTreeChildren.length ? spTreeChildren : fallbackShapeBlocks;

  const backgroundEls = [];
  const svgEls = [];

  // Base slide background: white. (Template has white page; the photo is a picture on top.)
  backgroundEls.push(
    `<rect x="0" y="0" width="100%" height="100%" fill="#ffffff" />`
  );

  // Iterate in order.
  for (const nodeXml of drawList) {
    if (nodeXml.startsWith("<p:sp")) {
      // 1) Shape fill.
      const fill = extractShapeFillColor(nodeXml);
      if (fill) {
        const { x, y, cx, cy } = extractShapeTransformEmu(nodeXml);
        if (cx && cy) {
          svgEls.push(
            `<rect x="${pxX(x)}" y="${pxY(y)}" width="${pxX(cx)}" height="${pxY(
              cy
            )}" fill="${escapeXmlAttr(fill)}" />`
          );
        }
      }

      // 2) Text (rendered per paragraph; within paragraph per run to support letter-spacing)
      if (nodeXml.includes("<p:txBody")) {
        const { x, y, cx, cy } = extractShapeTransformEmu(nodeXml);
        if (cx && cy) {
          const { lIns, tIns } = extractShapeTextBodyInsetsEmu(nodeXml);

          const boxX = pxX(x + lIns);
          const boxY = pxY(y + tIns);
          const boxW = pxX(Math.max(0, cx - lIns));
          const boxH = pxY(Math.max(0, cy - tIns));

          const runs = extractTextRunsFromShape(nodeXml);
          if (!runs.length) continue;

          // Group runs by paragraph.
          const paragraphs = new Map();
          for (const r of runs) {
            const key = r.paragraphIndex;
            const current = paragraphs.get(key) ?? { runs: [], seed: r };
            current.runs.push(r);
            paragraphs.set(key, current);
          }

          const sortedKeys = [...paragraphs.keys()].sort((a, b) => a - b);

          // Improve baseline placement: start at top of the text box and add
          // ascent-ish offset (0.85em) per line.
          let cursorY = boxY;

          for (const pIdx of sortedKeys) {
            const p = paragraphs.get(pIdx);
            if (!p) continue;

            const seedStyle = parseTextStyle({ rPrXml: p.seed.rPrXml, pPrXml: p.seed.pPrXml });
            const marL = parseLeftInsetEmu(p.seed.pPrXml);
            const insetX = pxX(marL);

            // Approx line height: 1.12 is closer to PPT default for the template.
            const lineHeight = Math.max(1, seedStyle.fontSizePx * 1.12);
            cursorY += lineHeight;

            let xPos = boxX + insetX;
            if (seedStyle.textAnchor === "middle") xPos = boxX + boxW / 2;
            if (seedStyle.textAnchor === "end") xPos = boxX + boxW;

            // Baseline within the line.
            const yPos = Math.min(boxY + boxH, cursorY - lineHeight * 0.18);

            // Render each run as a <tspan> with its own style so we preserve
            // letter-spacing differences (critical for THANK YOU).
            const tspans = [];
            for (const run of p.runs) {
              const style = parseTextStyle({ rPrXml: run.rPrXml, pPrXml: run.pPrXml });
              const text = normalizeWhitespaceForSvg(run.text);
              if (!text) continue;

              const parts = [];
              parts.push(`<tspan`);
              parts.push(` font-family="${escapeXmlAttr(style.fontFamily)}"`);
              parts.push(` font-size="${style.fontSizePx}"`);
              parts.push(` font-weight="${style.fontWeight}"`);
              parts.push(` font-style="${escapeXmlAttr(style.fontStyle)}"`);
              parts.push(
                ` text-decoration="${escapeXmlAttr(style.textDecoration)}"`
              );
              parts.push(` fill="${escapeXmlAttr(style.fill)}"`);

              // Per-run letter spacing.
              if (Math.abs(style.letterSpacingPx) > 0.01) {
                parts.push(` letter-spacing="${style.letterSpacingPx}"`);
              }

              // Keep x on the first tspan only; others flow naturally.
              parts.push(`>`);
              parts.push(`${escapeXmlAttr(text)}`);
              parts.push(`</tspan>`);
              tspans.push(parts.join(""));
            }

            if (!tspans.length) continue;

            svgEls.push(
              [
                `<text x="${xPos}" y="${yPos}"`,
                ` text-anchor="${escapeXmlAttr(seedStyle.textAnchor)}"`,
                // Preserve whitespace and prevent run collapsing.
                ` style="white-space: pre;"`,
                `>`,
                tspans.join(""),
                `</text>`,
              ].join("")
            );
          }
        }
      }
      continue;
    }

    if (nodeXml.startsWith("<p:pic")) {
      // Picture node
      const pic = extractPicFromPicXml(nodeXml);
      if (!pic.embed) continue;

      const target = relsById.get(pic.embed);
      const zipPath = resolveSlideRelTargetToZipPath(target);
      if (!zipPath) continue;

      const imgFile = zip.file(zipPath);
      if (!imgFile) continue;

      const bytes = await imgFile.async("uint8array");
      const mime = mimeFromZipPath(zipPath);
      const b64 = safeB64FromBytes(bytes);

      const x = pxX(pic.x);
      const y = pxY(pic.y);
      const w = pxX(pic.cx);
      const h = pxY(pic.cy);

      // Honor crop (common on the last slide background photo). We implement
      // crop by clipping to the picture rect and scaling/offsetting the image
      // inside to match the cropped region.
      //
      // srcRect fractions represent how much is cropped from each side.
      // visibleW = (1 - l - r), visibleH = (1 - t - b)
      // scale image by 1/visibleW & 1/visibleH and offset by -l/-t.
      if (pic.crop && (pic.crop.l || pic.crop.r || pic.crop.t || pic.crop.b)) {
        const visibleW = Math.max(0.0001, 1 - pic.crop.l - pic.crop.r);
        const visibleH = Math.max(0.0001, 1 - pic.crop.t - pic.crop.b);
        const imgW = w / visibleW;
        const imgH = h / visibleH;

        const dx = x - imgW * pic.crop.l;
        const dy = y - imgH * pic.crop.t;

        const clipId = `clip_${slideIndex}_${Math.random().toString(16).slice(2)}`;
        // Define clip path then draw transformed image inside it.
        svgEls.push(
          [
            `<defs>`,
            `<clipPath id="${clipId}">`,
            `<rect x="${x}" y="${y}" width="${w}" height="${h}" />`,
            `</clipPath>`,
            `</defs>`,
            `<image x="${dx}" y="${dy}" width="${imgW}" height="${imgH}"`,
            ` href="data:${mime};base64,${b64}"`,
            ` preserveAspectRatio="none"`,
            ` clip-path="url(#${clipId})"`,
            ` />`,
          ].join("")
        );
      } else {
        svgEls.push(
          `<image x="${x}" y="${y}" width="${w}" height="${h}" href="data:${mime};base64,${b64}" preserveAspectRatio="none" />`
        );
      }

      continue;
    }

    // Group shapes not explicitly supported: ignore for now (rare in template)
  }

  // Compose layers: base bg, then everything in z-order.
  const svgLayers = [...backgroundEls, ...svgEls];

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
