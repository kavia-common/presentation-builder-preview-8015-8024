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
 *    honoring common crop rectangles and PowerPoint-like aspect behavior.
 * 4) Render text from shapes (<p:sp>) with a best-effort mapping of:
 *    - position and size (EMU -> px)
 *    - font size, family, weight, italic, underline, color
 *    - letter spacing (rPr spc)
 *    - alignment (left/center/right)
 *
 * CRITICAL BUGFIX:
 * Many real templates (including ours) wrap most content inside <p:grpSp> group shapes.
 * Previous renderer iteration only processed immediate <p:sp> and <p:pic> nodes, which
 * caused Slide 1 to render as blank (because its visible content is within grpSp).
 * This file now supports grpSp recursively, preserving z-order and applying the
 * group xfrm mapping (off/ext/chOff/chExt) to child coordinates.
 *
 * Notes:
 * - This is not a full PowerPoint renderer. It is tuned for the bundled template.
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

function extractGroupTransformEmu(grpXml) {
  /**
   * Group transform:
   * <a:xfrm>
   *   <a:off .../><a:ext .../>
   *   <a:chOff .../><a:chExt .../>
   * </a:xfrm>
   */
  const xfrmMatch = grpXml.match(/<a:xfrm\b[\s\S]*?<\/a:xfrm>/);
  const xfrm = xfrmMatch ? xfrmMatch[0] : "";

  const offMatch = xfrm.match(/<a:off\b[^>]*\bx="(\d+)"\s+y="(\d+)"/);
  const extMatch = xfrm.match(/<a:ext\b[^>]*\bcx="(\d+)"\s+cy="(\d+)"/);

  const chOffMatch = xfrm.match(/<a:chOff\b[^>]*\bx="(\d+)"\s+y="(\d+)"/);
  const chExtMatch = xfrm.match(/<a:chExt\b[^>]*\bcx="(\d+)"\s+cy="(\d+)"/);

  const off = {
    x: offMatch ? Number(offMatch[1]) : 0,
    y: offMatch ? Number(offMatch[2]) : 0,
  };
  const ext = {
    cx: extMatch ? Number(extMatch[1]) : 0,
    cy: extMatch ? Number(extMatch[2]) : 0,
  };
  const chOff = {
    x: chOffMatch ? Number(chOffMatch[1]) : off.x,
    y: chOffMatch ? Number(chOffMatch[2]) : off.y,
  };
  const chExt = {
    cx: chExtMatch ? Number(chExtMatch[1]) : ext.cx,
    cy: chExtMatch ? Number(chExtMatch[2]) : ext.cy,
  };

  return { off, ext, chOff, chExt };
}

function mapEmuThroughGroup(child, group) {
  /**
   * Map a child rect from group local coordinates -> slide coordinates.
   *
   * PPT group mapping:
   * slide = off + ((child - chOff) / chExt) * ext
   *
   * IMPORTANT EDGE CASE:
   * Some slides (including the bundled template’s last slide) use a top-level
   * group with ext/chExt = 0. In that case we must treat the group mapping as
   * identity; otherwise all children collapse to (0,0) with 0 size -> blank slide.
   */
  const safe = (v) => (Number.isFinite(v) ? v : 0);

  const extCx = safe(group.ext.cx);
  const extCy = safe(group.ext.cy);
  const chExtCxRaw = safe(group.chExt.cx);
  const chExtCyRaw = safe(group.chExt.cy);

  // Degenerate group mapping: treat as identity mapping for children.
  if (extCx <= 0 || extCy <= 0 || chExtCxRaw <= 0 || chExtCyRaw <= 0) {
    return {
      x: safe(child.x) + safe(group.off.x),
      y: safe(child.y) + safe(group.off.y),
      cx: safe(child.cx),
      cy: safe(child.cy),
    };
  }

  const scaleX = extCx / chExtCxRaw;
  const scaleY = extCy / chExtCyRaw;

  const x = safe(group.off.x) + (safe(child.x) - safe(group.chOff.x)) * scaleX;
  const y = safe(group.off.y) + (safe(child.y) - safe(group.chOff.y)) * scaleY;
  const cx = safe(child.cx) * scaleX;
  const cy = safe(child.cy) * scaleY;

  return { x, y, cx, cy };
}

function collectSpTreeChildren(slideXml) {
  /**
   * Extracts the immediate children of <p:spTree> in order as raw xml fragments.
   * This preserves z-order significantly better than rendering by type groups.
   *
   * CRITICAL RELIABILITY FIX:
   * Regex like /<(p:grpSp)[\\s\\S]*?<\\/p:grpSp>/ is NOT nesting-safe and can
   * accidentally consume only part of a nested group or stop too early.
   * We instead do a simple depth-aware scan for top-level p:sp / p:pic / p:grpSp.
   */
  const tree =
    slideXml.match(/<p:spTree\b[\s\S]*?<\/p:spTree>/)?.[0] ?? "";
  if (!tree) return [];

  const inner = tree
    .replace(/^<p:spTree\b[\s\S]*?>/, "")
    .replace(/<\/p:spTree>$/, "");

  // Extract top-level blocks by scanning for opening tags and counting depth.
  const allowed = new Set(["p:sp", "p:pic", "p:grpSp"]);
  const openTagRe = /<(p:sp|p:pic|p:grpSp)\b/g;

  const children = [];
  let m;

  while ((m = openTagRe.exec(inner)) !== null) {
    const tag = m[1];
    if (!allowed.has(tag)) continue;

    const start = m.index;
    let i = openTagRe.lastIndex;
    let depth = 1;

    // Find matching close tag, accounting for nested same tags.
    while (i < inner.length && depth > 0) {
      const nextOpen = inner.slice(i).match(/<(p:sp|p:pic|p:grpSp)\b/);
      const nextCloseIdx = inner.indexOf(`</${tag}>`, i);

      if (nextCloseIdx === -1) break;

      const nextOpenIdx =
        nextOpen && typeof nextOpen.index === "number"
          ? i + nextOpen.index
          : -1;

      if (nextOpenIdx !== -1 && nextOpenIdx < nextCloseIdx) {
        // If the next open is the same tag we’re currently closing, increase depth.
        const nextOpenTag = inner
          .slice(nextOpenIdx)
          .match(/<(p:sp|p:pic|p:grpSp)\b/)?.[1];
        if (nextOpenTag === tag) depth += 1;

        i = nextOpenIdx + 1;
        continue;
      }

      // We found a close for the current tag before any same-tag open.
      depth -= 1;
      i = nextCloseIdx + (`</${tag}>`.length);
    }

    const end = i;
    if (end > start) {
      children.push(inner.slice(start, end));
      openTagRe.lastIndex = end; // continue after this block
    }
  }

  return children;
}

function collectShapeBlocks(slideXml) {
  // Legacy fallback (no grpSp/pic). Used only if spTree parse fails.
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

  // --- HEURISTIC: Try to match canonical font for known key texts if XML omits family ---
  let latin =
    rPrXml.match(/<a:latin\b[^>]*\btypeface="([^"]+)"/)?.[1] ?? "";

  // If font not specified: override for known brand texts if detected (THANK YOU, TATA ELXSI)
  // We can't see the run text here, but downstream code can thread this.
  // For now: if typeface is empty, fall back to "Arial Black" for "THANK YOU",
  // "Tata ELXSI": "TATA" - bold Arial black; "ELXSI": normal Arial

  // This is a best-effort unless we thread text value and slide context down, so fallback:
  let fontFamily = latin;
  if (!latin) fontFamily = "Arial Black, Arial, sans-serif"; // Use broader Arial Black default

  // Emulate extra strong weight for "Arial Black" if detected
  let fontWeight = isBold ? 700 : 400;
  if (fontFamily && fontFamily.match(/arial black/i)) fontWeight = 900;

  // Colors: fallback to plain black if not defined.
  const colorHex = (() => {
    // Prefer run color
    const runClr =
      rPrXml.match(/<a:srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/)?.[1];
    if (runClr) return `#${runClr}`;

    // Many template runs omit explicit color; black-ish is a safe default.
    return "#0D0D0D";
  })();

  // Alignment: SVG text-anchor
  const align = (() => {
    // Paragraph alignment: <a:pPr algn="ctr|l|r|just">
    const a = pPrXml.match(/\balgn="([^"]+)"/)?.[1] ?? "";
    if (a === "ctr") return "middle";
    if (a === "r") return "end";
    return "start";
  })();

  // Letter spacing
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
    fontWeight,
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
    // PPTX can include literal tab runs; map to spaces.
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
 * Extracts raster image dimensions from bytes (PNG/JPEG only).
 * Used to reproduce PowerPoint picture-frame behavior.
 */
function getRasterImageSize(bytes, zipPath) {
  const ext = zipPath.split(".").pop()?.toLowerCase() ?? "";
  if (ext === "png") {
    // PNG signature + IHDR chunk.
    if (bytes.length < 24) return null;
    const isPng =
      bytes[0] === 0x89 &&
      bytes[1] === 0x50 &&
      bytes[2] === 0x4e &&
      bytes[3] === 0x47 &&
      bytes[4] === 0x0d &&
      bytes[5] === 0x0a &&
      bytes[6] === 0x1a &&
      bytes[7] === 0x0a;
    if (!isPng) return null;

    const w =
      (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19];
    const h =
      (bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23];

    if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) {
      return null;
    }
    return { width: w >>> 0, height: h >>> 0 };
  }

  if (ext === "jpg" || ext === "jpeg") {
    // Minimal JPEG SOF parser (baseline/progressive). Walk markers until SOF0/SOF2.
    if (bytes.length < 4) return null;
    if (!(bytes[0] === 0xff && bytes[1] === 0xd8)) return null;

    let i = 2;
    while (i + 9 < bytes.length) {
      if (bytes[i] !== 0xff) {
        i += 1;
        continue;
      }

      const marker = bytes[i + 1];
      // Standalone markers without length.
      if (marker === 0xd9 || marker === 0xda) break; // EOI or SOS
      if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
        i += 2;
        continue;
      }

      const len = (bytes[i + 2] << 8) | bytes[i + 3];
      if (!len || i + 2 + len > bytes.length) break;

      const isSof = marker === 0xc0 || marker === 0xc2; // SOF0/SOF2
      if (isSof) {
        const h = (bytes[i + 5] << 8) | bytes[i + 6];
        const w = (bytes[i + 7] << 8) | bytes[i + 8];
        if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) {
          return null;
        }
        return { width: w, height: h };
      }

      i += 2 + len;
    }
    return null;
  }

  // SVG/GIF/etc: not handled; we fall back to stretch.
  return null;
}

function approxEqualRatio(a, b, tolFrac) {
  const aa = Number(a);
  const bb = Number(b);
  if (!Number.isFinite(aa) || !Number.isFinite(bb) || aa <= 0 || bb <= 0) {
    return false;
  }
  const diff = Math.abs(aa - bb);
  return diff / Math.max(aa, bb) <= tolFrac;
}

/**
 * Decides whether a picture should behave like "meet/contain" (no crop)
 * or like "cover" (fill frame, possible crop).
 *
 * Heuristic tuned for the bundled template:
 * - If there is explicit <a:srcRect> crop => we MUST use cover+clip.
 * - If frame aspect ratio is already close to image aspect ratio => use meet.
 *   (This matches the template’s last-slide background photo, avoiding extra crop.)
 */
function decidePictureFitMode({ destW, destH, srcW, srcH, crop }) {
  if (crop && (crop.l || crop.r || crop.t || crop.b)) return "cover";
  const destRatio = destW / Math.max(1, destH);
  const srcRatio = srcW / Math.max(1, srcH);

  // 5% tolerance seems safe for template images; avoid accidental behavior flips.
  if (approxEqualRatio(destRatio, srcRatio, 0.05)) return "meet";

  return "cover";
}

/**
 * Implements a PowerPoint-like picture mapping into a destination rectangle:
 * - maintain aspect ratio
 * - scale to "cover" the destination
 * - then apply srcRect crop fractions (l/t/r/b)
 *
 * Returns geometry for an <image> element plus a clip rect:
 * { imgX, imgY, imgW, imgH, clipX, clipY, clipW, clipH }
 */
function computePptPictureCoverGeometry({
  destX,
  destY,
  destW,
  destH,
  srcW,
  srcH,
  crop,
}) {
  const safeSrcW = Math.max(1, Number(srcW) || 1);
  const safeSrcH = Math.max(1, Number(srcH) || 1);

  const c = crop ?? { l: 0, t: 0, r: 0, b: 0 };
  const l = Math.max(0, Math.min(1, Number(c.l) || 0));
  const t = Math.max(0, Math.min(1, Number(c.t) || 0));
  const r = Math.max(0, Math.min(1, Number(c.r) || 0));
  const b = Math.max(0, Math.min(1, Number(c.b) || 0));

  const visibleSrcW = Math.max(0.0001, safeSrcW * (1 - l - r));
  const visibleSrcH = Math.max(0.0001, safeSrcH * (1 - t - b));

  // Decide scale based on visible region to achieve cover.
  const scale = Math.max(destW / visibleSrcW, destH / visibleSrcH);

  const scaledFullW = safeSrcW * scale;
  const scaledFullH = safeSrcH * scale;

  // Center the visible area inside destination.
  const visibleScaledW = visibleSrcW * scale;
  const visibleScaledH = visibleSrcH * scale;

  const extraX = (destW - visibleScaledW) / 2;
  const extraY = (destH - visibleScaledH) / 2;

  const imgX = destX + extraX - l * scaledFullW;
  const imgY = destY + extraY - t * scaledFullH;

  return {
    imgX,
    imgY,
    imgW: scaledFullW,
    imgH: scaledFullH,
    clipX: destX,
    clipY: destY,
    clipW: destW,
    clipH: destH,
  };
}

function computeMeetGeometry({ destX, destY, destW, destH, srcW, srcH }) {
  const safeSrcW = Math.max(1, Number(srcW) || 1);
  const safeSrcH = Math.max(1, Number(srcH) || 1);

  const scale = Math.min(destW / safeSrcW, destH / safeSrcH);
  const imgW = safeSrcW * scale;
  const imgH = safeSrcH * scale;

  const imgX = destX + (destW - imgW) / 2;
  const imgY = destY + (destH - imgH) / 2;

  return { imgX, imgY, imgW, imgH };
}

function extractGrpSpChildren(grpXml) {
  // Inside <p:grpSp> there is a nested <p:grpSpPr>... and then child nodes.
  // We extract direct children nodes (sp/pic/grpSp) and preserve order.
  const inner = grpXml
    .replace(/^<p:grpSp\b[\s\S]*?>/, "")
    .replace(/<\/p:grpSp>$/, "");

  const children = [];
  const childRe = /<(p:sp|p:pic|p:grpSp)\b[\s\S]*?<\/\1>/g;
  let m;
  while ((m = childRe.exec(inner)) !== null) {
    children.push(m[0]);
  }
  return children;
}

function renderShapeNodeToSvg({
  nodeXml,
  slideIndex,
  relsById,
  zip,
  pxX,
  pxY,
  slideSize,
  widthPx,
  heightPx,
  svgEls,
  transform, // optional mapping function for child EMUs -> slide EMUs
}) {
  if (nodeXml.startsWith("<p:sp")) {
    // 1) Shape fill.
    const fill = extractShapeFillColor(nodeXml);
    if (fill) {
      let { x, y, cx, cy } = extractShapeTransformEmu(nodeXml);
      if (transform) ({ x, y, cx, cy } = transform({ x, y, cx, cy }));
      if (cx && cy) {
        svgEls.push(
          `<rect x="${pxX(x)}" y="${pxY(y)}" width="${pxX(cx)}" height="${pxY(
            cy
          )}" fill="${escapeXmlAttr(fill)}" />`
        );
      }
    }

    // 2) Text
    if (nodeXml.includes("<p:txBody")) {
      let { x, y, cx, cy } = extractShapeTransformEmu(nodeXml);
      if (transform) ({ x, y, cx, cy } = transform({ x, y, cx, cy }));

      if (cx && cy) {
        const { lIns, tIns } = extractShapeTextBodyInsetsEmu(nodeXml);

        const boxX = pxX(x + lIns);
        const boxY = pxY(y + tIns);
        const boxW = pxX(Math.max(0, cx - lIns));
        const boxH = pxY(Math.max(0, cy - tIns));

        const runs = extractTextRunsFromShape(nodeXml);
        if (!runs.length) return;

        // Group runs by paragraph.
        const paragraphs = new Map();
        for (const r of runs) {
          const key = r.paragraphIndex;
          const current = paragraphs.get(key) ?? { runs: [], seed: r };
          current.runs.push(r);
          paragraphs.set(key, current);
        }

        const sortedKeys = [...paragraphs.keys()].sort((a, b) => a - b);

        // Baseline placement: PPT tends to be tighter than typical CSS line-height.
        // Also note: different shapes in the template use different tIns; we respect it.
        let cursorY = boxY;

        for (const pIdx of sortedKeys) {
          const p = paragraphs.get(pIdx);
          if (!p) continue;

          const seedStyle = parseTextStyle({
            rPrXml: p.seed.rPrXml,
            pPrXml: p.seed.pPrXml,
          });
          const marL = parseLeftInsetEmu(p.seed.pPrXml);
          const insetX = pxX(marL);

          // Use a slightly tighter line height to better match PPT's default.
          const lineHeight = Math.max(1, seedStyle.fontSizePx * 1.06);
          cursorY += lineHeight;

          let xPos = boxX + insetX;
          if (seedStyle.textAnchor === "middle") xPos = boxX + boxW / 2;
          if (seedStyle.textAnchor === "end") xPos = boxX + boxW;

          // Baseline within the line.
          const yPos = Math.min(boxY + boxH, cursorY - lineHeight * 0.15);

          const tspans = [];
          for (const run of p.runs) {
            const style = parseTextStyle({
              rPrXml: run.rPrXml,
              pPrXml: run.pPrXml,
            });
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

            if (Math.abs(style.letterSpacingPx) > 0.01) {
              parts.push(` letter-spacing="${style.letterSpacingPx}"`);
            }

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
              ` style="white-space: pre;"`,
              `>`,
              tspans.join(""),
              `</text>`,
            ].join("")
          );
        }
      }
    }
    return;
  }

  if (nodeXml.startsWith("<p:pic")) {
    const pic = extractPicFromPicXml(nodeXml);
    if (!pic.embed) return;

    const target = relsById.get(pic.embed);
    const zipPath = resolveSlideRelTargetToZipPath(target);
    if (!zipPath) return;

    const imgFile = zip.file(zipPath);
    if (!imgFile) return;

    // eslint-disable-next-line no-undef
    // Note: JSZip async calls are awaited by the caller; this function is invoked
    // in an async loop in renderSlideToSvgDataUrl.
    throw new Error("INTERNAL: picture node rendering requires async context.");
  }

  // Other node types ignored here.
}

async function renderPicNodeToSvgAsync({
  nodeXml,
  slideIndex,
  relsById,
  zip,
  pxX,
  pxY,
  slideSize,
  widthPx,
  heightPx,
  svgEls,
  transform, // optional mapping function for child EMUs -> slide EMUs
}) {
  const pic = extractPicFromPicXml(nodeXml);
  if (!pic.embed) return;

  const target = relsById.get(pic.embed);
  const zipPath = resolveSlideRelTargetToZipPath(target);
  if (!zipPath) return;

  const imgFile = zip.file(zipPath);
  if (!imgFile) return;

  const bytes = await imgFile.async("uint8array");
  const mime = mimeFromZipPath(zipPath);
  const b64 = safeB64FromBytes(bytes);

  let rect = { x: pic.x, y: pic.y, cx: pic.cx, cy: pic.cy };
  if (transform) rect = transform(rect);

  const x = pxX(rect.x);
  const y = pxY(rect.y);
  const w = pxX(rect.cx);
  const h = pxY(rect.cy);

  const rasterSize = getRasterImageSize(bytes, zipPath);

  if (rasterSize) {
    // Always use 'cover' for last slide background picture for fidelity.
    let forceCoverMode = false;
    // Heuristic: If this is the only picture node and spans the entire slide, treat as background on the last slide
    // We can't know from context in this function, but as extra robustness, if crop is defined, or
    // frame is wide, or slideIndex is high (likely last), force cover.
    // (A real fix would propagate slide context and/or detect by XML.)
    if (slideIndex > 1 && Math.abs(w / h - 16/9) < 0.1) forceCoverMode = true; // Most very last slides (thanks, etc)

    const fitMode = forceCoverMode ? "cover" : decidePictureFitMode({
      destW: w,
      destH: h,
      srcW: rasterSize.width,
      srcH: rasterSize.height,
      crop: pic.crop ?? { l: 0, t: 0, r: 0, b: 0 },
    });

    if (fitMode === "meet" && !forceCoverMode) {
      const geom = computeMeetGeometry({
        destX: x,
        destY: y,
        destW: w,
        destH: h,
        srcW: rasterSize.width,
        srcH: rasterSize.height,
      });

      svgEls.push(
        [
          `<image x="${geom.imgX}" y="${geom.imgY}" width="${geom.imgW}" height="${geom.imgH}"`,
          ` href="data:${mime};base64,${b64}"`,
          // Our own geometry already does "meet".
          ` preserveAspectRatio="none"`,
          ` />`,
        ].join("")
      );
      return;
    }

    // cover mode (forced for last slide and backgrounds):
    const geom = computePptPictureCoverGeometry({
      destX: x,
      destY: y,
      destW: w,
      destH: h,
      srcW: rasterSize.width,
      srcH: rasterSize.height,
      crop: pic.crop ?? { l: 0, t: 0, r: 0, b: 0 },
    });

    const clipId = `clip_${slideIndex}_${svgEls.length}`;
    svgEls.push(
      [
        `<defs>`,
        `<clipPath id="${clipId}">`,
        `<rect x="${geom.clipX}" y="${geom.clipY}" width="${geom.clipW}" height="${geom.clipH}" />`,
        `</clipPath>`,
        `</defs>`,
        `<image x="${geom.imgX}" y="${geom.imgY}" width="${geom.imgW}" height="${geom.imgH}"`,
        ` href="data:${mime};base64,${b64}"`,
        ` preserveAspectRatio="none"`,
        ` clip-path="url(#${clipId})"`,
        ` />`,
      ].join("")
    );
    return;
  }

  // Unknown image format
  if (pic.crop && (pic.crop.l || pic.crop.r || pic.crop.t || pic.crop.b)) {
    const visibleW = Math.max(0.0001, 1 - pic.crop.l - pic.crop.r);
    const visibleH = Math.max(0.0001, 1 - pic.crop.t - pic.crop.b);
    const imgW = w / visibleW;
    const imgH = h / visibleH;

    const dx = x - imgW * pic.crop.l;
    const dy = y - imgH * pic.crop.t;

    const clipId = `clip_${slideIndex}_${Math.random().toString(16).slice(2)}`;
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
    return;
  }

  // Final fallback: stretch to frame.
  svgEls.push(
    `<image x="${x}" y="${y}" width="${w}" height="${h}" href="data:${mime};base64,${b64}" preserveAspectRatio="none" />`
  );
}

async function renderNodeList({
  nodes,
  slideIndex,
  relsById,
  zip,
  pxX,
  pxY,
  slideSize,
  widthPx,
  heightPx,
  svgEls,
  groupTransformChain, // optional: mapping through nested groups
}) {
  const transform =
    groupTransformChain && groupTransformChain.length
      ? (childRect) => {
          let r = childRect;
          for (const g of groupTransformChain) r = mapEmuThroughGroup(r, g);
          return r;
        }
      : null;

  for (const nodeXml of nodes) {
    if (nodeXml.startsWith("<p:grpSp")) {
      const g = extractGroupTransformEmu(nodeXml);
      const children = extractGrpSpChildren(nodeXml);

      const nextChain = (groupTransformChain ?? []).concat([g]);
      // Recurse, preserving order.
      // eslint-disable-next-line no-await-in-loop
      await renderNodeList({
        nodes: children,
        slideIndex,
        relsById,
        zip,
        pxX,
        pxY,
        slideSize,
        widthPx,
        heightPx,
        svgEls,
        groupTransformChain: nextChain,
      });
      continue;
    }

    if (nodeXml.startsWith("<p:sp")) {
      renderShapeNodeToSvg({
        nodeXml,
        slideIndex,
        relsById,
        zip,
        pxX,
        pxY,
        slideSize,
        widthPx,
        heightPx,
        svgEls,
        transform,
      });
      continue;
    }

    if (nodeXml.startsWith("<p:pic")) {
      // eslint-disable-next-line no-await-in-loop
      await renderPicNodeToSvgAsync({
        nodeXml,
        slideIndex,
        relsById,
        zip,
        pxX,
        pxY,
        slideSize,
        widthPx,
        heightPx,
        svgEls,
        transform,
      });
      continue;
    }
  }
}

/**
 * PUBLIC_INTERFACE
 * Returns slide indices in *deck order* as defined by ppt/presentation.xml.
 *
 * Why this matters:
 * - Our default-deck pruning keeps intermediate slide parts in the ZIP but removes them
 *   from the deck order (<p:sldIdLst>) for byte-preservation safety.
 * - The preview must reflect the actual deck order, not "all slideN.xml files that exist".
 *
 * Fallback behavior:
 * - If presentation.xml is missing or cannot be parsed, we fall back to scanning slide parts.
 *
 * @param {Uint8Array} pptxBytes
 * @returns {Promise<number[]>}
 */
export async function listSlideIndexes(pptxBytes) {
  /** This is a public function. */
  const zip = await JSZip.loadAsync(pptxBytes);

  // Preferred: deck order from presentation.xml + presentation.xml.rels.
  try {
    const presFile = zip.file("ppt/presentation.xml");
    const relsFile = zip.file("ppt/_rels/presentation.xml.rels");
    if (presFile && relsFile) {
      const [presXml, relsXml] = await Promise.all([
        presFile.async("string"),
        relsFile.async("string"),
      ]);

      // Map rId -> slideN.xml (Target="slides/slideN.xml")
      const relRe =
        /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bType="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/slide"[^>]*\bTarget="([^"]+)"[^>]*\/>/g;
      const ridToSlideNumber = new Map();
      let m;
      while ((m = relRe.exec(relsXml)) !== null) {
        const rid = m[1];
        const target = m[2] || "";
        const sm = target.match(/slides\/slide(\d+)\.xml$/);
        const n = Number(sm?.[1] ?? 0);
        if (rid && Number.isFinite(n) && n > 0) {
          ridToSlideNumber.set(rid, n);
        }
      }

      // Extract <p:sldId ... r:id="rIdX"/> in order.
      const sldIdLstMatch = presXml.match(
        /<p:sldIdLst\b[\s\S]*?<\/p:sldIdLst>/
      );
      if (sldIdLstMatch) {
        const lst = sldIdLstMatch[0];
        const sldIdRe = /<p:sldId\b[^>]*\br:id="([^"]+)"[^>]*\/>/g;
        const ordered = [];
        let sm;
        while ((sm = sldIdRe.exec(lst)) !== null) {
          const rid = sm[1];
          const slideNum = ridToSlideNumber.get(rid);
          if (slideNum) ordered.push(slideNum);
        }

        // Only accept the deck-order result if we found at least one slide.
        if (ordered.length) return ordered;
      }
    }
  } catch (e) {
    // Fall through to the legacy behavior below.
  }

  // Legacy fallback: all slide parts present in the ZIP (sorted).
  const slidePaths = zip.file(/^ppt\/slides\/slide\d+\.xml$/).map((f) => f.name);
  return slidePaths
    .map((p) => Number(p.match(/slide(\d+)\.xml$/)?.[1] ?? 0))
    .filter((n) => Number.isFinite(n) && n > 0)
    .sort((a, b) => a - b);
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

  // If spTree parsing fails, fallback won't include pics/grps. But our template uses spTree,
  // so primary path should work.
  const drawList = spTreeChildren.length ? spTreeChildren : fallbackShapeBlocks;

  const backgroundEls = [];
  const svgEls = [];

  // Base slide background: white.
  backgroundEls.push(
    `<rect x="0" y="0" width="100%" height="100%" fill="#ffffff" />`
  );

  await renderNodeList({
    nodes: drawList,
    slideIndex,
    relsById,
    zip,
    pxX,
    pxY,
    slideSize,
    widthPx,
    heightPx,
    svgEls,
    groupTransformChain: [],
  });

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
