import JSZip from "jszip";
import {
  coerceBulletLines,
  formatDateRangeDDMonYYYY,
} from "./skillFactoryModel";

/**
 * NOTE ABOUT TEMPLATE SAFETY:
 * - We only edit *newly inserted* Skill Factory slides (which are copies).
 * - We never touch the original last slide XML bytes.
 * - We never touch the original slide 1 except for date-only edit (handled elsewhere).
 */

const SLIDES_DIR = "ppt/slides";

function getSlidePath(n) {
  return `${SLIDES_DIR}/slide${n}.xml`;
}

function escapeXmlText(text) {
  return String(text ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function replaceAllLiteral(haystack, needle, replacement) {
  return haystack.split(needle).join(replacement);
}

function replaceTokenOrFail(xml, token, value) {
  const escaped = escapeXmlText(value);
  if (!xml.includes(token)) {
    throw new Error(
      `Skill Factory slide template mismatch: could not find token ${token} in slide XML.`
    );
  }
  return replaceAllLiteral(xml, token, escaped);
}

function buildBulletsXml(lines) {
  // Best-effort bullet rendering via text: bullets as "• ".
  // We avoid relying on paragraph bullet properties because the copied slide may differ.
  if (!lines.length) return "• ";
  return lines.map((l) => `• ${l}`).join("\n");
}

function buildTeamTablePlainText(table) {
  const cols = table?.columns ?? [];
  const rows = table?.rows ?? [];
  const header = cols.map((c) => String(c ?? "").trim()).join(" | ");
  const sep = cols.map(() => "---").join(" | ");
  const body = rows.map((r) => (r ?? []).map((c) => String(c ?? "").trim()).join(" | "));
  return [header, sep, ...body].join("\n");
}

/**
 * PUBLIC_INTERFACE
 * Applies Skill Factory Slide 1 content to the given slide index (newly inserted slide).
 *
 * Implementation detail:
 * - We expect the copied slide to contain placeholder tokens. For now, the shipped
 *   template slide does not, so we fallback to minimal "append label" style replacements
 *   ONLY if we detect those labels exist.
 *
 * To keep this iteration safe, we:
 * - Insert an identifiable "Skill Factory" header block by replacing existing text runs
 *   that contain known labels from the screenshot:
 *     "Digital Applications:" (append factory name)
 *     "PROJECT HIGHLIGHTS" / "PROJECT LOWLIGHTS" / "TEAM MEMBERS"
 *     "Key Activities completed in previous week" / "Key Activities planned for current week"
 *
 * If those labels are absent in the copied slide, we throw an explicit error so the UI
 * can still function for default deck; factories only apply when user clicked Add.
 */
export async function applySkillFactorySlide1Content(zip, slideIndex, form) {
  /** This is a public function. */
  const slidePath = getSlidePath(slideIndex);
  const slideFile = zip.file(slidePath);
  if (!slideFile) throw new Error(`Missing ${slidePath} in PPTX.`);

  let xml = await slideFile.async("string");

  const factoryName = String(form?.factoryName ?? "").trim();
  const sprintNumber = String(form?.sprintNumber ?? "").trim();
  const dateRange = formatDateRangeDDMonYYYY(form?.sprintStartISO, form?.sprintEndISO);

  const highlights = coerceBulletLines(form?.projectHighlightsText);
  const lowlights = coerceBulletLines(form?.projectLowlightsText);
  const completed = coerceBulletLines(form?.keyActivitiesCompletedText);
  const planned = coerceBulletLines(form?.keyActivitiesPlannedText);

  const smeName = String(form?.smeName ?? "").trim();
  const teamText = buildTeamTablePlainText(form?.teamTable);

  // Prefer token replacements if present (future-proofing).
  const tokenBased = xml.includes("{{SF_FACTORY_NAME}}");
  if (tokenBased) {
    xml = replaceTokenOrFail(xml, "{{SF_FACTORY_NAME}}", factoryName);
    xml = replaceTokenOrFail(xml, "{{SF_SPRINT_HEADER}}", `Sprint ${sprintNumber} (${dateRange})`);
    xml = replaceTokenOrFail(xml, "{{SF_HIGHLIGHTS}}", buildBulletsXml(highlights));
    xml = replaceTokenOrFail(xml, "{{SF_LOWLIGHTS}}", buildBulletsXml(lowlights));
    xml = replaceTokenOrFail(xml, "{{SF_TEAM_TABLE}}", teamText);
    xml = replaceTokenOrFail(xml, "{{SF_SME_NAME}}", smeName);
    xml = replaceTokenOrFail(xml, "{{SF_KEY_COMPLETED}}", buildBulletsXml(completed));
    xml = replaceTokenOrFail(xml, "{{SF_KEY_PLANNED}}", buildBulletsXml(planned));
    zip.file(slidePath, xml);
    return;
  }

  // Label-based replacement (best effort).
  // 1) Replace label above Name with "TATA ELXSI" using shape/mock-name targeting via XML text node.
  // Search for known label string in slide XML, e.g. "Digital Applications:", "Company Name:", etc.
  // If not found, fallback to the first <a:t> run in the upper 1/4 of the slide with candidate label text.

  let labelMatched = false;
  // List of candidate original label texts to replace
  const candidateLabels = [
    "Digital Applications:",
    "Company Name:",
    "Organization:", 
    "Label", 
    "Company:", 
    "Name:",
    "Org Name:"
  ];
  for (const label of candidateLabels) {
    if (xml.includes(label)) {
      // Replace in-place and only first occurrence
      // Find <a:t> containing this label, replace whole node text to "TATA ELXSI"
      xml = xml.replace(
        new RegExp(`<a:t[^>]*>${label.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&')}<\\/a:t>`), 
        `<a:t>TATA ELXSI</a:t>`
      );
      labelMatched = true;
      break;
    }
  }

  if (!labelMatched) {
    // fallback: Replace the first <a:t> that looks like a label (heuristic: before or near "Name" field)
    // This is safest as long as only used for slide 1 template
    // Find the first <a:t>...</a:t> tag containing a candidate label-ish string
    const labelLike = /<a:t[^>]*>([^<]*Label[^<]*|Company [^<]*|Org[^<]*|Name[^<]*:)\s*<\/a:t>/i;
    if (labelLike.test(xml)) {
      xml = xml.replace(labelLike, `<a:t>TATA ELXSI</a:t>`);
      labelMatched = true;
    }
  }
  if (!labelMatched) {
    // As a last resort: just replace the first <a:t> node entirely
    xml = xml.replace(/<a:t[^>]*>[^<]*<\/a:t>/, `<a:t>TATA ELXSI</a:t>`);
  }

  // 2) Sprint header: if 'Sprint' exists, replace the entire first occurrence of "Sprint" line heuristically.
  // Since we don't have stable tokens, we find any existing "Sprint" run and replace sprint-like text.
  if (xml.includes("Sprint") && sprintNumber && dateRange) {
    // Replace occurrences of "Sprint" followed by digits and range-ish content. Conservative regex.
    const replacement = escapeXmlText(`Sprint ${sprintNumber} (${dateRange})`);
    xml = xml.replace(/Sprint\s*\d+[\s\S]{0,40}?\)/, replacement);
    // If regex doesn't match (different formatting), at least replace first "Sprint" word with full string.
    if (!xml.includes(replacement)) {
      xml = xml.replace("Sprint", replacement);
    }
  }

  // 3) Replace section bodies by locating common headings and injecting nearby.
  // We can't reliably target shapes; use token-like markers by inserting aText after headings if a nearby placeholder exists.
  // As a safe fallback, replace any existing bullet blocks that contain '•' within the slide for respective sections.
  const bulletsHighlights = escapeXmlText(buildBulletsXml(highlights));
  const bulletsLowlights = escapeXmlText(buildBulletsXml(lowlights));
  const bulletsCompleted = escapeXmlText(buildBulletsXml(completed));
  const bulletsPlanned = escapeXmlText(buildBulletsXml(planned));

  // Replace the first bullet block after each heading by simple heuristics.
  const replaceFirstBulletBlockAfter = (heading, bulletText) => {
    const idx = xml.indexOf(heading);
    if (idx < 0) return false;

    // Search forward for the next <a:t>...</a:t> that contains a bullet symbol or is empty.
    const forward = xml.slice(idx);
    const tRe = /<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g;
    let m;
    while ((m = tRe.exec(forward)) !== null) {
      const inner = m[1] ?? "";
      if (inner.includes("•") || inner.trim() === "" || inner.includes("Bullet")) {
        const fullNode = m[0];
        const newNode = fullNode.replace(
          /(<a:t\b[^>]*>)([\s\S]*?)(<\/a:t>)/,
          `$1${bulletText}$3`
        );
        xml = xml.replace(fullNode, newNode);
        return true;
      }
    }
    return false;
  };

  replaceFirstBulletBlockAfter("PROJECT HIGHLIGHTS", bulletsHighlights);
  replaceFirstBulletBlockAfter("PROJECT LOWLIGHTS", bulletsLowlights);
  replaceFirstBulletBlockAfter("Key Activities completed", bulletsCompleted);
  replaceFirstBulletBlockAfter("Key Activities planned", bulletsPlanned);

  // 4) Team members: replace first text node after TEAM MEMBERS with a plain-text table.
  if (xml.includes("TEAM MEMBERS")) {
    const teamEsc = escapeXmlText(teamText);
    replaceFirstBulletBlockAfter("TEAM MEMBERS", teamEsc);
  }

  // 5) SME name: attempt to replace a nearby "SME" or blank placeholder after team block.
  if (smeName) {
    if (xml.includes("SME")) {
      // Replace "SME" line (best-effort) with "SME: {name}".
      xml = xml.replace(/SME\s*:?\s*[^<]{0,50}/, `SME: ${escapeXmlText(smeName)}`);
    } else {
      // If there's no SME label, do nothing (still acceptable).
    }
  }

  zip.file(slidePath, xml);
}

/**
 * PUBLIC_INTERFACE
 * Applies placeholders for slides 2–4 (scaffold). For now, ensure each slide includes a visible
 * placeholder title to indicate it is not implemented.
 */
export async function applySkillFactoryPlaceholderSlides(zip, slideIndexes, label) {
  /** This is a public function. */
  const safeLabel = String(label ?? "Skill Factory").trim() || "Skill Factory";
  const titles = [
    `${safeLabel} — Slide 2 (Placeholder)`,
    `${safeLabel} — Slide 3 (Placeholder)`,
    `${safeLabel} — Slide 4 (Placeholder)`,
  ];

  for (let i = 0; i < 3; i += 1) {
    const slideIndex = slideIndexes[i];
    const slidePath = getSlidePath(slideIndex);
    const slideFile = zip.file(slidePath);
    if (!slideFile) continue;

    let xml = await slideFile.async("string");

    // Replace first <a:t> node content with the placeholder title (best-effort).
    // This is safe for copied slides only.
    xml = xml.replace(
      /<a:t\b[^>]*>[\s\S]*?<\/a:t>/,
      `<a:t>${escapeXmlText(titles[i])}</a:t>`
    );

    zip.file(slidePath, xml);
  }
}
