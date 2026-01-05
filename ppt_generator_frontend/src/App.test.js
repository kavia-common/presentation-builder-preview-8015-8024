import React from "react";
import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import JSZip from "jszip";
import App from "./App";
import { assertLastSlideUnchanged, updatePptxDateOnly } from "./pptx/templateEditor";
import { renderSlideToSvgDataUrl } from "./pptx/pptxSlideRenderer";

// Increase default test timeout: JSZip + SVG slide rendering can be slower in CI.
// This is a regression guard only; production behavior is unaffected.
jest.setTimeout(20000);

async function makeMinimalPptxArrayBuffer() {
  // Minimal PPTX-like zip with slide1.xml containing the expected strict date runs.
  // Also include a "last slide" with the highest slide number so we can assert it remains untouched.
  //
  // NOTE: The app now prunes the deck to only Slide 1 + last slide by editing
  // ppt/presentation.xml and ppt/_rels/presentation.xml.rels, so those parts must exist.
  const zip = new JSZip();

  zip.file(
    "ppt/slides/slide1.xml",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
      "<p:cSld><p:spTree>",
      '<p:sp><p:nvSpPr><p:cNvPr id="1" name="TextBox 4"/></p:nvSpPr><p:txBody>',
      "<a:p>",
      "<a:r><a:t>Date</a:t></a:r>",
      "<a:r><a:t> </a:t></a:r>",
      "<a:r><a:t>:</a:t></a:r>",
      // IMPORTANT: Use real NBSP characters to match strict template expectations in templateEditor.js
      "<a:r><a:t>\u00a0 24\u00a0</a:t></a:r>",
      "<a:r><a:t>Dec</a:t></a:r>",
      "<a:r><a:t> </a:t></a:r>",
      "<a:r><a:t>202</a:t></a:r>",
      "<a:r><a:t>5</a:t></a:r>",
      "</a:p>",
      "</p:txBody></p:sp>",
      "</p:spTree></p:cSld>",
      "</p:sld>",
    ].join("")
  );

  zip.file("ppt/slides/slide9.xml", "<last-slide>DO NOT TOUCH</last-slide>");

  zip.file(
    "ppt/presentation.xml",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
      "<p:sldIdLst>",
      '<p:sldId id="256" r:id="rId1"/>',
      '<p:sldId id="257" r:id="rId2"/>',
      "</p:sldIdLst>",
      "</p:presentation>",
    ].join("")
  );

  zip.file(
    "ppt/_rels/presentation.xml.rels",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>',
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide9.xml"/>',
      "</Relationships>",
    ].join("")
  );

  const bytes = await zip.generateAsync({ type: "uint8array" });
  return bytes.buffer;
}

describe("PPTX preview regeneration", () => {
  // Date-only regression suite (Slide 1 date editable; no label/name edits).
  test("changing the date regenerates the blob URL and keeps a visible Open/Download link", async () => {
    const prevFetch = global.fetch;

    // Mock only the template PPTX fetch for this test.
    global.fetch = jest.fn(async (input) => {
      const url = String(input || "");
      if (url === "/assets/template.pptx") {
        return {
          ok: true,
          status: 200,
          arrayBuffer: async () => await makeMinimalPptxArrayBuffer(),
        };
      }
      // Fall back to any existing polyfill for other URLs.
      if (prevFetch) return await prevFetch(input);
      throw new Error(`Unhandled fetch in test: ${url}`);
    });

    render(<App />);

    // Wait until generation has produced a PPTX URL (link is rendered when bytes exist).
    const firstLink = await screen.findByRole(
      "link",
      { name: "Download / Open PPTX" },
      { timeout: 8000 }
    );
    const firstHref = firstLink.getAttribute("href");
    expect(firstHref).toBeTruthy();

    // Changing date should create a new blob URL (new href).
    const dateInput = screen.getByLabelText("Slide 1 Date");
    fireEvent.change(dateInput, { target: { value: "2026-01-06" } });

    await waitFor(
      () => {
        const nextLink = screen.getByRole("link", {
          name: "Download / Open PPTX",
        });
        const nextHref = nextLink.getAttribute("href");
        expect(nextHref).toBeTruthy();
        expect(nextHref).not.toEqual(firstHref);
      },
      { timeout: 5000 }
    );

    // Carousel nav should exist (regression guard).
    expect(screen.getByRole("button", { name: "Prev" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Next" })).toBeInTheDocument();

    // Default deck should now be exactly 2 slides (Slide 1 + last).
    await waitFor(
      () => {
        const indicator = screen.getByLabelText("Slide indicator");
        expect(indicator.textContent).toMatch(/\/\s*2/);
      },
      { timeout: 5000 }
    );

    global.fetch = prevFetch;
  });

  test("strict invariant: last slide remains byte-for-byte unchanged", async () => {
    const template = await makeMinimalPptxArrayBuffer();
    const { updatedPptxBytes } = await updatePptxDateOnly(template, "2026-01-06");
    await expect(assertLastSlideUnchanged(template, updatedPptxBytes)).resolves.toEqual(true);
  });

  test("renderer regression: slide 1 (bundled template) renders non-empty SVG (grpSp supported)", async () => {
    // This guards the reported regression: Slide 1 disappeared when grpSp nodes were skipped.
    // The renderer is read-only, so this cannot affect PPTX invariants.
    const res = await fetch("/assets/template.pptx");
    const buf = await res.arrayBuffer();
    const bytes = new Uint8Array(buf);

    const out = await renderSlideToSvgDataUrl(bytes, 1, { widthPx: 520 });
    expect(out.dataUrl).toMatch(/^data:image\/svg\+xml;base64,/);

    const svg = decodeURIComponent(
      escape(atob(out.dataUrl.split(",")[1] || ""))
    );

    // Template slide 1 is largely picture-backed; ensure we at least have an image element.
    expect(svg.includes("<image")).toBe(true);
    // Slide 1 includes a visible "Date" label in the left block.
    expect(svg.includes(">Date<") || svg.includes("Date")).toBe(true);
  });
});
