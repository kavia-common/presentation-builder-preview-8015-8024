import React from "react";
import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import JSZip from "jszip";
import App from "./App";
import { assertLastSlideUnchanged, updatePptxDateOnly } from "./pptx/templateEditor";

async function makeMinimalPptxArrayBuffer() {
  // Minimal PPTX-like zip with slide1.xml containing the expected strict date runs.
  // Also include a "last slide" with the highest slide number so we can assert it remains untouched.
  // Our invariant checker detects the last slide dynamically (highest slideN.xml).
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

  const bytes = await zip.generateAsync({ type: "uint8array" });
  return bytes.buffer;
}

describe("PPTX preview regeneration", () => {
  test("changing the date regenerates the blob URL and refreshes the preview iframe src", async () => {
    global.fetch = jest.fn(async () => ({
      ok: true,
      arrayBuffer: async () => await makeMinimalPptxArrayBuffer(),
    }));

    render(<App />);

    const iframe = await screen.findByTitle("PPTX Preview");
    const firstSrc = iframe.getAttribute("src");
    expect(firstSrc).toBeTruthy();

    const dateInput = screen.getByLabelText("Slide 1 Date");
    fireEvent.change(dateInput, { target: { value: "2026-01-06" } });

    await waitFor(() => {
      const updated = screen.getByTitle("PPTX Preview");
      const nextSrc = updated.getAttribute("src");
      expect(nextSrc).toBeTruthy();
      expect(nextSrc).not.toEqual(firstSrc);
    });
  });

  test("strict invariant: last slide remains byte-for-byte unchanged", async () => {
    const template = await makeMinimalPptxArrayBuffer();
    const { updatedPptxBytes } = await updatePptxDateOnly(template, "2026-01-06");
    await expect(assertLastSlideUnchanged(template, updatedPptxBytes)).resolves.toEqual(true);
  });
});
