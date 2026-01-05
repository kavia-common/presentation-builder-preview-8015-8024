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
  test("changing the date regenerates the blob URL and keeps a visible Open/Download link", async () => {
    global.fetch = jest.fn(async () => ({
      ok: true,
      arrayBuffer: async () => await makeMinimalPptxArrayBuffer(),
    }));

    render(<App />);

    // The UI always provides a non-blank preview surface.
    // Once the PPTX is generated, it must show the Download/Open link (reliable fallback).
    const firstLink = await screen.findByRole("link", {
      name: "Download / Open PPTX",
    });
    const firstHref = firstLink.getAttribute("href");
    expect(firstHref).toBeTruthy();

    const dateInput = screen.getByLabelText("Slide 1 Date");
    fireEvent.change(dateInput, { target: { value: "2026-01-06" } });

    await waitFor(() => {
      const nextLink = screen.getByRole("link", { name: "Download / Open PPTX" });
      const nextHref = nextLink.getAttribute("href");
      expect(nextHref).toBeTruthy();
      expect(nextHref).not.toEqual(firstHref);
    });
  });

  test("strict invariant: last slide remains byte-for-byte unchanged", async () => {
    const template = await makeMinimalPptxArrayBuffer();
    const { updatedPptxBytes } = await updatePptxDateOnly(template, "2026-01-06");
    await expect(assertLastSlideUnchanged(template, updatedPptxBytes)).resolves.toEqual(true);
  });
});
