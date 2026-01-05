import React from "react";
import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import App from "./App";

function mockArrayBuffer(bytes = [1, 2, 3]) {
  const arr = new Uint8Array(bytes);
  return arr.buffer;
}

describe("PPTX preview regeneration", () => {
  test("changing the date regenerates the blob URL and refreshes the preview iframe src", async () => {
    // Mock fetch for /assets/template.pptx
    global.fetch = jest.fn(async () => ({
      ok: true,
      arrayBuffer: async () => mockArrayBuffer([7, 7, 7, 7]),
    }));

    render(<App />);

    // Wait for initial preview iframe to appear (template loaded + first generation).
    const iframe = await screen.findByTitle("PPTX Preview");

    const firstSrc = iframe.getAttribute("src");
    expect(firstSrc).toBeTruthy();

    const dateInput = screen.getByLabelText("Slide 1 Date");
    // Change date -> should trigger regeneration and assign a new blob URL.
    fireEvent.change(dateInput, { target: { value: "2026-01-06" } });

    await waitFor(() => {
      const updated = screen.getByTitle("PPTX Preview");
      const nextSrc = updated.getAttribute("src");
      expect(nextSrc).toBeTruthy();
      expect(nextSrc).not.toEqual(firstSrc);
    });
  });
});
