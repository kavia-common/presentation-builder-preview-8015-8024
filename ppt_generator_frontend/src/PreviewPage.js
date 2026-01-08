import React, { useEffect, useRef, useState } from "react";
import PptxPreview from "./pptx/PptxPreview";
import {
  fetchBundledTemplatePptx,
  updatePptxDateOnly,
  prunePptxToFirstAndLastSlides,
  createPptxObjectUrl,
} from "./pptx/templateEditor";

/**
 * PUBLIC_INTERFACE
 * PreviewPage displays the preview carousel for the current slides, generating the preview PPTX in-browser.
 * Handles template fetching, date-only edit, deck pruning, and passes pptxBytes to PptxPreview.
 * Shows clear errors if any step fails.
 *
 * Props:
 *   - slides: Array of slide objects, in deck order.
 *      ONLY Slide 1's date text is editable (reflected in preview).
 *   - selectedSidebarIdx: number, the current selected slide index in the preview.
 *   - setSelectedSidebarIdx: function, updates the slide index.
 *   - onBack: function, called to return to the main editor page.
 */
function PreviewPage({ slides, selectedSidebarIdx, setSelectedSidebarIdx, onBack }) {
  // Preview pipeline state
  const [pptxBytes, setPptxBytes] = useState(null);
  const [blobUrl, setBlobUrl] = useState("");
  const [error, setError] = useState("");
  const [loadingStep, setLoadingStep] = useState("");
  const prevBlobUrlRef = useRef("");

  // Infer date from Slide 1
  const previewDate = (slides && slides[0] && slides[0].date)
    ? slides[0].date
    : undefined;

  useEffect(() => {
    let cancelled = false;
    setPptxBytes(null);
    setBlobUrl("");
    setError("");
    setLoadingStep("fetch-template");

    async function generatePreviewPptx() {
      try {
        // 1) Fetch the template PPTX
        setLoadingStep("fetch-template");
        const templateBuf = await fetchBundledTemplatePptx();

        // 2) Apply date-only edit to Slide 1 (preserving all else)
        setLoadingStep("edit-date");
        const { updatedPptxBytes } = await updatePptxDateOnly(templateBuf, previewDate);

        // 3) Prune deck to first+last slides (deck order)
        setLoadingStep("prune-to-first-last");
        const { updatedPptxBytes: prunedBytes } = await prunePptxToFirstAndLastSlides(updatedPptxBytes);

        if (cancelled) return;

        setPptxBytes(prunedBytes);

        // 4) Generate Blob URL for download/open
        setLoadingStep("generate-blob");
        const url = createPptxObjectUrl(prunedBytes);
        setBlobUrl(url);

        // Clean up previous blob URLs to release memory
        if (prevBlobUrlRef.current && prevBlobUrlRef.current !== url) {
          try {
            URL.revokeObjectURL(prevBlobUrlRef.current);
          } catch (e) {
            /* ignore */
          }
        }
        prevBlobUrlRef.current = url;

        setLoadingStep("");
      } catch (err) {
        if (cancelled) return;
        setError(err && err.message ? err.message : String(err));
        setLoadingStep("");
        setPptxBytes(null);
        setBlobUrl("");
      }
    }

    generatePreviewPptx();
    return () => {
      cancelled = true;
    };
    // Trigger whenever Slide 1's date changes only (other slide state is irrelevant to preview)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [previewDate]);

  // Show error messages in the preview UI as described
  const errorMessage =
    error ||
    (loadingStep
      ? `Generating preview: ${loadingStep.replace(/-/g, " ")}...`
      : "");

  return (
    <div className="main-content">
      <div style={{ margin: "32px 0 0 0" }}>
        <button
          className="btn btn-secondary"
          type="button"
          onClick={onBack}
          style={{
            marginBottom: "30px",
            padding: "7px 22px",
            borderRadius: "8px",
            background: "#e5e7eb",
            color: "#1a1a1a",
            fontWeight: 500,
          }}
        >
          ⬅ Back to Editor
        </button>
      </div>
      <div id="preview-section" className="preview-block">
        <PptxPreview
          pptxBytes={pptxBytes}
          url={blobUrl}
          filename={"Preview.pptx"}
          selectedIdx={selectedSidebarIdx}
          setSlideIdx={setSelectedSidebarIdx}
          errorMessage={errorMessage}
        />
        {error && (
          <div
            style={{
              color: "#EF4444",
              background: "#FEF2F2",
              marginTop: "24px",
              padding: "10px 16px",
              borderRadius: "8px",
              fontWeight: 500,
            }}
            role="alert"
            aria-live="assertive"
          >
            <strong>Preview could not be generated:</strong>{" "}
            {error}
          </div>
        )}
      </div>
    </div>
  );
}

export default PreviewPage;
