import React from "react";
import PptxPreview from "./pptx/PptxPreview";

// PUBLIC_INTERFACE
/**
 * PreviewPage displays the preview carousel for the current slides.
 * Navigates back to the main editor via the supplied `onBack` handler.
 * Props:
 *   - slides: Array of slide objects, in deck order.
 *   - selectedSidebarIdx: number, the current selected slide index in the preview.
 *   - setSelectedSidebarIdx: function, updates the slide index.
 *   - onBack: function, called to return to the main editor page.
 */
function PreviewPage({ slides, selectedSidebarIdx, setSelectedSidebarIdx, onBack }) {
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
          slides={slides}
          selectedIdx={selectedSidebarIdx}
          setSlideIdx={setSelectedSidebarIdx}
        />
      </div>
    </div>
  );
}

export default PreviewPage;
