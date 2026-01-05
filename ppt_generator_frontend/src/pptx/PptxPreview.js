import React, { useEffect, useMemo, useState } from "react";
import {
  getSlideSizeEmu,
  listSlideIndexes,
  renderSlideToSvgDataUrl,
} from "./pptxSlideRenderer";

/**
 * PUBLIC_INTERFACE
 * PptxPreview renders a slide-by-slide (carousel) preview for a PPTX.
 *
 * Goals:
 * - Always show a non-blank preview surface even when browsers cannot render PPTX inline.
 * - Render each slide deterministically as an SVG (landscape) based on slide embedded images.
 * - Provide Next/Prev navigation across all slides.
 * - Keep invariants: the preview is read-only; PPTX mutations are done elsewhere (date-only on slide 1).
 */
export default function PptxPreview({
  url,
  filename,
  pptxBytes = null,
  debug = false,
  errorMessage = "",
  pipeline = null,
  skillFactoryUI = null,
}) {
  const [slideIndexes, setSlideIndexes] = useState([]);
  const [currentIdx, setCurrentIdx] = useState(0); // index into slideIndexes[]
  const [slideSize, setSlideSize] = useState(null);
  const [rendered, setRendered] = useState(null); // {dataUrl,widthPx,heightPx}
  const [renderError, setRenderError] = useState("");

  const label = useMemo(() => filename || "Generated.pptx", [filename]);

  // Keep a resilient link even if the parent URL prop is briefly empty.
  // The preview renderer is read-only; this does not affect PPTX invariants.
  const [localUrl, setLocalUrl] = useState("");

  useEffect(() => {
    if (!pptxBytes || !pptxBytes.length) {
      setLocalUrl("");
      return () => {};
    }

    const next = URL.createObjectURL(
      new Blob([pptxBytes], {
        type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      })
    );
    setLocalUrl(next);

    return () => {
      try {
        URL.revokeObjectURL(next);
      } catch (e) {
        // ignore
      }
    };
  }, [pptxBytes]);

  // When new PPTX bytes arrive, enumerate slides and reset selection to slide 1.
  useEffect(() => {
    let cancelled = false;

    async function loadMeta() {
      setRenderError("");
      setRendered(null);
      setSlideIndexes([]);
      setSlideSize(null);
      setCurrentIdx(0);

      if (!pptxBytes || !pptxBytes.length) return;

      try {
        const [idxs, sz] = await Promise.all([
          listSlideIndexes(pptxBytes),
          getSlideSizeEmu(pptxBytes),
        ]);
        if (cancelled) return;
        setSlideIndexes(idxs);
        setSlideSize(sz);
        setCurrentIdx(0);
      } catch (e) {
        if (cancelled) return;
        setRenderError(e instanceof Error ? e.message : String(e));
      }
    }

    loadMeta();

    return () => {
      cancelled = true;
    };
  }, [pptxBytes]);

  // Clamp the current carousel index when slide count changes.
  // This keeps the UI in sync after Skill Factory insertions/pruning and prevents
  // out-of-range indexes from showing incorrect counters or rendering attempts.
  useEffect(() => {
    if (!slideIndexes.length) {
      if (currentIdx !== 0) setCurrentIdx(0);
      return;
    }

    const maxIdx = Math.max(0, slideIndexes.length - 1);
    if (currentIdx > maxIdx) setCurrentIdx(maxIdx);
  }, [slideIndexes, currentIdx]);

  // Render the currently selected slide.
  useEffect(() => {
    let cancelled = false;

    async function doRender() {
      setRenderError("");
      setRendered(null);

      if (!pptxBytes || !pptxBytes.length) return;
      if (!slideIndexes.length) return;

      const slideIndex =
        slideIndexes[Math.min(currentIdx, slideIndexes.length - 1)];
      try {
        const result = await renderSlideToSvgDataUrl(pptxBytes, slideIndex, {
          widthPx: 1040,
        });
        if (cancelled) return;
        setRendered(result);
      } catch (e) {
        if (cancelled) return;
        setRenderError(e instanceof Error ? e.message : String(e));
      }
    }

    doRender();

    return () => {
      cancelled = true;
    };
  }, [pptxBytes, slideIndexes, currentIdx]);

  const effectiveUrl = url || localUrl;
  const showError = Boolean((!effectiveUrl && errorMessage) || renderError);
  const canNavigate = slideIndexes.length > 1;

  // Carousel position (1-based) and total count for the requested X/Y counter.
  const totalSlides = slideIndexes.length;
  const currentSlidePos = totalSlides ? Math.min(currentIdx, totalSlides - 1) + 1 : 0;

  // Actual PPT slide number (e.g., slide part index in the PPTX) for labels/badges.
  const currentSlideNumber = totalSlides
    ? slideIndexes[Math.min(currentIdx, totalSlides - 1)]
    : 0;

  if (!pptxBytes || !pptxBytes.length) {
    return (
      <div className="pptx-preview">
        <div className="preview-toolbar" aria-label="Preview actions">
          <div className="hint">
            {showError
              ? "Preview could not be generated."
              : "Preview will appear here once the template is loaded and the PPTX is generated."}
          </div>
          {debug && pipeline ? (
            <div className="hint">
              Debug: <strong>{pipeline.step}</strong>
              {pipeline.detail ? (
                <>
                  {" "}
                  (<code>{pipeline.detail}</code>)
                </>
              ) : null}
            </div>
          ) : null}
        </div>

        <div className="empty-preview" role={showError ? "alert" : undefined}>
          {showError ? (
            <>
              <strong>Error:</strong> {errorMessage || renderError}
            </>
          ) : (
            "Waiting for generation output…"
          )}
        </div>
      </div>
    );
  }

  return (
    <div className="pptx-preview">
      <div className="preview-toolbar" aria-label="Preview actions">
        <div className="preview-actions-row">
          {/* IMPORTANT: Keep link always visible when PPTX bytes exist (prop url OR local fallback). */}
          {effectiveUrl ? (
            <a className="btn btn-secondary" href={effectiveUrl} download={label}>
              Download / Open PPTX
            </a>
          ) : null}

          {/* Skill Factory control: single, obvious action in the preview panel. */}
          {skillFactoryUI ? (
            <div className="add-mode" aria-label="Add Skill Factory">
              <div className="add-mode-label">Add Skill Factory</div>

              <label className="field" style={{ minWidth: "220px" }}>
                <span className="hint" style={{ display: "block", marginBottom: 6 }}>
                  Factory
                </span>
                <select
                  className="skill-factory-select"
                  value={skillFactoryUI.selectedKind || "java"}
                  onChange={(e) =>
                    skillFactoryUI.onChangeKind &&
                    skillFactoryUI.onChangeKind(e.target.value)
                  }
                >
                  <option value="java">Java</option>
                  <option value="dataEngineering">Data Engineering</option>
                </select>
              </label>

              <button
                className="btn btn-primary"
                type="button"
                onClick={() =>
                  skillFactoryUI.onAddFactory && skillFactoryUI.onAddFactory()
                }
              >
                Add Skill Factory (4 slides)
              </button>

              {skillFactoryUI.addedFactories?.length ? (
                <div className="hint">
                  Added:{" "}
                  <strong>
                    {skillFactoryUI.addedFactories.map((f) => f.label).join(", ")}
                  </strong>
                </div>
              ) : (
                <div className="hint">
                  Default deck is 2 slides (Slide 1 + final). Clicking Add inserts 4 slides per
                  factory before the final slide. Slide 1 remains date-only editable; the final
                  slide stays unchanged.
                </div>
              )}
            </div>
          ) : null}
        </div>

        <div className="hint">
          Slide-by-slide preview (landscape). Default output contains only Slide 1 and the
          final slide; only the Slide 1 date is edited in the PPTX. Factory slides are inserted
          only after you click “Add Skill Factory”, and they appear before the final slide.
          The final slide remains byte-identical.
          {debug && slideSize ? (
            <>
              {" "}
              <span>
                (slides: <strong>{slideIndexes.length}</strong>, sizeEMU:{" "}
                <code>
                  {slideSize.cx}×{slideSize.cy}
                </code>
                {pipeline ? (
                  <>
                    {", pipeline: "}
                    <strong>{pipeline.step}</strong>
                  </>
                ) : null}
                )
              </span>
            </>
          ) : null}
        </div>
      </div>

      <div className="preview-frame-wrap preview-carousel">
        <div className="carousel-toolbar" aria-label="Slide navigation">
          <button
            className="btn btn-secondary"
            type="button"
            disabled={!canNavigate || currentIdx <= 0}
            onClick={() => setCurrentIdx((v) => Math.max(0, v - 1))}
          >
            Prev
          </button>

          <div className="carousel-counter" aria-label="Slide counter">
            <strong>{totalSlides ? currentSlidePos : "—"}</strong>
            <span className="carousel-counter-sep">/</span>
            <strong>{totalSlides || "—"}</strong>
          </div>

          <button
            className="btn btn-secondary"
            type="button"
            disabled={!canNavigate || currentIdx >= slideIndexes.length - 1}
            onClick={() =>
              setCurrentIdx((v) => Math.min(slideIndexes.length - 1, v + 1))
            }
          >
            Next
          </button>

          <div className="carousel-indicator" aria-label="Slide indicator">
            Slide <strong>{totalSlides ? currentSlideNumber : "—"}</strong> /{" "}
            <strong>{totalSlides || "—"}</strong>
            {currentSlideNumber === 1 ? (
              <span className="carousel-badge editable">date editable</span>
            ) : currentSlideNumber === slideIndexes[slideIndexes.length - 1] ? (
              <span className="carousel-badge locked">locked (last)</span>
            ) : (
              <span className="carousel-badge locked">locked</span>
            )}
          </div>
        </div>

        <div className="carousel-stage" aria-label="Slide preview stage">
          <div className="slide-frame" role="img" aria-label={`Slide ${currentSlideNumber}`}>
            {renderError ? (
              <div className="empty-preview" role="alert">
                <strong>Render error:</strong> {renderError}
              </div>
            ) : !rendered ? (
              <div className="empty-preview">Rendering slide…</div>
            ) : (
              <img
                className="slide-image"
                alt={`Slide ${currentSlideNumber}`}
                src={rendered.dataUrl}
              />
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
