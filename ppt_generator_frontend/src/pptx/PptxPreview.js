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
  onSelectAddMode = null,
  addMode = "dateOnly",
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

  // Render the currently selected slide.
  useEffect(() => {
    let cancelled = false;

    async function doRender() {
      setRenderError("");
      setRendered(null);

      if (!pptxBytes || !pptxBytes.length) return;
      if (!slideIndexes.length) return;

      const slideIndex = slideIndexes[Math.min(currentIdx, slideIndexes.length - 1)];
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
  const currentSlideNumber = slideIndexes.length
    ? slideIndexes[Math.min(currentIdx, slideIndexes.length - 1)]
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

          <div className="add-mode">
            <div className="add-mode-label">Add remaining slides</div>
            <label className="radio" htmlFor="addModeDateOnly">
              <input
                id="addModeDateOnly"
                type="radio"
                name="addMode"
                value="dateOnly"
                checked={addMode === "dateOnly"}
                onChange={() => onSelectAddMode && onSelectAddMode("dateOnly")}
              />
              Date-only (locked)
            </label>
            <label className="radio" htmlFor="addModeSkillFactory">
              <input
                id="addModeSkillFactory"
                type="radio"
                name="addMode"
                value="skillFactory"
                checked={addMode === "skillFactory"}
                onChange={() =>
                  onSelectAddMode && onSelectAddMode("skillFactory")
                }
              />
              Skill Factory
            </label>

            {skillFactoryUI?.enabled ? (
              <div className="skill-factory-controls" aria-label="Add Skill Factory">
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
                  onClick={() => skillFactoryUI.onAddFactory && skillFactoryUI.onAddFactory()}
                >
                  Add Factory (4 slides)
                </button>

                {skillFactoryUI.addedFactories?.length ? (
                  <div className="hint" style={{ marginTop: 8 }}>
                    Added:{" "}
                    <strong>
                      {skillFactoryUI.addedFactories.map((f) => f.label).join(", ")}
                    </strong>
                  </div>
                ) : (
                  <div className="hint" style={{ marginTop: 8 }}>
                    Adds 4 slides per factory after the existing slides, before the last slide.
                    (Only the first factory slide is fully implemented in this iteration.)
                  </div>
                )}
              </div>
            ) : null}
          </div>
        </div>

        <div className="hint">
          Slide-by-slide preview (landscape). Slide 1 matches the template visuals; only
          the date text is edited in the PPTX. The last slide remains byte-identical.
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

          <div className="carousel-indicator" aria-label="Slide indicator">
            Slide <strong>{slideIndexes.length ? currentSlideNumber : "—"}</strong>{" "}
            / <strong>{slideIndexes.length || "—"}</strong>
            {currentSlideNumber === 1 ? (
              <span className="carousel-badge editable">date editable</span>
            ) : currentSlideNumber === slideIndexes[slideIndexes.length - 1] ? (
              <span className="carousel-badge locked">locked (last)</span>
            ) : (
              <span className="carousel-badge locked">locked</span>
            )}
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
