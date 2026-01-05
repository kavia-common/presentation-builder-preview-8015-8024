import React, { useEffect, useMemo, useRef, useState } from "react";
import "./App.css";
import {
  createPptxObjectUrl,
  downloadPptxBytes,
  fetchBundledTemplatePptx,
  todayIsoDate,
  updatePptxDateOnly,
} from "./pptx/templateEditor";

// PUBLIC_INTERFACE
function App() {
  const [theme, setTheme] = useState("light");

  const [templateBytes, setTemplateBytes] = useState(null); // ArrayBuffer

  const [dateISO, setDateISO] = useState(todayIsoDate());

  const [generatedBytes, setGeneratedBytes] = useState(null); // Uint8Array
  const [previewUrl, setPreviewUrl] = useState("");
  const [status, setStatus] = useState({ kind: "idle", message: "" });
  const [detectionInfo, setDetectionInfo] = useState(null);

  // Apply theme to document
  useEffect(() => {
    document.documentElement.setAttribute("data-theme", theme);
  }, [theme]);

  // Load bundled template (Default Template mode)
  useEffect(() => {
    let cancelled = false;

    async function loadBundled() {
      try {
        setStatus({ kind: "loading", message: "Loading Default Template…" });
        const bytes = await fetchBundledTemplatePptx();
        if (cancelled) return;
        setTemplateBytes(bytes);
        setStatus({ kind: "ready", message: "Default Template loaded." });
      } catch (e) {
        if (cancelled) return;
        setStatus({
          kind: "error",
          message:
            e instanceof Error ? e.message : "Failed to load Default Template.",
        });
      }
    }

    loadBundled();

    return () => {
      cancelled = true;
    };
  }, []);

  // Regenerate PPTX any time date or template changes.
  // Debounced to keep typing/rapid changes smooth, while still updating "immediately"
  // after the user input settles (lightweight in-memory mutation of slide1 only).
  const regenSeqRef = useRef(0);
  useEffect(() => {
    if (!templateBytes) return () => {};

    const seq = (regenSeqRef.current += 1);
    let cancelled = false;

    const DEBOUNCE_MS = 150;
    const timeoutId = window.setTimeout(() => {
      async function regenerate() {
        try {
          setStatus({ kind: "loading", message: "Generating preview (date only)…" });

          const { updatedPptxBytes, detected } = await updatePptxDateOnly(
            templateBytes,
            dateISO
          );

          // Ignore out-of-date results (date changed again while we were generating).
          if (cancelled || regenSeqRef.current !== seq) return;

          setGeneratedBytes(updatedPptxBytes);
          setDetectionInfo(detected);

          // Refresh preview URL: changing the URL forces the iframe to load the new PPTX blob.
          setPreviewUrl((prev) => {
            if (prev) URL.revokeObjectURL(prev);
            return createPptxObjectUrl(updatedPptxBytes);
          });

          setStatus({ kind: "ready", message: "Preview updated." });
        } catch (e) {
          if (cancelled || regenSeqRef.current !== seq) return;

          setGeneratedBytes(null);
          setDetectionInfo(null);
          setPreviewUrl((prev) => {
            if (prev) URL.revokeObjectURL(prev);
            return "";
          });

          setStatus({
            kind: "error",
            message:
              e instanceof Error
                ? e.message
                : "Failed to generate PPTX preview.",
          });
        }
      }

      regenerate();
    }, DEBOUNCE_MS);

    return () => {
      cancelled = true;
      window.clearTimeout(timeoutId);
    };
  }, [templateBytes, dateISO]);

  useEffect(() => {
    return () => {
      if (previewUrl) URL.revokeObjectURL(previewUrl);
    };
  }, [previewUrl]);

  // PUBLIC_INTERFACE
  const toggleTheme = () => {
    setTheme((prevTheme) => (prevTheme === "light" ? "dark" : "light"));
  };

  const canDownload = useMemo(() => !!generatedBytes, [generatedBytes]);

  return (
    <div className="App ocean">
      <header className="ocean-header">
        <div className="ocean-topbar">
          <div className="brand">
            <div className="brand-mark" aria-hidden="true" />
            <div className="brand-text">
              <div className="brand-title">Default Template (Date Only)</div>
              <div className="brand-subtitle">
                Only Slide 1 date is editable • all other content locked • last slide preserved exactly
              </div>
            </div>
          </div>

          <button
            className="btn btn-secondary"
            onClick={toggleTheme}
            aria-label={`Switch to ${theme === "light" ? "dark" : "light"} mode`}
            type="button"
          >
            {theme === "light" ? "Dark mode" : "Light mode"}
          </button>
        </div>

        <main className="ocean-main">
          <section className="card controls" aria-label="Template controls">
            <div className="card-header">
              <h2>Default Template</h2>
              <p>
                The app ships with a built-in PPTX template. All content is locked
                except the Slide 1 date field, which is updated in-place with the
                exact original formatting and placement.
              </p>
            </div>

            <div className="field-row">
              <div className="field">
                <label htmlFor="dateInput">Slide 1 Date</label>
                <input
                  id="dateInput"
                  className="date-input"
                  type="date"
                  value={dateISO}
                  onChange={(e) => setDateISO(e.target.value)}
                />
                <div className="hint">
                  Updates only the exact date text runs on slide 1 (preserving font,
                  size, color, spacing, and locale format). All other slides/files,
                  including the last slide, are left untouched.
                </div>
              </div>
            </div>

            <div className="status-row" role="status" aria-live="polite">
              <span
                className={`status-pill ${
                  status.kind === "error"
                    ? "error"
                    : status.kind === "loading"
                      ? "loading"
                      : "ok"
                }`}
              >
                {status.kind.toUpperCase()}
              </span>
              <span className="status-text">{status.message}</span>
            </div>

            <div className="actions">
              <button
                className="btn btn-primary"
                type="button"
                disabled={!canDownload}
                onClick={() => {
                  if (!generatedBytes) return;
                  const name = `Updated_Template_${dateISO}.pptx`;
                  downloadPptxBytes(generatedBytes, name);
                }}
              >
                Download PPTX
              </button>

              <div className="meta">
                {detectionInfo ? (
                  <div className="hint">
                    Placeholder detection:{" "}
                    <strong>{detectionInfo.mode}</strong>
                    {detectionInfo.mode === "token" ? (
                      <> (token: <code>{detectionInfo.token}</code>)</>
                    ) : null}
                  </div>
                ) : (
                  <div className="hint">
                    Tip: For best reliability across templates, place a{" "}
                    <code>{"{{DATE}}"}</code> token in slide 1.
                  </div>
                )}
              </div>
            </div>
          </section>

          <section className="card preview" aria-label="Preview">
            <div className="card-header">
              <h2>Preview</h2>
              <p>
                Browser preview support for PPTX varies. This embeds the generated
                PPTX as a blob URL. If your browser cannot render it, use Download.
              </p>
            </div>

            {previewUrl ? (
              <div className="preview-frame-wrap">
                {/* Many browsers won't natively render PPTX; still provides a consistent "preview area".
                   If unsupported, user will see a download prompt or blank frame. */}
                <iframe
                  title="PPTX Preview"
                  className="preview-frame"
                  src={previewUrl}
                />
              </div>
            ) : (
              <div className="empty-preview">
                Preview will appear here once the template is loaded and the PPTX
                is generated.
              </div>
            )}
          </section>
        </main>

        <footer className="ocean-footer">
          <div className="footer-note">
            Locking behavior: The app edits only the Slide 1 date text runs inside{" "}
            <code>ppt/slides/slide1.xml</code>. All other files in the PPTX zip are
            left unchanged, so the last slide remains byte-for-byte identical.
          </div>
        </footer>
      </header>
    </div>
  );
}

export default App;
