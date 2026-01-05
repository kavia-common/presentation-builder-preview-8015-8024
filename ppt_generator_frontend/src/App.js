import React, { useEffect, useMemo, useRef, useState } from "react";
import "./App.css";
import PptxPreview from "./pptx/PptxPreview";
import {
  createPptxObjectUrl,
  downloadPptxBytes,
  fetchBundledTemplatePptx,
  todayIsoDate,
  updatePptxDateOnly,
} from "./pptx/templateEditor";
import {
  addSkillFactoryToPptx,
  getSkillFactoryScaffold,
} from "./pptx/skillFactory";

/**
 * Returns true when UI debug logging is enabled.
 * Gated by REACT_APP_LOG_LEVEL=debug to avoid noisy logs in normal usage.
 */
function isDebugEnabled() {
  return String(process.env.REACT_APP_LOG_LEVEL || "").toLowerCase() === "debug";
}

/**
 * Debug logger helper (no-op unless debug is enabled).
 */
function debugLog(...args) {
  if (!isDebugEnabled()) return;
  // eslint-disable-next-line no-console
  console.log("[pptx-preview]", ...args);
}

// PUBLIC_INTERFACE
function App() {
  const [theme, setTheme] = useState("light");

  const [templateBytes, setTemplateBytes] = useState(null); // ArrayBuffer

  const [dateISO, setDateISO] = useState(todayIsoDate());

  const [generatedBytes, setGeneratedBytes] = useState(null); // Uint8Array
  const [previewUrl, setPreviewUrl] = useState("");
  const previousPreviewUrlRef = useRef("");
  const [previewKey, setPreviewKey] = useState(0);
  const [status, setStatus] = useState({ kind: "idle", message: "" });
  const [detectionInfo, setDetectionInfo] = useState(null);

  // Controls future behavior for adding remaining slides.
  // For this subtask:
  // - dateOnly: only slide 1 date editing is applied.
  // - skillFactory: user may append a Skill Factory (4 slides) before the last slide.
  const [addMode, setAddMode] = useState("dateOnly"); // dateOnly | skillFactory

  // Skill Factory state
  const [skillFactories, setSkillFactories] = useState([]); // [{ kind, label }]
  const [selectedFactoryKind, setSelectedFactoryKind] = useState("java"); // java | dataEngineering

  // Pipeline step visibility (to avoid blank preview and aid debugging)
  const [pipeline, setPipeline] = useState({
    step: "init", // init | fetch:loading | fetch:ok | fetch:error | edit:loading | edit:ok | edit:error | blob:ok
    detail: "",
  });

  // Apply theme to document
  useEffect(() => {
    document.documentElement.setAttribute("data-theme", theme);
  }, [theme]);

  // Revoke the previous blob URL only after the new URL is committed.
  // This prevents a class of "blank preview" issues where a URL gets revoked too early
  // while the iframe is still loading or before React commits the new src.
  useEffect(() => {
    const prev = previousPreviewUrlRef.current;
    if (prev && prev !== previewUrl) {
      window.setTimeout(() => {
        try {
          URL.revokeObjectURL(prev);
          debugLog("revokeObjectURL(prev) ok");
        } catch (e) {
          debugLog("revokeObjectURL(prev) failed", e);
        }
      }, 0);
    }
    previousPreviewUrlRef.current = previewUrl;
  }, [previewUrl]);

  // Load bundled template (Default Template mode)
  useEffect(() => {
    let cancelled = false;

    async function loadBundled() {
      try {
        setPipeline({ step: "fetch:loading", detail: "/assets/template.pptx" });
        setStatus({ kind: "loading", message: "Loading Default Template…" });
        debugLog("fetch start", "/assets/template.pptx");

        const bytes = await fetchBundledTemplatePptx();
        if (cancelled) return;

        debugLog("fetch ok", { byteLength: bytes?.byteLength });
        setPipeline({
          step: "fetch:ok",
          detail: `${bytes?.byteLength ?? 0} bytes`,
        });

        setTemplateBytes(bytes);
        setStatus({ kind: "ready", message: "Default Template loaded." });
      } catch (e) {
        if (cancelled) return;
        debugLog("fetch error", e);
        setPipeline({
          step: "fetch:error",
          detail: e instanceof Error ? e.message : String(e),
        });
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
          setPipeline({ step: "edit:loading", detail: `date=${dateISO}` });
          setStatus({
            kind: "loading",
            message: "Generating preview (date only)…",
          });
          debugLog("edit start", { dateISO, templateBytes: templateBytes.byteLength });

          const { updatedPptxBytes: dateOnlyBytes, detected } =
            await updatePptxDateOnly(templateBytes, dateISO);

          let finalBytes = dateOnlyBytes;

          // Optional: append Skill Factory slides (4 per factory), inserted before last slide.
          // IMPORTANT: This must NOT modify the template's last slide bytes.
          if (addMode === "skillFactory" && skillFactories.length) {
            setPipeline({
              step: "edit:loading",
              detail: `date=${dateISO}, skillFactories=${skillFactories.length}`,
            });

            // Apply factories in sequence to ensure ordering is stable.
            // eslint-disable-next-line no-restricted-syntax
            for (const f of skillFactories) {
              // eslint-disable-next-line no-await-in-loop
              const res = await addSkillFactoryToPptx(finalBytes, f);
              finalBytes = res.updatedPptxBytes;
            }
          }

          // Ignore out-of-date results (date changed again while we were generating).
          if (cancelled || regenSeqRef.current !== seq) return;

          debugLog("edit ok", {
            updatedBytes: finalBytes?.byteLength ?? finalBytes?.length,
            detected,
            addMode,
            skillFactories,
          });
          setPipeline({
            step: "edit:ok",
            detail: `${finalBytes?.byteLength ?? finalBytes?.length ?? 0} bytes`,
          });

          setGeneratedBytes(finalBytes);
          setDetectionInfo(detected);

          const nextUrl = createPptxObjectUrl(finalBytes);
          debugLog("blob url created", nextUrl);

          setPipeline({ step: "blob:ok", detail: nextUrl });

          // Force reload of embed surfaces with a new URL + key.
          setPreviewUrl(nextUrl);
          setPreviewKey((k) => k + 1);

          setStatus({ kind: "ready", message: "Preview updated." });
        } catch (e) {
          if (cancelled || regenSeqRef.current !== seq) return;

          debugLog("edit error", e);
          setPipeline({
            step: "edit:error",
            detail: e instanceof Error ? e.message : String(e),
          });

          setGeneratedBytes(null);
          setDetectionInfo(null);
          setPreviewUrl("");

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
  }, [templateBytes, dateISO, addMode, skillFactories]);

  useEffect(() => {
    return () => {
      // Cleanup current URL on unmount.
      if (previewUrl) {
        try {
          URL.revokeObjectURL(previewUrl);
        } catch (e) {
          // ignore
        }
      }
      // Also cleanup any previous url we may still hold.
      const prev = previousPreviewUrlRef.current;
      if (prev && prev !== previewUrl) {
        try {
          URL.revokeObjectURL(prev);
        } catch (e) {
          // ignore
        }
      }
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
                Only Slide 1 date is editable • all other content locked • last
                slide preserved exactly
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
                  Updates only the exact date text runs on slide 1 (preserving
                  font, size, color, spacing, and locale format). All other
                  slides/files, including the last slide, are left untouched.
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

            {isDebugEnabled() ? (
              <div className="field-row">
                <div className="field">
                  <div className="hint">
                    Debug pipeline: <strong>{pipeline.step}</strong>
                    {pipeline.detail ? (
                      <>
                        {" "}
                        (<code>{pipeline.detail}</code>)
                      </>
                    ) : null}
                  </div>
                </div>
              </div>
            ) : null}

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
                    Placeholder detection: <strong>{detectionInfo.mode}</strong>
                    {detectionInfo.mode === "token" ? (
                      <>
                        {" "}
                        (token: <code>{detectionInfo.token}</code>)
                      </>
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
                Most browsers don’t natively render PPTX inline. We attempt an
                embedded preview, and always provide a “Download / Open” fallback
                so you never get a blank preview.
              </p>
            </div>

            <PptxPreview
              key={previewKey}
              url={previewUrl}
              filename={
                generatedBytes
                  ? `Updated_Template_${dateISO}.pptx`
                  : "Updated_Template.pptx"
              }
              pptxBytes={generatedBytes}
              addMode={addMode}
              onSelectAddMode={setAddMode}
              debug={isDebugEnabled()}
              errorMessage={
                status.kind === "error" && status.message ? status.message : ""
              }
              pipeline={
                isDebugEnabled()
                  ? { step: pipeline.step, detail: pipeline.detail }
                  : null
              }
              skillFactoryUI={{
                enabled: addMode === "skillFactory",
                selectedKind: selectedFactoryKind,
                onChangeKind: setSelectedFactoryKind,
                onAddFactory: () => {
                  const scaffold = getSkillFactoryScaffold(selectedFactoryKind);
                  // Add only the factory descriptor; slide insertion happens in regeneration.
                  setSkillFactories((prev) => {
                    // Avoid duplicates of same kind for now.
                    if (prev.some((x) => x.kind === scaffold.kind)) return prev;
                    return [...prev, { kind: scaffold.kind, label: scaffold.label }];
                  });
                },
                addedFactories: skillFactories,
              }}
            />
          </section>
        </main>

        <footer className="ocean-footer">
          <div className="footer-note">
            Locking behavior: The app edits only the Slide 1 date text runs inside{" "}
            <code>ppt/slides/slide1.xml</code>. All other files in the PPTX zip
            are left unchanged, so the last slide remains byte-for-byte identical.
          </div>
        </footer>
      </header>
    </div>
  );
}

export default App;
