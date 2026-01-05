import React, { useEffect, useMemo, useRef, useState } from "react";
import "./App.css";
import PptxPreview from "./pptx/PptxPreview";
import {
  createPptxObjectUrl,
  downloadPptxBytes,
  fetchBundledTemplatePptx,
  prunePptxToFirstAndLastSlides,
  todayIsoDate,
  updatePptxDateOnly,
} from "./pptx/templateEditor";
import {
  addSkillFactoryToPptx,
  getSkillFactoryScaffold,
} from "./pptx/skillFactory";
import {
  addTeamColumn,
  addTeamRow,
  createDefaultSkillFactoryForm,
  removeTeamColumn,
  removeTeamRow,
  updateTeamCell,
  updateTeamColumnHeader,
} from "./pptx/skillFactoryModel";

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

  // Skill Factory state
  const [skillFactories, setSkillFactories] = useState([]); // [{ kind, label, slide1 }]
  const [selectedFactoryKind, setSelectedFactoryKind] = useState("java"); // java | dataEngineering
  const [skillFactoryForm, setSkillFactoryForm] = useState(
    createDefaultSkillFactoryForm()
  );

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

          // Default-deck behavior: keep only Slide 1 and the template's last slide.
          // This ensures preview/download start as a 2-slide deck, while preserving:
          // - Slide 1 is date-only editable (already applied above)
          // - Last slide remains unchanged (we never touch last slide XML bytes)
          const { updatedPptxBytes: prunedBytes } =
            await prunePptxToFirstAndLastSlides(dateOnlyBytes);

          let finalBytes = prunedBytes;

          // Optional: insert Skill Factory slides (4 per factory), inserted before the last slide.
          // IMPORTANT: This must NOT modify the template's last slide bytes.
          // Note: skillFactories is empty by default; factories are only added when user clicks "Add".
          if (skillFactories.length) {
            setPipeline({
              step: "edit:loading",
              detail: `date=${dateISO}, pruned=1, skillFactories=${skillFactories.length}`,
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
  }, [templateBytes, dateISO, skillFactories]);

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
                Default deck: Slide 1 + final slide • only Slide 1 date is editable • final
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

            <div className="field-row">
              <div className="field">
                <div className="skill-factory-panel" aria-label="Skill Factory slide 1 fields">
                  <div className="skill-factory-panel-head">
                    <div className="skill-factory-panel-title">Skill Factory — Slide 1</div>
                    <div className="skill-factory-panel-subtitle">
                      These fields populate the Skill Factory slide(s) inserted only after you click{" "}
                      <strong>Add Skill Factory</strong>.
                    </div>
                  </div>

                  <div className="sf-grid">
                    <label className="sf-field">
                      <span className="sf-label">Skill Factory name (after “Digital Applications:”)</span>
                      <input
                        className="sf-input"
                        type="text"
                        value={skillFactoryForm.factoryName}
                        onChange={(e) =>
                          setSkillFactoryForm((p) => ({ ...p, factoryName: e.target.value }))
                        }
                        placeholder="e.g., Java"
                      />
                    </label>

                    <div className="sf-row-3">
                      <label className="sf-field">
                        <span className="sf-label">Sprint number</span>
                        <input
                          className="sf-input"
                          type="text"
                          value={skillFactoryForm.sprintNumber}
                          onChange={(e) =>
                            setSkillFactoryForm((p) => ({ ...p, sprintNumber: e.target.value }))
                          }
                          placeholder="e.g., 14"
                        />
                      </label>

                      <label className="sf-field">
                        <span className="sf-label">Sprint start date</span>
                        <input
                          className="sf-input"
                          type="date"
                          value={skillFactoryForm.sprintStartISO}
                          onChange={(e) =>
                            setSkillFactoryForm((p) => ({ ...p, sprintStartISO: e.target.value }))
                          }
                        />
                      </label>

                      <label className="sf-field">
                        <span className="sf-label">Sprint end date</span>
                        <input
                          className="sf-input"
                          type="date"
                          value={skillFactoryForm.sprintEndISO}
                          onChange={(e) =>
                            setSkillFactoryForm((p) => ({ ...p, sprintEndISO: e.target.value }))
                          }
                        />
                      </label>
                    </div>

                    <div className="sf-two-col">
                      <label className="sf-field">
                        <span className="sf-label">Project Highlights (one per line)</span>
                        <textarea
                          className="sf-textarea"
                          rows={5}
                          value={skillFactoryForm.projectHighlightsText}
                          onChange={(e) =>
                            setSkillFactoryForm((p) => ({
                              ...p,
                              projectHighlightsText: e.target.value,
                            }))
                          }
                          placeholder={"• Item 1\n• Item 2"}
                        />
                      </label>

                      <label className="sf-field">
                        <span className="sf-label">Project Lowlights (one per line)</span>
                        <textarea
                          className="sf-textarea"
                          rows={5}
                          value={skillFactoryForm.projectLowlightsText}
                          onChange={(e) =>
                            setSkillFactoryForm((p) => ({
                              ...p,
                              projectLowlightsText: e.target.value,
                            }))
                          }
                          placeholder={"• Item 1\n• Item 2"}
                        />
                      </label>
                    </div>

                    <div className="sf-section">
                      <div className="sf-section-title">Team Members</div>

                      <div className="sf-table-actions">
                        <button
                          className="btn btn-secondary btn-sm"
                          type="button"
                          onClick={() =>
                            setSkillFactoryForm((p) => ({
                              ...p,
                              teamTable: addTeamRow(p.teamTable),
                            }))
                          }
                        >
                          + Row
                        </button>
                        <button
                          className="btn btn-secondary btn-sm"
                          type="button"
                          onClick={() =>
                            setSkillFactoryForm((p) => ({
                              ...p,
                              teamTable: addTeamColumn(p.teamTable, "New Column"),
                            }))
                          }
                        >
                          + Column
                        </button>
                      </div>

                      <div className="sf-table-wrap" role="region" aria-label="Team members table">
                        <table className="sf-table">
                          <thead>
                            <tr>
                              {skillFactoryForm.teamTable.columns.map((c, colIdx) => (
                                <th key={`h-${colIdx}`}>
                                  <div className="sf-th">
                                    <input
                                      className="sf-input sf-th-input"
                                      type="text"
                                      value={c}
                                      onChange={(e) =>
                                        setSkillFactoryForm((p) => ({
                                          ...p,
                                          teamTable: updateTeamColumnHeader(
                                            p.teamTable,
                                            colIdx,
                                            e.target.value
                                          ),
                                        }))
                                      }
                                    />
                                    <button
                                      className="sf-icon-btn"
                                      type="button"
                                      aria-label={`Remove column ${colIdx + 1}`}
                                      onClick={() =>
                                        setSkillFactoryForm((p) => ({
                                          ...p,
                                          teamTable: removeTeamColumn(p.teamTable, colIdx),
                                        }))
                                      }
                                    >
                                      ×
                                    </button>
                                  </div>
                                </th>
                              ))}
                              <th className="sf-actions-col" aria-hidden="true" />
                            </tr>
                          </thead>
                          <tbody>
                            {skillFactoryForm.teamTable.rows.map((row, rowIdx) => (
                              <tr key={`r-${rowIdx}`}>
                                {skillFactoryForm.teamTable.columns.map((_, colIdx) => (
                                  <td key={`c-${rowIdx}-${colIdx}`}>
                                    <input
                                      className="sf-input sf-cell-input"
                                      type="text"
                                      value={row[colIdx] ?? ""}
                                      onChange={(e) =>
                                        setSkillFactoryForm((p) => ({
                                          ...p,
                                          teamTable: updateTeamCell(
                                            p.teamTable,
                                            rowIdx,
                                            colIdx,
                                            e.target.value
                                          ),
                                        }))
                                      }
                                    />
                                  </td>
                                ))}
                                <td className="sf-actions-col">
                                  <button
                                    className="sf-icon-btn"
                                    type="button"
                                    aria-label={`Remove row ${rowIdx + 1}`}
                                    onClick={() =>
                                      setSkillFactoryForm((p) => ({
                                        ...p,
                                        teamTable: removeTeamRow(p.teamTable, rowIdx),
                                      }))
                                    }
                                  >
                                    ×
                                  </button>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>

                      <label className="sf-field" style={{ marginTop: 10 }}>
                        <span className="sf-label">SME name (shown below the table)</span>
                        <input
                          className="sf-input"
                          type="text"
                          value={skillFactoryForm.smeName}
                          onChange={(e) =>
                            setSkillFactoryForm((p) => ({ ...p, smeName: e.target.value }))
                          }
                          placeholder="e.g., John Smith"
                        />
                      </label>
                    </div>

                    <div className="sf-two-col">
                      <label className="sf-field">
                        <span className="sf-label">Key Activities completed (previous week)</span>
                        <textarea
                          className="sf-textarea"
                          rows={5}
                          value={skillFactoryForm.keyActivitiesCompletedText}
                          onChange={(e) =>
                            setSkillFactoryForm((p) => ({
                              ...p,
                              keyActivitiesCompletedText: e.target.value,
                            }))
                          }
                          placeholder={"• Item 1\n• Item 2"}
                        />
                      </label>

                      <label className="sf-field">
                        <span className="sf-label">Key Activities planned (next week)</span>
                        <textarea
                          className="sf-textarea"
                          rows={5}
                          value={skillFactoryForm.keyActivitiesPlannedText}
                          onChange={(e) =>
                            setSkillFactoryForm((p) => ({
                              ...p,
                              keyActivitiesPlannedText: e.target.value,
                            }))
                          }
                          placeholder={"• Item 1\n• Item 2"}
                        />
                      </label>
                    </div>

                    <div className="hint" style={{ marginTop: 6 }}>
                      Slide header will format as:{" "}
                      <code>Sprint N (DD MMM YYYY – DD MMM YYYY)</code>.
                    </div>
                  </div>
                </div>
              </div>
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
                selectedKind: selectedFactoryKind,
                onChangeKind: setSelectedFactoryKind,
                onAddFactory: () => {
                  const scaffold = getSkillFactoryScaffold(selectedFactoryKind);
                  // Add only the factory descriptor; slide insertion happens in regeneration.
                  // IMPORTANT: Default deck remains 2 slides until user clicks Add.
                  setSkillFactories((prev) => {
                    // Avoid duplicates of same kind for now.
                    if (prev.some((x) => x.kind === scaffold.kind)) return prev;
                    return [
                      ...prev,
                      { kind: scaffold.kind, label: scaffold.label, slide1: skillFactoryForm },
                    ];
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
