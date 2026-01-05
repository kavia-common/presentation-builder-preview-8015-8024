import React, { useEffect, useMemo, useState } from "react";

/**
 * PUBLIC_INTERFACE
 * PptxPreview provides a robust client-side preview surface for a PPTX blob URL.
 *
 * Rationale:
 * - Most browsers cannot natively render PPTX content in an iframe/object.
 * - We still try to embed for the browsers/environments that support it.
 * - We ALWAYS provide a user-visible fallback: a direct link that opens/downloads the PPTX.
 * - If any step fails, we show a clear error message instead of a blank state.
 *
 * This component does not parse or modify the PPTX. It only displays the already-generated blob URL.
 */
export default function PptxPreview({
  url,
  filename,
  debug = false,
  errorMessage = "",
  pipeline = null,
}) {
  const [hasLoadEvent, setHasLoadEvent] = useState(false);
  const [embedError, setEmbedError] = useState("");

  useEffect(() => {
    setHasLoadEvent(false);
    setEmbedError("");
  }, [url]);

  const label = useMemo(() => filename || "Generated.pptx", [filename]);

  // Never render a blank surface:
  // - When URL exists: ALWAYS show the Open/Download link.
  // - When URL missing: show either an error (if provided) or a waiting message.
  const showError = Boolean(!url && errorMessage);

  if (!url) {
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
              <strong>Error:</strong> {errorMessage}
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
        <a className="btn btn-secondary" href={url} download={label}>
          Download / Open PPTX
        </a>

        <div className="hint">
          Most browsers can’t render PPTX inline. If slides are not visible below,
          use “Download / Open PPTX”.
          {debug ? (
            <>
              {" "}
              <span>
                (iframe load: <strong>{hasLoadEvent ? "fired" : "pending"}</strong>
                {embedError ? (
                  <>
                    {", error: "}
                    <strong>{embedError}</strong>
                  </>
                ) : null}
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

      <div className="preview-frame-wrap">
        {/* Attempt iframe embed first. Some environments may at least show an Office prompt UI. */}
        <iframe
          title="PPTX Preview"
          className="preview-frame"
          src={url}
          onLoad={() => setHasLoadEvent(true)}
          onError={() => setEmbedError("iframe error")}
        />

        {/* Secondary attempt via <object> (sometimes behaves differently than iframe). */}
        <object
          className="preview-object"
          data={url}
          type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
          aria-label="PPTX Object Preview"
          onError={() => setEmbedError("object error")}
        >
          <div className="empty-preview">
            Inline PPTX preview is not supported in this browser. Use the link
            above to open the generated file.
          </div>
        </object>
      </div>
    </div>
  );
}
