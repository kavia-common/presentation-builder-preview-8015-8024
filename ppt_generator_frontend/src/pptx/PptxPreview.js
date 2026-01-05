import React, { useEffect, useMemo, useState } from "react";

/**
 * PUBLIC_INTERFACE
 * PptxPreview provides a robust client-side preview surface for a PPTX blob URL.
 *
 * Rationale:
 * - Most browsers cannot natively render PPTX content in an iframe/object, resulting
 *   in a blank area and a "no preview slides shown" user experience.
 * - We still try to embed for the browsers/environments that support it.
 * - We ALWAYS provide a user-visible fallback: a direct link that opens/downloads the PPTX.
 *
 * This component does not parse or modify the PPTX. It only displays the already-generated blob URL.
 */
export default function PptxPreview({ url, filename, debug = false }) {
  const [hasLoadEvent, setHasLoadEvent] = useState(false);
  const [embedError, setEmbedError] = useState("");

  useEffect(() => {
    setHasLoadEvent(false);
    setEmbedError("");
  }, [url]);

  const label = useMemo(() => filename || "Generated.pptx", [filename]);

  // Always provide a user-visible surface (never a blank card).
  // If url is empty, show an explanatory placeholder. If url exists, ALWAYS show the link.
  if (!url) {
    return (
      <div className="pptx-preview">
        <div className="preview-toolbar" aria-label="Preview actions">
          <div className="hint">
            Preview will appear here once the template is loaded and the PPTX is generated.
          </div>
        </div>

        <div className="empty-preview">
          Waiting for generation output…
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
          Most browsers can’t render PPTX inline. If slides are not visible below, use “Download /
          Open PPTX”.
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
            Inline PPTX preview is not supported in this browser. Use the link above to open the
            generated file.
          </div>
        </object>
      </div>
    </div>
  );
}
