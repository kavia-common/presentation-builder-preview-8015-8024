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

  useEffect(() => {
    setHasLoadEvent(false);
  }, [url]);

  const label = useMemo(() => filename || "Generated.pptx", [filename]);

  if (!url) {
    return (
      <div className="empty-preview">
        Preview will appear here once the template is loaded and the PPTX is generated.
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
          If slides are not visible below, your browser likely can’t render PPTX inline.
          Use “Download / Open PPTX”.
          {debug ? (
            <>
              {" "}
              <span>
                (embed load event: <strong>{hasLoadEvent ? "fired" : "pending"}</strong>)
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
        />

        {/* Secondary attempt via <object> (sometimes behaves differently than iframe). */}
        <object
          className="preview-object"
          data={url}
          type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
          aria-label="PPTX Object Preview"
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
