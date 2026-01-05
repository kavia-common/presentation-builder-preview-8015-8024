import "@testing-library/jest-dom";

// JSDOM may not implement blob URL helpers; mock them for preview tests.
if (!global.URL.createObjectURL) {
  global.URL.createObjectURL = jest.fn(
    () => "blob:mock-url-" + Math.random().toString(16).slice(2)
  );
}
if (!global.URL.revokeObjectURL) {
  global.URL.revokeObjectURL = jest.fn();
}

/**
 * Provide a minimal fetch() polyfill for tests so that code paths which rely on
 * fetching the bundled template work under Jest/JSDOM.
 *
 * Notes:
 * - Individual tests may still override global.fetch (e.g., to inject a minimal zip).
 * - This polyfill only handles GET /assets/template.pptx and otherwise throws.
 */
/**
 * Ensure fetch() exists in Jest/JSDOM for /assets/template.pptx.
 *
 * Some Jest environments provide a partial fetch; we still wrap it to guarantee
 * `arrayBuffer()` works for the bundled template fetch used by the app + renderer tests.
 */
{
  // eslint-disable-next-line no-undef
  const fs = require("fs");
  // eslint-disable-next-line no-undef
  const path = require("path");

  const prevFetch = global.fetch;

  global.fetch = jest.fn(async (input, init) => {
    const url = String(input || "");
    if (url === "/assets/template.pptx") {
      const abs = path.resolve(__dirname, "../public/assets/template.pptx");
      const buf = fs.readFileSync(abs);
      return {
        ok: true,
        status: 200,
        arrayBuffer: async () =>
          buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
      };
    }
    if (prevFetch) return await prevFetch(input, init);
    throw new Error(`setupTests fetch polyfill: unhandled url: ${url}`);
  });
}
