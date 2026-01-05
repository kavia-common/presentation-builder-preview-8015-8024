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
