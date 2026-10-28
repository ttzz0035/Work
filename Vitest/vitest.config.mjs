import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    reporters: ["minimal"],
    onStackTrace: () => false,
    printConsoleTrace: false,
  },
});