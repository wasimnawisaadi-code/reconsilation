import { defineConfig } from "vitest/config";

// Standalone test config — intentionally does NOT load the app's vite.config
// (TanStack Start / Nitro plugins), so unit tests run fast in a plain Node
// environment against the pure reconciliation logic.
export default defineConfig({
  test: {
    environment: "node",
    include: ["src/**/*.test.ts"],
  },
});
