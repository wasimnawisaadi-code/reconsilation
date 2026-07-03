import { describe, it, expect } from "vitest";
import { requireUser, assertAiRateLimit } from "./server-auth";

describe("requireUser", () => {
  // .env.local is loaded by vitest, so Supabase counts as configured here.
  it("rejects calls with no access token when auth is configured", async () => {
    await expect(requireUser(undefined)).rejects.toThrow(/sign in required/i);
  });
});

describe("assertAiRateLimit", () => {
  it("allows calls under the limit and blocks the one over it", () => {
    const key = `test-${Date.now()}`;
    for (let i = 0; i < 5; i++) expect(() => assertAiRateLimit(key, 5)).not.toThrow();
    expect(() => assertAiRateLimit(key, 5)).toThrow(/limit reached/i);
  });

  it("tracks users independently", () => {
    const a = `a-${Date.now()}`;
    const b = `b-${Date.now()}`;
    for (let i = 0; i < 3; i++) assertAiRateLimit(a, 3);
    expect(() => assertAiRateLimit(b, 3)).not.toThrow();
  });
});
