import { createServerFn } from "@tanstack/react-start";
import { discoverSchema, matchRowsWithAi, generateExecutiveBrief } from "./ai-service";
import { requireUser, assertAiRateLimit } from "./server-auth";

/** Auth + rate-limit gate shared by every AI endpoint. No-op when Supabase
 *  is not configured (open dev mode); otherwise a valid signed-in session is
 *  required and each user gets a per-hour AI budget. */
async function guard(accessToken: string | undefined): Promise<void> {
  const user = await requireUser(accessToken);
  if (user) assertAiRateLimit(user.id);
}

export const analyzeSchema = createServerFn({ method: "POST" })
  .inputValidator((d: any) => d as { ours: any[]; partner: any[]; accessToken?: string })
  .handler(async ({ data }) => {
    try {
      await guard(data.accessToken);
      const result = await discoverSchema(data.ours, data.partner);
      return { data: result };
    } catch (e) {
      console.error("[Server Action] analyzeSchema error:", e);
      throw e;
    }
  });

export const performAiMatching = createServerFn({ method: "POST" })
  .inputValidator(
    (d: any) => d as { unmatchedOurs: any[]; unmatchedPartner: any[]; accessToken?: string },
  )
  .handler(async ({ data }) => {
    try {
      await guard(data.accessToken);
      const result = await matchRowsWithAi(data.unmatchedOurs, data.unmatchedPartner);
      return { data: result };
    } catch (e) {
      console.error("[Server Action] performAiMatching error:", e);
      throw e;
    }
  });

export const aiExecutiveBrief = createServerFn({ method: "POST" })
  .inputValidator((d: any) => d as { payload: unknown; accessToken?: string })
  .handler(async ({ data }) => {
    try {
      await guard(data.accessToken);
      const result = await generateExecutiveBrief(data.payload);
      return { data: result };
    } catch (e) {
      console.error("[Server Action] aiExecutiveBrief error:", e);
      throw e;
    }
  });
