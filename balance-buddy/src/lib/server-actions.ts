import { createServerFn } from "@tanstack/react-start";
import { discoverSchema, matchRowsWithAi } from "./ai-service";

export const analyzeSchema = createServerFn({ method: "POST" })
  .inputValidator((d: any) => d as { ours: any[]; partner: any[] })
  .handler(async ({ data }) => {
    try {
      const result = await discoverSchema(data.ours, data.partner);
      return { data: result };
    } catch (e) {
      console.error("[Server Action] analyzeSchema error:", e);
      throw e;
    }
  });

export const performAiMatching = createServerFn({ method: "POST" })
  .inputValidator((d: any) => d as { unmatchedOurs: any[]; unmatchedPartner: any[] })
  .handler(async ({ data }) => {
    try {
      const result = await matchRowsWithAi(data.unmatchedOurs, data.unmatchedPartner);
      return { data: result };
    } catch (e) {
      console.error("[Server Action] performAiMatching error:", e);
      throw e;
    }
  });
