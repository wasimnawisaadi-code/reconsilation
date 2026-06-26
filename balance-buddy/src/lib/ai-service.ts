// Vertex AI access for schema discovery and residual semantic matching.
// All calls run server-side only (the service-account key never reaches the browser).

const PRIMARY_MODEL = "gemini-2.5-flash";
const PRO_MODEL = "gemini-2.5-pro";
const FALLBACK_MODEL = "gemini-2.5-flash-lite";

const getVertexAI = async () => {
  if (typeof window !== "undefined") {
    throw new Error("Vertex AI can only be initialized on the server.");
  }

  const { VertexAI } = await import("@google-cloud/vertexai");
  const { default: fs } = await import("node:fs");
  const { default: path } = await import("node:path");

  const keyPath = path.resolve(process.cwd(), "service-account.json");
  const keyFile = JSON.parse(fs.readFileSync(keyPath, "utf8"));

  process.env.GOOGLE_APPLICATION_CREDENTIALS = keyPath;

  return new VertexAI({
    project: keyFile.project_id,
    location: "us-central1",
  });
};

export type ColumnMapping = {
  date?: string;
  passport?: string;
  paxName?: string;
  description?: string;
  reference?: string;
  charge?: string;
  credit?: string;
};

export type SchemaResult = {
  ours: ColumnMapping;
  partner: ColumnMapping;
  /** Regex/relationship hints the engine can apply (e.g. ticket-number shape). */
  patterns: string[];
  /** 0–1 confidence the mapping is correct. */
  confidence: number;
  /** Short human explanation of the detected structure & join logic. */
  logic: string;
};

/** Low-temperature, JSON-only generation config for deterministic structured output. */
const jsonConfig = {
  temperature: 0,
  topP: 0.1,
  maxOutputTokens: 4096,
  responseMimeType: "application/json",
} as const;

/** Pull text out of a Vertex response and parse the first JSON object found. */
function parseJsonResponse<T>(text: string, fallback: T): T {
  if (!text) return fallback;
  // Strip markdown fences if the model added them despite JSON mode.
  const cleaned = text
    .replace(/```json/gi, "")
    .replace(/```/g, "")
    .trim();
  try {
    return JSON.parse(cleaned) as T;
  } catch {
    const m = cleaned.match(/\{[\s\S]*\}/);
    if (m) {
      try {
        return JSON.parse(m[0]) as T;
      } catch {
        /* fall through */
      }
    }
    return fallback;
  }
}

async function generate(prompt: string, model: string, fallbackModel: string): Promise<string> {
  const vertex = await getVertexAI();
  const run = async (name: string) => {
    const m = vertex.getGenerativeModel({ model: name, generationConfig: jsonConfig as any });
    const r = await m.generateContent(prompt);
    return r.response.candidates?.[0]?.content?.parts?.[0]?.text ?? "";
  };
  try {
    return await run(model);
  } catch (e) {
    console.warn(`[AI] Model ${model} failed, falling back to ${fallbackModel}:`, e);
    return await run(fallbackModel);
  }
}

/**
 * Schema Discovery: analyse the headers + sample rows of BOTH ledgers and map
 * each side's columns to the canonical reconciliation roles. Travel/visa aware.
 */
export async function discoverSchema(ours: any[], partner: any[]): Promise<SchemaResult> {
  const sampleOurs = ours.slice(0, 12);
  const samplePartner = partner.slice(0, 12);

  const prompt = `
You are a financial data engineer building a reconciliation platform for a TRAVEL & VISA AGENCY.
The ledgers you process are: internal visa-agency ledger vs supplier/partner statements.

You are given the first rows (row 0 is usually the header) of TWO ledgers recording the SAME
transactions from opposite sides.

TRAVEL/VISA-SPECIFIC CONTEXT — critical for column mapping:
- "DOCNO" or similar columns contain codes like "VS26/1996" (visa charge), "VR26/82" (reversal/refund),
  "PY26/766" (payment), "IS25/xxx" (airline/interline). VS = visa sale, VR = void/reversal, PY = payment.
- "TICKET NO." or similar may contain "3VS XXXXX" where XXXXX is a passport number with a check digit.
  The passport identifier is everything AFTER "3VS " with the LAST character stripped.
- "SECTOR / DESCRIPTION" is the visa service type: "30 DAYS", "60 DAYS", "60 DAYS MULTI",
  "1M EXTENSION", "SECURITY DEPOSIT" etc. For VR rows it may contain the reversal reason.
- "PAX NAME" is the passenger/client name.
- Security deposits (tagged in description/service type) must NEVER be matched with visa charges.
- Partner ledgers may have passport in a "Type" column (2nd Type column in Format B), Comments
  column for passenger name, Description for service type.

Map each side's columns to these canonical roles (use EXACT header text, or null if absent):
- date: transaction date
- passport: primary identity linking the two sides — passport number, national ID, or account no.
- paxName: passenger / client / counterparty name
- description: free-text narration / service type / particulars
- reference: document reference — voucher, DOCNO, booking ID, invoice, PNR, ticket number
- charge: amount billed / debit (money the agency owes or was charged)
- credit: amount received / payment / refund / top-up

Notes:
- Signed single-amount column: map BOTH charge and credit to it.
- Separate DR/CR columns: figure out which is charge vs credit from context and values.

Internal ledger rows:
${JSON.stringify(sampleOurs)}

Partner ledger rows:
${JSON.stringify(samplePartner)}

Return ONLY this JSON shape:
{
  "ours": { "date": "...", "passport": "...", "paxName": "...", "description": "...", "reference": "...", "charge": "...", "credit": "..." },
  "partner": { ... same keys ... },
  "patterns": ["e.g. passport embedded as '3VS P0411511C' → strip '3VS ' and last char", "security deposits tagged in description"],
  "confidence": 0.0,
  "logic": "1-2 sentence explanation of the structure and how the two sides join"
}
`.trim();

  const text = await generate(prompt, PRIMARY_MODEL, FALLBACK_MODEL);
  const parsed = parseJsonResponse<Partial<SchemaResult>>(text, {});
  return {
    ours: parsed.ours ?? {},
    partner: parsed.partner ?? {},
    patterns: Array.isArray(parsed.patterns) ? parsed.patterns.map(String) : [],
    confidence: typeof parsed.confidence === "number" ? parsed.confidence : 0.75,
    logic: parsed.logic ?? "Heuristic column mapping applied.",
  };
}

export type AiMatch = {
  oursIndex: number;
  partnerIndex: number;
  confidence: number;
  reason: string;
};

/**
 * Residual semantic matching: only the rows the rule engine could NOT match are
 * sent here. The AI proposes pairings with reasons; the rule layer re-validates
 * every suggestion before it is accepted, so accuracy stays high.
 */
export async function matchRowsWithAi(
  unmatchedOurs: any[],
  unmatchedPartner: any[],
): Promise<{ pairs: AiMatch[]; insights: string }> {
  if (unmatchedOurs.length === 0 || unmatchedPartner.length === 0) {
    return { pairs: [], insights: "Nothing left to match." };
  }

  const a = unmatchedOurs.slice(0, 60).map((r, i) => ({ i, ...r }));
  const b = unmatchedPartner.slice(0, 60).map((r, i) => ({ i, ...r }));

  const prompt = `
You are a meticulous reconciliation auditor for a travel/visa agency. Two lists of UNMATCHED ledger
rows remain after deterministic rule-based matching. Your job is to pair rows from Ledger A (ours)
with rows from Ledger B (partner) that represent the SAME real transaction.

CRITICAL RULES — never violate these:
1. NEVER match a "security_deposit" scenario row with a "visa_charge" scenario row — they are
   completely different financial instruments even if amounts coincide.
2. NEVER match a wrong-invoice/wrong-client refund (isReversal=true) with a normal visa charge.
3. A row may match AT MOST one row on the other side. Do not force uncertain matches — omit them.
4. "i" is the index within each list below; return those indices.
5. confidence is your 0–1 certainty.

Matching clues to use:
- passport / ID numbers (allow check-digit/format differences; "3VS P0411511C" → P0411511)
- party/passenger names (allow typos, honorifics, word-order differences)
- amounts (allow ±1 rounding; for multi-pax groups allow N× differences)
- dates within ±7 days (visa processing often has a lag)
- document references (booking IDs, invoice numbers, PNRs)
- visaType should broadly agree (30 DAYS vs 30 DAYS = strong; 30 DAYS vs 60 DAYS = uncertain)
- scenario field — prefer matching same-scenario rows

Ledger A (ours): ${JSON.stringify(a)}
Ledger B (partner): ${JSON.stringify(b)}

Return ONLY:
{
  "pairs": [ { "oursIndex": 0, "partnerIndex": 0, "confidence": 0.0, "reason": "why" } ],
  "insights": "brief overall observation"
}
`.trim();

  try {
    const text = await generate(prompt, PRO_MODEL, PRIMARY_MODEL);
    const parsed = parseJsonResponse<{ pairs: AiMatch[]; insights: string }>(text, {
      pairs: [],
      insights: "",
    });
    return {
      pairs: Array.isArray(parsed.pairs) ? parsed.pairs : [],
      insights: parsed.insights ?? "",
    };
  } catch (e) {
    console.error("[AI Matcher] Error:", e);
    return { pairs: [], insights: "AI matching engine unavailable." };
  }
}
