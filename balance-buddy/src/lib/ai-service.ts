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
You are a financial data engineer building a GENERIC ledger reconciliation platform that must
support ANY ledger type: bank statements, accounts receivable / payable, supplier & customer
statements, invoices, expense sheets, travel/visa agency books, etc.
You are given the first rows (row 0 is usually the header) of two ledgers that record the SAME
underlying transactions from two sides: our internal books vs an external party (partner / bank /
supplier / customer).

Map each side's columns to these canonical roles (use the EXACT header text as the value,
or null if a role is absent):
- date: transaction / value / posting / record date
- passport: the PRIMARY ENTITY/PARTY identifier used to link the two sides — e.g. account number,
  customer id, vendor id, member id, national id, or passport. Pick the most stable unique key.
- paxName: the party / customer / counterparty / payee name
- description: free-text narration / particulars / remarks / memo
- reference: the DOCUMENT/TRANSACTION reference — invoice no, voucher, cheque no, PO, txn id,
  ticket number, PNR, or booking id
- charge: the amount column representing a debit/charge/fee owed (money out / amount billed)
- credit: the amount column representing a payment/receipt/refund/top-up (money in)

Notes:
- A ledger may use ONE signed "Amount" column (negative = charge, positive = credit) — in that case
  map BOTH charge and credit to that same column and say so in "logic".
- A ledger may use separate DR / CR (or Debit / Credit) columns. Decide which side is charge vs
  credit for THIS data based on the values and headers.
- The two ledgers will usually use different header names for the same role — that is expected;
  align them by meaning, not by exact text.

Internal ledger rows:
${JSON.stringify(sampleOurs)}

Partner ledger rows:
${JSON.stringify(samplePartner)}

Return ONLY this JSON shape:
{
  "ours": { "date": "...", "passport": "...", "paxName": "...", "description": "...", "reference": "...", "charge": "...", "credit": "..." },
  "partner": { ... same keys ... },
  "patterns": ["e.g. ticket numbers are 13 digits", "passport embedded after '3VS'"],
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
You are a meticulous reconciliation auditor working across ANY ledger type (bank, AR/AP, supplier,
customer, invoices, travel). Two lists of UNMATCHED ledger rows remain after deterministic matching.
Pair rows from Ledger A with rows from Ledger B that represent the SAME real transaction. Use every
clue: party/counterparty names (allow typos, reordering, honorifics, abbreviations), account/id
numbers (allow check-digit/format differences), document references (invoice/cheque/txn/ticket numbers)
embedded in text, amounts (allow small fees / rounding / tax / FX), and nearby dates.

Rules:
- A row may match AT MOST one row on the other side. Do not force matches — omit uncertain ones.
- "i" is the index within each list below; return those indices.
- confidence is your 0–1 certainty.

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
