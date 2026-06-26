import * as XLSX from "xlsx";
import * as XLSXStyle from "xlsx-js-style";
import Papa from "papaparse";

export type Side = "ours" | "partner";

/**
 * Scenario tags for each row — used for color-coding, filtering, and
 * preventing cross-category matches (e.g. security deposit ↔ visa charge).
 */
export type Scenario =
  | "visa_charge"      // Normal VS visa charge
  | "security_deposit" // Security / guarantee deposit (different from visa fee)
  | "wrong_invoice"    // VR reversal: wrong invoice was billed
  | "wrong_client"     // VR reversal: wrong client was billed
  | "duplicate"        // VR reversal: duplicate entry
  | "refund"           // Other reversal / refund
  | "bank_transfer"    // Inter-party bank transfer / settlement / top-up
  | "multi_passenger"  // Group row covering N passengers (bundled charge)
  | "flight"           // Airline / interline ticket (IS rows)
  | "addon";           // Supplier add-on service (meal, bus, hotel…)

export type LedgerRow = {
  side: Side;
  index: number; // original order
  date: string;
  passport: string | null;
  paxName: string;
  description: string;
  reference: string;
  /** Positive money outflow from our perspective (visa charge). 0 if not a charge. */
  charge: number;
  /** Positive money inflow (top-up / refund / bank transfer). 0 if not. */
  credit: number;
  /** Row category. */
  kind: "charge" | "credit" | "other";
  /**
   * True when the row is an inter-party money movement (bank transfer, wire, TT,
   * settlement, deposit, top-up …) rather than a per-item charge/invoice. Used to
   * drive the dedicated "Payments & Settlements" reconciliation view. Works for
   * ANY ledger type — detection is keyword-driven, not format-specific.
   */
  settlement?: boolean;
  /**
   * True for supplier rows that are add-on services (meal, bus, hotel, transfer…)
   * rather than visa/travel charges. These will never match our ledger and should
   * be counted separately so they don't dilute the visa reconciliation rate.
   */
  addon?: boolean;
  /** Scenario tag for this row (used for UI color-coding and match gating). */
  scenario?: Scenario;
  /**
   * Visa service type extracted from the description (e.g. "30 DAYS", "60 DAYS",
   * "1M EXTENSION", "SECURITY DEPOSIT", "60 DAYS MULTI").
   * Used to prevent a 30-day visa from matching a security deposit.
   */
  visaType?: string;
  /** True for VR (reversal) rows — charge correction, NOT a bank transfer. */
  isReversal?: boolean;
  /**
   * Duplicate detection WITHIN this same ledger. When ≥2 rows share the same
   * passport + amount + visa type, they are flagged as a duplicate group.
   * `duplicateCount` is the group size; `duplicateIndex` is this row's 1-based
   * position in the group. Security deposits vs visa charges are NOT duplicates
   * (different visa types) so they're never grouped together.
   */
  duplicateCount?: number;
  duplicateIndex?: number;
  /** Original 0-based row index in the uploaded sheet (header = row 0). */
  srcRow?: number;
  raw: Record<string, unknown>;
};

const num = (v: unknown): number => {
  if (v === null || v === undefined || v === "") return 0;
  const s = String(v).replace(/,/g, "").trim();
  if (!s) return 0;
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
};

/**
 * Generic detector for inter-party settlement / payment rows across ANY ledger.
 * Matches the vocabulary banks and accounting systems use for money movements:
 * transfers, wires, TT, remittances, deposits, top-ups, settlements, NEFT/RTGS/IMPS,
 * SWIFT, and explicit "payment received/made" narrations.
 */
const SETTLEMENT_RE =
  /\b(bank\s*transfer|wire(?:\s*transfer)?|telegraphic|t\/?t|remittance|remit|settlement|settle|top[\s-]*up|deposit|funds?\s*transfer|account\s*transfer|neft|rtgs|imps|swift|ach|giro|payment\s*(?:received|made|in|out)?|received\s*with\s*thanks|cash\s*(?:deposit|payment|received)|opening\s*balance\s*payment|on\s*account)\b/i;

/** Bank-name / account narration markers (incl. common UAE banks). */
const BANK_ACCOUNT_RE =
  /\bA\/?C[-\s]?\d|commercial bank|islamic bank|national bank|\badcb\b|\benbd\b|\bfab\b|mashreq|emirates\s+islamic|rak\s*bank|\badib\b|bank\s+of|\bllc\b/i;

export function isSettlementText(...parts: (string | null | undefined)[]): boolean {
  const text = parts.filter(Boolean).join(" ");
  if (!text) return false;
  return SETTLEMENT_RE.test(text) || BANK_ACCOUNT_RE.test(text);
}

/** Matches "SECURITY DEPOSIT" or "SECURITY" standalone (not in the middle of a word). */
const SECURITY_DEPOSIT_RE = /\bsecurity\s*deposit\b|\bsecurity\b(?!\s*(?:code|number|check|scan))/i;

/** Matches common visa service duration types in description fields. */
const VISA_TYPE_RE = /\b(\d{1,3})\s*days?\b|\b(1M|2M|3M|6M)\s*ext(?:ension)?\b|\bmulti\b|\bsingle\b|\bsecurity\s*deposit\b/i;

/** Detect if description text indicates a security deposit row. */
function isSecurityDepositText(...parts: (string | null | undefined)[]): boolean {
  return SECURITY_DEPOSIT_RE.test(parts.filter(Boolean).join(" "));
}

/**
 * Extract a normalised visa-service type label from description text.
 * Returns e.g. "30 DAYS", "60 DAYS MULTI", "1M EXTENSION", "SECURITY DEPOSIT", or undefined.
 */
function extractVisaType(text: string): string | undefined {
  if (!text) return undefined;
  if (SECURITY_DEPOSIT_RE.test(text)) return "SECURITY DEPOSIT";
  const m = text.match(VISA_TYPE_RE);
  if (!m) return undefined;
  return m[0].toUpperCase().replace(/\s+/g, " ").trim();
}

/**
 * Determine the scenario for a row based on docno prefix, description, flags, etc.
 * Used in both our-ledger and partner-ledger parsers.
 */
function detectScenario(opts: {
  isReversal?: boolean;
  isSettle?: boolean;
  isJV?: boolean;
  isFlightRow?: boolean;
  isSecDep?: boolean;
  isGroupRow?: boolean;
  isAddon?: boolean;
  desc?: string;
}): Scenario {
  if (opts.isAddon) return "addon";
  if (opts.isSettle || opts.isJV) return "bank_transfer";
  if (opts.isReversal) {
    const d = (opts.desc ?? "").toLowerCase();
    if (/wrong\s*invoice/.test(d)) return "wrong_invoice";
    if (/wrong\s*client/.test(d)) return "wrong_client";
    if (/duplicate/.test(d)) return "duplicate";
    return "refund";
  }
  if (opts.isFlightRow) return "flight";
  if (opts.isSecDep) return "security_deposit";
  if (opts.isGroupRow) return "multi_passenger";
  return "visa_charge";
}

export type ColumnMapping = {
  date?: string;
  passport?: string;
  paxName?: string;
  description?: string;
  reference?: string;
  charge?: string;
  credit?: string;
};

const normPassport = (p: string | null | undefined): string | null => {
  if (!p) return null;
  const s = String(p).toUpperCase().replace(/[^A-Z0-9]/g, "");
  return s.length >= 4 ? s : null;
};

/** Look for passport in a blob of narration text. */
function extractPassportFromText(text: string): string | null {
  if (!text) return null;
  // 1. "Passport : XXXXX, COUNTRY"
  let m = text.match(/Passport\s*[:\-]?\s*([A-Z0-9]+)/i);
  if (m) return normPassport(m[1]);
  // 2. "File No. : XXXX"
  m = text.match(/File\s*No\.?\s*[:\-]?\s*([A-Z0-9]+)/i);
  if (m) {
    const code = m[1].toUpperCase();
    return normPassport(code.length > 7 && /\d$/.test(code) ? code.slice(0, -1) : code);
  }
  // 3. Loose token in upper-case 6-12 alphanumerics
  m = text.match(/\b([A-Z]{1,3}\d{6,9}[A-Z]?)\b/);
  if (m) return normPassport(m[1]);
  return null;
}

/** Extract the primary bank transaction reference number from a payment comment (10-13 digits). */
function extractBankRef(text: string): string {
  if (!text) return "";
  const m = text.match(/\b(\d{10,13})\b/);
  return m ? m[1] : "";
}

/** A cell that is purely a money value (digits, optional commas/decimals) — not a date or ref. */
const MONEY_CELL = /^-?\d{1,3}(,\d{3})*(\.\d{1,2})?$|^-?\d+(\.\d{1,2})?$/;

/**
 * Recover a settlement/transfer amount from a row whose mapped DR/CR column was
 * shifted (a real quirk: NST bank-transfer rows are off by one column). Scans for
 * pure money cells, skipping dates/refs and the given columns; the running balance
 * is always the LAST money value, so the transaction amount is the first.
 */
function recoverSettlementAmount(row: unknown[], skip: number[] = []): number {
  const vals: number[] = [];
  for (let i = 0; i < row.length; i++) {
    if (skip.includes(i)) continue;
    const c = row[i];
    if (c instanceof Date) continue;
    const s = String(c ?? "").trim();
    if (!s || !MONEY_CELL.test(s)) continue;
    const v = Math.abs(num(s));
    if (v > 0) vals.push(v);
  }
  if (!vals.length) return 0;
  return vals[0]; // first money cell = the amount; balance comes after
}

/** Extract common reference patterns like Ticket Numbers (13 digits), Invoices, PNRs (6 chars), or SSR codes. */
function extractAdvancedRefs(text: string): string[] {
  if (!text) return [];
  const patterns = [
    /\b\d{13}\b/g, // Ticket numbers (13 digits)
    /\b[A-Z]{3}-\d{6,11}\b/gi, // Format: ABC-123456
    /\b[A-Z]{2,3}\d{6,11}\b/gi, // Invoice IDs / Booking IDs
    /\b[A-Z]{1}\d{5,10}\b/gi, // Short numeric refs prefixed with letter
    /\b[A-Z0-9]{6}\b/g, // PNR codes (usually 6 alpha-numerics)
    /\b[0-9A-Z]{5,10}\b/g, // General 5-10 digit references (loose)
  ];
  const found: string[] = [];
  const upper = text.toUpperCase();

  patterns.forEach((p) => {
    const matches = upper.match(p);
    if (matches) found.push(...matches.map((m) => m.trim()));
  });

  // Filter out common Noise (Dates, purely short numbers, etc)
  return [...new Set(found)].filter((s) => {
    if (/^\d{1,4}$/.test(s)) return false; // Too short to be a unique ref
    if (/^\d{8}$/.test(s)) return false; // Likely a date (YYYYMMDD)
    return true;
  });
}

/** Pull passport out of a partner ticket-number style "3VS XXXXX1". */
function extractPassportFromTicket(ticket: string): string | null {
  if (!ticket) return null;
  const m = String(ticket)
    .trim()
    .match(/^3VS\s*([A-Z0-9]+)$/i);
  if (!m) return null;
  return normPassport(m[1].slice(0, -1));
}

/** A passport-number token: 1-2 letters + 6-8 digits + optional trailing letter,
 *  OR (rarer) 2-3 letters embedded e.g. "24CE85273". */
const PASSPORT_TOKEN = /\b([A-Z]{1,2}\d{6,8}[A-Z]?|\d{2}[A-Z]{2}\d{5,6})\b/g;

/** Find every distinct passport-like token in a blob of text. */
export function extractAllPassports(text: string): string[] {
  if (!text) return [];
  const up = text.toUpperCase();
  const out: string[] = [];
  let m: RegExpExecArray | null;
  PASSPORT_TOKEN.lastIndex = 0;
  while ((m = PASSPORT_TOKEN.exec(up)) !== null) {
    const tok = m[1];
    // Skip pure booking/voucher codes like BK-2031 (handled separately).
    if (/^\d{8}$/.test(tok)) continue; // looks like a date YYYYMMDD
    out.push(tok);
  }
  return [...new Set(out)];
}

/**
 * Parse a multi-passenger cell (e.g. a Maverick "Passengers" column) into ordered
 * { name, passport } entries. Names and passports usually alternate line-by-line:
 *   "KEVIN CARL MEDINA / P9617010B / RAYMOND RIPOTOLA / P6655327B / ..."
 */
export function extractPaxEntries(text: string): { name: string; passport: string }[] {
  if (!text) return [];
  // Split on newlines, slashes, commas, semicolons, or " | ".
  const parts = String(text)
    .split(/[\n\r\/;|]+|,(?![^(]*\))/)
    .map((s) => s.trim())
    .filter(Boolean);
  const entries: { name: string; passport: string }[] = [];
  let pendingName = "";
  for (const part of parts) {
    // "NAME (PASSPORT)" inline form
    const inline = part.match(/^(.*?)\s*\(([A-Z0-9]{5,12})\)\s*$/i);
    if (inline && extractAllPassports(inline[2]).length) {
      entries.push({ name: inline[1].trim(), passport: normPassport(inline[2]) ?? "" });
      pendingName = "";
      continue;
    }
    const pass = extractAllPassports(part);
    if (pass.length === 1 && part.replace(/[^A-Z0-9]/gi, "").length <= pass[0].length + 2) {
      // This token IS a passport → attach to the most recent pending name.
      entries.push({ name: pendingName, passport: pass[0] });
      pendingName = "";
    } else if (pass.length === 0 && /[A-Za-z]/.test(part) && !/^BK-|^AB-|^\d+\s*\*\s*\d+/i.test(part)) {
      // A name line (skip booking codes like "BK-2017 260*3").
      pendingName = part;
    }
  }
  return entries;
}

function paxNameFromNarration(n1: string): string {
  // e.g. "MR. JIMMY ABAS MUSALI x 2" or "20 PAX 60 DAYS VISA x 1"
  return n1
    .replace(/\s*x\s*\d+\s*$/i, "")
    .replace(/^MR\.?\s*/i, "")
    .trim();
}

/**
 * Explode any row that bundles MULTIPLE passengers (≥2 passports) into one
 * sub-row per passenger, splitting the amount evenly. This lets a supplier's
 * batched booking (e.g. a Maverick "3 pax in one row") reconcile against the
 * individual per-passenger entries on the other ledger. Settlement / payment
 * rows and single-passenger rows pass through untouched.
 */
export function explodeMultiPax(rows: LedgerRow[]): LedgerRow[] {
  const out: LedgerRow[] = [];
  for (const row of rows) {
    const amount = row.charge > 0 ? row.charge : row.credit;
    // The passenger block can land in different columns depending on the export
    // (Maverick puts multi-pax bookings in Description, single ones in Passengers).
    // Pick whichever available field carries the most passports so we never miss it.
    const candidates = [
      typeof row.raw?.paxText === "string" ? (row.raw.paxText as string) : "",
      row.description,
      row.paxName,
    ].filter(Boolean);
    let paxSrc = "";
    let maxP = 0;
    for (const c of candidates) {
      const n = extractAllPassports(c).length;
      if (n > maxP) {
        maxP = n;
        paxSrc = c;
      }
    }
    if (!paxSrc) paxSrc = `${row.paxName} ${row.description}`;
    const passports = extractAllPassports(`${paxSrc} ${row.reference}`);
    if (row.settlement || passports.length < 2 || amount <= 0) {
      out.push(row);
      continue;
    }
    const entries = extractPaxEntries(paxSrc);
    const list =
      entries.length >= 2 ? entries : passports.map((p) => ({ name: "", passport: p }));
    const n = list.length;
    const per = Math.round((amount / n) * 100) / 100;
    const isCharge = row.charge > 0;
    list.forEach((e, i) => {
      // Put any rounding remainder on the first sub-row so the total is exact.
      const amt = i === 0 ? +(amount - per * (n - 1)).toFixed(2) : per;
      out.push({
        ...row,
        passport: normPassport(e.passport) ?? row.passport,
        paxName: e.name || row.paxName,
        charge: isCharge ? amt : 0,
        credit: isCharge ? 0 : amt,
        kind: row.kind,
        raw: { ...row.raw, explodedFrom: row.reference || row.paxName, paxIndex: i + 1, paxCount: n, explodedGroupAmt: amount },
      });
    });
  }
  out.forEach((r, i) => (r.index = i));
  return out;
}

/* ------------------------------------------------------------------ */
/* OUR LEDGER PARSING                                                  */
/* ------------------------------------------------------------------ */

/**
 * Parse "our" ledger. We try in order:
 *   (a) Real XLS/XLSX workbook
 *   (b) Tab-separated text saved with .xls extension
 *   (c) CSV
 */
export async function parseOurLedger(file: File): Promise<LedgerRow[]> {
  const buf = await file.arrayBuffer();
  // Try XLSX
  let aoa: unknown[][] | null = null;
  try {
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    aoa = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, raw: true });
  } catch {
    aoa = null;
  }
  if (!aoa || aoa.length === 0) {
    // Fallback to text parse
    const text = new TextDecoder("utf-8").decode(buf);
    const delim = text.includes("\t") ? "\t" : ",";
    const parsed = Papa.parse<string[]>(text, { delimiter: delim, skipEmptyLines: true });
    aoa = parsed.data;
  }
  if (!aoa.length) return [];

  const headerRow = (aoa[0] as unknown[]).map((c) =>
    String(c ?? "")
      .trim()
      .toUpperCase(),
  );
  const hasCode = headerRow.includes("CODE") && headerRow.includes("AMOUNT");
  const hasDRCR = headerRow.includes("DR") && headerRow.includes("CR");

  if (hasCode) return explodeMultiPax(parseOurNarrationStyle(aoa));
  if (hasDRCR) return explodeMultiPax(parseOurDrCrStyle(aoa));

  // Unknown layout → generic auto-detecting parser (works for any ledger).
  return explodeMultiPax(parseGenericLedger(aoa, "ours"));
}

/* ------------------------------------------------------------------ */
/* GENERIC AUTO-DETECTING PARSER  (works for ANY ledger layout)        */
/* ------------------------------------------------------------------ */

// Keyword groups used to recognise columns in an unknown ledger. Trailing word
// boundaries are intentionally omitted so plurals ("Particulars", "Comments",
// "Charges") still match. Debit/Credit are anchored to the start to avoid false
// hits inside unrelated words.
const COLREGEX = {
  date: /\b(date|posting|value\s*dt|record\s*time|\btime\b|\bday\b)/i,
  debit: /^\s*(dr\b|debit|withdraw|paid\s*out|out\s*flow|payment\s*out)/i,
  credit: /^\s*(cr\b|credit|deposit|paid\s*in|in\s*flow|payment\s*in|received)/i,
  amount: /\b(amount|amt|value|net\b|total|sum)/i,
  balance: /\b(balance|bal\b|running|closing)/i,
  reference: /\b(ref|voucher|invoice|inv\b|bill|cheque|chq|transaction|txn|utr|document|doc\s*no|pnr|ticket|receipt)/i,
  name: /\b(name|party|particular|customer|supplier|vendor|pax|beneficiary|narration|description|detail|comment|memo|remark)/i,
  id: /\b(passport|national\s*id|nat\s*id|id\s*no|emirates\s*id|iqama|trn)/i,
};

/** Score how header-like a row is (count of cells matching known column words). */
function headerScore(row: unknown[]): number {
  let score = 0;
  for (const cell of row) {
    const s = String(cell ?? "").trim();
    if (!s || s.length > 40) continue;
    if (
      COLREGEX.date.test(s) ||
      COLREGEX.debit.test(s) ||
      COLREGEX.credit.test(s) ||
      COLREGEX.amount.test(s) ||
      COLREGEX.balance.test(s) ||
      COLREGEX.reference.test(s) ||
      COLREGEX.name.test(s) ||
      COLREGEX.id.test(s)
    )
      score++;
  }
  return score;
}

/** Find the most header-like row within the first 20 rows. */
function detectHeaderRow(aoa: unknown[][]): number {
  let best = 0;
  let bestScore = -1;
  const limit = Math.min(20, aoa.length);
  for (let r = 0; r < limit; r++) {
    const row = (aoa[r] as unknown[]) ?? [];
    const sc = headerScore(row);
    if (sc > bestScore) {
      bestScore = sc;
      best = r;
    }
  }
  return bestScore >= 2 ? best : 0;
}

/**
 * Parse an arbitrary ledger: detect the header row, map columns by keyword, then
 * extract date / id / reference / party / charge / credit per row. Supports both
 * separate Debit/Credit columns and a single signed Amount column, and flags
 * inter-party settlement (bank-transfer) rows generically.
 */
export function parseGenericLedger(aoa: unknown[][], side: Side): LedgerRow[] {
  if (!aoa.length) return [];
  const hRow = detectHeaderRow(aoa);
  const header = (aoa[hRow] as unknown[]).map((c) => String(c ?? "").trim());

  const find = (re: RegExp, exclude?: RegExp): number => {
    for (let i = 0; i < header.length; i++) {
      const h = header[i];
      if (!h) continue;
      if (exclude && exclude.test(h)) continue;
      if (re.test(h)) return i;
    }
    return -1;
  };

  const idxDate = find(COLREGEX.date);
  const idxDebit = find(COLREGEX.debit);
  const idxCredit = find(COLREGEX.credit);
  // Only treat a generic "amount" column as signed when there is no DR/CR split.
  const idxAmount = idxDebit >= 0 && idxCredit >= 0 ? -1 : find(COLREGEX.amount, COLREGEX.balance);
  const idxRef = find(COLREGEX.reference);
  const idxId = find(COLREGEX.id);
  const idxName = find(COLREGEX.name);

  const rows: LedgerRow[] = [];
  for (let r = hRow + 1; r < aoa.length; r++) {
    const row = (aoa[r] as unknown[]) ?? [];
    if (!row.length || row.every((c) => c === null || c === undefined || c === "")) continue;

    const dateRaw = idxDate >= 0 ? row[idxDate] : "";
    let dateStr = "";
    if (dateRaw instanceof Date) dateStr = dateRaw.toISOString().slice(0, 10);
    else if (dateRaw !== undefined && dateRaw !== null) dateStr = String(dateRaw);

    const name = idxName >= 0 ? String(row[idxName] ?? "") : "";
    const ref = idxRef >= 0 ? String(row[idxRef] ?? "") : "";
    const idRaw = idxId >= 0 ? String(row[idxId] ?? "") : "";

    let charge = 0;
    let credit = 0;
    if (idxAmount >= 0) {
      const v = num(row[idxAmount]);
      if (v < 0) charge = Math.abs(v);
      else if (v > 0) credit = v;
    } else {
      const dr = idxDebit >= 0 ? num(row[idxDebit]) : 0;
      const cr = idxCredit >= 0 ? num(row[idxCredit]) : 0;
      if (dr > 0) charge = dr;
      if (cr > 0) credit = cr;
    }

    let kind: LedgerRow["kind"] = "other";
    if (charge > 0) kind = "charge";
    else if (credit > 0) kind = "credit";
    if (/brought forward|b\/f|opening balance|c\/f/i.test(`${name} ${ref}`)) kind = "other";

    const isPassBankTransfer = /^bank\s*transfer$/i.test(idRaw.trim());
    const passport = isPassBankTransfer
      ? null
      : (normPassport(idRaw) ?? extractPassportFromText(name));
    const settlement =
      kind !== "other" &&
      (isPassBankTransfer || (!passport && isSettlementText(name, ref)));
    const bankRef = settlement ? extractBankRef(`${name} ${ref}`) : "";

    const isSecDep = isSecurityDepositText(name, ref);
    const visaType = extractVisaType(`${name} ${ref}`);
    const scenario = detectScenario({ isSettle: settlement, isSecDep });

    rows.push({
      side,
      index: rows.length,
      date: dateStr,
      passport,
      paxName: isPassBankTransfer ? "BANK TRANSFER" : name,
      description: name,
      reference: bankRef || ref,
      charge,
      credit,
      kind,
      settlement,
      scenario,
      visaType,
      srcRow: r,
      raw: { row, paxText: name },
    });
  }
  return rows;
}

function parseOurNarrationStyle(aoa: unknown[][]): LedgerRow[] {
  const header = (aoa[0] as unknown[]).map((c) =>
    String(c ?? "")
      .trim()
      .toUpperCase(),
  );
  const col = (n: string) => header.indexOf(n);
  const idxDate = col("VOUCHER DATE");
  const idxVno = col("VOUCHER NO");
  const idxRef = col("REFERENCE");
  const idxAmt = col("AMOUNT");
  const narrIdx: number[] = [];
  for (let i = 1; i <= 6; i++) {
    const k = col(`NARRATION # ${i}`);
    if (k >= 0) narrIdx.push(k);
  }
  const rows: LedgerRow[] = [];
  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r] ?? [];
    if (!row.length) continue;
    const dateRaw = row[idxDate];
    const vno = String(row[idxVno] ?? "").trim();
    if (!vno && !dateRaw) continue;
    const amt = num(row[idxAmt]);
    const narrationText = narrIdx.map((i) => String(row[i] ?? "")).join(" | ");
    const passport = extractPassportFromText(narrationText);
    const paxName = paxNameFromNarration(String(row[narrIdx[0]] ?? ""));
    const isBF = /B\/F BALANCE/i.test(narrationText) || /BF\/BAL/i.test(vno);
    let kind: LedgerRow["kind"] = "other";
    let charge = 0,
      credit = 0;
    if (isBF) {
      kind = "other";
    } else if (amt < 0) {
      kind = "charge";
      charge = Math.abs(amt);
    } else if (amt > 0) {
      kind = "credit";
      credit = amt;
    }
    let dateStr = "";
    if (dateRaw instanceof Date) dateStr = dateRaw.toISOString().slice(0, 10);
    else if (dateRaw) dateStr = String(dateRaw);
    // In narration ledgers every positive AMOUNT is an incoming payment/top-up
    // (visa charges are negative). "PY" vouchers are payment vouchers. Either way
    // a credit here is a settlement, not a per-passenger charge.
    const isPaymentVoucher = /^PY/i.test(vno);
    const settlement =
      kind === "credit" && (isPaymentVoucher || isSettlementText(narrationText, paxName) || credit > 0);

    const isSecDep = isSecurityDepositText(narrationText);
    // Extract visa type from narration (e.g. "Visa For : UAE, 60 DAYS VISA")
    const visaType = extractVisaType(narrationText);
    const scenario = detectScenario({
      isSettle: settlement,
      isSecDep,
      isGroupRow: /\bx\s*\d+\b/.test(narrIdx[0] >= 0 ? String(row[narrIdx[0]] ?? "") : ""),
    });

    rows.push({
      side: "ours",
      index: rows.length,
      date: dateStr,
      passport: kind === "charge" ? passport : null,
      paxName,
      description: narrationText.slice(0, 200),
      reference: vno || String(row[idxRef] ?? ""),
      charge,
      credit,
      kind,
      settlement,
      scenario,
      visaType,
      srcRow: r,
      raw: { vno, amt, narrationText },
    });
  }
  return rows;
}

function parseOurDrCrStyle(aoa: unknown[][]): LedgerRow[] {
  const header = (aoa[0] as unknown[]).map((c) =>
    String(c ?? "")
      .trim()
      .toUpperCase(),
  );
  const col = (n: string) => header.indexOf(n);
  const idxDate   = col("DOC DATE");
  const idxDocNo  = col("DOCNO");    // "DocNo" in NST Maverick → "DOCNO" uppercase
  const idxTicket = col("TICKET NO.");
  const idxPnr    = col("PNR NO.");  // populated only for IS25 (flight) rows
  const idxDesc   = col("SECTOR / DESCRIPTION");
  const idxPax    = col("PAX NAME");
  const idxDR     = col("DR");
  const idxCR     = col("CR");
  const rows: LedgerRow[] = [];
  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r] ?? [];
    if (!row.length || row.every((c) => c === null || c === undefined || c === "")) continue;
    // Normalize DocNo — PY25/VR25 rows sometimes have trailing \n in the cell.
    const docno = (idxDocNo >= 0 ? String(row[idxDocNo] ?? "") : "").replace(/\s+/g, " ").trim();
    const ticket = String(row[idxTicket] ?? "").trim();
    const pax    = String(row[idxPax]    ?? "").trim();
    const desc   = String(row[idxDesc]   ?? "").trim();
    const dr = num(row[idxDR]);
    const cr = num(row[idxCR]);

    // IS25/ DocNo = interline / airline ticket. Ticket No. is a PNR-style "001 XXXXXX"
    // — NOT a "3VS PASSPORT" value. Treat these as simple charge rows whose
    // passport cannot be extracted from the ticket field.
    const isFlightRow = /^IS/i.test(docno);

    // "N PAX OMAN BUS SERVICE" — group-booking summary row where our system
    // recorded the combined amount for all passengers in one line.
    const isGroupRow = /^\d+\s+PAX\b/i.test(pax);

    // Reversal rows: DocNo prefix "VR" OR CR is negative (void / corrected charge).
    const isReversal = /^VR/i.test(docno) || (cr < 0 && dr === 0 && !isGroupRow);

    // JV25/ = journal voucher (TABBY payment received).
    const isJV = /^JV/i.test(docno);

    // Extract passport from Ticket No. only for "3VS XXXXX" format rows.
    const passport = isFlightRow ? null : extractPassportFromTicket(ticket);

    const isSettle = isSettlementText(desc, pax);

    // Security deposit detection — common in visa agency ledgers.
    const isSecDep = isSecurityDepositText(desc);

    let kind: LedgerRow["kind"] = "other";
    let charge = 0, credit = 0;

    if (isSettle || isJV) {
      kind = "credit";
      credit = recoverSettlementAmount(row, [idxDate]) || dr || Math.abs(cr);
    } else if (isReversal) {
      // VR rows are charge REVERSALS, not bank transfers.
      // They get kind="credit" but settlement=false so they can still match
      // the partner's corresponding refund row via passport + amount.
      kind = "credit";
      credit = Math.abs(cr) || dr;
    } else if (cr > 0) {
      kind = "charge";
      charge = cr;
    } else if (dr > 0) {
      kind = "credit";
      credit = dr;
    }

    const dateRaw = row[idxDate];
    let dateStr = "";
    if (dateRaw instanceof Date) dateStr = dateRaw.toISOString().slice(0, 10);
    else if (dateRaw) dateStr = String(dateRaw);

    // For flight rows: use the PNR as the reference (more useful than the ticket number).
    const pnr = idxPnr >= 0 ? String(row[idxPnr] ?? "").trim() : "";

    // Visa type from SECTOR / DESCRIPTION (e.g. "30 DAYS", "60 DAYS MULTI", "SECURITY DEPOSIT").
    // For reversal rows the desc may contain the refund reason instead.
    const visaType = isReversal ? undefined : extractVisaType(desc);

    // Scenario detection.
    const scenario = detectScenario({
      isReversal,
      isSettle,
      isJV,
      isFlightRow,
      isSecDep,
      isGroupRow,
      desc,
    });

    // Keep passport on reversal rows so they can match the partner's refund entry.
    const rowPassport =
      kind === "charge" || isReversal
        ? passport
        : null;

    rows.push({
      side: "ours",
      index: rows.length,
      date: dateStr,
      passport: rowPassport,
      paxName: isSettle || isJV ? (pax || desc.toUpperCase()) : pax,
      description: desc,
      reference: (isSettle || isJV)
        ? extractBankRef(`${desc} ${pax}`) || docno
        : isFlightRow
          ? pnr || ticket
          : (docno || ticket),
      charge,
      credit,
      kind,
      // VR reversals are NOT bank transfers — they're charge corrections.
      settlement: isSettle || isJV,
      isReversal,
      scenario,
      visaType,
      srcRow: r,
      raw: { docno, ticket, pax, dr, cr, paxText: pax, isGroupRow, isFlightRow },
    });
  }
  return rows;
}

/**
 * Dynamic parser that uses AI-detected mappings to extract ledger data.
 */
export function parseDynamicLedger(
  aoa: unknown[][],
  side: Side,
  mapping: ColumnMapping,
): LedgerRow[] {
  if (!aoa.length) return [];
  const header = (aoa[0] as unknown[]).map((c) =>
    String(c ?? "")
      .trim()
      .toUpperCase(),
  );
  // Fuzzy column resolver: tolerate small differences between the AI-returned
  // column name and the real header (spacing, punctuation, "No" vs "No.").
  const normH = header.map((h) => h.replace(/[^A-Z0-9]/gi, "").toUpperCase());
  const col = (n: string | undefined): number => {
    if (!n) return -1;
    const key = n.replace(/[^A-Z0-9]/gi, "").toUpperCase();
    if (!key) return -1;
    let i = normH.indexOf(key);
    if (i >= 0) return i;
    i = normH.findIndex((h) => h.length >= 3 && (h.includes(key) || key.includes(h)));
    return i;
  };

  const idxDate = col(mapping.date);
  const idxPass = col(mapping.passport);
  const idxPax = col(mapping.paxName);
  const idxDesc = col(mapping.description);
  const idxRef = col(mapping.reference);
  const idxCharge = col(mapping.charge);
  const idxCredit = col(mapping.credit);

  const rows: LedgerRow[] = [];
  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r] ?? [];
    if (!row.length || row.every((c) => c === null || c === undefined || c === "")) continue;

    const passRaw = idxPass >= 0 ? String(row[idxPass] ?? "") : "";
    // "Bank Transfer" in the passport column is a payment-row marker, not a real passport
    const isPassBankTransfer = /^bank\s*transfer$/i.test(passRaw.trim());
    const desc = idxDesc >= 0 ? String(row[idxDesc] ?? "") : "";
    const comm = idxPax >= 0 ? String(row[idxPax] ?? "") : "";

    // Resolve charge / credit robustly across the three common layouts:
    //   (a) separate DR & CR columns (positive values)
    //   (b) a single SIGNED "Amount" column (negative = charge, positive = credit)
    //   (c) a single column the AI mapped to both roles
    let charge = 0;
    let credit = 0;
    const sameAmtCol = idxCharge >= 0 && idxCharge === idxCredit;
    if (sameAmtCol) {
      const v = num(row[idxCharge]);
      if (v < 0) charge = Math.abs(v);
      else if (v > 0) credit = v;
    } else {
      const c = idxCharge >= 0 ? num(row[idxCharge]) : 0;
      const cr = idxCredit >= 0 ? num(row[idxCredit]) : 0;
      charge = c > 0 ? c : 0;
      credit = cr > 0 ? cr : 0;
      // A lone signed value sitting in only one of the two columns.
      if (charge === 0 && credit === 0) {
        if (c < 0) charge = Math.abs(c);
        else if (cr < 0) credit = Math.abs(cr);
      }
    }

    let kind: LedgerRow["kind"] = "other";
    if (charge > 0) kind = "charge";
    else if (credit > 0) kind = "credit";

    const dateRaw = idxDate >= 0 ? row[idxDate] : "";
    let dateStr = "";
    if (dateRaw instanceof Date) dateStr = dateRaw.toISOString().slice(0, 10);
    else if (dateRaw) dateStr = String(dateRaw);

    // Capture an identity key for every row. Strip a "3VS …" ticket prefix +
    // check digit so a ticket-embedded passport matches the bare passport on the
    // other ledger exactly (e.g. "3VS U90772971" → "U9077297").
    const passport = isPassBankTransfer
      ? null
      : (/^3VS/i.test(passRaw.trim())
          ? extractPassportFromTicket(passRaw)
          : normPassport(passRaw)) ?? extractPassportFromText(`${desc} ${comm}`);
    const refStr = idxRef >= 0 ? String(row[idxRef] ?? "") : "";

    // Generic settlement detection: a money-movement row with no per-item ID.
    // Independent of which column (DR/CR) the AI mapped the amount into, so it
    // works even when a ledger labels debits/credits the opposite way.
    const settlement =
      kind !== "other" &&
      (isPassBankTransfer || (!passport && isSettlementText(desc, comm, refStr)));

    // For settlements, recover the true amount from the row (handles ledgers like
    // NST whose transfer rows are shifted one column out of the mapped DR/CR).
    if (settlement) {
      const recovered = recoverSettlementAmount(row, [idxDate, idxPass].filter((x) => x >= 0));
      if (recovered > 0) {
        charge = 0;
        credit = recovered;
        kind = "credit";
      }
    }
    const bankRef = settlement ? extractBankRef(`${comm} ${desc} ${refStr}`) : "";

    const isSecDep = isSecurityDepositText(desc, comm, refStr);
    const visaType = extractVisaType(`${desc} ${comm}`);
    const scenario = detectScenario({ isSettle: settlement, isSecDep });

    rows.push({
      side,
      index: rows.length,
      date: dateStr,
      passport,
      paxName: isPassBankTransfer
        ? "BANK TRANSFER"
        : comm || (idxPax >= 0 ? String(row[idxPax] ?? "") : ""),
      description: desc,
      reference: bankRef || refStr,
      charge,
      credit,
      kind,
      settlement,
      scenario,
      visaType,
      srcRow: r,
      raw: { row, paxText: `${comm} ${desc}` },
    });
  }
  return explodeMultiPax(rows);
}

/* ------------------------------------------------------------------ */
/* PARTNER LEDGER PARSING                                              */
/* ------------------------------------------------------------------ */

export async function parsePartnerLedger(file: File): Promise<LedgerRow[]> {
  const buf = await file.arrayBuffer();
  // Try XLSX first
  let aoa: unknown[][] | null = null;
  try {
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    aoa = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, raw: true });
    // If aoa has only a single column, probably text → fallback
    if (aoa.length && aoa[0].length < 4) aoa = null;
  } catch {
    aoa = null;
  }
  if (!aoa) {
    const text = new TextDecoder("utf-8").decode(buf).replace(/^\ufeff/, "");
    const delim = text.split("\n", 2)[0].includes("\t") ? "\t" : ",";
    const parsed = Papa.parse<string[]>(text, { delimiter: delim, skipEmptyLines: true });
    aoa = parsed.data;
  }
  if (!aoa.length) return [];

  const header = (aoa[0] as unknown[]).map((c) => String(c ?? "").trim());
  const upper = header.map((h) => h.toUpperCase());

  // Maverick format: Transaction ID,Transaction Date,Agency,Credit,Debit,Balance,Description,Receipt,Passengers
  if (upper.some((h) => /TRANSACTION\s*ID/.test(h)) && upper.some((h) => /PASSENGER/.test(h)))
    return explodeMultiPax(parseMaverickSupplier(aoa, upper));

  // Format A: Record Time,Description,Passport No.,Comments,DR,CR,Balance
  if (upper.includes("PASSPORT NO.")) return explodeMultiPax(parsePartnerFormatA(aoa, upper));

  // Format B: Itinerary,Type,ID,Reference,Record Time,Description,Dates,Type,Comments,Status,VAT,DR,CR,Balance
  // Passport sits in the 2nd "Type" column at index 7.
  if (upper.includes("ITINERARY") && upper.includes("COMMENTS"))
    return explodeMultiPax(parsePartnerFormatB(aoa, upper));

  // Unknown layout → generic auto-detecting parser (works for any ledger).
  return explodeMultiPax(parseGenericLedger(aoa, "partner"));
}

/**
 * Maverick add-on / service charge rows — these are supplier extras (meals, bus,
 * hotel, driver, lounge…) that have no corresponding visa entry on our side.
 * Tagging them `addon: true` lets the UI separate them from genuine visa rows so
 * the real reconciliation rate isn't diluted by expected-unmatched service charges.
 */
const ADDON_RE =
  /\b(add[- ]?on|meal|breakfast|lunch|dinner|food|snack|water|beverage|bus(?:\s*transfer)?|hotel|accommodation|room|stay|meet\s*(?:&|and)?\s*greet|lounge|parking|insurance|luggage|baggage|transfer\s*fee|driver|van|vehicle|service\s*charge|handling\s*fee|admin(?:istration)?\s*fee)\b/i;

/**
 * Parse the Maverick supplier export:
 *   Transaction ID, Transaction Date, Agency, Credit, Debit, Balance, Descirption (typo), Reciept (typo), Passengers
 *
 * Row taxonomy:
 *   Booking BK-XXXXX [for PASSPORT] [Confirmed|Blocked] → visa/bus booking charge (Debit)
 *   Addon SERVICETYPE purchased for booking BK-XXXXX    → add-on service charge (Debit), no passport
 *   Payment PY-XXXXX [Approved]                         → money we sent (Credit)
 *   Refund for voided booking BK-XXXXX                  → reversal/cancellation (Credit)
 *
 * Multi-passenger rows have multiple "NAME (PASSPORT)" entries in the Passengers cell,
 * newline-separated. These are split per-passenger by explodeMultiPax downstream.
 */
function parseMaverickSupplier(aoa: unknown[][], upper: string[]): LedgerRow[] {
  const col = (re: RegExp) => upper.findIndex((h) => re.test(h));
  const iId   = col(/TRANSACTION\s*ID/);
  const iDate = col(/DATE/);
  const iCredit = col(/CREDIT/);
  const iDebit  = col(/DEBIT/);
  // Column header has a typo in Maverick's export: "Descirption"
  const iDesc = col(/DESC(?:I?R?I?P|RIPTION)/);
  // Column header has a typo: "Reciept"
  const iRcpt = col(/REC[EI]{1,2}PT/);
  const iPax  = col(/PASSENGER/);

  const rows: LedgerRow[] = [];
  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r] ?? [];
    if (!row.length || row.every((c) => c === null || c === undefined || c === "")) continue;
    const transId = String(row[iId]    ?? "").trim();
    const credit  = iCredit >= 0 ? num(row[iCredit]) : 0;
    const debit   = iDebit  >= 0 ? num(row[iDebit])  : 0;
    const desc    = String(row[iDesc]  ?? "").trim();
    const receipt = String(row[iRcpt]  ?? "").trim();
    const paxText = iPax >= 0 ? String(row[iPax] ?? "").trim() : "";

    // Dates in Maverick may be Date objects (Excel serial) or text strings with time.
    const dateCell = row[iDate];
    const date = dateCell instanceof Date
      ? dateCell.toISOString().slice(0, 10)
      : String(dateCell ?? "").replace(/\s+\d{1,2}:\d{2}(:\d{2})?$/, "").trim();

    // ── Row type detection ──────────────────────────────────────────────────────
    // Maverick addon rows ALWAYS start with the word "Addon" — most reliable signal.
    const isAddon = /^Addon\b/i.test(desc);

    // Refund/cancellation rows (Credit column, description starts with "Refund").
    const isRefund = /^Refund\b/i.test(desc);

    // Payment rows that represent money we sent them.
    const isPayment = isSettlementText(desc, receipt) || /^Payment\b/i.test(desc);

    const passports = extractAllPassports(`${paxText} ${desc}`);

    // ── Settlement / payment / refund path ──────────────────────────────────────
    const isSettle = isPayment || isRefund || (credit > 0 && passports.length === 0 && !isAddon);
    if (isSettle) {
      rows.push({
        side: "partner",
        index: rows.length,
        date,
        passport: null,
        paxName: isRefund ? "REFUND / VOID" : "BANK TRANSFER",
        description: desc,
        reference: extractBankRef(`${desc} ${receipt}`) || receipt || transId,
        charge: 0,
        credit: credit || debit,
        kind: "credit",
        settlement: true,
        scenario: isRefund ? "refund" : "bank_transfer",
        srcRow: r,
        raw: { transId, desc, receipt, paxText },
      });
      continue;
    }

    // ── Booking charge (Debit) ──────────────────────────────────────────────────
    // "Booking BK-XXXXX for PASSPORT Confirmed" or "Booking BK-XXXXX Blocked"
    const amount = debit || credit;
    const entries = extractPaxEntries(paxText);
    // Best passenger name: from Passengers column (full name without passport);
    // fall back to first line of paxText; last resort = description.
    const bestName = entries[0]?.name || paxText.split(/[\n\r]/)[0].replace(/\s*\([^)]*\)\s*$/, "").trim() || desc;

    // Booking reference: prefer the BK-XXXXX booking ID (in Receipt column).
    // For Blocked bookings this is still the correct BK reference.
    const bookingRef = receipt || transId;

    const isSecDepMav = isSecurityDepositText(desc, receipt);
    const visaTypeMav = extractVisaType(`${desc} ${receipt}`);
    const scenarioMav = detectScenario({ isSecDep: isSecDepMav, isAddon });

    rows.push({
      side: "partner",
      index: rows.length,
      date,
      passport: entries[0]?.passport ? normPassport(entries[0].passport) : passports[0] ?? null,
      paxName: bestName,
      description: desc,
      reference: bookingRef,
      charge: amount,
      credit: 0,
      kind: amount > 0 ? "charge" : "other",
      settlement: false,
      addon: isAddon,
      scenario: scenarioMav,
      visaType: visaTypeMav,
      srcRow: r,
      raw: { transId, desc, receipt, paxText },
    });
  }

  // ── Backfill passport + passenger name onto addon rows ─────────────────────
  //
  // Maverick creates ONE addon row per passenger per service type for multi-pax
  // bookings. E.g. BK-11785 (2 pax: Y5767509, Y1107639) has:
  //   AB-7067  BREAK FAST (5)              ← for passenger 1
  //   AB-7066  LUNCH & DINNER (15)         ← for passenger 1
  //   AB-7065  BREAK FAST (5)              ← for passenger 2
  //   AB-7064  LUNCH & DINNER (15)         ← for passenger 2
  //   BK-11785 Booking 420  pax: Y5767509, Y1107639
  //
  // Strategy:
  //   1. Build BK → ordered passenger list from raw paxText (all pax, not just first).
  //   2. Group addon rows by (BK, normalised addon type).
  //   3. Within each group assign passports sequentially:
  //      group[0] → pax[0], group[1] → pax[1], group[2] → pax[2], …
  //      using modulo for the rare case of more addons than passengers.

  // Step 1: BK → ordered passenger list.
  const bkPassengersMap = new Map<string, { passport: string; paxName: string }[]>();
  for (const row of rows) {
    if (row.addon || !row.reference.startsWith("BK-")) continue;
    if (bkPassengersMap.has(row.reference)) continue; // first booking row wins
    const paxText = (row.raw?.paxText as string) ?? "";
    const entries = extractPaxEntries(paxText);
    const passengers: { passport: string; paxName: string }[] = [];
    for (const e of entries) {
      const pp = normPassport(e.passport);
      if (pp) passengers.push({ passport: pp, paxName: e.name || row.paxName });
    }
    if (!passengers.length && row.passport)
      passengers.push({ passport: row.passport, paxName: row.paxName });
    if (passengers.length) bkPassengersMap.set(row.reference, passengers);
  }

  // Step 2: Group addon rows by (BK, normalised addon type).
  // The addon type is the service name extracted from the description.
  const bkAddonGroups = new Map<string, Map<string, LedgerRow[]>>();
  for (const row of rows) {
    if (!row.addon) continue;
    const bkM = row.description.match(/\bBK-(\d+)\b/i);
    if (!bkM) continue;
    const bk = `BK-${bkM[1]}`;
    if (!bkAddonGroups.has(bk)) bkAddonGroups.set(bk, new Map());
    const typeMap = bkAddonGroups.get(bk)!;
    // "Addon BREAK FAST purchased for booking BK-11785" → "BREAK FAST"
    const typeKey = row.description
      .replace(/^Addon\s+/i, "")
      .replace(/\s+purchased\s+for\s+booking\s+BK-\d+.*/i, "")
      .trim()
      .toUpperCase();
    if (!typeMap.has(typeKey)) typeMap.set(typeKey, []);
    typeMap.get(typeKey)!.push(row);
  }

  // Step 3: Assign passports sequentially within each (BK, type) group.
  for (const [bk, typeMap] of bkAddonGroups) {
    const passengers = bkPassengersMap.get(bk);
    if (!passengers || passengers.length === 0) continue;
    const n = passengers.length;
    for (const addonRows of typeMap.values()) {
      addonRows.forEach((row, idx) => {
        const pax = passengers[idx % n];
        row.passport = pax.passport;
        row.paxName = pax.paxName;
        (row.raw as Record<string, unknown>).addonParentBK = bk;
        (row.raw as Record<string, unknown>).addonPaxIndex = idx;
      });
    }
  }

  // Step 4: Merge per-BK addon totals into their booking rows, then drop the
  // individual addon rows from the output — but ONLY when the booking row is
  // present in this file.
  //
  // NST invoices ONE line per passenger (booking fare + all addon services
  // bundled). Maverick records them as separate rows. To compare apples to
  // apples we fold addon debit into the parent booking charge before
  // explodeMultiPax runs, so the per-pax split produces amounts that match NST.
  //
  // E.g. BK-11785 (2 pax, booking=420) + 4 addon rows (5+15+5+15=40)
  //   → merged booking charge = 460
  //   → explodeMultiPax splits evenly: 230 per pax
  //   → NST shows 230 per pax  → exact match!
  //
  // If a booking row is absent (older BK from a different export period), the
  // addon row is kept as-is so it can still match an NST entry independently.

  // Only merge addons whose parent BK booking row exists in this file.
  const bkAddonTotal = new Map<string, number>();
  for (const row of rows) {
    if (!row.addon) continue;
    const bkM = row.description.match(/\bBK-(\d+)\b/i);
    if (!bkM) continue;
    const bk = `BK-${bkM[1]}`;
    // Skip orphaned addons — booking row not in this file.
    if (!bkPassengersMap.has(bk)) continue;
    bkAddonTotal.set(bk, (bkAddonTotal.get(bk) ?? 0) + row.charge);
  }

  return rows
    .filter((row) => {
      if (!row.addon) return true;
      // Remove only addon rows whose parent booking was merged.
      const bkM = row.description.match(/\bBK-(\d+)\b/i);
      if (!bkM) return true;
      return !bkAddonTotal.has(`BK-${bkM[1]}`);
    })
    .map((row) => {
      if (row.settlement || row.addon) return row;
      const extra = bkAddonTotal.get(row.reference) ?? 0;
      if (!extra) return row;
      return {
        ...row,
        charge: +(row.charge + extra).toFixed(2),
        raw: { ...(row.raw as Record<string, unknown>), addonMerged: extra, bookingOnlyCharge: row.charge },
      };
    });
}

/**
 * Format A: Record Time | Description | Passport No. | Comments | DR | CR | Balance
 *
 * Used by:
 *  • Maverick Travel Visa  — DR = charge to us, CR = payment from us
 *    Passport No. = raw passport (no "3VS" prefix needed)
 *    Payment rows: "AC-4 Rak Bank" in Description, "Bank Transfer" in Passport No.
 *    TABBY/Other payments: "PK Other" or "FS Other" in Passport No., no DR/CR switch
 *    Numeric passports (e.g. Algerian 309858136) are stored as integers in XLS.
 *  • SmartTrip — same column layout but charge/credit may swap DR/CR
 */
function parsePartnerFormatA(aoa: unknown[][], upper: string[]): LedgerRow[] {
  const col = (n: string) => upper.indexOf(n);
  const idxTime = col("RECORD TIME");
  const idxDesc = col("DESCRIPTION");
  const idxPass = col("PASSPORT NO.");
  const idxComm = col("COMMENTS");
  const idxDR = col("DR");
  const idxCR = col("CR");
  const rows: LedgerRow[] = [];
  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r] ?? [];
    if (!row.length || row.every((c) => c === null || c === undefined || c === "")) continue;

    const desc = String(row[idxDesc] ?? "").trim();
    // Passport No. may be an integer in XLS (e.g. Algerian passport 309858136).
    // Stringify without ".0" suffix — JS Number → String drops trailing zeros.
    const passRaw = row[idxPass];
    const pass = passRaw != null && passRaw !== "" ? String(passRaw).trim() : "";
    const comm = String(row[idxComm] ?? "").trim();

    // Dates may be Date objects (Excel serial) or text strings with a time component.
    const dateCell = row[idxTime];
    const date = dateCell instanceof Date
      ? dateCell.toISOString().slice(0, 10)
      : String(dateCell ?? "").replace(/\s+\d{1,2}:\d{2}(:\d{2})?$/, "").trim();

    const dr = num(row[idxDR]);
    const cr = num(row[idxCR]); // Maverick VISA CR can be comma-formatted "5,000.00" — num() handles this

    const isBF = /brought forward|opening\s*balance/i.test(desc);

    // Payment / settlement row detection:
    // • SmartTrip: "Bank Transfer" in Passport No. column
    // • Maverick VISA bank payment: "AC-4 Rak Bank" in Description, "Bank Transfer" in Passport
    // • Maverick VISA TABBY/other: "PK Other" or "FS Other" in Passport column
    const isPassportBankTransfer = /^bank\s*transfer$/i.test(pass);
    const isOtherCollection    = /^(PK|FS)\s*Other\b/i.test(pass);
    const isBankDesc           = isSettlementText(desc, comm) || /^AC-\d+\b/i.test(desc);
    const isBankish = isPassportBankTransfer || isOtherCollection || isBankDesc;

    let kind: LedgerRow["kind"] = "other";
    let charge = 0, credit = 0;

    if (isBF) {
      kind = "other";
    } else if (isBankish) {
      // Payment / settlement / collection row.
      kind = "credit";
      credit = cr || dr;
    } else if (dr > 0) {
      // Charge row — Maverick VISA: they billed us via DR.
      kind = "charge";
      charge = dr;
    } else if (cr > 0) {
      // SmartTrip-style: charge shown in CR column (their receivable from us).
      kind = "credit";
      credit = cr;
    }

    const settlement = kind === "credit" && (isPassportBankTransfer || isOtherCollection || isBankDesc);

    // Build clean passport: skip non-passport values that appear in Passport No. column.
    const passportSkip = isPassportBankTransfer || isOtherCollection || !pass;
    const passport =
      kind === "charge"
        ? (normPassport(passportSkip ? "" : pass) ?? extractPassportFromText(comm))
        : null;

    const isSecDepA = isSecurityDepositText(desc, comm);
    const visaTypeA = extractVisaType(`${desc} ${comm}`);
    const scenarioA = detectScenario({ isSettle: settlement, isSecDep: isSecDepA });

    rows.push({
      side: "partner",
      index: rows.length,
      date,
      passport,
      paxName: isPassportBankTransfer || isOtherCollection ? "BANK TRANSFER" : (comm || pass),
      description: desc,
      reference: settlement ? extractBankRef(`${comm} ${desc}`) : "",
      charge,
      credit,
      kind,
      settlement,
      scenario: scenarioA,
      visaType: visaTypeA,
      srcRow: r,
      raw: { dr, cr, pass, comm, desc },
    });
  }
  return rows;
}

function parsePartnerFormatB(aoa: unknown[][], upper: string[]): LedgerRow[] {
  const idxTime = upper.indexOf("RECORD TIME");
  const idxDesc = upper.indexOf("DESCRIPTION");
  const idxComm = upper.indexOf("COMMENTS");
  const idxStatus = upper.indexOf("STATUS");
  const idxRef = upper.indexOf("REFERENCE");
  const idxDR = upper.indexOf("DR");
  const idxCR = upper.indexOf("CR");
  // The 2nd "Type" column (passport) sits right before "Comments" (often index 7).
  let idxPass = -1;
  for (let i = upper.length - 1; i >= 0; i--) {
    if (upper[i] === "TYPE") {
      idxPass = i;
      break;
    }
  }
  // Skip the first "Type" which is the row category (VS / TR / PY).
  const firstType = upper.indexOf("TYPE");
  if (idxPass === firstType) idxPass = -1;

  const rows: LedgerRow[] = [];
  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r] ?? [];
    if (!row.length || row.every((c) => c === null || c === undefined || c === "")) continue;
    const desc = String(row[idxDesc] ?? "");
    const comm = String(row[idxComm] ?? "");
    const status = String(row[idxStatus] ?? "");
    const ref = String(row[idxRef] ?? "");
    const passRaw = idxPass >= 0 ? String(row[idxPass] ?? "") : "";
    const dr = num(row[idxDR]);
    const cr = num(row[idxCR]);
    const isBF = /brought forward/i.test(desc);
    const isRefund = /refund/i.test(desc) || /refund/i.test(passRaw);
    const isBank = isSettlementText(desc, ref);
    let kind: LedgerRow["kind"] = "other";
    let charge = 0,
      credit = 0;
    let passport: string | null = null;
    if (isBF) kind = "other";
    else if (cr > 0 && !isRefund && !isBank) {
      // Visa charge to us → partner shows it as CR (their receivable).
      kind = "charge";
      charge = cr;
      passport = normPassport(passRaw) ?? extractPassportFromText(comm);
    } else if (dr > 0) {
      kind = "credit";
      credit = dr;
    } else if (cr > 0) {
      kind = "credit";
      credit = cr;
    }
    const settlement = kind === "credit" && (isBank || isRefund);

    // Security deposit detection: Format B often has "Refund Security" in reference
    // or "Security Deposit" in the description.
    const isSecDepB = isSecurityDepositText(ref, desc) || /security/i.test(ref);
    const visaTypeB = extractVisaType(`${desc} ${comm} ${ref}`);
    const scenarioB = detectScenario({
      isSettle: settlement,
      isSecDep: isSecDepB && kind === "charge",
      isReversal: isRefund && !isBank,
      desc,
    });

    rows.push({
      side: "partner",
      index: rows.length,
      date: String(row[idxTime] ?? ""),
      passport,
      paxName: comm,
      description: [desc, status].filter(Boolean).join(" · "),
      reference: settlement ? extractBankRef(`${comm} ${desc} ${ref}`) || ref : ref,
      charge,
      credit,
      kind,
      settlement,
      scenario: scenarioB,
      visaType: visaTypeB,
      srcRow: r,
      raw: { dr, cr, passRaw, comm, desc },
    });
  }
  return rows;
}

/* ------------------------------------------------------------------ */
/* RECONCILIATION                                                      */
/* ------------------------------------------------------------------ */

export type PairStatus =
  | "matched" // paired, amounts agree
  | "amount_diff" // paired, amounts differ
  | "missing_partner" // exists only in OUR ledger  → "Only Ours"
  | "missing_ours"; // exists only in PARTNER ledger → "Only Partner"

/** Per-signal breakdown used both for scoring and for UI evidence display. */
export type MatchEvidence = {
  /** 0–1 strength of passport/ID agreement. */
  passport: number;
  /** 0–1 strength of ticket/voucher/reference agreement. */
  reference: number;
  /** 0–1 pax/name similarity. */
  name: number;
  /** 0–1 raw amount closeness (unadjusted). */
  amount: number;
  /** 0–1 effective amount closeness after group-row correction. */
  effectiveAmount: number;
  /** 0–1 date proximity. */
  date: number;
  /** How this pair was established. */
  method: "rule" | "ai" | "none";
  /** absolute day gap between the two dates, null if unknown. */
  dateDeltaDays: number | null;
};

export type Pair = {
  key: string;
  status: PairStatus;
  /** "charge" pair (visa/service) or "credit" pair (payment/refund/top-up). */
  kind: "charge" | "credit";
  ours: LedgerRow | null;
  partner: LedgerRow | null;
  oursAmt: number;
  partnerAmt: number;
  diff: number; // partner - ours
  note: string;
  /** Overall confidence 0–1 that this pairing is correct. */
  confidence?: number;
  /** Raw matching score 0–1 before classification. */
  score?: number;
  /** True when confidence is below the auto-accept threshold → human should verify. */
  needsReview?: boolean;
  /** Per-signal evidence breakdown. */
  evidence?: MatchEvidence;
  aiInsight?: string;
};

export type ReconResult = {
  pairs: Pair[];
  totals: {
    oursRows: number;
    partnerRows: number;
    oursCharges: number;
    partnerCharges: number;
    oursCredits: number;
    partnerCredits: number;
    /** Paired with agreeing amounts. */
    matched: number;
    /** Paired but amounts differ. */
    amountIssues: number;
    /** Only in our ledger. */
    onlyOurs: number;
    /** Only in partner ledger. */
    onlyPartner: number;
    /** Net absolute-amount difference across matched pairs. */
    netAmountDiff: number;
    /** Average confidence across all paired (non-missing) rows, 0–1. */
    avgConfidence: number;
    /** Count of pairs flagged for human review (confidence 0.5–0.85). */
    needsReview: number;
    /** Count of pairs whose match came from the AI residual stage. */
    aiAssisted: number;
  };
};

const MONTHS = [
  "jan", "feb", "mar", "apr", "may", "jun",
  "jul", "aug", "sep", "oct", "nov", "dec",
];

/**
 * Parse a date from many real-world ledger formats → epoch ms (NaN if unknown).
 * Handles: DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY, YYYY-MM-DD, DD-MMM-YY/YYYY,
 * "DD MMM YYYY", "MMM DD, YYYY", and bare Excel date serial numbers.
 * Day/month order is inferred when one value is > 12.
 */
function parseDate(s: string): number {
  if (!s && s !== "0") return NaN;
  const t = String(s).trim();
  if (!t) return NaN;

  // Excel serial date (e.g. "45800" or 45800.5) — days since 1899-12-30.
  if (/^\d{4,6}(\.\d+)?$/.test(t)) {
    const serial = parseFloat(t);
    if (serial > 20000 && serial < 80000) {
      return Math.round((serial - 25569) * 86400 * 1000);
    }
  }

  // ISO-ish: YYYY-MM-DD or YYYY/MM/DD
  let m = t.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]).getTime();

  // Numeric D/M/Y or M/D/Y with / - . separators
  m = t.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) {
    let a = +m[1],
      b = +m[2];
    let yy = +m[3];
    if (yy < 100) yy += yy < 50 ? 2000 : 1900;
    // Infer order: if first > 12 it's the day; if second > 12 the first is month.
    let dd = a,
      mm = b;
    if (a > 12 && b <= 12) {
      dd = a;
      mm = b;
    } else if (b > 12 && a <= 12) {
      dd = b;
      mm = a;
    } // else assume DD/MM (most non-US ledgers)
    return new Date(yy, mm - 1, dd).getTime();
  }

  // DD-MMM-YY / "DD MMM YYYY"  (e.g. "19-May-26", "1 Jun 2026")
  m = t.match(/^(\d{1,2})[\s\/\-]([A-Za-z]{3,})[\s\/\-](\d{2,4})/);
  if (m) {
    const dd = +m[1];
    const mo = MONTHS.indexOf(m[2].slice(0, 3).toLowerCase());
    if (mo >= 0) {
      let yy = +m[3];
      if (yy < 100) yy += yy < 50 ? 2000 : 1900;
      return new Date(yy, mo, dd).getTime();
    }
  }

  // "MMM DD, YYYY"  (e.g. "May 19, 2026")
  m = t.match(/^([A-Za-z]{3,})[\s\/\-](\d{1,2}),?\s*(\d{2,4})/);
  if (m) {
    const mo = MONTHS.indexOf(m[1].slice(0, 3).toLowerCase());
    if (mo >= 0) {
      let yy = +m[3];
      if (yy < 100) yy += yy < 50 ? 2000 : 1900;
      return new Date(yy, mo, +m[2]).getTime();
    }
  }

  const d = Date.parse(t);
  return isNaN(d) ? NaN : d;
}

/* ------------------------------------------------------------------ */
/* SCORING ENGINE  (multi-signal, rule-based)                          */
/* ------------------------------------------------------------------ */

/** Levenshtein edit distance (bounded). */
function lev(a: string, b: string): number {
  const m = a.length,
    n = b.length;
  if (!m) return n;
  if (!n) return m;
  let prev = Array.from({ length: n + 1 }, (_, i) => i);
  let curr = new Array(n + 1).fill(0);
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min(prev[j] + 1, curr[j - 1] + 1, prev[j - 1] + cost);
    }
    [prev, curr] = [curr, prev];
  }
  return prev[n];
}

/** Jaro-Winkler similarity 0–1, good for short names. */
function jaroWinkler(s1: string, s2: string): number {
  if (s1 === s2) return 1;
  if (!s1.length || !s2.length) return 0;
  const md = Math.max(0, Math.floor(Math.max(s1.length, s2.length) / 2) - 1);
  const s1m = new Array(s1.length).fill(false);
  const s2m = new Array(s2.length).fill(false);
  let matches = 0;
  for (let i = 0; i < s1.length; i++) {
    const lo = Math.max(0, i - md);
    const hi = Math.min(i + md + 1, s2.length);
    for (let j = lo; j < hi; j++) {
      if (s2m[j] || s1[i] !== s2[j]) continue;
      s1m[i] = s2m[j] = true;
      matches++;
      break;
    }
  }
  if (!matches) return 0;
  let t = 0,
    k = 0;
  for (let i = 0; i < s1.length; i++) {
    if (!s1m[i]) continue;
    while (!s2m[k]) k++;
    if (s1[i] !== s2[k]) t++;
    k++;
  }
  t /= 2;
  const jaro = (matches / s1.length + matches / s2.length + (matches - t) / matches) / 3;
  let prefix = 0;
  for (let i = 0; i < Math.min(4, s1.length, s2.length); i++) {
    if (s1[i] === s2[i]) prefix++;
    else break;
  }
  return jaro + prefix * 0.1 * (1 - jaro);
}

/** Normalise a pax/customer name for comparison. */
function normName(s: string): string {
  return (s || "")
    .toUpperCase()
    .replace(/\b(MR|MRS|MS|MISS|MSTR|MAS|DR|HAJI|HJ)\.?\b/g, " ")
    .replace(/\s*X\s*\d+\s*$/i, "")
    .replace(/[^A-Z0-9 ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/** Name similarity combining token-set overlap and Jaro-Winkler. */
export function nameSimilarity(a: string, b: string): number {
  const na = normName(a);
  const nb = normName(b);
  if (!na || !nb) return 0;
  if (na === nb) return 1;
  const ta = na.split(" ").filter(Boolean);
  const tb = nb.split(" ").filter(Boolean);
  const sa = new Set(ta);
  const sb = new Set(tb);
  const inter = [...sa].filter((t) => sb.has(t)).length;
  const union = new Set([...sa, ...sb]).size;
  const jacc = union ? inter / union : 0;
  // token-sorted jaro-winkler handles reordered name parts (FIRST LAST vs LAST FIRST)
  const jw = jaroWinkler([...ta].sort().join(" "), [...tb].sort().join(" "));
  return Math.max(jacc, jw, 0.55 * jacc + 0.45 * jw);
}

/** Passport / national-ID agreement, tolerant of check digits and formatting. */
export function passportMatch(a: string | null, b: string | null): number {
  if (!a || !b) return 0;
  const ca = a.toUpperCase().replace(/[^A-Z0-9]/g, "");
  const cb = b.toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (!ca || !cb || ca.length < 4 || cb.length < 4) return 0;
  if (ca === cb) return 1;
  // one fully contains the other (prefix/suffix codes, embedded check digits)
  if (ca.includes(cb) || cb.includes(ca)) return 0.88;
  // share a long common prefix (trailing check digit differs)
  const min = Math.min(ca.length, cb.length);
  if (min >= 6 && ca.slice(0, min - 1) === cb.slice(0, min - 1)) return 0.82;
  // close edit distance
  const d = lev(ca, cb);
  if (d === 1) return 0.78;
  if (d === 2 && min >= 7) return 0.6;
  return 0;
}

/** Reference / ticket / voucher agreement via extracted regex tokens. */
export function referenceMatch(a: LedgerRow, b: LedgerRow): number {
  const ra = new Set(extractAdvancedRefs(`${a.description} ${a.reference}`));
  if (!ra.size) return 0;
  const rb = extractAdvancedRefs(`${b.description} ${b.reference}`);
  for (const r of rb) if (ra.has(r)) return 1;
  return 0;
}

/** Amount closeness, exact within 0.5, decaying with relative difference. */
function amountCloseness(x: number, y: number): number {
  const d = Math.abs(x - y);
  if (d < 0.5) return 1;
  const base = Math.max(Math.abs(x), Math.abs(y), 1);
  const rel = d / base;
  if (rel <= 0.01) return 0.92;
  if (rel <= 0.02) return 0.82;
  if (rel <= 0.05) return 0.62;
  if (rel <= 0.1) return 0.4;
  if (rel <= 0.2) return 0.2;
  return 0.05;
}

/** Date proximity, full credit within a day, decaying to a month. */
function dateProximity(a: string, b: string): number {
  const da = parseDate(a);
  const db = parseDate(b);
  if (isNaN(da) || isNaN(db)) return 0.35; // unknown → mild neutral
  const days = Math.abs(da - db) / 86400000;
  if (days <= 1) return 1;
  if (days <= 3) return 0.85;
  if (days <= 7) return 0.65;
  if (days <= 14) return 0.45;
  if (days <= 30) return 0.25;
  return 0.05;
}

function dayGap(a: string, b: string): number | null {
  const da = parseDate(a);
  const db = parseDate(b);
  if (isNaN(da) || isNaN(db)) return null;
  return Math.round(Math.abs(da - db) / 86400000);
}

export type ScoreResult = { score: number; evidence: MatchEvidence };

/** The signed magnitude of a row, regardless of debit/credit side. */
function absAmount(r: LedgerRow): number {
  return r.charge > 0 ? r.charge : r.credit;
}

const normRef = (s: string) => (s || "").replace(/[^A-Z0-9]/gi, "").toUpperCase();

/** Description token overlap (0–1). */
function descSimilarity(a: string, b: string): number {
  const ta = normName(a)
    .split(" ")
    .filter((t) => t.length >= 3);
  const tb = new Set(
    normName(b)
      .split(" ")
      .filter((t) => t.length >= 3),
  );
  if (!ta.length || !tb.size) return 0;
  const inter = ta.filter((t) => tb.has(t)).length;
  return inter / Math.max(ta.length, tb.size);
}

/** True when a row is categorised as a security deposit. */
function isSecDep(r: LedgerRow): boolean {
  return (
    r.scenario === "security_deposit" ||
    (r.visaType ?? "").toUpperCase().includes("SECURITY DEPOSIT") ||
    isSecurityDepositText(r.description)
  );
}

/** Normalised visa-type label for comparison (strips country/suffix noise). */
function normVisaType(vt: string | undefined): string {
  if (!vt) return "";
  return vt
    .toUpperCase()
    .replace(/\s*,\s*\w+$/, "") // strip trailing country ", INDIA"
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Unified row-pair scoring used for EVERY pair (charges and payments alike).
 * Weights follow the spec: 40% ID, 30% amount, 20% date, 10% text (name/desc).
 * Identity is the primary key; amount validates it; date & text break ties.
 */
export function scoreRowPair(o: LedgerRow, p: LedgerRow): ScoreResult {
  // ── HARD GATES ──────────────────────────────────────────────────────────────
  // A security deposit must NEVER match a visa charge and vice versa.
  // They are separate financial instruments even when the amounts coincide.
  const oSD = isSecDep(o);
  const pSD = isSecDep(p);
  if (oSD !== pSD) {
    return {
      score: 0,
      evidence: {
        passport: 0, reference: 0, name: 0, amount: 0, effectiveAmount: 0,
        date: 0, method: "rule", dateDeltaDays: null,
      },
    };
  }

  const passport = passportMatch(o.passport, p.passport);
  const refExact = normRef(o.reference) && normRef(o.reference) === normRef(p.reference) ? 1 : 0;
  const reference = Math.max(referenceMatch(o, p), refExact);
  const idSim = Math.max(passport, reference);
  const amount = amountCloseness(absAmount(o), absAmount(p));
  const date = dateProximity(o.date, p.date);
  const name = nameSimilarity(o.paxName, p.paxName);
  const text = Math.max(name, descSimilarity(o.description, p.description));

  // Conflicting identity: both rows carry a valid passport but they are clearly
  // different people. Two different passengers must NOT pair just because their
  // fee and date happen to coincide — this is the #1 source of false matches.
  const idA = normRef(o.passport ?? "");
  const idB = normRef(p.passport ?? "");
  const idConflict = idA.length >= 5 && idB.length >= 5 && passport === 0 && reference < 0.5;

  // Group-booking rows in our NST file (e.g. "03 PAX OMAN BUS SERVICE") hold the
  // combined amount for N passengers. After the partner side is exploded into per-
  // passenger sub-rows the amounts differ by a factor of N. Two strategies:
  //
  //  1. If the partner row was exploded, compare our group total against the
  //     pre-explosion total stored in raw.explodedGroupAmt — that gives a clean
  //     amount match without needing passport overlap (NST group rows have none).
  //
  //  2. If the raw ratio of amounts is a clean integer (2..9), treat amount as
  //     neutral (0.5) to avoid a heavy penalty when passport already matched.
  let effectiveAmount = amount;
  const isGroupRow = !!(o.raw?.isGroupRow || p.raw?.isGroupRow);
  if (isGroupRow) {
    // Strategy 1: partner sub-row carries its group total → direct comparison.
    const explodedGroupAmt = (p.raw?.explodedGroupAmt ?? o.raw?.explodedGroupAmt) as number | undefined;
    if (explodedGroupAmt && explodedGroupAmt > 0) {
      const groupSide = o.raw?.isGroupRow ? absAmount(o) : absAmount(p);
      effectiveAmount = Math.max(effectiveAmount, amountCloseness(groupSide, explodedGroupAmt));
    }
    // Strategy 2: ratio check (works even without explodedGroupAmt, e.g. un-exploded partner).
    if (effectiveAmount < 0.5) {
      const bigAmt = Math.max(absAmount(o), absAmount(p));
      const smlAmt = Math.min(absAmount(o), absAmount(p));
      if (smlAmt > 0) {
        const ratio = bigAmt / smlAmt;
        const nearest = Math.round(ratio);
        if (nearest >= 2 && nearest <= 9 && Math.abs(ratio - nearest) < 0.15) {
          effectiveAmount = Math.max(effectiveAmount, 0.5);
        }
      }
    }
  }

  // ── Visa-type agreement bonus / penalty ─────────────────────────────────────
  // When both rows carry a visa type, matching types get a small boost and
  // mismatching types get a penalty (e.g. "30 DAYS" vs "60 DAYS").
  // This prevents a 30-day visa from edging out the correct 60-day match.
  const ovt = normVisaType(o.visaType);
  const pvt = normVisaType(p.visaType);
  let visaTypeFactor = 0; // neutral
  if (ovt && pvt) {
    if (ovt === pvt) visaTypeFactor = 0.06;      // bonus for matching type
    else if (ovt.includes(pvt) || pvt.includes(ovt)) visaTypeFactor = 0.02; // partial
    else visaTypeFactor = -0.10;                  // penalty for clearly different types
  }

  let score = 0.4 * idSim + 0.3 * effectiveAmount + 0.2 * date + 0.1 * text + visaTypeFactor;
  score = Math.max(0, score); // clamp after potential visa-type penalty

  // Confidence boosts for strong combinations.
  if (idSim >= 0.99 && amount >= 0.99) {
    score = Math.max(score, 0.99);
  } else if (idSim >= 0.99 && !isGroupRow) {
    // Exact passport match, any amount — but NOT for group rows.
    // Group rows carry the LEAD passenger's passport and the GROUP total,
    // so their passport would coincidentally match the lead pax's individual
    // Maverick booking (smaller amount) — that's a false positive we suppress.
    score = Math.max(score, 0.9);
  } else if (amount >= 0.99 && text >= 0.7) {
    score = Math.max(score, 0.86);
  } else if (amount >= 0.99 && date >= 0.85) {
    score = Math.max(score, 0.8);
  }

  // Group-row boost: NST N-PAX summary row paired with an exploded Maverick
  // sub-row whose group total (raw.explodedGroupAmt) matches the NST total.
  // Score just above "needs review" so a human can confirm the match.
  if (isGroupRow && idSim >= 0.99 && effectiveAmount >= 0.9 && date >= 0.65) {
    score = Math.max(score, 0.88);
  } else if (isGroupRow && effectiveAmount >= 0.9 && date >= 0.65) {
    score = Math.max(score, 0.72);
  }

  // Settlement / bank-transfer pairs: there is no per-item ID, so the evidence is
  // amount + date + the payment label. Scored independently of identity so they
  // match across ANY two ledgers — and tolerant of small bank-fee differences and
  // of unparseable / missing dates (date contributes, but never blocks the match).
  const bothSettle = !!(o.settlement && p.settlement);
  const settleLike =
    bothSettle || (o.kind === "credit" && p.kind === "credit" && idSim === 0);
  if (settleLike && amount >= 0.6) {
    // A shared bank reference number is decisive proof.
    if (reference >= 0.99 && amount >= 0.9) {
      score = Math.max(score, 0.99);
    } else {
      // amount carries most weight; date refines but a neutral/unknown date
      // (0.35) still lands a flagged-for-review match rather than a miss.
      const base = bothSettle ? 0.2 : 0.1;
      score = Math.max(score, base + 0.55 * amount + 0.25 * date);
    }
  }

  // Hard cap on conflicting identities — keep below the acceptance gate so two
  // different passengers are never paired on coincidental amount/date.
  if (idConflict) score = Math.min(score, 0.35);

  return {
    score: Math.min(1, score),
    evidence: {
      passport,
      reference,
      name: text,
      amount,
      effectiveAmount,
      date,
      method: "rule",
      dateDeltaDays: dayGap(o.date, p.date),
    },
  };
}

/** Acceptance gate so the engine never invents low-evidence matches. */
function acceptRow(s: number, e: MatchEvidence): boolean {
  if (s < 0.5) return false; // below the review band
  const id = Math.max(e.passport, e.reference);
  // Use effectiveAmount (group-corrected) so N-pax group rows aren't blocked.
  return id > 0 || e.effectiveAmount >= 0.6;
}

/** Auto-accept threshold; pairs below this (but matched) are flagged for review. */
const MATCH_THRESHOLD = 0.85;

type Cand = { oi: number; pi: number; score: number; evidence: MatchEvidence };

/**
 * Generate candidate pairs using blocking (only compare rows that share at
 * least one weak signal) then assign greedily by descending score so each row
 * is used at most once. Blocking keeps this scalable for large ledgers.
 */
function matchSet(
  oursRows: LedgerRow[],
  partnerRows: LedgerRow[],
  amountKey: (r: LedgerRow) => number,
  scoreFn: (o: LedgerRow, p: LedgerRow) => ScoreResult,
  acceptFn: (s: number, e: MatchEvidence) => boolean,
): { pairs: Cand[]; usedO: Set<number>; usedP: Set<number> } {
  // Build blocking indexes over partner rows.
  const byPassPrefix = new Map<string, number[]>();
  const byRef = new Map<string, number[]>();
  const byAmt = new Map<number, number[]>();
  const byNameTok = new Map<string, number[]>();
  const push = (m: Map<string, number[]> | Map<number, number[]>, k: any, i: number) => {
    const a = (m as Map<any, number[]>).get(k) ?? [];
    a.push(i);
    (m as Map<any, number[]>).set(k, a);
  };
  // Strong identity keys: full normalised passport + reference, plus regex-extracted refs.
  const keyTokens = (r: LedgerRow): string[] => {
    const out: string[] = [];
    const kp = normRef(r.passport ?? "");
    if (kp.length >= 4) out.push("K" + kp);
    const kr = normRef(r.reference);
    if (kr.length >= 4) out.push("K" + kr);
    extractAdvancedRefs(`${r.description} ${r.reference}`).forEach((x) => out.push("R" + x));
    return out;
  };

  // Settlement rows (bank transfers / payments) carry no per-item ID, so they are
  // all cross-listed under one bucket — guaranteeing every settlement on one side
  // is compared against every settlement on the other, regardless of amount/name.
  const settleP: number[] = [];

  partnerRows.forEach((p, i) => {
    if (p.passport) {
      const cp = normRef(p.passport);
      if (cp.length >= 4) push(byPassPrefix, cp.slice(0, 4), i);
    }
    keyTokens(p).forEach((r) => push(byRef, r, i));
    const amt = Math.round(amountKey(p));
    [amt - 1, amt, amt + 1].forEach((a) => push(byAmt, a, i));
    normName(p.paxName)
      .split(" ")
      .filter((t) => t.length >= 3)
      .slice(0, 3)
      .forEach((t) => push(byNameTok, t, i));
    if (p.settlement) settleP.push(i);
  });

  const cands: Cand[] = [];
  oursRows.forEach((o, oi) => {
    const candIdx = new Set<number>();
    if (o.passport) {
      const cp = normRef(o.passport);
      if (cp.length >= 4) (byPassPrefix.get(cp.slice(0, 4)) ?? []).forEach((i) => candIdx.add(i));
    }
    keyTokens(o).forEach((r) => (byRef.get(r) ?? []).forEach((i) => candIdx.add(i)));
    (byAmt.get(Math.round(amountKey(o))) ?? []).forEach((i) => candIdx.add(i));
    normName(o.paxName)
      .split(" ")
      .filter((t) => t.length >= 3)
      .slice(0, 3)
      .forEach((t) => (byNameTok.get(t) ?? []).forEach((i) => candIdx.add(i)));
    if (o.settlement) settleP.forEach((i) => candIdx.add(i));

    // N-PAX group rows (NST combined entry) have no per-passenger amount.
    // Also look for partner rows whose per-pax amount × N ≈ our group total,
    // covering Maverick multi-pax bookings exploded into individual sub-rows.
    if (o.raw?.isGroupRow) {
      const groupAmt = Math.round(amountKey(o));
      for (let n = 2; n <= 9; n++) {
        const perPax = Math.round(groupAmt / n);
        if (perPax >= 1) {
          [perPax - 1, perPax, perPax + 1].forEach((a) =>
            (byAmt.get(a) ?? []).forEach((i) => candIdx.add(i)),
          );
        }
      }
    }

    const oIsSD = isSecDep(o);
    candIdx.forEach((pi) => {
      // Fast-path security-deposit isolation before calling the full scorer.
      if (oIsSD !== isSecDep(partnerRows[pi])) return;
      const { score, evidence } = scoreFn(o, partnerRows[pi]);
      if (acceptFn(score, evidence)) cands.push({ oi, pi, score, evidence });
    });
  });

  cands.sort((a, b) => b.score - a.score);
  const usedO = new Set<number>();
  const usedP = new Set<number>();
  const pairs: Cand[] = [];
  for (const c of cands) {
    if (usedO.has(c.oi) || usedP.has(c.pi)) continue;
    usedO.add(c.oi);
    usedP.add(c.pi);
    pairs.push(c);
  }
  return { pairs, usedO, usedP };
}

/**
 * Flag duplicate entries WITHIN a single ledger. A duplicate group is ≥2 rows
 * that share the same passport, (near-)identical amount, AND the same visa type.
 * This catches double-bookings (often one of them later refunded via a VR row).
 *
 * Security deposits and visa charges for the SAME passport are deliberately NOT
 * grouped — their visa types differ ("SECURITY DEPOSIT" vs "60 DAYS"), so the
 * machine treats them as the distinct financial items they are.
 */
export function flagDuplicates(rows: LedgerRow[]): void {
  const groups = new Map<string, LedgerRow[]>();
  for (const r of rows) {
    // Only consider genuine per-passenger charge/credit rows.
    if (!r.passport || r.settlement || r.isReversal) continue;
    const amt = r.charge > 0 ? r.charge : r.credit;
    if (amt <= 0) continue;
    const pass = r.passport.toUpperCase().replace(/[^A-Z0-9]/g, "");
    const vt = (r.visaType ?? "").toUpperCase().replace(/\s+/g, " ").trim();
    const key = `${pass}|${Math.round(amt)}|${vt}`;
    const arr = groups.get(key);
    if (arr) arr.push(r);
    else groups.set(key, [r]);
  }
  for (const grp of groups.values()) {
    if (grp.length < 2) continue;
    grp.forEach((r, i) => {
      r.duplicateCount = grp.length;
      r.duplicateIndex = i + 1;
    });
  }
}

export function reconcile(ours: LedgerRow[], partner: LedgerRow[]): ReconResult {
  // Tag duplicate entries inside each ledger before matching, so the UI can
  // surface double-bookings and the reviewer can see them at a glance.
  flagDuplicates(ours);
  flagDuplicates(partner);

  const pairs: Pair[] = [];

  // Single unified pass over ALL rows (charges and payments together) so a debit
  // on one side can pair with the mirrored entry on the other regardless of how
  // each ledger labels it. ID-first via blocking, then weighted fuzzy scoring.
  const m = matchSet(ours, partner, absAmount, scoreRowPair, acceptRow);

  for (const c of m.pairs) {
    const o = ours[c.oi];
    const p = partner[c.pi];
    const ao = absAmount(o);
    const ap = absAmount(p);
    const diff = +(ap - ao).toFixed(2);
    const exact = Math.abs(diff) < 0.5;
    pairs.push({
      key: `m-${o.index}-${p.index}`,
      status: exact ? "matched" : "amount_diff",
      kind: o.kind === "credit" ? "credit" : "charge",
      ours: o,
      partner: p,
      oursAmt: ao,
      partnerAmt: ap,
      diff: exact ? 0 : diff,
      score: c.score,
      confidence: c.score,
      needsReview: c.score < MATCH_THRESHOLD,
      evidence: c.evidence,
      note: explainMatch(c.evidence, diff, exact, c.score, o, p),
    });
  }
  ours.forEach((o, oi) => {
    if (m.usedO.has(oi)) return;
    pairs.push({
      key: `oo-${o.index}`,
      status: "missing_partner",
      kind: o.kind === "credit" ? "credit" : "charge",
      ours: o,
      partner: null,
      oursAmt: absAmount(o),
      partnerAmt: 0,
      diff: -absAmount(o),
      note: "Only in our ledger — no matching row found in partner ledger.",
    });
  });
  partner.forEach((p, pi) => {
    if (m.usedP.has(pi)) return;
    pairs.push({
      key: `op-${p.index}`,
      status: "missing_ours",
      kind: p.kind === "credit" ? "credit" : "charge",
      ours: null,
      partner: p,
      oursAmt: 0,
      partnerAmt: absAmount(p),
      diff: absAmount(p),
      note: "Only in partner ledger — no matching row found in our ledger.",
    });
  });

  return { pairs, totals: computeTotals(ours, partner, pairs) };
}

/** Explainable "why matched" string, e.g. "ID matched exactly · amount 0% · date 1 day". */
function explainMatch(
  e: MatchEvidence,
  diff: number,
  exact: boolean,
  score: number,
  o?: LedgerRow,
  p?: LedgerRow,
): string {
  const parts: string[] = [];
  if (e.passport >= 0.99) parts.push("ID matched exactly");
  else if (e.passport > 0) parts.push("ID similar");
  if (e.reference >= 0.99) parts.push("reference matched");
  if (e.name >= 0.85) parts.push("name matched");
  else if (e.name >= 0.6) parts.push("name similar");
  // Visa type agreement
  const ovt = normVisaType(o?.visaType);
  const pvt = normVisaType(p?.visaType);
  if (ovt && pvt && ovt !== pvt) parts.push(`type mismatch (${ovt} vs ${pvt})`);
  else if (ovt && pvt && ovt === pvt) parts.push(`type matched (${ovt})`);
  parts.push(exact ? "amount difference 0%" : `amount differs by ${Math.abs(diff).toFixed(2)}`);
  if (e.dateDeltaDays !== null)
    parts.push(`date difference ${e.dateDeltaDays} day${e.dateDeltaDays === 1 ? "" : "s"}`);
  if (e.method === "ai") parts.push("AI-assisted");
  parts.push(`confidence ${(score * 100).toFixed(0)}%`);
  return parts.join(" · ");
}

/** Recompute all totals from the final pair set (handles AI-merged pairs too). */
export function computeTotals(
  ours: LedgerRow[],
  partner: LedgerRow[],
  pairs: Pair[],
): ReconResult["totals"] {
  const sum = (arr: LedgerRow[], k: "charge" | "credit") =>
    +arr.reduce((s, r) => s + r[k], 0).toFixed(2);
  const count = (st: PairStatus) => pairs.filter((p) => p.status === st).length;

  const confident = pairs.filter((p) => p.ours && p.partner && typeof p.confidence === "number");
  const avgConfidence = confident.length
    ? +(confident.reduce((s, p) => s + (p.confidence ?? 0), 0) / confident.length).toFixed(3)
    : 0;
  const netAmountDiff = +pairs
    .filter((p) => p.ours && p.partner)
    .reduce((s, p) => s + p.diff, 0)
    .toFixed(2);

  return {
    oursRows: ours.length,
    partnerRows: partner.length,
    oursCharges: sum(ours, "charge"),
    partnerCharges: sum(partner, "charge"),
    oursCredits: sum(ours, "credit"),
    partnerCredits: sum(partner, "credit"),
    matched: count("matched"),
    amountIssues: count("amount_diff"),
    onlyOurs: count("missing_partner"),
    onlyPartner: count("missing_ours"),
    netAmountDiff,
    avgConfidence,
    needsReview: pairs.filter((p) => p.needsReview).length,
    aiAssisted: pairs.filter((p) => p.evidence?.method === "ai").length,
  };
}

export function exportPairsCSV(pairs: Pair[]): string {
  const headers = [
    "Status",
    "Confidence %",
    "Needs Review",
    "Method",
    "Passport",
    "Our Date",
    "Our Pax",
    "Our Ref",
    "Our Charge",
    "Our Credit",
    "Partner Date",
    "Partner Pax",
    "Partner Ref",
    "Partner Charge",
    "Partner Credit",
    "Diff (Partner − Our)",
    "Note",
  ];
  const rows = pairs.map((p) => [
    p.status,
    typeof p.confidence === "number" ? Math.round(p.confidence * 100) : "",
    p.needsReview ? "YES" : "",
    p.evidence?.method === "ai" ? "AI" : p.ours && p.partner ? "RULE" : "",
    p.ours?.passport ?? p.partner?.passport ?? "",
    p.ours?.date ?? "",
    p.ours?.paxName ?? "",
    p.ours?.reference ?? "",
    p.ours?.charge ?? "",
    p.ours?.credit ?? "",
    p.partner?.date ?? "",
    p.partner?.paxName ?? "",
    p.partner?.reference ?? "",
    p.partner?.charge ?? "",
    p.partner?.credit ?? "",
    p.diff,
    p.note,
  ]);
  return Papa.unparse([headers, ...rows]);
}

export const STATUS_LABEL: Record<PairStatus, string> = {
  matched: "Matched",
  amount_diff: "Amount Difference",
  missing_partner: "Only in Our Ledger",
  missing_ours: "Only in Partner Ledger",
};

/* ---- colour palette (ARGB-less RGB hex) for the styled export ---- */
const COL = {
  navy: "0C2E5F",
  gold: "C9A23A",
  white: "FFFFFF",
  headerText: "FFFFFF",
  matchedFill: "D8F3E3",
  matchedText: "047857",
  diffFill: "FEF3C7",
  diffText: "B45309",
  onlyOursFill: "E0E7FF",
  onlyOursText: "4338CA",
  onlyPartnerFill: "FFE4E6",
  onlyPartnerText: "BE123C",
  missingFill: "FCA5A5",
  missingText: "7F1D1D",
  reviewFill: "FEF08A",
  sectionFill: "EAF0F8",
  border: "D6DEE8",
  zebra: "F7F9FC",
};

const statusFill: Record<PairStatus, string> = {
  matched: COL.matchedFill,
  amount_diff: COL.diffFill,
  missing_partner: COL.onlyOursFill,
  missing_ours: COL.onlyPartnerFill,
};
const statusText: Record<PairStatus, string> = {
  matched: COL.matchedText,
  amount_diff: COL.diffText,
  missing_partner: COL.onlyOursText,
  missing_ours: COL.onlyPartnerText,
};

const thinBorder = {
  top: { style: "thin", color: { rgb: COL.border } },
  bottom: { style: "thin", color: { rgb: COL.border } },
  left: { style: "thin", color: { rgb: COL.border } },
  right: { style: "thin", color: { rgb: COL.border } },
};
const solid = (rgb: string) => ({ patternType: "solid", fgColor: { rgb } });

const PAIR_HEADERS = [
  "Status",
  "Confidence %",
  "Needs Review",
  "Method",
  "Category",
  "Scenario",
  "Visa Type",
  "ID / Passport",
  "Our Date",
  "Our Party",
  "Our Reference",
  "Our Charge",
  "Our Credit",
  "Partner Date",
  "Partner Party",
  "Partner Reference",
  "Partner Charge",
  "Partner Credit",
  "Variance (Partner − Our)",
  "Date Gap (days)",
  "Note",
];
const PAIR_COL_WIDTHS = [18, 9, 9, 8, 11, 16, 14, 15, 12, 24, 16, 11, 11, 12, 24, 16, 11, 11, 16, 9, 44];

const SCENARIO_LABEL: Partial<Record<Scenario, string>> = {
  visa_charge: "Visa Charge",
  security_deposit: "Security Deposit",
  wrong_invoice: "Wrong Invoice Refund",
  wrong_client: "Wrong Client Refund",
  duplicate: "Duplicate Refund",
  refund: "Refund / Reversal",
  bank_transfer: "Bank Transfer",
  multi_passenger: "Multi-Passenger",
  flight: "Flight / Airline",
  addon: "Add-On Service",
};

/** One data row for a pair (matches PAIR_HEADERS order). */
function pairRow(p: Pair): (string | number)[] {
  const scenario = p.ours?.scenario ?? p.partner?.scenario;
  const visaType = p.ours?.visaType ?? p.partner?.visaType ?? "";
  return [
    STATUS_LABEL[p.status],
    typeof p.confidence === "number" ? Math.round(p.confidence * 100) : "",
    p.needsReview ? "YES" : "",
    p.evidence?.method === "ai" ? "AI" : p.ours && p.partner ? "RULE" : "",
    p.ours?.settlement || p.partner?.settlement ? "Settlement" : "Charge",
    scenario ? (SCENARIO_LABEL[scenario] ?? scenario) : "",
    visaType,
    p.ours?.passport ?? p.partner?.passport ?? "",
    p.ours?.date ?? "",
    p.ours?.paxName ?? "",
    p.ours?.reference ?? "",
    p.ours?.charge ?? "",
    p.ours?.credit ?? "",
    p.partner?.date ?? "",
    p.partner?.paxName ?? "",
    p.partner?.reference ?? "",
    p.partner?.charge ?? "",
    p.partner?.credit ?? "",
    p.diff,
    p.evidence?.dateDeltaDays ?? "",
    p.note,
  ];
}

/** Build a fully colour-coded worksheet for a set of pairs. */
function styledPairSheet(pairs: Pair[]) {
  const aoa = [PAIR_HEADERS, ...pairs.map(pairRow)];
  const ws = XLSXStyle.utils.aoa_to_sheet(aoa);
  const ncols = PAIR_HEADERS.length;
  ws["!cols"] = PAIR_COL_WIDTHS.map((wch) => ({ wch }));
  ws["!freeze"] = { xSplit: 0, ySplit: 1 };
  // Column indices shifted +2 for the new Scenario and Visa Type columns.
  const rightCols = new Set([1, 11, 12, 16, 17, 18, 19]);
  const ourCols = new Set([8, 9, 10, 11, 12]);
  const partnerCols = new Set([13, 14, 15, 16, 17]);

  // Header row
  for (let c = 0; c < ncols; c++) {
    const ref = XLSXStyle.utils.encode_cell({ r: 0, c });
    ws[ref] = ws[ref] || { t: "s", v: PAIR_HEADERS[c] };
    ws[ref].s = {
      fill: solid(COL.navy),
      font: { bold: true, color: { rgb: COL.headerText }, sz: 10 },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder,
    };
  }

  // Data rows
  pairs.forEach((p, i) => {
    const r = i + 1;
    const ourMissing = !p.ours;
    const partnerMissing = !p.partner;
    for (let c = 0; c < ncols; c++) {
      const ref = XLSXStyle.utils.encode_cell({ r, c });
      ws[ref] = ws[ref] || { t: "s", v: "" };
      const isMissingCell =
        (ourMissing && ourCols.has(c)) || (partnerMissing && partnerCols.has(c));
      // Mark the first column of an absent side with an explicit MISSING flag.
      if (isMissingCell && (c === 8 || c === 13) && !ws[ref].v) {
        ws[ref].v = "⚠ MISSING";
        ws[ref].t = "s";
      }
      const fillRgb = isMissingCell ? COL.missingFill : statusFill[p.status];
      const style: Record<string, unknown> = {
        fill: solid(fillRgb),
        border: thinBorder,
        alignment: {
          horizontal: rightCols.has(c) ? "right" : "left",
          vertical: "center",
        },
      };
      if (isMissingCell) {
        style.font = { bold: true, color: { rgb: COL.missingText }, sz: 9 };
      } else if (c === 0) {
        style.font = { bold: true, color: { rgb: statusText[p.status] }, sz: 10 };
      } else if (c === 2 && p.needsReview) {
        style.fill = solid(COL.reviewFill);
        style.font = { bold: true, color: { rgb: COL.diffText } };
      } else if (c === 18 && Math.abs(p.diff) > 0.5) {
        // Variance column (shifted +2 for Scenario + Visa Type columns)
        style.font = { bold: true, color: { rgb: COL.onlyPartnerText } };
      }
      ws[ref].s = style;
    }
  });

  if (pairs.length) ws["!autofilter"] = { ref: `A1:U${pairs.length + 1}` };
  return ws;
}

/**
 * Build a DETAILED, colour-coded sheet of one full uploaded ledger: every
 * original column is preserved and prefixed with reconciliation columns
 * (row #, match status, confidence, variance, the matched counterparty row).
 * Each row is tinted by its match status; unmatched/blank rows stay neutral.
 */
function styledLedgerSheet(
  aoa: unknown[][] | null | undefined,
  map: Map<number, Pair>,
  side: Side,
) {
  if (!aoa || aoa.length < 1) {
    return XLSXStyle.utils.aoa_to_sheet([["(no data uploaded for this ledger)"]]);
  }
  // True width = widest row anywhere in the sheet, so NO column is ever dropped
  // even when data rows have more cells than the header row.
  let dataWidth = 0;
  for (const row of aoa) dataWidth = Math.max(dataWidth, (row as unknown[])?.length ?? 0);
  const origHeader: string[] = [];
  for (let c = 0; c < dataWidth; c++) {
    const h = String((aoa[0] as unknown[])?.[c] ?? "").trim();
    origHeader.push(h || `Column ${c + 1}`);
  }

  const reconCols = [
    "Row #",
    "Match Status",
    "Confidence %",
    "Variance (Partner − Our)",
    "Matched-To Date",
    "Matched-To Party",
    "Matched-To Amount",
  ];
  const header = [...reconCols, ...origHeader];
  const ncols = header.length;
  const reconN = reconCols.length;

  const aoaOut: (string | number)[][] = [header];
  const rowStatus: (PairStatus | null)[] = [null]; // index aligned with aoaOut
  for (let i = 1; i < aoa.length; i++) {
    const cells = (aoa[i] as unknown[]) ?? [];
    if (!cells.length || cells.every((c) => c === null || c === undefined || c === "")) continue;
    const pair = map.get(i);
    const cp = pair ? (side === "ours" ? pair.partner : pair.ours) : null;
    const cpAmt = cp ? cp.charge || cp.credit : "";
    const origCells: (string | number)[] = [];
    for (let c = 0; c < dataWidth; c++) {
      const v = cells[c];
      origCells.push(v === null || v === undefined ? "" : typeof v === "number" ? v : String(v));
    }
    aoaOut.push([
      i,
      pair ? STATUS_LABEL[pair.status] : "Unmarked",
      pair && typeof pair.confidence === "number" ? Math.round(pair.confidence * 100) : "",
      pair && pair.ours && pair.partner ? pair.diff : "",
      cp?.date ?? "",
      cp?.paxName ?? "",
      cpAmt === 0 ? 0 : cpAmt || "",
      ...origCells,
    ]);
    rowStatus.push(pair ? pair.status : null);
  }

  const ws = XLSXStyle.utils.aoa_to_sheet(aoaOut);
  ws["!cols"] = [
    { wch: 7 },
    { wch: 20 },
    { wch: 11 },
    { wch: 14 },
    { wch: 13 },
    { wch: 22 },
    { wch: 13 },
    ...origHeader.map((h) => ({ wch: Math.min(34, Math.max(11, h.length + 3)) })),
  ];
  ws["!freeze"] = { xSplit: 0, ySplit: 1 };

  // Header styling
  for (let c = 0; c < ncols; c++) {
    const ref = XLSXStyle.utils.encode_cell({ r: 0, c });
    if (!ws[ref]) ws[ref] = { t: "s", v: header[c] };
    ws[ref].s = {
      fill: solid(c < reconN ? COL.gold : COL.navy),
      font: { bold: true, color: { rgb: c < reconN ? "1F2937" : COL.headerText }, sz: 10 },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder,
    };
  }

  // Data rows
  for (let r = 1; r < aoaOut.length; r++) {
    const st = rowStatus[r];
    const fillRgb = st ? statusFill[st] : COL.zebra;
    for (let c = 0; c < ncols; c++) {
      const ref = XLSXStyle.utils.encode_cell({ r, c });
      if (!ws[ref]) ws[ref] = { t: "s", v: "" };
      const style: Record<string, unknown> = {
        fill: solid(fillRgb),
        border: thinBorder,
        alignment: { vertical: "center", horizontal: c === 0 ? "center" : "left" },
      };
      if (c === 1) {
        style.font = {
          bold: true,
          color: { rgb: st ? statusText[st] : "94A3B8" },
          sz: 10,
        };
      } else if (c === 3 && st && (st === "amount_diff" || st === "missing_ours" || st === "missing_partner")) {
        style.font = { bold: true, color: { rgb: COL.onlyPartnerText } };
        style.alignment = { horizontal: "right", vertical: "center" };
      }
      ws[ref].s = style;
    }
  }

  const lastColLetter = XLSXStyle.utils.encode_col(ncols - 1);
  if (aoaOut.length > 1)
    ws["!autofilter"] = { ref: `A1:${lastColLetter}${aoaOut.length}` };
  return ws;
}

/**
 * Build a professional, colour-coded multi-sheet Excel workbook:
 *   1. Summary           – headline KPIs and totals (banded sections)
 *   2. Our Ledger (Full) – your whole uploaded sheet, all columns, marked
 *   3. Partner Ledger (Full) – their whole uploaded sheet, all columns, marked
 *   4. All Items         – every reconciled pair, colour-coded by status
 *   5. Settlements       – bank transfers / payments only
 *   6. Exceptions        – differences + unmatched (the action list)
 * Missing data is highlighted red and flagged "⚠ MISSING".
 */
export function buildReconciliationWorkbook(
  result: ReconResult,
  opts?: { oursAoa?: unknown[][] | null; partnerAoa?: unknown[][] | null },
): ArrayBuffer {
  const { pairs, totals } = result;
  const wb = XLSXStyle.utils.book_new();

  // srcRow → pair lookup for the detailed full-ledger sheets.
  const oursMap = new Map<number, Pair>();
  const partnerMap = new Map<number, Pair>();
  pairs.forEach((p) => {
    if (p.ours?.srcRow != null) oursMap.set(p.ours.srcRow, p);
    if (p.partner?.srcRow != null) partnerMap.set(p.partner.srcRow, p);
  });

  const settlementPairs = pairs.filter((p) => p.ours?.settlement || p.partner?.settlement);
  const exceptionPairs = pairs.filter((p) => p.status !== "matched");
  const settledValue = +pairs
    .filter((p) => p.status === "matched")
    .reduce((s, p) => s + p.partnerAmt, 0)
    .toFixed(2);
  const totalPaired = pairs.filter((p) => p.ours && p.partner).length;

  /* ---- Sheet 1: Summary (banded, coloured) ---- */
  type Row = { cells: (string | number)[]; kind: "title" | "section" | "kv" | "blank" };
  const S = (cells: (string | number)[], kind: Row["kind"]): Row => ({ cells, kind });
  const rows: Row[] = [
    S(["NAVVI SAADI — AI LEDGER RECONCILIATION REPORT"], "title"),
    S(["Generated", new Date().toLocaleString()], "kv"),
    S([], "blank"),
    S(["OVERVIEW"], "section"),
    S(["Our ledger rows", totals.oursRows], "kv"),
    S(["Partner ledger rows", totals.partnerRows], "kv"),
    S(["Total reconciled items", pairs.length], "kv"),
    S(["Rows paired", totalPaired], "kv"),
    S(["Match rate %", pairs.length ? Math.round((totalPaired / pairs.length) * 100) : 0], "kv"),
    S(["Average confidence %", Math.round(totals.avgConfidence * 100)], "kv"),
    S([], "blank"),
    S(["RESULTS"], "section"),
    S(["Matched (amounts agree)", totals.matched], "kv"),
    S(["Amount differences", totals.amountIssues], "kv"),
    S(["Only in our ledger", totals.onlyOurs], "kv"),
    S(["Only in partner ledger", totals.onlyPartner], "kv"),
    S(["Flagged for review", totals.needsReview], "kv"),
    S(["AI-assisted matches", totals.aiAssisted], "kv"),
    S([], "blank"),
    S(["FINANCIALS"], "section"),
    S(["Our total charges", totals.oursCharges], "kv"),
    S(["Partner total charges", totals.partnerCharges], "kv"),
    S(["Our total credits", totals.oursCredits], "kv"),
    S(["Partner total credits", totals.partnerCredits], "kv"),
    S(["Matched value (verified)", settledValue], "kv"),
    S(["Net amount difference", totals.netAmountDiff], "kv"),
    S([], "blank"),
    S(["SETTLEMENTS / BANK TRANSFERS"], "section"),
    S(["Settlement items", settlementPairs.length], "kv"),
    S(["Settlements matched", settlementPairs.filter((p) => p.status === "matched").length], "kv"),
    S(
      ["Settlements unmatched", settlementPairs.filter((p) => p.status !== "matched").length],
      "kv",
    ),
  ];
  const wsSummary = XLSXStyle.utils.aoa_to_sheet(rows.map((r) => r.cells));
  wsSummary["!cols"] = [{ wch: 32 }, { wch: 26 }];
  wsSummary["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
  rows.forEach((row, r) => {
    const a = XLSXStyle.utils.encode_cell({ r, c: 0 });
    const b = XLSXStyle.utils.encode_cell({ r, c: 1 });
    if (row.kind === "title") {
      if (wsSummary[a])
        wsSummary[a].s = {
          fill: solid(COL.navy),
          font: { bold: true, sz: 14, color: { rgb: COL.gold } },
          alignment: { horizontal: "center", vertical: "center" },
        };
    } else if (row.kind === "section") {
      if (wsSummary[a])
        wsSummary[a].s = {
          fill: solid(COL.sectionFill),
          font: { bold: true, sz: 11, color: { rgb: COL.navy } },
        };
    } else if (row.kind === "kv") {
      if (wsSummary[a])
        wsSummary[a].s = { font: { bold: true, color: { rgb: "334155" } } };
      if (wsSummary[b])
        wsSummary[b].s = {
          font: { color: { rgb: "0F172A" } },
          alignment: { horizontal: "right" },
        };
    }
  });
  XLSXStyle.utils.book_append_sheet(wb, wsSummary, "Summary");

  /* ---- Detailed full-ledger sheets (both uploaded files, every column) ---- */
  if (opts?.oursAoa)
    XLSXStyle.utils.book_append_sheet(
      wb,
      styledLedgerSheet(opts.oursAoa, oursMap, "ours"),
      "Our Ledger (Full)",
    );
  if (opts?.partnerAoa)
    XLSXStyle.utils.book_append_sheet(
      wb,
      styledLedgerSheet(opts.partnerAoa, partnerMap, "partner"),
      "Partner Ledger (Full)",
    );

  /* ---- Coloured pair tables ---- */
  XLSXStyle.utils.book_append_sheet(wb, styledPairSheet(pairs), "All Items");
  XLSXStyle.utils.book_append_sheet(wb, styledPairSheet(settlementPairs), "Settlements");
  XLSXStyle.utils.book_append_sheet(wb, styledPairSheet(exceptionPairs), "Exceptions");

  return XLSXStyle.write(wb, { bookType: "xlsx", type: "array" }) as ArrayBuffer;
}
