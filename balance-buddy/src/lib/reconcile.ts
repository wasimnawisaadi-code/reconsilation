import * as XLSX from "xlsx";
import * as XLSXStyle from "xlsx-js-style";
import Papa from "papaparse";

export type Side = "ours" | "partner";

/**
 * Cheap structural integrity check for an uploaded spreadsheet.
 *
 * XLSX.read will thrash for a very long time (effectively hanging the browser
 * tab) on a `.xlsx` whose internal ZIP container has been corrupted. A common
 * cause is re-saving or converting a binary Excel file through a text/UTF-8
 * editor: every non-UTF-8 byte becomes the 3-byte replacement character, which
 * both inflates the file and destroys the ZIP central directory. We detect that
 * here and fail fast with a clear, user-actionable message instead of freezing.
 *
 * Only ZIP-based formats (.xlsx / .xlsm — they start with the "PK" signature)
 * are checked; legacy `.xls` and CSV files are left to their own parsers.
 */
/** "%PDF" magic bytes — the upload is a PDF statement, not a spreadsheet. */
export function looksLikePdf(buf: ArrayBuffer): boolean {
  const u8 = new Uint8Array(buf);
  return u8.length > 4 && u8[0] === 0x25 && u8[1] === 0x50 && u8[2] === 0x44 && u8[3] === 0x46;
}

export function assertReadableSpreadsheet(buf: ArrayBuffer, fileName = "This file"): void {
  const u8 = new Uint8Array(buf);
  if (u8.length < 22) throw new Error(`"${fileName}" is empty or unreadable.`);
  if (!(u8[0] === 0x50 && u8[1] === 0x4b)) return; // not a ZIP container

  const corrupt = () =>
    new Error(
      `"${fileName}" appears to be corrupted — its internal structure is damaged, ` +
        `so it can't be opened. This usually happens when an Excel file is re-saved ` +
        `or converted through a text editor. Please re-export the original file from ` +
        `your accounting software and upload it again.`,
    );

  // Locate the End-Of-Central-Directory record (signature PK\x05\x06), scanning
  // back from the end past any trailing comment (max 65535 bytes).
  const minStart = Math.max(0, u8.length - (22 + 65535));
  let eocd = -1;
  for (let i = u8.length - 22; i >= minStart; i--) {
    if (u8[i] === 0x50 && u8[i + 1] === 0x4b && u8[i + 2] === 0x05 && u8[i + 3] === 0x06) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw corrupt();

  // The EOCD points at the central directory; verify a real CD header lives there.
  const cdOffset =
    (u8[eocd + 16] | (u8[eocd + 17] << 8) | (u8[eocd + 18] << 16) | (u8[eocd + 19] << 24)) >>> 0;
  if (cdOffset + 4 > u8.length) throw corrupt();
  if (
    !(
      u8[cdOffset] === 0x50 &&
      u8[cdOffset + 1] === 0x4b &&
      u8[cdOffset + 2] === 0x01 &&
      u8[cdOffset + 3] === 0x02
    )
  )
    throw corrupt();
}

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
   * ISO month label "YYYY-MM" — set when a file comes from a multi-month upload
   * so the UI can filter/group reconciliation results by month.
   */
  month?: string;
  /**
   * ISO currency code (e.g. "SAR", "AED", "USD") detected for this row. Read from
   * the supplier statement's "Currency" header on the partner side, and inferred
   * from the flight route (Saudi airports → SAR, UAE airports → AED) on the GDS
   * side. Lets the UI label amounts correctly and tell a genuine same-currency
   * over-charge apart from an expected cross-currency / retail-vs-wholesale gap.
   */
  currency?: string;
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

/** Map Arabic-Indic (٠١٢٣٤٥٦٧٨٩) and Extended-Arabic (۰…۹) digits to ASCII. */
const normalizeDigits = (s: string): string =>
  s
    .replace(/[٠-٩]/g, (d) => String(d.charCodeAt(0) - 0x0660))
    .replace(/[۰-۹]/g, (d) => String(d.charCodeAt(0) - 0x06f0));

const num = (v: unknown): number => {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  // Unicode minus / dashes some exports use instead of ASCII "-".
  let s = normalizeDigits(String(v).trim()).replace(/[−–—]/g, "-");
  if (!s) return 0;

  let neg = false;
  // Accounting negative: "(1,500.00)" → -1500
  const paren = /^\((.+)\)$/.exec(s);
  if (paren) {
    neg = true;
    s = paren[1].trim();
  }

  // Strip any prefix junk up to the first digit — currency codes ("SAR 1,500",
  // "Rs.1500"), symbols ($ € £ ﷼ د.إ …) — keeping a minus that directly
  // precedes the number. Pure numeric prefixes (".5", "-3") are left alone.
  const fd = s.search(/\d/);
  if (fd < 0) return 0;
  if (fd > 0) {
    const prefix = s.slice(0, fd);
    if (!/^[.,\-+\s]*$/.test(prefix)) {
      if (/-\s*$/.test(prefix)) neg = true;
      s = s.slice(fd);
    }
  }

  // Trailing-minus negatives ("1500-", common in ERP exports)…
  if (/-\s*$/.test(s)) {
    neg = true;
    s = s.replace(/-\s*$/, "");
  }
  // …then trailing junk: alpha codes ("465.00 INV", "1,500.00 DR"), %, symbols.
  s = s.replace(/[^\d.,]+$/, "").trim();
  // Space / NBSP / thin-space / apostrophe thousands ("1 234,56", "1'234.56")
  s = s.replace(/[\s  ']/g, "");
  if (!s) return 0;

  const hasComma = s.includes(",");
  const hasDot   = s.includes(".");

  if (hasComma && !hasDot) {
    // European decimal: "255,0" or "1.234,56" (no dot present)
    // If comma is followed by 1-2 digits at the very end → decimal separator
    if (/,\d{1,2}$/.test(s)) {
      s = s.replace(/\./g, "").replace(",", ".");
    } else {
      // Thousands separator: "1,234" → "1234"
      s = s.replace(/,/g, "");
    }
  } else if (hasComma && hasDot) {
    // Both present — determine which is decimal by position
    if (s.lastIndexOf(",") > s.lastIndexOf(".")) {
      // "1.234,56" → European
      s = s.replace(/\./g, "").replace(",", ".");
    } else {
      // "1,234.56" → US
      s = s.replace(/,/g, "");
    }
  }
  // else only dot → standard decimal, parse directly

  const n = parseFloat(s);
  if (isNaN(n)) return 0;
  return neg && n > 0 ? -n : n;
};

/* ------------------------------------------------------------------ */
/* CURRENCY DETECTION                                                   */
/* ------------------------------------------------------------------ */

/** IATA airport codes for the Gulf selling markets we operate in. */
const SAUDI_AIRPORTS =
  /\b(JED|RUH|DMM|MED|AHB|TIF|ELQ|GIZ|TUU|AJF|HAS|EAM|YNB|ULH|HOF|SHW|EJH|WAE|AQI|BHH|DWD|RAE|NUM|EJH)\b/;
const UAE_AIRPORTS = /\b(DXB|AUH|SHJ|RKT|FJR|DWC|AAN)\b/;

/** Explicit currency code appearing in free text ("... 465.00 SAR").
 *  Only unambiguous ISO codes — word-like codes (TRY, MAD, GEL, TOP, CUP, ALL)
 *  are deliberately excluded to avoid false hits inside normal narration. */
const CURRENCY_CODE_RE =
  /\b(SAR|AED|USD|QAR|KWD|BHD|OMR|EUR|GBP|PKR|INR|AFN|JOD|EGP|IQD|LBP|JPY|CNY|MYR|IDR|THB|SGD|HKD|KRW|CAD|AUD|NZD|CHF|ZAR|BDT|LKR|NPR|KES|NGN|GHS|ETB|TZS|UGX|DKK|SEK|NOK|RUB|UZS|AZN|KZT)\b/;

/**
 * Infer the ISO currency of a booking from any text that may carry a route or an
 * explicit currency code. An explicit code wins; otherwise the Gulf selling
 * market is read from the flight sector (Saudi airports → SAR, UAE → AED).
 * Returns undefined when nothing recognisable is present (caller leaves it blank).
 */
export function detectCurrency(...parts: (string | null | undefined)[]): string | undefined {
  const t = parts.filter(Boolean).join(" ").toUpperCase();
  if (!t) return undefined;
  const code = t.match(CURRENCY_CODE_RE);
  if (code) return code[1];
  if (UAE_AIRPORTS.test(t)) return "AED";
  if (SAUDI_AIRPORTS.test(t)) return "SAR";
  return undefined;
}

/**
 * Read the account currency declared in a supplier-statement header block, e.g.
 * a row like ["Jeddah Saudi Arabia", "Currency", "SAR"]. Scans the first rows for
 * a "Currency" label and returns the adjacent code, else the first bare currency
 * code found in the header. Returns undefined if none is declared.
 */
export function detectStatementCurrency(aoa: unknown[][]): string | undefined {
  for (let r = 0; r < Math.min(20, aoa.length); r++) {
    const row = (aoa[r] as unknown[]) ?? [];
    for (let c = 0; c < row.length; c++) {
      if (/^\s*currency\s*$/i.test(String(row[c] ?? ""))) {
        for (let k = c + 1; k < row.length; k++) {
          const v = String(row[k] ?? "").trim().toUpperCase();
          const m = v.match(CURRENCY_CODE_RE);
          if (m) return m[1];
        }
      }
    }
  }
  // Fallback: any standalone currency code in the header block.
  for (let r = 0; r < Math.min(20, aoa.length); r++) {
    for (const cell of (aoa[r] as unknown[]) ?? []) {
      const v = String(cell ?? "").trim();
      if (/^(SAR|AED|USD|QAR|KWD|BHD|OMR)$/i.test(v)) return v.toUpperCase();
    }
  }
  return undefined;
}

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
  // A "security deposit" is a visa-charge instrument, NOT a bank transfer — even
  // though it contains the word "deposit". Never classify it as a settlement
  // unless the row ALSO carries an explicit bank / account-transfer marker
  // (e.g. "AC-4 Rak Bank", "wire", "NEFT"). This stops security deposits from
  // being swept into the Payments view and losing their per-passenger identity.
  if (
    SECURITY_DEPOSIT_RE.test(text) &&
    !BANK_ACCOUNT_RE.test(text) &&
    !/\b(transfer|wire|remit|neft|rtgs|imps|swift|telegraphic)\b/i.test(text)
  ) {
    return false;
  }
  return SETTLEMENT_RE.test(text) || BANK_ACCOUNT_RE.test(text);
}

/**
 * Matches a security / guarantee deposit line. Real ledgers misspell "security"
 * constantly ("SECUTITY", "SECUTIRY", "SECRUITY", "SECURTIY"…), so rather than
 * enumerate typos we treat ANY "sec…" word immediately followed by "deposit" as a
 * security deposit. We also accept "guarantee deposit" and the correctly-spelled
 * "security" standalone. This keeps mis-spelled deposits out of the bank-transfer
 * bucket (the bare word "deposit" would otherwise look like a settlement).
 */
const SECURITY_DEPOSIT_RE =
  /\b(?:sec\w*|guarantee)\s*deposit\b|\bsecurity\b(?!\s*(?:code|number|check|scan))/i;

/** Matches common visa service duration types in description fields. */
const VISA_TYPE_RE =
  /\b(\d{1,3})\s*days?\b|\b(1M|2M|3M|6M)\s*ext(?:ension)?\b|\bmulti\b|\bsingle\b|\b(?:sec\w*|guarantee)\s*deposit\b/i;

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
  /** See the identical field on ai-service.ts's ColumnMapping (this is a
   *  structurally-duplicated type so the engine stays AI-code-free) — when
   *  charge/credit map to the same unsigned column, disambiguates a positive
   *  value's direction. Defaults to "credit" when absent. */
  unsignedAmountMeans?: "charge" | "credit";
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
  const refs = [...new Set(found)].filter((s) => {
    if (/^\d{1,4}$/.test(s)) return false; // Too short to be a unique ref
    if (/^\d{8}$/.test(s)) return false; // Likely a date (YYYYMMDD)
    return true;
  });

  // Some supplier systems append a 1-digit invoice-line sequence number onto the
  // 6-character PNR (e.g. GDS PNR "18GR56" is stored as "18GR561" in their
  // software entry report) rather than a distinct booking key. Also emit the
  // 6-char prefix as a candidate ref so it can still line up with the bare PNR
  // on the other ledger, without discarding the original 7-char token.
  const extra: string[] = [];
  for (const s of refs) {
    if (/^[A-Z0-9]{6}\d$/.test(s)) extra.push(s.slice(0, 6));
  }
  // Letter "O" and digit "0" are visually identical and get transcribed
  // inconsistently between systems for the same 6-7 char PNR-shaped code
  // (e.g. GDS "18PG0N" vs supplier "18PGON" for the same booking). Emit an
  // O-folded-to-0 variant so both ledgers' tokens line up, scoped to
  // PNR-length codes only so it can't fold an unrelated longer reference.
  for (const s of [...refs, ...extra]) {
    if (/^[A-Z0-9]{6,7}$/.test(s) && s.includes("O")) extra.push(s.replace(/O/g, "0"));
  }
  return extra.length ? [...new Set([...refs, ...extra])] : refs;
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

    // Group-booking row that bundles N passengers but lists only the LEAD passport
    // (our NST "04 PAX 60 DAYS VISA" style). The other passports are unrecoverable,
    // but the head-count IS known, so split the amount into N equal per-passenger
    // sub-rows. The first keeps the lead passport (matches the supplier's lead line
    // on ID); the rest carry no passport and reconcile on amount + date + visa type
    // against the supplier's remaining per-passenger lines.
    // Different suppliers phrase the headcount differently — "04 PAX", "PAX: 4",
    // "PAX-04", "4 PASSENGERS", "4 PERSONS", "GROUP OF 4" — try each in turn so a
    // new company's ledger format isn't silently dropped to the ungrouped path.
    const groupText = `${row.paxName} ${row.description}`;
    const groupM =
      groupText.match(/(?:^|\b)0*(\d{1,2})\s*(?:PAX|PASSENGERS?|PERSONS?|PEOPLE)\b/i) ??
      groupText.match(/\bPAX(?:ES)?[\s:#-]*0*(\d{1,2})\b/i) ??
      groupText.match(/\bGROUP\s*(?:OF)?[\s:#-]*0*(\d{1,2})\b/i);
    if (!row.settlement && passports.length < 2 && amount > 0 && groupM) {
      const gN = parseInt(groupM[1], 10);
      if (gN >= 2 && gN <= 50) {
        const per = Math.round((amount / gN) * 100) / 100;
        const isCharge = row.charge > 0;
        for (let i = 0; i < gN; i++) {
          const amt = i === 0 ? +(amount - per * (gN - 1)).toFixed(2) : per;
          out.push({
            ...row,
            passport: i === 0 ? row.passport : null,
            charge: isCharge ? amt : 0,
            credit: isCharge ? 0 : amt,
            raw: {
              ...row.raw,
              explodedFrom: row.reference || row.paxName,
              paxIndex: i + 1,
              paxCount: gN,
              explodedGroupAmt: amount,
              isGroupRow: true,
            },
          });
        }
        continue;
      }
    }

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
        raw: {
          ...row.raw,
          explodedFrom: row.reference || row.paxName,
          paxIndex: i + 1,
          paxCount: n,
          explodedGroupAmt: amount,
          isGroupRow: true,
        },
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
  assertReadableSpreadsheet(buf, file.name);
  // PDF statements (text-based) are converted to a table first.
  let aoa: unknown[][] | null = null;
  if (looksLikePdf(buf)) {
    const { pdfBufToAoa } = await import("./pdf-ledger");
    aoa = await pdfBufToAoa(buf, file.name);
  }
  // Try XLSX
  if (!aoa) {
    try {
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      aoa = readBestSheetAoa(wb, { raw: true });
    } catch {
      aoa = null;
    }
  }
  if (!aoa || aoa.length === 0) {
    // Fallback: HTML-table ".xls" or delimiter-auto-detected text.
    aoa = textToAoa(buf) as string[][];
  }
  if (!aoa.length) return [];
  let rows: LedgerRow[];
  if (isSoftwareEntryReport(aoa)) {
    rows = parseSoftwareEntryReport(aoa);
  } else if (looksLikeGdsMonthly(aoa)) {
    rows = parseGDSMonthlyLedger(aoa, monthFromFilename(file.name));
  } else {
    // The real header isn't always aoa[0] — a company-title row above it (like
    // AL Saffron's "AL SAFFRON TRAVEL AND TOURISM LLC — ...") would otherwise
    // make hasCode/hasDRCR both false and silently misroute a genuine DR/CR
    // layout into the generic fallback, losing the DocNo-prefix scenario
    // classification (VS/VR/IS/JV) that parseOurDrCrStyle applies.
    const hRow = detectHeaderRow(aoa);
    const headerRow = (aoa[hRow] as unknown[]).map((c) =>
      String(c ?? "")
        .trim()
        .toUpperCase(),
    );
    const hasCode = headerRow.includes("CODE") && headerRow.includes("AMOUNT");
    const hasDRCR = headerRow.includes("DR") && headerRow.includes("CR");

    if (hasCode) rows = parseOurNarrationStyle(aoa.slice(hRow));
    else if (hasDRCR) rows = parseOurDrCrStyle(aoa.slice(hRow));
    // Unknown layout → generic auto-detecting parser (works for any ledger).
    else rows = parseGenericLedger(aoa, "ours");
    if (hRow > 0) for (const r of rows) if (typeof r.srcRow === "number") r.srcRow += hRow;
  }

  // Safety net: any format-specific parser that produced nothing (a mis-detected
  // layout) falls back to the universal generic parser, so an unexpected file
  // still reconciles instead of coming back empty.
  if (!rows.length) rows = parseGenericLedger(aoa, "ours");
  return explodeMultiPax(rows);
}

/* ------------------------------------------------------------------ */
/* GENERIC AUTO-DETECTING PARSER  (works for ANY ledger layout)        */
/* ------------------------------------------------------------------ */

// Keyword groups used to recognise columns in an unknown ledger. Trailing word
// boundaries are intentionally omitted so plurals ("Particulars", "Comments",
// "Charges") still match. Debit/Credit are anchored to the start to avoid false
// hits inside unrelated words.
// English keywords are matched on the accent-stripped header (so French
// "Débit"/"Crédit" hit debit/credit); Arabic terms are matched as literal
// substrings (JS \b does not work with Arabic script).
const COLREGEX = {
  date: /\b(date|dated|posting|value\s*dt|record\s*time|\btime\b|\bday\b)|تاريخ|التاريخ/i,
  debit: /^\s*(dr\.?\b|debit|withdraw|paid\s*out|out\s*flow|money\s*out|payment\s*out|charge[sd]?\b|expense)|مدين/i,
  credit: /^\s*(cr\.?\b|credit|deposit|paid\s*in|in\s*flow|money\s*in|payment\s*in|received)|دائن/i,
  amount: /\b(amount|amt|value|net\b|total|sum)|المبلغ|مبلغ|القيمة/i,
  balance: /\b(balance|bal\b|running|closing)|الرصيد|رصيد/i,
  reference: /\b(ref|voucher|invoice|inv\b|bill|cheque|chq|transaction|txn|utr|document|doc\s*no|pnr|ticket|receipt|booking|order\s*no|folio)|مرجع|سند|فاتورة|مستند/i,
  name: /\b(name|party|particular|customer|supplier|vendor|pax|beneficiary|payee|payer|client|narration|description|detail|comment|memo|remark)|البيان|بيان|الاسم|اسم|العميل|الوصف|المستفيد/i,
  id: /\b(passport|national\s*id|nat\s*id|id\s*no|emirates\s*id|iqama|trn|cnic|civil\s*id|qid|cpr\b)|جواز|الهوية|هوية|إقامة|اقامة/i,
};

// A layout can carry BOTH a real person-identity column (e.g. "Pax Name") and a
// generic free-text column (e.g. "Narration") — when the free-text one appears
// first, the combined `COLREGEX.name` match would pick it even though it's
// blank/irrelevant boilerplate on many real exports. Person-identity keywords
// are tried first; the generic ones are only a fallback when no identity
// column exists at all (so narration-only layouts still work as before).
const NAME_STRONG = /\b(name|party|particular|customer|supplier|vendor|pax|beneficiary|payee|payer|client)\b|البيان|الاسم|اسم|العميل|المستفيد/i;

/** Strip combining accents so "Débit" matches the "debit" keyword. */
const stripAccents = (s: string): string =>
  s.normalize("NFD").replace(/[̀-ͯ]/g, "").normalize("NFC");

/**
 * True when a data row is a non-transaction summary line — "TOTAL",
 * "Grand Total", "Opening/Closing Balance", "Balance B/F", "Carried Forward".
 * Anchored to the start of a cell so a real narration merely containing the
 * word "total" mid-sentence is never skipped.
 */
function isSummaryRow(cells: unknown[]): boolean {
  for (const c of cells) {
    const s = String(c ?? "").trim();
    if (!s || s.length > 40) continue;
    if (
      /^((sub|grand|net|page|running)\s*-?\s*totals?\b|totals?\s*$|totals?\s*:|(opening|closing)\s*balance|balance\s*(b\/?f|c\/?f|brought|carried)|brought\s*forward|carried\s*forward)/i.test(
        s,
      )
    ) {
      return true;
    }
  }
  return false;
}

/**
 * Normalize a date cell to something displayable and parseable: Date instances
 * and Excel serial numbers (e.g. 45800) become ISO "YYYY-MM-DD"; everything
 * else is returned as its trimmed string.
 */
function normalizeDateCell(v: unknown): string {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  const s = String(v).trim();
  if (/^\d{4,6}(\.\d+)?$/.test(s)) {
    const t = parseDate(s);
    if (!isNaN(t)) {
      const d = new Date(t);
      const p = (n: number) => String(n).padStart(2, "0");
      return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
    }
  }
  return s;
}

/**
 * Read the workbook sheet that actually holds the ledger. Real exports often
 * put a cover / notes / parameters sheet first, so SheetNames[0] is not always
 * the data — pick the sheet (of the first 8) with the most rows that carry at
 * least two non-empty cells. Ties keep the earliest sheet, preserving the old
 * behaviour for normal single-sheet files.
 */
export function readBestSheetAoa(
  wb: XLSX.WorkBook,
  opts: XLSX.Sheet2JSONOpts = {},
): unknown[][] {
  let best: unknown[][] = [];
  let bestScore = -1;
  for (const name of wb.SheetNames.slice(0, 8)) {
    const ws = wb.Sheets[name];
    if (!ws) continue;
    const aoa = XLSX.utils.sheet_to_json<unknown[]>(ws, { ...opts, header: 1 });
    const score = aoa.reduce(
      (s, r) => s + ((r as unknown[]).filter((c) => c !== "" && c != null).length >= 2 ? 1 : 0),
      0,
    );
    if (score > bestScore) {
      bestScore = score;
      best = aoa;
    }
  }
  return best;
}

/**
 * Text fallback for uploads that are not real binary spreadsheets:
 * - HTML tables saved with an ".xls" name (very common in old ERP/bank
 *   exports) are parsed through SheetJS's HTML reader;
 * - everything else goes to Papa with the delimiter AUTO-detected, so
 *   semicolon (European), pipe and tab separated files all work.
 */
export function textToAoa(buf: ArrayBuffer): unknown[][] {
  const text = new TextDecoder("utf-8").decode(buf).replace(/^﻿/, "");
  if (/^\s*</.test(text) && /<table/i.test(text)) {
    try {
      const wb = XLSX.read(text, { type: "string" });
      const aoa = readBestSheetAoa(wb, { raw: true });
      if (aoa.length) return aoa;
    } catch {
      /* fall through to delimited-text parse */
    }
  }
  return Papa.parse<unknown[]>(text, { skipEmptyLines: true }).data;
}

/** Score how header-like a row is (count of cells matching known column words). */
function headerScore(row: unknown[]): number {
  let score = 0;
  for (const cell of row) {
    const s = stripAccents(String(cell ?? "").trim());
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

  // Two-row headers ("Amount" split into Debit/Credit on the line below): when
  // the NEXT row also reads like a header — keyword hits, no numbers, no dates —
  // merge its labels into the header and start the data one row later.
  let dataStart = hRow + 1;
  const nextRow = (aoa[hRow + 1] as unknown[]) ?? [];
  if (nextRow.length && headerScore(nextRow) >= 2) {
    const looksData = nextRow.some((c) => {
      const s = String(c ?? "").trim();
      if (!s) return false;
      return Math.abs(num(s)) > 0 || !isNaN(parseDate(s));
    });
    if (!looksData) {
      for (let i = 0; i < nextRow.length; i++) {
        const sub = String(nextRow[i] ?? "").trim();
        // Sub-label first: debit/credit keywords are start-anchored, so
        // "Amount"+"Debit" must merge to "Debit Amount", not "Amount Debit".
        if (sub) header[i] = header[i] ? `${sub} ${header[i]}` : sub;
      }
      dataStart = hRow + 2;
    }
  }

  // Accent-stripped copy for keyword matching ("Débit" → "Debit").
  const headerNorm = header.map(stripAccents);

  const find = (re: RegExp, exclude?: RegExp): number => {
    for (let i = 0; i < headerNorm.length; i++) {
      const h = headerNorm[i];
      if (!h) continue;
      if (exclude && exclude.test(h)) continue;
      if (re.test(h)) return i;
    }
    return -1;
  };

  // Print-style exports repeat the header block on every "page" of the sheet —
  // recognise a data row that is just the header again so it can be skipped.
  const isRepeatedHeader = (row: unknown[]): boolean => {
    let same = 0;
    let nonEmpty = 0;
    for (let i = 0; i < Math.max(row.length, header.length); i++) {
      const s = String(row[i] ?? "").trim();
      if (!s) continue;
      nonEmpty++;
      if (header[i] && s.toUpperCase() === header[i].toUpperCase()) same++;
    }
    return nonEmpty >= 2 && same / nonEmpty >= 0.6;
  };

  const idxDate = find(COLREGEX.date);
  const idxDebit = find(COLREGEX.debit);
  const idxCredit = find(COLREGEX.credit);
  // Only treat a generic "amount" column as signed when there is no DR/CR split.
  const idxAmount = idxDebit >= 0 && idxCredit >= 0 ? -1 : find(COLREGEX.amount, COLREGEX.balance);
  // Exclude date columns from reference detection so "Transaction Date" is read
  // as the date, not the reference (the word "transaction" matches both). A
  // layout often has SEVERAL reference-like columns (DocNo, Ticket No., PNR
  // No., Reference…) — collect all of them so none of that identity/matching
  // data is silently dropped just because one column matched first.
  const idxRefAll: number[] = [];
  for (let i = 0; i < headerNorm.length; i++) {
    const h = headerNorm[i];
    if (!h || COLREGEX.date.test(h)) continue;
    if (COLREGEX.reference.test(h)) idxRefAll.push(i);
  }
  const idxRef = idxRefAll[0] ?? -1;
  const idxId = find(COLREGEX.id);
  const idxName = (() => {
    const strong = find(NAME_STRONG);
    return strong >= 0 ? strong : find(COLREGEX.name);
  })();

  // Single-amount ledgers often carry the direction in a separate DR/CR marker
  // column ("1,500.00" + "DR"). Find it by VALUES, not header (a "Type" header
  // can mean anything): a column whose non-empty cells are ≥80% DR/CR markers.
  let idxDrCr = -1;
  if (idxAmount >= 0) {
    const scanEnd = Math.min(aoa.length, dataStart + 40);
    for (let c = 0; c < header.length && idxDrCr < 0; c++) {
      if (c === idxAmount || c === idxDate || idxRefAll.includes(c) || c === idxId || c === idxName) continue;
      let hits = 0;
      let nonEmpty = 0;
      for (let r = dataStart; r < scanEnd; r++) {
        const s = String((aoa[r] as unknown[] | undefined)?.[c] ?? "").trim().toUpperCase();
        if (!s) continue;
        nonEmpty++;
        if (/^(DR|CR|D|C|DEBIT|CREDIT)\.?$/.test(s)) hits++;
      }
      if (nonEmpty >= 3 && hits / nonEmpty >= 0.8) idxDrCr = c;
    }
  }

  const rows: LedgerRow[] = [];
  for (let r = dataStart; r < aoa.length; r++) {
    const row = (aoa[r] as unknown[]) ?? [];
    if (!row.length || row.every((c) => c === null || c === undefined || c === "")) continue;
    // Print-style page headers repeated mid-data are not transactions.
    if (isRepeatedHeader(row)) continue;

    const idRaw = idxId >= 0 ? String(row[idxId] ?? "") : "";
    // Skip totals / opening-closing balance / carried-forward summary lines —
    // they are not transactions and would otherwise pollute matching as fake
    // "missing" entries. Rows that carry a real ID are never skipped.
    if (!idRaw.trim() && isSummaryRow(row)) continue;

    const dateStr = idxDate >= 0 ? normalizeDateCell(row[idxDate]) : "";

    const name = idxName >= 0 ? String(row[idxName] ?? "") : "";
    // Concatenate every reference-like column (DocNo, Ticket No., PNR No., …)
    // instead of only the first match, so downstream regex-based matching and
    // passport recovery never lose data that lives in a column other than
    // whichever reference column happened to be leftmost.
    const ref = idxRefAll.map((i) => String(row[i] ?? "").trim()).filter(Boolean).join(" ");

    let charge = 0;
    let credit = 0;
    if (idxAmount >= 0) {
      const v = num(row[idxAmount]);
      const ind = idxDrCr >= 0 ? String(row[idxDrCr] ?? "").trim().toUpperCase() : "";
      if (/^(DR|D|DEBIT)\.?$/.test(ind)) charge = Math.abs(v);
      else if (/^(CR|C|CREDIT)\.?$/.test(ind)) credit = Math.abs(v);
      else if (v < 0) charge = Math.abs(v);
      else if (v > 0) credit = v;
    } else {
      const dr = idxDebit >= 0 ? num(row[idxDebit]) : 0;
      const cr = idxCredit >= 0 ? num(row[idxCredit]) : 0;
      if (dr > 0) charge = dr;
      if (cr > 0) credit = cr;
      // A lone signed value sitting in only one of the two columns.
      if (charge === 0 && credit === 0) {
        if (dr < 0) charge = Math.abs(dr);
        else if (cr < 0) credit = Math.abs(cr);
      }
    }

    let kind: LedgerRow["kind"] = "other";
    if (charge > 0) kind = "charge";
    else if (credit > 0) kind = "credit";
    if (/brought forward|b\/f|opening balance|c\/f/i.test(`${name} ${ref}`)) kind = "other";

    const isPassBankTransfer = /^bank\s*transfer$/i.test(idRaw.trim());
    // No literal "Passport"/"ID" column doesn't mean no passport data — GDS-style
    // ticket numbers ("3VS A38535385") embed it, and that text now lives in the
    // concatenated reference field once a matching column exists.
    const passport = isPassBankTransfer
      ? null
      : (normPassport(idRaw) ?? extractPassportFromText(name) ?? extractPassportFromText(ref));
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
    const isSecDep = isSecurityDepositText(narrationText);
    // In narration ledgers every positive AMOUNT is an incoming payment/top-up
    // (visa charges are negative). "PY" vouchers are payment vouchers. Either way
    // a credit here is a settlement, not a per-passenger charge — EXCEPT a security
    // deposit, which is a per-passenger item even when it lands in the credit column.
    const isPaymentVoucher = /^PY/i.test(vno);
    const settlement =
      kind === "credit" &&
      !isSecDep &&
      (isPaymentVoucher || isSettlementText(narrationText, paxName) || credit > 0);

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
      passport: kind === "charge" || isSecDep ? passport : null,
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

  // Auto-detect which column holds CHARGES vs CREDITS/PAYMENTS. Both the
  // "NST"/"NTS" family exports (CR = charge — verified: Mavrick, Climate,
  // Mirsal, SmartTrip) and other software configured the opposite way
  // (DR = charge — verified: AL Saffron) use these same generic "DR"/"CR"
  // headers, so the convention can't be told from header text alone.
  // Settlement rows (bank transfers/top-ups, via isSettlementText) are
  // unambiguously PAYMENTS, never charges — summing their DR vs CR values
  // over the whole sheet reveals which column is actually the credit side
  // for THIS file.
  let drSettleSum = 0;
  let crSettleSum = 0;
  for (let r = 1; r < aoa.length; r++) {
    const row = (aoa[r] ?? []) as unknown[];
    const pax = String(row[idxPax] ?? "").trim();
    const desc = String(row[idxDesc] ?? "").trim();
    if (isSettlementText(desc, pax)) {
      drSettleSum += num(row[idxDR]);
      crSettleSum += num(row[idxCR]);
    }
  }
  // Settlements are payments — if they land mostly in DR, DR is the
  // credit/payment column, which means CR must be the charge column (and
  // vice-versa). No settlement rows found (or a tie) → keep the original
  // NST/NTS-family default: CR = charge, DR = credit/payment.
  const crIsCharge = drSettleSum >= crSettleSum;

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
    const chargeVal = crIsCharge ? cr : dr;
    const creditVal = crIsCharge ? dr : cr;

    // IS25/ DocNo = interline / airline ticket. Ticket No. is a PNR-style "001 XXXXXX"
    // — NOT a "3VS PASSPORT" value. Treat these as simple charge rows whose
    // passport cannot be extracted from the ticket field.
    const isFlightRow = /^IS/i.test(docno);

    // "N PAX OMAN BUS SERVICE" — group-booking summary row where our system
    // recorded the combined amount for all passengers in one line.
    const isGroupRow = /^\d+\s+PAX\b/i.test(pax);

    // Reversal rows: DocNo prefix "VR" OR the charge column is negative (void / corrected charge).
    const isReversal = /^VR/i.test(docno) || (chargeVal < 0 && creditVal === 0 && !isGroupRow);

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
      credit = recoverSettlementAmount(row, [idxDate]) || creditVal || Math.abs(chargeVal);
    } else if (isReversal) {
      // VR rows are charge REVERSALS, not bank transfers.
      // They get kind="credit" but settlement=false so they can still match
      // the partner's corresponding refund row via passport + amount.
      kind = "credit";
      credit = Math.abs(chargeVal) || creditVal;
    } else if (chargeVal > 0) {
      kind = "charge";
      charge = chargeVal;
    } else if (creditVal > 0) {
      kind = "credit";
      credit = creditVal;
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
      raw: { docno, ticket, pax, dr, cr, paxText: pax, isGroupRow, isFlightRow, pnr },
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
    // Skip totals / balance summary lines (never rows carrying a real ID).
    if (!passRaw.trim() && isSummaryRow(row)) continue;
    // "Bank Transfer" in the passport column is a payment-row marker, not a real passport
    const isPassBankTransfer = /^bank\s*transfer$/i.test(passRaw.trim());
    const desc = idxDesc >= 0 ? String(row[idxDesc] ?? "") : "";
    const comm = idxPax >= 0 ? String(row[idxPax] ?? "") : "";

    // Resolve charge / credit robustly across the three common layouts:
    //   (a) separate DR & CR columns (positive values)
    //   (b) a single SIGNED "Amount" column (negative = charge, positive = credit)
    //   (c) a single UNSIGNED column the AI mapped to both roles — direction is
    //       genuinely ambiguous from the number alone, so it's resolved by
    //       `mapping.unsignedAmountMeans` (inferred by the AI from the column's
    //       header wording / narration; defaults to "credit" — the prior
    //       hardcoded behavior — when the mapping predates this field).
    let charge = 0;
    let credit = 0;
    const sameAmtCol = idxCharge >= 0 && idxCharge === idxCredit;
    if (sameAmtCol) {
      const v = num(row[idxCharge]);
      if (v < 0) charge = Math.abs(v);
      else if (v > 0) {
        if (mapping.unsignedAmountMeans === "charge") charge = v;
        else credit = v;
      }
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

    const dateStr = idxDate >= 0 ? normalizeDateCell(row[idxDate]) : "";

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
  assertReadableSpreadsheet(buf, file.name);
  // PDF statements (text-based) are converted to a table first.
  let aoa: unknown[][] | null = null;
  if (looksLikePdf(buf)) {
    const { pdfBufToAoa } = await import("./pdf-ledger");
    aoa = await pdfBufToAoa(buf, file.name);
  }
  // Try XLSX first
  if (!aoa) {
    try {
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      aoa = readBestSheetAoa(wb, { raw: true });
      // If aoa has only a single column, probably text → fallback
      if (aoa.length && aoa[0].length < 4) aoa = null;
    } catch {
      aoa = null;
    }
  }
  if (!aoa || !aoa.length) {
    // Fallback: HTML-table ".xls" or delimiter-auto-detected text.
    aoa = textToAoa(buf) as string[][];
  }
  if (!aoa.length) return [];
  const header = (aoa[0] as unknown[]).map((c) => String(c ?? "").trim());
  const upper = header.map((h) => h.toUpperCase());

  let rows: LedgerRow[];
  if (isSoftwareEntryReport(aoa)) {
    rows = parseSoftwareEntryReport(aoa);
  } else if (looksLikeGdsMonthly(aoa)) {
    rows = parseGDSMonthlyLedger(aoa, monthFromFilename(file.name));
  } else if (upper.some((h) => /TRANSACTION\s*ID/.test(h)) && upper.some((h) => /PASSENGER/.test(h))) {
    // Maverick format: Transaction ID,Transaction Date,Agency,Credit,Debit,Balance,Description,Receipt,Passengers
    rows = parseMaverickSupplier(aoa, upper);
  } else if (upper.includes("PASSPORT NO.")) {
    // Format A: Record Time,Description,Passport No.,Comments,DR,CR,Balance
    rows = parsePartnerFormatA(aoa, upper);
  } else if (upper.includes("ITINERARY") && upper.includes("COMMENTS")) {
    // Format B: Itinerary,Type,ID,Reference,Record Time,Description,Dates,Type,Comments,Status,VAT,DR,CR,Balance
    // Passport sits in the 2nd "Type" column at index 7.
    rows = parsePartnerFormatB(aoa, upper);
  } else if (
    upper.includes("RECORD TIME") &&
    upper.includes("TYPE") &&
    upper.includes("COMMENTS") &&
    upper.includes("DESCRIPTION")
  ) {
    // Format C (Climate / Mirsal): Record Time,Description,Type,Comments,DR,CR,Balance
    // Same shape as Format A but the passport/ID lives in the "Type" column, and a
    // few rows are "VS Penalty" fines whose passport is embedded in the Comments.
    rows = parsePartnerFormatC(aoa, upper);
  } else {
    // Unknown layout → generic auto-detecting parser (works for any ledger).
    rows = parseGenericLedger(aoa, "partner");
  }

  // Safety net: any format-specific parser that produced nothing (a mis-detected
  // layout) falls back to the universal generic parser, so an unexpected file
  // still reconciles instead of coming back empty.
  if (!rows.length) rows = parseGenericLedger(aoa, "partner");
  return explodeMultiPax(rows);
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
      kind = "charge";
      charge = cr;
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

/** A bare passport/ID token sitting alone in a cell (1-2 letters + 6-9 digits +
 *  optional trailing letter, or a purely numeric national ID). Used to decide
 *  whether the "Type" column of a Climate/Mirsal ledger holds a passport or a
 *  marker word ("VS Penalty", "Bank Transfer"). */
function looksLikePassportCell(s: string): boolean {
  const t = (s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (t.length < 6) return false;
  return /^[A-Z]{0,2}\d{6,9}[A-Z]?$/.test(t);
}

/**
 * Format C: Record Time | Description | Type | Comments | DR | CR | Balance
 *
 * Used by Climate Tourism and Mirsal Travel. It is the Format-A layout with two
 * twists:
 *   • the passport / national-ID lives in the "Type" column (not "Passport No."),
 *   • some rows are fines where Type = "VS Penalty" and the passport is buried in
 *     the Comments narration (e.g. " U5338066  fine 05-01-2026"), with the visa
 *     description left blank.
 *
 * Row taxonomy:
 *   Type = <passport>           → visa charge (DR), name in Comments, type in Description
 *   Type = "VS Penalty"         → overstay fine (DR), passport from Comments
 *   Type = "Bank Transfer" /
 *     Description "AC-… Bank"    → settlement / payment (CR)
 */
function parsePartnerFormatC(aoa: unknown[][], upper: string[]): LedgerRow[] {
  const col = (n: string) => upper.indexOf(n);
  const idxTime = col("RECORD TIME");
  const idxDesc = col("DESCRIPTION");
  const idxType = col("TYPE");
  const idxComm = col("COMMENTS");
  const idxDR = col("DR");
  const idxCR = col("CR");
  const rows: LedgerRow[] = [];
  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r] ?? [];
    if (!row.length || row.every((c) => c === null || c === undefined || c === "")) continue;

    const desc = String(row[idxDesc] ?? "").trim();
    const typeRaw = String(row[idxType] ?? "").trim();
    const comm = String(row[idxComm] ?? "").trim();
    const dr = num(row[idxDR]);
    const cr = num(row[idxCR]);

    const dateCell = row[idxTime];
    const date =
      dateCell instanceof Date
        ? dateCell.toISOString().slice(0, 10)
        : String(dateCell ?? "")
            .replace(/\s+\d{1,2}:\d{2}(:\d{2})?$/, "")
            .trim();

    const isBF = /brought forward|opening\s*balance/i.test(desc);
    // "VS Penalty" / overstay fine: passport hides inside the Comments text.
    const isPenalty = /\b(penalty|overstay|fine)\b/i.test(typeRaw) || /\bpenalty\b/i.test(desc);
    // Bank / payment row: "Bank Transfer" marker in Type, or an "AC-… Bank" desc.
    const isBankTransfer =
      /^bank\s*transfer$/i.test(typeRaw) || /^AC-?\d+/i.test(desc) || isSettlementText(desc, comm);

    let kind: LedgerRow["kind"] = "other";
    let charge = 0,
      credit = 0;
    if (isBF) {
      kind = "other";
    } else if (isBankTransfer) {
      kind = "credit";
      credit = cr || dr;
    } else if (dr > 0) {
      // Visa charge / penalty / escape-remove — billed to us via DR.
      kind = "charge";
      charge = dr;
    } else if (cr > 0) {
      kind = "credit";
      credit = cr;
    }

    const settlement = kind === "credit" && isBankTransfer;

    // Passport: from the Type column when it is a bare passport; for penalty rows
    // (and any non-passport Type marker) recover it from the Comments narration.
    let passport: string | null = null;
    if (kind === "charge") {
      passport = looksLikePassportCell(typeRaw)
        ? normPassport(typeRaw)
        : extractPassportFromText(`${comm} ${desc}`);
    }

    const isSecDepC = isSecurityDepositText(desc, comm);
    const visaTypeC = isPenalty ? "OVERSTAY FINE" : extractVisaType(`${desc} ${comm}`);
    const scenarioC = detectScenario({ isSettle: settlement, isSecDep: isSecDepC });

    rows.push({
      side: "partner",
      index: rows.length,
      date,
      passport,
      paxName: isBankTransfer ? "BANK TRANSFER" : comm || typeRaw,
      description: desc || (isPenalty ? "OVERSTAY FINE" : ""),
      reference: settlement ? extractBankRef(`${comm} ${desc}`) : "",
      charge,
      credit,
      kind,
      settlement,
      scenario: scenarioC,
      visaType: visaTypeC,
      srcRow: r,
      raw: { dr, cr, typeRaw, comm, desc, paxText: comm },
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

  // Auto-detect which column holds CHARGES vs CREDITS/PAYMENTS — this layout
  // is shared by suppliers with opposite conventions (verified: al-hadaf uses
  // CR = charge; AL Saffron's NINTHWARE export uses DR = charge), and it can't
  // be told from header text alone. Settlement rows ("Bank Transfer" in the
  // 2nd "Type" column, or elsewhere) are unambiguously PAYMENTS — summing
  // their DR vs CR values over the sheet reveals the credit column for THIS
  // file. No settlement rows found (or a tie) → keep the original default:
  // CR = charge, DR = credit/payment.
  let drSettleSum = 0;
  let crSettleSum = 0;
  for (let r = 1; r < aoa.length; r++) {
    const row = (aoa[r] ?? []) as unknown[];
    const desc = String(row[idxDesc] ?? "");
    const ref = String(row[idxRef] ?? "");
    const passRaw = idxPass >= 0 ? String(row[idxPass] ?? "") : "";
    const comm = String(row[idxComm] ?? "");
    if (isSettlementText(desc, ref, passRaw, comm)) {
      drSettleSum += num(row[idxDR]);
      crSettleSum += num(row[idxCR]);
    }
  }
  const crIsCharge = drSettleSum >= crSettleSum;

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
    const chargeVal = crIsCharge ? cr : dr;
    const creditVal = crIsCharge ? dr : cr;
    const isBF = /brought forward/i.test(desc);
    const isRefund = /refund/i.test(desc) || /refund/i.test(passRaw);
    const isBank = isSettlementText(desc, ref, passRaw, comm);
    let kind: LedgerRow["kind"] = "other";
    let charge = 0,
      credit = 0;
    let passport: string | null = null;
    if (isBF) kind = "other";
    else if (chargeVal > 0 && !isRefund && !isBank) {
      kind = "charge";
      charge = chargeVal;
      passport = normPassport(passRaw) ?? extractPassportFromText(comm);
    } else if (creditVal > 0) {
      kind = "credit";
      credit = creditVal;
    } else if (chargeVal > 0) {
      kind = "credit";
      credit = chargeVal;
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

    // A Date object (this file is read with cellDates:true) stringifies to a
    // verbose "Thu Jan 23 2025 08:01:30 GMT+0400 (Gulf Standard Time)" form if
    // not formatted — normalize it like every other parser in this file does.
    const timeCell = row[idxTime];
    const dateStr =
      timeCell instanceof Date
        ? timeCell.toISOString().slice(0, 10)
        : String(timeCell ?? "")
            .replace(/\s+\d{1,2}:\d{2}(:\d{2})?\s*(AM|PM)?$/i, "")
            .trim();

    rows.push({
      side: "partner",
      index: rows.length,
      date: dateStr,
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
  /** Set by the AI low-confidence-match audit pass: "confirmed" when a second,
   *  independent AI read agrees this is genuinely the same transaction as the
   *  rule engine's uncertain match; "suspicious" when the AI thinks it looks
   *  like two DIFFERENT transactions that only coincidentally scored high
   *  enough to pair. Purely informational — never changes `status` itself, so
   *  a wrong AI verdict can't silently corrupt a real match. */
  aiVerified?: "confirmed" | "suspicious";
  /** Enrichment attached from an optional Reference File (e.g. Imran report), matched by passport/name. */
  refInfo?: ReferenceRow;
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
    /**
     * Implied partner÷ours amount factor — the median ratio across confidently
     * matched flight bookings. Captures whatever consistent gap exists between
     * the two ledgers (a currency peg like USD→SAR ≈ 3.75, a markup, or both),
     * detected from the data itself with no currency input required. Lets the UI
     * show a rate-adjusted variance and flag bookings priced far from the norm.
     * 0 when there aren't enough matched amount pairs to infer it.
     */
    impliedRate: number;
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
  const t = normalizeDigits(String(s).trim());
  if (!t) return NaN;

  // Compact YYYYMMDD (e.g. "20260105" — common in bank/ERP exports).
  if (/^\d{8}$/.test(t)) {
    const yy = +t.slice(0, 4);
    const mm = +t.slice(4, 6);
    const dd = +t.slice(6, 8);
    if (yy >= 1990 && yy <= 2100 && mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) {
      return new Date(yy, mm - 1, dd).getTime();
    }
  }

  // Excel serial date (e.g. "45800" or 45800.5) — days since 1899-12-30.
  if (/^\d{4,6}(\.\d+)?$/.test(t)) {
    const serial = parseFloat(t);
    if (serial > 20000 && serial < 80000) {
      // Days since 1899-12-30. Convert to UTC date first, then extract calendar
      // components to build a local date. This is 100% immune to timezone shifts.
      const utcDays = Math.floor(serial - 25569);
      const utcMs = utcDays * 86400 * 1000;
      const tempDate = new Date(utcMs);
      const yy = tempDate.getUTCFullYear();
      const mm = tempDate.getUTCMonth();
      const dd = tempDate.getUTCDate();
      return new Date(yy, mm, dd).getTime();
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

/** True when a row is a refund / reversal / cancellation (money flowing back),
 *  as opposed to a normal forward charge or payment. */
function isReversalRow(r: LedgerRow): boolean {
  return (
    !!r.isReversal ||
    r.scenario === "refund" ||
    r.scenario === "wrong_invoice" ||
    r.scenario === "wrong_client" ||
    r.scenario === "duplicate"
  );
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
  const zeroScore = (): ScoreResult => ({
    score: 0,
    evidence: {
      passport: 0, reference: 0, name: 0, amount: 0, effectiveAmount: 0,
      date: 0, method: "rule", dateDeltaDays: null,
    },
  });

  // ── HARD GATES ──────────────────────────────────────────────────────────────
  // A security deposit must NEVER match a visa charge and vice versa.
  // They are separate financial instruments even when the amounts coincide.
  const oSD = isSecDep(o);
  const pSD = isSecDep(p);
  if (oSD !== pSD) return zeroScore();

  // A refund / reversal (money flowing back) must never pair with a normal
  // forward charge — they are opposite-direction entries. A reversal reconciles
  // against the OTHER ledger's reversal; a charge against a charge. Without this
  // gate the supplier's refund line (same PNR as its invoice) steals the match
  // from the invoice line, leaving the real invoice falsely "only partner".
  if (isReversalRow(o) !== isReversalRow(p)) return zeroScore();

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
      const groupSide = o.raw?.isGroupRow ? absAmount(p) : absAmount(o);
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

  // GDS ↔ Supplier Entry Report boost:
  // When a GDS row (our side, scenario=flight) is matched against a supplier invoice
  // (partner side), the amounts will ALWAYS differ because:
  //   - GDS "Amount" = what we collected from the customer (our retail price)
  //   - Supplier "Credit" = what Kam Air charges us (their published/wholesale rate)
  // Amount mismatch is EXPECTED and should NOT reduce the match score.
  // When the ticket number or PNR matches exactly, the pair is confirmed regardless
  // of the amount difference — treat it as a high-confidence match.
  const gdsToSupplier =
    (o.scenario === "flight" && p.scenario === "flight") &&
    (o.side === "ours" && p.side === "partner") &&
    reference >= 0.99;
  if (gdsToSupplier) {
    score = Math.max(score, 0.87);
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
 * Flag TRUE duplicate entries WITHIN a single ledger: ≥2 rows that share the same
 * passport, (near-)identical amount, the same visa type, AND the SAME date. The
 * date requirement is what separates a real double-entry (the exact same charge
 * keyed twice on one day) from a legitimate RECURRING charge — e.g. a daily
 * overstay fine of 225 booked every day, or a monthly extension billed each month.
 * Those share passport/amount/type but fall on different dates, so they are NOT
 * duplicates and must stay un-flagged for the numbers to be accurate.
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
    // Normalise the date to a day key so the same charge on the same day groups,
    // but the same charge on different days (recurring fines/extensions) does not.
    const dayMs = parseDate(r.date);
    const day = isNaN(dayMs) ? "?" : Math.floor(dayMs / 86400000);
    const key = `${pass}|${Math.round(amt)}|${vt}|${day}`;
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

/**
 * Consolidate multiple per-passenger component charges on ONE side into a single
 * synthetic charge when their SUM equals a single counterpart charge that shares
 * the same passport on the OTHER side.
 *
 * The motivating case (Mirsal / Climate style): a supplier BUNDLES the visa fee
 * and the refundable security deposit onto one invoice line (e.g. 560 + 1035 =
 * 1595), while our ledger books them as two separate rows. Matching them 1:1
 * leaves a permanent "amount difference" for every passenger. Summing our two
 * component rows and pairing the total against the supplier's single line makes
 * them reconcile cleanly.
 *
 * Guard rails — a merge happens ONLY when:
 *   • the components share a passport that also exists on the other side,
 *   • their SUM matches a single other-side charge (within 1.00), and
 *   • no individual component already matches that charge on its own
 *     (so ledgers that keep visa + deposit separate on BOTH sides are untouched).
 * The merged row keeps the NON-deposit row's description / visa type so the
 * security-deposit hard gate still recognises it as a visa charge. Component
 * source-rows are recorded in `raw.componentSrcRows` for the detailed export.
 */
function consolidateComponentCharges(src: LedgerRow[], dst: LedgerRow[]): LedgerRow[] {
  // Other-side charge amounts grouped by normalised passport.
  const dstByPass = new Map<string, number[]>();
  for (const r of dst) {
    if (r.charge <= 0 || r.settlement) continue;
    const p = normRef(r.passport ?? "");
    if (p.length < 4) continue;
    const a = dstByPass.get(p);
    if (a) a.push(r.charge);
    else dstByPass.set(p, [r.charge]);
  }
  if (!dstByPass.size) return src;

  // Our-side charge rows grouped by passport (skip settlements, reversals, group rows).
  const groups = new Map<string, LedgerRow[]>();
  for (const r of src) {
    if (r.charge <= 0 || r.settlement || r.isReversal || r.raw?.isGroupRow) continue;
    const p = normRef(r.passport ?? "");
    if (p.length < 4) continue;
    const a = groups.get(p);
    if (a) a.push(r);
    else groups.set(p, [r]);
  }

  const mergedAway = new Set<LedgerRow>();
  const synthetic: LedgerRow[] = [];
  for (const [p, rowsForPass] of groups) {
    if (rowsForPass.length < 2) continue;
    const dstAmts = dstByPass.get(p);
    if (!dstAmts || !dstAmts.length) continue;

    // A passport can carry several distinct bookings (e.g. a 60-day visa + deposit
    // in January AND a 1-month extension in March). Peel off the rows that already
    // match a supplier line 1:1 — those are tracked separately on both sides and
    // must be left alone — then try to BUNDLE what remains.
    const pool = [...dstAmts];
    const remaining: LedgerRow[] = [];
    for (const r of rowsForPass) {
      const i = pool.findIndex((a) => Math.abs(a - r.charge) < 1.0);
      if (i >= 0) pool.splice(i, 1);
      else remaining.push(r);
    }
    if (remaining.length < 2 || !pool.length) continue;

    // Components of one bundle share the SAME booking date (the visa fee and its
    // security deposit are booked together). Group the leftovers by date and merge
    // a date-group whose SUM equals an as-yet-unmatched supplier line.
    const byDate = new Map<string, LedgerRow[]>();
    for (const r of remaining) {
      const a = byDate.get(r.date);
      if (a) a.push(r);
      else byDate.set(r.date, [r]);
    }
    for (const rows of byDate.values()) {
      if (rows.length < 2) continue;
      const sum = +rows.reduce((s, r) => s + r.charge, 0).toFixed(2);
      const i = pool.findIndex((a) => Math.abs(a - sum) < 1.0);
      if (i < 0) continue;
      pool.splice(i, 1);
      // Face the merged row with the visa row (not the deposit) so it reads as a
      // visa charge for the security-deposit gate.
      const face = rows.find((r) => !isSecDep(r)) ?? rows[0];
      rows.forEach((r) => mergedAway.add(r));
      synthetic.push({
        ...face,
        charge: sum,
        credit: 0,
        kind: "charge",
        scenario: "visa_charge",
        raw: {
          ...face.raw,
          consolidated: true,
          componentSrcRows: rows.map((r) => r.srcRow).filter((x) => x != null),
          componentAmounts: rows.map((r) => r.charge),
          componentDesc: rows.map((r) => r.description),
        },
      });
    }
  }
  if (!mergedAway.size) return src;
  const out = src.filter((r) => !mergedAway.has(r));
  out.push(...synthetic);
  out.forEach((r, i) => (r.index = i));
  return out;
}

/**
 * Identity match between a GDS booking (our side) and a supplier statement line
 * (partner side) for the FLIGHT scenario.
 *
 * The two ledgers price the SAME ticket in different currencies (GDS = retail SAR
 * collected from the customer; supplier = wholesale AED charged to us), so the
 * amounts NEVER agree. Identity is therefore established purely by the shared
 * booking key — the ticket number (per-passenger) or the PNR (per-booking).
 * When that key matches, the pair is a confirmed match and the amount difference
 * is expected rather than a discrepancy.
 */
/**
 * The ticket / PNR booking keys live under different `raw` names depending on
 * the file: a GDS monthly export stores `ticketNum` / `pnr`, while the supplier
 * "Software Entry Report" stores `ticket` / `pnrRef`. Read BOTH names so identity
 * matching works no matter which file the user dropped on which side (the engine
 * was originally tuned for ours=GDS / partner=supplier; users routinely upload
 * them the other way round, which used to collapse every flight to amount_diff).
 */
const ticketKey = (r: LedgerRow): string =>
  normRef((r.raw?.ticketNum as string) ?? (r.raw?.ticket as string) ?? "");
const pnrKey = (r: LedgerRow): string =>
  normRef((r.raw?.pnr as string) ?? (r.raw?.pnrRef as string) ?? "");

function flightIdentityMatch(o: LedgerRow, p: LedgerRow): boolean {
  if (o.scenario !== "flight" || p.scenario !== "flight") return false;
  if (!(o.side === "ours" && p.side === "partner")) return false;
  // Ticket number (10-digit base) — strongest per-passenger key.
  const ot = ticketKey(o);
  const pt = ticketKey(p);
  if (ot.length >= 8 && ot === pt) return true;
  // PNR — booking-level key (one PNR may cover several passengers).
  const opnr = pnrKey(o);
  const ppnr = pnrKey(p);
  if (opnr.length >= 5 && opnr === ppnr) return true;
  // Some supplier systems append a 1-digit invoice-line sequence number onto the
  // bare 6-char PNR (e.g. GDS "18GR56" stored as "18GR561") — treat a 7-vs-6-char
  // PNR that agrees on the first 6 characters as the same booking.
  if (opnr.length === 6 && ppnr.length === 7 && ppnr.slice(0, 6) === opnr) return true;
  if (ppnr.length === 6 && opnr.length === 7 && opnr.slice(0, 6) === ppnr) return true;
  // Letter "O" and digit "0" are visually identical and get transcribed
  // inconsistently between systems for the same PNR (e.g. GDS "18PG0N" vs
  // supplier "18PGON") — fold both before comparing.
  if (
    opnr.length >= 5 &&
    opnr.length === ppnr.length &&
    opnr.replace(/O/g, "0") === ppnr.replace(/O/g, "0")
  ) {
    return true;
  }
  // Fallback: reference field equality (older exports that share the same key).
  const ro = normRef(o.reference);
  return ro.length >= 6 && ro === normRef(p.reference);
}

/**
 * Consolidate flight (GDS ↔ supplier) charge rows that share a PNR into ONE
 * booking row per PNR.
 *
 * Our GDS export lists one row PER PASSENGER, while the supplier's account
 * statement bills one combined line PER PNR (booking). Reconciling them
 * per-passenger leaves every extra passenger of a multi-passenger booking
 * falsely flagged as "missing from partner" — even though the booking itself
 * reconciles perfectly. Rolling each side up to the PNR grain (the grain the
 * supplier actually invoices at) makes a multi-passenger booking reconcile as a
 * single item. Every passenger's name, amount and source row is preserved in
 * `raw` so the detailed export can still break the booking down per passenger.
 *
 * Only FLIGHT charge rows are touched. Payments, refunds, rows without a PNR,
 * and the entire visa / security-deposit path pass through unchanged.
 */
function consolidateFlightBookings(rows: LedgerRow[]): LedgerRow[] {
  const groups = new Map<string, LedgerRow[]>();
  const passthrough: LedgerRow[] = [];
  for (const r of rows) {
    const pnr = normRef((r.raw?.pnr as string) ?? (r.raw?.pnrRef as string) ?? "");
    const isFlightCharge =
      r.scenario === "flight" && r.charge > 0 && !r.settlement && !isReversalRow(r);
    if (!isFlightCharge || pnr.length < 5) {
      passthrough.push(r);
      continue;
    }
    const a = groups.get(pnr);
    if (a) a.push(r);
    else groups.set(pnr, [r]);
  }
  const out = [...passthrough];
  for (const grp of groups.values()) {
    if (grp.length === 1) {
      out.push(grp[0]);
      continue;
    }
    const sum = +grp.reduce((s, r) => s + r.charge, 0).toFixed(2);
    const names = grp.map((r) => r.paxName).filter(Boolean);
    const face = grp[0];
    out.push({
      ...face,
      charge: sum,
      paxName: names.length > 1 ? `${names[0]} +${names.length - 1} more` : (names[0] ?? face.paxName),
      raw: {
        ...face.raw,
        pnrConsolidated: true,
        paxCount: grp.length,
        paxList: names,
        componentSrcRows: grp.map((r) => r.srcRow).filter((x) => x != null),
        componentAmounts: grp.map((r) => r.charge),
      },
    });
  }
  out.forEach((r, i) => (r.index = i));
  return out;
}

/**
 * Same shape as {@link consolidateFlightBookings} but for VISA charges: our
 * ledger often books one row per passenger under a shared DocNo (e.g.
 * "VS26/ 765" covering 2 passengers), while the supplier's statement bills
 * the whole group as one combined invoice line. Without this, every extra
 * passenger in the group falsely shows as "only in our ledger" even though
 * the group as a whole reconciles fine. Only touches rows that share a real
 * (non-blank) DocNo — a blank/missing DocNo never triggers grouping, so
 * unrelated single-passenger charges are never merged by accident.
 */
function consolidateVisaBookings(rows: LedgerRow[]): LedgerRow[] {
  const groups = new Map<string, LedgerRow[]>();
  const passthrough: LedgerRow[] = [];
  for (const r of rows) {
    const docno = normRef((r.raw?.docno as string) ?? "");
    const isVisaCharge =
      r.scenario === "visa_charge" && r.charge > 0 && !r.settlement && !isReversalRow(r);
    if (!isVisaCharge || docno.length < 5) {
      passthrough.push(r);
      continue;
    }
    const a = groups.get(docno);
    if (a) a.push(r);
    else groups.set(docno, [r]);
  }
  const out = [...passthrough];
  for (const grp of groups.values()) {
    if (grp.length === 1) {
      out.push(grp[0]);
      continue;
    }
    const sum = +grp.reduce((s, r) => s + r.charge, 0).toFixed(2);
    const names = grp.map((r) => r.paxName).filter(Boolean);
    const face = grp[0];
    out.push({
      ...face,
      charge: sum,
      paxName: names.length > 1 ? `${names[0]} +${names.length - 1} more` : (names[0] ?? face.paxName),
      raw: {
        ...face.raw,
        docConsolidated: true,
        paxCount: grp.length,
        paxList: names,
        componentSrcRows: grp.map((r) => r.srcRow).filter((x) => x != null),
        componentAmounts: grp.map((r) => r.charge),
      },
    });
  }
  out.forEach((r, i) => (r.index = i));
  return out;
}

export function reconcile(ours: LedgerRow[], partner: LedgerRow[]): ReconResult {
  // Combine bundled component charges (visa fee + security deposit booked as two
  // rows on one side, one line on the other) BEFORE matching so each bundle
  // reconciles as a single item instead of a per-passenger amount difference.
  ours = consolidateComponentCharges(ours, partner);
  partner = consolidateComponentCharges(partner, ours);

  // Roll multi-passenger flight bookings up to one row per PNR on each side so a
  // GDS export (one row per passenger) reconciles against the supplier statement
  // (one combined line per PNR). No-op for visa ledgers — only flight rows match.
  ours = consolidateFlightBookings(ours);
  partner = consolidateFlightBookings(partner);

  // Same idea for multi-passenger VISA bookings, grouped by DocNo instead of
  // PNR (visa charges have no PNR concept).
  ours = consolidateVisaBookings(ours);
  partner = consolidateVisaBookings(partner);

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

    // GDS ↔ Software Entry Report: amounts are in DIFFERENT CURRENCIES and represent
    // different values (GDS = retail price in SAR charged to the customer; Supplier =
    // Kam Air wholesale cost in AED charged to us). Amount mismatch is always expected
    // and should never show as "amount_diff" when the ticket/PNR reference matches.
    const isGdsSupplierPair = flightIdentityMatch(o, p);
    // A consolidated multi-passenger booking matched on PNR: note the head-count
    // so the reviewer knows how many passengers the single supplier line covers.
    const oPax = (o.raw?.paxCount as number) ?? 1;
    const oCur = o.currency;
    const pCur = p.currency;
    const curBasis =
      oCur && pCur && oCur !== pCur
        ? `Our amount is in ${oCur}, supplier amount in ${pCur} — different currencies, so they never match by value.`
        : `Our amount is the retail price charged to the customer; the supplier amount is their cost to us${oCur ? ` (both ${oCur})` : ""}. The pricing gap is normal.`;
    const currencyNote = isGdsSupplierPair
      ? `Confirmed by ticket/PNR reference. ${curBasis}${oPax > 1 ? ` Booking covers ${oPax} passengers on one supplier line.` : ""}`
      : "";

    pairs.push({
      key: `m-${o.index}-${p.index}`,
      status: (exact || isGdsSupplierPair) ? "matched" : "amount_diff",
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
      note: currencyNote || explainMatch(c.evidence, diff, exact, c.score, o, p),
    });
  }
  // Unmatched sub-rows that we split out of a single N-PAX group booking are
  // collapsed back into ONE representative row per booking, so the "only ours"
  // count and the totals reflect real bookings rather than synthetic per-pax
  // splits. (Group sub-rows share the original booking's srcRow.)
  const groupRemnant = new Map<number, LedgerRow[]>();
  ours.forEach((o, oi) => {
    if (m.usedO.has(oi)) return;
    if (o.raw?.isGroupRow && o.srcRow != null) {
      const a = groupRemnant.get(o.srcRow);
      if (a) a.push(o);
      else groupRemnant.set(o.srcRow, [o]);
      return;
    }
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
  for (const rows of groupRemnant.values()) {
    const o = rows[0];
    const amt = +rows.reduce((s, r) => s + absAmount(r), 0).toFixed(2);
    const cnt = (o.raw?.paxCount as number) ?? rows.length;
    const isCharge = o.charge > 0;
    pairs.push({
      key: `oo-grp-${o.srcRow}`,
      status: "missing_partner",
      kind: o.kind === "credit" ? "credit" : "charge",
      ours: { ...o, charge: isCharge ? amt : 0, credit: isCharge ? 0 : amt },
      partner: null,
      oursAmt: amt,
      partnerAmt: 0,
      diff: -amt,
      note: `Group booking — ${rows.length} of ${cnt} passenger(s) unmatched in partner ledger.`,
    });
  }
  // Mirror the `ours`-side group-remnant collapsing above: unmatched sub-rows
  // split out of a single N-PAX partner group booking are collapsed back into
  // ONE representative row per booking instead of spilling out as fragmented
  // per-passenger "missing_ours" rows.
  const partnerGroupRemnant = new Map<number, LedgerRow[]>();
  partner.forEach((p, pi) => {
    if (m.usedP.has(pi)) return;
    if (p.raw?.isGroupRow && p.srcRow != null) {
      const a = partnerGroupRemnant.get(p.srcRow);
      if (a) a.push(p);
      else partnerGroupRemnant.set(p.srcRow, [p]);
      return;
    }
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
  for (const rows of partnerGroupRemnant.values()) {
    const p = rows[0];
    const amt = +rows.reduce((s, r) => s + absAmount(r), 0).toFixed(2);
    const cnt = (p.raw?.paxCount as number) ?? rows.length;
    const isCharge = p.charge > 0;
    pairs.push({
      key: `op-grp-${p.srcRow}`,
      status: "missing_ours",
      kind: p.kind === "credit" ? "credit" : "charge",
      ours: null,
      partner: { ...p, charge: isCharge ? amt : 0, credit: isCharge ? 0 : amt },
      oursAmt: 0,
      partnerAmt: amt,
      diff: amt,
      note: `Group booking — ${rows.length} of ${cnt} passenger(s) unmatched in our ledger.`,
    });
  }

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
    impliedRate: inferImpliedRate(pairs),
  };
}

/** Max gap (in OUR currency units) for a converted match — the larger of a
 *  flat floor (float rounding on tiny amounts) and a small percentage of the
 *  amount itself. A truly PEGGED currency (e.g. 3.75 USD/SAR) has an exact
 *  rate, so the flat floor alone is enough. A FLOATING pair (e.g. AED/SAR)
 *  only has an approximate rate the user types with a handful of decimal
 *  digits — at real booking amounts (thousands), that rounding alone can
 *  exceed a flat cent-tolerance and wrongly flag every genuine match as
 *  "amount_diff". The percentage term absorbs that typed-rate rounding while
 *  staying far tighter than genuine price discrepancies (typically >1%). */
const FX_MATCH_MAX_DIFF = 0.01;
const FX_MATCH_MAX_DIFF_PCT = 0.0025; // 0.25% of the converted amount

/**
 * Re-decide matched vs amount-difference PURELY on amount agreement, after
 * expressing the partner amount in OUR currency via the user's FX peg
 * (1 partner unit = `rate` our units). A paired row keeps "matched" ONLY when
 * the converted amounts agree within a tight tolerance (see
 * {@link FX_MATCH_MAX_DIFF}/{@link FX_MATCH_MAX_DIFF_PCT} — a flat cent floor
 * plus a small percentage, so rounding in a manually-typed rate doesn't fail
 * a genuine match); a real gap, e.g. a double-digit percentage, makes it
 * "amount_diff" — amounts that don't reconcile are NOT counted as matched.
 * Rows that exist on only one side are left untouched.
 *
 * Returns a fresh result with recomputed totals. When conversion is inactive the
 * input is returned unchanged, so single-currency reconciliations behave exactly
 * as the engine decided. Applies identically in single-file and 1-Year mode
 * because both flow through this one result.
 */
export function applyFxMatchAccuracy(
  result: ReconResult,
  opts: { active: boolean; rate: number },
): ReconResult {
  if (!opts.active || !(opts.rate > 0)) return result;
  const { rate } = opts;

  const pairs: Pair[] = result.pairs.map((p) => {
    if (!p.ours || !p.partner) return p; // only-ours / only-partner: nothing to convert
    if (p.status !== "matched" && p.status !== "amount_diff") return p;

    const converted = p.partnerAmt * rate; // partner → our currency
    const diff = +(converted - p.oursAmt).toFixed(2);
    const tolerance = Math.max(FX_MATCH_MAX_DIFF, FX_MATCH_MAX_DIFF_PCT * Math.max(Math.abs(converted), Math.abs(p.oursAmt)));
    const agree = Math.abs(diff) < tolerance;

    return {
      ...p,
      status: agree ? "matched" : "amount_diff",
      diff: agree ? 0 : diff,
      note: agree
        ? `Amounts agree after conversion (1 = ${rate} → ${converted.toFixed(2)} vs ${p.oursAmt.toFixed(2)}).`
        : `Converted amount ${converted.toFixed(2)} ≠ our ${p.oursAmt.toFixed(2)} (Δ ${diff.toFixed(2)}) — not a value match.`,
    };
  });

  // Totals count charges/credits from the row arrays, so rebuild them from the
  // (unchanged) row objects still carried on each pair.
  const ours = pairs.map((p) => p.ours).filter(Boolean) as LedgerRow[];
  const partner = pairs.map((p) => p.partner).filter(Boolean) as LedgerRow[];
  return { pairs, totals: computeTotals(ours, partner, pairs) };
}

/**
 * Infer the consistent partner÷ours amount factor between the two ledgers from
 * the data itself — the MEDIAN ratio across confidently matched flight bookings
 * (single-passenger, both amounts present). The median ignores the handful of
 * mispriced outliers, so it reflects the true underlying rate (a currency peg,
 * a markup, or both) without anyone having to know or enter the currencies.
 * Returns 0 when there aren't enough samples to be meaningful.
 */
export function inferImpliedRate(pairs: Pair[]): number {
  const ratios = pairs
    .filter(
      (p) =>
        p.status === "matched" &&
        p.ours && p.partner &&
        p.oursAmt > 0 && p.partnerAmt > 0 &&
        (p.ours.scenario === "flight" || p.partner.scenario === "flight") &&
        !(p.ours.raw?.paxCount && (p.ours.raw.paxCount as number) > 1),
    )
    .map((p) => p.partnerAmt / p.oursAmt)
    .sort((a, b) => a - b);
  if (ratios.length < 5) return 0;
  return +ratios[Math.floor(ratios.length / 2)].toFixed(4);
}

/**
 * How far a matched pair's amount sits from what the implied rate predicts, as a
 * signed fraction: 0 means exactly on-rate, +0.5 means the supplier charged 50%
 * more than expected (possible over-charge), -0.5 means 50% less. Returns null
 * when the pair has no comparable amounts or the rate is unknown.
 */
export function rateDeviation(p: Pair, impliedRate: number): number | null {
  if (!impliedRate || !p.ours || !p.partner || p.oursAmt <= 0 || p.partnerAmt <= 0) return null;
  const expected = p.oursAmt * impliedRate;
  if (expected <= 0) return null;
  return +(p.partnerAmt / expected - 1).toFixed(4);
}

/* ------------------------------------------------------------------ */
/* ANALYTICS  (per-scenario + duplicate / reversal intelligence)       */
/* ------------------------------------------------------------------ */

export type ScenarioStat = {
  key: Scenario;
  label: string;
  total: number;
  matched: number;
  amountDiff: number;
  onlyOurs: number;
  onlyPartner: number;
  /** Verified (matched) value carried by this scenario. */
  matchedValue: number;
  /** Total absolute value moving through this scenario (any status). */
  grossValue: number;
};

export type ReversalReasonStat = {
  reason: Scenario;
  label: string;
  ours: number;
  partner: number;
};

export type ReconAnalytics = {
  /** Per-scenario rollup, ordered most-common first. */
  scenarios: ScenarioStat[];
  duplicates: {
    /** Rows that belong to a duplicate group, by the ledger they were found on. */
    rowsOurs: number;
    rowsPartner: number;
    /** Distinct duplicate groups (counted once, on the side they occur). */
    groupsOurs: number;
    groupsPartner: number;
    /** Value tied up in duplicate rows (excluding the first of each group). */
    redundantValue: number;
  };
  reversals: {
    ours: number;
    partner: number;
    byReason: ReversalReasonStat[];
  };
};

const REVERSAL_SCENARIOS: Scenario[] = ["wrong_invoice", "wrong_client", "duplicate", "refund"];

const SCENARIO_ORDER: Scenario[] = [
  "visa_charge",
  "security_deposit",
  "multi_passenger",
  "flight",
  "addon",
  "bank_transfer",
  "wrong_invoice",
  "wrong_client",
  "duplicate",
  "refund",
];

/**
 * Roll the reconciled pairs up into category-level intelligence: how every
 * scenario performed (matched / mismatched / unmatched + value), and where the
 * duplicate and reversal (VR) entries actually originate — on OUR ledger or the
 * partner's. This powers the advanced analytics dashboard so a reviewer can see,
 * at a glance, which ledger introduced a duplicate and what each refund was for.
 */
export function computeAnalytics(pairs: Pair[]): ReconAnalytics {
  const stats = new Map<Scenario, ScenarioStat>();
  const ensure = (key: Scenario): ScenarioStat => {
    let s = stats.get(key);
    if (!s) {
      s = {
        key,
        label: SCENARIO_LABEL[key] ?? key,
        total: 0,
        matched: 0,
        amountDiff: 0,
        onlyOurs: 0,
        onlyPartner: 0,
        matchedValue: 0,
        grossValue: 0,
      };
      stats.set(key, s);
    }
    return s;
  };

  const dupSeenOurs = new Set<string>();
  const dupSeenPartner = new Set<string>();
  const duplicates = {
    rowsOurs: 0,
    rowsPartner: 0,
    groupsOurs: 0,
    groupsPartner: 0,
    redundantValue: 0,
  };
  const reversalReasons = new Map<Scenario, ReversalReasonStat>();
  REVERSAL_SCENARIOS.forEach((r) =>
    reversalReasons.set(r, { reason: r, label: SCENARIO_LABEL[r] ?? r, ours: 0, partner: 0 }),
  );
  let reversalsOurs = 0;
  let reversalsPartner = 0;

  const dupKey = (r: LedgerRow) =>
    `${(r.passport ?? "").toUpperCase()}|${Math.round(r.charge > 0 ? r.charge : r.credit)}|${(r.visaType ?? "").toUpperCase()}`;

  for (const p of pairs) {
    const sc = (p.ours?.scenario ?? p.partner?.scenario) as Scenario | undefined;
    if (sc) {
      const s = ensure(sc);
      s.total++;
      const v = p.partnerAmt || p.oursAmt;
      s.grossValue += v;
      if (p.status === "matched") {
        s.matched++;
        s.matchedValue += v;
      } else if (p.status === "amount_diff") s.amountDiff++;
      else if (p.status === "missing_partner") s.onlyOurs++;
      else if (p.status === "missing_ours") s.onlyPartner++;
    }

    // Duplicate origin — flagged independently on each ledger by flagDuplicates.
    for (const side of ["ours", "partner"] as const) {
      const r = p[side];
      if (!r || !r.duplicateCount || r.duplicateCount < 2) continue;
      const amt = r.charge > 0 ? r.charge : r.credit;
      if (side === "ours") {
        duplicates.rowsOurs++;
        if ((r.duplicateIndex ?? 0) > 1) duplicates.redundantValue += amt;
        dupSeenOurs.add(dupKey(r));
      } else {
        duplicates.rowsPartner++;
        if ((r.duplicateIndex ?? 0) > 1) duplicates.redundantValue += amt;
        dupSeenPartner.add(dupKey(r));
      }
    }

    // Reversal / VR origin.
    if (p.ours?.isReversal || (p.ours && REVERSAL_SCENARIOS.includes(p.ours.scenario as Scenario))) {
      reversalsOurs++;
      const reason = (p.ours!.scenario as Scenario) ?? "refund";
      const rr = reversalReasons.get(REVERSAL_SCENARIOS.includes(reason) ? reason : "refund");
      if (rr) rr.ours++;
    }
    if (
      p.partner?.isReversal ||
      (p.partner && REVERSAL_SCENARIOS.includes(p.partner.scenario as Scenario))
    ) {
      reversalsPartner++;
      const reason = (p.partner!.scenario as Scenario) ?? "refund";
      const rr = reversalReasons.get(REVERSAL_SCENARIOS.includes(reason) ? reason : "refund");
      if (rr) rr.partner++;
    }
  }

  duplicates.groupsOurs = dupSeenOurs.size;
  duplicates.groupsPartner = dupSeenPartner.size;
  duplicates.redundantValue = +duplicates.redundantValue.toFixed(2);

  const scenarios = [...stats.values()].sort((a, b) => {
    const ai = SCENARIO_ORDER.indexOf(a.key);
    const bi = SCENARIO_ORDER.indexOf(b.key);
    return (ai < 0 ? 99 : ai) - (bi < 0 ? 99 : bi);
  });
  scenarios.forEach((s) => {
    s.matchedValue = +s.matchedValue.toFixed(2);
    s.grossValue = +s.grossValue.toFixed(2);
  });

  return {
    scenarios,
    duplicates,
    reversals: {
      ours: reversalsOurs,
      partner: reversalsPartner,
      byReason: [...reversalReasons.values()].filter((r) => r.ours + r.partner > 0),
    },
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
  "Ref: Nationality",
  "Ref: Supplier",
  "Ref: Quoted",
  "Ref: Remarks",
];
const PAIR_COL_WIDTHS = [18, 9, 9, 8, 11, 16, 14, 15, 12, 24, 16, 11, 11, 12, 24, 16, 11, 11, 16, 9, 44, 16, 20, 10, 28];

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
    p.refInfo?.nationality ?? "",
    p.refInfo?.supplier ?? "",
    p.refInfo?.quotedAmount ?? "",
    p.refInfo?.remarks ?? "",
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

/* ------------------------------------------------------------------ */
/* ADVANCED INSIGHT SHEETS  (duplicates · refunds/VR · insights)       */
/* ------------------------------------------------------------------ */

/** Format a number as a money string with thousands separators (2 decimals). */
function money(n: number): string {
  return (n ?? 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

type TableRowStyle = { tint?: string; bold?: boolean; section?: boolean };

/**
 * Generic styled-table builder: navy header, zebra body, optional per-row tint /
 * bold / full-width section banner. Keeps the advanced sheets visually consistent
 * with the rest of the workbook without repeating cell-styling boilerplate.
 */
function makeStyledTable(
  headers: string[],
  rows: (string | number)[][],
  opts: {
    widths?: number[];
    rightCols?: Set<number>;
    rowStyle?: (i: number) => TableRowStyle;
  } = {},
) {
  const aoa = [headers, ...rows];
  const ws = XLSXStyle.utils.aoa_to_sheet(aoa);
  const ncols = headers.length;
  ws["!cols"] = (opts.widths ?? headers.map(() => 16)).map((wch) => ({ wch }));
  ws["!freeze"] = { xSplit: 0, ySplit: 1 };
  const right = opts.rightCols ?? new Set<number>();

  for (let c = 0; c < ncols; c++) {
    const ref = XLSXStyle.utils.encode_cell({ r: 0, c });
    ws[ref] = ws[ref] || { t: "s", v: headers[c] };
    ws[ref].s = {
      fill: solid(COL.navy),
      font: { bold: true, color: { rgb: COL.headerText }, sz: 10 },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder,
    };
  }

  rows.forEach((_, i) => {
    const r = i + 1;
    const rs = opts.rowStyle?.(i) ?? {};
    const fill = rs.tint ?? (r % 2 === 0 ? COL.zebra : COL.white);
    for (let c = 0; c < ncols; c++) {
      const ref = XLSXStyle.utils.encode_cell({ r, c });
      ws[ref] = ws[ref] || { t: "s", v: "" };
      ws[ref].s = {
        fill: solid(fill),
        border: thinBorder,
        font: {
          bold: !!rs.bold || !!rs.section,
          color: { rgb: rs.section ? COL.navy : "1F2937" },
          sz: rs.section ? 10 : 9,
        },
        alignment: {
          horizontal: rs.section ? "left" : right.has(c) ? "right" : "left",
          vertical: "center",
        },
      };
    }
  });
  if (rows.length) ws["!autofilter"] = { ref: `A1:${XLSXStyle.utils.encode_col(ncols - 1)}${rows.length + 1}` };
  return ws;
}

const amtOf = (r: LedgerRow) => (r.charge > 0 ? r.charge : r.credit);
const sideLabel = (s: Side) => (s === "ours" ? "OUR LEDGER" : "PARTNER LEDGER");

/** One duplicate group: ≥2 rows on the SAME ledger sharing passport+amount+type. */
export type DupGroup = { side: Side; passport: string; amount: number; visaType: string; rows: { row: LedgerRow; status: PairStatus }[] };

/** Gather every duplicate group from the reconciled pairs, per ledger side. */
export function collectDuplicateGroups(pairs: Pair[]): DupGroup[] {
  const map = new Map<string, DupGroup>();
  for (const p of pairs) {
    for (const side of ["ours", "partner"] as const) {
      const row = p[side];
      if (!row || !row.duplicateCount || row.duplicateCount < 2) continue;
      const pass = (row.passport ?? "").toUpperCase().replace(/[^A-Z0-9]/g, "");
      const amt = Math.round(amtOf(row));
      const vt = (row.visaType ?? "").toUpperCase();
      const key = `${side}|${pass}|${amt}|${vt}`;
      let g = map.get(key);
      if (!g) {
        g = { side, passport: row.passport ?? "—", amount: amtOf(row), visaType: row.visaType ?? "", rows: [] };
        map.set(key, g);
      }
      g.rows.push({ row, status: p.status });
    }
  }
  // Sort copies within a group by their original sheet position.
  const groups = [...map.values()].filter((g) => g.rows.length >= 2);
  groups.forEach((g) => g.rows.sort((a, b) => (a.row.srcRow ?? 0) - (b.row.srcRow ?? 0)));
  // Order: our ledger first, then by passenger.
  groups.sort((a, b) => (a.side === b.side ? a.passport.localeCompare(b.passport) : a.side === "ours" ? -1 : 1));
  return groups;
}

/**
 * "Duplicate Entries" sheet — every duplicate grouped together so the reviewer
 * can see WHERE each duplicate occurs and EVERY copy of it side by side. A banner
 * row introduces each group (which ledger, passenger, amount, how many copies),
 * followed by one tinted row per copy. Nothing is collapsed or hidden.
 */
function buildDuplicatesSheet(pairs: Pair[]) {
  const groups = collectDuplicateGroups(pairs);
  const headers = [
    "Ledger",
    "Group",
    "Copy",
    "Passenger",
    "ID / Passport",
    "Visa Type",
    "Date",
    "Voucher / Ref",
    "Amount",
    "Match Status",
    "Note",
  ];
  const widths = [16, 7, 7, 26, 16, 14, 12, 18, 12, 22, 40];
  const right = new Set([8]);
  const rows: (string | number)[][] = [];
  const styles: TableRowStyle[] = [];

  if (!groups.length) {
    rows.push(["✓ No duplicate entries found on either ledger.", "", "", "", "", "", "", "", "", "", ""]);
    styles.push({ section: true, tint: COL.matchedFill });
  }

  groups.forEach((g, gi) => {
    const redundant = g.amount * (g.rows.length - 1);
    rows.push([
      `▼ ${sideLabel(g.side)}`,
      gi + 1,
      `${g.rows.length}×`,
      g.rows[0].row.paxName || "—",
      g.passport,
      g.visaType,
      "",
      "",
      g.amount,
      `${g.rows.length} copies · ${money(redundant)} redundant`,
      "One of these is likely a mistake — keep one, remove the rest.",
    ]);
    styles.push({ section: true, tint: COL.diffFill });
    g.rows.forEach((c, ci) => {
      rows.push([
        sideLabel(g.side),
        gi + 1,
        `${ci + 1}/${g.rows.length}`,
        c.row.paxName || "—",
        c.row.passport ?? "—",
        c.row.visaType ?? "",
        c.row.date || "—",
        c.row.reference || "—",
        amtOf(c.row),
        STATUS_LABEL[c.status],
        c.row.srcRow != null ? `Original sheet row ${c.row.srcRow + 1}` : "",
      ]);
      styles.push({ tint: ci === 0 ? COL.white : COL.onlyPartnerFill });
    });
  });

  return makeStyledTable(headers, rows, { widths, rightCols: right, rowStyle: (i) => styles[i] });
}

const REFUND_REASON_LABEL: Partial<Record<Scenario, string>> = {
  wrong_invoice: "Wrong Invoice",
  wrong_client: "Wrong Client",
  duplicate: "Duplicate Cancelled",
  refund: "Refund / Money Back",
};

export type RefundRow = {
  side: Side;
  reason: string;
  paxName: string;
  passport: string;
  date: string;
  reference: string;
  amount: number;
  matched: boolean;
  counterparty: string;
  note: string;
};

/** Every money-back / reversal (VR) entry on either ledger, with its reason and
 *  whether it matched the other side. Shared by the Excel sheet and the PDF. */
export function collectRefunds(pairs: Pair[]): RefundRow[] {
  const out: RefundRow[] = [];
  for (const p of pairs) {
    for (const side of ["ours", "partner"] as const) {
      const row = p[side];
      if (!row) continue;
      const sc = row.scenario as Scenario | undefined;
      const isRev =
        !!row.isReversal ||
        (!!sc && ["wrong_invoice", "wrong_client", "duplicate", "refund"].includes(sc));
      if (!isRev) continue;
      const counter = side === "ours" ? p.partner : p.ours;
      out.push({
        side: row.side,
        reason: REFUND_REASON_LABEL[sc ?? "refund"] ?? "Reversal (VR)",
        paxName: row.paxName || "—",
        passport: row.passport ?? "—",
        date: row.date || "—",
        reference: row.reference || "—",
        amount: amtOf(row),
        matched: p.status === "matched" || p.status === "amount_diff",
        counterparty: counter ? `${counter.paxName || counter.passport || "—"} · ${money(amtOf(counter))}` : "—",
        note: p.note,
      });
    }
  }
  return out;
}

/**
 * "Refunds & Reversals (VR)" sheet — every money-back / correction entry on either
 * ledger with its reason, the passenger it belongs to, the amount returned, and
 * whether it was matched to the other ledger. Lets the reviewer audit every
 * reversal without missing a line.
 */
function buildRefundsSheet(pairs: Pair[]) {
  const headers = [
    "Ledger",
    "Reason",
    "Passenger",
    "ID / Passport",
    "Date",
    "Voucher / Ref",
    "Amount Back",
    "Matched?",
    "Matched-To (other ledger)",
    "Note",
  ];
  const widths = [16, 18, 26, 16, 12, 18, 12, 12, 28, 40];
  const right = new Set([6]);
  const rows: (string | number)[][] = [];
  const styles: TableRowStyle[] = [];

  for (const f of collectRefunds(pairs)) {
    rows.push([
      sideLabel(f.side),
      f.reason,
      f.paxName,
      f.passport,
      f.date,
      f.reference,
      f.amount,
      f.matched ? "✓ Matched" : "✗ Only here",
      f.counterparty,
      f.note,
    ]);
    styles.push({ tint: f.matched ? COL.matchedFill : COL.onlyOursFill });
  }

  if (!rows.length) {
    rows.push(["✓ No refunds or reversals found on either ledger.", "", "", "", "", "", "", "", "", ""]);
    styles.push({ section: true, tint: COL.matchedFill });
  }

  return makeStyledTable(headers, rows, { widths, rightCols: right, rowStyle: (i) => styles[i] });
}

/**
 * "Insights" sheet — the management-level read-out: per-scenario performance
 * (matched / mismatch / unmatched + value), where duplicates originate (our vs
 * partner ledger) and the value tied up in them, and a reversal-by-reason split.
 */
function buildInsightsSheet(pairs: Pair[]) {
  const a = computeAnalytics(pairs);
  const headers = ["Category", "Total", "Matched", "Amount Diff", "Only Ours", "Only Partner", "Match %", "Matched Value", "Gross Value"];
  const widths = [22, 9, 9, 11, 10, 12, 9, 15, 15];
  const right = new Set([1, 2, 3, 4, 5, 6, 7, 8]);
  const rows: (string | number)[][] = [];
  const styles: TableRowStyle[] = [];

  rows.push(["PER-CATEGORY PERFORMANCE", "", "", "", "", "", "", "", ""]);
  styles.push({ section: true, tint: COL.sectionFill });
  for (const s of a.scenarios) {
    const pct = s.total ? Math.round((s.matched / s.total) * 100) : 0;
    rows.push([s.label, s.total, s.matched, s.amountDiff, s.onlyOurs, s.onlyPartner, `${pct}%`, s.matchedValue, s.grossValue]);
    styles.push({ tint: pct >= 90 ? COL.matchedFill : pct >= 60 ? COL.diffFill : COL.onlyPartnerFill });
  }

  rows.push(["", "", "", "", "", "", "", "", ""]);
  styles.push({});
  rows.push(["DUPLICATE ENTRIES — WHERE THEY ORIGINATE", "", "", "", "", "", "", "", ""]);
  styles.push({ section: true, tint: COL.sectionFill });
  rows.push(["Our ledger", a.duplicates.rowsOurs, "", "", "", "", "", "", `${a.duplicates.groupsOurs} groups`]);
  styles.push({ tint: COL.onlyOursFill });
  rows.push(["Partner ledger", a.duplicates.rowsPartner, "", "", "", "", "", "", `${a.duplicates.groupsPartner} groups`]);
  styles.push({ tint: COL.onlyPartnerFill });
  rows.push(["Redundant value (extra copies)", "", "", "", "", "", "", a.duplicates.redundantValue, ""]);
  styles.push({ bold: true, tint: COL.diffFill });

  rows.push(["", "", "", "", "", "", "", "", ""]);
  styles.push({});
  rows.push(["REVERSALS / REFUNDS — BY REASON", "", "", "", "", "", "", "", ""]);
  styles.push({ section: true, tint: COL.sectionFill });
  rows.push(["Reason", "Our ledger", "Partner ledger", "", "", "", "", "", ""]);
  styles.push({ bold: true });
  if (a.reversals.byReason.length) {
    for (const r of a.reversals.byReason) {
      rows.push([r.label, r.ours, r.partner, "", "", "", "", "", ""]);
      styles.push({});
    }
  } else {
    rows.push(["None found", "", "", "", "", "", "", "", ""]);
    styles.push({});
  }

  return makeStyledTable(headers, rows, { widths, rightCols: right, rowStyle: (i) => styles[i] });
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
  opts?: {
    oursAoa?: unknown[][] | null;
    partnerAoa?: unknown[][] | null;
    monthlyBreakdown?: MonthlyBreakdown[];
    /** Currency-conversion context so amounts can be shown in one base unit. */
    fx?: { rate: number; oursCcy: string; partnerCcy: string; active: boolean };
  },
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
    S(["NAWI SAADI — AI LEDGER RECONCILIATION REPORT"], "title"),
    S(["Generated", new Date().toLocaleString()], "kv"),
    S([], "blank"),
    S(["OVERVIEW"], "section"),
    S(["Our ledger rows", totals.oursRows], "kv"),
    S(["Partner ledger rows", totals.partnerRows], "kv"),
    S(["Total reconciled items", pairs.length], "kv"),
    S(["Rows paired", totalPaired], "kv"),
    // Exact match rate: only amount-agreeing matches count (mirrors the app UI),
    // so a paired-but-amount-off row is not counted as a match here either.
    S(["Match rate %", pairs.length ? Math.round((totals.matched / pairs.length) * 100) : 0], "kv"),
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

  /* ---- Action Items: the few things worth chasing, in one base currency ---- */
  {
    const fx = opts?.fx;
    const rate = fx?.active && fx.rate > 0 ? fx.rate : 0; // partner = ours * rate
    const conv = rate > 0 ? 1 / rate : 1; // supplier amount -> our currency
    const base = fx?.active ? fx.oursCcy : "";
    const m = (n: number) => `${base ? base + " " : ""}${(+n.toFixed(2)).toLocaleString()}`;

    const onlyOurs = pairs.filter((p) => p.status === "missing_partner");
    const onlyPartner = pairs.filter((p) => p.status === "missing_ours");
    const ooVal = onlyOurs.reduce((s, p) => s + (p.oursAmt || 0), 0);
    const opVal = onlyPartner.reduce((s, p) => s + (p.partnerAmt || 0) * conv, 0);

    const priced = pairs
      .filter((p) => p.status === "matched" && p.ours && p.partner && p.oursAmt > 0 && p.partnerAmt > 0)
      .map((p) => ({ p, dev: rate > 0 ? rateDeviation(p, rate) : null, impact: p.partnerAmt * conv - p.oursAmt }))
      .filter((x) => x.dev !== null && Math.abs(x.dev) > 0.15)
      .sort((a, b) => Math.abs(b.impact) - Math.abs(a.impact));
    const overTotal = priced.filter((x) => x.impact > 0).reduce((s, x) => s + x.impact, 0);
    const dups = pairs.filter((p) => (p.ours?.duplicateCount ?? 0) > 1 || (p.partner?.duplicateCount ?? 0) > 1).length;

    const ai: (string | number)[][] = [
      ["ACTION ITEMS — what to check first"],
      fx?.active ? [`All amounts in ${base} (supplier converted at 1 USD = ${fx.rate} SAR)`] : ["Amounts as recorded on each ledger"],
      [],
      ["Money at risk (unmatched)", ""],
      [`  Only in our ledger (${onlyOurs.length})`, m(ooVal)],
      [`  Only in supplier ledger (${onlyPartner.length})`, m(opVal)],
      ["  Total at risk", m(ooVal + opVal)],
      [],
      ["Price looks off", ""],
      [`  Bookings beyond ±15% of rate`, priced.length],
      ["  Total over-charge vs rate", m(overTotal)],
      ["Possible duplicate entries", dups],
      [],
      ["TOP PRICE DISCREPANCIES"],
      ["Reference", "Passenger", "Our amount", "Supplier (converted)", "% off", "Impact"],
      ...priced.slice(0, 25).map((x) => [
        x.p.ours?.reference || x.p.partner?.reference || "",
        x.p.ours?.paxName || x.p.partner?.paxName || "",
        +x.p.oursAmt.toFixed(2),
        +(x.p.partnerAmt * conv).toFixed(2),
        `${x.dev! > 0 ? "+" : ""}${Math.round(x.dev! * 100)}%`,
        +x.impact.toFixed(2),
      ]),
    ];
    const wsAi = XLSXStyle.utils.aoa_to_sheet(ai);
    wsAi["!cols"] = [{ wch: 22 }, { wch: 22 }, { wch: 16 }, { wch: 20 }, { wch: 8 }, { wch: 14 }];
    const titleStyle = { fill: solid(COL.navy), font: { bold: true, sz: 12, color: { rgb: COL.gold } } };
    if (wsAi["A1"]) wsAi["A1"].s = titleStyle;
    if (wsAi["A14"]) wsAi["A14"].s = { font: { bold: true, sz: 11, color: { rgb: COL.navy } }, fill: solid(COL.sectionFill) };
    ["A15", "B15", "C15", "D15", "E15", "F15"].forEach((a) => {
      if (wsAi[a]) wsAi[a].s = { fill: solid(COL.navy), font: { bold: true, color: { rgb: "FFFFFF" } } };
    });
    XLSXStyle.utils.book_append_sheet(wb, wsAi, "Action Items");
  }

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

  /* ---- Advanced insight sheets ---- */
  XLSXStyle.utils.book_append_sheet(wb, buildInsightsSheet(pairs), "Insights");
  XLSXStyle.utils.book_append_sheet(wb, buildDuplicatesSheet(pairs), "Duplicate Entries");
  XLSXStyle.utils.book_append_sheet(wb, buildRefundsSheet(pairs), "Refunds & Reversals");

  XLSXStyle.utils.book_append_sheet(wb, styledPairSheet(settlementPairs), "Settlements");
  XLSXStyle.utils.book_append_sheet(wb, styledPairSheet(exceptionPairs), "Exceptions");

  /* ---- Monthly Breakdown sheet (Year Mode only) ---- */
  if (opts?.monthlyBreakdown && opts.monthlyBreakdown.length > 0) {
    const mb = opts.monthlyBreakdown;
    const mbHeaders = ["Month", "Total", "Matched", "Amt Diff", "Only Ours", "Only Partner", "Match %", "Our Amount", "Partner Amount"];
    const mbRows: (string | number)[][] = mb.map((bk) => [
      bk.label,
      bk.total,
      bk.matched,
      bk.amountDiff,
      bk.onlyOurs,
      bk.onlyPartner,
      `${Math.round(bk.matchRate * 100)}%`,
      +bk.oursTotal.toFixed(2),
      +bk.partnerTotal.toFixed(2),
    ]);
    // Totals row
    const sumTotal = mb.reduce((s, b) => s + b.total, 0);
    const sumMatched = mb.reduce((s, b) => s + b.matched, 0);
    const sumAmtDiff = mb.reduce((s, b) => s + b.amountDiff, 0);
    const sumOnlyOurs = mb.reduce((s, b) => s + b.onlyOurs, 0);
    const sumOnlyPartner = mb.reduce((s, b) => s + b.onlyPartner, 0);
    const sumOursAmt = +mb.reduce((s, b) => s + b.oursTotal, 0).toFixed(2);
    const sumPartnerAmt = +mb.reduce((s, b) => s + b.partnerTotal, 0).toFixed(2);
    mbRows.push(["TOTAL", sumTotal, sumMatched, sumAmtDiff, sumOnlyOurs, sumOnlyPartner,
      `${sumTotal ? Math.round(sumMatched / sumTotal * 100) : 0}%`, sumOursAmt, sumPartnerAmt]);

    const wsMb = XLSXStyle.utils.aoa_to_sheet([mbHeaders, ...mbRows]);
    // Style header row
    mbHeaders.forEach((_, ci) => {
      const addr = XLSXStyle.utils.encode_cell({ r: 0, c: ci });
      if (wsMb[addr]) wsMb[addr].s = {
        fill: { patternType: "solid", fgColor: { rgb: COL.navy } },
        font: { bold: true, color: { rgb: "FFFFFF" } },
        alignment: { horizontal: "center" },
      };
    });
    // Style totals row
    const totR = mbRows.length; // 0-indexed last row = mbRows.length (header is row 0, data starts row 1)
    mbHeaders.forEach((_, ci) => {
      const addr = XLSXStyle.utils.encode_cell({ r: totR, c: ci });
      if (wsMb[addr]) wsMb[addr].s = { font: { bold: true, color: { rgb: COL.navy } }, fill: { patternType: "solid", fgColor: { rgb: COL.sectionFill } } };
    });
    wsMb["!cols"] = [{ wch: 14 }, { wch: 8 }, { wch: 9 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 9 }, { wch: 14 }, { wch: 16 }];
    XLSXStyle.utils.book_append_sheet(wb, wsMb, "Monthly Breakdown");

    /* ---- One tab PER MONTH (Year Mode): that month's bookings on its own sheet,
       so the export is browsable month-by-month, not just one giant list. ---- */
    const usedNames = new Set<string>();
    for (const bk of mb) {
      if (bk.total === 0) continue;
      const monthPairs = pairs.filter((p) => pairMonth(p) === bk.month);
      if (!monthPairs.length) continue;
      // Excel tab names: ≤31 chars, and none of : \ / ? * [ ]
      let name = bk.label.replace(/[:\\/?*[\]]/g, " ").trim().slice(0, 31) || bk.month;
      let n = 2;
      while (usedNames.has(name.toLowerCase())) name = `${bk.label} (${n++})`.slice(0, 31);
      usedNames.add(name.toLowerCase());
      XLSXStyle.utils.book_append_sheet(wb, styledPairSheet(monthPairs), name);
    }
  }

  return XLSXStyle.write(wb, { bookType: "xlsx", type: "array" }) as ArrayBuffer;
}

/* ------------------------------------------------------------------ */
/* MONTHLY GDS BOOKING PARSER  (our side — monthly exports)            */
/* ------------------------------------------------------------------ */

/**
 * Infer an ISO month string "YYYY-MM" from a file name like "JAN 2026.xlsx",
 * "september 2025.xlsx", "1-13MAY 2026.xlsx", etc.
 */
export function monthFromFilename(name: string): string {
  const up = name.toUpperCase();
  const MONTH_NAMES: Record<string, string> = {
    JAN: "01", FEB: "02", MAR: "03", APR: "04", MAY: "05", JUN: "06",
    JUL: "07", AUG: "08", SEP: "09", OCT: "10", NOV: "11", DEC: "12",
  };
  const yearM = up.match(/\b(20\d{2})\b/);
  const year = yearM ? yearM[1] : new Date().getFullYear().toString();
  for (const [abbr, num] of Object.entries(MONTH_NAMES)) {
    if (up.includes(abbr)) return `${year}-${num}`;
  }
  return "";
}

/**
 * Parse a monthly GDS booking export (our side).
 *
 * Two detected sub-formats (both share PNR + Full Name + Amount):
 *   OLD (Sept–Dec 2025): Date | PNR | Ticket No | Fare | Com. Amount | Payment type | Amount | User | First Name | Last Name | Full Name | Class And Segment
 *   NEW (Jan 2026+): Date | PNR | SSR | Tax | NR | Surcharge | Service Fee | Com. Amount | Other Charges | Amount | Payment Amount with Type | User | First Name | Last Name | Full Name | Class And Segment
 *
 * Each row becomes a LedgerRow with:
 *   reference = PNR
 *   paxName   = Full Name
 *   charge    = Amount (positive)
 *   month     = derived from filename
 */
export function parseGDSMonthlyLedger(aoa: unknown[][], monthLabel: string): LedgerRow[] {
  if (!aoa.length) return [];

  // Find header row (first row with "PNR" or "DATE" heading — scan up to 10 rows)
  let hRow = 0;
  for (let r = 0; r < Math.min(10, aoa.length); r++) {
    const row = (aoa[r] as unknown[]).map((c) => String(c ?? "").toUpperCase().trim());
    if (row.some((h) => h === "PNR" || h === "DATE")) {
      hRow = r;
      break;
    }
  }

  const header = (aoa[hRow] as unknown[]).map((c) => String(c ?? "").trim().toUpperCase());
  const col = (re: RegExp) => header.findIndex((h) => re.test(h));

  const idxDate   = col(/^DATE$/);
  const idxPnr    = col(/^PNR$/);
  const idxTicket = col(/TICKET/);
  // Amount: prefer plain "Amount" column (clean numeric). "Payment Amount with Type" has
  // mixed text like "498.00 EX 41.00 INV" that confuses parseFloat — only use it as
  // fallback when no plain Amount column exists.
  const idxPlainAmt = col(/^AMOUNT$/);
  const idxPayAmt   = col(/PAYMENT\s*AMOUNT/);
  const idxAmt      = idxPlainAmt >= 0 ? idxPlainAmt : idxPayAmt;
  const idxFull   = col(/FULL\s*NAME/);
  const idxFirst  = col(/FIRST\s*NAME/);
  const idxLast   = col(/LAST\s*NAME/);
  const idxClass  = col(/CLASS/);

  const rows: LedgerRow[] = [];

  for (let r = hRow + 1; r < aoa.length; r++) {
    const row = (aoa[r] as unknown[]) ?? [];
    if (!row.length || row.every((c) => c === null || c === undefined || c === "")) continue;

    const dateRaw = idxDate >= 0 ? row[idxDate] : "";
    let dateStr = "";
    if (dateRaw instanceof Date) dateStr = dateRaw.toISOString().slice(0, 10);
    else if (dateRaw) {
      // Strip time component (e.g. "03-01-2026 12:38" → "03-01-2026")
      dateStr = String(dateRaw).replace(/\s+\d{1,2}:\d{2}(:\d{2})?$/, "").trim();
    }

    const pnr     = idxPnr >= 0 ? String(row[idxPnr] ?? "").trim() : "";
    const ticket  = idxTicket >= 0 ? String(row[idxTicket] ?? "").trim() : "";
    const amtRaw  = idxAmt >= 0 ? String(row[idxAmt] ?? "").trim() : "";
    // Strip trailing " INV" / " NPC" etc. from payment-type cells like "465.00 INV"
    const amtClean = amtRaw.replace(/\s+[A-Z]+\s*$/i, "");
    const amount  = num(amtClean);

    const fullName = idxFull >= 0 ? String(row[idxFull] ?? "").trim() : "";
    const firstName = idxFirst >= 0 ? String(row[idxFirst] ?? "").trim() : "";
    const lastName  = idxLast  >= 0 ? String(row[idxLast]  ?? "").trim() : "";
    const classInfo = idxClass >= 0 ? String(row[idxClass] ?? "").trim() : "";

    // Build passenger name: full name is best; fallback to first + last
    const paxName = fullName || [firstName, lastName].filter(Boolean).join(" ");
    // Strip gender suffix like "MAROFA\F" → "MAROFA"
    const cleanPax = paxName.replace(/\\[MFI]\s*$/, "").trim();

    // Skip rows with no PNR and no amount, and rows with negative/zero amounts
    // (these are exchange reversals "EX" — they cancel a previous booking and
    // should not appear as unmatched flights in the output).
    if (!pnr && !amount) continue;
    if (amount <= 0 && pnr) {
      // A booking with a real PNR but an amount that came back <= 0 is usually
      // a genuine "EX" reversal — but if the source cell wasn't actually
      // numeric (parse failure defaulting to 0) rather than a real zero/
      // negative value, this is more likely a malformed row silently lost.
      // Surfaced as a console warning only — behavior is unchanged to avoid
      // regressing the documented baseline match counts.
      if (amtClean.trim() && !/^-?\d/.test(amtClean.trim())) {
        console.warn(
          `[parseGDSMonthlyLedger] Dropped row with PNR "${pnr}" — amount cell "${amtRaw}" did not parse as a number.`,
        );
      }
      continue;
    }

    // Ticket number: extract the 10-13 digit number from the ticket cell.
    // GDS stores the full IATA ticket with a 3-digit airline code prefix
    // (e.g. "3842407525000" = airline "384" + base ticket "2407525000").
    // The supplier's software entry report stores ONLY the base 10-digit number.
    // Normalize by stripping the 3-char airline code prefix so both sides match.
    const ticketFull = ticket.match(/\b(\d{10,13})\b/)?.[1] ?? "";
    // Strip leading 3-digit airline code if ticket is 13 digits
    const ticketNum = ticketFull.length === 13 ? ticketFull.slice(3) : ticketFull;
    // Primary reference: normalized ticket (matches supplier format); fallback to PNR
    const refStr = ticketNum || pnr;

    rows.push({
      side: "ours",
      index: rows.length,
      date: dateStr,
      passport: null,
      paxName: cleanPax,
      // description carries both PNR and class so the text scorer has context
      description: [classInfo, pnr ? `PNR:${pnr}` : ""].filter(Boolean).join(" "),
      reference: refStr,
      charge: amount > 0 ? amount : 0,
      credit: 0,
      kind: amount > 0 ? "charge" : "other",
      settlement: false,
      scenario: "flight",
      month: monthLabel,
      // Currency inferred from the flight route (Saudi airports → SAR, UAE → AED).
      currency: detectCurrency(classInfo),
      srcRow: r,
      raw: { pnr, ticket: ticketFull, ticketNum, fullName: cleanPax, classInfo },
    });
  }
  return rows;
}

/* ------------------------------------------------------------------ */
/* SOFTWARE ENTRY REPORT PARSER  (supplier / partner side)             */
/* ------------------------------------------------------------------ */

/**
 * Parse the supplier's "Software Entry Report" / account statement.
 *
 * Format (Nawi Saadi Travel accounting software):
 *   Row 0–12: company header, account info, date range
 *   Row 13: headers — Date | Doc No | Issue Date | Ticket/Voucher No | Type of sales |
 *            Type of services | Booking staff | Airline Code | Service Provider | Client Code |
 *            Client Name | Pax | Class | Travel Date | Confirmation | Supplier Invoice Number |
 *            Supplier Invoice Date | Cost Centre | Department | Debit | Credit | Balance |
 *            PNR/Reference | Narration | Voucher Reco No | Due Date | Created User
 *   Row 14: first data row
 *   Row N-1: totals row (starts with "TOTAL" or "Amount Due")
 *
 * INVOICE rows → charge (Credit column = amount billed by supplier to us)
 * PAYMENT rows → settlement/credit (Debit column = payment we made to them)
 */
export function parseSoftwareEntryReport(aoa: unknown[][]): LedgerRow[] {
  if (!aoa.length) return [];

  // Account currency declared in the statement header (e.g. "Currency: SAR").
  const stmtCurrency = detectStatementCurrency(aoa);

  // Find the header row (contains "Doc No", "Ticket / Voucher No", or "Type of Sales")
  let hRow = 0;
  for (let r = 0; r < Math.min(25, aoa.length); r++) {
    const row = (aoa[r] as unknown[]).map((c) => String(c ?? "").trim().toUpperCase());
    if (
      row.some((h) => h === "DOC NO" || h.includes("VOUCHER NO") || h === "TYPE OF SALES") ||
      (row.includes("DEBIT") && row.includes("CREDIT"))
    ) {
      hRow = r;
      break;
    }
  }

  const header = (aoa[hRow] as unknown[]).map((c) => String(c ?? "").trim().toUpperCase());
  const col = (re: RegExp) => header.findIndex((h) => re.test(h));

  const idxDate      = col(/^DATE$/);
  const idxDocNo     = col(/^DOC\s*NO$/);
  const idxTicket    = col(/TICKET\s*\/\s*VOUCHER/i) >= 0 ? col(/TICKET\s*\/\s*VOUCHER/i) : col(/TICKET/);
  const idxType      = col(/TYPE\s*OF\s*SALES/);
  const idxPax       = col(/^PAX$/);
  const idxClass     = col(/^CLASS$/);
  const idxTravelDt  = col(/TRAVEL\s*DATE/);
  const idxDebit     = col(/^DEBIT$/);
  const idxCredit    = col(/^CREDIT$/);
  const idxPnrRef    = col(/PNR\s*\/\s*REFERENCE|PNR.*REF/);
  const idxNarration = col(/^NARRATION$/);
  const idxClientNm  = col(/CLIENT\s*NAME/);

  const rows: LedgerRow[] = [];

  for (let r = hRow + 1; r < aoa.length; r++) {
    const row = (aoa[r] as unknown[]) ?? [];
    if (!row.length || row.every((c) => c === null || c === undefined || c === "")) continue;

    // Stop at totals / summary rows. The "TOTAL" label may appear in any cell
    // (not just column 0 — the supplier report puts it in column 18).
    const anyTotal = (row as unknown[]).some((c) =>
      /^\s*(TOTAL|AMOUNT\s*DUE)\s*$/i.test(String(c ?? ""))
    );
    if (anyTotal) break;

    const dateRaw = idxDate >= 0 ? row[idxDate] : "";
    let dateStr = "";
    if (dateRaw instanceof Date) dateStr = dateRaw.toISOString().slice(0, 10);
    else if (dateRaw) dateStr = String(dateRaw).replace(/\s+\d{1,2}:\d{2}.*$/, "").trim();

    const docNo    = idxDocNo  >= 0 ? String(row[idxDocNo]  ?? "").trim() : "";
    const ticket   = idxTicket >= 0 ? String(row[idxTicket] ?? "").trim() : "";
    const saleType = idxType   >= 0 ? String(row[idxType]   ?? "").trim().toUpperCase() : "";
    const pax      = idxPax    >= 0 ? String(row[idxPax]    ?? "").trim() : "";
    const classVal = idxClass  >= 0 ? String(row[idxClass]  ?? "").trim() : "";
    const travelDt = idxTravelDt >= 0 ? String(row[idxTravelDt] ?? "").trim() : "";
    const debit    = idxDebit  >= 0 ? num(row[idxDebit])  : 0;
    const credit   = idxCredit >= 0 ? num(row[idxCredit]) : 0;
    const pnrRef   = idxPnrRef >= 0 ? String(row[idxPnrRef] ?? "").trim() : "";
    const narr     = idxNarration >= 0 ? String(row[idxNarration] ?? "").trim() : "";
    const clientNm = idxClientNm >= 0 ? String(row[idxClientNm] ?? "").trim() : "";

    // Skip rows with no financial data and no PNR
    if (!debit && !credit && !pnrRef && !ticket) continue;

    const isPayment = saleType === "PAYMENT" || /^PV/.test(docNo);
    const isRefund  = saleType === "REFUND"  || /^RFD/i.test(docNo);
    const isInvoice = !isPayment && !isRefund && (saleType === "INVOICE" || /^INV/.test(docNo) || (credit > 0 && debit === 0));

    // Extract the clean PNR reference (strip leading "PNR" text if present).
    // Supplier stores PNRs as "PNR19Y35A" in the Ticket/Voucher field for newer
    // entries and as a plain 10-digit ticket number for older entries. Manual
    // data entry also produces typo'd variants of the same "PNR" label ("PNE",
    // "PNRF", "PNER" — R/E are adjacent keys). A real PNR code is a short
    // alphanumeric token (5-8 chars, no spaces) with no reliable constraint on
    // its first character — plenty of real PNRs start with a letter (e.g.
    // "KBEXI6", "IWUNC8"), so requiring a digit right after the label (as an
    // earlier version of this regex did) silently failed to strip most of
    // them. Requiring the REST of the string to look PNR-shaped (rather than
    // requiring a digit specifically) is what actually distinguishes the
    // label from real content — a passenger name/narration in this column
    // has spaces and won't match a bare 5-8 char alnum tail. The extra
    // letter(s) are restricted to R/E/F — the actual typo variants observed
    // in real data (R/E are adjacent QWERTY keys) — so an unrelated code that
    // happens to start with "PN" isn't mistaken for the label.
    const PNR_PREFIX = /^PN[REF]{1,2}(?=[A-Z0-9]{5,8}$)/i;
    const cleanPnr = pnrRef.replace(PNR_PREFIX, "").trim();
    // Also strip the same PNR-label prefix from the ticket field when it acts as a
    // PNR placeholder there instead.
    const ticketClean = ticket.replace(PNR_PREFIX, "").trim();
    // Some SOA-style exports (e.g. a Dubai client statement) glue the literal
    // "PNR" marker directly onto the code inside the Ticket field, after a
    // carrier/office prefix — "255 - PNRMNPDRQ" for the first passenger of a
    // booking, "255 - PNRMNPDRQ1".."...5" for the rest (trailing digit = pax
    // line number, same suffix convention as the Kam Air case above). Extract
    // the bare 6-char PNR so every passenger of one booking shares one key.
    const embeddedPnr = ticket.match(/PNR([A-Z0-9]{6})\d?\b/i)?.[1]?.toUpperCase() ?? "";
    // Prefer PNR from the dedicated column; else the embedded-in-ticket PNR;
    // else cleaned ticket; else docNo
    const reference = cleanPnr || embeddedPnr || ticketClean || docNo;

    // Raw ticket number: 10-digit base number (supplier never includes airline code prefix)
    const ticketNum = ticketClean.match(/\b(\d{10,13})\b/)?.[1] ?? ticketClean;

    let kind: LedgerRow["kind"] = "other";
    let charge = 0, creditAmt = 0;

    if (isPayment) {
      kind = "credit";
      creditAmt = debit > 0 ? debit : credit;
    } else if (isRefund) {
      // Refund: supplier returns money to us — treat as negative invoice (credit note)
      kind = "credit";
      creditAmt = debit > 0 ? debit : credit;
    } else if (isInvoice) {
      kind = "charge";
      charge = credit > 0 ? credit : debit;
    } else if (credit > 0) {
      kind = "charge";
      charge = credit;
    } else if (debit > 0) {
      kind = "credit";
      creditAmt = debit;
    }

    // Use the booking/issue date for month bucketing (matches the monthly file
    // each booking was exported in); fall back to travel date if no issue date.
    const effectiveDate = dateStr || travelDt;
    // Tag the supplier row with its month so the per-month filter works on the
    // supplier side too (supplier-only rows have no `ours` counterpart).
    const partnerMonth = monthKeyFromDate(dateStr || travelDt);

    rows.push({
      side: "partner",
      index: rows.length,
      date: effectiveDate,
      passport: null,
      paxName: pax || narr || clientNm,
      // Include reference in description so extractAdvancedRefs finds ticket/PNR via text scorer
      description: [classVal, reference ? `REF:${reference}` : "", narr].filter(Boolean).join(" ") || saleType,
      reference,
      charge,
      credit: creditAmt,
      kind,
      settlement: isPayment,
      isReversal: isRefund,
      scenario: isPayment ? "bank_transfer" : isRefund ? "refund" : "flight",
      month: partnerMonth || undefined,
      // Statement-level currency (from the header), refined per row if the
      // narration/class names a route or explicit code.
      currency: detectCurrency(classVal, narr) ?? stmtCurrency,
      srcRow: r,
      raw: { docNo, ticket: ticketNum, pnrRef: cleanPnr || embeddedPnr, saleType, pax, classVal, travelDt },
    });
  }
  return rows;
}

/* ------------------------------------------------------------------ */
/* MULTI-FILE MERGE                                                     */
/* ------------------------------------------------------------------ */

/**
 * Merge multiple parsed ledger arrays (from different monthly files) into one.
 * Re-assigns sequential indexes. Preserves the `month` field set per file.
 */
export function mergeLedgers(arrays: LedgerRow[][]): LedgerRow[] {
  const merged: LedgerRow[] = [];
  for (const arr of arrays) {
    for (const row of arr) {
      merged.push({ ...row, index: merged.length });
    }
  }
  return merged;
}

/**
 * Build a display AOA (header row + one row per entry) from parsed ledger rows.
 *
 * Used by the "Uploaded Source Files" panel, the Both-Sheets full ledger and the
 * Excel export when rows come from a MERGED multi-month (Year Mode) upload, where
 * there is no single original sheet to show. Each side gets a clean normalized
 * view of every entry across all uploaded months.
 *
 * Side effect: reassigns each row's `srcRow` to its 1-based position in this AOA
 * (header = row 0) so the full-ledger status map and the Excel export — both of
 * which look rows up by `srcRow` — line up exactly with these rows.
 */
export function ledgerRowsToAoa(rows: LedgerRow[]): unknown[][] {
  const aoa: unknown[][] = [
    ["Date", "Month", "Reference (PNR/Ticket)", "Passenger", "Description", "Charge", "Credit", "Currency"],
  ];
  rows.forEach((r, idx) => {
    r.srcRow = idx + 1;
    aoa.push([
      r.date || "",
      r.month ? monthLabel(r.month) : "",
      r.reference || "",
      r.paxName || "",
      r.description || "",
      r.charge > 0 ? r.charge : "",
      r.credit > 0 ? r.credit : "",
      r.currency || "",
    ]);
  });
  return aoa;
}

/**
 * Parse a single monthly GDS file from a File object.
 * Returns parsed rows tagged with the month derived from the filename.
 */
/** Read any spreadsheet/CSV file into a 2-D array of cell values. */
async function fileToAoa(file: File): Promise<unknown[][]> {
  const buf = await file.arrayBuffer();
  assertReadableSpreadsheet(buf, file.name);
  if (looksLikePdf(buf)) {
    const { pdfBufToAoa } = await import("./pdf-ledger");
    return pdfBufToAoa(buf, file.name);
  }
  try {
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    const aoa = readBestSheetAoa(wb, { raw: false, defval: "" });
    if (aoa.length) return aoa;
    return textToAoa(buf);
  } catch {
    return textToAoa(buf);
  }
}

/**
 * Detect whether an AoA looks like a Software Entry Report (supplier accounting
 * export) vs a GDS monthly booking export (our side).
 *
 * Identified by markers UNIQUE to that report — "Doc No", "Type of Sales", or the
 * combined "Ticket / Voucher No" column (note the slash). A bare "Voucher No"
 * column is intentionally NOT a trigger: other agency ledgers (e.g. the al-hadaf
 * "our account" narration export) carry a plain "Voucher No" and must keep
 * routing to their own parser instead of being force-read as a supplier statement.
 */
export function isSoftwareEntryReport(aoa: unknown[][]): boolean {
  for (let r = 0; r < Math.min(20, aoa.length); r++) {
    const row = (aoa[r] as unknown[]).map((c) => String(c ?? "").trim().toUpperCase());
    if (
      row.some(
        (h) => h === "DOC NO" || h === "TYPE OF SALES" || /TICKET\s*\/\s*VOUCHER/.test(h),
      )
    ) {
      return true;
    }
  }
  return false;
}

export async function parseMonthlyFile(file: File): Promise<LedgerRow[]> {
  const aoa = await fileToAoa(file);
  if (!aoa.length) return [];
  const monthLabel = monthFromFilename(file.name);
  return parseGDSMonthlyLedger(aoa, monthLabel);
}

/**
 * Parse the software entry report (supplier annual file) from a File object.
 */
export async function parseSoftwareEntryReportFile(file: File): Promise<LedgerRow[]> {
  const aoa = await fileToAoa(file);
  if (!aoa.length) return [];
  return parseSoftwareEntryReport(aoa);
}

/**
 * Does this sheet look like a GDS monthly booking export? It must carry a
 * booking key (PNR or a ticket number) AND an amount column. Anything else is
 * routed to the universal generic parser instead of being forced through the
 * GDS layout (which would otherwise drop every row and look "empty").
 */
export function looksLikeGdsMonthly(aoa: unknown[][]): boolean {
  for (let r = 0; r < Math.min(12, aoa.length); r++) {
    const row = (aoa[r] as unknown[]).map((c) => String(c ?? "").trim().toUpperCase());
    const hasKey = row.some((h) => h === "PNR" || /\bTICKET\b/.test(h));
    const hasAmt = row.some((h) => h === "AMOUNT" || /PAYMENT\s*AMOUNT/.test(h) || h === "FARE");
    if (hasKey && hasAmt) return true;
  }
  return false;
}

/**
 * Auto-detect format and parse ANY Year Mode file — built to be format-agnostic.
 *
 * Detection order, most specific first:
 *   1. Supplier "Software Entry Report" (annual statement)  → parseSoftwareEntryReport
 *   2. GDS monthly booking export (PNR/Ticket + Amount)     → parseGDSMonthlyLedger
 *   3. Anything else                                        → parseGenericLedger
 *
 * The generic parser auto-detects the header row and maps date / amount (or
 * DR-CR) / reference / name / id columns by keyword, so an unfamiliar ledger
 * still reconciles instead of erroring. As a final safety net, if the chosen
 * parser yields no rows we retry with the generic parser. Every row is tagged
 * with a month (filename first, else the row's own date) so the monthly
 * breakdown and the "only the months you uploaded" filter work for any file.
 */
export async function autoParseYearFile(
  file: File,
  side: "ours" | "partner",
): Promise<LedgerRow[]> {
  const aoa = await fileToAoa(file);
  if (!aoa.length) return [];

  const monthLabel = monthFromFilename(file.name);
  const header = (aoa[0] as unknown[]).map((c) => String(c ?? "").trim());
  const upper = header.map((h) => h.toUpperCase());

  let rows: LedgerRow[] = [];
  if (isSoftwareEntryReport(aoa)) {
    rows = parseSoftwareEntryReport(aoa);
  } else if (looksLikeGdsMonthly(aoa)) {
    rows = parseGDSMonthlyLedger(aoa, monthLabel);
  } else if (side === "ours") {
    const hasCode = upper.includes("CODE") && upper.includes("AMOUNT");
    const hasDRCR = upper.includes("DR") && upper.includes("CR");
    if (hasCode) {
      rows = parseOurNarrationStyle(aoa);
    } else if (hasDRCR) {
      rows = parseOurDrCrStyle(aoa);
    } else {
      rows = parseGenericLedger(aoa, "ours");
    }
  } else {
    if (upper.some((h) => /TRANSACTION\s*ID/.test(h)) && upper.some((h) => /PASSENGER/.test(h))) {
      rows = parseMaverickSupplier(aoa, upper);
    } else if (upper.includes("PASSPORT NO.")) {
      rows = parsePartnerFormatA(aoa, upper);
    } else if (upper.includes("ITINERARY") && upper.includes("COMMENTS")) {
      rows = parsePartnerFormatB(aoa, upper);
    } else if (
      upper.includes("RECORD TIME") &&
      upper.includes("TYPE") &&
      upper.includes("COMMENTS") &&
      upper.includes("DESCRIPTION")
    ) {
      rows = parsePartnerFormatC(aoa, upper);
    } else {
      rows = parseGenericLedger(aoa, "partner");
    }
  }

  // Safety net: a known-format guess that produced nothing falls back to generic.
  if (!rows.length) rows = parseGenericLedger(aoa, side);

  // Run the multi-passenger group splitting logic
  const exploded = explodeMultiPax(rows);

  return exploded.map((r) => ({
    ...r,
    side,
    month: r.month || monthLabel || monthKeyFromDate(r.date) || undefined,
  }));
}

/* ------------------------------------------------------------------ */
/* MONTHLY BREAKDOWN ANALYTICS                                          */
/* ------------------------------------------------------------------ */

export type MonthlyBreakdown = {
  month: string;        // "YYYY-MM"
  label: string;        // e.g. "Jan 2026"
  total: number;
  matched: number;
  amountDiff: number;
  onlyOurs: number;
  onlyPartner: number;
  oursTotal: number;    // sum of our amounts in this month
  partnerTotal: number; // sum of partner amounts in this month
  matchRate: number;    // 0–1
};

const MONTH_LABELS: Record<string, string> = {
  "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr", "05": "May", "06": "Jun",
  "07": "Jul", "08": "Aug", "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec",
};

/** "YYYY-MM" from any date string, or "" if unparseable. */
export function monthKeyFromDate(dateStr: string | null | undefined): string {
  if (!dateStr) return "";
  const d = parseDate(dateStr);
  if (isNaN(d)) return "";
  const dt = new Date(d);
  const mo = String(dt.getMonth() + 1).padStart(2, "0");
  return `${dt.getFullYear()}-${mo}`;
}

/** Pretty label for a "YYYY-MM" key, e.g. "Jan 2026". */
export function monthLabel(key: string): string {
  if (!key || key === "unknown") return "Undated";
  const [yr, mo] = key.split("-");
  return `${MONTH_LABELS[mo] ?? mo} ${yr}`;
}

/**
 * Resolve which month a reconciliation pair belongs to, looking at BOTH sides so
 * the per-month filter shows our entries AND supplier entries for that month:
 *   1. our-side explicit `month` tag (from the monthly filename)
 *   2. our-side date
 *   3. partner-side explicit `month` tag (supplier date-derived)
 *   4. partner-side date
 */
export function pairMonth(pair: Pair): string {
  return (
    pair.ours?.month ||
    monthKeyFromDate(pair.ours?.date) ||
    pair.partner?.month ||
    monthKeyFromDate(pair.partner?.date) ||
    "unknown"
  );
}

/**
 * Compute per-month reconciliation breakdown from a set of pairs.
 * Month is derived from either side via {@link pairMonth} so supplier-only rows
 * are counted in the right month too.
 */
export function computeMonthlyBreakdown(pairs: Pair[]): MonthlyBreakdown[] {
  const map = new Map<string, MonthlyBreakdown>();

  const getOrCreate = (key: string): MonthlyBreakdown => {
    if (!map.has(key)) {
      map.set(key, {
        month: key,
        label: monthLabel(key),
        total: 0, matched: 0, amountDiff: 0,
        onlyOurs: 0, onlyPartner: 0,
        oursTotal: 0, partnerTotal: 0,
        matchRate: 0,
      });
    }
    return map.get(key)!;
  };

  // Show ONLY the months that actually carry uploaded data — in CALENDAR-MONTH
  // order (January → December), regardless of which year each falls in (the final
  // sort handles this). Empty months are never shown: upload only January → you
  // see exactly January; a month you didn't upload is NOT padded with a blank
  // 0-card, and undated rows never spawn a stray month. When nothing is dated the
  // strip stays empty and the month selector hides itself.
  for (const pair of pairs) {
    const key = pairMonth(pair);
    // Skip undated pairs (pairMonth → "unknown") so they never spawn a stray
    // blank month card. They still appear in the "all" views, just not as a
    // month in the strip.
    if (!key.includes("-")) continue;
    const bk = getOrCreate(key);
    bk.total++;
    bk.oursTotal += pair.oursAmt ?? 0;
    bk.partnerTotal += pair.partnerAmt ?? 0;

    if (pair.status === "matched") bk.matched++;
    else if (pair.status === "amount_diff") bk.amountDiff++;
    else if (pair.status === "missing_partner") bk.onlyOurs++;
    else if (pair.status === "missing_ours") bk.onlyPartner++;
  }

  for (const bk of map.values()) {
    bk.matchRate = bk.total > 0 ? bk.matched / bk.total : 0;
  }

  // Sort by month-of-year (Jan=1 … Dec=12) so the strip always reads
  // January → December; ties (same month, different year) fall back to the
  // chronological key so e.g. Jan 2025 precedes Jan 2026.
  const monthOf = (k: string) => parseInt(k.split("-")[1] ?? "0", 10);
  return [...map.entries()]
    .sort(([a], [b]) => monthOf(a) - monthOf(b) || a.localeCompare(b))
    .map(([, v]) => v);
}

/* ------------------------------------------------------------------ */
/* FILE ORGANIZER — turn an unorganized ledger into a clean standard   */
/* sheet (Organize Mode). The output columns are exactly the upload    */
/* template's, so an organized file can go straight back into          */
/* Template / Single-File reconciliation.                              */
/* ------------------------------------------------------------------ */

/** Canonical column order of an organized sheet — matches the upload template. */
export const ORGANIZED_HEADERS = [
  "Date",
  "Passport",
  "Passenger Name",
  "Ticket No",
  "Description",
  "Debit",
  "Credit",
  "Currency",
];

export type OrganizeOutcome = {
  rows: LedgerRow[];
  /** Which parser produced the kept rows. */
  engine: "ai" | "heuristic";
  /** 0-based index of the detected header row in the original sheet. */
  headerRow: number;
};

/**
 * Normalize one unorganized sheet into clean ledger rows.
 *
 * Both parsers run: the keyword-driven heuristic (which also finds a header row
 * buried under junk lines) and — when an AI column mapping is supplied — the
 * deterministic dynamic parser on the sheet trimmed to that header row. The
 * result that recovers MORE rows wins (tie → AI, whose semantic mapping is
 * richer). Duplicates are flagged and each row gets a currency: its own
 * detected one, else the statement-level currency.
 */
export function organizeLedgerRows(
  aoa: unknown[][],
  mapping?: ColumnMapping | null,
): OrganizeOutcome {
  if (!aoa.length) return { rows: [], engine: "heuristic", headerRow: 0 };
  const headerRow = detectHeaderRow(aoa);

  const heuristic = parseGenericLedger(aoa, "ours");

  let aiRows: LedgerRow[] = [];
  const hasMapping = !!mapping && Object.values(mapping).some((v) => !!v);
  if (hasMapping) {
    try {
      aiRows = parseDynamicLedger(aoa.slice(headerRow), "ours", mapping!);
      // srcRow came from the trimmed sheet — restore original row numbers.
      for (const r of aiRows) if (typeof r.srcRow === "number") r.srcRow += headerRow;
    } catch {
      aiRows = [];
    }
  }

  // Judge each parse by rows that actually carry data (a date or an amount) so
  // a wrong AI mapping that "reads" every line as empty junk can never win.
  const meaningful = (rs: LedgerRow[]) =>
    rs.filter((r) => r.charge > 0 || r.credit > 0 || !!r.date).length;
  const useAi = meaningful(aiRows) > 0 && meaningful(aiRows) >= meaningful(heuristic);
  const rows = useAi ? aiRows : heuristic;

  flagDuplicates(rows);
  const statementCcy = detectStatementCurrency(aoa);
  for (const r of rows) {
    if (!r.currency) {
      r.currency =
        detectCurrency(r.description, r.reference, r.paxName) ?? statementCcy;
    }
  }
  return { rows, engine: useAi ? "ai" : "heuristic", headerRow };
}

/** Clean AOA (header + body) for organized rows, in ORGANIZED_HEADERS order. */
export function organizedRowsToAoa(rows: LedgerRow[]): (string | number)[][] {
  const aoa: (string | number)[][] = [[...ORGANIZED_HEADERS]];
  for (const r of rows) {
    aoa.push([
      r.date || "",
      r.passport ?? "",
      r.paxName || "",
      r.reference || "",
      r.description || "",
      r.charge > 0 ? r.charge : "",
      r.credit > 0 ? r.credit : "",
      r.currency || "",
    ]);
  }
  return aoa;
}

/** Excel-safe unique sheet name (≤31 chars, no \\ / ? * [ ] :). */
function organizedSheetName(fileName: string, used: Set<string>): string {
  const base =
    fileName
      .replace(/\.[^.]+$/, "")
      .replace(/[\\/?*[\]:]/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 27) || "Sheet";
  let name = base;
  let i = 2;
  while (used.has(name.toUpperCase())) name = `${base.slice(0, 24)} (${i++})`;
  used.add(name.toUpperCase());
  return name;
}

/**
 * Styled workbook of organized files — one clean sheet per input file, plus a
 * combined "All Rows" sheet (with a Source File column) when there are several.
 * Duplicate rows are tinted rose, settlement/top-up rows amber.
 */
export function buildOrganizedWorkbook(
  files: { name: string; rows: LedgerRow[] }[],
): ArrayBuffer {
  const wb = XLSXStyle.utils.book_new();
  const used = new Set<string>();

  const rowStyleFor = (rows: LedgerRow[]) => (i: number): TableRowStyle => {
    const r = rows[i];
    if (!r) return {};
    if ((r.duplicateCount ?? 0) >= 2) return { tint: "FDE8E8" };
    if (r.settlement) return { tint: "FDF3D7", bold: true };
    return {};
  };
  const widths = [12, 14, 26, 18, 38, 12, 12, 10];
  const rightCols = new Set([5, 6]);

  if (files.length > 1) {
    const all: LedgerRow[] = [];
    const srcOf = new Map<LedgerRow, string>();
    for (const f of files)
      for (const r of f.rows) {
        all.push(r);
        srcOf.set(r, f.name);
      }
    const body = all.map((r) => [srcOf.get(r) ?? "", ...organizedRowsToAoa([r])[1]]);
    const ws = makeStyledTable(["Source File", ...ORGANIZED_HEADERS], body, {
      widths: [24, ...widths],
      rightCols: new Set([...rightCols].map((c) => c + 1)),
      rowStyle: rowStyleFor(all),
    });
    XLSXStyle.utils.book_append_sheet(wb, ws, "All Rows");
    used.add("ALL ROWS");
  }

  for (const f of files) {
    const body = organizedRowsToAoa(f.rows).slice(1);
    const ws = makeStyledTable([...ORGANIZED_HEADERS], body, {
      widths,
      rightCols,
      rowStyle: rowStyleFor(f.rows),
    });
    XLSXStyle.utils.book_append_sheet(wb, ws, organizedSheetName(f.name, used));
  }

  return XLSXStyle.write(wb, { bookType: "xlsx", type: "array" }) as ArrayBuffer;
}

/* ------------------------------------------------------------------ */
/* REFERENCE FILES — optional third-party detail sheets (e.g. an       */
/* "Imran report") that carry richer per-passenger context (supplier   */
/* used, nationality, quoted price, remarks) than either ledger. These  */
/* never participate in matching — they're looked up by passport/name  */
/* AFTER reconciliation and attached to the resulting pairs as          */
/* read-only enrichment.                                                */
/* ------------------------------------------------------------------ */

export type ReferenceRow = {
  date: string;
  paxName: string;
  passport: string;
  nationality?: string;
  service?: string;
  supplier?: string;
  company?: string;
  remarks?: string;
  quotedAmount?: number;
  /** Which uploaded reference file this row came from (for multi-file attach). */
  sourceFile?: string;
};

const REF_COLREGEX: Record<string, RegExp> = {
  date: /^date$/i,
  paxName: /^name$/i,
  passport: /passport/i,
  nationality: /national/i,
  service: /service/i,
  supplier: /supplier/i,
  company: /^co\.?$|company/i,
  remarks: /remark|comment|note/i,
  quotedAmount: /quot|price|amount|rate/i,
};

/**
 * Detect whether a batch of ambiguous "A/B/Y" date strings predominantly uses
 * M/D/Y (US-style, common in reference/detail exports) or D/M/Y (the shared
 * `parseDate()` default, tuned for the Gulf-region ledger formats elsewhere in
 * this file) — using only the UNAMBIGUOUS samples in the SAME file (a
 * component >12 can only be the day) — then reformats every date to
 * unambiguous ISO so `parseDate()` never has to guess for dates where both
 * components are ≤12. Scoped to reference files only; the shared ledger
 * parsers and their D/M/Y default are untouched.
 */
function normalizeReferenceDates(rawDates: string[]): string[] {
  const parsed = rawDates.map((s) => {
    const m = String(s ?? "").trim().match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
    return m ? { a: +m[1], b: +m[2], y: +m[3] } : null;
  });
  let mdyVotes = 0;
  let dmyVotes = 0;
  for (const p of parsed) {
    if (!p) continue;
    if (p.a > 12 && p.b <= 12) dmyVotes++;
    else if (p.b > 12 && p.a <= 12) mdyVotes++;
  }
  const preferMdy = mdyVotes > dmyVotes;
  return rawDates.map((s, i) => {
    const p = parsed[i];
    if (!p) return s;
    let yy = p.y;
    if (yy < 100) yy += yy < 50 ? 2000 : 1900;
    let dd: number, mm: number;
    if (p.a > 12 && p.b <= 12) { dd = p.a; mm = p.b; }
    else if (p.b > 12 && p.a <= 12) { dd = p.b; mm = p.a; }
    else if (preferMdy) { mm = p.a; dd = p.b; }
    else { dd = p.a; mm = p.b; }
    if (mm < 1 || mm > 12 || dd < 1 || dd > 31) return s;
    return `${yy}-${String(mm).padStart(2, "0")}-${String(dd).padStart(2, "0")}`;
  });
}

/** Find the header row + column roles for a reference sheet like the Imran report. */
function detectReferenceColumns(
  aoa: unknown[][],
): { headerRow: number; cols: Partial<Record<keyof typeof REF_COLREGEX, number>> } | null {
  for (let r = 0; r < Math.min(15, aoa.length); r++) {
    const row = (aoa[r] as unknown[]).map((c) => String(c ?? "").trim());
    const cols: Partial<Record<keyof typeof REF_COLREGEX, number>> = {};
    row.forEach((h, i) => {
      if (!h) return;
      for (const key of Object.keys(REF_COLREGEX) as (keyof typeof REF_COLREGEX)[]) {
        if (cols[key] == null && REF_COLREGEX[key].test(h)) cols[key] = i;
      }
    });
    // Needs at least a name and a passport column to be usable as a reference sheet.
    if (cols.paxName != null && cols.passport != null) return { headerRow: r, cols };
  }
  return null;
}

/**
 * Parse a standalone reference/detail file (not a DR/CR ledger — no charge or
 * credit columns). Rows are matched to reconciled pairs by passport, so only
 * `passport` + `paxName` are required; every other column is best-effort.
 */
export function parseReferenceFile(aoa: unknown[][], sourceFile?: string): ReferenceRow[] {
  if (!aoa.length) return [];
  const detected = detectReferenceColumns(aoa);
  if (!detected) return [];
  const { headerRow, cols } = detected;
  const get = (row: unknown[], key: keyof typeof REF_COLREGEX): string => {
    const idx = cols[key];
    return idx == null ? "" : String(row[idx] ?? "").trim();
  };

  const out: ReferenceRow[] = [];
  for (let r = headerRow + 1; r < aoa.length; r++) {
    const row = aoa[r] as unknown[];
    if (!row || row.every((c) => c === "" || c == null)) continue;
    const paxName = get(row, "paxName");
    const passportRaw = get(row, "passport");
    const passport = normPassport(passportRaw) ?? passportRaw.toUpperCase();
    if (!paxName && !passport) continue;
    out.push({
      date: get(row, "date"),
      paxName,
      passport,
      nationality: get(row, "nationality") || undefined,
      service: get(row, "service") || undefined,
      supplier: get(row, "supplier") || undefined,
      company: get(row, "company") || undefined,
      remarks: get(row, "remarks") || undefined,
      quotedAmount: cols.quotedAmount != null ? num(row[cols.quotedAmount]) || undefined : undefined,
      sourceFile,
    });
  }

  const normDates = normalizeReferenceDates(out.map((r) => r.date));
  out.forEach((r, i) => { r.date = normDates[i]; });
  return out;
}

/** Parse a reference file from a File object (mirrors parseMonthlyFile / parseSoftwareEntryReportFile). */
export async function parseReferenceFileFromFile(file: File): Promise<ReferenceRow[]> {
  const aoa = await fileToAoa(file);
  return parseReferenceFile(aoa, file.name);
}

/** Extract a day-count visa duration ("30", "60", …) from free text, if present. */
function visaDaysToken(text: string): string | null {
  const m = String(text ?? "").match(/\b(\d{1,3})\s*DAYS?\b/i);
  return m ? m[1] : null;
}

/**
 * Correct ticket-embedded passports for "3VS XXXXX" style ticket numbers.
 * `extractPassportFromTicket` (used by `parseOurDrCrStyle`) always strips the
 * trailing character as a check digit — correct for some suppliers (verified:
 * Mavrick's "3VS V41005051" → real passport "V4100505") but WRONG for others
 * whose ticket suffix already IS the complete passport with no check digit
 * (verified: AL Saffron's "3VS A38535385" → real passport "A38535385", the
 * full 9 characters). This is inherently ambiguous from ticket text alone —
 * both conventions produce identically-shaped 1-letter+8-digit strings. When a
 * Reference File is attached, use it as ground truth: if the FULL
 * (un-truncated) ticket suffix matches a known reference passport but the
 * truncated default doesn't, correct the row. No-op whenever no reference
 * file is attached, so the existing (Mavrick-verified) default behavior is
 * completely unchanged without one — zero regression risk.
 */
export function correctTicketPassportsFromReference(rows: LedgerRow[], refRows: ReferenceRow[]): void {
  if (!refRows.length) return;
  const refPassports = new Set(
    refRows.map((r) => normPassport(r.passport)).filter((p): p is string => !!p),
  );
  for (const row of rows) {
    // A group booking's sub-rows all carry the SAME embedded ticket text
    // (explodeMultiPax spreads the original row into each sub-row) — only the
    // lead sub-row (paxIndex 1) is meant to hold that ticket's passport; the
    // rest are deliberately left blank for enrichGroupRowsFromReference to
    // fill in per-passenger below. Without this guard every sub-row got
    // stamped with the SAME lead passport, which then blocked the per-
    // passenger fill (it only fills an already-empty passport).
    if (row.raw?.isGroupRow && row.raw?.paxIndex !== 1) continue;
    const ticket = typeof row.raw?.ticket === "string" ? (row.raw.ticket as string) : "";
    const m = ticket.trim().match(/^3VS\s*([A-Z0-9]+)$/i);
    if (!m) continue;
    const full = normPassport(m[1]);
    const truncated = normPassport(m[1].slice(0, -1));
    if (full && refPassports.has(full) && !(truncated && refPassports.has(truncated))) {
      row.passport = full;
    }
  }
}

/**
 * Fill in real passenger identity (name + passport) for anonymous multi-pax
 * group-booking sub-rows — e.g. "03 PAX 60 DAYS VISA" split by `explodeMultiPax`
 * into 3 rows that only carry the lead passenger's passport — using an attached
 * Reference File that lists each passenger individually (date, name, passport,
 * service). Never touches amounts or overwrites an already-known passport;
 * mutates `rows` in place. Prefers same-day + same-visa-duration candidates
 * when there are enough of them; falls back to the broader same-day pool
 * otherwise. Best-effort and partial: if the reference file has fewer
 * candidates than the group size, it still fills as many sub-rows as it can
 * instead of filling none — a partially-identified group still improves match
 * confidence for the rows it does cover.
 */
export function enrichGroupRowsFromReference(rows: LedgerRow[], refRows: ReferenceRow[]): void {
  if (!refRows.length) return;

  // Bucket reference rows by calendar day (via parseDate, so "9/8/25" and
  // "2025-09-08" compare equal regardless of source format).
  // Bucket by LOCAL calendar day, not raw epoch ms: a reference row whose date
  // string falls through to the Date.parse() fallback (e.g. a full
  // "Tue Sep 09 2025 23:59:48 GMT+0400 (...)" stringified JS Date) keeps its
  // time-of-day, while a plain "2025-09-09" ledger date parses to local
  // midnight — same calendar day, different exact millisecond, so raw-epoch
  // keys silently never matched. (A naive `Math.floor(ms / 86400000)` day-index
  // is NOT a fix here — it buckets by UTC day, so in any timezone ahead of UTC
  // a local midnight and a same-day 23:59 reference timestamp land on
  // DIFFERENT UTC-day indices, aliasing the wrong day instead. Reading the
  // local Y/M/D back off the Date is immune to that: both representations
  // resolve to the same local calendar day, which is what "same day" means to
  // every date string in this file.)
  const dayIndex = (ms: number) => {
    const d = new Date(ms);
    return d.getFullYear() * 10000 + d.getMonth() * 100 + d.getDate();
  };
  const byDay = new Map<number, ReferenceRow[]>();
  for (const ref of refRows) {
    const d = parseDate(ref.date);
    if (Number.isNaN(d)) continue;
    const k = dayIndex(d);
    (byDay.get(k) ?? byDay.set(k, []).get(k)!).push(ref);
  }
  const used = new Set<ReferenceRow>();

  let i = 0;
  while (i < rows.length) {
    const r = rows[i];
    const n = r.raw?.isGroupRow ? (r.raw.paxCount as number) : 0;
    if (n >= 2 && r.raw?.paxIndex === 1) {
      // Sub-rows for one group booking are pushed consecutively by explodeMultiPax.
      const block: LedgerRow[] = [];
      let j = i;
      while (
        j < rows.length &&
        block.length < n &&
        rows[j].raw?.isGroupRow &&
        rows[j].raw?.explodedFrom === r.raw!.explodedFrom &&
        rows[j].raw?.paxCount === n
      ) {
        block.push(rows[j]);
        j++;
      }
      if (block.length === n) {
        // Snapshot the shared generic label BEFORE any mutation — block[0] is
        // the same object as `r`, so writing block[0].paxName below would
        // otherwise corrupt this comparison for every later k in the loop.
        const genericPaxName = r.paxName;
        const day = dayIndex(parseDate(r.date));
        const dayCandidates = (byDay.get(day) ?? []).filter((ref) => !used.has(ref));
        const days = visaDaysToken(`${r.paxName} ${r.description}`);
        const filtered = days
          ? dayCandidates.filter((ref) => visaDaysToken(`${ref.service ?? ""}`) === days)
          : dayCandidates;
        // Prefer a pool that can fill the WHOLE group; when neither pool is big
        // enough, still prefer the more precise (duration-filtered) pool for a
        // partial fill over the broader unfiltered one.
        const candidates =
          filtered.length >= n
            ? filtered
            : dayCandidates.length >= n
              ? dayCandidates
              : filtered.length
                ? filtered
                : dayCandidates;
        // A sub-row can already carry a known passport (e.g. the lead sub-row,
        // corrected earlier from its ticket text by correctTicketPassportsFromReference)
        // — pair it with the candidate that actually HAS that passport rather than
        // whichever candidate happens to sit at the same array index, or its name
        // gets overwritten with an unrelated person's while the passport stays
        // correct, silently mismatching the two.
        const remaining = candidates.slice();
        for (const sub of block) {
          if (!sub.passport) continue;
          const idx = remaining.findIndex((ref) => normPassport(ref.passport) === sub.passport);
          if (idx < 0) continue;
          const [ref] = remaining.splice(idx, 1);
          used.add(ref);
          if (sub.paxName === genericPaxName) sub.paxName = ref.paxName || sub.paxName;
        }
        // Fill the still-anonymous sub-rows positionally from whatever's left.
        let ri = 0;
        for (const sub of block) {
          if (sub.passport) continue;
          const ref = remaining[ri++];
          if (!ref) break;
          used.add(ref);
          sub.passport = normPassport(ref.passport);
          if (sub.paxName === genericPaxName) sub.paxName = ref.paxName || sub.paxName;
        }
      }
      i = j;
      continue;
    }
    i++;
  }
}

/**
 * Attach reference-file enrichment to already-reconciled pairs, purely by
 * passport (falling back to name similarity when no passport is available on
 * either side). Never changes match status — read-only annotation.
 */
export function attachReferenceData(pairs: Pair[], refRows: ReferenceRow[]): Pair[] {
  if (!refRows.length) return pairs;

  const byPassport = new Map<string, ReferenceRow>();
  // Blocking index for the fuzzy passport fallback (first 4 chars, same
  // scheme as matchSet's byPassPrefix) — avoids an O(pairs × refRows) scan.
  const byPassPrefix = new Map<string, ReferenceRow[]>();
  // Blocking index for the name fallback (first 3-char-plus token).
  const byNameTok = new Map<string, ReferenceRow[]>();
  for (const ref of refRows) {
    const key = normPassport(ref.passport);
    if (key) {
      if (!byPassport.has(key)) byPassport.set(key, ref);
      const prefix = key.slice(0, 4);
      (byPassPrefix.get(prefix) ?? byPassPrefix.set(prefix, []).get(prefix)!).push(ref);
    }
    for (const tok of ref.paxName.toUpperCase().split(/\s+/).filter((t) => t.length >= 3).slice(0, 3)) {
      (byNameTok.get(tok) ?? byNameTok.set(tok, []).get(tok)!).push(ref);
    }
  }

  for (const p of pairs) {
    const passport = p.ours?.passport ?? p.partner?.passport ?? null;
    const name = p.ours?.paxName ?? p.partner?.paxName ?? "";

    let best: ReferenceRow | undefined;
    const key = normPassport(passport);
    if (key && byPassport.has(key)) {
      best = byPassport.get(key);
    } else if (passport && key) {
      // Fuzzy passport match (embedded/prefix forms, e.g. ticket-derived passports)
      // — only against reference rows sharing the same 4-char prefix bucket.
      const candidates = byPassPrefix.get(key.slice(0, 4)) ?? [];
      let bestScore = 0;
      for (const ref of candidates) {
        const s = passportMatch(passport, ref.passport);
        if (s > bestScore) {
          bestScore = s;
          best = ref;
        }
      }
      if (bestScore < 0.78) best = undefined;
    }
    if (!best && name) {
      const tokens = name.toUpperCase().split(/\s+/).filter((t) => t.length >= 3);
      const candidates = new Set<ReferenceRow>();
      for (const tok of tokens) for (const ref of byNameTok.get(tok) ?? []) candidates.add(ref);
      let bestScore = 0;
      for (const ref of candidates) {
        const s = nameSimilarity(name, ref.paxName);
        if (s > bestScore) {
          bestScore = s;
          best = ref;
        }
      }
      if (bestScore < 0.75) best = undefined;
    }
    if (best) p.refInfo = best;
  }
  return pairs;
}
