import { describe, it, expect } from "vitest";
import * as XLSX from "xlsx";
import {
  applyFxMatchAccuracy,
  computeTotals,
  organizeLedgerRows,
  organizedRowsToAoa,
  ORGANIZED_HEADERS,
  parseGenericLedger,
  monthKeyFromDate,
  readBestSheetAoa,
  textToAoa,
  looksLikePdf,
  explodeMultiPax,
  enrichGroupRowsFromReference,
  referenceMatch,
  parseSoftwareEntryReport,
  correctTicketPassportsFromReference,
  reconcile,
  type LedgerRow,
  type Pair,
  type ReconResult,
  type ReferenceRow,
} from "./reconcile";
import { pdfItemsToAoa, type PdfTextItem } from "./pdf-ledger";

/* ---------- minimal fixtures ---------- */
const row = (o: Partial<LedgerRow>): LedgerRow =>
  ({
    side: "ours",
    index: 0,
    date: "2026-01-01",
    passport: null,
    paxName: "",
    description: "",
    reference: "",
    charge: 0,
    credit: 0,
    kind: "charge",
    raw: {},
    ...o,
  }) as LedgerRow;

const refRow = (o: Partial<ReferenceRow>): ReferenceRow => ({
  date: "2026-01-10",
  paxName: "",
  passport: "",
  ...o,
});

const pair = (o: Partial<Pair>): Pair => ({
  key: "k",
  status: "matched",
  kind: "charge",
  ours: null,
  partner: null,
  oursAmt: 0,
  partnerAmt: 0,
  diff: 0,
  note: "",
  ...o,
});

const result = (pairs: Pair[]): ReconResult => {
  const ours = pairs.map((p) => p.ours).filter(Boolean) as LedgerRow[];
  const partner = pairs.map((p) => p.partner).filter(Boolean) as LedgerRow[];
  return { pairs, totals: computeTotals(ours, partner, pairs) };
};

/* ---------- exact, currency-aware matching ---------- */
describe("applyFxMatchAccuracy", () => {
  it("keeps a pair matched when converted amounts are exactly equal", () => {
    const p = pair({
      status: "matched",
      ours: row({ charge: 375 }),
      partner: row({ side: "partner", charge: 100, currency: "USD" }),
      oursAmt: 375,
      partnerAmt: 100,
    });
    const out = applyFxMatchAccuracy(result([p]), { active: true, rate: 3.75 });
    expect(out.pairs[0].status).toBe("matched");
    expect(out.totals.matched).toBe(1);
  });

  it("tolerates a small gap explainable by a manually-typed rate's rounding", () => {
    // A real (non-pegged) rate typed to 4 decimals, e.g. AED/SAR, leaves a
    // sub-cent-scale gap at real transaction sizes purely from that rounding
    // — this must still count as matched, or currency conversion is useless
    // for any pair without a perfectly round peg.
    const p = pair({
      status: "matched",
      ours: row({ charge: 375.1 }),
      partner: row({ side: "partner", charge: 100, currency: "USD" }),
      oursAmt: 375.1,
      partnerAmt: 100,
    });
    const out = applyFxMatchAccuracy(result([p]), { active: true, rate: 3.75 });
    expect(out.pairs[0].status).toBe("matched");
    expect(out.totals.matched).toBe(1);
  });

  it("marks a pair as amount_diff when the gap is a genuine discrepancy, not rate rounding", () => {
    const p = pair({
      status: "matched",
      ours: row({ charge: 390 }), // ~4% off — far beyond any plausible rate-rounding gap
      partner: row({ side: "partner", charge: 100, currency: "USD" }),
      oursAmt: 390,
      partnerAmt: 100,
    });
    const out = applyFxMatchAccuracy(result([p]), { active: true, rate: 3.75 });
    expect(out.pairs[0].status).toBe("amount_diff");
    expect(out.totals.matched).toBe(0);
    expect(out.totals.amountIssues).toBe(1);
  });

  it("tolerates sub-cent rounding (0.004 rounds to 0 → still matched)", () => {
    const p = pair({
      status: "matched",
      ours: row({ charge: 3750 }),
      partner: row({ side: "partner", charge: 1000, currency: "USD" }),
      oursAmt: 3750,
      partnerAmt: 1000.001, // ×3.75 = 3750.00375 → diff 0.00 after 2dp
    });
    const out = applyFxMatchAccuracy(result([p]), { active: true, rate: 3.75 });
    expect(out.pairs[0].status).toBe("matched");
  });

  it("returns the input unchanged when conversion is inactive", () => {
    const r = result([
      pair({
        status: "matched",
        ours: row({ charge: 100 }),
        partner: row({ side: "partner", charge: 100 }),
        oursAmt: 100,
        partnerAmt: 100,
      }),
    ]);
    expect(applyFxMatchAccuracy(r, { active: false, rate: 3.75 })).toBe(r);
    expect(applyFxMatchAccuracy(r, { active: true, rate: 0 })).toBe(r);
  });

  it("leaves only-ours / only-partner pairs untouched", () => {
    const p = pair({
      status: "missing_partner",
      ours: row({ charge: 100 }),
      partner: null,
      oursAmt: 100,
      partnerAmt: 0,
    });
    const out = applyFxMatchAccuracy(result([p]), { active: true, rate: 3.75 });
    expect(out.pairs[0].status).toBe("missing_partner");
    expect(out.totals.onlyOurs).toBe(1);
  });
});

/* ---------- totals ---------- */
describe("computeTotals", () => {
  it("counts each status bucket", () => {
    const pairs = [
      pair({ status: "matched", ours: row({ charge: 100 }), partner: row({ side: "partner", charge: 100 }) }),
      pair({ status: "amount_diff", ours: row({ charge: 100 }), partner: row({ side: "partner", charge: 90 }) }),
      pair({ status: "missing_partner", ours: row({ charge: 100 }) }),
      pair({ status: "missing_ours", partner: row({ side: "partner", charge: 100 }) }),
    ];
    const t = computeTotals(
      pairs.map((p) => p.ours).filter(Boolean) as LedgerRow[],
      pairs.map((p) => p.partner).filter(Boolean) as LedgerRow[],
      pairs,
    );
    expect(t.matched).toBe(1);
    expect(t.amountIssues).toBe(1);
    expect(t.onlyOurs).toBe(1);
    expect(t.onlyPartner).toBe(1);
  });
});

/* ---------- Organize mode: messy sheet → clean standard rows ---------- */
describe("organizeLedgerRows", () => {
  // A typical unorganized export: junk title lines, blank row, buried header.
  const messy: unknown[][] = [
    ["MY COMPANY LLC", "", "", "", ""],
    ["Statement of Account — 2026", "", "", "", ""],
    ["", "", "", "", ""],
    ["Date", "Passport No", "Client Name", "Debit", "Credit"],
    ["2026-01-05", "A1234567", "JOHN SMITH", "1500", ""],
    ["2026-01-06", "B7654321", "JANE DOE", "1500", ""],
    ["2026-01-10", "", "BANK TRANSFER RECEIVED", "", "3000"],
  ];

  it("finds the buried header row and parses only the data rows", () => {
    const { rows, headerRow } = organizeLedgerRows(messy);
    expect(headerRow).toBe(3);
    expect(rows.length).toBe(3);
    expect(rows[0].passport).toBe("A1234567");
    expect(rows[0].charge).toBe(1500);
    expect(rows[2].credit).toBe(3000);
  });

  it("uses the AI mapping when it recovers at least as many data rows", () => {
    const { engine, rows } = organizeLedgerRows(messy, {
      date: "Date",
      passport: "Passport No",
      paxName: "Client Name",
      charge: "Debit",
      credit: "Credit",
    });
    expect(engine).toBe("ai");
    expect(rows.length).toBe(3);
    expect(rows[1].paxName).toBe("JANE DOE");
  });

  it("falls back to the heuristic parser when the mapping matches nothing", () => {
    const { engine, rows } = organizeLedgerRows(messy, { date: "ZZZZZZ", charge: "QQQQQQ" });
    expect(engine).toBe("heuristic");
    expect(rows.length).toBe(3);
  });

  it("returns no rows for an empty sheet", () => {
    expect(organizeLedgerRows([]).rows).toEqual([]);
  });
});

describe("organizedRowsToAoa", () => {
  it("emits the standard template header plus one line per row", () => {
    const { rows } = organizeLedgerRows([
      ["Date", "Passport No", "Client Name", "Debit", "Credit"],
      ["2026-01-05", "A1234567", "JOHN SMITH", "1500", ""],
    ]);
    const aoa = organizedRowsToAoa(rows);
    expect(aoa[0]).toEqual(ORGANIZED_HEADERS);
    expect(aoa.length).toBe(rows.length + 1);
    expect(aoa[1][1]).toBe("A1234567"); // Passport column
    expect(aoa[1][5]).toBe(1500); // Debit column
  });
});

/* ---------- universal-ledger hardening ---------- */
describe("parseGenericLedger — universal formats", () => {
  it("reads a DR/CR indicator column next to a single Amount column", () => {
    const rows = parseGenericLedger(
      [
        ["Date", "Particulars", "Ref", "Amount", "Ind"],
        ["05/01/2026", "VISA FEE JOHN SMITH", "V-1", "1,500.00", "DR"],
        ["06/01/2026", "PAYMENT RECEIVED", "P-1", "3,000.00", "CR"],
        ["07/01/2026", "VISA FEE JANE DOE", "V-2", "1,500.00", "DR"],
        ["08/01/2026", "REFUND ISSUED", "R-1", "250.00", "CR"],
      ],
      "ours",
    );
    expect(rows.length).toBe(4);
    expect(rows[0].charge).toBe(1500);
    expect(rows[1].credit).toBe(3000);
    expect(rows[3].credit).toBe(250);
  });

  it("skips TOTAL / Closing Balance summary lines", () => {
    const rows = parseGenericLedger(
      [
        ["Date", "Description", "Debit", "Credit"],
        ["2026-01-05", "Visa fee", "1500", ""],
        ["2026-01-06", "Top-up received", "", "3000"],
        ["", "TOTAL", "1500", "3000"],
        ["", "Closing Balance", "", "1500"],
        ["", "Balance C/F", "", "1500"],
      ],
      "ours",
    );
    expect(rows.length).toBe(2);
    expect(rows[0].charge).toBe(1500);
  });

  it("parses accounting negatives, currency prefixes and exotic thousands", () => {
    const rows = parseGenericLedger(
      [
        ["Date", "Description", "Debit", "Credit"],
        ["2026-01-05", "Accounting negative", "(1,500.00)", ""], // → charge 1500
        ["2026-01-06", "Currency prefix", "SAR 2,250.50", ""],
        ["2026-01-07", "European format", "", "1.234,56"],
        ["2026-01-08", "Apostrophe thousands", "1'000", ""],
        ["2026-01-09", "Arabic-Indic digits", "٧٥٠", ""],
        ["2026-01-10", "Trailing minus", "", "500-"], // negative credit → credit 500
      ],
      "ours",
    );
    expect(rows[0].charge).toBe(1500); // (1,500.00) in a debit column
    expect(rows[1].charge).toBe(2250.5);
    expect(rows[2].credit).toBe(1234.56);
    expect(rows[3].charge).toBe(1000);
    expect(rows[4].charge).toBe(750);
    expect(rows[5].credit).toBe(500);
  });

  it("normalizes Excel serial dates to ISO for display and month grouping", () => {
    const rows = parseGenericLedger(
      [
        ["Date", "Description", "Debit", "Credit"],
        [46027, "Visa fee", "1500", ""], // serial 46027 = 2026-01-05
      ],
      "ours",
    );
    expect(rows[0].date).toBe("2026-01-05");
    expect(monthKeyFromDate(rows[0].date)).toBe("2026-01");
  });
});

describe("parseDate via monthKeyFromDate", () => {
  it("reads compact YYYYMMDD dates", () => {
    expect(monthKeyFromDate("20260105")).toBe("2026-01");
  });
  it("reads Arabic-Indic digit dates", () => {
    expect(monthKeyFromDate("٠٥/٠١/٢٠٢٦")).toBe("2026-01");
  });
});

describe("parseGenericLedger — international & messy headers", () => {
  it("reads Arabic column headers (التاريخ / مدين / دائن)", () => {
    const rows = parseGenericLedger(
      [
        ["التاريخ", "البيان", "مدين", "دائن"],
        ["05/01/2026", "رسوم تأشيرة", "1500", ""],
        ["06/01/2026", "دفعة مستلمة", "", "3000"],
      ],
      "ours",
    );
    expect(rows.length).toBe(2);
    expect(rows[0].charge).toBe(1500);
    expect(rows[1].credit).toBe(3000);
  });

  it("reads accented French headers (Débit / Crédit)", () => {
    const rows = parseGenericLedger(
      [
        ["Date", "Libellé Description", "Débit", "Crédit"],
        ["05/01/2026", "FRAIS VISA", "1500", ""],
        ["06/01/2026", "PAIEMENT", "", "3000"],
      ],
      "ours",
    );
    expect(rows[0].charge).toBe(1500);
    expect(rows[1].credit).toBe(3000);
  });

  it("merges two-row split headers (Amount over Debit/Credit)", () => {
    const rows = parseGenericLedger(
      [
        ["Date", "Description", "Amount", ""],
        ["", "", "Debit", "Credit"],
        ["2026-01-05", "Visa fee", "1500", ""],
        ["2026-01-06", "Top-up", "", "3000"],
      ],
      "ours",
    );
    expect(rows.length).toBe(2);
    expect(rows[0].charge).toBe(1500);
    expect(rows[1].credit).toBe(3000);
  });

  it("skips print-style repeated header rows inside the data", () => {
    const rows = parseGenericLedger(
      [
        ["Date", "Description", "Debit", "Credit"],
        ["2026-01-05", "Visa fee", "1500", ""],
        ["Date", "Description", "Debit", "Credit"], // page 2 header
        ["2026-02-05", "Visa fee", "1500", ""],
      ],
      "ours",
    );
    expect(rows.length).toBe(2);
  });
});

describe("textToAoa — HTML .xls and exotic delimiters", () => {
  const toBuf = (s: string): ArrayBuffer => new TextEncoder().encode(s).buffer as ArrayBuffer;

  it("parses an HTML table saved as .xls (old ERP exports)", () => {
    const html = `<html><body><table>
      <tr><td>Date</td><td>Description</td><td>Debit</td><td>Credit</td></tr>
      <tr><td>2026-01-05</td><td>Visa fee</td><td>1500</td><td></td></tr>
      <tr><td>2026-01-06</td><td>Top-up</td><td></td><td>3000</td></tr>
    </table></body></html>`;
    const aoa = textToAoa(toBuf(html));
    expect(String(aoa[0][0])).toBe("Date");
    const rows = parseGenericLedger(aoa, "ours");
    expect(rows.length).toBe(2);
    expect(rows[0].charge).toBe(1500);
  });

  it("auto-detects semicolon-delimited CSVs (European exports)", () => {
    const csv = "Date;Description;Debit;Credit\r\n2026-01-05;Visa fee;1500;\r\n2026-01-06;Top-up;;3000";
    const aoa = textToAoa(toBuf(csv));
    expect(aoa[0]).toEqual(["Date", "Description", "Debit", "Credit"]);
    const rows = parseGenericLedger(aoa, "ours");
    expect(rows[0].charge).toBe(1500);
    expect(rows[1].credit).toBe(3000);
  });

  it("auto-detects pipe-delimited files", () => {
    const txt = "Date|Description|Debit|Credit\n2026-01-05|Visa fee|1500|";
    const aoa = textToAoa(toBuf(txt));
    expect(aoa[0]).toEqual(["Date", "Description", "Debit", "Credit"]);
  });
});

describe("PDF statements", () => {
  it("detects PDF magic bytes", () => {
    expect(looksLikePdf(new TextEncoder().encode("%PDF-1.7 ...").buffer as ArrayBuffer)).toBe(true);
    expect(looksLikePdf(new TextEncoder().encode("PK").buffer as ArrayBuffer)).toBe(false);
  });

  it("rebuilds a table from positioned PDF text fragments", () => {
    // Simulated bank-statement layout: header line + 2 data lines. The Debit
    // column is RIGHT-aligned (x varies) — coverage-based bands must still
    // put all amounts in one column.
    const line = (y: number, cells: [string, number, number][]): PdfTextItem[] =>
      cells.map(([str, x, w]) => ({ str, x, y, w }));
    const page: PdfTextItem[] = [
      ...line(700, [["Date", 40, 30], ["Description", 120, 70], ["Debit", 320, 35], ["Credit", 400, 40]]),
      ...line(680, [["05/01/2026", 40, 60], ["VISA FEE JOHN", 120, 90], ["1,500.00", 305, 50]]),
      ...line(660, [["06/01/2026", 40, 60], ["PAYMENT RECEIVED", 120, 110], ["3,000.00", 392, 48]]),
    ];
    const aoa = pdfItemsToAoa([page]);
    expect(aoa.length).toBe(3);
    const rows = parseGenericLedger(aoa, "ours");
    expect(rows.length).toBe(2);
    expect(rows[0].charge).toBe(1500);
    expect(rows[1].credit).toBe(3000);
  });
});

describe("readBestSheetAoa", () => {
  it("picks the data sheet when a cover sheet comes first", () => {
    const wb = XLSX.utils.book_new();
    const cover = XLSX.utils.aoa_to_sheet([["Statement of Account"], ["Prepared 2026"]]);
    const data = XLSX.utils.aoa_to_sheet([
      ["Date", "Description", "Debit", "Credit"],
      ["2026-01-05", "Visa fee", 1500, ""],
      ["2026-01-06", "Top-up", "", 3000],
      ["2026-01-07", "Visa fee", 1500, ""],
    ]);
    XLSX.utils.book_append_sheet(wb, cover, "Cover");
    XLSX.utils.book_append_sheet(wb, data, "Ledger");
    const aoa = readBestSheetAoa(wb, { defval: "" });
    expect(String(aoa[0][0])).toBe("Date");
    expect(aoa.length).toBe(4);
  });

  it("keeps the first sheet for normal single-sheet files", () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.aoa_to_sheet([
        ["Date", "Debit"],
        ["2026-01-05", 100],
      ]),
      "Sheet1",
    );
    const aoa = readBestSheetAoa(wb, { defval: "" });
    expect(String(aoa[0][0])).toBe("Date");
  });
});

/* ---------- many-to-one group-passenger matching ---------- */
describe("explodeMultiPax", () => {
  it("splits an anonymous 'N PAX' group row into N sub-rows, keeping the lead passport", () => {
    const [out] = [
      explodeMultiPax([
        row({ paxName: "04 PAX 60 DAYS VISA", description: "Visa fee", charge: 4000, passport: "P1000000" }),
      ]),
    ];
    expect(out.length).toBe(4);
    expect(out[0].passport).toBe("P1000000");
    expect(out.slice(1).every((r) => r.passport === null)).toBe(true);
    expect(out.every((r) => r.raw?.isGroupRow)).toBe(true);
    expect(out.every((r) => r.raw?.paxCount === 4)).toBe(true);
    expect(+out.reduce((s, r) => s + r.charge, 0).toFixed(2)).toBe(4000);
  });

  it("recognizes 'PAX: N' and 'N PASSENGERS' and 'GROUP OF N' phrasing, not just 'N PAX'", () => {
    const colon = explodeMultiPax([
      row({ paxName: "", description: "PAX: 03 60 DAYS VISA", charge: 3000 }),
    ]);
    expect(colon.length).toBe(3);

    const passengers = explodeMultiPax([
      row({ paxName: "", description: "4 PASSENGERS 90 DAYS VISA", charge: 4400 }),
    ]);
    expect(passengers.length).toBe(4);

    const groupOf = explodeMultiPax([
      row({ paxName: "", description: "GROUP OF 5 VISA RENEWAL", charge: 5500 }),
    ]);
    expect(groupOf.length).toBe(5);
  });

  it("leaves single-passenger and settlement rows untouched", () => {
    const single = explodeMultiPax([row({ paxName: "John Doe", description: "Visa fee", charge: 1000 })]);
    expect(single.length).toBe(1);
    expect(single[0].raw?.isGroupRow).toBeUndefined();

    const settlement = explodeMultiPax([
      row({ paxName: "03 PAX 60 DAYS VISA", description: "Visa fee", charge: 3000, settlement: true as never }),
    ]);
    expect(settlement.length).toBe(1);
  });
});

describe("referenceMatch — PNR sequence-suffix tolerance", () => {
  it("matches a GDS bare PNR against a supplier PNR with a 1-digit invoice-line suffix", () => {
    const gds = row({
      side: "ours",
      scenario: "flight" as never,
      reference: "18GR56",
      description: "15/10/2025 RQ-957-L KBL-JED PNR:18GR56",
    });
    const supplier = row({
      side: "partner",
      scenario: "flight" as never,
      reference: "18GR561",
      description: "Economy Class REF:18GR561 SAMIR",
    });
    expect(referenceMatch(gds, supplier)).toBe(1);
  });

  it("does not fuzz-match two genuinely different 6-char PNRs", () => {
    const gds = row({ side: "ours", reference: "18GR56", description: "PNR:18GR56" });
    const supplier = row({ side: "partner", reference: "19ZZ99", description: "REF:19ZZ99" });
    expect(referenceMatch(gds, supplier)).toBe(0);
  });

  it("matches a PNR that differs only by letter-O vs digit-0 transcription", () => {
    const gds = row({ side: "ours", reference: "18PG0N", description: "PNR:18PG0N" });
    const supplier = row({ side: "partner", reference: "18PGON", description: "REF:18PGON" });
    expect(referenceMatch(gds, supplier)).toBe(1);
  });
});

describe("parseSoftwareEntryReport — PNR label typo tolerance", () => {
  const header = [
    "Date", "Doc No", "Ticket / Voucher No", "Type of Sales", "Pax", "Class",
    "Travel Date", "Debit", "Credit", "Balance", "PNR / Reference", "Narration",
  ];
  const aoa = (pnrCell: string) => [
    header,
    ["29/01/2026", "INV26010500", "", "INVOICE", "SHARAFAT KHAN", "Economy", "", 0, 240, 0, pnrCell, ""],
  ];

  it("strips the standard 'PNR' label", () => {
    const rows = parseSoftwareEntryReport(aoa("PNR19U9RA"));
    expect(rows[0].raw?.pnrRef).toBe("19U9RA");
  });

  it("strips typo'd label variants ('PNE', 'PNRF', 'PNER') since every real PNR starts with a digit", () => {
    expect(parseSoftwareEntryReport(aoa("PNE19U9RA"))[0].raw?.pnrRef).toBe("19U9RA");
    expect(parseSoftwareEntryReport(aoa("PNRF19U9RA"))[0].raw?.pnrRef).toBe("19U9RA");
    expect(parseSoftwareEntryReport(aoa("PNER19U9RA"))[0].raw?.pnrRef).toBe("19U9RA");
  });

  it("does not strip a real code that just happens to start with 'PN' but isn't a label typo", () => {
    // Doesn't match PNR_PREFIX because no digit immediately follows the letter run.
    const rows = parseSoftwareEntryReport(aoa("PNXABCDEF"));
    expect(rows[0].raw?.pnrRef).toBe("PNXABCDEF");
  });
});

describe("correctTicketPassportsFromReference — group-row guard", () => {
  it("only corrects the lead sub-row of a group, leaving the rest blank for per-passenger enrichment", () => {
    // explodeMultiPax spreads the SAME embedded ticket text onto every sub-row of
    // a group booking; only the lead (paxIndex 1) is meant to carry a passport —
    // the others must stay null so enrichGroupRowsFromReference can fill them in
    // per-passenger, not get stamped with the lead's passport first.
    const rows = explodeMultiPax([
      row({
        paxName: "03 PAX 60 DAYS VISA",
        description: "Visa fee",
        charge: 3000,
        raw: { ticket: "3VS P1000000" },
      }),
    ]);
    const refRows: ReferenceRow[] = [refRow({ paxName: "Lead", passport: "P1000000" })];
    correctTicketPassportsFromReference(rows, refRows);
    expect(rows[0].passport).toBe("P1000000");
    expect(rows[1].passport).toBeNull();
    expect(rows[2].passport).toBeNull();
  });
});

describe("enrichGroupRowsFromReference", () => {
  it("fills every sub-row when the reference file has a full same-day pool", () => {
    const rows = explodeMultiPax([
      row({
        paxName: "03 PAX 60 DAYS VISA",
        description: "Visa fee",
        charge: 3000,
        date: "2026-01-10",
        passport: "P1000000",
      }),
    ]);
    const refRows: ReferenceRow[] = [
      // Lead sub-row's passport (already known, e.g. from ticket-text correction)
      // matches Alice specifically — the fill must pair by THAT identity, not by
      // array position, or the lead's name and passport end up referring to two
      // different people.
      refRow({ paxName: "Alice", passport: "P1000000", service: "60 DAYS VISA" }),
      refRow({ paxName: "Bob", passport: "P2000002", service: "60 DAYS VISA" }),
      refRow({ paxName: "Carol", passport: "P2000003", service: "60 DAYS VISA" }),
    ];
    enrichGroupRowsFromReference(rows, refRows);
    // Lead passport (already known) is never overwritten, but its generic name is
    // filled from the candidate that actually carries that same passport.
    expect(rows[0].passport).toBe("P1000000");
    expect(rows[0].paxName).toBe("Alice");
    expect(rows[1].passport).toBe("P2000002");
    expect(rows[1].paxName).toBe("Bob");
    expect(rows[2].passport).toBe("P2000003");
    expect(rows[2].paxName).toBe("Carol");
  });

  it("does not mismatch a sub-row's known passport to an unrelated candidate's name", () => {
    const rows = explodeMultiPax([
      row({
        paxName: "02 PAX 60 DAYS VISA",
        description: "Visa fee",
        charge: 2000,
        date: "2026-01-10",
        passport: "P1000000",
      }),
    ]);
    // Neither candidate carries the lead's passport P1000000 — the lead must stay
    // anonymous rather than borrowing Alice's name just because she's listed first.
    const refRows: ReferenceRow[] = [
      refRow({ paxName: "Alice", passport: "P2000001", service: "60 DAYS VISA" }),
      refRow({ paxName: "Bob", passport: "P2000002", service: "60 DAYS VISA" }),
    ];
    enrichGroupRowsFromReference(rows, refRows);
    expect(rows[0].passport).toBe("P1000000");
    expect(rows[0].paxName).toBe("02 PAX 60 DAYS VISA");
    // The second (identity-less) sub-row still gets filled from the pool.
    expect(rows[1].passport).toBe("P2000001");
    expect(rows[1].paxName).toBe("Alice");
  });

  it("partially fills a group when the reference file has fewer candidates than the group size", () => {
    const rows = explodeMultiPax([
      row({
        paxName: "03 PAX 60 DAYS VISA",
        description: "Visa fee",
        charge: 3000,
        date: "2026-01-10",
      }),
    ]);
    const refRows: ReferenceRow[] = [
      refRow({ paxName: "Alice", passport: "P2000001", service: "60 DAYS VISA" }),
      refRow({ paxName: "Bob", passport: "P2000002", service: "60 DAYS VISA" }),
    ];
    enrichGroupRowsFromReference(rows, refRows);
    expect(rows[0].passport).toBe("P2000001");
    expect(rows[0].paxName).toBe("Alice");
    expect(rows[1].passport).toBe("P2000002");
    expect(rows[1].paxName).toBe("Bob");
    // Third sub-row has no candidate left — stays anonymous rather than being guessed.
    expect(rows[2].passport).toBeNull();
    expect(rows[2].paxName).toBe("03 PAX 60 DAYS VISA");
  });

  it("does nothing when there are no same-day reference candidates", () => {
    const rows = explodeMultiPax([
      row({ paxName: "02 PAX 60 DAYS VISA", description: "Visa fee", charge: 2000, date: "2026-01-10" }),
    ]);
    const refRows: ReferenceRow[] = [refRow({ date: "2026-05-05", paxName: "Zed", passport: "P9999999" })];
    expect(() => enrichGroupRowsFromReference(rows, refRows)).not.toThrow();
    expect(rows.every((r) => r.paxName === "02 PAX 60 DAYS VISA")).toBe(true);
  });

  it("matches a reference date stringified as a full JS Date (with time-of-day) to a plain ISO ledger date on the same calendar day", () => {
    // A reference file read via XLSX cellDates:true can hand back a JS Date that
    // gets stringified to "Tue Sep 09 2025 23:59:48 GMT+0400 (...)" — this must
    // still bucket under the SAME day as the ledger's plain "2025-09-09", even
    // though the two are hours apart and, in a timezone ahead of UTC, can even
    // fall on different UTC calendar days.
    const rows = explodeMultiPax([
      row({ paxName: "02 PAX 60 DAYS VISA", description: "Visa fee", charge: 2000, date: "2025-09-09" }),
    ]);
    const refRows: ReferenceRow[] = [
      refRow({ date: "Tue Sep 09 2025 23:59:48 GMT+0400 (Gulf Standard Time)", paxName: "Alice", passport: "P1" }),
      refRow({ date: "Tue Sep 09 2025 23:59:48 GMT+0400 (Gulf Standard Time)", paxName: "Bob", passport: "P2" }),
    ];
    enrichGroupRowsFromReference(rows, refRows);
    expect(rows[0].paxName).toBe("Alice");
    expect(rows[1].paxName).toBe("Bob");
  });
});

describe("reconcile — group booking remnants", () => {
  it("collapses unmatched ours-side group sub-rows into one summarized pair", () => {
    const ours = explodeMultiPax([
      row({ paxName: "03 PAX 60 DAYS VISA", description: "Visa fee", charge: 3000, date: "2026-03-01", srcRow: 7 }),
    ]);
    const out = reconcile(ours, []);
    const grp = out.pairs.filter((p) => p.status === "missing_partner" && p.key.startsWith("oo-grp-"));
    expect(grp.length).toBe(1);
    expect(grp[0].note).toContain("3 of 3 passenger(s) unmatched");
    expect(grp[0].oursAmt).toBe(3000);
  });

  it("collapses unmatched partner-side group sub-rows into one summarized pair (mirrors ours-side)", () => {
    const partner = explodeMultiPax([
      row({
        side: "partner",
        paxName: "03 PAX 60 DAYS VISA",
        description: "Visa fee",
        charge: 3000,
        date: "2026-03-01",
        srcRow: 9,
      }),
    ]);
    const out = reconcile([], partner);
    const grp = out.pairs.filter((p) => p.status === "missing_ours" && p.key.startsWith("op-grp-"));
    expect(grp.length).toBe(1);
    expect(grp[0].note).toContain("3 of 3 passenger(s) unmatched");
    expect(grp[0].partnerAmt).toBe(3000);
    // No fragmented per-pax rows should leak through alongside the summary.
    expect(out.pairs.filter((p) => p.status === "missing_ours").length).toBe(1);
  });
});
