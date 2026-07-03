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
  type LedgerRow,
  type Pair,
  type ReconResult,
} from "./reconcile";

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

  it("marks a pair as amount_diff when off by even 0.10 after conversion", () => {
    const p = pair({
      status: "matched",
      ours: row({ charge: 375.1 }),
      partner: row({ side: "partner", charge: 100, currency: "USD" }),
      oursAmt: 375.1,
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
