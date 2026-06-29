import { createFileRoute } from "@tanstack/react-router";
import React, { useMemo, useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  parseOurLedger,
  parsePartnerLedger,
  assertReadableSpreadsheet,
  reconcile,
  exportPairsCSV,
  buildReconciliationWorkbook,
  parseDynamicLedger,
  computeTotals,
  computeAnalytics,
  collectDuplicateGroups,
  collectRefunds,
  scoreRowPair,
  parseMonthlyFile,
  parseSoftwareEntryReportFile,
  autoParseYearFile,
  mergeLedgers,
  ledgerRowsToAoa,
  computeMonthlyBreakdown,
  monthFromFilename,
  monthKeyFromDate,
  rateDeviation,
  pairMonth,
  monthLabel,
  type ReconResult,
  type Pair,
  type LedgerRow,
  type ColumnMapping,
  type MatchEvidence,
  type MonthlyBreakdown,
} from "@/lib/reconcile";
import { analyzeSchema, performAiMatching } from "@/lib/server-actions";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
} from "recharts";
import {
  Brain,
  Sparkles,
  AlertCircle,
  CheckCircle2,
  TrendingUp,
  Info,
  Download,
  ShieldCheck,
  Gauge,
  Cpu,
  Search,
  FileSpreadsheet,
  FileText,
  UploadCloud,
  Table2,
  ChevronDown,
  CreditCard,
  Calendar,
  ArrowLeftRight,
  Filter,
  Users,
  RefreshCw,
  Landmark,
} from "lucide-react";
import type { Scenario } from "@/lib/reconcile";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Navvi Saadi | AI Ledger Reconciliation" },
      {
        name: "description",
        content:
          "Navvi Saadi Travel & Tourism — AI-powered, high-accuracy financial reconciliation.",
      },
    ],
  }),
  component: Index,
});

/* ---------- brand ---------- */
const NAVY = "#0c2e5f";
const GOLD = "#c9a23a";

/**
 * Visible build stamp. Bump this every deploy so it's obvious at a glance whether
 * the live site is serving the latest bundle or a cached/old one. Shown in the
 * footer — if the footer doesn't show this tag, the browser/CDN is stale.
 */
const BUILD_TAG = "2026-06-29 · build r7";

/** The Navvi Saadi gold arch / kufic dome mark, recreated as crisp vector. */
function BrandMark({ className = "" }: { className?: string }) {
  // Graduated minaret bars rising to a central apex, under a double arch.
  const bars = [24, 31.2, 38.4, 45.6, 52.8, 60, 67.2, 74.4, 81.6, 88.8, 96].map((x) => {
    const h = 14 + 26 * (1 - Math.abs(x - 60) / 36);
    return { x, top: 60 - h };
  });
  return (
    <svg viewBox="0 0 120 70" className={className} aria-hidden>
      <path
        d="M8 62 C8 26 60 9 60 9 C60 9 112 26 112 62"
        fill="none"
        stroke={GOLD}
        strokeWidth="3.4"
        strokeLinecap="round"
      />
      <path
        d="M16 62 C16 33 60 18 60 18 C60 18 104 33 104 62"
        fill="none"
        stroke={GOLD}
        strokeWidth="1.1"
        opacity="0.5"
      />
      <g stroke={GOLD} strokeWidth="2.9" strokeLinecap="round">
        {bars.map((b, i) => (
          <line key={i} x1={b.x} y1={60} x2={b.x} y2={b.top} />
        ))}
      </g>
    </svg>
  );
}

/** Vector lockup of the full brand: arch mark + wordmark + Arabic subtitle. */
function BrandLogoVector({ light = true }: { light?: boolean }) {
  return (
    <div className="flex items-center gap-3">
      <BrandMark className="h-11 w-auto shrink-0" />
      <div className="leading-tight">
        <div
          className={`text-[17px] font-semibold tracking-[0.26em] ${light ? "text-white" : "text-slate-800"}`}
        >
          NAVVI SAADI
        </div>
        <div className="flex items-center gap-2">
          <span
            className="text-[8.5px] font-semibold tracking-[0.32em]"
            style={{ color: GOLD }}
          >
            TRAVEL &amp; TOURISM
          </span>
          <span className="text-[9px]" style={{ color: GOLD }} dir="rtl">
            للسفر والسياحة
          </span>
        </div>
      </div>
    </div>
  );
}

/**
 * Company logo. Shows the polished vector brand immediately, and silently
 * upgrades to your exact logo image if it is present at:
 *   public/navvi-saadi-logo.png
 * (Preloaded so a missing file never flashes a broken-image icon.)
 */
function BrandLogo({ light = true }: { light?: boolean }) {
  const [hasImg, setHasImg] = useState(false);
  useEffect(() => {
    const img = new Image();
    img.onload = () => {
      if (img.naturalWidth > 1 && img.naturalHeight > 1) setHasImg(true);
    };
    img.src = "/navvi-saadi-logo.png";
  }, []);
  if (hasImg)
    return (
      <img
        src="/navvi-saadi-logo.png"
        alt="Navvi Saadi Travel & Tourism"
        className="h-12 w-auto shrink-0 object-contain"
        style={{ maxWidth: 240 }}
      />
    );
  return <BrandLogoVector light={light} />;
}

/* ---------- formatting helpers ---------- */
const money = (n: number) =>
  (n ?? 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const signed = (n: number) => (n > 0 ? "+" : "") + money(n);
const pct = (n: number) => `${Math.round((n ?? 0) * 100)}%`;

const confColor = (c: number) => (c >= 0.85 ? "#10b981" : c >= 0.7 ? "#f59e0b" : "#ef4444");
const confLabel = (c: number) => (c >= 0.85 ? "High" : c >= 0.7 ? "Medium" : "Low");

/** True when a ledger row is an inter-party settlement / bank transfer (not a per-item charge). */
const isTransfer = (row: LedgerRow | null | undefined): boolean =>
  !!row?.settlement || (row?.paxName ?? "").trim().toUpperCase() === "BANK TRANSFER";

/** True when either side of a pair is a settlement. */
const pairIsSettlement = (p: Pair): boolean => isTransfer(p.ours) || isTransfer(p.partner);

type StatusFilter =
  | "all"
  | "review"
  | "payments"
  | "fullledger"
  | "security_deposit"
  | "refunds"
  | "multi_passenger"
  | "duplicates"
  | "price_off"
  | Pair["status"];

/** A matched pair is "off-rate" when the supplier amount deviates more than this
 *  fraction from what the auto-detected ledger rate predicts (possible mispricing). */
const RATE_OFF_THRESHOLD = 0.15;

/** Per-scenario UI styling — used in table rows and detail panel badges. */
const SCENARIO_STYLE: Record<
  Scenario,
  { bg: string; border: string; text: string; label: string; dot: string }
> = {
  visa_charge: {
    bg: "bg-blue-50",
    border: "border-blue-200",
    text: "text-blue-700",
    label: "Visa Charge",
    dot: "bg-blue-500",
  },
  security_deposit: {
    bg: "bg-orange-50",
    border: "border-orange-200",
    text: "text-orange-700",
    label: "Security Deposit",
    dot: "bg-orange-500",
  },
  wrong_invoice: {
    bg: "bg-yellow-50",
    border: "border-yellow-300",
    text: "text-yellow-800",
    label: "Wrong Invoice Refund",
    dot: "bg-yellow-500",
  },
  wrong_client: {
    bg: "bg-yellow-50",
    border: "border-yellow-300",
    text: "text-yellow-800",
    label: "Wrong Client Refund",
    dot: "bg-yellow-500",
  },
  duplicate: {
    bg: "bg-purple-50",
    border: "border-purple-200",
    text: "text-purple-700",
    label: "Duplicate Refund",
    dot: "bg-purple-500",
  },
  refund: {
    bg: "bg-pink-50",
    border: "border-pink-200",
    text: "text-pink-700",
    label: "Refund / Reversal",
    dot: "bg-pink-500",
  },
  bank_transfer: {
    bg: "bg-emerald-50",
    border: "border-emerald-200",
    text: "text-emerald-700",
    label: "Bank Transfer",
    dot: "bg-emerald-500",
  },
  multi_passenger: {
    bg: "bg-indigo-50",
    border: "border-indigo-200",
    text: "text-indigo-700",
    label: "Multi-Passenger",
    dot: "bg-indigo-500",
  },
  flight: {
    bg: "bg-sky-50",
    border: "border-sky-200",
    text: "text-sky-700",
    label: "Flight / Airline",
    dot: "bg-sky-500",
  },
  addon: {
    bg: "bg-slate-50",
    border: "border-slate-200",
    text: "text-slate-500",
    label: "Add-On Service",
    dot: "bg-slate-400",
  },
};

function ScenarioBadge({ scenario }: { scenario: Scenario | undefined }) {
  if (!scenario) return null;
  const s = SCENARIO_STYLE[scenario];
  if (!s) return null;
  return (
    <span
      className={`inline-flex items-center gap-1 text-[9px] font-black px-1.5 py-0.5 rounded border ${s.bg} ${s.border} ${s.text}`}
    >
      <span className={`size-1.5 rounded-full ${s.dot}`} />
      {s.label}
    </span>
  );
}

type Aoa = unknown[][];

/* ================================================================== */
/*  MAIN                                                               */
/* ================================================================== */

function Index() {
  const [oursFile, setOursFile] = useState<File | null>(null);
  const [partnerFile, setPartnerFile] = useState<File | null>(null);
  /** Multiple monthly files — used in Year Mode (Our side) */
  const [oursFiles, setOursFiles] = useState<File[]>([]);
  /** Multiple monthly files — used in Year Mode (Partner/Supplier side) */
  const [partnerFiles, setPartnerFiles] = useState<File[]>([]);
  /** true = multi-file year reconciliation mode */
  const [yearMode, setYearMode] = useState(false);
  /** Per-side upload type in Year Mode */
  const [oursUploadType, setOursUploadType] = useState<"single" | "multi">("multi");
  // Partner side defaults to a single file — the supplier's annual statement
  // (e.g. "Copy of software entry report.xls") is one 1-year ledger.
  const [partnerUploadType, setPartnerUploadType] = useState<"single" | "multi">("single");
  const [rawOurs, setRawOurs] = useState<Aoa | null>(null);
  const [rawPartner, setRawPartner] = useState<Aoa | null>(null);
  const [result, setResult] = useState<ReconResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [isClient, setIsClient] = useState(false);
  const [aiStatus, setAiStatus] = useState<string>("");
  const [engineMode, setEngineMode] = useState<"ai" | "heuristic">("ai");
  const [monthBreakdown, setMonthBreakdown] = useState<MonthlyBreakdown[]>([]);
  const [monthFilter, setMonthFilter] = useState<string>("all");

  useEffect(() => setIsClient(true), []);

  const [filter, setFilter] = useState<StatusFilter>("all");
  const [query, setQuery] = useState("");
  const [sortByConf, setSortByConf] = useState(false);
  const [selected, setSelected] = useState<Pair | null>(null);
  const [schema, setSchema] = useState<any>(null);
  const [showSource, setShowSource] = useState(false);

  const getAoa = async (file: File): Promise<Aoa> => {
    const buf = await file.arrayBuffer();
    assertReadableSpreadsheet(buf, file.name);
    try {
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      // defval:"" pads every row to the full column width so no trailing column
      // is ever lost. Blank rows are kept so source-row indices stay aligned with
      // the parser (the export/full-ledger view skips empty rows on output).
      return XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: "" });
    } catch {
      const text = new TextDecoder("utf-8").decode(buf);
      return Papa.parse<unknown[]>(text, { skipEmptyLines: true }).data;
    }
  };

  const selectFile = async (side: "ours" | "partner", file: File | null) => {
    if (side === "ours") setOursFile(file);
    else setPartnerFile(file);
    if (!file) {
      side === "ours" ? setRawOurs(null) : setRawPartner(null);
      return;
    }
    try {
      const aoa = await getAoa(file);
      side === "ours" ? setRawOurs(aoa) : setRawPartner(aoa);
    } catch {
      /* preview is best-effort */
    }
  };

  /* ---- rule-validated AI residual matching ---- */
  const aiResidualMatch = async (
    oursRows: LedgerRow[],
    partnerRows: LedgerRow[],
    scoreFn: (o: LedgerRow, p: LedgerRow) => { score: number; evidence: MatchEvidence },
  ) => {
    if (!oursRows.length || !partnerRows.length) return [];
    const resp: any = await performAiMatching({
      data: { unmatchedOurs: oursRows, unmatchedPartner: partnerRows },
    });
    const aiPairs: any[] = resp?.data?.pairs ?? [];
    const usedO = new Set<number>();
    const usedP = new Set<number>();
    const out: {
      o: LedgerRow;
      p: LedgerRow;
      confidence: number;
      evidence: MatchEvidence;
      reason: string;
    }[] = [];
    aiPairs
      .slice()
      .sort((a, b) => (b.confidence ?? 0) - (a.confidence ?? 0))
      .forEach((m) => {
        const o = oursRows[m.oursIndex];
        const p = partnerRows[m.partnerIndex];
        if (!o || !p || usedO.has(m.oursIndex) || usedP.has(m.partnerIndex)) return;
        const { score, evidence } = scoreFn(o, p);
        const aiConf = typeof m.confidence === "number" ? m.confidence : 0.6;
        if (score < 0.4 && aiConf < 0.85) return; // rule re-validation
        usedO.add(m.oursIndex);
        usedP.add(m.partnerIndex);
        out.push({
          o,
          p,
          confidence: Math.min(0.97, 0.55 * score + 0.45 * aiConf),
          evidence: { ...evidence, method: "ai" },
          reason: m.reason || "semantic pairing",
        });
      });
    return out;
  };

  const runSmartRecon = async () => {
    setError(null);
    setBusy(true);
    setResult(null);
    setSelected(null);
    setSchema(null);
    setFilter("all");
    setShowSource(false);
    setMonthBreakdown([]);
    setMonthFilter("all");
    // Clear any AOA from a previous run so switching files/modes never reuses a
    // stale sheet (single mode reads fresh; year mode rebuilds from parsed rows).
    setRawOurs(null);
    setRawPartner(null);
    setAiStatus("Reading files…");
    try {
      let ours: LedgerRow[];
      let partner: LedgerRow[];
      let mode: "ai" | "heuristic" = "heuristic";

      /* ── YEAR MODE: each side independently single or multi-file ──────────
         Auto-detects format (GDS monthly vs Software Entry Report) per file.
         ─────────────────────────────────────────────────────────────────── */
      const yearOursReady = oursUploadType === "multi" ? oursFiles.length > 0 : !!oursFile;
      const yearPartnerReady = partnerUploadType === "multi" ? partnerFiles.length > 0 : !!partnerFile;

      if (yearMode && yearOursReady && yearPartnerReady) {
        // ── Parse Our Ledger ───────────────────────────────────────────────
        if (oursUploadType === "multi") {
          const sorted = [...oursFiles].sort((a, b) =>
            monthFromFilename(a.name).localeCompare(monthFromFilename(b.name))
          );
          const allParsed: LedgerRow[][] = [];
          for (let i = 0; i < sorted.length; i++) {
            setAiStatus(`Parsing Our Ledger: ${sorted[i].name} (${i + 1}/${sorted.length})…`);
            const rows = await autoParseYearFile(sorted[i], "ours");
            if (rows.length === 0) {
              throw new Error(
                `Could not read any rows from "${sorted[i].name}". ` +
                `The sheet should have a header row with at least a date and an amount column.`
              );
            }
            allParsed.push(rows);
          }
          ours = mergeLedgers(allParsed);
        } else {
          setAiStatus(`Parsing Our Ledger: ${oursFile!.name}…`);
          ours = await autoParseYearFile(oursFile!, "ours");
          if (!ours.length) throw new Error(`Could not read any rows from "${oursFile!.name}".`);
        }

        setAiStatus(`Our Ledger: ${ours.length} rows loaded.`);

        // ── Parse Partner Ledger ───────────────────────────────────────────
        if (partnerUploadType === "multi") {
          const sorted = [...partnerFiles].sort((a, b) =>
            monthFromFilename(a.name).localeCompare(monthFromFilename(b.name))
          );
          const allParsed: LedgerRow[][] = [];
          for (let i = 0; i < sorted.length; i++) {
            setAiStatus(`Parsing Partner Ledger: ${sorted[i].name} (${i + 1}/${sorted.length})…`);
            const rows = await autoParseYearFile(sorted[i], "partner");
            if (rows.length === 0) {
              throw new Error(
                `Could not read any rows from partner file "${sorted[i].name}". ` +
                `The sheet should have a header row with at least a date and an amount column.`
              );
            }
            allParsed.push(rows);
          }
          partner = mergeLedgers(allParsed);
        } else {
          setAiStatus(`Parsing Partner Ledger: ${partnerFile!.name}…`);
          partner = await autoParseYearFile(partnerFile!, "partner");
          if (!partner.length) throw new Error(`Could not read any rows from "${partnerFile!.name}".`);
        }

        // One side is usually the FULL-YEAR statement while the other is just the
        // month(s) you actually uploaded (e.g. only January). Reconcile ONLY those
        // months: restrict the broader side to the months present on the narrower
        // side — in EITHER direction.
        const getIntendedMonths = (uploadType: "single" | "multi", files: File[], rows: LedgerRow[]) => {
          if (uploadType === "multi") {
            const explicit = new Set(files.map((f) => monthFromFilename(f.name)).filter(Boolean));
            if (explicit.size > 0) return explicit;
          }
          return new Set(rows.map((r) => r.month || monthKeyFromDate(r.date)).filter(Boolean) as string[]);
        };

        const restrictTo = (rows: LedgerRow[], focus: Set<string>) =>
          rows.filter((r) => { const m = r.month || monthKeyFromDate(r.date); return !m || focus.has(m); });

        const oursMonths = getIntendedMonths(oursUploadType, oursUploadType === "multi" ? oursFiles : oursFile ? [oursFile] : [], ours);
        const partnerMonths = getIntendedMonths(partnerUploadType, partnerUploadType === "multi" ? partnerFiles : partnerFile ? [partnerFile] : [], partner);

        if (oursMonths.size > 0 && partnerMonths.size > 0 && oursMonths.size !== partnerMonths.size) {
          // Focus on the narrower (monthly) side; trim the broader (annual) side.
          const focus = oursMonths.size < partnerMonths.size ? oursMonths : partnerMonths;
          const trimOurs = oursMonths.size > partnerMonths.size;
          const label = [...focus].sort().map(monthLabel).join(", ");
          if (trimOurs) {
            const before = ours.length;
            ours = restrictTo(ours, focus);
            setAiStatus(`Reconciling only your uploaded month(s): ${label} (our ledger ${before} → ${ours.length} rows)…`);
          } else {
            const before = partner.length;
            partner = restrictTo(partner, focus);
            setAiStatus(`Reconciling only your uploaded month(s): ${label} (supplier ${before} → ${partner.length} rows)…`);
          }
        }

        setAiStatus(`Partner Ledger: ${partner.length} rows loaded. Running reconciliation…`);
        setEngineMode("heuristic");
        setSchema(null);
        mode = "heuristic";
      } else {
        /* ── SINGLE FILE MODE (original logic) ───────────────────────────── */
        if (!oursFile || !partnerFile) throw new Error("Please upload both ledger files.");

        const aoaOurs = rawOurs ?? (await getAoa(oursFile));
        const aoaPartner = rawPartner ?? (await getAoa(partnerFile));
        setRawOurs(aoaOurs);
        setRawPartner(aoaPartner);

        try {
          setAiStatus("AI analysing column structure…");
          const schemaResponse: any = await analyzeSchema({
            data: { ours: aoaOurs.slice(0, 50), partner: aoaPartner.slice(0, 50) },
          });
          if (!schemaResponse?.data) throw new Error("AI schema discovery returned no data.");
          const sc = schemaResponse.data;
          setSchema(sc);

          setAiStatus("Mapping & parsing rows…");
          ours = parseDynamicLedger(aoaOurs, "ours", sc.ours as ColumnMapping);
          partner = parseDynamicLedger(aoaPartner, "partner", sc.partner as ColumnMapping);
          if (!ours.length || !partner.length) throw new Error("AI mapping produced no usable rows.");
          mode = "ai";
        } catch (aiErr) {
          console.warn("[Recon] AI schema discovery failed, using heuristic parser:", aiErr);
          setAiStatus("AI unavailable — using built-in heuristic parser…");
          mode = "heuristic";
          setSchema(null);
          ours = await parseOurLedger(oursFile);
          partner = await parsePartnerLedger(partnerFile);
        }

        // Safety net: if the AI column-mapping reconciled almost nothing, it was
        // probably wrong — fall back to the deterministic heuristic parser and keep
        // whichever reconciles more rows.
        if (mode === "ai") {
          const aiTry = reconcile(ours, partner);
          const minRows = Math.min(ours.length, partner.length) || 1;
          if (aiTry.totals.matched / minRows < 0.15) {
            try {
              setAiStatus("AI mapping weak — verifying with heuristic engine…");
              const ho = await parseOurLedger(oursFile);
              const hp = await parsePartnerLedger(partnerFile);
              if (reconcile(ho, hp).totals.matched > aiTry.totals.matched) {
                ours = ho;
                partner = hp;
                mode = "heuristic";
                setSchema(null);
              }
            } catch (e) {
              console.warn("[Recon] heuristic verification failed:", e);
            }
          }
        }
        setEngineMode(mode);
      }

      setAiStatus("Running multi-signal rule engine…");
      const baseResult = reconcile(ours, partner);

      // ── YEAR MODE result ──────────────────────────────────────────────────
      // Multi-passenger group bookings (the supplier invoices one combined line
      // per PNR while the GDS lists one row per passenger) are already resolved
      // inside reconcile(): same-PNR flight rows are consolidated to one booking
      // per side BEFORE matching, so each booking reconciles as a single pair
      // with no double-counting. Just publish the result and the monthly summary.
      if (yearMode) {
        // Year mode merges several monthly files, so there is no single original
        // sheet to show. Build a normalized all-entries AOA per side from the
        // reconciled rows so the "Uploaded Source Files" panel, the Both-Sheets
        // full ledger and the Excel export all show real data (instead of "—").
        // ledgerRowsToAoa also reassigns each row's srcRow so the status maps and
        // month filtering in the full ledger line up.
        const oursRows = baseResult.pairs.map((p) => p.ours).filter(Boolean) as LedgerRow[];
        const partnerRows = baseResult.pairs.map((p) => p.partner).filter(Boolean) as LedgerRow[];
        setRawOurs(ledgerRowsToAoa(oursRows));
        setRawPartner(ledgerRowsToAoa(partnerRows));
        setResult(baseResult);
        setMonthBreakdown(computeMonthlyBreakdown(baseResult.pairs));
        setAiStatus("");
        setBusy(false);
        return;
      }

      setResult(baseResult);

      const onlyOursRows = baseResult.pairs
        .filter((p) => p.status === "missing_partner" && p.ours)
        .map((p) => p.ours!);
      const onlyPartnerRows = baseResult.pairs
        .filter((p) => p.status === "missing_ours" && p.partner)
        .map((p) => p.partner!);

      if (mode === "ai" && onlyOursRows.length && onlyPartnerRows.length) {
        setAiStatus("Deep AI semantic matching on residuals…");
        const matches = await aiResidualMatch(onlyOursRows, onlyPartnerRows, scoreRowPair);

        if (matches.length) {
          setAiStatus("Validating & merging AI matches…");
          let merged = [...baseResult.pairs];
          for (const mm of matches) {
            merged = merged.filter(
              (pr) =>
                !(pr.ours?.index === mm.o.index && pr.status === "missing_partner") &&
                !(pr.partner?.index === mm.p.index && pr.status === "missing_ours"),
            );
            const oAmt = mm.o.charge > 0 ? mm.o.charge : mm.o.credit;
            const pAmt = mm.p.charge > 0 ? mm.p.charge : mm.p.credit;
            const diff = +(pAmt - oAmt).toFixed(2);
            const exact = Math.abs(diff) < 0.5;
            merged.push({
              key: `ai-${mm.o.index}-${mm.p.index}`,
              status: exact ? "matched" : "amount_diff",
              kind: mm.o.kind === "credit" ? "credit" : "charge",
              ours: mm.o,
              partner: mm.p,
              oursAmt: oAmt,
              partnerAmt: pAmt,
              diff: exact ? 0 : diff,
              score: mm.confidence,
              confidence: mm.confidence,
              needsReview: mm.confidence < 0.85,
              evidence: mm.evidence,
              note: `AI match: ${mm.reason}`,
              aiInsight: mm.reason,
            });
          }
          const finalResult = { pairs: merged, totals: computeTotals(ours, partner, merged) };
          setResult(finalResult);
          if (yearMode) setMonthBreakdown(computeMonthlyBreakdown(finalResult.pairs));
        }
      }

      setAiStatus("");
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
      setAiStatus("");
    } finally {
      setBusy(false);
    }
  };

  const downloadBlob = (data: BlobPart, type: string, ext: string) => {
    const blob = new Blob([data], { type });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `navvi-saadi-reconciliation-${new Date().toISOString().slice(0, 10)}.${ext}`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const exportCsv = () => {
    if (!result) return;
    downloadBlob(exportPairsCSV(result.pairs), "text/csv;charset=utf-8;", "csv");
  };

  const exportExcel = () => {
    if (!result) return;
    const buf = buildReconciliationWorkbook(result, {
      oursAoa: rawOurs,
      partnerAoa: rawPartner,
      monthlyBreakdown: monthBreakdown.length > 0 ? monthBreakdown : undefined,
    });
    downloadBlob(
      buf,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "xlsx",
    );
  };

  const exportPdf = async () => {
    if (!result) return;
    setAiStatus("Building PDF report…");
    try {
      const [{ default: JsPDF }, autoTableMod] = await Promise.all([
        import("jspdf"),
        import("jspdf-autotable"),
      ]);
      const autoTable = (autoTableMod as any).default ?? (autoTableMod as any);
      const doc = new JsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
      const W = doc.internal.pageSize.getWidth();
      const t = result.totals;
      const paired = result.pairs.filter((p) => p.ours && p.partner).length;
      const matchRatePct = result.pairs.length
        ? Math.round((paired / result.pairs.length) * 100)
        : 0;

      /* ---- Header band ---- */
      doc.setFillColor(12, 46, 95);
      doc.rect(0, 0, W, 56, "F");
      doc.setTextColor(201, 162, 58);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(16);
      doc.text("NAVVI SAADI", 40, 26);
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(11);
      doc.text("AI Ledger Reconciliation Report", 40, 44);
      doc.setFontSize(8);
      doc.setFont("helvetica", "normal");
      doc.text(`Generated ${new Date().toLocaleString()}`, W - 40, 30, { align: "right" });

      /* ---- KPI strip ---- */
      const kpis: [string, string][] = [
        ["Total Items", String(result.pairs.length)],
        ["Match Rate", `${matchRatePct}%`],
        ["Matched", String(t.matched)],
        ["Amount Diff", String(t.amountIssues)],
        ["Only Ours", String(t.onlyOurs)],
        ["Only Partner", String(t.onlyPartner)],
        ["Needs Review", String(t.needsReview)],
        ["Net Diff", money(t.netAmountDiff)],
      ];
      const kpiW = (W - 80) / kpis.length;
      kpis.forEach(([label, val], i) => {
        const x = 40 + i * kpiW;
        doc.setFillColor(247, 249, 252);
        doc.roundedRect(x, 70, kpiW - 8, 40, 4, 4, "F");
        doc.setTextColor(100, 116, 139);
        doc.setFontSize(7);
        doc.setFont("helvetica", "bold");
        doc.text(label.toUpperCase(), x + 8, 84);
        doc.setTextColor(12, 46, 95);
        doc.setFontSize(13);
        doc.text(val, x + 8, 102);
      });

      /* ---- Status colour map ---- */
      const fillFor: Record<string, [number, number, number]> = {
        Matched: [216, 243, 227],
        "Amount Difference": [254, 243, 199],
        "Only in Our Ledger": [224, 231, 255],
        "Only in Partner Ledger": [255, 228, 230],
      };
      const textFor: Record<string, [number, number, number]> = {
        Matched: [4, 120, 87],
        "Amount Difference": [180, 83, 9],
        "Only in Our Ledger": [67, 56, 202],
        "Only in Partner Ledger": [190, 18, 60],
      };
      const STATUS_TEXT: Record<string, string> = {
        matched: "Matched",
        amount_diff: "Amount Difference",
        missing_partner: "Only in Our Ledger",
        missing_ours: "Only in Partner Ledger",
      };

      const body = result.pairs.map((p) => {
        const oAmt = p.ours ? p.ours.charge || p.ours.credit : null;
        const pAmt = p.partner ? p.partner.charge || p.partner.credit : null;
        return [
          STATUS_TEXT[p.status],
          typeof p.confidence === "number" ? `${Math.round(p.confidence * 100)}%` : "",
          p.ours?.settlement || p.partner?.settlement ? "Settlement" : "Charge",
          p.ours?.passport ?? p.partner?.passport ?? "",
          p.ours?.date ?? "⚠ MISSING",
          oAmt === null ? "⚠ MISSING" : money(oAmt),
          p.partner?.date ?? "⚠ MISSING",
          pAmt === null ? "⚠ MISSING" : money(pAmt),
          p.status === "matched" ? "✓" : signed(p.diff),
          p.evidence?.dateDeltaDays ?? "",
        ];
      });

      autoTable(doc, {
        startY: 122,
        head: [
          [
            "Status",
            "Conf",
            "Category",
            "ID / Passport",
            "Our Date",
            "Our Amount",
            "Partner Date",
            "Partner Amount",
            "Variance",
            "Gap",
          ],
        ],
        body,
        theme: "grid",
        styles: { fontSize: 7, cellPadding: 3, lineColor: [214, 222, 232], lineWidth: 0.5 },
        headStyles: { fillColor: [12, 46, 95], textColor: [255, 255, 255], fontStyle: "bold" },
        columnStyles: {
          5: { halign: "right" },
          7: { halign: "right" },
          8: { halign: "right" },
          9: { halign: "center" },
        },
        didParseCell: (data: any) => {
          if (data.section !== "body") return;
          const status = body[data.row.index][0] as string;
          // Whole-row tint by status
          const fill = fillFor[status];
          if (fill) data.cell.styles.fillColor = fill;
          // Status cell: bold coloured text
          if (data.column.index === 0 && textFor[status]) {
            data.cell.styles.textColor = textFor[status];
            data.cell.styles.fontStyle = "bold";
          }
          // Missing cells → strong red
          if (String(data.cell.raw).includes("MISSING")) {
            data.cell.styles.fillColor = [252, 165, 165];
            data.cell.styles.textColor = [127, 29, 29];
            data.cell.styles.fontStyle = "bold";
          }
          // Variance emphasis
          if (data.column.index === 8 && String(data.cell.raw) !== "✓") {
            data.cell.styles.textColor = [190, 18, 60];
            data.cell.styles.fontStyle = "bold";
          }
        },
        didDrawPage: (data: any) => {
          doc.setFontSize(7);
          doc.setTextColor(150, 150, 150);
          doc.text(
            `Navvi Saadi Travel & Tourism · Page ${doc.getNumberOfPages()}`,
            W - 40,
            doc.internal.pageSize.getHeight() - 12,
            { align: "right" },
          );
        },
      });

      /* ---- Section title helper ---- */
      const sectionTitle = (title: string, subtitle: string) => {
        doc.addPage();
        doc.setFillColor(12, 46, 95);
        doc.rect(0, 0, W, 38, "F");
        doc.setTextColor(201, 162, 58);
        doc.setFont("helvetica", "bold");
        doc.setFontSize(13);
        doc.text(title, 40, 22);
        doc.setTextColor(220, 226, 234);
        doc.setFont("helvetica", "normal");
        doc.setFontSize(8);
        doc.text(subtitle, 40, 33);
      };

      /* ---- INSIGHTS: per-category performance ---- */
      const analytics = computeAnalytics(result.pairs);
      sectionTitle(
        "Insights — Performance by Category",
        "How every type of entry reconciled, and the value verified on each.",
      );
      autoTable(doc, {
        startY: 52,
        head: [["Category", "Total", "Matched", "Amount Diff", "Only Ours", "Only Partner", "Match %", "Matched Value"]],
        body: analytics.scenarios.map((s) => [
          s.label,
          s.total,
          s.matched,
          s.amountDiff,
          s.onlyOurs,
          s.onlyPartner,
          `${s.total ? Math.round((s.matched / s.total) * 100) : 0}%`,
          money(s.matchedValue),
        ]),
        theme: "grid",
        styles: { fontSize: 8, cellPadding: 4, lineColor: [214, 222, 232], lineWidth: 0.5 },
        headStyles: { fillColor: [12, 46, 95], textColor: [255, 255, 255], fontStyle: "bold" },
        columnStyles: { 1: { halign: "right" }, 2: { halign: "right" }, 3: { halign: "right" }, 4: { halign: "right" }, 5: { halign: "right" }, 6: { halign: "right" }, 7: { halign: "right" } },
      });

      /* ---- DUPLICATE ENTRIES: grouped, every copy shown ---- */
      const dupGroups = collectDuplicateGroups(result.pairs);
      sectionTitle(
        "Duplicate Entries — Where They Happen",
        "Same charge written more than once on one ledger. Every copy is listed; one of each group is likely a mistake to remove.",
      );
      const dupBody: any[] = [];
      const dupGroupStart = new Set<number>();
      if (!dupGroups.length) {
        dupBody.push(["✓ No duplicate entries found on either ledger.", "", "", "", "", "", "", ""]);
      } else {
        dupGroups.forEach((g, gi) => {
          dupGroupStart.add(dupBody.length);
          dupBody.push([
            `▼ ${g.side === "ours" ? "OUR LEDGER" : "PARTNER LEDGER"} · ${g.rows.length} copies · ${money(g.amount * (g.rows.length - 1))} redundant`,
            "", "", "", "", "", "", "",
          ]);
          g.rows.forEach((c, ci) => {
            dupBody.push([
              g.side === "ours" ? "Our" : "Partner",
              `${ci + 1}/${g.rows.length}`,
              c.row.paxName || "—",
              c.row.passport ?? "—",
              c.row.visaType ?? "",
              c.row.date || "—",
              c.row.reference || "—",
              money(c.row.charge > 0 ? c.row.charge : c.row.credit),
            ]);
          });
        });
      }
      autoTable(doc, {
        startY: 52,
        head: [["Ledger", "Copy", "Passenger", "ID / Passport", "Visa Type", "Date", "Voucher / Ref", "Amount"]],
        body: dupBody,
        theme: "grid",
        styles: { fontSize: 7.5, cellPadding: 3, lineColor: [214, 222, 232], lineWidth: 0.5 },
        headStyles: { fillColor: [12, 46, 95], textColor: [255, 255, 255], fontStyle: "bold" },
        columnStyles: { 7: { halign: "right" } },
        didParseCell: (data: any) => {
          if (data.section !== "body") return;
          if (dupGroupStart.has(data.row.index)) {
            data.cell.styles.fillColor = [254, 243, 199];
            data.cell.styles.textColor = [146, 64, 14];
            data.cell.styles.fontStyle = "bold";
          } else if (data.row.index > 0) {
            // 2nd+ copy rows tinted to stand out as the redundant ones
            const colCopy = dupBody[data.row.index][1];
            if (typeof colCopy === "string" && colCopy.startsWith("1/") === false) {
              data.cell.styles.fillColor = [255, 235, 238];
            }
          }
        },
      });

      /* ---- REFUNDS & REVERSALS (VR) ---- */
      const refunds = collectRefunds(result.pairs);
      sectionTitle(
        "Refunds & Reversals (VR) — Money Back",
        "Every refund and reversal (wrong invoice / wrong client / duplicate cancelled), and whether it matched the other ledger.",
      );
      autoTable(doc, {
        startY: 52,
        head: [["Ledger", "Reason", "Passenger", "ID / Passport", "Date", "Voucher / Ref", "Amount Back", "Matched?", "Matched-To"]],
        body: refunds.length
          ? refunds.map((f) => [
              f.side === "ours" ? "Our" : "Partner",
              f.reason,
              f.paxName,
              f.passport,
              f.date,
              f.reference,
              money(f.amount),
              f.matched ? "✓ Matched" : "✗ Only here",
              f.counterparty,
            ])
          : [["✓ No refunds or reversals found.", "", "", "", "", "", "", "", ""]],
        theme: "grid",
        styles: { fontSize: 7.5, cellPadding: 3, lineColor: [214, 222, 232], lineWidth: 0.5 },
        headStyles: { fillColor: [12, 46, 95], textColor: [255, 255, 255], fontStyle: "bold" },
        columnStyles: { 6: { halign: "right" }, 7: { halign: "center" } },
        didParseCell: (data: any) => {
          if (data.section !== "body" || !refunds.length) return;
          const matched = refunds[data.row.index]?.matched;
          data.cell.styles.fillColor = matched ? [216, 243, 227] : [224, 231, 255];
        },
      });

      /* ---- MONTHLY BREAKDOWN (Year Mode only) ---- */
      if (monthBreakdown.length > 0) {
        sectionTitle(
          "Monthly Breakdown — 1-Year Reconciliation",
          "Per-month reconciliation rate, match counts, and amounts verified across all uploaded months.",
        );
        const mbTotals = {
          total: monthBreakdown.reduce((s, b) => s + b.total, 0),
          matched: monthBreakdown.reduce((s, b) => s + b.matched, 0),
          onlyOurs: monthBreakdown.reduce((s, b) => s + b.onlyOurs, 0),
          onlyPartner: monthBreakdown.reduce((s, b) => s + b.onlyPartner, 0),
          oursTotal: monthBreakdown.reduce((s, b) => s + b.oursTotal, 0),
          partnerTotal: monthBreakdown.reduce((s, b) => s + b.partnerTotal, 0),
        };
        autoTable(doc, {
          startY: 52,
          head: [["Month", "Total", "Matched", "Only Ours", "Only Partner", "Match %", "Our Amount", "Supplier Amount"]],
          body: [
            ...monthBreakdown.map((bk) => [
              bk.label,
              bk.total,
              bk.matched,
              bk.onlyOurs,
              bk.onlyPartner,
              `${Math.round(bk.matchRate * 100)}%`,
              money(bk.oursTotal),
              money(bk.partnerTotal),
            ]),
            [
              "TOTAL",
              mbTotals.total,
              mbTotals.matched,
              mbTotals.onlyOurs,
              mbTotals.onlyPartner,
              `${mbTotals.total ? Math.round(mbTotals.matched / mbTotals.total * 100) : 0}%`,
              money(mbTotals.oursTotal),
              money(mbTotals.partnerTotal),
            ],
          ],
          theme: "grid",
          styles: { fontSize: 8, cellPadding: 4, lineColor: [214, 222, 232], lineWidth: 0.5 },
          headStyles: { fillColor: [12, 46, 95], textColor: [255, 255, 255], fontStyle: "bold" },
          columnStyles: { 1: { halign: "right" }, 2: { halign: "right" }, 3: { halign: "right" }, 4: { halign: "right" }, 5: { halign: "right" }, 6: { halign: "right" }, 7: { halign: "right" } },
          didParseCell: (data: any) => {
            if (data.section !== "body") return;
            const isTotal = data.row.index === monthBreakdown.length;
            if (isTotal) {
              data.cell.styles.fillColor = [229, 231, 235];
              data.cell.styles.fontStyle = "bold";
            } else {
              const pct = monthBreakdown[data.row.index]?.matchRate ?? 0;
              if (data.column.index === 5) {
                data.cell.styles.textColor = pct >= 0.85 ? [4, 120, 87] : pct >= 0.6 ? [180, 83, 9] : [190, 18, 60];
                data.cell.styles.fontStyle = "bold";
              }
            }
          },
        });
      }

      downloadBlob(doc.output("blob"), "application/pdf", "pdf");
    } catch (e) {
      setError(e instanceof Error ? e.message : "PDF export failed");
    } finally {
      setAiStatus("");
    }
  };

  /* ---- derived view data ---- */
  const filteredPairs = useMemo(() => {
    if (!result) return [];
    const q = query.trim().toLowerCase();
    let list = result.pairs.filter((p) => {
      // Month filter (Year Mode) — checks BOTH sides so the month view shows our
      // entries and supplier entries (including supplier-only rows) for the month.
      if (monthFilter !== "all") {
        if (pairMonth(p) !== monthFilter) return false;
      }
      if (filter === "review") {
        if (!p.needsReview) return false;
      } else if (filter === "payments") {
        if (!isTransfer(p.ours) && !isTransfer(p.partner)) return false;
      } else if (filter === "security_deposit") {
        const sc = p.ours?.scenario ?? p.partner?.scenario;
        if (sc !== "security_deposit") return false;
      } else if (filter === "refunds") {
        const sc = p.ours?.scenario ?? p.partner?.scenario;
        if (!sc || !["wrong_invoice", "wrong_client", "duplicate", "refund"].includes(sc)) return false;
      } else if (filter === "multi_passenger") {
        const sc = p.ours?.scenario ?? p.partner?.scenario;
        if (sc !== "multi_passenger") return false;
      } else if (filter === "duplicates") {
        const dup = (p.ours?.duplicateCount ?? 0) > 1 || (p.partner?.duplicateCount ?? 0) > 1;
        if (!dup) return false;
      } else if (filter === "price_off") {
        const d = rateDeviation(p, totals?.impliedRate ?? 0);
        if (d === null || Math.abs(d) <= RATE_OFF_THRESHOLD) return false;
      } else if (filter !== "all" && p.status !== filter) return false;
      if (!q) return true;
      const hay = [
        p.ours?.passport,
        p.partner?.passport,
        p.ours?.paxName,
        p.partner?.paxName,
        p.ours?.description,
        p.partner?.description,
        p.ours?.reference,
        p.partner?.reference,
        p.ours?.visaType,
        p.partner?.visaType,
      ]
        .filter(Boolean)
        .join(" ")
        .toLowerCase();
      return hay.includes(q);
    });
    if (sortByConf) list = [...list].sort((a, b) => (a.confidence ?? 1.1) - (b.confidence ?? 1.1));
    return list;
  }, [result, filter, query, sortByConf, monthFilter]);

  // Pairs in scope of the active month filter only (no status filter). The filter
  // TAB COUNTS are computed from this so they reflect the month you're viewing —
  // otherwise the tabs show whole-year totals while the table shows one month.
  const monthPairs = useMemo(() => {
    if (!result) return [];
    if (monthFilter === "all") return result.pairs;
    return result.pairs.filter((p) => pairMonth(p) === monthFilter);
  }, [result, monthFilter]);

  const activeTotals = useMemo(() => {
    if (!result) return null;
    if (monthFilter === "all") return result.totals;
    const oursRows = monthPairs.map((p) => p.ours).filter(Boolean) as LedgerRow[];
    const partnerRows = monthPairs.map((p) => p.partner).filter(Boolean) as LedgerRow[];
    return computeTotals(oursRows, partnerRows, monthPairs);
  }, [result, monthFilter, monthPairs]);

  const totals = activeTotals;

  const chartData = useMemo(() => {
    if (!result || !totals) return [];
    const t = totals;
    return [
      { name: "Matched", value: t.matched, color: "#10b981" },
      { name: "Amount Diff", value: t.amountIssues, color: "#f59e0b" },
      { name: "Only Ours", value: t.onlyOurs, color: "#6366f1" },
      { name: "Only Partner", value: t.onlyPartner, color: "#ef4444" },
    ];
  }, [result, totals]);

  const confHist = useMemo(() => {
    const buckets = [
      { name: "<60", value: 0, color: "#ef4444" },
      { name: "60-80", value: 0, color: "#f59e0b" },
      { name: "80-95", value: 0, color: "#3b82f6" },
      { name: "95+", value: 0, color: "#10b981" },
    ];
    if (!result) return buckets;
    monthPairs.forEach((p) => {
      if (typeof p.confidence !== "number") return;
      const c = p.confidence;
      if (c < 0.6) buckets[0].value++;
      else if (c < 0.8) buckets[1].value++;
      else if (c < 0.95) buckets[2].value++;
      else buckets[3].value++;
    });
    return buckets;
  }, [result, monthPairs]);

  const matchRate = useMemo(() => {
    if (!result) return 0;
    const paired = monthPairs.filter((p) => p.ours && p.partner).length;
    return paired / (monthPairs.length || 1);
  }, [result, monthPairs]);

  const matchedValue = useMemo(() => {
    if (!result) return 0;
    return +monthPairs
      .filter((p) => p.status === "matched")
      .reduce((s, p) => s + p.partnerAmt, 0)
      .toFixed(2);
  }, [result, monthPairs]);

  // Detected currency per side (most common code across the rows). Drives the
  // currency note + amount-column labels so they reflect the actual files, not
  // a hardcoded SAR/AED assumption. Returns the full set too, for mixed uploads.
  const currencies = useMemo(() => {
    const tally = (pick: (p: Pair) => string | undefined) => {
      const m = new Map<string, number>();
      if (result) for (const p of result.pairs) {
        const c = pick(p);
        if (c) m.set(c, (m.get(c) ?? 0) + 1);
      }
      const sorted = [...m.entries()].sort((a, b) => b[1] - a[1]);
      return { top: sorted[0]?.[0] as string | undefined, all: sorted.map(([c]) => c) };
    };
    const ours = tally((p) => p.ours?.currency);
    const partner = tally((p) => p.partner?.currency);
    return { ours: ours.top, partner: partner.top, oursAll: ours.all, partnerAll: partner.all };
  }, [result]);
  const CURRENCY_NAME: Record<string, string> = {
    SAR: "Saudi Riyal", AED: "UAE Dirham", USD: "US Dollar", QAR: "Qatari Riyal",
    KWD: "Kuwaiti Dinar", BHD: "Bahraini Dinar", OMR: "Omani Rial",
  };

  const hasData = !!(rawOurs || rawPartner || oursFile || partnerFile ||
    (yearMode && (oursFiles.length > 0 || partnerFiles.length > 0)));

  return (
    <div className="min-h-screen bg-gradient-to-b from-slate-50 to-white text-slate-900 font-sans">
      {/* ---------------- HEADER ---------------- */}
      <header
        className="sticky top-0 z-50 border-b border-amber-400/20 shadow-lg"
        style={{ background: `linear-gradient(100deg, #0a2547 0%, ${NAVY} 60%, #103a73 100%)` }}
      >
        <div className="mx-auto max-w-[1600px] px-6 py-3 flex items-center gap-5 flex-wrap">
          <BrandLogo />
          <div className="hidden md:block h-9 w-px bg-white/15" />
          <div className="hidden md:flex flex-col">
            <span className="text-[11px] font-bold text-white/90 tracking-wide">
              AI Ledger Reconciliation
            </span>
            <span
              className="text-[9px] font-semibold flex items-center gap-1"
              style={{ color: GOLD }}
            >
              <span className="size-1.5 rounded-full bg-emerald-400 animate-pulse" /> HYBRID
              PRECISION ENGINE
            </span>
          </div>

          <div className="ml-auto flex items-center gap-2.5 flex-wrap">
            {/* Year Mode toggle */}
            <button
              onClick={() => {
                setYearMode((y) => !y);
                setOursFile(null); setOursFiles([]);
                setPartnerFile(null); setPartnerFiles([]);
                setRawOurs(null); setRawPartner(null);
                setResult(null); setMonthBreakdown([]);
              }}
              className={`flex items-center gap-1.5 rounded-xl border px-3 py-2 text-[10px] font-bold transition-all ${yearMode ? "border-amber-400/60 bg-amber-400/20 text-amber-200" : "border-white/20 bg-white/5 text-white/60 hover:bg-white/15"}`}
              title={yearMode ? "Switch to single-file mode" : "Switch to 1-Year multi-month mode"}
            >
              <Calendar className="size-3.5" />
              {yearMode ? "1-Year Mode ✓" : "1-Year Mode"}
            </button>
            {yearMode ? (
              <>
                <HeaderMultiChip label="Our Ledger" files={oursFiles} onChange={setOursFiles} />
                <HeaderMultiChip label="Partner Ledger" files={partnerFiles} onChange={setPartnerFiles} />
              </>
            ) : (
              <>
                <HeaderChip label="Our Ledger" file={oursFile} onChange={(f) => selectFile("ours", f)} />
                <HeaderChip label="Partner Ledger" file={partnerFile} onChange={(f) => selectFile("partner", f)} />
              </>
            )}
            {result && (
              <>
                <button
                  onClick={exportExcel}
                  className="flex items-center gap-2 rounded-xl border border-amber-400/40 bg-amber-400/15 px-3.5 py-2.5 text-xs font-bold text-white hover:bg-amber-400/25 transition-all"
                  title="Download colour-coded multi-sheet Excel report"
                >
                  <FileSpreadsheet className="size-4" style={{ color: GOLD }} /> Excel
                </button>
                <button
                  onClick={exportPdf}
                  className="flex items-center gap-2 rounded-xl border border-rose-400/40 bg-rose-400/15 px-3.5 py-2.5 text-xs font-bold text-white hover:bg-rose-400/25 transition-all"
                  title="Download colour-coded PDF report"
                >
                  <FileText className="size-4 text-rose-300" /> PDF
                </button>
                <button
                  onClick={exportCsv}
                  className="flex items-center gap-2 rounded-xl border border-white/20 bg-white/10 px-3.5 py-2.5 text-xs font-semibold text-white hover:bg-white/20 transition-all"
                  title="Download flat CSV"
                >
                  <Download className="size-4" /> CSV
                </button>
              </>
            )}
            <button
              onClick={runSmartRecon}
              disabled={yearMode
                ? (
                    (oursUploadType === "multi" ? oursFiles.length === 0 : !oursFile) ||
                    (partnerUploadType === "multi" ? partnerFiles.length === 0 : !partnerFile)
                  )
                : (!oursFile || !partnerFile)
                || busy}
              className="group relative overflow-hidden rounded-xl px-5 py-2.5 text-sm font-bold shadow-lg transition-all disabled:opacity-40 active:scale-95"
              style={{ background: `linear-gradient(90deg, #d4af37, ${GOLD})`, color: NAVY }}
            >
              <div className="flex items-center gap-2">
                {busy ? (
                  <>
                    <span className="size-4 rounded-full border-2 border-[#0c2e5f]/30 border-t-[#0c2e5f] animate-spin" />
                    <span className="text-xs">{aiStatus || "Processing…"}</span>
                  </>
                ) : (
                  <>
                    <Sparkles className="size-4 group-hover:rotate-12 transition-transform" />
                    <span>Smart Reconcile</span>
                  </>
                )}
              </div>
            </button>
          </div>
        </div>
        {error && (
          <div className="bg-rose-500/90 text-white text-sm px-6 py-2.5 flex items-center gap-2">
            <AlertCircle className="size-4" /> {error}
          </div>
        )}
      </header>

      <main className="mx-auto max-w-[1600px] px-6 py-8 space-y-7">
        {/* ---------------- UPLOAD HERO (centered, pre-result) ---------------- */}
        {!result && (
          <UploadHero
            oursFile={oursFile}
            oursFiles={oursFiles}
            oursUploadType={oursUploadType}
            partnerFile={partnerFile}
            partnerFiles={partnerFiles}
            partnerUploadType={partnerUploadType}
            onPick={selectFile}
            onOursFilesChange={setOursFiles}
            onPartnerFilesChange={setPartnerFiles}
            onOursUploadTypeChange={setOursUploadType}
            onPartnerUploadTypeChange={setPartnerUploadType}
            onRun={runSmartRecon}
            busy={busy}
            yearMode={yearMode}
            onToggleYearMode={() => {
              setYearMode((y) => !y);
              setOursFile(null); setOursFiles([]);
              setPartnerFile(null); setPartnerFiles([]);
              setRawOurs(null); setRawPartner(null);
              // Our Ledger = month-wise GDS files (multi); Partner = the annual
              // supplier statement, one file (single) — matches the real workflow.
              setOursUploadType("multi"); setPartnerUploadType("single");
              setResult(null); setMonthBreakdown([]);
            }}
          />
        )}

        {/* ---------------- SOURCE FILES (full upload preview) ---------------- */}
        {hasData && (
          <section className="rounded-2xl border border-slate-200/70 bg-white shadow-sm overflow-hidden">
            <button
              onClick={() => setShowSource((s) => !s)}
              className="w-full flex items-center gap-3 px-6 py-4 hover:bg-slate-50 transition-colors"
            >
              <Table2 className="size-4" style={{ color: NAVY }} />
              <span className="text-sm font-bold text-slate-700">Uploaded Source Files</span>
              <span className="text-[11px] font-semibold text-slate-400">
                {rawOurs ? `Ours: ${Math.max(0, rawOurs.length - 1)} rows` : "Ours: —"} ·{" "}
                {rawPartner ? `Partner: ${Math.max(0, rawPartner.length - 1)} rows` : "Partner: —"}
              </span>
              <ChevronDown
                className={`ml-auto size-4 text-slate-400 transition-transform ${showSource ? "rotate-180" : ""}`}
              />
            </button>
            {showSource && (
              <div className="grid gap-4 px-6 pb-6 lg:grid-cols-2">
                <SourceTable title="Our Ledger" file={oursFile} aoa={rawOurs} accent={NAVY} />
                <SourceTable
                  title="Partner Ledger"
                  file={partnerFile}
                  aoa={rawPartner}
                  accent={GOLD}
                />
              </div>
            )}
          </section>
        )}

        {result && (
          <>
            {/* ---------------- YEAR MODE MONTH SELECTOR (top-level, most prominent) ---------------- */}
            {yearMode && monthBreakdown.length > 0 && (
              <MonthSelectorBar
                breakdown={monthBreakdown}
                selected={monthFilter}
                onSelect={(m) => { setMonthFilter(m); if (filter !== "fullledger") setFilter("all"); }}
              />
            )}

            {/* ---- YEAR MODE: how amounts compare (rate-based, no currency assumptions) ---- */}
            {yearMode && (
              <div className="flex items-start gap-2 rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-[11px] text-amber-800">
                <span className="text-base leading-none mt-0.5">ℹ️</span>
                <span>
                  <strong>How amounts compare:</strong> your ledger and the supplier statement record amounts on different bases, so the raw numbers won't match — bookings are matched by <strong>ticket number / PNR</strong>, not by amount.
                  {totals && totals.impliedRate > 0 ? (
                    <> {" "}The app <strong>auto-detected</strong> that the supplier amount is about{" "}
                      <strong>{totals.impliedRate.toFixed(2)}× your amount</strong> on average (a currency/markup factor — no currency setting needed). The{" "}
                      <strong>Variance</strong> column compares each booking to this rate:{" "}
                      <span className="text-emerald-700 font-bold">✓ on rate</span> = priced as expected, and a red{" "}
                      <span className="text-rose-700 font-bold">%</span> = the supplier charged that much more / less than expected — check those in the <strong>Price Looks Off</strong> tab.</>
                  ) : (
                    <> {" "}The <strong>Variance</strong> column shows the difference per booking.</>
                  )}
                </span>
              </div>
            )}

            {/* ---------------- SUMMARY STRIP ---------------- */}
            <section className="grid gap-3 grid-cols-2 sm:grid-cols-3 lg:grid-cols-5">
              <SummaryCard
                label="Total Rows"
                value={monthPairs.length}
                onClick={() => setFilter("all")}
                active={filter === "all"}
                accent={NAVY}
              />
              <SummaryCard
                label="Matched"
                value={totals?.matched ?? 0}
                onClick={() => setFilter("matched")}
                active={filter === "matched"}
                accent="#10b981"
              />
              <SummaryCard
                label="Only Ours"
                value={totals?.onlyOurs ?? 0}
                onClick={() => setFilter("missing_partner")}
                active={filter === "missing_partner"}
                accent="#6366f1"
              />
              <SummaryCard
                label="Only Partner"
                value={totals?.onlyPartner ?? 0}
                onClick={() => setFilter("missing_ours")}
                active={filter === "missing_ours"}
                accent="#ef4444"
              />
              <SummaryCard
                label="Needs Review"
                value={totals?.needsReview ?? 0}
                onClick={() => setFilter("review")}
                active={filter === "review"}
                accent={GOLD}
              />
            </section>

            {/* ---------------- KPI ROW ---------------- */}
            <section className="grid gap-5 sm:grid-cols-2 lg:grid-cols-4">
              <RingCard
                title="Match Rate"
                value={matchRate}
                caption={`${monthPairs.filter((p) => p.ours && p.partner).length} of ${monthPairs.length} rows paired`}
                icon={<CheckCircle2 className="size-4 text-emerald-500" />}
                color="#10b981"
              />
              <RingCard
                title="Avg Confidence"
                value={totals?.avgConfidence ?? 0}
                caption={confLabel(totals?.avgConfidence ?? 0) + " certainty"}
                icon={<ShieldCheck className="size-4" style={{ color: NAVY }} />}
                color={confColor(totals?.avgConfidence ?? 0)}
              />
              <StatCard
                title="Matched Value"
                value={money(matchedValue)}
                caption="Verified on both ledgers"
                icon={<TrendingUp className="size-4 text-emerald-500" />}
                tone="emerald"
              />
              <StatCard
                title="Needs Review"
                value={String(totals?.needsReview ?? 0)}
                caption={
                  totals && totals.aiAssisted > 0
                    ? `${totals.aiAssisted} AI-assisted matches`
                    : "Low-confidence pairs"
                }
                icon={<Cpu className="size-4" style={{ color: GOLD }} />}
                tone="amber"
                onClick={() => setFilter("review")}
              />
            </section>

            {/* ---------------- ANALYTICS ---------------- */}
            {isClient && (
              <section className="grid gap-5 lg:grid-cols-3">
                <div className="rounded-2xl border border-slate-200/70 bg-white p-6 shadow-sm">
                  <div className="flex items-center justify-between mb-5">
                    <h3 className="text-xs font-bold uppercase tracking-wider text-slate-400">
                      Distribution
                    </h3>
                    <span
                      className="text-[10px] font-bold px-2 py-1 rounded-md text-white"
                      style={{ background: NAVY }}
                    >
                      {engineMode === "ai" ? "AI + Rules" : "Heuristic"}
                    </span>
                  </div>
                  <div className="h-[200px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={chartData} barCategoryGap="30%">
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis
                          dataKey="name"
                          axisLine={false}
                          tickLine={false}
                          tick={{ fontSize: 11, fill: "#64748b" }}
                        />
                        <YAxis
                          axisLine={false}
                          tickLine={false}
                          tick={{ fontSize: 11, fill: "#64748b" }}
                          allowDecimals={false}
                        />
                        <Tooltip
                          contentStyle={{
                            borderRadius: 12,
                            border: "none",
                            boxShadow: "0 8px 24px -6px rgb(0 0 0 / 0.15)",
                          }}
                          cursor={{ fill: "#f8fafc" }}
                        />
                        <Bar dataKey="value" radius={[8, 8, 0, 0]}>
                          {chartData.map((e, i) => (
                            <Cell key={i} fill={e.color} />
                          ))}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="rounded-2xl border border-slate-200/70 bg-white p-6 shadow-sm">
                  <h3 className="text-xs font-bold uppercase tracking-wider text-slate-400 mb-5">
                    Confidence Spread
                  </h3>
                  <div className="h-[200px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={confHist} barCategoryGap="25%">
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis
                          dataKey="name"
                          axisLine={false}
                          tickLine={false}
                          tick={{ fontSize: 11, fill: "#64748b" }}
                          unit="%"
                        />
                        <YAxis
                          axisLine={false}
                          tickLine={false}
                          tick={{ fontSize: 11, fill: "#64748b" }}
                          allowDecimals={false}
                        />
                        <Tooltip
                          contentStyle={{
                            borderRadius: 12,
                            border: "none",
                            boxShadow: "0 8px 24px -6px rgb(0 0 0 / 0.15)",
                          }}
                          cursor={{ fill: "#f8fafc" }}
                        />
                        <Bar dataKey="value" radius={[8, 8, 0, 0]}>
                          {confHist.map((e, i) => (
                            <Cell key={i} fill={e.color} />
                          ))}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="rounded-2xl border border-slate-200/70 bg-white p-6 shadow-sm flex flex-col">
                  <h3 className="text-xs font-bold uppercase tracking-wider text-slate-400 mb-3">
                    Quality Split
                  </h3>
                  <div className="relative flex-1 min-h-[150px]">
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <Pie
                          data={chartData}
                          innerRadius={50}
                          outerRadius={72}
                          paddingAngle={4}
                          dataKey="value"
                          stroke="none"
                        >
                          {chartData.map((e, i) => (
                            <Cell key={i} fill={e.color} />
                          ))}
                        </Pie>
                        <Tooltip
                          contentStyle={{
                            borderRadius: 12,
                            border: "none",
                            boxShadow: "0 8px 24px -6px rgb(0 0 0 / 0.15)",
                          }}
                        />
                      </PieChart>
                    </ResponsiveContainer>
                    <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                      <span className="text-2xl font-black text-slate-800">{pct(matchRate)}</span>
                      <span className="text-[9px] text-slate-400 font-bold uppercase">Paired</span>
                    </div>
                  </div>
                  <div className="mt-2 space-y-1.5">
                    {chartData.map((e) => (
                      <div key={e.name} className="flex items-center gap-2 text-[11px]">
                        <span className="size-2 rounded-full" style={{ background: e.color }} />
                        <span className="text-slate-500 font-medium">{e.name}</span>
                        <span className="ml-auto font-bold text-slate-700">{e.value}</span>
                      </div>
                    ))}
                  </div>
                </div>
              </section>
            )}

            {/* ---------------- AI MAPPING ---------------- */}
            {schema && (
              <section className="grid gap-5 md:grid-cols-2">
                <AiMappingCard title="Our Ledger Mapping" mapping={schema.ours} schema={schema} />
                <AiMappingCard
                  title="Partner Ledger Mapping"
                  mapping={schema.partner}
                  schema={schema}
                />
              </section>
            )}

            {/* ---------------- TOTALS ---------------- */}
            <section className="grid gap-5 md:grid-cols-2 lg:grid-cols-3">
              <TotalsCard
                title="Internal Ledger"
                rows={totals?.oursRows ?? 0}
                charges={totals?.oursCharges ?? 0}
                credits={totals?.oursCredits ?? 0}
                icon={<CheckCircle2 className="size-4" style={{ color: NAVY }} />}
              />
              <TotalsCard
                title="Partner Ledger"
                rows={totals?.partnerRows ?? 0}
                charges={totals?.partnerCharges ?? 0}
                credits={totals?.partnerCredits ?? 0}
                icon={<TrendingUp className="size-4 text-emerald-500" />}
              />
              <div
                className="rounded-2xl border border-slate-200/70 p-6 shadow-sm"
                style={{ background: "linear-gradient(135deg, rgba(12,46,95,0.04), white)" }}
              >
                <div className="flex items-center gap-2 text-[11px] font-bold uppercase tracking-widest text-slate-400 mb-5">
                  <AlertCircle className="size-4" style={{ color: GOLD }} /> Discrepancy Analysis
                </div>
                <div className="space-y-4">
                  <DiffRow label="Net Amount Diff" value={totals?.netAmountDiff ?? 0} />
                  <DiffRow label="Amount Mismatches" value={totals?.amountIssues ?? 0} raw />
                  <div className="h-px bg-slate-100" />
                  <div
                    className="flex justify-between items-center text-white p-3 rounded-xl shadow-lg"
                    style={{ background: `linear-gradient(90deg, ${NAVY}, #103a73)` }}
                  >
                    <div className="text-[10px] font-bold uppercase opacity-80">
                      Unmatched Items
                    </div>
                    <div className="text-xl font-black">
                      {(totals?.onlyOurs ?? 0) + (totals?.onlyPartner ?? 0)}
                    </div>
                  </div>
                </div>
              </div>
            </section>

            {/* ---------------- TABLE + DETAIL ---------------- */}
            <section
              className={`grid gap-5 ${filter === "fullledger" ? "grid-cols-1" : "lg:grid-cols-[1fr_400px]"}`}
            >
              <div className="space-y-4">
                <div className="flex flex-wrap items-center gap-1.5 bg-white p-2 rounded-2xl border border-slate-200/70 shadow-sm">
                  {(
                    (() => {
                      // All tab counts are scoped to the active month (monthPairs)
                      // so they match the table you're looking at.
                      const c = (pred: (p: Pair) => boolean) => monthPairs.filter(pred).length;
                      const scen = (p: Pair) => p.ours?.scenario ?? p.partner?.scenario;
                      return [
                        ["all", "All", monthPairs.length],
                        ["matched", "Matched", c((p) => p.status === "matched")],
                        ["amount_diff", "Amount Not Same", c((p) => p.status === "amount_diff")],
                        ...(totals && totals.impliedRate
                          ? [[
                              "price_off",
                              "Price Looks Off",
                              c((p) => {
                                const d = rateDeviation(p, totals.impliedRate);
                                return d !== null && Math.abs(d) > RATE_OFF_THRESHOLD;
                              }),
                            ]] as Array<[StatusFilter, string, number]>
                          : []),
                        ["payments", "Bank Transfers", c((p) => isTransfer(p.ours) || isTransfer(p.partner))],
                        ["security_deposit", "Security Deposit", c((p) => scen(p) === "security_deposit")],
                        ["refunds", "Refunds (Money Back)", c((p) => {
                          const sc = scen(p);
                          return !!sc && ["wrong_invoice", "wrong_client", "duplicate", "refund"].includes(sc);
                        })],
                        ["multi_passenger", "Group (Many People)", c((p) => scen(p) === "multi_passenger")],
                        ["duplicates", "Same Entry Twice", c((p) =>
                          (p.ours?.duplicateCount ?? 0) > 1 || (p.partner?.duplicateCount ?? 0) > 1)],
                        ["missing_partner", "Only Ours", c((p) => p.status === "missing_partner")],
                        ["missing_ours", "Only Partner", c((p) => p.status === "missing_ours")],
                        ["review", "Needs Review", c((p) => !!p.needsReview)],
                        [
                          "fullledger",
                          "Both Sheets (Full)",
                          monthPairs.reduce((acc, p) => acc + (p.ours ? 1 : 0) + (p.partner ? 1 : 0), 0),
                        ],
                      ] as Array<[StatusFilter, string, number]>;
                    })()
                  ).map(([k, label, count]) => (
                    <button
                      key={k}
                      onClick={() => setFilter(k)}
                      className={`text-[11px] px-3 py-2 rounded-lg font-bold transition-all ${
                        filter === k
                          ? k === "review"
                            ? "bg-amber-500 text-white shadow-md"
                            : "text-white shadow-md"
                          : k === "review" && count > 0
                            ? "text-amber-600 hover:bg-amber-50"
                            : k === "fullledger"
                              ? "text-slate-600 hover:bg-slate-50"
                              : "text-slate-500 hover:bg-slate-50"
                      }`}
                      style={filter === k && k !== "review" ? { background: NAVY } : undefined}
                    >
                      {label} <span className="ml-0.5 opacity-60">({count})</span>
                    </button>
                  ))}
                  {filter !== "fullledger" && (
                    <div className="ml-auto flex items-center gap-2 pl-3 border-l border-slate-200">
                      {/* Active month badge */}
                      {monthFilter !== "all" && (
                        <span
                          className="flex items-center gap-1 text-[9px] font-black px-2 py-1 rounded-lg cursor-pointer hover:opacity-80"
                          style={{ background: NAVY, color: GOLD }}
                          onClick={() => { setMonthFilter("all"); setFilter("all"); }}
                          title="Click to clear month filter"
                        >
                          <Calendar className="size-2.5" />
                          {monthLabel(monthFilter)} ✕
                        </span>
                      )}
                      <button
                        onClick={() => setSortByConf((s) => !s)}
                        className={`text-[10px] px-2.5 py-1.5 rounded-lg font-bold transition-all ${
                          sortByConf ? "bg-slate-800 text-white" : "text-slate-400 hover:bg-slate-50"
                        }`}
                        title="Sort lowest confidence first"
                      >
                        ↑ Confidence
                      </button>
                      <div className="flex items-center gap-1.5">
                        <Search className="size-3.5 text-slate-300" />
                        <input
                          value={query}
                          onChange={(e) => setQuery(e.target.value)}
                          placeholder="Filter…"
                          className="min-w-[140px] bg-transparent text-sm focus:outline-none"
                        />
                      </div>
                    </div>
                  )}
                </div>

                {filter === "payments" ? (
                  <PaymentFinderView
                    pairs={filteredPairs}
                    onSelect={setSelected}
                    selected={selected}
                  />
                ) : filter === "fullledger" ? (
                  <FullLedgerView ours={rawOurs} partner={rawPartner} result={result} pairs={monthPairs} />
                ) : (
                  <PairsTable
                    pairs={filteredPairs}
                    onSelect={setSelected}
                    selected={selected}
                    rawOurs={rawOurs}
                    rawPartner={rawPartner}
                    yearMode={yearMode}
                    impliedRate={totals?.impliedRate ?? 0}
                  />
                )}
              </div>

              {filter !== "fullledger" && <DetailPanel pair={selected} />}
            </section>
          </>
        )}
      </main>

      <footer className="mx-auto max-w-[1600px] px-6 py-6 text-center text-[11px] text-slate-400">
        Navvi Saadi Travel &amp; Tourism · AI Ledger Reconciliation · Hybrid Precision Engine
        <span className="mx-2 text-slate-300">·</span>
        <span className="font-mono text-[10px] font-bold text-slate-500">{BUILD_TAG}</span>
      </footer>
    </div>
  );
}

/* ================================================================== */
/*  UPLOAD                                                             */
/* ================================================================== */

/** One side's upload panel in Year Mode — toggle single/multi + month subcategory list */
function YearSideUploadPanel({
  label, accentColor, uploadType, onUploadTypeChange,
  singleFile, onSingleFile, multiFiles, onMultiFiles,
}: {
  label: string;
  accentColor: string;
  uploadType: "single" | "multi";
  onUploadTypeChange: (t: "single" | "multi") => void;
  singleFile: File | null;
  onSingleFile: (f: File | null) => void;
  multiFiles: File[];
  onMultiFiles: (f: File[]) => void;
}) {
  const [dragging, setDragging] = useState(false);
  const inputRef = React.useRef<HTMLInputElement>(null);
  const inputMultiRef = React.useRef<HTMLInputElement>(null);

  const addMultiFiles = (incoming: File[]) => {
    const existing = new Set(multiFiles.map((f) => f.name));
    const merged = [...multiFiles, ...incoming.filter((f) => !existing.has(f.name))];
    merged.sort((a, b) => monthFromFilename(a.name).localeCompare(monthFromFilename(b.name)));
    onMultiFiles(merged);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragging(false);
    const dropped = Array.from(e.dataTransfer.files).filter((f) => /\.(xlsx?|csv|tsv|txt)$/i.test(f.name));
    if (uploadType === "single") { onSingleFile(dropped[0] ?? null); }
    else addMultiFiles(dropped);
  };

  const ready = uploadType === "single" ? !!singleFile : multiFiles.length > 0;

  return (
    <div className="flex flex-col gap-0 rounded-2xl border overflow-hidden shadow-sm"
      style={{ borderColor: ready ? `${accentColor}50` : "#e2e8f0" }}>

      {/* Header bar */}
      <div className="flex items-center gap-3 px-4 py-3 border-b" style={{ background: `${accentColor}0d`, borderColor: `${accentColor}20` }}>
        <span className="size-2.5 rounded-full" style={{ background: accentColor }} />
        <span className="text-xs font-black uppercase tracking-wider text-slate-700">{label}</span>
        {ready && (
          <span className="ml-auto text-[9px] font-black px-2 py-0.5 rounded-md text-white" style={{ background: accentColor }}>
            {uploadType === "single" ? "1 file" : `${multiFiles.length} file${multiFiles.length > 1 ? "s" : ""}`}
          </span>
        )}
      </div>

      {/* Single / Multi toggle */}
      <div className="flex border-b border-slate-100">
        {(["single", "multi"] as const).map((t) => (
          <button
            key={t}
            onClick={() => onUploadTypeChange(t)}
            className={`flex-1 flex items-center justify-center gap-1.5 py-2.5 text-[11px] font-bold transition-all ${uploadType === t ? "text-white" : "text-slate-400 hover:text-slate-600 bg-slate-50 hover:bg-slate-100"}`}
            style={uploadType === t ? { background: accentColor } : undefined}
          >
            {t === "single" ? <FileSpreadsheet className="size-3" /> : <Calendar className="size-3" />}
            {t === "single" ? "Single File" : "Multiple Files (Month-wise)"}
          </button>
        ))}
      </div>

      {/* Upload body */}
      <div
        className="p-4 transition-colors"
        style={{ background: dragging ? `${accentColor}08` : "white" }}
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={handleDrop}
      >
        {uploadType === "single" ? (
          /* ── Single file drop zone ── */
          <label
            className="flex flex-col items-center justify-center gap-2 rounded-xl border-2 border-dashed p-5 cursor-pointer transition-all"
            style={{ borderColor: singleFile ? accentColor : dragging ? accentColor : "#cbd5e1", background: singleFile ? `${accentColor}06` : "transparent" }}
          >
            <div className="size-10 rounded-xl flex items-center justify-center shadow-sm" style={{ background: singleFile ? accentColor : "#f1f5f9" }}>
              {singleFile ? <CheckCircle2 className="size-5 text-white" /> : <UploadCloud className="size-5 text-slate-400" />}
            </div>
            {singleFile ? (
              <div className="text-center">
                <div className="text-xs font-black text-slate-800 truncate max-w-[180px]">{singleFile.name}</div>
                <div className="text-[10px] text-slate-400 mt-0.5">{Math.round(singleFile.size / 1024)} KB</div>
              </div>
            ) : (
              <div className="text-center">
                <div className="text-xs font-bold text-slate-600">Drag & drop or click</div>
                <div className="text-[10px] text-slate-400">Single ledger file</div>
              </div>
            )}
            <input ref={inputRef} type="file" accept=".xls,.xlsx,.csv,.tsv,.txt" className="hidden"
              onChange={(e) => { onSingleFile(e.target.files?.[0] ?? null); e.target.value = ""; }} />
          </label>
        ) : (
          /* ── Multi-file: explicit 12-month slot grid ── */
          (() => {
            const monthOf = (f: File) => monthFromFilename(f.name); // "YYYY-MM" or ""
            const slotFile: Record<string, File | undefined> = {};
            for (const mm of MONTH_SLOTS) {
              slotFile[mm] = multiFiles.find((f) => (monthOf(f).split("-")[1] ?? "") === mm);
            }
            const shown = new Set(Object.values(slotFile).filter(Boolean) as File[]);
            const extraFiles = multiFiles.filter((f) => !shown.has(f)); // unknown month or duplicate
            const filled = MONTH_SLOTS.filter((mm) => slotFile[mm]).length;
            return (
              <div className="flex flex-col gap-3">
                {/* Bulk drop zone — drop all months at once (optional) */}
                <label
                  className="flex items-center gap-3 rounded-xl border-2 border-dashed px-4 py-2.5 cursor-pointer transition-all"
                  style={{ borderColor: dragging ? accentColor : multiFiles.length > 0 ? `${accentColor}60` : "#cbd5e1", background: dragging ? `${accentColor}06` : "transparent" }}
                >
                  <div className="size-8 rounded-lg flex items-center justify-center shrink-0" style={{ background: multiFiles.length > 0 ? accentColor : "#f1f5f9" }}>
                    <UploadCloud className="size-4" style={{ color: multiFiles.length > 0 ? "white" : "#94a3b8" }} />
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className="text-xs font-bold text-slate-700">
                      {filled > 0 ? `${filled} of 12 months uploaded` : "Drop all months at once — or upload each below"}
                    </div>
                    <div className="text-[10px] text-slate-400">One file per month · auto-sorted by date</div>
                  </div>
                  <input ref={inputMultiRef} type="file" accept=".xls,.xlsx,.csv,.tsv,.txt" className="hidden" multiple
                    onChange={(e) => { addMultiFiles(Array.from(e.target.files ?? [])); e.target.value = ""; }} />
                </label>

                {/* 12 month slots — upload each month manually */}
                <div className="grid grid-cols-3 gap-2">
                  {MONTH_SLOTS.map((mm) => {
                    const f = slotFile[mm];
                    const color = MONTH_COLORS[mm] ?? "#64748b";
                    const abbr = MONTH_ABBR[mm] ?? "??";
                    const yr = f ? (monthOf(f).split("-")[0] ?? "") : "";
                    return (
                      <label
                        key={mm}
                        className="relative flex flex-col items-center justify-center gap-1 rounded-xl border-2 border-dashed p-2 cursor-pointer transition-all min-h-[64px] text-center"
                        style={{ borderColor: f ? color : "#e2e8f0", background: f ? `${color}0d` : "transparent" }}
                        title={f ? f.name : `Upload ${abbr} file`}
                      >
                        <span className="text-[9px] font-black uppercase px-1.5 py-0.5 rounded text-white" style={{ background: f ? color : "#cbd5e1" }}>
                          {abbr}{yr ? ` '${yr.slice(2)}` : ""}
                        </span>
                        {f ? (
                          <>
                            <span className="text-[9px] font-semibold text-slate-600 truncate max-w-full px-1 leading-tight">{f.name}</span>
                            <button
                              onClick={(e) => { e.preventDefault(); onMultiFiles(multiFiles.filter((x) => x.name !== f.name)); }}
                              className="absolute -top-1.5 -right-1.5 size-4 rounded-full bg-white border border-slate-200 flex items-center justify-center text-slate-400 hover:text-rose-500 hover:border-rose-300 text-[10px] shadow-sm"
                              title="Remove"
                            >×</button>
                          </>
                        ) : (
                          <span className="text-[9px] text-slate-400 flex items-center gap-0.5"><UploadCloud className="size-3" /> add</span>
                        )}
                        <input type="file" accept=".xls,.xlsx,.csv,.tsv,.txt" className="hidden"
                          onChange={(e) => { const sel = e.target.files?.[0]; if (sel) addMultiFiles([sel]); e.target.value = ""; }} />
                      </label>
                    );
                  })}
                </div>

                {/* Files whose month couldn't be read from the filename (still reconciled) */}
                {extraFiles.length > 0 && (
                  <div className="rounded-xl border border-amber-200 bg-amber-50 px-3 py-2">
                    <div className="text-[10px] font-bold text-amber-700 mb-1">
                      ⚠ Month not detected from these file names — they'll still be reconciled, but won't sit in a month slot:
                    </div>
                    {extraFiles.map((f) => (
                      <div key={f.name} className="flex items-center gap-2 py-0.5">
                        <FileSpreadsheet className="size-3 shrink-0 text-amber-500" />
                        <span className="flex-1 text-[10px] text-slate-700 truncate">{f.name}</span>
                        <button onClick={() => onMultiFiles(multiFiles.filter((x) => x.name !== f.name))} className="text-slate-300 hover:text-rose-500 text-xs">×</button>
                      </div>
                    ))}
                  </div>
                )}

                {multiFiles.length > 0 && (
                  <div className="flex items-center justify-between px-1">
                    <span className="text-[10px] text-slate-400">{multiFiles.length} file{multiFiles.length > 1 ? "s" : ""} ready</span>
                    <button onClick={() => onMultiFiles([])} className="text-[10px] font-bold text-rose-400 hover:text-rose-600 transition-colors">Clear all</button>
                  </div>
                )}
              </div>
            );
          })()
        )}
      </div>
    </div>
  );
}

function UploadHero({
  oursFile, oursFiles, oursUploadType,
  partnerFile, partnerFiles, partnerUploadType,
  onPick, onOursFilesChange, onPartnerFilesChange,
  onOursUploadTypeChange, onPartnerUploadTypeChange,
  onRun, busy, yearMode, onToggleYearMode,
}: {
  oursFile: File | null;
  oursFiles: File[];
  oursUploadType: "single" | "multi";
  partnerFile: File | null;
  partnerFiles: File[];
  partnerUploadType: "single" | "multi";
  onPick: (side: "ours" | "partner", f: File | null) => void;
  onOursFilesChange: (files: File[]) => void;
  onPartnerFilesChange: (files: File[]) => void;
  onOursUploadTypeChange: (t: "single" | "multi") => void;
  onPartnerUploadTypeChange: (t: "single" | "multi") => void;
  onRun: () => void;
  busy: boolean;
  yearMode: boolean;
  onToggleYearMode: () => void;
}) {
  const oursReady = yearMode
    ? (oursUploadType === "multi" ? oursFiles.length > 0 : !!oursFile)
    : !!oursFile;
  const partnerReady = yearMode
    ? (partnerUploadType === "multi" ? partnerFiles.length > 0 : !!partnerFile)
    : !!partnerFile;
  const canRun = oursReady && partnerReady;

  // Month coverage matrix (Year Mode, multi on at least one side)
  const ourMonths = Array.from(new Set(
    (oursUploadType === "multi" ? oursFiles : oursFile ? [oursFile] : [])
      .map((f) => monthFromFilename(f.name)).filter(Boolean)
  )).sort();
  const partnerMonths = Array.from(new Set(
    (partnerUploadType === "multi" ? partnerFiles : partnerFile ? [partnerFile] : [])
      .map((f) => monthFromFilename(f.name)).filter(Boolean)
  )).sort();
  const allMonths = Array.from(new Set([...ourMonths, ...partnerMonths])).sort();

  return (
    <div className="rounded-3xl border border-slate-200/70 bg-white shadow-sm overflow-hidden">
      <div className="p-8">
        <div className="text-center max-w-2xl mx-auto">
          <div
            className="mx-auto size-16 rounded-2xl flex items-center justify-center mb-5 shadow-xl"
            style={{ background: `linear-gradient(135deg, ${NAVY}, #103a73)` }}
          >
            <Brain className="size-8" style={{ color: GOLD }} />
          </div>
          <h1 className="text-3xl font-black tracking-tight text-slate-800">
            AI Financial Reconciliation
          </h1>
          <p className="mt-3 text-sm text-slate-500 leading-relaxed">
            {yearMode
              ? "1-Year Mode — each side lets you upload a single file or multiple month-wise files. Each month's files are shown as sub-categories below."
              : "Upload both statements — a hybrid AI engine maps columns, matches every entry on multiple signals, and scores each result by confidence."}
          </p>

          {/* Mode toggle */}
          <div className="mt-5 flex justify-center gap-3">
            <button
              onClick={onToggleYearMode}
              className={`flex items-center gap-2 px-4 py-2 rounded-xl border text-xs font-bold transition-all ${!yearMode ? "text-white shadow-md" : "border-slate-200 text-slate-500 hover:bg-slate-50"}`}
              style={!yearMode ? { background: NAVY } : undefined}
            >
              <FileSpreadsheet className="size-3.5" /> Single-File Mode
            </button>
            <button
              onClick={onToggleYearMode}
              className={`flex items-center gap-2 px-4 py-2 rounded-xl border text-xs font-bold transition-all ${yearMode ? "text-white shadow-md" : "border-slate-200 text-slate-500 hover:bg-slate-50"}`}
              style={yearMode ? { background: `linear-gradient(90deg, #d4af37, ${GOLD})`, color: NAVY } : undefined}
            >
              <Calendar className="size-3.5" /> 1-Year Mode (Multi-Month)
            </button>
          </div>
        </div>

        {/* Upload panels */}
        <div className="mt-8 grid gap-5 md:grid-cols-2 max-w-4xl mx-auto">
          {yearMode ? (
            <>
              <YearSideUploadPanel
                label="Our Ledger"
                accentColor={NAVY}
                uploadType={oursUploadType}
                onUploadTypeChange={onOursUploadTypeChange}
                singleFile={oursFile}
                onSingleFile={(f) => onPick("ours", f)}
                multiFiles={oursFiles}
                onMultiFiles={onOursFilesChange}
              />
              <YearSideUploadPanel
                label="Partner Ledger"
                accentColor="#c9a23a"
                uploadType={partnerUploadType}
                onUploadTypeChange={onPartnerUploadTypeChange}
                singleFile={partnerFile}
                onSingleFile={(f) => onPick("partner", f)}
                multiFiles={partnerFiles}
                onMultiFiles={onPartnerFilesChange}
              />
            </>
          ) : (
            <>
              <UploadZone label="Our Ledger" file={oursFile} onChange={(f) => onPick("ours", f)} accent={NAVY} />
              <UploadZone label="Partner Ledger" file={partnerFile} onChange={(f) => onPick("partner", f)} accent={GOLD} />
            </>
          )}
        </div>

        <div className="mt-7 flex justify-center">
          <button
            onClick={onRun}
            disabled={!canRun || busy}
            className="rounded-xl px-8 py-3 text-sm font-bold shadow-lg transition-all disabled:opacity-40 active:scale-95 flex items-center gap-2"
            style={{ background: `linear-gradient(90deg, #d4af37, ${GOLD})`, color: NAVY }}
          >
            <Sparkles className="size-4" />
            {yearMode
              ? `Reconcile ${allMonths.length > 1 ? allMonths.length + " Months" : "Full Year"}`
              : "Smart Reconcile"}
          </button>
        </div>

        <div className="mt-7 flex flex-wrap justify-center gap-6">
          {(yearMode
            ? ["Both Sides Flexible Upload", "Month Sub-categories", "Monthly Breakdown", "Annual Summary"]
            : ["Auto-Schema Detection", "Multi-Column Matching", "AI Residual Resolver", "Confidence Scoring"]
          ).map((feat) => (
            <div key={feat} className="flex items-center gap-2 text-[11px] font-bold text-slate-500 uppercase">
              <CheckCircle2 className="size-4 text-emerald-500" /> {feat}
            </div>
          ))}
        </div>
      </div>

      {/* ── Month Coverage Matrix (Year Mode) ── */}
      {yearMode && allMonths.length > 0 && (
        <div className="border-t border-slate-100 px-8 py-5" style={{ background: "linear-gradient(180deg,#f8fafc,white)" }}>
          <div className="flex items-center gap-2 mb-4">
            <Calendar className="size-4" style={{ color: NAVY }} />
            <span className="text-xs font-black uppercase tracking-wider text-slate-700">Month Coverage</span>
            <span className="text-[10px] font-bold px-2 py-0.5 rounded-md text-white ml-1" style={{ background: NAVY }}>
              {allMonths.length} month{allMonths.length > 1 ? "s" : ""}
            </span>
          </div>
          <div className="overflow-x-auto">
            <table className="text-[11px] w-full">
              <thead>
                <tr>
                  <th className="text-left pr-4 py-1 text-[10px] font-black uppercase text-slate-400 whitespace-nowrap">Side</th>
                  {allMonths.map((m) => {
                    const [yr, mo] = m.split("-");
                    return (
                      <th key={m} className="px-2 py-1 text-center text-[10px] font-black text-slate-500 whitespace-nowrap">
                        <div>{MONTH_ABBR[mo] ?? mo}</div>
                        <div className="font-normal text-slate-400">{yr.slice(2)}</div>
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {([
                  { lab: "Our Ledger", months: ourMonths, files: oursUploadType === "multi" ? oursFiles : oursFile ? [oursFile] : [], color: NAVY },
                  { lab: "Partner Ledger", months: partnerMonths, files: partnerUploadType === "multi" ? partnerFiles : partnerFile ? [partnerFile] : [], color: "#c9a23a" },
                ] as const).map(({ lab, months, files, color }) => (
                  <tr key={lab}>
                    <td className="pr-4 py-1.5 font-bold text-slate-600 whitespace-nowrap">
                      <div className="flex items-center gap-1.5">
                        <span className="size-2 rounded-full" style={{ background: color }} />
                        {lab}
                        <span className="text-[9px] text-slate-400 font-normal">({files.length} file{files.length !== 1 ? "s" : ""})</span>
                      </div>
                    </td>
                    {allMonths.map((m) => {
                      const has = months.includes(m);
                      const file = (files as File[]).find((f) => monthFromFilename(f.name) === m);
                      return (
                        <td key={m} className="px-2 py-1.5 text-center">
                          {has ? (
                            <span className="inline-flex items-center justify-center size-7 rounded-lg text-[9px] font-black text-white shadow-sm" style={{ background: color }} title={file?.name}>✓</span>
                          ) : (
                            <span className="inline-flex items-center justify-center size-7 rounded-lg text-[9px] text-slate-300 border border-dashed border-slate-200">—</span>
                          )}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {(() => {
            const missing = allMonths.filter((m) => !ourMonths.includes(m) || !partnerMonths.includes(m));
            if (!missing.length || ourMonths.length === 0 || partnerMonths.length === 0) return null;
            return (
              <div className="mt-3 flex items-start gap-2 rounded-xl bg-amber-50 border border-amber-200 px-3 py-2">
                <AlertCircle className="size-4 text-amber-500 shrink-0 mt-0.5" />
                <div className="text-[10px] text-amber-700 font-medium">
                  <strong>Coverage gap:</strong> {missing.map((m) => monthLabel(m)).join(", ")} — one side is missing. Those entries will show as "Only Ours" or "Only Partner".
                </div>
              </div>
            );
          })()}
        </div>
      )}
    </div>
  );
}

function UploadZone({
  label,
  file,
  onChange,
  accent,
}: {
  label: string;
  file: File | null;
  onChange: (f: File | null) => void;
  accent: string;
}) {
  return (
    <label
      className="group flex flex-col items-center justify-center gap-3 rounded-2xl border-2 border-dashed p-8 cursor-pointer transition-all hover:bg-slate-50"
      style={{ borderColor: file ? accent : "#cbd5e1" }}
    >
      <div
        className="size-12 rounded-xl flex items-center justify-center transition-transform group-hover:scale-110"
        style={{ background: file ? accent : "#f1f5f9", color: file ? "white" : "#94a3b8" }}
      >
        {file ? <FileSpreadsheet className="size-6" /> : <UploadCloud className="size-6" />}
      </div>
      <div className="text-center">
        <div className="text-xs font-bold uppercase tracking-widest text-slate-400">{label}</div>
        <div className="mt-1 text-sm font-bold text-slate-700 truncate max-w-[240px]">
          {file ? file.name : "Click to upload"}
        </div>
        <div className="text-[10px] text-slate-400 mt-0.5">.xlsx · .xls · .csv · .tsv</div>
      </div>
      <input
        type="file"
        accept=".xls,.xlsx,.csv,.tsv,.txt"
        className="hidden"
        onChange={(e) => onChange(e.target.files?.[0] ?? null)}
      />
    </label>
  );
}

/** Multi-file upload zone for year-mode monthly files */
/** Month color palette for the multi-upload file list */
const MONTH_COLORS: Record<string, string> = {
  "01": "#3b82f6", "02": "#8b5cf6", "03": "#10b981",
  "04": "#f59e0b", "05": "#ef4444", "06": "#06b6d4",
  "07": "#f97316", "08": "#84cc16", "09": "#6366f1",
  "10": "#ec4899", "11": "#14b8a6", "12": "#c9a23a",
};
const MONTH_ABBR: Record<string, string> = {
  "01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"May","06":"Jun",
  "07":"Jul","08":"Aug","09":"Sep","10":"Oct","11":"Nov","12":"Dec",
};
/** The 12 calendar months, as the upload grid renders one slot per month. */
const MONTH_SLOTS = ["01","02","03","04","05","06","07","08","09","10","11","12"];

function MultiUploadZone({
  label, accentColor, files, onChange,
}: {
  label: string;
  accentColor: string;
  files: File[];
  onChange: (f: File[]) => void;
}) {
  const [dragging, setDragging] = useState(false);

  const addFiles = (newFiles: File[]) => {
    const existing = new Set(files.map((f) => f.name));
    const merged = [...files, ...newFiles.filter((f) => !existing.has(f.name))];
    merged.sort((a, b) => monthFromFilename(a.name).localeCompare(monthFromFilename(b.name)));
    onChange(merged);
  };

  const handleInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    addFiles(Array.from(e.target.files ?? []));
    e.target.value = "";
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragging(false);
    addFiles(Array.from(e.dataTransfer.files).filter((f) =>
      /\.(xlsx?|csv|tsv|txt)$/i.test(f.name)
    ));
  };

  const allMonthsFound = files.length === 0 || files.every((f) => !!monthFromFilename(f.name));

  return (
    <div className="flex flex-col gap-2.5">
      {/* Drop zone */}
      <label
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={handleDrop}
        className="group relative flex flex-col items-center justify-center gap-3 rounded-2xl border-2 border-dashed p-7 cursor-pointer transition-all"
        style={{
          borderColor: dragging ? GOLD : files.length > 0 ? accentColor : "#cbd5e1",
          background: dragging ? `rgba(201,162,58,0.06)` : files.length > 0 ? `${accentColor}08` : "#fafafa",
        }}
      >
        <div
          className="size-14 rounded-2xl flex items-center justify-center shadow-lg transition-transform group-hover:scale-110"
          style={{ background: dragging ? `linear-gradient(135deg,${GOLD},#b08020)` : files.length > 0 ? accentColor : "#f1f5f9" }}
        >
          {files.length > 0
            ? <CheckCircle2 className="size-7 text-white" />
            : <UploadCloud className="size-7" style={{ color: dragging ? "white" : "#94a3b8" }} />
          }
        </div>

        <div className="text-center space-y-1">
          <div className="text-[10px] font-black uppercase tracking-widest" style={{ color: accentColor }}>
            {label} · Monthly Files
          </div>
          {files.length > 0 ? (
            <div className="text-sm font-black text-slate-800">
              {files.length} file{files.length > 1 ? "s" : ""} loaded
              {allMonthsFound ? " ✓" : " — some months undetected"}
            </div>
          ) : (
            <div className="text-sm font-bold text-slate-600">Drag & drop or click to select</div>
          )}
          <div className="text-[10px] text-slate-400">Select one file per month — sorted automatically</div>
        </div>
        <input type="file" accept=".xls,.xlsx,.csv,.tsv,.txt" className="hidden" multiple onChange={handleInput} />
      </label>

      {/* File list */}
      {files.length > 0 && (
        <div className="rounded-2xl border border-slate-200 overflow-hidden shadow-sm">
          <div className="flex items-center justify-between px-3 py-2 border-b border-slate-100" style={{ background: `${accentColor}0a` }}>
            <span className="text-[10px] font-black uppercase tracking-wider text-slate-500">
              {files.length} file{files.length > 1 ? "s" : ""} · sorted by month
            </span>
            <button onClick={() => onChange([])} className="text-[10px] font-bold text-rose-400 hover:text-rose-600 transition-colors">
              Clear all
            </button>
          </div>
          <div className="divide-y divide-slate-50 max-h-[220px] overflow-y-auto">
            {files.map((f, i) => {
              const mo = monthFromFilename(f.name);
              const moNum = mo.split("-")[1] ?? "";
              const color = MONTH_COLORS[moNum] ?? "#64748b";
              const abbr = MONTH_ABBR[moNum] ?? "??";
              const yr = mo.split("-")[0] ?? "";
              const kb = Math.round(f.size / 1024);
              return (
                <div key={f.name} className="flex items-center gap-3 px-3 py-2 hover:bg-slate-50 transition-colors">
                  <div
                    className="size-9 rounded-xl flex flex-col items-center justify-center shrink-0 shadow-sm"
                    style={{ background: `${color}18`, border: `1.5px solid ${color}40` }}
                  >
                    <span className="text-[8px] font-black uppercase" style={{ color }}>{abbr}</span>
                    <span className="text-[8px] font-bold text-slate-400">{yr.slice(2)}</span>
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className="text-[11px] font-bold text-slate-700 truncate">{f.name}</div>
                    <div className="text-[9px] text-slate-400">{kb} KB · {mo || "month not detected"}</div>
                  </div>
                  <button
                    onClick={() => onChange(files.filter((_, j) => j !== i))}
                    className="size-6 rounded-lg flex items-center justify-center text-slate-300 hover:text-rose-500 hover:bg-rose-50 transition-all shrink-0"
                    title="Remove"
                  >×</button>
                </div>
              );
            })}
          </div>
          <label className="flex items-center justify-center gap-2 px-3 py-2 bg-slate-50 border-t border-slate-100 cursor-pointer hover:bg-slate-100 transition-colors">
            <UploadCloud className="size-3.5 text-slate-400" />
            <span className="text-[10px] font-bold text-slate-500">Add more files</span>
            <input type="file" accept=".xls,.xlsx,.csv,.tsv,.txt" className="hidden" multiple onChange={handleInput} />
          </label>
        </div>
      )}
    </div>
  );
}

/**
 * Premium month selector — calendar-style cards with match-rate fill bars,
 * a dropdown for quick jump, and a both-sides stats strip for the selected month.
 */
function MonthSelectorBar({
  breakdown,
  selected,
  onSelect,
}: {
  breakdown: MonthlyBreakdown[];
  selected: string;
  onSelect: (m: string) => void;
}) {
  const active = selected !== "all" ? breakdown.find((b) => b.month === selected) : null;
  const activeIdx = breakdown.findIndex((b) => b.month === selected);

  const rateColor = (r: number) =>
    r >= 0.85 ? "#10b981" : r >= 0.6 ? "#f59e0b" : "#ef4444";
  const rateBg = (r: number) =>
    r >= 0.85 ? "#d1fae5" : r >= 0.6 ? "#fef3c7" : "#fee2e2";
  const rateText = (r: number) =>
    r >= 0.85 ? "text-emerald-700" : r >= 0.6 ? "text-amber-700" : "text-rose-700";

  return (
    <div
      className="rounded-2xl border border-slate-200/70 bg-white shadow-sm overflow-hidden"
      style={{ boxShadow: "0 2px 16px -4px rgba(12,46,95,0.10)" }}
    >
      {/* ── Header strip ─────────────────────────────────────────── */}
      <div className="flex items-center justify-between px-4 py-3 border-b border-slate-100"
        style={{ background: `linear-gradient(90deg, rgba(12,46,95,0.04), white)` }}>
        <div className="flex items-center gap-2">
          <div className="size-7 rounded-lg flex items-center justify-center" style={{ background: NAVY }}>
            <Calendar className="size-3.5 text-white" />
          </div>
          <div>
            <div className="text-[11px] font-black uppercase tracking-wider text-slate-700">
              Monthly Filter
            </div>
            <div className="text-[9px] text-slate-400 font-medium">
              Both sides filtered · Click a month to drill in
            </div>
          </div>
        </div>
        <div className="flex items-center gap-2">
          {/* Prev/Next navigation */}
          {selected !== "all" && (
            <div className="flex items-center gap-1">
              <button
                disabled={activeIdx <= 0}
                onClick={() => activeIdx > 0 && onSelect(breakdown[activeIdx - 1].month)}
                className="size-7 rounded-lg flex items-center justify-center text-slate-400 hover:bg-slate-100 disabled:opacity-30 disabled:cursor-not-allowed transition-colors"
                title="Previous month"
              >‹</button>
              <button
                disabled={activeIdx >= breakdown.length - 1}
                onClick={() => activeIdx < breakdown.length - 1 && onSelect(breakdown[activeIdx + 1].month)}
                className="size-7 rounded-lg flex items-center justify-center text-slate-400 hover:bg-slate-100 disabled:opacity-30 disabled:cursor-not-allowed transition-colors"
                title="Next month"
              >›</button>
            </div>
          )}
          {/* Dropdown jump */}
          <select
            value={selected}
            onChange={(e) => onSelect(e.target.value)}
            className="rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-[11px] font-bold text-slate-700 focus:outline-none focus:ring-2 focus:ring-amber-300 cursor-pointer shadow-sm"
          >
            <option value="all">All Months — {breakdown.reduce((s, b) => s + b.total, 0)} rows</option>
            {breakdown.map((b) => (
              <option key={b.month} value={b.month}>
                {b.label} · {b.total} rows · {Math.round(b.matchRate * 100)}% matched
              </option>
            ))}
          </select>
          {selected !== "all" && (
            <button
              onClick={() => onSelect("all")}
              className="text-[10px] font-bold text-slate-400 hover:text-slate-600 px-2 py-1.5 rounded-lg hover:bg-slate-50 transition-colors"
            >✕ Clear</button>
          )}
        </div>
      </div>

      {/* ── Calendar-card scroll strip ────────────────────────────── */}
      <div className="overflow-x-auto">
        <div className="flex gap-2 px-4 py-3 min-w-max">
          {/* "All" card */}
          <button
            onClick={() => onSelect("all")}
            className={`flex flex-col items-center gap-1 rounded-xl px-4 py-2.5 border-2 transition-all min-w-[68px] ${
              selected === "all"
                ? "border-slate-700 shadow-md"
                : "border-slate-100 hover:border-slate-200 hover:shadow-sm"
            }`}
            style={selected === "all" ? { background: NAVY } : { background: "#f8fafc" }}
          >
            <span className={`text-[10px] font-black uppercase ${selected === "all" ? "text-white" : "text-slate-500"}`}>
              All
            </span>
            <span className={`text-xs font-black ${selected === "all" ? "text-white" : "text-slate-700"}`}>
              {breakdown.reduce((s, b) => s + b.total, 0)}
            </span>
            <span className={`text-[9px] ${selected === "all" ? "text-white/70" : "text-slate-400"}`}>rows</span>
          </button>

          {breakdown.map((b) => {
            const rate = b.matchRate;
            const rateRounded = Math.round(rate * 100);
            const isSel = selected === b.month;
            const color = rateColor(rate);
            const bg = rateBg(rate);
            return (
              <button
                key={b.month}
                onClick={() => onSelect(b.month)}
                className={`relative flex flex-col rounded-xl border-2 transition-all overflow-hidden min-w-[82px] ${
                  isSel ? "border-current shadow-lg scale-105" : "border-slate-100 hover:border-slate-200 hover:shadow-sm hover:scale-102"
                }`}
                style={isSel
                  ? { background: NAVY, borderColor: NAVY, color: "white" }
                  : { background: "white" }
                }
              >
                {/* Match-rate fill bar at bottom */}
                {!isSel && (
                  <div
                    className="absolute bottom-0 left-0 right-0 h-1 rounded-b-xl opacity-70"
                    style={{ background: color, width: `${rateRounded}%` }}
                  />
                )}
                <div className="px-3 py-2.5 flex flex-col items-center gap-0.5">
                  <span className={`text-[9px] font-black uppercase tracking-wider ${isSel ? "text-white/70" : "text-slate-400"}`}>
                    {b.label.split(" ")[0]}
                  </span>
                  <span className={`text-[11px] font-black ${isSel ? "text-white/60" : "text-slate-400"}`}>
                    {b.label.split(" ")[1]}
                  </span>
                  <span className={`text-base font-black mt-0.5 ${isSel ? "text-white" : "text-slate-800"}`}>
                    {rateRounded}%
                  </span>
                  <span
                    className={`text-[9px] font-bold px-1.5 py-0.5 rounded-full ${isSel ? "bg-white/20 text-white" : rateText(rate)}`}
                    style={!isSel ? { background: bg } : undefined}
                  >
                    {b.matched}/{b.total}
                  </span>
                </div>
              </button>
            );
          })}
        </div>
      </div>

      {/* ── Selected-month both-sides breakdown ──────────────────── */}
      {active && (
        <div className="border-t border-slate-100 px-4 py-3" style={{ background: "linear-gradient(180deg, #f8fafc, white)" }}>
          <div className="text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2 flex items-center gap-1.5">
            <span className="size-1.5 rounded-full bg-amber-400" />
            {active.label} — Both Sides
          </div>
          <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2">
            {/* Our side */}
            <div className="col-span-2 sm:col-span-1 lg:col-span-2 rounded-xl p-3 flex flex-col gap-1"
              style={{ background: `rgba(12,46,95,0.06)`, border: `1px solid rgba(12,46,95,0.12)` }}>
              <div className="flex items-center gap-1.5 text-[9px] font-black uppercase tracking-wider" style={{ color: NAVY }}>
                <ArrowLeftRight className="size-3" /> Our Ledger (Monthly File)
              </div>
              <div className="text-2xl font-black" style={{ color: NAVY }}>{money(active.oursTotal)}</div>
              <div className="text-[10px] text-slate-500">{active.matched + active.onlyOurs} entries</div>
            </div>

            {/* Supplier side */}
            <div className="col-span-2 sm:col-span-1 lg:col-span-2 rounded-xl p-3 flex flex-col gap-1"
              style={{ background: `rgba(201,162,58,0.08)`, border: `1px solid rgba(201,162,58,0.25)` }}>
              <div className="flex items-center gap-1.5 text-[9px] font-black uppercase tracking-wider" style={{ color: "#b08020" }}>
                <Landmark className="size-3" /> Supplier Ledger (Annual)
              </div>
              <div className="text-2xl font-black" style={{ color: "#b08020" }}>{money(active.partnerTotal)}</div>
              <div className="text-[10px] text-slate-500">{active.matched + active.onlyPartner} entries</div>
            </div>

            {/* Stats */}
            <div className="lg:col-span-2 grid grid-cols-2 gap-2">
              <div className="rounded-xl p-2.5 bg-emerald-50 border border-emerald-100 flex flex-col items-center justify-center">
                <div className="text-[8px] font-black uppercase text-emerald-600">Matched</div>
                <div className="text-xl font-black text-emerald-700">{active.matched}</div>
              </div>
              <div className="rounded-xl p-2.5 flex flex-col items-center justify-center"
                style={{ background: `${rateColor(active.matchRate)}18`, border: `1px solid ${rateColor(active.matchRate)}30` }}>
                <div className="text-[8px] font-black uppercase" style={{ color: rateColor(active.matchRate) }}>Match %</div>
                <div className="text-xl font-black" style={{ color: rateColor(active.matchRate) }}>
                  {Math.round(active.matchRate * 100)}%
                </div>
              </div>
              <div className="rounded-xl p-2.5 bg-indigo-50 border border-indigo-100 flex flex-col items-center justify-center">
                <div className="text-[8px] font-black uppercase text-indigo-600">Only Ours</div>
                <div className="text-xl font-black text-indigo-700">{active.onlyOurs}</div>
              </div>
              <div className="rounded-xl p-2.5 bg-rose-50 border border-rose-100 flex flex-col items-center justify-center">
                <div className="text-[8px] font-black uppercase text-rose-600">Only Supplier</div>
                <div className="text-xl font-black text-rose-700">{active.onlyPartner}</div>
              </div>
            </div>
          </div>

          {/* Variance banner */}
          {(() => {
            const diff = active.oursTotal - active.partnerTotal;
            const absDiff = Math.abs(diff);
            if (absDiff < 0.5) return null;
            return (
              <div className={`mt-2 rounded-xl px-3 py-2 flex items-center justify-between text-xs font-bold ${
                absDiff > 5000 ? "bg-rose-50 text-rose-700 border border-rose-200" : "bg-amber-50 text-amber-700 border border-amber-200"
              }`}>
                <span>⚠ Amount variance this month</span>
                <span className="text-sm font-black">{diff > 0 ? "+" : ""}{money(diff)}</span>
              </div>
            );
          })()}
        </div>
      )}
    </div>
  );
}

function MonthStat({ label, value, accent }: { label: string; value: string | number; accent: string }) {
  return (
    <div className="rounded-lg border border-slate-100 px-3 py-2" style={{ background: `${accent}0d` }}>
      <div className="text-[9px] font-bold uppercase tracking-wider text-slate-400">{label}</div>
      <div className="text-lg font-black" style={{ color: accent }}>{value}</div>
    </div>
  );
}

/** Header chip for multi-file year mode */
function HeaderMultiChip({ label, files, onChange }: { label: string; files: File[]; onChange: (f: File[]) => void }) {
  const handleAdd = (e: React.ChangeEvent<HTMLInputElement>) => {
    const newFiles = Array.from(e.target.files ?? []);
    if (!newFiles.length) return;
    const existing = new Set(files.map((f) => f.name));
    onChange([...files, ...newFiles.filter((f) => !existing.has(f.name))]);
    e.target.value = "";
  };
  return (
    <label
      className={`flex items-center gap-2.5 rounded-xl border px-3.5 py-2 text-xs font-bold cursor-pointer transition-all ${
        files.length > 0 ? "border-amber-400/40 bg-white/10" : "border-white/20 bg-white/5 hover:bg-white/15"
      }`}
    >
      <Calendar className="size-4" style={{ color: files.length > 0 ? GOLD : "rgba(255,255,255,0.6)" }} />
      <div className="flex flex-col">
        <span className="uppercase tracking-tighter text-[8px] text-white/50">{label}</span>
        <span className="text-white/90 truncate max-w-[110px]">
          {files.length > 0 ? `${files.length} file${files.length > 1 ? "s" : ""}` : "Select files"}
        </span>
      </div>
      <input
        type="file"
        accept=".xls,.xlsx,.csv,.tsv,.txt"
        className="hidden"
        multiple
        onChange={handleAdd}
      />
    </label>
  );
}

function HeaderChip({
  label,
  file,
  onChange,
}: {
  label: string;
  file: File | null;
  onChange: (f: File | null) => void;
}) {
  return (
    <label
      className={`flex items-center gap-2.5 rounded-xl border px-3.5 py-2 text-xs font-bold cursor-pointer transition-all ${
        file ? "border-amber-400/40 bg-white/10" : "border-white/20 bg-white/5 hover:bg-white/15"
      }`}
    >
      <FileSpreadsheet
        className="size-4"
        style={{ color: file ? GOLD : "rgba(255,255,255,0.6)" }}
      />
      <div className="flex flex-col">
        <span className="uppercase tracking-tighter text-[8px] text-white/50">{label}</span>
        <span className="text-white/90 truncate max-w-[110px]">
          {file ? file.name : "Select file"}
        </span>
      </div>
      <input
        type="file"
        accept=".xls,.xlsx,.csv,.tsv,.txt"
        className="hidden"
        onChange={(e) => onChange(e.target.files?.[0] ?? null)}
      />
    </label>
  );
}

/* ---------- source data preview ---------- */
function SourceTable({
  title,
  file,
  aoa,
  accent,
}: {
  title: string;
  file: File | null;
  aoa: Aoa | null;
  accent: string;
}) {
  if (!aoa || !aoa.length) {
    return (
      <div className="rounded-xl border border-dashed border-slate-200 p-8 text-center text-xs font-medium text-slate-400">
        No file uploaded for {title}.
      </div>
    );
  }
  const header = (aoa[0] as unknown[]) ?? [];
  const cap = 500;
  const body = aoa.slice(1, cap + 1);
  const more = Math.max(0, aoa.length - 1 - body.length);
  return (
    <div className="rounded-xl border border-slate-200 overflow-hidden">
      <div className="flex items-center gap-2 px-3 py-2 border-b border-slate-100">
        <span className="size-2 rounded-full" style={{ background: accent }} />
        <span className="text-xs font-bold text-slate-700">{title}</span>
        <span className="text-[10px] text-slate-400 truncate">{file?.name}</span>
        <span className="ml-auto text-[10px] font-semibold text-slate-400">
          {Math.max(0, aoa.length - 1)} rows
        </span>
      </div>
      <div className="max-h-[360px] overflow-auto">
        <table className="min-w-full text-[10px]">
          <thead className="bg-slate-50 sticky top-0">
            <tr>
              <th className="px-2 py-1.5 text-left font-black text-slate-300">#</th>
              {header.map((h, i) => (
                <th
                  key={i}
                  className="px-2 py-1.5 text-left font-black text-slate-500 whitespace-nowrap"
                >
                  {String(h ?? "")}
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {body.map((row, ri) => (
              <tr key={ri} className="hover:bg-slate-50/60">
                <td className="px-2 py-1.5 text-slate-300 tabular-nums">{ri + 1}</td>
                {header.map((_, ci) => (
                  <td
                    key={ci}
                    className="px-2 py-1.5 text-slate-600 whitespace-nowrap max-w-[180px] truncate"
                  >
                    {String((row as unknown[])?.[ci] ?? "")}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {more > 0 && (
        <div className="px-3 py-1.5 text-[10px] text-slate-400 bg-slate-50 border-t border-slate-100">
          + {more} more rows not shown
        </div>
      )}
    </div>
  );
}

/* ================================================================== */
/*  PRIMITIVES                                                         */
/* ================================================================== */

function Ring({
  value,
  size = 116,
  stroke = 11,
  color,
  children,
}: {
  value: number;
  size?: number;
  stroke?: number;
  color: string;
  children?: React.ReactNode;
}) {
  const r = (size - stroke) / 2;
  const c = 2 * Math.PI * r;
  const off = c * (1 - Math.max(0, Math.min(1, value)));
  return (
    <div className="relative" style={{ width: size, height: size }}>
      <svg width={size} height={size} className="-rotate-90">
        <circle
          cx={size / 2}
          cy={size / 2}
          r={r}
          stroke="#eef2ff"
          strokeWidth={stroke}
          fill="none"
        />
        <circle
          cx={size / 2}
          cy={size / 2}
          r={r}
          stroke={color}
          strokeWidth={stroke}
          fill="none"
          strokeDasharray={c}
          strokeDashoffset={off}
          strokeLinecap="round"
          style={{ transition: "stroke-dashoffset 0.8s cubic-bezier(.4,0,.2,1)" }}
        />
      </svg>
      <div className="absolute inset-0 flex flex-col items-center justify-center">{children}</div>
    </div>
  );
}

function RingCard({
  title,
  value,
  caption,
  icon,
  color,
}: {
  title: string;
  value: number;
  caption: string;
  icon: React.ReactNode;
  color: string;
}) {
  return (
    <div className="rounded-2xl border border-slate-200/70 bg-white p-5 shadow-sm flex items-center gap-4">
      <Ring value={value} color={color}>
        <span className="text-xl font-black text-slate-800">{pct(value)}</span>
      </Ring>
      <div>
        <div className="flex items-center gap-1.5 text-[10px] font-bold uppercase tracking-widest text-slate-400">
          {icon} {title}
        </div>
        <div className="mt-1.5 text-xs text-slate-500 font-medium leading-snug max-w-[140px]">
          {caption}
        </div>
      </div>
    </div>
  );
}

function StatCard({
  title,
  value,
  caption,
  icon,
  tone = "indigo",
  onClick,
}: {
  title: string;
  value: string;
  caption: string;
  icon: React.ReactNode;
  tone?: "indigo" | "rose" | "emerald" | "amber" | "violet";
  onClick?: () => void;
}) {
  const tones: Record<string, string> = {
    indigo: "text-indigo-600",
    rose: "text-rose-600",
    emerald: "text-emerald-600",
    amber: "text-amber-600",
    violet: "text-violet-600",
  };
  return (
    <button
      onClick={onClick}
      disabled={!onClick}
      className="text-left rounded-2xl border border-slate-200/70 bg-white p-5 shadow-sm transition-all enabled:hover:-translate-y-0.5 enabled:hover:shadow-md disabled:cursor-default"
    >
      <div className="flex items-center justify-between mb-3">
        <span className="text-[10px] font-bold uppercase tracking-widest text-slate-400">
          {title}
        </span>
        <span className="p-1.5 rounded-lg bg-slate-50">{icon}</span>
      </div>
      <div className={`text-2xl font-black tabular-nums tracking-tight ${tones[tone]}`}>
        {value}
      </div>
      <div className="mt-1 text-[11px] text-slate-400 font-medium">{caption}</div>
    </button>
  );
}

function DiffRow({ label, value, raw = false }: { label: string; value: number; raw?: boolean }) {
  const bad = raw ? value > 0 : Math.abs(value) > 1;
  return (
    <div className="flex justify-between items-baseline">
      <span className="text-sm font-medium text-slate-500">{label}</span>
      <span
        className={`text-lg font-bold tabular-nums ${bad ? "text-rose-600" : "text-slate-800"}`}
      >
        {raw ? value : signed(value)}
      </span>
    </div>
  );
}

function SummaryCard({
  label,
  value,
  onClick,
  active,
  accent,
}: {
  label: string;
  value: number;
  onClick: () => void;
  active: boolean;
  accent: string;
}) {
  return (
    <button
      onClick={onClick}
      className={`text-left rounded-2xl border bg-white p-4 shadow-sm transition-all hover:-translate-y-0.5 hover:shadow-md ${
        active ? "ring-2" : "border-slate-200/70"
      }`}
      style={
        active
          ? ({ "--tw-ring-color": accent, borderColor: accent } as React.CSSProperties)
          : undefined
      }
    >
      <div className="flex items-center gap-2 mb-1.5">
        <span className="size-2 rounded-full" style={{ background: accent }} />
        <span className="text-[10px] font-bold uppercase tracking-widest text-slate-400">
          {label}
        </span>
      </div>
      <div className="text-2xl font-black tabular-nums" style={{ color: accent }}>
        {value.toLocaleString()}
      </div>
    </button>
  );
}

function AiMappingCard({ title, mapping, schema }: { title: string; mapping: any; schema: any }) {
  if (!mapping) return null;
  const entries = Object.entries(mapping).filter(
    ([, col]) => col !== null && col !== undefined && col !== "",
  );
  if (!entries.length) return null;
  return (
    <div className="rounded-2xl border border-slate-200/70 bg-white p-5 shadow-sm">
      <div className="flex items-center justify-between mb-4">
        <h4 className="text-xs font-bold uppercase text-slate-400 tracking-wider flex items-center gap-2">
          <Brain className="size-3.5" style={{ color: NAVY }} /> {title}
        </h4>
        {typeof schema?.confidence === "number" && (
          <div className="flex items-center gap-1.5 px-2 py-0.5 rounded-full bg-emerald-50 border border-emerald-100 text-[10px] font-bold text-emerald-700">
            <span className="size-1.5 rounded-full bg-emerald-500 animate-pulse" />
            {pct(schema.confidence)} Confidence
          </div>
        )}
      </div>
      <div className="flex flex-wrap gap-2">
        {entries.map(([key, col]) => (
          <div
            key={key}
            className="px-2.5 py-1 bg-slate-50 border border-slate-100 rounded-md text-[10px] font-medium"
          >
            <span className="text-slate-400 capitalize">{key}:</span>{" "}
            <span className="font-bold" style={{ color: NAVY }}>
              {String(col)}
            </span>
          </div>
        ))}
      </div>
      {schema?.logic && (
        <div
          className="mt-3 text-[10px] text-slate-500 italic bg-slate-50 p-2 rounded-lg border-l-2 leading-relaxed"
          style={{ borderColor: GOLD }}
        >
          <strong>Engine Note:</strong> {schema.logic}
        </div>
      )}
    </div>
  );
}

function TotalsCard({
  title,
  rows,
  charges,
  credits,
  icon,
}: {
  title: string;
  rows: number;
  charges: number;
  credits: number;
  icon: React.ReactNode;
}) {
  return (
    <div className="rounded-2xl border border-slate-200/70 bg-white p-6 shadow-sm transition-transform hover:-translate-y-0.5">
      <div className="flex items-center justify-between mb-4">
        <div className="text-[11px] font-bold uppercase tracking-widest text-slate-400">
          {title}
        </div>
        {icon}
      </div>
      <div className="space-y-3">
        <div className="flex justify-between items-center text-[10px] font-bold text-slate-400 uppercase">
          <span>Rows</span>
          <span className="text-slate-800">{rows}</span>
        </div>
        <div className="h-px bg-slate-50" />
        <KV k="Total Charges" v={money(charges)} />
        <KV k="Total Credits" v={money(credits)} />
        <div className="pt-1">
          <div className="flex justify-between items-center bg-slate-50 p-2.5 rounded-lg">
            <span className="text-[10px] font-black uppercase text-slate-500">Net Exposure</span>
            <span className="text-sm font-black" style={{ color: NAVY }}>
              {signed(+(credits - charges).toFixed(2))}
            </span>
          </div>
        </div>
      </div>
    </div>
  );
}

function KV({ k, v }: { k: string; v: string }) {
  return (
    <div className="flex justify-between gap-3 text-xs">
      <span className="text-slate-500 font-medium">{k}</span>
      <span className="tabular-nums font-semibold text-slate-700">{v}</span>
    </div>
  );
}

const STATUS_STYLE: Record<
  Pair["status"],
  { dot: string; label: string; row: string; text: string }
> = {
  matched: {
    dot: "bg-emerald-500",
    label: "Matched",
    row: "hover:bg-emerald-50/40",
    text: "text-emerald-700",
  },
  amount_diff: {
    dot: "bg-amber-500",
    label: "Amount Diff",
    row: "bg-amber-50/30 hover:bg-amber-50/60",
    text: "text-amber-700",
  },
  missing_ours: {
    dot: "bg-rose-500",
    label: "Only in Partner",
    row: "bg-rose-50/30 hover:bg-rose-50/60",
    text: "text-rose-700",
  },
  missing_partner: {
    dot: "bg-indigo-500",
    label: "Only in Ours",
    row: "bg-indigo-50/20 hover:bg-indigo-50/40",
    text: "text-indigo-700",
  },
};

function ConfidenceChip({ pair }: { pair: Pair }) {
  if (typeof pair.confidence !== "number") return <span className="text-slate-300">—</span>;
  const c = pair.confidence;
  const col = confColor(c);
  return (
    <span className="inline-flex items-center gap-1.5">
      <span className="relative inline-block w-10 h-1.5 rounded-full bg-slate-100 overflow-hidden">
        <span
          className="absolute inset-y-0 left-0 rounded-full"
          style={{ width: `${Math.round(c * 100)}%`, background: col }}
        />
      </span>
      <span className="text-[10px] font-bold tabular-nums" style={{ color: col }}>
        {Math.round(c * 100)}
      </span>
      {pair.evidence?.method === "ai" && (
        <span className="text-[8px] font-black px-1 py-0.5 rounded bg-violet-100 text-violet-600">
          AI
        </span>
      )}
    </span>
  );
}

/**
 * Renders the COMPLETE original row from an uploaded sheet — every column with
 * its header label and value — so the reviewer sees the full detail, not a summary.
 */
function OriginalRowDetail({
  title,
  accent,
  aoa,
  srcRow,
  row,
}: {
  title: string;
  accent: string;
  aoa: Aoa | null;
  srcRow?: number;
  row: LedgerRow | null;
}) {
  if (!row) {
    return (
      <div className="rounded-xl border-2 border-dashed border-rose-200 bg-rose-50/40 p-5 text-center">
        <div className="text-[10px] font-black uppercase tracking-widest text-rose-400 mb-1">
          {title}
        </div>
        <div className="text-sm font-black text-rose-500">⚠ NO MATCHING ROW</div>
        <div className="text-[10px] text-rose-400 font-semibold mt-1">
          This entry exists only on the other side.
        </div>
      </div>
    );
  }
  const header = aoa && aoa.length ? (aoa[0] as unknown[]) : [];
  const cells =
    aoa && srcRow != null && srcRow > 0 && srcRow < aoa.length
      ? (aoa[srcRow] as unknown[])
      : null;

  return (
    <div className="rounded-xl border border-slate-200 overflow-hidden bg-white">
      <div
        className="flex items-center gap-2 px-3 py-2 border-b"
        style={{ borderColor: `${accent}30`, background: `${accent}0a` }}
      >
        <span className="size-2 rounded-full" style={{ background: accent }} />
        <span className="text-[11px] font-black" style={{ color: accent }}>
          {title}
        </span>
        {srcRow != null && (
          <span className="ml-auto text-[9px] font-bold text-slate-400">
            source row #{srcRow}
          </span>
        )}
      </div>
      {/* Parsed/normalised key fields */}
      <div className="grid grid-cols-2 gap-x-4 gap-y-1.5 px-3 py-2.5 bg-slate-50/60 border-b border-slate-100">
        <KvMini k="Date" v={row.date || "—"} />
        <KvMini k="Passport / ID" v={row.passport || "—"} mono />
        <KvMini k="Passenger" v={row.paxName || "—"} wide />
        {row.visaType && <KvMini k="Visa Type" v={row.visaType} />}
        <KvMini k="Reference" v={row.reference || "—"} mono />
        <KvMini k="Charge (DR)" v={money(row.charge)} />
        <KvMini k="Credit (CR)" v={money(row.credit)} />
        {row.description && row.description !== row.paxName && (
          <KvMini k="Description" v={row.description} wide />
        )}
      </div>
      {/* Every original column from the uploaded file */}
      {cells ? (
        <div className="px-3 py-2.5">
          <div className="text-[8px] font-black uppercase tracking-widest text-slate-300 mb-1.5">
            Full original row — all columns
          </div>
          <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-3 gap-y-1">
            {header.map((h, c) => {
              const val = String((cells as unknown[])?.[c] ?? "").trim();
              if (!val) return null;
              return (
                <div key={c} className="min-w-0">
                  <div className="text-[8px] font-bold uppercase text-slate-400 truncate">
                    {String(h ?? `Col ${c + 1}`)}
                  </div>
                  <div className="text-[10px] font-semibold text-slate-700 break-words">
                    {val}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      ) : (
        <div className="px-3 py-2 text-[10px] text-slate-400 italic">
          Original source row unavailable.
        </div>
      )}
    </div>
  );
}

function KvMini({
  k,
  v,
  mono = false,
  wide = false,
}: {
  k: string;
  v: string;
  mono?: boolean;
  wide?: boolean;
}) {
  return (
    <div className={wide ? "col-span-2" : ""}>
      <div className="text-[8px] font-bold uppercase text-slate-400">{k}</div>
      <div className={`text-[11px] font-bold text-slate-700 break-words ${mono ? "font-mono" : ""}`}>
        {v}
      </div>
    </div>
  );
}

function PairsTable({
  pairs,
  onSelect,
  selected,
  rawOurs,
  rawPartner,
  yearMode = false,
  impliedRate = 0,
}: {
  pairs: Pair[];
  onSelect: (p: Pair) => void;
  selected: Pair | null;
  rawOurs: Aoa | null;
  rawPartner: Aoa | null;
  yearMode?: boolean;
  impliedRate?: number;
}) {
  const [expanded, setExpanded] = useState<string | null>(null);

  // Dominant currency on each side (for the amount-column labels).
  const sideCurrency = (pick: (p: Pair) => string | undefined): string => {
    const m = new Map<string, number>();
    for (const p of pairs) { const c = pick(p); if (c) m.set(c, (m.get(c) ?? 0) + 1); }
    return [...m.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? "";
  };
  const oursCur = sideCurrency((p) => p.ours?.currency);
  const partnerCur = sideCurrency((p) => p.partner?.currency);

  if (pairs.length === 0) {
    return (
      <div className="rounded-2xl border-2 border-dashed border-slate-100 bg-white p-12 text-center text-sm font-medium text-slate-400">
        No records match the current criteria.
      </div>
    );
  }
  return (
    <div className="rounded-2xl border border-slate-200/70 bg-white overflow-hidden shadow-sm">
      <div className="overflow-x-auto">
        <table className="min-w-full text-[11px]">
          <thead className="bg-slate-50/80 text-slate-400 border-b border-slate-100">
            <tr className="uppercase tracking-widest font-black">
              <th className="px-2 py-3 text-left w-6" />
              <th className="px-3 py-3 text-left">Status</th>
              <th className="px-3 py-3 text-left">Conf.</th>
              <th className="px-3 py-3 text-left">Scenario / Type</th>
              <th className="px-3 py-3 text-left">Passenger</th>
              <th className="px-3 py-3 text-left">ID / Passport</th>
              <th className="px-3 py-3 text-left border-l border-slate-100">Our Date</th>
              <th className="px-3 py-3 text-right border-l border-slate-100">
                Our Amt
              </th>
              <th className="px-3 py-3 text-left border-l border-slate-100">Partner Date</th>
              <th className="px-3 py-3 text-right border-l border-slate-100">
                Partner Amt
              </th>
              <th className="px-3 py-3 text-right border-l border-slate-100">Variance</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {pairs.map((p, idx) => {
              const s = STATUS_STYLE[p.status];
              const isSel = selected?.key === p.key;
              const isExpanded = expanded === p.key;
              const oursAmt = p.ours ? p.ours.charge || p.ours.credit : 0;
              const partnerAmt = p.partner ? p.partner.charge || p.partner.credit : 0;
              const scenario = p.ours?.scenario ?? p.partner?.scenario;
              const visaType = p.ours?.visaType ?? p.partner?.visaType;
              const paxName = p.ours?.paxName ?? p.partner?.paxName ?? "—";
              const passport = p.ours?.passport ?? p.partner?.passport;
              const dupOurs = (p.ours?.duplicateCount ?? 0) > 1;
              const dupPartner = (p.partner?.duplicateCount ?? 0) > 1;
              const dup = dupOurs ? p.ours!.duplicateCount : dupPartner ? p.partner!.duplicateCount : undefined;
              const dupIdx = dupOurs ? p.ours!.duplicateIndex : p.partner?.duplicateIndex;
              const dupSide = dupOurs && dupPartner ? "Both" : dupOurs ? "Ours" : dupPartner ? "Partner" : "";
              // Scenario row tint — layered under status color
              const scStyle = scenario ? SCENARIO_STYLE[scenario] : null;
              return (
                <React.Fragment key={`${p.key}-${idx}`}>
                  <tr
                    onClick={() => {
                      onSelect(p);
                      setExpanded(isExpanded ? null : p.key);
                    }}
                    className={`cursor-pointer transition-all ${s.row} ${
                      isSel ? "ring-2 ring-inset ring-amber-400 bg-amber-50/20" : ""
                    } ${p.needsReview ? "border-l-4 border-l-amber-400" : ""}`}
                  >
                    <td className="px-2 py-2.5 text-center">
                      <ChevronDown
                        className={`size-3.5 text-slate-300 transition-transform ${isExpanded ? "rotate-180" : ""}`}
                      />
                    </td>
                    <td className="px-3 py-2.5">
                      <span className="flex items-center gap-1.5">
                        <span className={`size-2 rounded-full shrink-0 ${s.dot}`} />
                        <span className={`font-bold ${s.text}`}>{s.label}</span>
                      </span>
                      {p.needsReview && (
                        <span className="mt-0.5 inline-flex text-[8px] font-black text-amber-600 bg-amber-50 px-1 py-0.5 rounded">
                          ⚠ Review
                        </span>
                      )}
                    </td>
                    <td className="px-3 py-2.5">
                      <ConfidenceChip pair={p} />
                    </td>
                    <td className="px-3 py-2.5">
                      <div className="flex flex-col gap-0.5">
                        {scenario && scStyle && (
                          <span
                            className={`inline-flex items-center gap-1 text-[9px] font-black px-1.5 py-0.5 rounded border w-fit ${scStyle.bg} ${scStyle.border} ${scStyle.text}`}
                          >
                            <span className={`size-1.5 rounded-full ${scStyle.dot}`} />
                            {scStyle.label}
                          </span>
                        )}
                        {visaType && (
                          <span className="text-[9px] font-bold text-slate-500 bg-slate-100 px-1.5 py-0.5 rounded w-fit">
                            {visaType}
                          </span>
                        )}
                        {(() => {
                          const mk = pairMonth(p);
                          return mk && mk !== "unknown" ? (
                            <span className="inline-flex items-center gap-1 text-[8px] font-black text-indigo-600 bg-indigo-50 border border-indigo-200 px-1 py-0.5 rounded w-fit">
                              <Calendar className="size-2" /> {monthLabel(mk)}
                            </span>
                          ) : null;
                        })()}
                        {dup && dup > 1 && (
                          <span
                            className="text-[8px] font-black text-rose-600 bg-rose-50 border border-rose-200 px-1 py-0.5 rounded w-fit"
                            title={`This same entry appears ${dup} times in the ${dupSide} sheet. This is copy ${dupIdx} of ${dup}.`}
                          >
                            ⧉ Same entry {dupIdx} of {dup} · in {dupSide}
                          </span>
                        )}
                      </div>
                    </td>
                    <td className="px-3 py-2.5 max-w-[160px]">
                      <div className="font-semibold text-slate-700 truncate" title={paxName}>
                        {paxName}
                      </div>
                      {p.ours?.description && p.ours.description !== paxName && (
                        <div
                          className="text-[9px] text-slate-400 truncate"
                          title={p.ours.description}
                        >
                          {p.ours.description}
                        </div>
                      )}
                    </td>
                    <td className="px-3 py-2.5 font-mono font-bold text-slate-600 text-[10px]">
                      {passport ?? "—"}
                    </td>
                    <td className="px-3 py-2.5 text-slate-400 tabular-nums border-l border-slate-100">
                      {p.ours?.date ?? <span className="text-slate-200">—</span>}
                    </td>
                    <td className="px-3 py-2.5 text-right font-bold tabular-nums border-l border-slate-100">
                      {p.ours ? money(oursAmt) : <span className="text-slate-200">—</span>}
                    </td>
                    <td className="px-3 py-2.5 text-slate-400 tabular-nums border-l border-slate-100">
                      {p.partner?.date ?? <span className="text-slate-200">—</span>}
                    </td>
                    <td className="px-3 py-2.5 text-right font-bold tabular-nums border-l border-slate-100">
                      {p.partner ? money(partnerAmt) : <span className="text-slate-200">—</span>}
                    </td>
                    <td
                      className="px-3 py-2.5 text-right tabular-nums border-l border-slate-100 font-black"
                      title={
                        p.status === "matched" && impliedRate && rateDeviation(p, impliedRate) !== null
                          ? `Supplier charged ${money(partnerAmt)}; at the usual rate (${impliedRate.toFixed(2)}×) it should be ≈ ${money(oursAmt * impliedRate)}. ${p.note || ""}`
                          : p.note || undefined
                      }
                    >
                      {(() => {
                        if (p.status !== "matched")
                          return (
                            <span className={Math.abs(p.diff) > 0.5 ? "text-rose-600" : "text-slate-300"}>
                              {signed(p.diff)}
                            </span>
                          );
                        const dev = impliedRate ? rateDeviation(p, impliedRate) : null;
                        if (dev !== null) {
                          const off = Math.abs(dev) > RATE_OFF_THRESHOLD;
                          return off ? (
                            <span className="text-rose-600">
                              {dev > 0 ? "+" : ""}{Math.round(dev * 100)}%
                            </span>
                          ) : (
                            <span className="text-emerald-500">✓ on&nbsp;rate</span>
                          );
                        }
                        return yearMode && Math.abs(p.diff) > 0.5 ? (
                          <span className="text-[9px] font-semibold text-emerald-600">✓ Ref</span>
                        ) : (
                          <span className="text-emerald-500">✓</span>
                        );
                      })()}
                    </td>
                  </tr>
                  {isExpanded && (
                    <tr className={s.row}>
                      <td colSpan={11} className="px-4 py-4 bg-slate-50/40">
                        <div className="grid gap-3 lg:grid-cols-2">
                          <OriginalRowDetail
                            title="Our Ledger — Full Row"
                            accent={NAVY}
                            aoa={rawOurs}
                            srcRow={p.ours?.srcRow}
                            row={p.ours}
                          />
                          <OriginalRowDetail
                            title="Partner Ledger — Full Row"
                            accent="#7c3aed"
                            aoa={rawPartner}
                            srcRow={p.partner?.srcRow}
                            row={p.partner}
                          />
                        </div>
                        {/* Reconciliation verdict line */}
                        <div className="mt-3 flex flex-wrap items-center gap-2 text-[11px]">
                          <span className="font-black uppercase tracking-widest text-slate-400">
                            Verdict:
                          </span>
                          <span className={`font-bold ${s.text}`}>{p.note}</span>
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
      <div className="px-4 py-2 text-[10px] text-slate-400 font-semibold bg-slate-50/60 border-t border-slate-100">
        Tip: click any row to expand the full original detail from both ledgers.
      </div>
    </div>
  );
}

/* ================================================================== */
/*  FULL LEDGER VIEW  — both uploaded sheets, every column, marked     */
/* ================================================================== */

const STATUS_BG: Record<Pair["status"], string> = {
  matched: "bg-emerald-50",
  amount_diff: "bg-amber-50",
  missing_partner: "bg-indigo-50",
  missing_ours: "bg-rose-50",
};

function LedgerSheet({
  title,
  aoa,
  map,
  accent,
  query,
  issuesOnly,
  isFiltered,
}: {
  title: string;
  aoa: Aoa | null;
  map: Map<number, Pair>;
  accent: string;
  query: string;
  issuesOnly: boolean;
  isFiltered?: boolean;
}) {
  if (!aoa || aoa.length < 1) {
    return (
      <div className="rounded-xl border border-dashed border-slate-200 p-8 text-center text-xs font-medium text-slate-400">
        No data for {title}.
      </div>
    );
  }
  const header = (aoa[0] as unknown[]) ?? [];
  const q = query.trim().toLowerCase();

  // Build the visible row list (keep original index for status lookup).
  const rows: { i: number; cells: unknown[]; pair: Pair | undefined }[] = [];
  for (let i = 1; i < aoa.length; i++) {
    const cells = (aoa[i] as unknown[]) ?? [];
    const pair = map.get(i);
    if (isFiltered && !pair) continue;
    if (issuesOnly && (!pair || pair.status === "matched")) continue;
    if (q) {
      let hay = cells.map((c) => String(c ?? "")).join(" ").toLowerCase();
      if (pair) {
        const st = pair.status;
        if (st) hay += " " + STATUS_STYLE[st].label.toLowerCase();
      } else {
        hay += " unmarked";
      }
      if (!hay.includes(q)) continue;
    }
    rows.push({ i, cells, pair });
  }

  return (
    <div className="rounded-xl border border-slate-200 overflow-hidden">
      <div className="flex items-center gap-2 px-3 py-2 border-b border-slate-100 bg-slate-50/70">
        <span className="size-2.5 rounded-full" style={{ background: accent }} />
        <span className="text-xs font-black text-slate-700">{title}</span>
        <span className="ml-auto text-[10px] font-semibold text-slate-400">
          {rows.length} of {Math.max(0, aoa.length - 1)} rows
        </span>
      </div>
      <div className="max-h-[460px] overflow-auto">
        <table className="min-w-full text-[10px]">
          <thead className="bg-slate-100 sticky top-0 z-10">
            <tr>
              <th className="px-2 py-1.5 text-left font-black text-slate-400 whitespace-nowrap">
                #
              </th>
              <th className="px-2 py-1.5 text-left font-black text-slate-500 whitespace-nowrap">
                Match Status
              </th>
              {header.map((h, c) => (
                <th
                  key={c}
                  className="px-2 py-1.5 text-left font-black text-slate-500 whitespace-nowrap"
                >
                  {String(h ?? "")}
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {rows.map(({ i, cells, pair }) => {
              const st = pair?.status;
              const style = st ? STATUS_STYLE[st] : null;
              const bg = st ? STATUS_BG[st] : "";
              return (
                <tr key={i} className={`${bg} hover:brightness-95 transition`}>
                  <td className="px-2 py-1.5 text-slate-300 tabular-nums">{i}</td>
                  <td className="px-2 py-1.5 whitespace-nowrap">
                    {style ? (
                      <span className="inline-flex items-center gap-1.5">
                        <span className={`size-1.5 rounded-full ${style.dot}`} />
                        <span className={`font-bold ${style.text}`}>{style.label}</span>
                        {pair?.diff != null && Math.abs(pair.diff) > 0.5 && (
                          <span className="text-rose-600 font-black">{signed(pair.diff)}</span>
                        )}
                      </span>
                    ) : (
                      <span className="text-slate-300 font-semibold">—</span>
                    )}
                  </td>
                  {header.map((_, c) => (
                    <td
                      key={c}
                      className="px-2 py-1.5 text-slate-600 whitespace-nowrap max-w-[200px] truncate"
                    >
                      {String((cells as unknown[])?.[c] ?? "")}
                    </td>
                  ))}
                </tr>
              );
            })}
            {rows.length === 0 && (
              <tr>
                <td
                  colSpan={header.length + 2}
                  className="px-4 py-8 text-center text-slate-400 font-medium"
                >
                  No rows match the current filter.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function FullLedgerView({
  ours,
  partner,
  result,
  pairs,
}: {
  ours: Aoa | null;
  partner: Aoa | null;
  result: ReconResult;
  pairs: Pair[];
}) {
  const [query, setQuery] = useState("");
  const [issuesOnly, setIssuesOnly] = useState(false);

  const isFiltered = pairs !== result.pairs;

  const getAoaAndMap = useCallback((side: "ours" | "partner", rawAoa: Aoa | null) => {
    if (rawAoa) {
      const m = new Map<number, Pair>();
      pairs.forEach((p) => {
        if (p[side]?.srcRow != null) m.set(p[side]!.srcRow!, p);
      });
      return { aoa: rawAoa, map: m };
    }
    // Synthetic fallback for multi-file mode
    const allPairs = isFiltered ? pairs : result.pairs;
    const rows = allPairs.map(p => p[side]).filter(Boolean) as LedgerRow[];
    const unique = Array.from(new Set(rows));
    
    const aoa: Aoa = [["Date", "Reference", "Name", "Charge", "Credit", "Description"]];
    const m = new Map<number, Pair>();
    
    unique.forEach((r, idx) => {
      const srcRow = idx + 1;
      aoa.push([
        r.date || r.month || "",
        r.reference || "",
        r.paxName || "",
        r.charge > 0 ? r.charge : "",
        r.credit > 0 ? r.credit : "",
        r.description || ""
      ]);
      const p = allPairs.find(x => x[side] === r);
      if (p) m.set(srcRow, p);
    });
    return { aoa, map: m };
  }, [pairs, result.pairs, isFiltered]);

  const { aoa: oursAoa, map: oursMap } = useMemo(() => getAoaAndMap("ours", ours), [getAoaAndMap, ours]);
  const { aoa: partnerAoa, map: partnerMap } = useMemo(() => getAoaAndMap("partner", partner), [getAoaAndMap, partner]);

  const legend: { label: string; dot: string; text: string }[] = [
    { label: "Matched", dot: "bg-emerald-500", text: "text-emerald-700" },
    { label: "Amount Diff", dot: "bg-amber-500", text: "text-amber-700" },
    { label: "Only in Ours", dot: "bg-indigo-500", text: "text-indigo-700" },
    { label: "Only in Partner", dot: "bg-rose-500", text: "text-rose-700" },
    { label: "Unmarked (B/F, blank…)", dot: "bg-slate-300", text: "text-slate-500" },
  ];

  return (
    <div className="space-y-4">
      {/* Toolbar */}
      <div className="bg-white rounded-2xl border border-slate-200/70 p-3 shadow-sm flex flex-wrap items-center gap-3">
        <div className="flex items-center gap-1.5 text-[11px] font-black uppercase tracking-wider text-slate-500">
          <Table2 className="size-4" style={{ color: NAVY }} />
          Both Sheets — Full Ledger
        </div>
        <div className="flex flex-wrap items-center gap-x-3 gap-y-1">
          {legend.map((l) => (
            <span key={l.label} className="flex items-center gap-1.5 text-[10px] font-bold">
              <span className={`size-2 rounded-full ${l.dot}`} />
              <span className={l.text}>{l.label}</span>
            </span>
          ))}
        </div>
        <div className="ml-auto flex items-center gap-2">
          <button
            onClick={() => setIssuesOnly((s) => !s)}
            className={`text-[10px] px-2.5 py-1.5 rounded-lg font-bold transition-all ${
              issuesOnly ? "bg-rose-500 text-white shadow" : "text-slate-400 hover:bg-slate-50"
            }`}
            title="Show only rows that need attention"
          >
            Issues only
          </button>
          <div className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 rounded-lg px-2 py-1">
            <Search className="size-3.5 text-slate-300" />
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search any cell…"
              className="min-w-[150px] bg-transparent text-[11px] focus:outline-none"
            />
          </div>
        </div>
      </div>

      <LedgerSheet
        title="Our Ledger (uploaded file, all columns)"
        aoa={oursAoa}
        map={oursMap}
        accent={NAVY}
        query={query}
        issuesOnly={issuesOnly}
        isFiltered={isFiltered}
      />
      <LedgerSheet
        title="Partner Ledger (uploaded file, all columns)"
        aoa={partnerAoa}
        map={partnerMap}
        accent={GOLD}
        query={query}
        issuesOnly={issuesOnly}
        isFiltered={isFiltered}
      />
    </div>
  );
}

/* ================================================================== */
/*  PAYMENT FINDER VIEW                                                */
/* ================================================================== */

type PaySubFilter = "all" | "matched" | "amount_diff" | "only_ours" | "only_partner";

/* ---- small summary tile for the payment panel ---- */
function PayTile({
  label,
  value,
  sub,
  color,
  netSign,
}: {
  label: string;
  value: number;
  sub: string;
  color: string;
  netSign?: number; // when set, prefixes + / - and colours by sign
}) {
  const display =
    netSign !== undefined ? (
      <span style={{ color: Math.abs(netSign) < 1 ? "#10b981" : "#ef4444" }}>
        {netSign > 0 ? "+" : netSign < 0 ? "−" : ""}
        {money(Math.abs(value))}
      </span>
    ) : (
      <span>{money(value)}</span>
    );
  return (
    <div className="bg-white rounded-2xl border border-slate-200/70 px-4 py-3 shadow-sm">
      <div className="text-[9px] font-black uppercase tracking-widest mb-1" style={{ color }}>
        {label}
      </div>
      <div className="text-[13px] font-black text-slate-800 tabular-nums">{display}</div>
      <div className="text-[9px] text-slate-400 font-semibold mt-0.5">{sub}</div>
    </div>
  );
}

/* ---- date-gap badge ---- */
function DayGapBadge({ days }: { days: number | null | undefined }) {
  if (days === null || days === undefined)
    return <span className="text-slate-200 text-[10px]">—</span>;
  if (days === 0)
    return (
      <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-600">
        Same day
      </span>
    );
  if (days <= 2)
    return (
      <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-amber-50 text-amber-600">
        +{days}d
      </span>
    );
  return (
    <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-rose-50 text-rose-600">
      +{days}d
    </span>
  );
}

function PaymentFinderView({
  pairs,
  onSelect,
  selected,
}: {
  pairs: Pair[];
  onSelect: (p: Pair) => void;
  selected: Pair | null;
}) {
  const [subFilter, setSubFilter] = useState<PaySubFilter>("all");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [bankRefQuery, setBankRefQuery] = useState("");

  const matchedPairs = useMemo(() => pairs.filter((p) => p.status === "matched"), [pairs]);
  const amtDiffPairs = useMemo(() => pairs.filter((p) => p.status === "amount_diff"), [pairs]);
  const onlyOursPairs = useMemo(() => pairs.filter((p) => p.status === "missing_partner"), [pairs]);
  const onlyPartnerPairs = useMemo(
    () => pairs.filter((p) => p.status === "missing_ours"),
    [pairs],
  );

  const oursSentTotal = useMemo(
    () => pairs.reduce((s, p) => s + (p.ours?.credit || p.ours?.charge || 0), 0),
    [pairs],
  );
  const partnerReceivedTotal = useMemo(
    () => pairs.reduce((s, p) => s + (p.partner?.credit || p.partner?.charge || 0), 0),
    [pairs],
  );
  const netDiff = oursSentTotal - partnerReceivedTotal;
  const matchedAmt = matchedPairs.reduce((s, p) => s + p.oursAmt, 0);
  const unmatchedOursAmt = onlyOursPairs.reduce((s, p) => s + p.oursAmt, 0);
  const unmatchedPartnerAmt = onlyPartnerPairs.reduce((s, p) => s + p.partnerAmt, 0);

  const filtered = useMemo(() => {
    let list = pairs;
    if (subFilter === "matched") list = list.filter((p) => p.status === "matched");
    else if (subFilter === "amount_diff") list = list.filter((p) => p.status === "amount_diff");
    else if (subFilter === "only_ours") list = list.filter((p) => p.status === "missing_partner");
    else if (subFilter === "only_partner") list = list.filter((p) => p.status === "missing_ours");

    if (bankRefQuery.trim()) {
      const q = bankRefQuery.trim().toLowerCase();
      list = list.filter((p) =>
        [p.partner?.reference, p.ours?.reference, p.partner?.description, p.ours?.description]
          .filter(Boolean)
          .join(" ")
          .toLowerCase()
          .includes(q),
      );
    }
    if (dateFrom)
      list = list.filter((p) => {
        const d = p.ours?.date || p.partner?.date || "";
        return d >= dateFrom;
      });
    if (dateTo)
      list = list.filter((p) => {
        const d = p.ours?.date || p.partner?.date || "";
        return d <= dateTo;
      });
    return list;
  }, [pairs, subFilter, bankRefQuery, dateFrom, dateTo]);

  const subTabs: [PaySubFilter, string, number][] = [
    ["all", "All", pairs.length],
    ["matched", "Matched", matchedPairs.length],
    ["amount_diff", "Amt Diff", amtDiffPairs.length],
    ["only_ours", "Sent — Unconfirmed", onlyOursPairs.length],
    ["only_partner", "Partner Only", onlyPartnerPairs.length],
  ];

  return (
    <div className="space-y-4">
      {/* ── Header banner ─────────────────────────────────────────── */}
      <div
        className="rounded-2xl p-5 text-white"
        style={{ background: `linear-gradient(120deg, #0a2547 0%, ${NAVY} 60%, #103a73 100%)` }}
      >
        <div className="flex items-center gap-3 mb-4">
          <div
            className="size-9 rounded-xl flex items-center justify-center"
            style={{ background: "rgba(201,162,58,0.2)" }}
          >
            <ArrowLeftRight className="size-5" style={{ color: GOLD }} />
          </div>
          <div>
            <div className="text-[11px] font-black uppercase tracking-widest opacity-70">
              Bank Transfer Reconciliation
            </div>
            <div className="text-base font-black">
              {pairs.length} transfer{pairs.length !== 1 ? "s" : ""} —{" "}
              <span style={{ color: GOLD }}>{matchedPairs.length} matched</span>
            </div>
          </div>
          <div className="ml-auto text-right">
            <div className="text-[9px] font-bold uppercase tracking-widest opacity-50">
              Net Gap
            </div>
            <div
              className="text-lg font-black tabular-nums"
              style={{ color: Math.abs(netDiff) < 1 ? "#34d399" : "#f87171" }}
            >
              {netDiff > 0 ? "+" : netDiff < 0 ? "−" : ""}
              {money(Math.abs(netDiff))}
            </div>
          </div>
        </div>
        {/* Summary stat row */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          <div className="bg-white/10 rounded-xl px-3 py-2.5">
            <div className="text-[9px] font-black uppercase opacity-60 mb-0.5">Our Ledger — Sent</div>
            <div className="text-sm font-black tabular-nums">{money(oursSentTotal)}</div>
            <div className="text-[9px] opacity-50">
              {matchedPairs.length + onlyOursPairs.length + amtDiffPairs.length} transfers
            </div>
          </div>
          <div className="bg-white/10 rounded-xl px-3 py-2.5">
            <div className="text-[9px] font-black uppercase opacity-60 mb-0.5">
              Partner Recorded
            </div>
            <div className="text-sm font-black tabular-nums">{money(partnerReceivedTotal)}</div>
            <div className="text-[9px] opacity-50">
              {matchedPairs.length + onlyPartnerPairs.length + amtDiffPairs.length} transfers
            </div>
          </div>
          <div className="bg-white/10 rounded-xl px-3 py-2.5">
            <div className="text-[9px] font-black uppercase opacity-60 mb-0.5">
              Matched Value
            </div>
            <div className="text-sm font-black tabular-nums text-emerald-300">
              {money(matchedAmt)}
            </div>
            <div className="text-[9px] opacity-50">{matchedPairs.length} confirmed</div>
          </div>
          <div className="bg-white/10 rounded-xl px-3 py-2.5">
            <div className="text-[9px] font-black uppercase opacity-60 mb-0.5">Unconfirmed</div>
            <div
              className="text-sm font-black tabular-nums"
              style={{ color: onlyOursPairs.length > 0 ? "#fbbf24" : "#34d399" }}
            >
              {money(unmatchedOursAmt || unmatchedPartnerAmt)}
            </div>
            <div className="text-[9px] opacity-50">
              {onlyOursPairs.length} ours · {onlyPartnerPairs.length} partner
            </div>
          </div>
        </div>
      </div>

      {/* ── Filters ───────────────────────────────────────────────── */}
      <div className="bg-white rounded-2xl border border-slate-200/70 p-3 shadow-sm">
        <div className="flex flex-wrap items-center gap-1.5">
          <Filter className="size-3.5 text-slate-300 shrink-0" />
          {subTabs.map(([k, lbl, cnt]) => (
            <button
              key={k}
              onClick={() => setSubFilter(k)}
              className={`text-[10px] px-2.5 py-1.5 rounded-lg font-bold transition-all ${
                subFilter === k ? "text-white shadow" : "text-slate-400 hover:bg-slate-50"
              }`}
              style={subFilter === k ? { background: NAVY } : undefined}
            >
              {lbl}
              {cnt > 0 && (
                <span
                  className="ml-1 opacity-60"
                  style={
                    subFilter !== k && (k === "only_ours" || k === "only_partner") && cnt > 0
                      ? { color: "#ef4444", opacity: 1 }
                      : undefined
                  }
                >
                  ({cnt})
                </span>
              )}
            </button>
          ))}

          <div className="ml-auto flex items-center gap-2 flex-wrap">
            {/* Bank ref search */}
            <div className="flex items-center gap-1 bg-slate-50 border border-slate-200 rounded-lg px-2 py-1">
              <Search className="size-3 text-slate-300 shrink-0" />
              <input
                value={bankRefQuery}
                onChange={(e) => setBankRefQuery(e.target.value)}
                placeholder="Bank ref…"
                className="w-28 bg-transparent text-[10px] focus:outline-none"
              />
            </div>
            {/* Date range */}
            <div className="flex items-center gap-1 text-[10px] text-slate-400">
              <Calendar className="size-3 shrink-0" />
              <input
                type="date"
                value={dateFrom}
                onChange={(e) => setDateFrom(e.target.value)}
                className="bg-slate-50 border border-slate-200 rounded-lg px-2 py-1 text-[10px] focus:outline-none"
              />
              <span>–</span>
              <input
                type="date"
                value={dateTo}
                onChange={(e) => setDateTo(e.target.value)}
                className="bg-slate-50 border border-slate-200 rounded-lg px-2 py-1 text-[10px] focus:outline-none"
              />
            </div>
            {(bankRefQuery || dateFrom || dateTo) && (
              <button
                onClick={() => {
                  setBankRefQuery("");
                  setDateFrom("");
                  setDateTo("");
                }}
                className="text-[9px] font-black text-slate-300 hover:text-rose-400 transition-colors"
              >
                Clear ×
              </button>
            )}
            <span className="text-[10px] font-bold text-slate-400">
              {filtered.length}/{pairs.length}
            </span>
          </div>
        </div>
      </div>

      {/* ── Main transfer table ────────────────────────────────────── */}
      <div className="rounded-2xl border border-slate-200/70 bg-white overflow-hidden shadow-sm">
        <div className="overflow-x-auto">
          <table className="min-w-full text-[11px]">
            <thead>
              {/* Group header */}
              <tr className="text-[9px] font-black uppercase tracking-widest text-slate-400 bg-slate-50/80 border-b border-slate-100">
                <th className="px-3 py-3 text-left" rowSpan={2}>
                  Status
                </th>
                <th
                  className="px-3 py-2 text-left"
                  colSpan={2}
                  style={{ borderLeft: `3px solid ${NAVY}` }}
                >
                  <span style={{ color: NAVY }}>●</span> Our Ledger — Sent
                </th>
                <th
                  className="px-3 py-2 text-center w-10"
                  style={{ borderLeft: "1px solid #f1f5f9", borderRight: "1px solid #f1f5f9" }}
                />
                <th className="px-3 py-2 text-left" colSpan={2}>
                  <span style={{ color: "#7c3aed" }}>●</span> Partner Ledger — Recorded
                </th>
                <th className="px-3 py-2 text-left">Bank Reference</th>
                <th className="px-3 py-2 text-center">Date Gap</th>
                <th
                  className="px-3 py-2 text-right"
                  style={{ borderLeft: "2px solid #e2e8f0" }}
                >
                  Variance
                </th>
                <th className="px-3 py-2 text-center">Conf.</th>
              </tr>
              <tr className="text-[9px] font-bold text-slate-300 bg-slate-50/50 border-b border-slate-100">
                <th
                  className="px-3 pb-2 text-left"
                  style={{ borderLeft: `3px solid ${NAVY}` }}
                >
                  Date
                </th>
                <th className="px-3 pb-2 text-right">Amount</th>
                <th
                  style={{ borderLeft: "1px solid #f1f5f9", borderRight: "1px solid #f1f5f9" }}
                />
                <th className="px-3 pb-2 text-left">Date</th>
                <th className="px-3 pb-2 text-right">Amount</th>
                <th className="px-3 pb-2 text-left" />
                <th />
                <th style={{ borderLeft: "2px solid #e2e8f0" }} />
                <th />
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50">
              {filtered.length === 0 ? (
                <tr>
                  <td
                    colSpan={10}
                    className="px-6 py-12 text-center text-sm font-medium text-slate-400"
                  >
                    No bank transfers match the current filter.
                  </td>
                </tr>
              ) : (
                filtered.map((p, idx) => {
                  const s = STATUS_STYLE[p.status];
                  const isSel = selected?.key === p.key;
                  const oAmt = p.ours?.credit || p.ours?.charge || 0;
                  const pAmt = p.partner?.credit || p.partner?.charge || 0;
                  const bankRef = p.partner?.reference || p.ours?.reference || "";
                  const gap = p.evidence?.dateDeltaDays;
                  const matchIcon =
                    p.status === "matched"
                      ? { icon: "↔", col: "#10b981" }
                      : p.status === "amount_diff"
                        ? { icon: "≈", col: "#f59e0b" }
                        : p.status === "missing_partner"
                          ? { icon: "⟶", col: "#6366f1" }
                          : { icon: "⟵", col: "#ef4444" };
                  return (
                    <tr
                      key={`pt-${p.key}-${idx}`}
                      onClick={() => onSelect(p)}
                      className={`cursor-pointer transition-all ${s.row} ${
                        isSel ? "ring-2 ring-inset ring-amber-400 bg-amber-50/20" : ""
                      } ${p.needsReview ? "border-l-2 border-l-amber-400" : ""}`}
                    >
                      {/* Status */}
                      <td className="px-3 py-3 whitespace-nowrap">
                        <div className="flex items-center gap-1.5 mb-1">
                          <span className={`size-2 rounded-full ${s.dot}`} />
                          <span className={`text-[10px] font-bold ${s.text}`}>{s.label}</span>
                        </div>
                      </td>
                      {/* NST side */}
                      <td
                        className="px-3 py-3 text-slate-500 tabular-nums whitespace-nowrap"
                        style={{ borderLeft: `3px solid ${NAVY}` }}
                      >
                        {p.ours?.date ?? <span className="text-slate-200">—</span>}
                      </td>
                      <td className="px-3 py-3 text-right font-black tabular-nums whitespace-nowrap"
                        style={{ color: p.ours ? NAVY : "#e2e8f0" }}>
                        {p.ours ? money(oAmt) : "—"}
                      </td>
                      {/* Arrow */}
                      <td
                        className="px-2 py-3 text-center text-base font-black whitespace-nowrap"
                        style={{
                          borderLeft: "1px solid #f1f5f9",
                          borderRight: "1px solid #f1f5f9",
                          color: matchIcon.col,
                        }}
                      >
                        {matchIcon.icon}
                      </td>
                      {/* Partner side */}
                      <td className="px-3 py-3 text-slate-500 tabular-nums whitespace-nowrap">
                        {p.partner?.date ?? <span className="text-slate-200">—</span>}
                      </td>
                      <td className="px-3 py-3 text-right font-black tabular-nums whitespace-nowrap"
                        style={{ color: p.partner ? "#7c3aed" : "#e2e8f0" }}>
                        {p.partner ? money(pAmt) : "—"}
                      </td>
                      {/* Bank Reference (from SmartTrip comments) */}
                      <td className="px-3 py-3 font-mono text-[10px] text-slate-400 whitespace-nowrap">
                        {bankRef ? (
                          <span
                            className="px-1.5 py-0.5 rounded bg-slate-50 border border-slate-100"
                            title={bankRef}
                          >
                            {bankRef}
                          </span>
                        ) : (
                          <span className="text-slate-200">—</span>
                        )}
                      </td>
                      {/* Date gap */}
                      <td className="px-3 py-3 text-center">
                        <DayGapBadge days={gap} />
                      </td>
                      {/* Variance */}
                      <td
                        className={`px-3 py-3 text-right tabular-nums font-black whitespace-nowrap ${
                          Math.abs(p.diff) > 0.5 ? "text-rose-600" : "text-slate-300"
                        }`}
                        style={{ borderLeft: "2px solid #e2e8f0" }}
                      >
                        {p.status === "matched" ? (
                          <span className="text-emerald-500 text-base">✓</span>
                        ) : (
                          signed(p.diff)
                        )}
                      </td>
                      {/* Confidence */}
                      <td className="px-3 py-3">
                        <ConfidenceChip pair={p} />
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>

        {/* Footer totals */}
        {filtered.length > 0 && (
          <div
            className="px-5 py-3 border-t border-slate-100 flex flex-wrap justify-between items-center gap-3 text-[10px] font-bold text-slate-400"
            style={{ background: "#fafafa" }}
          >
            <span>
              Showing{" "}
              <strong className="text-slate-600">{filtered.length}</strong> of {pairs.length}{" "}
              bank transfers
            </span>
            <span className="flex gap-6">
              <span>
                Our total:{" "}
                <span className="font-black tabular-nums" style={{ color: NAVY }}>
                  {money(filtered.reduce((s, p) => s + (p.ours?.credit || p.ours?.charge || 0), 0))}
                </span>
              </span>
              <span>
                Partner total:{" "}
                <span className="font-black tabular-nums" style={{ color: "#7c3aed" }}>
                  {money(
                    filtered.reduce(
                      (s, p) => s + (p.partner?.credit || p.partner?.charge || 0),
                      0,
                    ),
                  )}
                </span>
              </span>
              <span>
                Variance:{" "}
                <span
                  className="font-black tabular-nums"
                  style={{
                    color:
                      Math.abs(
                        filtered.reduce((s, p) => s + p.diff, 0),
                      ) < 1
                        ? "#10b981"
                        : "#ef4444",
                  }}
                >
                  {signed(+(filtered.reduce((s, p) => s + p.diff, 0)).toFixed(2))}
                </span>
              </span>
            </span>
          </div>
        )}
      </div>
    </div>
  );
}

function SignalBar({ label, value }: { label: string; value: number }) {
  const col =
    value >= 0.85 ? "#10b981" : value >= 0.5 ? "#f59e0b" : value > 0 ? "#ef4444" : "#cbd5e1";
  return (
    <div>
      <div className="flex justify-between text-[10px] font-bold text-slate-400 uppercase mb-1">
        <span>{label}</span>
        <span style={{ color: value > 0 ? col : undefined }}>{Math.round(value * 100)}%</span>
      </div>
      <div className="h-1.5 w-full rounded-full bg-slate-100 overflow-hidden">
        <div
          className="h-full rounded-full transition-all duration-500"
          style={{ width: `${Math.round(value * 100)}%`, background: col }}
        />
      </div>
    </div>
  );
}

function DetailPanel({ pair }: { pair: Pair | null }) {
  if (!pair) {
    return (
      <aside className="rounded-2xl border border-slate-200/70 bg-white p-8 text-center h-fit lg:sticky lg:top-24 shadow-sm">
        <Info className="size-8 text-slate-200 mx-auto mb-4" />
        <div className="text-sm font-bold text-slate-400 uppercase tracking-widest mb-1">
          Row Inspector
        </div>
        <p className="text-xs text-slate-400 leading-relaxed">
          Select any record to view its confidence breakdown, matching evidence, and raw dual-ledger
          comparison.
        </p>
      </aside>
    );
  }
  const s = STATUS_STYLE[pair.status];
  const e = pair.evidence;
  const conf = pair.confidence ?? 0;
  const scenario = pair.ours?.scenario ?? pair.partner?.scenario;
  const scStyle = scenario ? SCENARIO_STYLE[scenario] : null;
  const oursVT = pair.ours?.visaType;
  const partnerVT = pair.partner?.visaType;
  const vtMismatch = oursVT && partnerVT && oursVT !== partnerVT;
  return (
    <aside className="rounded-2xl border border-slate-200/70 bg-white p-6 h-fit lg:sticky lg:top-24 space-y-4 shadow-xl">
      {/* Status + Scenario row */}
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div className="flex items-center gap-2 flex-wrap">
          <span className={`size-3 rounded-full ${s.dot}`} />
          <span className="text-sm font-black text-slate-800">{s.label}</span>
          {scStyle && (
            <span
              className={`text-[9px] font-black px-2 py-0.5 rounded border ${scStyle.bg} ${scStyle.border} ${scStyle.text}`}
            >
              {scStyle.label}
            </span>
          )}
        </div>
        {pair.needsReview && (
          <span className="text-[9px] font-black bg-amber-100 text-amber-700 px-2 py-1 rounded-md uppercase">
            ⚠ Review
          </span>
        )}
      </div>

      {/* Visa type info */}
      {(oursVT || partnerVT) && (
        <div className="flex flex-wrap gap-2 text-[10px]">
          {oursVT && (
            <span className="font-bold bg-slate-100 text-slate-600 px-2 py-0.5 rounded">
              Our type: {oursVT}
            </span>
          )}
          {partnerVT && (
            <span className="font-bold bg-slate-100 text-slate-600 px-2 py-0.5 rounded">
              Partner type: {partnerVT}
            </span>
          )}
          {vtMismatch && (
            <span className="font-black text-rose-600 bg-rose-50 border border-rose-200 px-2 py-0.5 rounded">
              ⚠ Type mismatch!
            </span>
          )}
        </div>
      )}

      {typeof pair.confidence === "number" && (
        <div className="flex items-center gap-4 rounded-2xl bg-slate-50 p-4">
          <Ring value={conf} size={84} stroke={9} color={confColor(conf)}>
            <span className="text-base font-black text-slate-800">{pct(conf)}</span>
          </Ring>
          <div>
            <div className="text-[10px] font-bold uppercase tracking-widest text-slate-400">
              Confidence
            </div>
            <div className="text-lg font-black" style={{ color: confColor(conf) }}>
              {confLabel(conf)}
            </div>
            <div className="text-[10px] text-slate-400 font-semibold mt-0.5">
              {e?.method === "ai" ? "AI semantic + rule validated" : "Rule engine"}
            </div>
          </div>
        </div>
      )}

      <p className="text-xs text-slate-500 font-medium bg-slate-50 p-3 rounded-xl border border-slate-100 leading-relaxed">
        <strong>Insight:</strong> {pair.note}
      </p>

      {e && (
        <div className="rounded-2xl border border-slate-100 p-4 space-y-3">
          <div className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
            Matching Evidence
          </div>
          <SignalBar label="Passport / ID" value={e.passport} />
          <SignalBar label="Reference" value={e.reference} />
          <SignalBar label="Name" value={e.name} />
          <SignalBar label="Amount" value={e.effectiveAmount ?? e.amount} />
          <SignalBar label="Date" value={e.date} />
          {e.dateDeltaDays !== null && (
            <div className="text-[10px] text-slate-400 font-semibold">
              Date gap: {e.dateDeltaDays} day{e.dateDeltaDays === 1 ? "" : "s"}
            </div>
          )}
        </div>
      )}

      <div className="grid grid-cols-2 gap-3">
        <SidePanel title="Internal Ledger" row={pair.ours} accent={NAVY} />
        <SidePanel title="Partner Ledger" row={pair.partner} accent="#7c3aed" />
      </div>

      <div className="rounded-2xl p-4 text-white space-y-2.5" style={{ background: NAVY }}>
        <div className="flex justify-between text-[10px] font-bold text-white/50 uppercase">
          <span>Variance</span>
          <AlertCircle className="size-3.5 text-rose-400" />
        </div>
        <RowWhite k="Internal" v={money(pair.oursAmt)} />
        <RowWhite k="Partner" v={money(pair.partnerAmt)} />
        <div className="flex justify-between pt-1.5 border-t border-white/10">
          <span className="text-[10px] font-black uppercase" style={{ color: GOLD }}>
            Gap
          </span>
          <span
            className={`text-sm font-black ${Math.abs(pair.diff) > 1 ? "text-rose-400" : "text-white"}`}
          >
            {signed(pair.diff)}
          </span>
        </div>
      </div>

      {pair.aiInsight && (
        <div className="p-4 bg-violet-50/60 border border-violet-100 rounded-2xl">
          <div className="text-[10px] font-black text-violet-600 uppercase mb-2 flex items-center gap-1">
            <Sparkles className="size-3.5" /> AI Reasoning
          </div>
          <p className="text-xs font-semibold text-violet-900 leading-relaxed italic">
            "{pair.aiInsight}"
          </p>
        </div>
      )}
    </aside>
  );
}

function RowWhite({ k, v }: { k: string; v: string }) {
  return (
    <div className="flex justify-between text-[11px]">
      <span className="text-white/50 font-bold">{k}</span>
      <span className="font-black tabular-nums">{v}</span>
    </div>
  );
}

function SidePanel({ title, row, accent }: { title: string; row: LedgerRow | null; accent?: string }) {
  return (
    <div className="rounded-xl border border-slate-100 bg-slate-50/40 p-3 text-[10px]">
      <div
        className="text-[9px] uppercase font-black mb-2 tracking-tighter"
        style={{ color: accent ?? "#64748b" }}
      >
        {title}
      </div>
      {!row ? (
        <div className="italic text-slate-400 py-6 text-center font-bold">⚠ MISSING</div>
      ) : (
        <div className="space-y-2 text-slate-700">
          <Field k="Date" v={row.date || "—"} />
          <Field k="Party / Passenger" v={row.paxName || "—"} mono={false} />
          <Field k="ID / Passport" v={row.passport || "—"} />
          <Field k="Reference" v={row.reference || "—"} />
          {row.description && row.description !== row.paxName && (
            <Field k="Description" v={row.description} mono={false} />
          )}
          {row.visaType && <Field k="Visa Type" v={row.visaType} />}
          <div className="h-px bg-slate-200/60" />
          <div className="flex justify-between font-black text-slate-800">
            <span>DR {money(row.charge)}</span>
            <span>CR {money(row.credit)}</span>
          </div>
          {row.isReversal && (
            <div className="text-[9px] font-bold text-yellow-800 bg-yellow-50 border border-yellow-200 px-2 py-1 rounded leading-relaxed">
              ↩ <b>Reversal (VR).</b> This entry takes money back / cancels an earlier
              charge — for example a wrong invoice, wrong client, or a refund.
            </div>
          )}
          {(row.duplicateCount ?? 0) > 1 && (
            <div className="text-[9px] font-bold text-rose-700 bg-rose-50 border border-rose-200 px-2 py-1 rounded leading-relaxed">
              ⧉ <b>Same entry twice.</b> This exact charge is written {row.duplicateCount} times in
              the {title} sheet. This is copy {row.duplicateIndex} of {row.duplicateCount} — one of
              them is likely a mistake to remove.
            </div>
          )}
          {!!row.raw?.consolidated && Array.isArray(row.raw?.componentAmounts) && (
            <div className="text-[9px] font-bold text-blue-700 bg-blue-50 border border-blue-200 px-2 py-1 rounded leading-relaxed">
              ⊕ <b>Visa + deposit joined.</b> We added{" "}
              {(row.raw.componentAmounts as number[]).map((a) => money(a)).join(" + ")} ={" "}
              {money(row.charge)} so it matches the supplier's single line.
            </div>
          )}
          {!!row.raw?.isGroupRow && row.raw?.paxCount != null && (
            <div className="text-[9px] font-bold text-indigo-700 bg-indigo-50 border border-indigo-200 px-2 py-1 rounded leading-relaxed">
              👥 <b>Group booking.</b> One row for {String(row.raw.paxCount)} people — this is person{" "}
              {String(row.raw.paxIndex)} of {String(row.raw.paxCount)}
              {row.raw.explodedGroupAmt != null && `, full total ${money(Number(row.raw.explodedGroupAmt))}`}.
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function Field({ k, v, mono = true }: { k: string; v: string; mono?: boolean }) {
  return (
    <div>
      <div className="text-[8px] font-bold uppercase text-slate-400">{k}</div>
      <div className={`font-bold break-words ${mono ? "font-mono" : ""}`}>{v}</div>
    </div>
  );
}
