import { createFileRoute } from "@tanstack/react-router";
import React, { useMemo, useState, useEffect } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  parseOurLedger,
  parsePartnerLedger,
  reconcile,
  exportPairsCSV,
  buildReconciliationWorkbook,
  parseDynamicLedger,
  computeTotals,
  computeAnalytics,
  scoreRowPair,
  type ReconResult,
  type Pair,
  type LedgerRow,
  type ColumnMapping,
  type MatchEvidence,
  type ReconAnalytics,
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
  Tag,
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
  | Pair["status"];

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

/** Map a scenario category to the table filter that isolates it. */
function scenarioToFilter(key: Scenario): StatusFilter {
  switch (key) {
    case "security_deposit":
      return "security_deposit";
    case "multi_passenger":
      return "multi_passenger";
    case "bank_transfer":
      return "payments";
    case "wrong_invoice":
    case "wrong_client":
    case "duplicate":
    case "refund":
      return "refunds";
    default:
      return "all";
  }
}

/* ---- Advanced analytics: per-scenario performance breakdown ---- */
function ScenarioIntelligenceCard({
  analytics,
  onPick,
}: {
  analytics: ReconAnalytics;
  onPick: (key: Scenario) => void;
}) {
  const rows = analytics.scenarios;
  const maxTotal = Math.max(1, ...rows.map((r) => r.total));
  return (
    <div className="rounded-2xl border border-slate-200/70 bg-white p-6 shadow-sm">
      <div className="flex items-center justify-between mb-5">
        <h3 className="text-xs font-bold uppercase tracking-wider text-slate-400">
          Scenario Intelligence
        </h3>
        <span className="inline-flex items-center gap-1 text-[10px] font-bold text-slate-400">
          <Tag className="size-3" /> matched / total · value
        </span>
      </div>
      <div className="space-y-3">
        {rows.map((r) => {
          const st = SCENARIO_STYLE[r.key];
          const rate = r.total ? Math.round((r.matched / r.total) * 100) : 0;
          const seg = (n: number) => (r.total ? (n / r.total) * 100 : 0);
          return (
            <button
              key={r.key}
              onClick={() => onPick(r.key)}
              className="w-full text-left group"
              title={`${r.label} — click to filter`}
            >
              <div className="flex items-center justify-between mb-1">
                <span className="flex items-center gap-1.5 text-[11px] font-bold text-slate-600">
                  <span className={`size-2 rounded-full ${st?.dot ?? "bg-slate-400"}`} />
                  {r.label}
                </span>
                <span className="flex items-center gap-2 text-[11px]">
                  <span className="font-black text-slate-700">
                    {r.matched}/{r.total}
                  </span>
                  <span className="text-slate-400 tabular-nums">{money(r.matchedValue)}</span>
                  <span
                    className={`font-bold tabular-nums ${rate >= 90 ? "text-emerald-600" : rate >= 60 ? "text-amber-600" : "text-rose-600"}`}
                  >
                    {rate}%
                  </span>
                </span>
              </div>
              {/* Stacked status bar */}
              <div
                className="flex h-2.5 w-full overflow-hidden rounded-full bg-slate-100"
                style={{ opacity: 0.95 }}
              >
                <span className="bg-emerald-500" style={{ width: `${seg(r.matched)}%` }} />
                <span className="bg-amber-400" style={{ width: `${seg(r.amountDiff)}%` }} />
                <span className="bg-indigo-400" style={{ width: `${seg(r.onlyOurs)}%` }} />
                <span className="bg-rose-400" style={{ width: `${seg(r.onlyPartner)}%` }} />
              </div>
              <div className="mt-0.5 flex gap-3 text-[9px] font-semibold text-slate-400 opacity-0 group-hover:opacity-100 transition-opacity">
                {r.amountDiff > 0 && <span className="text-amber-600">⚠ {r.amountDiff} diff</span>}
                {r.onlyOurs > 0 && <span className="text-indigo-500">◀ {r.onlyOurs} only ours</span>}
                {r.onlyPartner > 0 && (
                  <span className="text-rose-500">▶ {r.onlyPartner} only partner</span>
                )}
              </div>
              <div
                className="mt-1 h-px bg-slate-50"
                style={{ width: `${(r.total / maxTotal) * 100}%` }}
              />
            </button>
          );
        })}
      </div>
      <div className="mt-4 flex flex-wrap gap-3 text-[9px] font-bold uppercase tracking-wide text-slate-400">
        <Legend color="bg-emerald-500" label="Matched" />
        <Legend color="bg-amber-400" label="Amount diff" />
        <Legend color="bg-indigo-400" label="Only ours" />
        <Legend color="bg-rose-400" label="Only partner" />
      </div>
    </div>
  );
}

function Legend({ color, label }: { color: string; label: string }) {
  return (
    <span className="flex items-center gap-1">
      <span className={`size-2 rounded-full ${color}`} /> {label}
    </span>
  );
}

/* ---- Advanced analytics: where duplicates & reversals originate ---- */
function DuplicateReversalCard({
  analytics,
  onPick,
}: {
  analytics: ReconAnalytics;
  onPick: (f: StatusFilter) => void;
}) {
  const d = analytics.duplicates;
  const rev = analytics.reversals;
  const totalDupRows = d.rowsOurs + d.rowsPartner;
  return (
    <div className="rounded-2xl border border-slate-200/70 bg-white p-6 shadow-sm flex flex-col gap-5">
      {/* Duplicates */}
      <div>
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-xs font-bold uppercase tracking-wider text-slate-400">
            Duplicate Origin
          </h3>
          <button
            onClick={() => onPick("duplicates")}
            className="text-[10px] font-black text-rose-600 hover:underline"
          >
            View {totalDupRows} →
          </button>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <OriginTile
            label="Our Ledger"
            rows={d.rowsOurs}
            groups={d.groupsOurs}
            tone="indigo"
          />
          <OriginTile
            label="Partner Ledger"
            rows={d.rowsPartner}
            groups={d.groupsPartner}
            tone="rose"
          />
        </div>
        {d.redundantValue > 0 && (
          <div className="mt-2 flex items-center justify-between rounded-lg bg-rose-50 border border-rose-100 px-3 py-1.5 text-[11px]">
            <span className="font-semibold text-rose-700">Redundant value (excl. first of group)</span>
            <span className="font-black text-rose-700 tabular-nums">{money(d.redundantValue)}</span>
          </div>
        )}
        {totalDupRows === 0 && (
          <p className="text-[11px] text-slate-400 italic">No duplicate entries detected.</p>
        )}
      </div>

      <div className="h-px bg-slate-100" />

      {/* Reversals / VR */}
      <div>
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-xs font-bold uppercase tracking-wider text-slate-400">
            Reversals / VR by Reason
          </h3>
          <button
            onClick={() => onPick("refunds")}
            className="text-[10px] font-black text-pink-600 hover:underline"
          >
            View {rev.ours + rev.partner} →
          </button>
        </div>
        {rev.byReason.length === 0 ? (
          <p className="text-[11px] text-slate-400 italic">No reversal / refund entries detected.</p>
        ) : (
          <div className="space-y-1.5">
            <div className="flex items-center text-[9px] font-black uppercase tracking-wide text-slate-300">
              <span className="flex-1">Reason</span>
              <span className="w-14 text-right text-indigo-400">Ours</span>
              <span className="w-14 text-right text-rose-400">Partner</span>
            </div>
            {rev.byReason.map((r) => {
              const st = SCENARIO_STYLE[r.reason];
              return (
                <div
                  key={r.reason}
                  className="flex items-center text-[11px] py-1 border-b border-slate-50 last:border-0"
                >
                  <span className="flex-1 flex items-center gap-1.5 font-semibold text-slate-600">
                    <span className={`size-2 rounded-full ${st?.dot ?? "bg-slate-400"}`} />
                    {r.label}
                  </span>
                  <span className="w-14 text-right font-black text-indigo-600 tabular-nums">
                    {r.ours || "—"}
                  </span>
                  <span className="w-14 text-right font-black text-rose-600 tabular-nums">
                    {r.partner || "—"}
                  </span>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

function OriginTile({
  label,
  rows,
  groups,
  tone,
}: {
  label: string;
  rows: number;
  groups: number;
  tone: "indigo" | "rose";
}) {
  const tones = {
    indigo: "from-indigo-50 to-white border-indigo-100 text-indigo-700",
    rose: "from-rose-50 to-white border-rose-100 text-rose-700",
  };
  return (
    <div className={`rounded-xl border bg-gradient-to-br p-3 ${tones[tone]}`}>
      <div className="text-[9px] font-black uppercase tracking-wide opacity-70">{label}</div>
      <div className="mt-1 text-2xl font-black tabular-nums">{rows}</div>
      <div className="text-[10px] font-semibold opacity-70">
        {groups} group{groups === 1 ? "" : "s"}
      </div>
    </div>
  );
}

type Aoa = unknown[][];

/* ================================================================== */
/*  MAIN                                                               */
/* ================================================================== */

function Index() {
  const [oursFile, setOursFile] = useState<File | null>(null);
  const [partnerFile, setPartnerFile] = useState<File | null>(null);
  const [rawOurs, setRawOurs] = useState<Aoa | null>(null);
  const [rawPartner, setRawPartner] = useState<Aoa | null>(null);
  const [result, setResult] = useState<ReconResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [isClient, setIsClient] = useState(false);
  const [aiStatus, setAiStatus] = useState<string>("");
  const [engineMode, setEngineMode] = useState<"ai" | "heuristic">("ai");

  useEffect(() => setIsClient(true), []);

  const [filter, setFilter] = useState<StatusFilter>("all");
  const [query, setQuery] = useState("");
  const [sortByConf, setSortByConf] = useState(false);
  const [selected, setSelected] = useState<Pair | null>(null);
  const [schema, setSchema] = useState<any>(null);
  const [showSource, setShowSource] = useState(false);

  const getAoa = async (file: File): Promise<Aoa> => {
    const buf = await file.arrayBuffer();
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
    setAiStatus("Reading files…");
    try {
      if (!oursFile || !partnerFile) throw new Error("Please upload both ledger files.");

      const aoaOurs = rawOurs ?? (await getAoa(oursFile));
      const aoaPartner = rawPartner ?? (await getAoa(partnerFile));
      setRawOurs(aoaOurs);
      setRawPartner(aoaPartner);

      let ours: LedgerRow[];
      let partner: LedgerRow[];
      let mode: "ai" | "heuristic" = "ai";

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
      // whichever reconciles more rows. Guarantees we never silently show 0 matches.
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

      setAiStatus("Running multi-signal rule engine…");
      const baseResult = reconcile(ours, partner);
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
          setResult({ pairs: merged, totals: computeTotals(ours, partner, merged) });
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
  }, [result, filter, query, sortByConf]);

  const chartData = useMemo(() => {
    if (!result) return [];
    const t = result.totals;
    return [
      { name: "Matched", value: t.matched, color: "#10b981" },
      { name: "Amount Diff", value: t.amountIssues, color: "#f59e0b" },
      { name: "Only Ours", value: t.onlyOurs, color: "#6366f1" },
      { name: "Only Partner", value: t.onlyPartner, color: "#ef4444" },
    ];
  }, [result]);

  const confHist = useMemo(() => {
    const buckets = [
      { name: "<60", value: 0, color: "#ef4444" },
      { name: "60-80", value: 0, color: "#f59e0b" },
      { name: "80-95", value: 0, color: "#3b82f6" },
      { name: "95+", value: 0, color: "#10b981" },
    ];
    if (!result) return buckets;
    result.pairs.forEach((p) => {
      if (typeof p.confidence !== "number") return;
      const c = p.confidence;
      if (c < 0.6) buckets[0].value++;
      else if (c < 0.8) buckets[1].value++;
      else if (c < 0.95) buckets[2].value++;
      else buckets[3].value++;
    });
    return buckets;
  }, [result]);

  const matchRate = useMemo(() => {
    if (!result) return 0;
    const paired = result.pairs.filter((p) => p.ours && p.partner).length;
    return paired / (result.pairs.length || 1);
  }, [result]);

  const matchedValue = useMemo(() => {
    if (!result) return 0;
    return +result.pairs
      .filter((p) => p.status === "matched")
      .reduce((s, p) => s + p.partnerAmt, 0)
      .toFixed(2);
  }, [result]);

  const analytics = useMemo<ReconAnalytics | null>(
    () => (result ? computeAnalytics(result.pairs) : null),
    [result],
  );

  const hasData = !!(rawOurs || rawPartner);

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
            <HeaderChip
              label="Our Ledger"
              file={oursFile}
              onChange={(f) => selectFile("ours", f)}
            />
            <HeaderChip
              label="Partner Ledger"
              file={partnerFile}
              onChange={(f) => selectFile("partner", f)}
            />
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
              disabled={!oursFile || !partnerFile || busy}
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
            partnerFile={partnerFile}
            onPick={selectFile}
            onRun={runSmartRecon}
            busy={busy}
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
            {/* ---------------- SUMMARY STRIP ---------------- */}
            <section className="grid gap-3 grid-cols-2 sm:grid-cols-3 lg:grid-cols-5">
              <SummaryCard
                label="Total Rows"
                value={result.pairs.length}
                onClick={() => setFilter("all")}
                active={filter === "all"}
                accent={NAVY}
              />
              <SummaryCard
                label="Matched"
                value={result.totals.matched}
                onClick={() => setFilter("matched")}
                active={filter === "matched"}
                accent="#10b981"
              />
              <SummaryCard
                label="Only Ours"
                value={result.totals.onlyOurs}
                onClick={() => setFilter("missing_partner")}
                active={filter === "missing_partner"}
                accent="#6366f1"
              />
              <SummaryCard
                label="Only Partner"
                value={result.totals.onlyPartner}
                onClick={() => setFilter("missing_ours")}
                active={filter === "missing_ours"}
                accent="#ef4444"
              />
              <SummaryCard
                label="Needs Review"
                value={result.totals.needsReview}
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
                caption={`${result.pairs.filter((p) => p.ours && p.partner).length} of ${result.pairs.length} rows paired`}
                icon={<CheckCircle2 className="size-4 text-emerald-500" />}
                color="#10b981"
              />
              <RingCard
                title="Avg Confidence"
                value={result.totals.avgConfidence}
                caption={confLabel(result.totals.avgConfidence) + " certainty"}
                icon={<ShieldCheck className="size-4" style={{ color: NAVY }} />}
                color={confColor(result.totals.avgConfidence)}
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
                value={String(result.totals.needsReview)}
                caption={
                  result.totals.aiAssisted > 0
                    ? `${result.totals.aiAssisted} AI-assisted matches`
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

            {/* ---------------- SCENARIO INTELLIGENCE ---------------- */}
            {isClient && analytics && analytics.scenarios.length > 0 && (
              <section className="grid gap-5 lg:grid-cols-[1.6fr_1fr]">
                <ScenarioIntelligenceCard
                  analytics={analytics}
                  onPick={(key) => setFilter(scenarioToFilter(key))}
                />
                <DuplicateReversalCard analytics={analytics} onPick={setFilter} />
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
                rows={result.totals.oursRows}
                charges={result.totals.oursCharges}
                credits={result.totals.oursCredits}
                icon={<CheckCircle2 className="size-4" style={{ color: NAVY }} />}
              />
              <TotalsCard
                title="Partner Ledger"
                rows={result.totals.partnerRows}
                charges={result.totals.partnerCharges}
                credits={result.totals.partnerCredits}
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
                  <DiffRow label="Net Amount Diff" value={result.totals.netAmountDiff} />
                  <DiffRow label="Amount Mismatches" value={result.totals.amountIssues} raw />
                  <div className="h-px bg-slate-100" />
                  <div
                    className="flex justify-between items-center text-white p-3 rounded-xl shadow-lg"
                    style={{ background: `linear-gradient(90deg, ${NAVY}, #103a73)` }}
                  >
                    <div className="text-[10px] font-bold uppercase opacity-80">
                      Unmatched Items
                    </div>
                    <div className="text-xl font-black">
                      {result.totals.onlyOurs + result.totals.onlyPartner}
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
                    [
                      ["all", "All", result.pairs.length],
                      ["matched", "Matched", result.totals.matched],
                      ["amount_diff", "Amount Mismatch", result.totals.amountIssues],
                      [
                        "payments",
                        "Bank Transfers",
                        result.pairs.filter((p) => isTransfer(p.ours) || isTransfer(p.partner)).length,
                      ],
                      [
                        "security_deposit",
                        "Security Dep.",
                        result.pairs.filter((p) => {
                          const sc = p.ours?.scenario ?? p.partner?.scenario;
                          return sc === "security_deposit";
                        }).length,
                      ],
                      [
                        "refunds",
                        "Refunds / VR",
                        result.pairs.filter((p) => {
                          const sc = p.ours?.scenario ?? p.partner?.scenario;
                          return !!sc && ["wrong_invoice", "wrong_client", "duplicate", "refund"].includes(sc);
                        }).length,
                      ],
                      [
                        "multi_passenger",
                        "Multi-Pax",
                        result.pairs.filter((p) => {
                          const sc = p.ours?.scenario ?? p.partner?.scenario;
                          return sc === "multi_passenger";
                        }).length,
                      ],
                      [
                        "duplicates",
                        "Duplicates",
                        result.pairs.filter(
                          (p) =>
                            (p.ours?.duplicateCount ?? 0) > 1 ||
                            (p.partner?.duplicateCount ?? 0) > 1,
                        ).length,
                      ],
                      ["missing_partner", "Only Ours", result.totals.onlyOurs],
                      ["missing_ours", "Only Partner", result.totals.onlyPartner],
                      ["review", "Needs Review", result.totals.needsReview],
                      [
                        "fullledger",
                        "Both Sheets (Full)",
                        (rawOurs ? Math.max(0, rawOurs.length - 1) : 0) +
                          (rawPartner ? Math.max(0, rawPartner.length - 1) : 0),
                      ],
                    ] as Array<[StatusFilter, string, number]>
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
                  <FullLedgerView ours={rawOurs} partner={rawPartner} result={result} />
                ) : (
                  <PairsTable
                    pairs={filteredPairs}
                    onSelect={setSelected}
                    selected={selected}
                    rawOurs={rawOurs}
                    rawPartner={rawPartner}
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
      </footer>
    </div>
  );
}

/* ================================================================== */
/*  UPLOAD                                                             */
/* ================================================================== */

function UploadHero({
  oursFile,
  partnerFile,
  onPick,
  onRun,
  busy,
}: {
  oursFile: File | null;
  partnerFile: File | null;
  onPick: (side: "ours" | "partner", f: File | null) => void;
  onRun: () => void;
  busy: boolean;
}) {
  return (
    <div className="rounded-3xl border border-slate-200/70 bg-white p-10 shadow-sm">
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
          A common reconciliation platform for <strong>any ledger type</strong> — bank, AR/AP,
          supplier, customer, invoices or travel. Upload both statements; a hybrid engine maps the
          columns with AI, matches every charge and payment on multiple columns, then lets AI
          resolve the hardest residual rows — each result carries a confidence score.
        </p>
      </div>

      <div className="mt-8 grid gap-5 md:grid-cols-2 max-w-3xl mx-auto">
        <UploadZone
          label="Our Ledger"
          file={oursFile}
          onChange={(f) => onPick("ours", f)}
          accent={NAVY}
        />
        <UploadZone
          label="Partner Ledger"
          file={partnerFile}
          onChange={(f) => onPick("partner", f)}
          accent={GOLD}
        />
      </div>

      <div className="mt-7 flex justify-center">
        <button
          onClick={onRun}
          disabled={!oursFile || !partnerFile || busy}
          className="rounded-xl px-7 py-3 text-sm font-bold shadow-lg transition-all disabled:opacity-40 active:scale-95 flex items-center gap-2"
          style={{ background: `linear-gradient(90deg, #d4af37, ${GOLD})`, color: NAVY }}
        >
          <Sparkles className="size-4" /> Smart Reconcile
        </button>
      </div>

      <div className="mt-8 flex flex-wrap justify-center gap-6">
        {[
          "Auto-Schema Detection",
          "Multi-Column Matching",
          "AI Residual Resolver",
          "Confidence Scoring",
        ].map((f) => (
          <div
            key={f}
            className="flex items-center gap-2 text-[11px] font-bold text-slate-500 uppercase"
          >
            <CheckCircle2 className="size-4 text-emerald-500" /> {f}
          </div>
        ))}
      </div>
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
}: {
  pairs: Pair[];
  onSelect: (p: Pair) => void;
  selected: Pair | null;
  rawOurs: Aoa | null;
  rawPartner: Aoa | null;
}) {
  const [expanded, setExpanded] = useState<string | null>(null);

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
              <th className="px-3 py-3 text-right border-l border-slate-100">Our Amt</th>
              <th className="px-3 py-3 text-left border-l border-slate-100">Partner Date</th>
              <th className="px-3 py-3 text-right border-l border-slate-100">Partner Amt</th>
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
                        {dup && dup > 1 && (
                          <span
                            className="text-[8px] font-black text-rose-600 bg-rose-50 border border-rose-200 px-1 py-0.5 rounded w-fit"
                            title={`Duplicate found in ${dupSide} ledger — entry ${dupIdx} of ${dup}`}
                          >
                            ⧉ Duplicate {dupIdx}/{dup} · {dupSide}
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
                      className={`px-3 py-2.5 text-right tabular-nums border-l border-slate-100 font-black ${
                        Math.abs(p.diff) > 0.5 ? "text-rose-600" : "text-slate-300"
                      }`}
                    >
                      {p.status === "matched" ? "✓" : signed(p.diff)}
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
}: {
  title: string;
  aoa: Aoa | null;
  map: Map<number, Pair>;
  accent: string;
  query: string;
  issuesOnly: boolean;
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
    if (issuesOnly && (!pair || pair.status === "matched")) continue;
    if (q) {
      const hay = cells.map((c) => String(c ?? "")).join(" ").toLowerCase();
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
}: {
  ours: Aoa | null;
  partner: Aoa | null;
  result: ReconResult;
}) {
  const [query, setQuery] = useState("");
  const [issuesOnly, setIssuesOnly] = useState(false);

  const { oursMap, partnerMap } = useMemo(() => {
    const om = new Map<number, Pair>();
    const pm = new Map<number, Pair>();
    result.pairs.forEach((p) => {
      if (p.ours?.srcRow != null) om.set(p.ours.srcRow, p);
      if (p.partner?.srcRow != null) pm.set(p.partner.srcRow, p);
    });
    return { oursMap: om, partnerMap: pm };
  }, [result]);

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
        aoa={ours}
        map={oursMap}
        accent={NAVY}
        query={query}
        issuesOnly={issuesOnly}
      />
      <LedgerSheet
        title="Partner Ledger (uploaded file, all columns)"
        aoa={partner}
        map={partnerMap}
        accent={GOLD}
        query={query}
        issuesOnly={issuesOnly}
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
            <div className="text-[9px] font-black text-yellow-700 bg-yellow-50 border border-yellow-200 px-2 py-1 rounded">
              ↩ Reversal / Correction Entry
            </div>
          )}
          {(row.duplicateCount ?? 0) > 1 && (
            <div className="text-[9px] font-black text-rose-700 bg-rose-50 border border-rose-200 px-2 py-1 rounded">
              ⧉ Duplicate {row.duplicateIndex}/{row.duplicateCount} — found in {title} ledger
            </div>
          )}
          {!!row.raw?.consolidated && Array.isArray(row.raw?.componentAmounts) && (
            <div className="text-[9px] font-bold text-blue-700 bg-blue-50 border border-blue-200 px-2 py-1 rounded">
              ⊕ Combined {(row.raw.componentAmounts as number[]).length} lines:{" "}
              {(row.raw.componentAmounts as number[]).map((a) => money(a)).join(" + ")} = {money(row.charge)}
            </div>
          )}
          {!!row.raw?.isGroupRow && row.raw?.paxCount != null && (
            <div className="text-[9px] font-bold text-indigo-700 bg-indigo-50 border border-indigo-200 px-2 py-1 rounded">
              👥 Group booking · passenger {String(row.raw.paxIndex)} of {String(row.raw.paxCount)}
              {row.raw.explodedGroupAmt != null && ` · total ${money(Number(row.raw.explodedGroupAmt))}`}
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
