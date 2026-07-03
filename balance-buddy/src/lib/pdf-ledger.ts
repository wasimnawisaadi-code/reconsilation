// PDF statement → table (AOA) conversion.
//
// Many banks and suppliers only send PDF statements. For TEXT-BASED PDFs we
// extract every positioned text fragment with pdf.js, rebuild lines from the
// Y coordinates, and rebuild columns from the X coverage gaps — producing the
// same AOA shape the spreadsheet readers emit, so the whole existing pipeline
// (header detection, column mapping, matching) works unchanged.
//
// Scanned/image PDFs have no text layer; they fail with a clear message
// telling the user to export a digital statement instead.
//
// pdf.js is imported dynamically inside pdfBufToAoa only — importing THIS
// module stays cheap and Node-safe, so the pure geometry functions below are
// unit-testable without pdf.js.

/** One positioned text fragment from a PDF page. */
export type PdfTextItem = {
  str: string;
  /** X of the fragment's left edge (PDF points). */
  x: number;
  /** Y of the fragment's baseline (PDF points; larger = higher on page). */
  y: number;
  /** Rendered width in points. */
  w: number;
};

/** Vertical tolerance: fragments within this many points share a line. */
const LINE_TOL = 3;
/** A horizontal gap in text coverage at least this wide splits two columns. */
const COL_GAP = 7;

/**
 * Rebuild a table from positioned text fragments.
 *
 * Column edges are derived from the WHOLE document (all pages) so every page
 * lands in the same columns: project all fragments of multi-fragment lines
 * onto the X axis, merge overlapping/near intervals, and treat each surviving
 * band as one column. Right-aligned amount columns work because the band is
 * built from coverage, not from left edges.
 */
export function pdfItemsToAoa(pages: PdfTextItem[][]): unknown[][] {
  // ── 1. Group each page's fragments into lines by Y ──────────────────────
  type Line = { y: number; items: PdfTextItem[] };
  const allLines: Line[] = [];
  for (const items of pages) {
    const sorted = [...items]
      .filter((i) => i.str.trim())
      .sort((a, b) => b.y - a.y || a.x - b.x);
    let line: Line | null = null;
    for (const it of sorted) {
      if (!line || Math.abs(line.y - it.y) > LINE_TOL) {
        line = { y: it.y, items: [] };
        allLines.push(line);
      }
      line.items.push(it);
    }
  }

  // ── 2. Column bands from X coverage of table-ish lines (≥3 fragments) ───
  const spans: { a: number; b: number }[] = [];
  for (const ln of allLines) {
    if (ln.items.length < 3) continue;
    for (const it of ln.items) spans.push({ a: it.x, b: it.x + Math.max(it.w, 1) });
  }
  if (!spans.length) {
    // No table-like lines at all — emit one cell per line so downstream
    // parsers still see the text.
    return allLines.map((ln) => [ln.items.map((i) => i.str.trim()).join(" ")]);
  }
  spans.sort((s, t) => s.a - t.a);
  const bands: { a: number; b: number }[] = [];
  for (const s of spans) {
    const last = bands[bands.length - 1];
    if (last && s.a <= last.b + COL_GAP) last.b = Math.max(last.b, s.b);
    else bands.push({ ...s });
  }

  // ── 3. Assign fragments to bands; join fragments sharing a cell ─────────
  const bandOf = (it: PdfTextItem): number => {
    const cx = it.x + Math.max(it.w, 1) / 2;
    for (let i = 0; i < bands.length; i++) {
      if (cx >= bands[i].a - COL_GAP && cx <= bands[i].b + COL_GAP) return i;
    }
    return bands.length - 1;
  };
  const aoa: unknown[][] = [];
  for (const ln of allLines) {
    const row: string[] = new Array(bands.length).fill("");
    for (const it of ln.items.sort((a, b) => a.x - b.x)) {
      const c = bandOf(it);
      row[c] = row[c] ? `${row[c]} ${it.str.trim()}` : it.str.trim();
    }
    if (row.some((c) => c !== "")) aoa.push(row);
  }
  return aoa;
}

/**
 * Convert a PDF buffer to an AOA using pdf.js (dynamic import; sets up the
 * worker in the browser). Throws a user-actionable error for scanned PDFs.
 */
export async function pdfBufToAoa(buf: ArrayBuffer, fileName = "This PDF"): Promise<unknown[][]> {
  const pdfjs = await import("pdfjs-dist");
  if (typeof window !== "undefined" && !pdfjs.GlobalWorkerOptions.workerSrc) {
    const worker = await import("pdfjs-dist/build/pdf.worker.min.mjs?url");
    pdfjs.GlobalWorkerOptions.workerSrc = worker.default;
  }

  const doc = await pdfjs.getDocument({ data: new Uint8Array(buf.slice(0)) }).promise;
  const pages: PdfTextItem[][] = [];
  try {
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      const tc = await page.getTextContent();
      pages.push(
        (tc.items as Array<{ str?: string; transform?: number[]; width?: number }>)
          .filter((i) => typeof i.str === "string" && i.str.trim() && Array.isArray(i.transform))
          .map((i) => ({
            str: i.str as string,
            x: i.transform![4],
            y: i.transform![5],
            w: i.width ?? 0,
          })),
      );
    }
  } finally {
    await (doc as { destroy?: () => Promise<void> }).destroy?.().catch(() => undefined);
  }

  const totalFragments = pages.reduce((s, p) => s + p.length, 0);
  if (totalFragments < 5) {
    throw new Error(
      `"${fileName}" looks like a scanned/image PDF with no selectable text — ` +
        `it can't be read directly. Please export a digital statement (Excel/CSV, ` +
        `or a text-based PDF) from the bank/supplier portal instead.`,
    );
  }
  return pdfItemsToAoa(pages);
}
