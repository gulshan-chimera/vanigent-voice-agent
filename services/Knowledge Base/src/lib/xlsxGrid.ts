import * as XLSX from "xlsx";
export interface SheetGrid {
  sheetName: string;
  rowCount: number;
  columnCount: number;
  /** Merged ranges in A1 notation, e.g. "A3:A4". A label spanning rows
   *  or columns only appears once in the cell data, so the model needs
   *  these to know a header covers more than one row or column. */
  mergedRanges: string[];
  /** One line per non-empty cell: "A3: 2026 Class" */
  cellLines: string[];
  /** True when rows were dropped to stay inside the row cap. */
  truncated: boolean;
}

export interface SkippedSheet {
  sheetName: string;
  reason: string;
}

export interface WorkbookGrid {
  sheets: SheetGrid[];
  skipped: SkippedSheet[];
  totalSheets: number;
}

// A sheet beyond this many rows can't fit in a prompt alongside its own
// instructions. Above the cap we keep the first N rows so headers and a
// representative sample survive, and flag it.
const MAX_ROWS_PER_SHEET = 300;

// Workbooks routinely carry a blank "Sheet2", or a three-cell notes tab.
// Processing those costs a model call and returns nothing useful.
const MIN_NON_EMPTY_CELLS = 10;

/**
 * Reads a cell's DISPLAY value where one exists, falling back to the raw
 * value. This matters more than it looks: Excel stores a date as a serial
 * number and keeps the "$" of a currency in the format string, so raw
 * values would give a model five-digit integers to describe.
 */
function readCell(cell: XLSX.CellObject | undefined): string {
  if (!cell) return "";

  // `w` is the formatted text as Excel would display it.
  if (typeof cell.w === "string" && cell.w.trim().length > 0) {
    return cell.w.trim();
  }

  if (cell.v === undefined || cell.v === null) return "";

  if (cell.v instanceof Date) {
    return cell.v.toISOString().slice(0, 10);
  }

  return String(cell.v).trim();
}

/**
 * Reads the hidden flag from workbook metadata. SheetJS reports
 * 0 = visible, 1 = hidden, 2 = very hidden. Hidden sheets usually hold
 * lookup tables or working calculations that nobody intended to publish,
 * so they're excluded — worth knowing this is a deliberate choice, since
 * a hidden sheet that DOES matter will silently never be indexed.
 */
function isHidden(workbook: XLSX.WorkBook, sheetIndex: number): boolean {
  const state = workbook.Workbook?.Sheets?.[sheetIndex]?.Hidden;
  return typeof state === "number" && state > 0;
}

export function extractWorkbookGrid(base64Content: string): WorkbookGrid | null {
  try {
    const buffer = Buffer.from(base64Content, "base64");

    // cellDates converts date serials to Date objects; cellNF/cellText
    // keep number formats and formatted text so readCell can use them.
    const workbook = XLSX.read(buffer, {
      type: "buffer",
      cellDates: true,
      cellNF: true,
      cellText: true,
    });

    const includeHidden = process.env.KB_INCLUDE_HIDDEN_SHEETS === "true";
    const sheets: SheetGrid[] = [];
    const skipped: SkippedSheet[] = [];

    workbook.SheetNames.forEach((sheetName, sheetIndex) => {
    if (!includeHidden && isHidden(workbook, sheetIndex)) {
      console.log(`[XLSX-GRID] Sheet "${sheetName}" is hidden — skipping.`);
      skipped.push({ sheetName, reason: "hidden" });
      return;
    }

      const sheet = workbook.Sheets[sheetName];

      if (!sheet?.["!ref"]) {
        console.log(`[XLSX-GRID] Sheet "${sheetName}" is empty — skipping.`);
        skipped.push({ sheetName, reason: "empty" });
        return;
      }

      const range = XLSX.utils.decode_range(sheet["!ref"]);
      const totalRows = range.e.r - range.s.r + 1;
      const columnCount = range.e.c - range.s.c + 1;

      const lastRow = Math.min(range.e.r, range.s.r + MAX_ROWS_PER_SHEET - 1);
      const truncated = lastRow < range.e.r;

      if (truncated) {
        console.warn(
          `[XLSX-GRID] Sheet "${sheetName}" has ${totalRows} rows — truncating to ${MAX_ROWS_PER_SHEET}.`
        );
      }

      const cellLines: string[] = [];

      for (let r = range.s.r; r <= lastRow; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
          const address = XLSX.utils.encode_cell({ r, c });
          const value = readCell(sheet[address] as XLSX.CellObject | undefined);

          // Skip empty cells — including them would multiply prompt size
          // for no information gain.
          if (value.length === 0) continue;

          // Collapse newlines: a wrapped cell would otherwise break the
          // one-line-per-cell format the model relies on.
          cellLines.push(`${address}: ${value.replace(/\s*\n\s*/g, " / ")}`);
        }
      }

      if (cellLines.length < MIN_NON_EMPTY_CELLS) {
        console.log(
          `[XLSX-GRID] Sheet "${sheetName}" has only ${cellLines.length} non-empty cell(s) — skipping as junk.`
        );
        skipped.push({
          sheetName,
          reason: `too few cells (${cellLines.length})`,
        });
        return;
      }

      const mergedRanges = (sheet["!merges"] ?? []).map((m) =>
        XLSX.utils.encode_range(m)
      );

      sheets.push({
        sheetName,
        rowCount: totalRows,
        columnCount,
        mergedRanges,
        cellLines,
        truncated,
      });

      console.log(
        `[XLSX-GRID] Sheet "${sheetName}": ${totalRows} row(s), ${columnCount} column(s), ${cellLines.length} non-empty cell(s), ${mergedRanges.length} merged range(s).`
      );
    });

    if (sheets.length === 0) {
      console.warn(
        `[XLSX-GRID] No usable sheets found. Skipped: ${skipped.map((s) => `${s.sheetName} (${s.reason})`).join(", ")}`
      );
      return null;
    }

    return { sheets, skipped, totalSheets: workbook.SheetNames.length };
  } catch (error) {
    console.error(`[XLSX-GRID] Extraction failed: ${(error as Error).message}`);
    return null;
  }
}