import type { StockRow } from "../App";

// ---------------------------------------------------------------------------
// Manual changes to the stock list.
//
// The tool works out what a project needs from the layout, and it is right
// about the layout - but it does not know that the last job left six spare
// cables on the truck, or that the client is supplying the distro, or that a
// line simply should not be pulled this time. So every row can have its
// quantity typed over, or be taken off the list altogether.
//
// An edit is a note on ONE row of ONE project: it is saved in the project file
// and travels with it, unlike the Rentman stock overrides next door, which are
// a fact about the warehouse and are stored per browser (see stockOverrides).
//
// The calculated figure is never thrown away. A row that has been edited
// carries what the tool worked out alongside the number that replaced it, so
// the table, the CSV and the report can all show both - and "Reset to
// calculated" is just dropping these notes, never a recalculation that could
// come out differently.
// ---------------------------------------------------------------------------

/** One row's manual change: a replacement order quantity, or off the list. */
export type StockEdit = {
  /** Replacement TOTAL required (the order/pull quantity). */
  qty?: number;
  /** Taken off the list entirely - no row, in the table or any export. */
  removed?: boolean;
};

/** Manual changes for one project, keyed by the stock row's code. */
export type StockEdits = Record<string, StockEdit>;

/** The order quantity a row asks for before any manual change. */
export const calculatedTotalOf = (row: StockRow): number => row.rounded ?? row.required;

/**
 * Overlays manual changes onto stock rows: removed rows disappear, an edited
 * quantity replaces the total (and what is left over against stock), and every
 * other row passes through untouched - so an empty map is a full no-op.
 */
export function applyStockEdits(rows: StockRow[], edits: StockEdits): StockRow[] {
  const out: StockRow[] = [];
  for (const row of rows) {
    const edit = edits[row.code];
    if (!edit) {
      out.push(row);
      continue;
    }
    if (edit.removed) continue;
    const calculated = calculatedTotalOf(row);
    if (edit.qty === undefined || edit.qty === calculated) {
      out.push(row);
      continue;
    }
    out.push({ ...row, rounded: edit.qty, net: row.stock - edit.qty, calculated, edited: true });
  }
  return out;
}

/** The rows a manual change has taken off the list, for the "put it back" strip. */
export function removedStockRows(rows: StockRow[], edits: StockEdits): StockRow[] {
  return rows.filter((row) => edits[row.code]?.removed);
}

/** How many rows carry a manual change - what the "edited" count in the header reads. */
export function stockEditCount(rows: StockRow[], edits: StockEdits): number {
  return rows.reduce((count, row) => {
    const edit = edits[row.code];
    if (!edit) return count;
    if (edit.removed) return count + 1;
    return edit.qty !== undefined && edit.qty !== calculatedTotalOf(row) ? count + 1 : count;
  }, 0);
}

/** Sets one row's quantity, dropping the note entirely when it matches the calculated figure again. */
export function withStockQty(edits: StockEdits, code: string, qty: number | null, calculated: number): StockEdits {
  const next = { ...edits };
  const existing = next[code];
  if (qty === null || qty === calculated) {
    if (!existing) return edits;
    if (existing.removed) next[code] = { removed: true };
    else delete next[code];
    return next;
  }
  next[code] = { ...existing, qty: Math.max(0, Math.round(qty)) };
  return next;
}

/** Takes a row off the list, or puts it back. */
export function withStockRemoved(edits: StockEdits, code: string, removed: boolean): StockEdits {
  const next = { ...edits };
  const existing = next[code];
  if (removed) next[code] = { ...existing, removed: true };
  else if (!existing) return edits;
  else if (existing.qty === undefined) delete next[code];
  else next[code] = { qty: existing.qty };
  return next;
}

/** Drops every manual change on one row, quantity and removal alike. */
export function withStockRowReset(edits: StockEdits, code: string): StockEdits {
  if (!edits[code]) return edits;
  const next = { ...edits };
  delete next[code];
  return next;
}

/** Whatever was loaded from a project file, reduced to changes this app understands. */
export function normalizeStockEdits(raw: unknown): StockEdits {
  if (!raw || typeof raw !== "object") return {};
  const out: StockEdits = {};
  for (const [code, value] of Object.entries(raw as Record<string, unknown>)) {
    if (!value || typeof value !== "object") continue;
    const edit = value as { qty?: unknown; removed?: unknown };
    const next: StockEdit = {};
    if (typeof edit.qty === "number" && Number.isFinite(edit.qty)) next.qty = Math.max(0, Math.round(edit.qty));
    if (edit.removed === true) next.removed = true;
    if (next.qty !== undefined || next.removed) out[code] = next;
  }
  return out;
}
