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

/** One row's manual change: a replacement order quantity, a row off the list, or a row put on it. */
export type StockEdit = {
  /** Replacement TOTAL required (the order/pull quantity). */
  qty?: number;
  /** Taken off the list entirely - no row, in the table or any export. */
  removed?: boolean;
  /**
   * An item the tool works out no requirement for, put on the list by hand.
   * Only the code is stored: the name and the shelf quantity are looked up in
   * the catalogue when the list is built, so a project opened after the
   * catalogue changes shows today's figures rather than a stale copy.
   */
  added?: boolean;
};

/** The catalogue entry behind a code, for rows that are on the list by hand alone. */
export type StockCatalogLookup = (code: string) => { name: string; stock: number } | null;

/** Manual changes for one project, keyed by the stock row's code. */
export type StockEdits = Record<string, StockEdit>;

/** The order quantity a row asks for before any manual change. */
export const calculatedTotalOf = (row: StockRow): number => row.rounded ?? row.required;

/**
 * Overlays manual changes onto stock rows: removed rows disappear, an edited
 * quantity replaces the total (and what is left over against stock), and every
 * other row passes through untouched - so an empty map is a full no-op.
 */
export function applyStockEdits(rows: StockRow[], edits: StockEdits, lookup?: StockCatalogLookup): StockRow[] {
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
  // Then the items somebody put on the list themselves, which no calculation
  // produced a row for. They go at the end, in the order they were added, and
  // an added code that the list already holds is left to the loop above - the
  // calculated row wins, so the same item can never appear twice.
  if (lookup) {
    const present = new Set(rows.map((row) => row.code));
    for (const [code, edit] of Object.entries(edits)) {
      if (!edit.added || edit.removed || present.has(code)) continue;
      const item = lookup(code);
      if (!item) continue;
      const qty = edit.qty ?? 0;
      out.push({
        code,
        name: item.name,
        required: 0,
        stock: item.stock,
        net: item.stock - qty,
        method: "added to this project by hand",
        spare: 0,
        spareRounded: 0,
        rounded: qty,
        calculated: 0,
        edited: true,
        manual: true,
      });
    }
  }
  return out;
}

/** Codes that are on the list by hand alone, in the order they were added. */
export function addedStockCodes(edits: StockEdits): string[] {
  return Object.entries(edits)
    .filter(([, edit]) => edit.added && !edit.removed)
    .map(([code]) => code);
}

/** Puts a catalogue item on the list with a starting quantity. */
export function withStockAdded(edits: StockEdits, code: string, qty: number): StockEdits {
  return { ...edits, [code]: { added: true, qty: Math.max(0, Math.round(qty)) } };
}

/** The rows a manual change has taken off the list, for the "put it back" strip. */
export function removedStockRows(rows: StockRow[], edits: StockEdits): StockRow[] {
  return rows.filter((row) => edits[row.code]?.removed);
}

/** How many rows carry a manual change - what the "edited" count in the header reads. */
export function stockEditCount(rows: StockRow[], edits: StockEdits): number {
  const present = new Set(rows.map((row) => row.code));
  const added = Object.entries(edits).filter(([code, edit]) => edit.added && !edit.removed && !present.has(code)).length;
  return added + rows.reduce((count, row) => {
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
  // A row that is only on the list because somebody added it has no calculated
  // figure to fall back to - clearing its box leaves it at zero rather than
  // taking the row away underneath them.
  if (existing?.added) {
    next[code] = { ...existing, qty: qty === null ? 0 : Math.max(0, Math.round(qty)) };
    return next;
  }
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
  if (removed && existing?.added) {
    // Nothing calculated it, so there is no row underneath to put back.
    delete next[code];
    return next;
  }
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
    const edit = value as { qty?: unknown; removed?: unknown; added?: unknown };
    const next: StockEdit = {};
    if (typeof edit.qty === "number" && Number.isFinite(edit.qty)) next.qty = Math.max(0, Math.round(edit.qty));
    if (edit.removed === true) next.removed = true;
    if (edit.added === true) next.added = true;
    if (next.qty !== undefined || next.removed || next.added) out[code] = next;
  }
  return out;
}
