import type { StockRow } from "../App";
import type { LiveStockEntry } from "./rentmanClient";

// Shaped panels (Triangle/Curved) emit one StockRow per rotation
// orientation, coded `${baseCode}-${orientation}` (e.g. "12398-LU" for
// Triangle Left-Up) - see stockRows in App.tsx. All orientations of the
// same physical item share one Rentman equipment record, so lookups key
// off the orientation-stripped base code.
export const baseCodeOf = (code: string): string => code.replace(/-(LU|LD|RU|RD)$/, "");

const STORAGE_KEY = "ledCablingRentmanStockOverrides:v1";

/** Confirmed current Rentman quantity per base code, applied globally (this is a fact about the business's inventory, not one project) - see App.tsx's exportJson comment for why this deliberately isn't saved in the project file. */
export type StockOverrides = Record<string, number>;

export function loadStockOverrides(): StockOverrides {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (err) {
    console.error("Rentman stock overrides were invalid, ignoring", err);
    return {};
  }
}

export function saveStockOverrides(overrides: StockOverrides): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(overrides));
  } catch (err) {
    console.error("Failed to save Rentman stock overrides", err);
  }
}

/**
 * Overlays confirmed Rentman stock onto stockRows, keyed by each row's
 * orientation-stripped code. A row with no override passes through
 * unchanged - this is what makes the whole feature optional: an empty
 * overrides map is a full no-op, so nothing about the existing
 * required/spare/rounded math or its consumers changes until you actually
 * click Get Current Stock and apply a result.
 */
export function applyStockOverrides(rows: StockRow[], overrides: StockOverrides): StockRow[] {
  return rows.map((row) => {
    const override = overrides[baseCodeOf(row.code)];
    if (override === undefined) return row;
    const rounded = row.rounded ?? row.required;
    return { ...row, stock: override, net: override - rounded };
  });
}

export type StockComparisonRow = {
  code: string;
  localName: string;
  /** Rentman's own name for this code, or null if the code didn't resolve to any Rentman equipment - shown alongside the quantity so a code that resolves to the WRONG item (confirmed to happen for a few codes in this catalog) is visible before applying. */
  rentmanName: string | null;
  oldQuantity: number;
  newQuantity: number | null;
};

/** Pairs each distinct catalog row's currently-effective stock with what Rentman just returned, for the Get Current Stock review modal. `currentRows` should come from visibleStockRows (i.e. already reflects any previously-applied override), one entry per distinct base code. */
export function buildStockComparison(
  currentRows: Array<{ code: string; name: string; stock: number }>,
  fetched: Record<string, LiveStockEntry | null>,
): StockComparisonRow[] {
  return currentRows.map((row) => {
    const entry = fetched[row.code];
    return {
      code: row.code,
      localName: row.name,
      rentmanName: entry ? entry.name : null,
      oldQuantity: row.stock,
      newQuantity: entry ? entry.currentQuantity : null,
    };
  });
}
