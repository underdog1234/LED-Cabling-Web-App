import type { StockRow } from "../App";

// Shaped panels (Triangle/Curved) emit one StockRow per rotation
// orientation, coded `${baseCode}-${orientation}` (e.g. "12398-LU" for
// Triangle Left-Up) - see stockRows in App.tsx. All orientations of the
// same physical item share one Rentman equipment record, so the mapping
// and live-data lookups key off the orientation-stripped base code.
export const baseCodeOf = (code: string): string => code.replace(/-(LU|LD|RU|RD)$/, "");

// Synthetic, non-equipment stock rows with no Rentman equivalent - derived
// "boxes required" display lines where stock is deliberately set equal to
// required so net is always 0 (see stockRows in App.tsx). Never offered in
// the equipment-mapping UI and never overlaid with live data.
const NON_EQUIPMENT_CODES = new Set(["BOX-MG9", "BOX-MT"]);

export const isMappableStockRow = (row: Pick<StockRow, "code">): boolean => !NON_EQUIPMENT_CODES.has(row.code);

export type LiveStockEntry = { name: string; currentQuantity: number };

/**
 * Overlays live Rentman stock/availability onto stockRows, keyed by each
 * row's orientation-stripped code. A row whose code has no mapping, or is
 * mapped but missing from the fetched data (stale/deleted Rentman
 * equipment, or one failed lookup in a batch), passes through completely
 * unchanged - never NaN/undefined. This is what makes the whole feature
 * optional: empty liveStock/liveAvailable maps are a full no-op, so nothing
 * about the existing required/spare/rounded math or its consumers changes
 * when Rentman isn't configured.
 */
export function applyLiveRentmanData(
  rows: StockRow[],
  liveStock: Record<string, LiveStockEntry>,
  liveAvailable: Record<string, number>,
): StockRow[] {
  return rows.map((row) => {
    const live = liveStock[baseCodeOf(row.code)];
    if (!live) return row;
    const rounded = row.rounded ?? row.required;
    const next: StockRow = { ...row, stock: live.currentQuantity, net: live.currentQuantity - rounded };
    const available = liveAvailable[baseCodeOf(row.code)];
    if (typeof available === "number") next.available = available;
    return next;
  });
}
