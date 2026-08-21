// The code -> Rentman equipment mapping is account-wide (equipment IDs
// belong to the user's Rentman account, not to any one wall/project), so it
// lives in localStorage rather than a project's saved JSON. Unlike this
// codebase's other localStorage keys (QUICK_LAYOUT_TRANSFER_KEY,
// "ledCablingTestPattern:v1" in App.tsx), which are one-shot handoffs read
// once then removed, this is a durable setting loaded on mount and
// persisted on every edit - App.tsx owns the actual useState and calls
// these two functions rather than touching localStorage directly.

const STORAGE_KEY = "ledCablingRentmanMapping:v1";

export type RentmanEquipmentRef = { id: number; code: string; name: string };
export type EquipmentMapping = Record<string, RentmanEquipmentRef>;

export function loadEquipmentMapping(): EquipmentMapping {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (err) {
    console.error("Rentman equipment mapping was invalid, ignoring", err);
    return {};
  }
}

export function saveEquipmentMapping(mapping: EquipmentMapping): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(mapping));
  } catch (err) {
    console.error("Failed to save Rentman equipment mapping", err);
  }
}
