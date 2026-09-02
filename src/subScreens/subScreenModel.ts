// Pure, framework-free helpers for the Sub-Screens feature. Kept separate
// from App.tsx so the sub-screen list UI can import just the data logic it
// needs without pulling in the rest of the monolith.
import { type RectMm, activeBBox } from "../model/panels";
import { SUB_SCREEN_COLORS, type Cell, type SubScreen } from "../App";

let subScreenIdCounter = 0;
export const newSubScreenId = (): string => {
  try {
    return crypto.randomUUID();
  } catch {
    subScreenIdCounter += 1;
    return `ss-${Date.now().toString(36)}-${subScreenIdCounter}`;
  }
};

/** `existingCount` picks the next default identity colour, so consecutively created sub-screens never start out the same colour. */
export const makeSubScreen = (name: string, createdAt = Date.now(), existingCount = 0): SubScreen => ({
  id: newSubScreenId(),
  name,
  canvasX: 0,
  canvasY: 0,
  createdAt,
  color: SUB_SCREEN_COLORS[existingCount % SUB_SCREEN_COLORS.length],
});

// A rect-producing helper matching the shape App.tsx's own `cellRect` uses,
// passed in rather than imported to avoid a circular dependency on App.tsx.
export type CellRectFn = (cell: Cell) => RectMm;

/** Bounding box (workspace mm) of a sub-screen's active member panels. */
export const subScreenBBoxOf = (panels: Cell[], subScreenId: string, cellRect: CellRectFn): RectMm =>
  activeBBox(
    panels
      .filter((cell) => cell.subScreenId === subScreenId && !cell.isRemoved)
      .map(cellRect),
  );

/** Panel count + physical size for a sub-screen, for list/summary display. */
export const subScreenPanelCount = (panels: Cell[], subScreenId: string): number =>
  panels.filter((cell) => cell.subScreenId === subScreenId && !cell.isRemoved).length;
