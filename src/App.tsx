import { Wand2, Zap, Download, Upload, FileText } from "lucide-react";
import React, { useEffect, useMemo, useRef, useState } from "react";
import { ImageDown, Video, LayoutGrid } from "lucide-react";
import { HelpCircle, Redo2, Undo2 } from "lucide-react";
import { Button, Card, CardHeader, CardContent, CardTitle, Input, ControlGroup, StatusChip } from "./components/ui";
import {
  type RectMm,
  type PanelAnchorSpec,
  type PanelShape,
  activeBBox,
  bandPanels,
  computeAnchorSnapDelta,
  connectedGroupsByGeom,
  findOverlaps,
  joinedGroupIdsByGeom,
  panelsAnchorJoined,
  rectsJoined,
  MODULE_MM,
} from "./model/panels";
import { parseYesTechLayout, type ImportResult } from "./import/yesTechLayout";
import SubScreenPanel from "./subScreens/SubScreenPanel";
import { makeSubScreen, subScreenBBoxOf } from "./subScreens/subScreenModel";
import OutputCanvasPanel from "./canvasView/OutputCanvasPanel";
import { finalCanvasPositionOf, subScreenResolutionOf } from "./canvasView/canvasModel";
import { subScreenPanelCount } from "./subScreens/subScreenModel";
import { type TestPatternProject, LOOP_SECONDS, DRAW_FPS, computeTestPatternLayout, drawTestPatternFrame } from "./testPattern/drawTestPattern";
import { PROCESSOR_SPECS, PROCESSOR_MODEL_IDS, type ProcessorModelId } from "./novastar/processorModels";
import { buildExportSummaryAndCabinets, buildNovaStarExport, WHOLE_LAYOUT_KEY, type CanvasEntryInput, type InputMode } from "./novastar/exportBuilder";
import NovaStarExportPanel from "./novastar/NovaStarExportPanel";
import { applyLiveRentmanData, baseCodeOf, isMappableStockRow, type LiveStockEntry } from "./rentman/applyLiveRentmanData";
import { fetchEquipmentStock, fetchEquipmentAvailability } from "./rentman/rentmanClient";
import { loadEquipmentMapping, saveEquipmentMapping, type EquipmentMapping, type RentmanEquipmentRef } from "./rentman/equipmentMapping";
import RentmanPanel, { type MappableItem } from "./rentman/RentmanPanel";

const SIGNAL_PORT_COUNT = 20;
const CELL_SIZE = 78;
const GRID_GAP = 8;
const MAX_PIXELS_PER_PORT = 650000;
const VOLTAGE = 230;
const MAX_OUTLET_AMPS = 16;
const POWER_COLOR = "#f97316";
// Chain-start (and backup-loop end) indicator outlines drawn alongside the
// existing panel borders. Blue = first panel of a signal chain (and the last
// panel too when the backup signal loop is on); orange = first panel of a power chain.
const SIGNAL_START_COLOR = "#2563eb";
const POWER_START_COLOR = POWER_COLOR;
const APP_VERSION = "0.33.0";

// Target resolution for the Panel Layout PNG embedded in the full PDF
// report (see buildLayoutCanvas) - a fixed print DPI at the page's own
// print size, not a flat pixel multiplier. The image always ends up
// shrunk to fit the SAME ~277x152mm page area (drawLayoutPage's usable
// width/height below) regardless of the wall's actual size, so scaling
// canvas resolution with the wall's mm dimensions (as a flat multiplier
// does) makes huge walls render at far more pixels - and file size - than
// that fixed print area could ever show, with zero visible quality gain.
const PDF_LAYOUT_IMAGE_DPI = 300;
const PDF_LAYOUT_USABLE_WIDTH_MM = 277; // matches drawLayoutPage's usableWidth (pageWidth - 20)
const PDF_LAYOUT_USABLE_HEIGHT_MM = 152; // matches drawLayoutPage's usableHeight (pageHeight - 58)

export const PANEL_TYPES = {
  MG9: {
    name: "MG9",
    w: 0.5,
    h: 0.5,
    pixW: 168,
    pixH: 168,
    weight: 7.4,
    power: { maxW: 175, maxA: 0.77, avgW: 59, avgA: 0.26 },
    defaults: {
      powerPanelsPerOutlet: 21,
      signalPanelsPerPort: 23,
      spareRatio: 0.07,
      panelsPerBox: 10,
      signalSpareRatio: 0.3,
      powerSpareRatio: 0.2,
      flyBarWeight: 1.9,
      slingWeight: 1.5,
    },
    stock: {
      panels: 319,
      vx1000: 2,
      vx2000: 2,
      distro32: 4,
      distro63: 4,
      powerCable15m: 85,
      signalCable15m: 53,
      hangingBar: 40,
      reinforcementPlate: 160,
      reinforcementScrew: 400,
    },
  },
  MT: {
    name: "MT",
    w: 1,
    h: 0.5,
    pixW: 256,
    pixH: 64,
    weight: 9.4,
    power: { maxW: 250, maxA: 1.09, avgW: 100, avgA: 0.44 },
    defaults: {
      powerPanelsPerOutlet: 14,
      signalPanelsPerPort: 39,
      spareRatio: 0,
      panelsPerBox: 6,
      signalSpareRatio: 0.3,
      powerSpareRatio: 0.2,
      flyBarWeight: 5.9,
      slingWeight: 1.5,
    },
    stock: {
      panels: 100,
      distro32: 0,
      distro63: 0,
      powerCable15m: 0,
      signalCable15m: 0,
      hangingBar: 10,
      reinforcementPlate: 100,
      reinforcementScrew: 400,
    },
  },
} as const;

export const POWER_DISTROS = {
  "32A": { id: "32A", label: "32A distro (9 ports)", portCount: 9, safePhaseWatts: 6900 },
  "63A": { id: "63A", label: "63A distro (18 ports)", portCount: 18, safePhaseWatts: 14500 },
} as const;

const DEPLOYMENT_TYPES = {
  FLOWN: "Flown",
  GROUND: "Ground",
  NO_SUPPORT: "No Support",
  FLOOR: "Floor",
} as const;

const STOCK_CATALOG = {
  prodCase: { code: "12317", name: "LED Prod Case", stock: 1 },
  signalJoiner: { code: "12280", name: "SEETRONIC SE8FF-05 F/M - F/M Joiner", stock: 10 },
  signalJoinerCable: { code: "12312", name: "SEETRONIC F/M - F/M Cable", stock: 11 },
  modularFrameScrew: { code: "12253", name: "YES TECH Modular Frame Installation Screw", stock: 384 },
  modularFrameUCoupler: { code: "12255", name: "YES TECH Modular Frame To Panel U-Coupler", stock: 100 },
  danceFloorRampCorner: { code: "12266", name: "YES TECH Modular Frame Dance Floor Ramp Corner", stock: 4 },
  danceFloorRamp: { code: "12267", name: "YES TECH Modular Frame Dance Floor Ramp", stock: 96 },
  modularFrame950: { code: "12268", name: "YES TECH Modular Frame 950mm x 500mm", stock: 96 },
  modularFrame860: { code: "12269", name: "YES TECH Modular Frame 860mm x 500mm (Side Piece)", stock: 3 },
  bottomBeam1m: { code: "12270", name: "YES TECH Modular Frame Bottom Beam 1m", stock: 8 },
  connectingJoint: { code: "12273", name: "YES TECH Modular Frame Connecting Joint", stock: 192 },
  danceFloorFeet: { code: "12276", name: "YES TECH Modular Frame Feet for Dance Floor Mode", stock: 384 },
  floorReinforcementBar: { code: "12274", name: "YES TECH Modular Frame Floor Reinforcement Bar", stock: 384 },
  floorTaperPin: { code: "12275", name: "YES TECH Modular Frame Floor Taper Mounting Pin", stock: 1536 },
  temperedGlass: { code: "12272", name: "YES TECH 500mm x 500mm Tempered Glass Floor Cover", stock: 384 },
  mg12Triangle: { code: "12398", name: "Triangle Panel", stock: 20 },
  mg13Curved: { code: "12399", name: "1/4 Curved Panel", stock: 20 },
  mg9Corner: { code: "12225", name: "YES TECH MG9 P2.9 500mm x 500mm LED Corner Panel", stock: 80 },
  cornerFlatConnector: { code: "12260", name: "YES TECH MG9 150 Corner Panels as Flat Connector", stock: 240 },
  cornerCornerConnector: { code: "12258", name: "YES TECH MG9 Corner Connector", stock: 160 },
} as const;

export const PANEL_VARIANTS = {
  STANDARD: { id: "STANDARD", label: "Standard MG9", symbol: "", stockItem: null, shape: "rect" },
  TRIANGLE: { id: "TRIANGLE", label: "MG12 Triangle Panel", symbol: "△", stockItem: STOCK_CATALOG.mg12Triangle, shape: "triangle" },
  CURVED: { id: "CURVED", label: "MG13 1/4 Curved Panel", symbol: "◜", stockItem: STOCK_CATALOG.mg13Curved, shape: "curve" },
  CORNER: { id: "CORNER", label: "MG9 LED Corner Panel", symbol: "Corner", stockItem: STOCK_CATALOG.mg9Corner, shape: "corner" },
} as const;

// Shaped panels (MG12 triangle / MG13 quarter circle) are physical one-way
// pieces: the location of the right-angle corner after rotation decides which
// stock unit is consumed. Mapping matches the YES TECH layout tool exactly.
const SHAPE_ORIENTATIONS = {
  LU: { key: "LU", icon: "↖", label: "Left Up" },
  LD: { key: "LD", icon: "↙", label: "Left Down" },
  RU: { key: "RU", icon: "↗", label: "Right Up" },
  RD: { key: "RD", icon: "↘", label: "Right Down" },
} as const;
type ShapeOrientationKey = keyof typeof SHAPE_ORIENTATIONS;
// Right-angle corner after clockwise rotation -> orientation bucket.
// Base shapes (rotation 0): triangle corner bottom-left (LD); sector corner bottom-right (RD).
const TRIANGLE_ORIENTATION: Record<number, ShapeOrientationKey> = { 0: "LD", 90: "LU", 180: "RU", 270: "RD" };
const SECTOR_ORIENTATION: Record<number, ShapeOrientationKey> = { 0: "RD", 90: "LD", 180: "LU", 270: "RU" };
// Per-orientation stock on the shelf (matches the layout tool's inventory).
const SHAPED_STOCK_PER_ORIENTATION = { TRIANGLE: 5, CURVED: 5 } as const;

const normalizeRotation = (rotation: number | undefined | null) =>
  ((Math.round((Number(rotation) || 0) / 90) * 90) % 360 + 360) % 360;

const getShapeOrientation = (variant: PanelVariantKey, rotation: number | undefined | null): ShapeOrientationKey | null => {
  const rot = normalizeRotation(rotation);
  if (variant === "TRIANGLE") return TRIANGLE_ORIENTATION[rot] ?? null;
  if (variant === "CURVED") return SECTOR_ORIENTATION[rot] ?? null;
  return null;
};

export const PORT_COLORS = [
  "#48d7d2",
  "#d58cff",
  "#69d54c",
  "#4968f0",
  "#fff230",
  "#f6a548",
  "#ef8c8f",
  "#71f08d",
  "#7f84ff",
  "#ffd6d8",
  "#cfe6ff",
  "#e8c7ff",
  "#a9ece7",
  "#ffc98c",
  "#fff7b8",
  "#c4ebb0",
  "#ff5bc6",
  "#944fff",
  "#27f0a4",
  "#f35c64",
];

export type PanelTypeKey = keyof typeof PANEL_TYPES;
export type PanelVariantKey = keyof typeof PANEL_VARIANTS;
export type PowerDistroKey = keyof typeof POWER_DISTROS;
type DeploymentType = (typeof DEPLOYMENT_TYPES)[keyof typeof DEPLOYMENT_TYPES];

export type StockRow = {
  code: string;
  name: string;
  required: number;
  stock: number;
  net: number;
  method: string;
  spare?: number;
  rounded?: number;
  /** Live availability for the selected date range, from Rentman - see src/rentman/. */
  available?: number;
};

// A panel in the free workspace. x/y are the TOP-LEFT corner in workspace
// millimetres - panels are no longer bound to a rows x cols grid (the grid
// generator just emits panels on a 500mm pitch). MT is a plain 1000x500mm
// record; the old head/tail module pairing exists only in legacy migration.
export type Cell = {
  /** Stable identity - selection, patching stats and joins key off this. */
  id: string;
  /** Top-left, workspace millimetres. */
  x: number;
  /** Top-left, workspace millimetres. */
  y: number;
  assignedPort: number | null;
  sequence: number | null;
  assignedPowerPort: number | null;
  powerSequence: number | null;
  powerManual: boolean;
  isRemoved: boolean;
  panelVariant: PanelVariantKey;
  rotation: number;
  panelType: PanelTypeKey;
  /** Sub-screen membership - null = unassigned. */
  subScreenId: string | null;
};

// Copy/paste clipboard: each panel's offset from the copied selection's own
// top-left, so pasting anywhere reproduces the exact spacing/arrangement.
// Patching (port assignment) is deliberately not copied - same "paste
// un-patched" convention as importing a layout.
type ClipboardPanel = { dx: number; dy: number; panelType: PanelTypeKey; panelVariant: PanelVariantKey; rotation: number };
type ClipboardSelection = { panels: ClipboardPanel[]; w: number; h: number };

// A named grouping of panels ("Centre Screen", "Stage Left Tower", ...).
// Resolution/physical size is intentionally NOT stored here - it's always
// derived from the sub-screen's member panels, same as the whole-layout
// stats, so there's only ever one source of truth.
export type SubScreen = {
  id: string;
  name: string;
  /** Output-canvas position, top-left origin, canvas pixels. */
  canvasX: number;
  canvasY: number;
  /** Stable insertion-order sort key. */
  createdAt: number;
};

export type OutputCanvasPreset = { w: number; h: number };
export const OUTPUT_CANVAS_PRESETS: OutputCanvasPreset[] = [
  { w: 1920, h: 1080 },
  { w: 2560, h: 1440 },
  { w: 3840, h: 1080 },
  { w: 3840, h: 2160 },
  { w: 4096, h: 2160 },
  { w: 7680, h: 2160 },
  { w: 7680, h: 4320 },
];

type LayoutSnapshot = {
  panels: Cell[];
  subScreens: SubScreen[];
  outputCanvasW: number;
  outputCanvasH: number;
  wholeLayoutCanvasX: number;
  wholeLayoutCanvasY: number;
};

// Hand-off payload from the standalone Quick Panel Layout tab (see
// src/quickLayout/QuickLayoutView.tsx), written to localStorage right before
// it navigates this same tab back to the plain app URL.
const QUICK_LAYOUT_TRANSFER_KEY = "ledCablingQuickLayoutTransfer:v1";
type QuickLayoutTransfer = { panelType: PanelTypeKey; cols: number; rows: number; projectName?: string };

type SignalPortStat = {
  panels: number;
  path: Cell[];
  firstKey: string | null;
  lastKey: string | null;
};

type PowerPortStat = {
  panels: number;
  maxWatts: number;
  maxAmps: number;
  avgWatts: number;
  avgAmps: number;
  utilisation: number;
  phase: string;
  manualPanels: number;
  path: Cell[];
  firstKey: string | null;
  lastKey: string | null;
};

// Legacy (formatVersion 1) grid cell as stored by older saves.
type LegacyGridCell = {
  x: number;
  y: number;
  assignedPort?: number | null;
  sequence?: number | null;
  assignedPowerPort?: number | null;
  powerSequence?: number | null;
  powerManual?: boolean;
  isRemoved?: boolean;
  panelVariant?: PanelVariantKey;
  rotation?: number;
  panelType?: PanelTypeKey;
  mtTail?: boolean;
  id?: string;
  subScreenId?: string | null;
};

type OpenJsonPayload = {
  formatVersion?: number;
  projectName?: string;
  surfaceName?: string;
  panelType?: PanelTypeKey;
  powerDistro?: PowerDistroKey;
  backupSignalLoop?: boolean;
  includeReinforcementPlate?: boolean;
  deploymentType?: DeploymentType | "";
  wall?: {
    cols?: number;
    rows?: number;
  };
  /** v2: flat list of mm-positioned panels. */
  panels?: Cell[];
  /** v1 legacy: rows x cols grid of cells. */
  patching?: {
    grid?: LegacyGridCell[][];
  };
  /** v3: named sub-screen groupings. */
  subScreens?: SubScreen[];
  /** v3: output-canvas resolution. */
  outputCanvas?: { w?: number; h?: number };
  /** v3: whole-layout canvas position, used only when subScreens is empty. */
  wholeLayoutCanvasPos?: { x?: number; y?: number };
  /** v4: selected NovaStar processor model, "" = none selected. */
  processorModel?: ProcessorModelId | "";
  /** v4: per-canvas-entry input assignment, keyed by sub-screen id (or the whole-layout sentinel). */
  canvasInputs?: Record<string, number | null>;
  /** v5: "perEntry" (default) or "whole" - see App's inputMode state. */
  inputMode?: InputMode;
  /** v5: used only when inputMode is "whole". */
  wholeCanvasInputId?: number | null;
  /** v6: Rentman availability date range - see src/rentman/. */
  rentmanDateFrom?: string;
  rentmanDateTo?: string;
};

const gcd = (a: number, b: number): number => {
  const absA = Math.abs(a);
  const absB = Math.abs(b);
  if (absB === 0) return absA;
  return gcd(absB, absA % absB);
};

const makeSignalPorts = (count: number) =>
  Array.from({ length: count }, (_, i) => ({
    id: i + 1,
    name: `Port ${i + 1}`,
    color: PORT_COLORS[i % PORT_COLORS.length],
  }));

const makePowerPorts = (count: number) =>
  Array.from({ length: count }, (_, i) => ({
    id: i + 1,
    name: `Plug ${i + 1}`,
    color: POWER_COLOR,
    phase: `P${(i % 3) + 1}`,
  }));

let cellIdCounter = 0;
const newCellId = () => {
  try {
    return crypto.randomUUID();
  } catch {
    cellIdCounter += 1;
    return `c-${Date.now().toString(36)}-${cellIdCounter}`;
  }
};

const findCellById = (panels: Cell[], id: string | null | undefined): Cell | null => {
  if (!id) return null;
  return panels.find((cell) => cell.id === id) ?? null;
};

const makePanelAt = (xMm: number, yMm: number, panelType: PanelTypeKey = "MG9", subScreenId: string | null = null): Cell => ({
  id: newCellId(),
  x: xMm,
  y: yMm,
  assignedPort: null,
  sequence: null,
  assignedPowerPort: null,
  powerSequence: null,
  powerManual: false,
  isRemoved: false,
  panelVariant: "STANDARD",
  rotation: 0,
  panelType,
  subScreenId,
});

// Grid generator: cols x rows of the given type on its own pitch (MG9 500mm,
// MT 1000mm wide). After generation every panel is freely movable. New panels
// join whichever sub-screen is currently being edited (null = unassigned,
// e.g. Canvas View or no sub-screens created yet).
export const makeGridPanels = (cols: number, rows: number, panelType: PanelTypeKey = "MG9", subScreenId: string | null = null): Cell[] => {
  const wMm = PANEL_TYPES[panelType].w * 1000;
  const hMm = PANEL_TYPES[panelType].h * 1000;
  const panels: Cell[] = [];
  for (let y = 0; y < rows; y += 1) {
    for (let x = 0; x < cols; x += 1) {
      panels.push(makePanelAt(x * wMm, y * hMm, panelType, subScreenId));
    }
  }
  return panels;
};

export const cellPanelType = (cell: Cell): PanelTypeKey => cell.panelType ?? "MG9";

// Spare-panel bucketing: MG9's shaped variants (Triangle/Curved) and its
// Corner variant are each a separate physical stock item from Standard MG9,
// and MT is a separate panel type entirely - so spare stock is computed per
// bucket, not as one combined MG9 number (a single combined ceil() would
// under-count once several buckets each need their own rounding-up). See
// sparePanelSurfaces in the App component for the per-surface breakdown
// this feeds.
export type SpareBucketKey = "MG9_STANDARD" | "MG9_TRIANGLE" | "MG9_CURVED" | "MG9_CORNER" | "MT";
export const SPARE_BUCKETS: Array<{ key: SpareBucketKey; label: string }> = [
  { key: "MG9_STANDARD", label: "MG9 Standard" },
  { key: "MG9_TRIANGLE", label: "MG9 Triangle" },
  { key: "MG9_CURVED", label: "MG9 Curved" },
  { key: "MG9_CORNER", label: "MG9 Corner" },
  { key: "MT", label: "MT" },
];
export const spareBucketOfCell = (cell: Cell): SpareBucketKey => {
  if (cellPanelType(cell) !== "MG9") return "MT";
  const variant = cell.panelVariant ?? "STANDARD";
  if (variant === "TRIANGLE") return "MG9_TRIANGLE";
  if (variant === "CURVED") return "MG9_CURVED";
  if (variant === "CORNER") return "MG9_CORNER";
  return "MG9_STANDARD";
};
const SPARE_BUCKET_RATIO: Record<SpareBucketKey, number> = {
  MG9_STANDARD: PANEL_TYPES.MG9.defaults.spareRatio,
  MG9_TRIANGLE: PANEL_TYPES.MG9.defaults.spareRatio,
  MG9_CURVED: PANEL_TYPES.MG9.defaults.spareRatio,
  MG9_CORNER: PANEL_TYPES.MG9.defaults.spareRatio,
  MT: PANEL_TYPES.MT.defaults.spareRatio,
};
// Box size to round each bucket's (used + spare) up to - null means shaped
// panels (Triangle/Curved), which are one-way physical pieces bought
// individually, not boxed, so the spare is just added as-is, unrounded.
const SPARE_BUCKET_BOX_SIZE: Record<SpareBucketKey, number | null> = {
  MG9_STANDARD: PANEL_TYPES.MG9.defaults.panelsPerBox,
  MG9_TRIANGLE: null,
  MG9_CURVED: null,
  MG9_CORNER: PANEL_TYPES.MG9.defaults.panelsPerBox,
  MT: PANEL_TYPES.MT.defaults.panelsPerBox,
};
export const spareForBucket = (used: number, bucket: SpareBucketKey): { spare: number; rounded: number } => {
  const spare = Math.ceil(used * SPARE_BUCKET_RATIO[bucket]);
  const boxSize = SPARE_BUCKET_BOX_SIZE[bucket];
  const rounded = boxSize ? roundUpToBox(used + spare, boxSize) : used + spare;
  return { spare, rounded };
};

// Footprint in workspace mm, honouring rotation (90/270 swaps width/height).
export const cellSizeMm = (cell: Cell) => {
  const spec = PANEL_TYPES[cellPanelType(cell)];
  const rot = ((Math.round((cell.rotation ?? 0) / 90) * 90) % 360 + 360) % 360;
  const wMm = spec.w * 1000;
  const hMm = spec.h * 1000;
  return rot === 90 || rot === 270 ? { wMm: hMm, hMm: wMm } : { wMm, hMm };
};

export const cellRect = (cell: Cell): RectMm => {
  const { wMm, hMm } = cellSizeMm(cell);
  return { x: cell.x, y: cell.y, w: wMm, h: hMm };
};

// Connector-anchor descriptor for a panel: centre + unrotated base half extents
// + rotation + shape. Drives snapping and join detection (see model/panels.ts).
const cellGeom = (cell: Cell): PanelAnchorSpec => {
  const spec = PANEL_TYPES[cellPanelType(cell)];
  const { wMm, hMm } = cellSizeMm(cell);
  return {
    cx: cell.x + wMm / 2,
    cy: cell.y + hMm / 2,
    halfW: (spec.w * 1000) / 2,
    halfH: (spec.h * 1000) / 2,
    rotation: cell.rotation ?? 0,
    shape: (PANEL_VARIANTS[cell.panelVariant ?? "STANDARD"].shape as PanelShape),
  };
};

// The old grid model called real panels "heads" (vs MT tail modules). In the
// free model every active record is a panel; keep the name for call sites.
export const isPanelHead = (cell: Cell | null | undefined): cell is Cell => isActiveCell(cell);

const cloneGrid = (panels: Cell[]): Cell[] => panels.map((cell) => ({ ...cell }));

// Validate/repair a v2 panel list from a file.
export const normalizePanels = (raw: unknown): Cell[] => {
  if (!Array.isArray(raw)) return [];
  const seen = new Set<string>();
  const panels: Cell[] = [];
  raw.forEach((item) => {
    const cell = item as Partial<Cell> | null;
    if (!cell || !Number.isFinite(Number(cell.x)) || !Number.isFinite(Number(cell.y))) return;
    let id = typeof cell.id === "string" && cell.id ? cell.id : newCellId();
    if (seen.has(id)) id = newCellId();
    seen.add(id);
    panels.push({
      id,
      x: Number(cell.x),
      y: Number(cell.y),
      assignedPort: cell.assignedPort ?? null,
      sequence: cell.sequence ?? null,
      assignedPowerPort: cell.assignedPowerPort ?? null,
      powerSequence: cell.powerSequence ?? null,
      powerManual: Boolean(cell.powerManual),
      isRemoved: Boolean(cell.isRemoved),
      panelVariant: cell.panelVariant && PANEL_VARIANTS[cell.panelVariant] ? cell.panelVariant : "STANDARD",
      rotation: Number.isFinite(cell.rotation) ? ((Number(cell.rotation) % 360) + 360) % 360 : 0,
      panelType: cell.panelType && PANEL_TYPES[cell.panelType] ? cell.panelType : "MG9",
      subScreenId: typeof cell.subScreenId === "string" ? cell.subScreenId : null,
    });
  });
  return panels;
};

// Validate/repair a sub-screen list from a file - drop malformed entries
// rather than letting a corrupt name/id/position crash the app.
export const normalizeSubScreens = (raw: unknown): SubScreen[] => {
  if (!Array.isArray(raw)) return [];
  const seen = new Set<string>();
  const subScreens: SubScreen[] = [];
  raw.forEach((item, index) => {
    const entry = item as Partial<SubScreen> | null;
    if (!entry || typeof entry.id !== "string" || !entry.id || typeof entry.name !== "string") return;
    if (seen.has(entry.id)) return;
    seen.add(entry.id);
    subScreens.push({
      id: entry.id,
      name: entry.name,
      canvasX: Number.isFinite(Number(entry.canvasX)) ? Number(entry.canvasX) : 0,
      canvasY: Number.isFinite(Number(entry.canvasY)) ? Number(entry.canvasY) : 0,
      createdAt: Number.isFinite(Number(entry.createdAt)) ? Number(entry.createdAt) : index,
    });
  });
  return subScreens;
};

// A settings file saved before per-cell panel types existed has cells with no
// `panelType` field. Those all-one-type files stored one panel per grid cell.
const isLegacyUntypedGrid = (rawGrid: unknown): boolean =>
  Array.isArray(rawGrid) &&
  rawGrid.some((row) => Array.isArray(row) && row.some((cell) => cell && (cell as LegacyGridCell).panelType === undefined));

// Migrate a legacy formatVersion-1 grid (rows x cols of cells, MT stored as a
// head module + mtTail module) onto the free mm workspace. Tail modules are
// absorbed into their head, which becomes a single 1000x500mm MT record.
// `legacyAllType` handles pre-panelType files where the wall was one type.
const gridCellsToPanels = (rawGrid: LegacyGridCell[][], legacyAllType: PanelTypeKey | null = null): Cell[] => {
  const panels: Cell[] = [];
  rawGrid.forEach((row, y) => {
    if (!Array.isArray(row)) return;
    row.forEach((cell, x) => {
      if (!cell) return;
      if (cell.mtTail) return; // absorbed into its head
      const cellType: PanelTypeKey =
        legacyAllType ?? (cell.panelType && PANEL_TYPES[cell.panelType] ? cell.panelType : "MG9");
      // Legacy grid columns are 0.5m modules, except pre-panelType MT files
      // where each column was a full 1m MT panel.
      const pitchX = legacyAllType === "MT" ? 1000 : 500;
      panels.push({
        id: typeof cell.id === "string" && cell.id ? cell.id : newCellId(),
        x: (Number(cell.x) || x) * pitchX,
        y: (Number(cell.y) || y) * 500,
        assignedPort: cell.assignedPort ?? null,
        sequence: cell.sequence ?? null,
        assignedPowerPort: cell.assignedPowerPort ?? null,
        powerSequence: cell.powerSequence ?? null,
        powerManual: Boolean(cell.powerManual),
        isRemoved: Boolean(cell.isRemoved),
        panelVariant: cell.panelVariant && PANEL_VARIANTS[cell.panelVariant] ? cell.panelVariant : "STANDARD",
        rotation: Number.isFinite(cell.rotation) ? ((Number(cell.rotation) % 360) + 360) % 360 : 0,
        panelType: cellType,
        subScreenId: null,
      });
    });
  });
  return panels;
};

const isActiveCell = (cell: Cell | null | undefined) => Boolean(cell && !cell.isRemoved);

// Change a panel's type in place (mutates a cloned list). Converting MG9 -> MT
// doubles the footprint: if a standard MG9 sits flush in the newly covered
// space it is absorbed (removed); any other overlap is left to the overlap
// warning. Converting MT -> MG9 halves the footprint and backfills the freed
// half-module with a fresh MG9 so the wall keeps its outline.
const convertPanelTypeInList = (panels: Cell[], id: string, type: PanelTypeKey): Cell[] => {
  const target = panels.find((p) => p.id === id);
  if (!target || target.isRemoved || cellPanelType(target) === type) return panels;
  if (type === "MT") {
    const absorbRect: RectMm = { x: target.x + 500, y: target.y, w: 500, h: 500 };
    const survivors = panels.filter((p) => {
      if (p.id === target.id || p.isRemoved) return true;
      if (cellPanelType(p) !== "MG9" || p.panelVariant !== "STANDARD") return true;
      const r = cellRect(p);
      const flush = Math.abs(r.x - absorbRect.x) < 1 && Math.abs(r.y - absorbRect.y) < 1 && Math.abs(r.w - 500) < 1;
      return !flush;
    });
    target.panelType = "MT";
    target.panelVariant = "STANDARD";
    return survivors;
  }
  // MT -> MG9: shrink in place and backfill the freed right half.
  target.panelType = "MG9";
  const filler = makePanelAt(target.x + 500, target.y, "MG9", target.subScreenId);
  return [...panels, filler];
};

const formatNumber = (value: number, digits = 0) =>
  Number(value || 0).toLocaleString(undefined, {
    maximumFractionDigits: digits,
    minimumFractionDigits: digits,
  });

// Rounds to at most 2 decimal places and trims trailing zeros (19.384... -> "19.38", 3.50 -> "3.5", 4.00 -> "4").
const formatMeters = (value: number) => (Number(value) || 0).toFixed(2).replace(/\.?0+$/, "");

const getStatusColor = (percent: number) => {
  if (percent > 100) return "#ef4444";
  if (percent >= 80) return "#f59e0b";
  return "#22c55e";
};

const clampActivePort = (value: number, max: number) => Math.min(Math.max(value, 1), max);

const clearSignalOnGrid = (panels: Cell[], scopeIds: Set<string> | null = null) =>
  panels.map((cell) =>
    scopeIds && !scopeIds.has(cell.id) ? cell : { ...cell, assignedPort: null, sequence: null },
  );

const clearPowerOnGrid = (panels: Cell[], scopeIds: Set<string> | null = null) =>
  panels.map((cell) =>
    scopeIds && !scopeIds.has(cell.id)
      ? cell
      : { ...cell, assignedPowerPort: null, powerSequence: null, powerManual: false },
  );

const getNextSequence = (
  panels: Cell[],
  portField: "assignedPort" | "assignedPowerPort",
  sequenceField: "sequence" | "powerSequence",
  portId: number,
) => {
  let max = 0;
  for (const cell of panels) {
    if (!isActiveCell(cell)) continue;
    if (cell[portField] === portId && (cell[sequenceField] ?? 0) > max) {
      max = cell[sequenceField] ?? 0;
    }
  }
  return max + 1;
};

const getPowerPortLoadWatts = (
  panels: Cell[],
  portId: number,
  _legacyMaxW: number,
  excludeId: string | null = null,
) => {
  // Each assigned panel draws its own type's max watts (MG9 vs MT differ).
  let watts = 0;
  for (const cell of panels) {
    if (!isActiveCell(cell)) continue;
    if (excludeId && cell.id === excludeId) continue;
    if (cell.assignedPowerPort === portId) watts += PANEL_TYPES[cellPanelType(cell)].power.maxW;
  }
  return watts;
};

const getPortPanelCount = (panels: Cell[], portField: "assignedPort" | "assignedPowerPort", portId: number) =>
  panels.filter((cell) => isActiveCell(cell) && cell[portField] === portId).length;

// Column banding (for TB/BT snake): group active panels into visual columns by
// their horizontal centre, columns left->right and panels top->bottom within.
const bandPanelsByColumn = (panels: Cell[]): Cell[][] => {
  const active = panels.filter((p) => isActiveCell(p));
  const entries = active
    .map((p) => ({ p, r: cellRect(p) }))
    .sort((a, b) => a.r.x + a.r.w / 2 - (b.r.x + b.r.w / 2));
  const bands: { centerX: number; items: { p: Cell; r: RectMm }[] }[] = [];
  entries.forEach((e) => {
    const cx = e.r.x + e.r.w / 2;
    const band = bands.find((b) => Math.abs(b.centerX - cx) < MODULE_MM / 2);
    if (band) {
      band.items.push(e);
      band.centerX = band.items.reduce((s, i) => s + i.r.x + i.r.w / 2, 0) / band.items.length;
    } else {
      bands.push({ centerX: cx, items: [e] });
    }
  });
  return bands.map((b) => b.items.sort((a, c) => a.r.y - c.r.y).map((i) => i.p));
};

// Reading order for auto-snake over a free layout: row bands (or column bands
// for TB/BT) with optional alternation - the non-uniform generalisation of the
// old rows x cols walk. LOOP_TOGETHER pairs row bands into left/right loops.
const orderPanelsForSnake = (panels: Cell[], snakeDirection: string, snakeAlternates = true): Cell[][] => {
  if (snakeDirection === "TB" || snakeDirection === "BT") {
    const columns = bandPanelsByColumn(panels);
    return [
      columns.flatMap((column, index) => {
        let col = [...column];
        if (snakeDirection === "BT") col.reverse();
        if (snakeAlternates && index % 2 === 1) col.reverse();
        return col;
      }),
    ];
  }

  const rows = bandPanels(panels, cellRect) as Cell[][];

  if (snakeDirection === "LOOP_TOGETHER") {
    // Pair adjacent row bands; each pair splits into a left loop and a right
    // loop that both start at the middle, mirroring the old grid behaviour.
    const segments: Cell[][] = [];
    for (let pairStart = 0; pairStart < rows.length; pairStart += 2) {
      const top = rows[pairStart];
      const bottom = pairStart + 1 < rows.length ? rows[pairStart + 1] : null;
      const splitAt = (row: Cell[]) => Math.floor(row.length / 2);
      const topSplit = splitAt(top);
      const leftSegment = [...top.slice(0, topSplit)].reverse();
      const rightSegment = top.slice(topSplit);
      if (bottom) {
        const bottomSplit = splitAt(bottom);
        leftSegment.push(...bottom.slice(0, bottomSplit));
        rightSegment.push(...[...bottom.slice(bottomSplit)].reverse());
      }
      if (leftSegment.length) segments.push(leftSegment);
      if (rightSegment.length) segments.push(rightSegment);
    }
    return segments;
  }

  const startFromBottom = snakeDirection === "LRB" || snakeDirection === "RLB";
  const rightToLeft = snakeDirection === "RL" || snakeDirection === "RLB";
  const orderedRows = startFromBottom ? [...rows].reverse() : rows;
  return [
    orderedRows.flatMap((row, index) => {
      let out = [...row];
      if (rightToLeft) out.reverse();
      if (snakeAlternates && index % 2 === 1) out.reverse();
      return out;
    }),
  ];
};

// Mirror a mm rect horizontally inside the wall bbox (front view).
export const mirrorRectX = (rect: RectMm, bbox: RectMm): RectMm => ({
  ...rect,
  x: 2 * bbox.x + bbox.w - rect.x - rect.w,
});

// Depth-first bottom->top traversal of one letter (a connected group of
// panels). Starts at the bottom-most/left-most panel and, at each step, walks
// to the highest unvisited joined neighbour first; on reaching the top of a
// branch it backtracks to the fork and takes the next branch. This yields the
// "patch up one branch, jump back to the fork, continue" order.
const orderLetterBottomUp = (cells: Cell[]): Cell[] => {
  if (cells.length <= 1) return [...cells];
  const rectOf = new Map(cells.map((c) => [c.id, cellRect(c)]));
  const geomOf = new Map(cells.map((c) => [c.id, cellGeom(c)]));
  const byId = new Map(cells.map((c) => [c.id, c]));
  const adj = new Map<string, string[]>();
  cells.forEach((c) => adj.set(c.id, []));
  for (let i = 0; i < cells.length; i += 1) {
    for (let j = i + 1; j < cells.length; j += 1) {
      if (panelsAnchorJoined(geomOf.get(cells[i].id)!, geomOf.get(cells[j].id)!)) {
        adj.get(cells[i].id)!.push(cells[j].id);
        adj.get(cells[j].id)!.push(cells[i].id);
      }
    }
  }
  const start = [...cells].sort((a, b) => {
    const ra = rectOf.get(a.id)!;
    const rb = rectOf.get(b.id)!;
    return rb.y + rb.h - (ra.y + ra.h) || ra.x - rb.x; // lowest bottom edge, then left-most
  })[0];
  const visited = new Set<string>();
  const order: Cell[] = [];
  const visit = (id: string) => {
    if (visited.has(id)) return;
    visited.add(id);
    order.push(byId.get(id)!);
    const neighbours = adj
      .get(id)!
      .filter((n) => !visited.has(n))
      .sort((a, b) => {
        const ra = rectOf.get(a)!;
        const rb = rectOf.get(b)!;
        return ra.y - rb.y || ra.x - rb.x; // go up (smaller y) first
      });
    neighbours.forEach(visit);
  };
  visit(start.id);
  cells.forEach((c) => {
    if (!visited.has(c.id)) order.push(c);
  });
  return order;
};

// Order panels for letter-shaped layouts: split active panels into connected
// "letters", order the letters left->right within top->bottom text lines, and
// return each letter as a bottom-up traversal. snakePatch keeps a whole letter
// on one port where it fits.
const orderPanelsForLetters = (panels: Cell[]): Cell[][] => {
  const active = panels.filter((c) => isActiveCell(c));
  if (!active.length) return [];
  const groups = connectedGroupsByGeom(active, cellGeom);
  const byGroup = new Map<number, Cell[]>();
  active.forEach((c) => {
    const g = groups.get(c.id);
    if (g === undefined) return;
    const arr = byGroup.get(g) ?? [];
    arr.push(c);
    byGroup.set(g, arr);
  });
  const letters = [...byGroup.values()].map((cells) => {
    const bb = activeBBox(cells.map(cellRect));
    return { cells, bb, cx: bb.x + bb.w / 2, cy: bb.y + bb.h / 2 };
  });
  // Band letters into text lines by vertical centre, lines top->bottom.
  letters.sort((a, b) => a.cy - b.cy);
  const lines: (typeof letters)[] = [];
  letters.forEach((letter) => {
    const line = lines.find((l) => Math.abs(l[0].cy - letter.cy) < letter.bb.h * 0.6 + 250);
    if (line) line.push(letter);
    else lines.push([letter]);
  });
  const ordered: Cell[][] = [];
  lines.forEach((line) => {
    line.sort((a, b) => a.bb.x - b.bb.x); // left -> right
    line.forEach((letter) => ordered.push(orderLetterBottomUp(letter.cells)));
  });
  return ordered;
};

// `required` is always the raw quantity needed to build the wall, with no
// spare or packaging rounding folded in - `rounded` (required + spare,
// packaging-rounded where relevant) is the real order/pull quantity, so
// `net` (shortfall) is checked against THAT, not the bare required count.
const makeStockRow = (
  item: { code: string; name: string; stock: number },
  required: number,
  method: string,
  spare = 0,
  rounded = required + spare,
): StockRow => ({
  code: item.code,
  name: item.name,
  required,
  stock: item.stock,
  net: item.stock - rounded,
  method,
  spare,
  rounded,
});

const roundUpToBox = (value: number, boxSize = 10) => Math.ceil(Math.max(value, 0) / boxSize) * boxSize;

const getSelectedIds = (selectedCells: Set<string>, selectedId: string | null) => {
  if (selectedCells.size > 0) return selectedCells;
  return selectedId ? new Set([selectedId]) : new Set<string>();
};

// SVG outline path (in a 0..100 box) matching each variant's on-screen shape,
// used to draw the signal/power indicator outlines so they follow the panel shape.
const variantOutlineSvgPath = (shape: string): string => {
  if (shape === "triangle") return "M0 0 L100 100 L0 100 Z";
  if (shape === "curve") return "M0 100 A100 100 0 0 1 100 0 L100 100 Z"; // matches circle(farthest-side at 100% 100%)
  return "M0 0 H100 V100 H0 Z";
};

const getPanelSymbol = (cell: Cell) => {
  const variant = PANEL_VARIANTS[cell.panelVariant ?? "STANDARD"];
  const parts = [];
  if (variant.symbol) parts.push(variant.symbol);
  if (cell.rotation) parts.push("🔄");
  return parts.join(" ");
};

// Cabling endpoints between two panel rects (px space). Side-by-side panels
// connect edge to edge at the middle of their vertical overlap; stacked panels
// connect at the middle of their horizontal overlap; anything else runs
// centre to centre.
const getLineEndpointsPx = (a: RectMm, b: RectMm, offsetY = 0) => {
  const vOverlapLo = Math.max(a.y, b.y);
  const vOverlapHi = Math.min(a.y + a.h, b.y + b.h);
  const hOverlapLo = Math.max(a.x, b.x);
  const hOverlapHi = Math.min(a.x + a.w, b.x + b.w);

  if (vOverlapHi - vOverlapLo > 4) {
    const y = (vOverlapLo + vOverlapHi) / 2 + offsetY;
    if (b.x >= a.x + a.w - 1) return { x1: a.x + a.w - 1, y1: y, x2: b.x + 1, y2: y };
    if (a.x >= b.x + b.w - 1) return { x1: a.x + 1, y1: y, x2: b.x + b.w - 1, y2: y };
  }
  if (hOverlapHi - hOverlapLo > 4) {
    const x = (hOverlapLo + hOverlapHi) / 2 + offsetY; // offset separates signal/power runs
    if (b.y >= a.y + a.h - 1) return { x1: x, y1: a.y + a.h - 1, x2: x, y2: b.y + 1 };
    if (a.y >= b.y + b.h - 1) return { x1: x, y1: a.y + 1, x2: x, y2: b.y + b.h - 1 };
  }
  return {
    x1: a.x + a.w / 2 + offsetY,
    y1: a.y + a.h / 2 + offsetY,
    x2: b.x + b.w / 2 + offsetY,
    y2: b.y + b.h / 2 + offsetY,
  };
};

// Do two panel rects share an edge (touching, with real overlap)? Used to
// decide when a cable may run straight between panels vs route around.
const rectsAdjacentPx = (a: RectMm, b: RectMm, tol = 3) => {
  const vOverlap = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
  const hOverlap = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
  const touchV = Math.abs(a.x + a.w - b.x) <= tol || Math.abs(b.x + b.w - a.x) <= tol;
  const touchH = Math.abs(a.y + a.h - b.y) <= tol || Math.abs(b.y + b.h - a.y) <= tol;
  return (touchV && vOverlap > tol) || (touchH && hOverlap > tol);
};

// Orthogonal (Manhattan) cable route between two panel rects in px space.
// Returns a polyline with only horizontal/vertical segments and 90-degree
// corners. Adjacent panels connect straight through their shared edge; other
// panels leave the facing edge and turn in the gap between them (never
// through a panel centre). `offset` shifts the run so signal/power don't overlap.
const routeCablePx = (a: RectMm, b: RectMm, offset = 0): Array<{ x: number; y: number }> => {
  const aC = { x: a.x + a.w / 2, y: a.y + a.h / 2 };
  const bC = { x: b.x + b.w / 2, y: b.y + b.h / 2 };
  const horizontal = Math.abs(bC.x - aC.x) >= Math.abs(bC.y - aC.y);
  if (horizontal) {
    const rightward = bC.x >= aC.x;
    const ax = rightward ? a.x + a.w : a.x;
    const bx = rightward ? b.x : b.x + b.w;
    const ay = aC.y + offset;
    const by = bC.y + offset;
    if (Math.abs(ay - by) < 1) return [{ x: ax, y: ay }, { x: bx, y: by }];
    const midX = (ax + bx) / 2; // in the horizontal gap between the facing edges
    return [{ x: ax, y: ay }, { x: midX, y: ay }, { x: midX, y: by }, { x: bx, y: by }];
  }
  const downward = bC.y >= aC.y;
  const ay = downward ? a.y + a.h : a.y;
  const by = downward ? b.y : b.y + b.h;
  const ax = aC.x + offset;
  const bx = bC.x + offset;
  if (Math.abs(ax - bx) < 1) return [{ x: ax, y: ay }, { x: bx, y: by }];
  const midY = (ay + by) / 2;
  return [{ x: ax, y: ay }, { x: ax, y: midY }, { x: bx, y: midY }, { x: bx, y: by }];
};

// Trace a panel's true outline (triangle / quarter-circle / rectangle) in the
// local 0..w,0..h space, so fills, strokes, indicator rings and any other
// consumer (on-screen, PDF, PNG, animated test pattern) all share ONE
// implementation and therefore agree on where a panel's rotation actually
// puts its cut corner. (The PNG exporter used to trace a curve with its
// right-angle corner at the opposite corner from every other renderer via a
// since-removed `testPattern` path variant - that was the source of curved
// panels appearing incorrectly rotated in the PNG relative to the PDF.)
export const tracePanelShapePath = (ctx: CanvasRenderingContext2D, w: number, h: number, shape: PanelShape) => {
  ctx.beginPath();
  if (shape === "triangle") {
    ctx.moveTo(0, 0);
    ctx.lineTo(w, h);
    ctx.lineTo(0, h);
    ctx.closePath();
  } else if (shape === "curve") {
    ctx.moveTo(w, 0);
    ctx.lineTo(w, h);
    ctx.lineTo(0, h);
    ctx.quadraticCurveTo(0, 0, w, 0);
    ctx.closePath();
  } else {
    ctx.rect(0, 0, w, h);
  }
};

// Establish a panel's local frame at (x,y,w,h): front view mirrors the shape
// via scaleX(-1) without touching its stored rotation. Shared by drawPanelShape
// and any other consumer that needs to clip/draw in a panel's true footprint.
export const applyPanelFrame = (ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number, rotation: number, mirrorX = false) => {
  ctx.translate(x + w / 2, y + h / 2);
  if (mirrorX) ctx.scale(-1, 1);
  ctx.rotate((rotation * Math.PI) / 180);
  ctx.translate(-w / 2, -h / 2);
};

const drawPanelShape = (
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  cell: Cell,
  fill: string,
  stroke: string,
  lineWidth = 2,
  options: { hatchStep?: number; signalBadges?: number[]; powerBadge?: number | null; mirrorX?: boolean } = {},
) => {
  const variant = PANEL_VARIANTS[cell.panelVariant ?? "STANDARD"];
  const traceShape = () => tracePanelShapePath(ctx, w, h, variant.shape);
  const applyFrame = () => applyPanelFrame(ctx, x, y, w, h, cell.rotation ?? 0, options.mirrorX);

  ctx.save();
  applyFrame();
  ctx.fillStyle = fill;
  ctx.strokeStyle = stroke;
  ctx.lineWidth = lineWidth;
  traceShape();
  ctx.fill();
  ctx.stroke();

  if (variant.shape === "corner") {
    ctx.strokeStyle = "rgba(2, 6, 23, 0.45)";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.rect(0, 0, w, h);
    ctx.clip();
    for (let i = -h; i < w + h; i += options.hatchStep ?? 10) {
      ctx.beginPath();
      ctx.moveTo(i, h);
      ctx.lineTo(i + h, 0);
      ctx.stroke();
    }
  }
  ctx.restore();

  // Chain-start indicator rings that FOLLOW the panel shape - shown together
  // with the port-number badges below (both requested). Blue = signal chain
  // start (and backup-loop end, when the backup badge below applies), orange
  // = power chain start. Clipping to the shape and stroking the outline
  // gives a constant-thickness outline that hugs the true edge; when both
  // apply the wider (power) band is drawn first and the signal band sits on
  // top, so they nest concentrically and stay distinct.
  const hasSignalRing = (options.signalBadges ?? []).length > 0;
  const hasPowerRing = !!options.powerBadge;
  if (hasSignalRing || hasPowerRing) {
    const ringW = Math.max(2, Math.round(Math.min(w, h) * 0.06));
    ctx.save();
    applyFrame();
    traceShape();
    ctx.clip();
    if (hasPowerRing) {
      ctx.strokeStyle = POWER_START_COLOR;
      ctx.lineWidth = ringW * 2 * (hasSignalRing ? 2 : 1);
      traceShape();
      ctx.stroke();
    }
    if (hasSignalRing) {
      ctx.strokeStyle = SIGNAL_START_COLOR;
      ctx.lineWidth = ringW * 2;
      traceShape();
      ctx.stroke();
    }
    ctx.restore();
  }

  // Port-number badges: small filled circles with the port number, all in
  // the panel's own top-left corner, side by side (signal first, then
  // power) - a single neatly-spaced, non-overlapping row. Deliberately
  // drawn in absolute (x,y,w,h) space, OUTSIDE the panel's rotate/mirror
  // frame (applyFrame, used above only for the fill/outline) - the digit
  // always stays upright and legible even on a rotated panel, and
  // "top-left" always means the panel's own unrotated footprint corner. A
  // chain's first panel gets its primary signal port number; when the
  // backup signal loop is enabled, the chain's last panel also gets a
  // second signal badge with the backup port number (see
  // getPanelIndicators) - if a chain is a single panel, both land on it.
  const signalBadges = options.signalBadges ?? [];
  const cornerBadges: Array<{ color: string; text: string }> = signalBadges.map((portNum) => ({ color: SIGNAL_START_COLOR, text: String(portNum) }));
  if (options.powerBadge) cornerBadges.push({ color: POWER_START_COLOR, text: String(options.powerBadge) });
  if (cornerBadges.length) {
    const badgeR = Math.max(6, Math.round(Math.min(w, h) * 0.15));
    const pad = Math.max(2, Math.round(badgeR * 0.35));
    const fontPx = Math.round(badgeR * 1.15);
    const drawBadge = (cx: number, cy: number, color: string, text: string) => {
      ctx.save();
      ctx.beginPath();
      ctx.arc(cx, cy, badgeR, 0, Math.PI * 2);
      ctx.fillStyle = color;
      ctx.fill();
      ctx.strokeStyle = "#0f172a";
      ctx.lineWidth = 1;
      ctx.stroke();
      ctx.fillStyle = "#ffffff";
      ctx.font = `bold ${fontPx}px Arial`;
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(text, cx, cy + 0.5);
      ctx.restore();
    };
    cornerBadges.forEach((b, i) => {
      drawBadge(x + pad + badgeR + i * (badgeR * 2 + pad), y + pad + badgeR, b.color, b.text);
    });
  }
};

const drawCanvasArrowHead = (
  ctx: CanvasRenderingContext2D,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  color: string,
  size = 16,
) => {
  const angle = Math.atan2(y2 - y1, x2 - x1);
  const baseX1 = x2 - size * Math.cos(angle - Math.PI / 6);
  const baseY1 = y2 - size * Math.sin(angle - Math.PI / 6);
  const baseX2 = x2 - size * Math.cos(angle + Math.PI / 6);
  const baseY2 = y2 - size * Math.sin(angle + Math.PI / 6);
  ctx.save();
  ctx.strokeStyle = "#020617";
  ctx.lineWidth = 2;
  ctx.fillStyle = color;
  ctx.beginPath();
  ctx.moveTo(x2, y2);
  ctx.lineTo(baseX1, baseY1);
  ctx.lineTo(baseX2, baseY2);
  ctx.closePath();
  ctx.stroke();
  ctx.fill();
  ctx.restore();
};

function UtilBar({ percent }: { percent: number }) {
  const color = getStatusColor(percent);
  return (
    <div className="h-2 w-full rounded border border-white/30 bg-black/30">
      <div className="h-2 rounded" style={{ width: `${Math.min(percent, 100)}%`, background: color }} />
    </div>
  );
}

function HelpModal({ onClose }: { onClose: () => void }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 no-print" onMouseDown={onClose}>
      <div className="max-w-2xl rounded-xl border border-slate-600 bg-slate-900 p-5 text-white shadow-2xl" onMouseDown={(event) => event.stopPropagation()}>
        <div className="mb-3 flex items-center justify-between gap-4">
          <div className="text-lg font-bold">LED Planner Help</div>
          <Button variant="outline" onClick={onClose}>Close</Button>
        </div>
        <div className="grid gap-4 text-sm md:grid-cols-2">
          <div className="space-y-2">
            <div className="font-semibold text-sky-200">Workflow</div>
            <div><b>Patch</b>: click or drag panels to patch the selected signal port or power plug.</div>
            <div><b>Select</b>: click a panel or drag a box (Shift adds). Then change type, rotate, clear, delete, or restore.</div>
            <div><b>Move</b>: drag panels to reposition freely; edges snap and join. Toggle Snap for fine positioning.</div>
            <div><b>Import Project</b>: bring in a layout from the YES TECH Layout Tool.</div>
          </div>
          <div className="space-y-2">
            <div className="font-semibold text-sky-200">Shortcuts</div>
            <div><b>S</b>: Select mode</div>
            <div><b>M</b>: Move mode</div>
            <div><b>P</b>: Patch mode</div>
            <div><b>R</b>: Rotate selected panels</div>
            <div><b>C</b>: Clear selected panel patching</div>
            <div><b>Delete</b>: Delete selected panels (Remove / Mark Inactive)</div>
            <div><b>Ctrl+Z</b>: Undo</div>
            <div><b>Ctrl+Y</b> or <b>Ctrl+Shift+Z</b>: Redo</div>
            <div><b>Escape</b>: Clear selection or leave the current mode</div>
          </div>
        </div>
      </div>
    </div>
  );
}

function DeleteConfirmModal({
  count,
  onRemove,
  onMarkInactive,
  onCancel,
}: {
  count: number;
  onRemove: () => void;
  onMarkInactive: () => void;
  onCancel: () => void;
}) {
  const label = count === 1 ? "this panel" : `these ${count} panels`;
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 no-print" onMouseDown={onCancel}>
      <div className="max-w-md rounded-xl border border-slate-600 bg-slate-900 p-5 text-white shadow-2xl" onMouseDown={(event) => event.stopPropagation()}>
        <div className="mb-2 text-lg font-bold">Delete {label}?</div>
        <div className="mb-4 text-sm text-slate-300">
          Choose how to handle {label}. Inactive panels stay in position but are excluded from totals, patching and exported outputs.
        </div>
        <div className="flex flex-col gap-2">
          <Button intent="danger" onClick={onRemove}>Remove Panel{count === 1 ? "" : "s"}</Button>
          <Button intent="secondary" onClick={onMarkInactive}>Mark as Inactive</Button>
          <Button intent="ghost" onClick={onCancel}>Cancel</Button>
        </div>
      </div>
    </div>
  );
}

function GridSizeConfirmModal({ onConfirm, onCancel }: { onConfirm: () => void; onCancel: () => void }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 no-print" onMouseDown={onCancel}>
      <div className="max-w-md rounded-xl border border-slate-600 bg-slate-900 p-5 text-white shadow-2xl" onMouseDown={(event) => event.stopPropagation()}>
        <div className="mb-2 text-lg font-bold">Apply new grid size?</div>
        <div className="mb-4 text-sm text-slate-300">
          Applying this grid size will remove all panels currently in the layout. Do you want to continue?
        </div>
        <div className="flex flex-col gap-2">
          <Button intent="danger" onClick={onConfirm}>Remove Panels and Apply Grid</Button>
          <Button intent="ghost" onClick={onCancel}>Cancel</Button>
        </div>
      </div>
    </div>
  );
}

function DownloadFormatModal({
  format,
  onFormatChange,
  onDownload,
  onCancel,
}: {
  format: "webm" | "mp4";
  onFormatChange: (format: "webm" | "mp4") => void;
  onDownload: () => void;
  onCancel: () => void;
}) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 no-print" onMouseDown={onCancel}>
      <div className="w-full max-w-md rounded-xl border border-slate-600 bg-slate-900 p-5 text-white shadow-2xl" onMouseDown={(event) => event.stopPropagation()}>
        <div className="mb-2 text-lg font-bold">Download Moving Test Pattern</div>
        <div className="mb-3 text-sm text-slate-300">Choose an output format for the recorded video.</div>
        <div className="mb-3 space-y-2">
          <label className="flex items-center gap-2 rounded-lg border border-slate-700 bg-slate-800 p-2 text-sm">
            <input type="radio" name="download-format" checked={format === "webm"} onChange={() => onFormatChange("webm")} />
            <span>
              <span className="font-semibold">WebM</span>
              <span className="text-slate-400"> - fast, recorded directly in the browser</span>
            </span>
          </label>
          <label className="flex items-center gap-2 rounded-lg border border-slate-700 bg-slate-800 p-2 text-sm">
            <input type="radio" name="download-format" checked={format === "mp4"} onChange={() => onFormatChange("mp4")} />
            <span>
              <span className="font-semibold">MP4</span>
              <span className="text-slate-400"> - widely compatible, encoded in the browser after recording</span>
            </span>
          </label>
        </div>
        {format === "mp4" ? (
          <div className="mb-3 rounded-lg border border-amber-400 bg-amber-500/15 p-2 text-xs text-amber-200">
            ⚠ MP4 requires an extra encoding pass after recording and can take significantly longer than WebM, especially for larger walls.
          </div>
        ) : null}
        <div className="flex justify-end gap-2">
          <Button intent="ghost" onClick={onCancel}>Cancel</Button>
          <Button intent="primary" onClick={onDownload}>Download</Button>
        </div>
      </div>
    </div>
  );
}

function ImportPreviewModal({
  result,
  hasUnsavedWork,
  onCancel,
  onApply,
}: {
  result: ImportResult;
  hasUnsavedWork: boolean;
  onCancel: () => void;
  onApply: (result: ImportResult, mode: "replace" | "new") => void;
}) {
  const typeLabel: Record<string, string> = { MG9: "MG9 square", MG12: "MG12 triangle", MG13: "MG13 quarter-circle" };
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 no-print" onMouseDown={onCancel}>
      <div className="max-h-[85vh] w-full max-w-xl overflow-auto rounded-xl border border-slate-600 bg-slate-900 p-5 text-white shadow-2xl" onMouseDown={(event) => event.stopPropagation()}>
        <div className="mb-3 flex items-center justify-between gap-4">
          <div className="text-lg font-bold">Import Project</div>
          <Button variant="outline" onClick={onCancel}>Close</Button>
        </div>

        {result.ok ? (
          <>
            <div className="rounded-lg border border-slate-700 bg-slate-800 p-3 text-sm">
              <div className="mb-2 font-semibold text-sky-200">Detected</div>
              <dl className="grid grid-cols-2 gap-x-4 gap-y-1">
                <dt className="text-slate-400">Project name</dt>
                <dd>{result.projectName}</dd>
                <dt className="text-slate-400">Source version</dt>
                <dd>{result.summary.sourceVersion ?? "unknown"}</dd>
                <dt className="text-slate-400">Active panels</dt>
                <dd>{result.summary.panelCount}</dd>
                <dt className="text-slate-400">Panel types</dt>
                <dd>{Object.entries(result.summary.typeCounts).map(([t, n]) => `${n}× ${typeLabel[t] ?? t}`).join(", ")}</dd>
                <dt className="text-slate-400">Wall size</dt>
                <dd>{result.summary.widthM.toFixed(2)}m × {result.summary.heightM.toFixed(2)}m</dd>
                <dt className="text-slate-400">Signal / power</dt>
                <dd>0 outputs (imported un-patched)</dd>
                <dt className="text-slate-400">Backup loop</dt>
                <dd>Unchanged</dd>
              </dl>
            </div>

            {result.converted.length ? (
              <div className="mt-3 rounded-lg border border-sky-800 bg-sky-950/40 p-3 text-xs text-sky-200">
                <div className="mb-1 font-semibold">Converted</div>
                <ul className="list-disc space-y-0.5 pl-4">{result.converted.map((c, i) => <li key={i}>{c}</li>)}</ul>
              </div>
            ) : null}
            {result.warnings.length ? (
              <div className="mt-3 rounded-lg border border-amber-700 bg-amber-950/40 p-3 text-xs text-amber-200">
                <div className="mb-1 font-semibold">Notes</div>
                <ul className="list-disc space-y-0.5 pl-4">{result.warnings.map((c, i) => <li key={i}>{c}</li>)}</ul>
              </div>
            ) : null}
            {result.skipped.length ? (
              <div className="mt-3 rounded-lg border border-rose-800 bg-rose-950/40 p-3 text-xs text-rose-200">
                <div className="mb-1 font-semibold">Skipped ({result.skipped.length})</div>
                <ul className="list-disc space-y-0.5 pl-4">{result.skipped.slice(0, 8).map((c, i) => <li key={i}>{c}</li>)}</ul>
                {result.skipped.length > 8 ? <div className="pl-4">…and {result.skipped.length - 8} more.</div> : null}
              </div>
            ) : null}

            {hasUnsavedWork ? (
              <div className="mt-3 rounded-lg border border-amber-500 bg-amber-500/15 p-2 text-xs text-amber-200">
                ⚠ Your current project has patching that will be replaced. Save it first if you want to keep it.
              </div>
            ) : null}

            <div className="mt-4 flex flex-wrap justify-end gap-2">
              <Button variant="outline" onClick={onCancel}>Cancel</Button>
              <Button intent="secondary" onClick={() => onApply(result, "replace")}>Replace current</Button>
              <Button intent="primary" onClick={() => onApply(result, "new")}>Import as new project</Button>
            </div>
          </>
        ) : (
          <div className="rounded-lg border border-rose-700 bg-rose-950/40 p-3 text-sm text-rose-200">
            <div className="mb-1 font-semibold">Could not import this file</div>
            <div>{result.error}</div>
            {result.skipped.length ? (
              <ul className="mt-2 list-disc space-y-0.5 pl-4 text-xs">{result.skipped.slice(0, 8).map((c, i) => <li key={i}>{c}</li>)}</ul>
            ) : null}
            <div className="mt-4 flex justify-end">
              <Button variant="outline" onClick={onCancel}>Close</Button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function QuickLayoutTransferModal({
  payload,
  onCancel,
  onReplace,
  onAdd,
}: {
  payload: QuickLayoutTransfer;
  onCancel: () => void;
  onReplace: () => void;
  onAdd: () => void;
}) {
  const typeLabel = payload.panelType === "MT" ? "MT" : "MG9";
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 no-print" onMouseDown={onCancel}>
      <div className="w-full max-w-md rounded-xl border border-slate-600 bg-slate-900 p-5 text-white shadow-2xl" onMouseDown={(event) => event.stopPropagation()}>
        <div className="mb-3 text-lg font-bold">Quick Panel Layout</div>
        <p className="text-sm text-slate-300">
          Bring in {payload.cols}×{payload.rows} {typeLabel} panels from Quick Panel Layout. This project already has panels on it - replace the current layout, or add the new grid alongside it?
        </p>
        <div className="mt-4 flex flex-wrap justify-end gap-2">
          <Button variant="outline" onClick={onCancel}>Cancel</Button>
          <Button intent="secondary" onClick={onAdd}>Add to canvas</Button>
          <Button intent="primary" onClick={onReplace}>Replace current</Button>
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const importInputRef = useRef<HTMLInputElement | null>(null);
  const workspaceRef = useRef<HTMLDivElement | null>(null);
  // The scrollable viewport AROUND workspaceRef (the overflow-auto wrapper) -
  // its current visible size is what "Fit to View" measures against.
  const workspaceViewportRef = useRef<HTMLDivElement | null>(null);
  const [importPreview, setImportPreview] = useState<ImportResult | null>(null);
  const [pendingQuickLayoutTransfer, setPendingQuickLayoutTransfer] = useState<QuickLayoutTransfer | null>(null);

  const [projectName, setProjectName] = useState("Untitled Project");
  const [surfaceName, setSurfaceName] = useState("");
  const [panelType, setPanelType] = useState<PanelTypeKey>("MG9");
  const [includeFlyBar, setIncludeFlyBar] = useState(false);
  const [includeSling, setIncludeSling] = useState(false);
  const [includePowerCable, setIncludePowerCable] = useState(false);
  const [includeSignalCable, setIncludeSignalCable] = useState(false);
  const [includeCustomWeight, setIncludeCustomWeight] = useState(false);
  const [customWeight, setCustomWeight] = useState(0);
  const [cols, setCols] = useState(24);
  const [rows, setRows] = useState(8);
  const [draftCols, setDraftCols] = useState("24");
  const [draftRows, setDraftRows] = useState("8");
  const [grid, setGrid] = useState<Cell[]>(() => []);
  const [activePort, setActivePort] = useState(1);
  const [activePowerPort, setActivePowerPort] = useState(1);
  const [patchMode, setPatchMode] = useState<"signal" | "power">("signal");
  const [powerDistro, setPowerDistro] = useState<PowerDistroKey>("32A");
  const [isDragging, setIsDragging] = useState(false);
  const [dragVisited, setDragVisited] = useState<Set<string>>(() => new Set());
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [selectedCells, setSelectedCells] = useState<Set<string>>(() => new Set());
  // Workspace editor mode: patch (default click-to-patch), select (click/marquee
  // selection), move (free drag repositioning).
  const [editMode, setEditMode] = useState<"patch" | "select" | "move">("patch");
  const [isSelectingPanels, setIsSelectingPanels] = useState(false);
  // Marquee corners in workspace mm while select-dragging.
  const [selectionStart, setSelectionStart] = useState<{ x: number; y: number } | null>(null);
  const [selectionEnd, setSelectionEnd] = useState<{ x: number; y: number } | null>(null);
  // Free-move gesture: which panels are moving and the live mm delta.
  const [moveDrag, setMoveDrag] = useState<{ ids: string[]; startX: number; startY: number; dx: number; dy: number } | null>(null);
  // Live snap/join preview shown while dragging: display-px outlines of where the
  // moving panels will land, plus the shared edges they will join along.
  const [snapGuide, setSnapGuide] = useState<{ ghosts: RectMm[]; edges: { x1: number; y1: number; x2: number; y2: number }[] } | null>(null);
  const [snapEnabled, setSnapEnabled] = useState(true);
  const [allowOverlaps, setAllowOverlaps] = useState(false);
  const [moveJoinedGroup, setMoveJoinedGroup] = useState(false);
  const [zoom, setZoom] = useState(1);
  const [overlapNotice, setOverlapNotice] = useState<string | null>(null);
  const [undoStack, setUndoStack] = useState<LayoutSnapshot[]>([]);
  const [redoStack, setRedoStack] = useState<LayoutSnapshot[]>([]);
  const [showHelp, setShowHelp] = useState(false);
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  const [showGridSizeConfirm, setShowGridSizeConfirm] = useState(false);
  const [customRotationDeg, setCustomRotationDeg] = useState("15");
  const [showCentreLine, setShowCentreLine] = useState(true);
  const [includeCentreLineInExport, setIncludeCentreLineInExport] = useState(true);
  const [clipboard, setClipboard] = useState<ClipboardSelection | null>(null);
  const [isPasting, setIsPasting] = useState(false);
  const [pasteAnchor, setPasteAnchor] = useState<{ x: number; y: number } | null>(null);
  const [isRecordingVideo, setIsRecordingVideo] = useState(false);
  const [videoRecordSeconds, setVideoRecordSeconds] = useState(0);
  const [isEncodingMp4, setIsEncodingMp4] = useState(false);
  const [mp4EncodeProgress, setMp4EncodeProgress] = useState(0);
  const [showDownloadFormatModal, setShowDownloadFormatModal] = useState(false);
  const [downloadFormat, setDownloadFormat] = useState<"webm" | "mp4">("webm");
  const [snakeDirection, setSnakeDirection] = useState<"LR" | "RL" | "LRB" | "RLB" | "TB" | "BT" | "LOOP_TOGETHER" | "LETTERS">("LR");
  const [snakeAlternates, setSnakeAlternates] = useState(true);
  const [isFlippedView, setIsFlippedView] = useState(false);
  const [backupSignalLoop, setBackupSignalLoop] = useState(true);
  const [includeReinforcementPlate, setIncludeReinforcementPlate] = useState(false);
  const [deploymentType, setDeploymentType] = useState<DeploymentType | "">("");

  // --- Sub-screens + output-canvas positioning -----------------------------
  const [subScreens, setSubScreens] = useState<SubScreen[]>([]);
  // null = "Canvas View" (whole layout). Use resolvedActiveSubScreenId below
  // for any read - it treats a dangling id (e.g. after undoing a sub-screen
  // creation) as Canvas View instead of crashing/misbehaving.
  const [activeSubScreenId, setActiveSubScreenId] = useState<string | null>(null);
  const [outputCanvasW, setOutputCanvasW] = useState(1920);
  const [outputCanvasH, setOutputCanvasH] = useState(1080);
  // Canvas position of the whole layout, used only when no sub-screens exist.
  const [wholeLayoutCanvasX, setWholeLayoutCanvasX] = useState(0);
  const [wholeLayoutCanvasY, setWholeLayoutCanvasY] = useState(0);
  const [canvasSnapEnabled, setCanvasSnapEnabled] = useState(true);
  // --- NovaStar processor configuration export -----------------------------
  // Defaults to VX2000 Pro for a new project. "" (no processor selected) is
  // still a valid state - openJson always explicitly sets this from the
  // loaded file (falling back to "" when absent/invalid), so this default
  // only affects a brand-new project, never the format-migration behaviour
  // for older saved files.
  const [processorModel, setProcessorModel] = useState<ProcessorModelId | "">("VX2000_PRO");
  // "perEntry" (default, original behavior): a separate input per sub-screen
  // (or a single whole-layout entry). "whole": one input for the entire
  // output canvas regardless of sub-screen boundaries.
  const [inputMode, setInputMode] = useState<InputMode>("perEntry");
  // Per-canvas-entry (sub-screen, or WHOLE_LAYOUT_KEY when none exist) input
  // assignment - the FK into the selected processor's input list. Kept even
  // while inputMode is "whole" so switching back to "perEntry" restores it.
  const [canvasInputs, setCanvasInputs] = useState<Record<string, number | null>>({});
  // Used only when inputMode === "whole".
  const [wholeCanvasInputId, setWholeCanvasInputId] = useState<number | null>(null);
  const [isGeneratingNovaStarFile, setIsGeneratingNovaStarFile] = useState(false);
  // Rentman Integration (see src/rentman/) - project-specific date range is
  // regular state (saved/loaded with the project); the code -> Rentman
  // equipment mapping is account-wide, so it's loaded from/persisted to
  // localStorage instead (see equipmentMapping.ts), not this project's JSON.
  const [rentmanDateFrom, setRentmanDateFrom] = useState("");
  const [rentmanDateTo, setRentmanDateTo] = useState("");
  const [equipmentMapping, setEquipmentMapping] = useState<EquipmentMapping>(() => loadEquipmentMapping());
  const [liveStock, setLiveStock] = useState<Record<string, LiveStockEntry>>({});
  const [liveAvailable, setLiveAvailable] = useState<Record<string, number>>({});
  const [rentmanRefreshing, setRentmanRefreshing] = useState(false);
  const [rentmanRefreshError, setRentmanRefreshError] = useState<string | null>(null);
  const [rentmanLastRefreshedAt, setRentmanLastRefreshedAt] = useState<Date | null>(null);
  // Drag gesture for repositioning a sub-screen (or the whole layout, id=null)
  // on the output canvas - separate pixel-space analogue of moveDrag.
  const [canvasDrag, setCanvasDrag] = useState<{ id: string | null; startX: number; startY: number; dx: number; dy: number } | null>(null);
  // Which sub-screen the "Assign Selected" control (next to Undo/Redo) will
  // assign the current selection to.
  const [assignTargetSubScreenId, setAssignTargetSubScreenId] = useState("");

  const panel = PANEL_TYPES[panelType];
  // Workspace scale: CELL_SIZE px per 0.5m module at zoom 1.
  const pxPerMm = (CELL_SIZE / MODULE_MM) * zoom;
  const panelSelectMode = editMode === "select";
  const powerSpec = panel.power;
  const distro = POWER_DISTROS[powerDistro];
  const powerPorts = useMemo(() => makePowerPorts(distro.portCount), [distro.portCount]);
  // Total selectable signal ports matches the selected NovaStar processor's
  // Ethernet output count (10 for VX1000 Pro, 20 for VX2000 Pro) so the
  // Signal Patching UI never offers more ports than the processor actually
  // has; falls back to the pre-NovaStar default of 20 when none is selected.
  const totalSignalPorts = processorModel ? PROCESSOR_SPECS[processorModel].ethernetOutputCount : SIGNAL_PORT_COUNT;
  // With the backup signal loop enabled, the second half of the ports are
  // reserved as backups for the first half - port N backs up port
  // (N - primarySignalPortCount), e.g. for a 20-port processor, port 11
  // backs up port 1 (see getPanelIndicators's backup badge below) - so only
  // the first half remains available for primary assignment.
  const primarySignalPortCount = backupSignalLoop ? Math.max(1, Math.floor(totalSignalPorts / 2)) : totalSignalPorts;
  const signalPorts = useMemo(() => makeSignalPorts(totalSignalPorts), [totalSignalPorts]);

  const [panelsPerPowerOutlet, setPanelsPerPowerOutlet] = useState<number>(panel.defaults.powerPanelsPerOutlet);
  const [panelsPerSignalPort, setPanelsPerSignalPort] = useState<number>(panel.defaults.signalPanelsPerPort);

  const selectedPanel = findCellById(grid, selectedId);
  const activeSelectedKeys = getSelectedIds(selectedCells, selectedId);
  const selectedCount = activeSelectedKeys.size;
  const isPatchTargetActive = patchMode === "signal" ? activePort > 0 : activePowerPort > 0;

  // A dangling activeSubScreenId (e.g. left pointing at a sub-screen an undo
  // just removed) must behave as Canvas View everywhere, not crash/misscope.
  const resolvedActiveSubScreenId = useMemo(
    () => (activeSubScreenId && subScreens.some((s) => s.id === activeSubScreenId) ? activeSubScreenId : null),
    [activeSubScreenId, subScreens],
  );

  // A panel is "dimmed" (visible, non-interactive) when a sub-screen is being
  // edited and the panel isn't part of it. Unassigned panels are dimmed too -
  // otherwise a scoped edit session could silently reach out and touch a
  // panel nobody has categorised yet. Forces an explicit assign-to-sub-screen
  // step (or returning to Canvas View) before it can be selected/patched.
  const isPanelDimmed = (cell: Cell) =>
    subScreens.length > 0 && resolvedActiveSubScreenId !== null && cell.subScreenId !== resolvedActiveSubScreenId;

  // The set of panel ids the active sub-screen (if any) is allowed to
  // touch/patch/reorder, or null for "whole grid" (Canvas View, or no
  // sub-screens created - i.e. today's exact unscoped behaviour).
  const currentScopeIds = (): Set<string> | null =>
    subScreens.length && resolvedActiveSubScreenId !== null
      ? new Set(grid.filter((c) => c.subScreenId === resolvedActiveSubScreenId).map((c) => c.id))
      : null;

  const captureLayout = (): LayoutSnapshot => ({
    panels: cloneGrid(grid),
    subScreens: subScreens.map((s) => ({ ...s })),
    outputCanvasW,
    outputCanvasH,
    wholeLayoutCanvasX,
    wholeLayoutCanvasY,
  });
  const restoreLayout = (snapshot: LayoutSnapshot) => {
    setGrid(cloneGrid(snapshot.panels));
    setSubScreens(snapshot.subScreens.map((s) => ({ ...s })));
    setOutputCanvasW(snapshot.outputCanvasW);
    setOutputCanvasH(snapshot.outputCanvasH);
    setWholeLayoutCanvasX(snapshot.wholeLayoutCanvasX);
    setWholeLayoutCanvasY(snapshot.wholeLayoutCanvasY);
    setSelectedId(null);
    setSelectedCells(new Set());
    setDragVisited(new Set());
    setIsDragging(false);
    setIsSelectingPanels(false);
    setMoveDrag(null);
    setCanvasDrag(null);
  };
  const pushUndoSnapshot = (snapshot = captureLayout()) => {
    setUndoStack((prev) => [...prev.slice(-49), snapshot]);
    setRedoStack([]);
  };
  const commitGridUpdate = (updater: (prev: Cell[]) => Cell[]) => {
    const snapshot = captureLayout();
    setGrid((prev) => updater(prev));
    pushUndoSnapshot(snapshot);
  };
  // Sub-screen CRUD (create/rename/delete/assign/...) and output-canvas
  // position changes route through the same single undo stack as panel
  // edits - one mental model for "undo", not a second parallel history.
  const commitSubScreensUpdate = (updater: (prev: SubScreen[]) => SubScreen[]) => {
    const snapshot = captureLayout();
    setSubScreens((prev) => updater(prev));
    pushUndoSnapshot(snapshot);
  };
  const commitCanvasUpdate = (updater: () => void) => {
    const snapshot = captureLayout();
    updater();
    pushUndoSnapshot(snapshot);
  };
  const undoLayout = () => {
    setUndoStack((prev) => {
      if (!prev.length) return prev;
      const next = [...prev];
      const snapshot = next.pop()!;
      setRedoStack((redoPrev) => [...redoPrev.slice(-49), captureLayout()]);
      restoreLayout(snapshot);
      return next;
    });
  };
  const redoLayout = () => {
    setRedoStack((prev) => {
      if (!prev.length) return prev;
      const next = [...prev];
      const snapshot = next.pop()!;
      setUndoStack((undoPrev) => [...undoPrev.slice(-49), captureLayout()]);
      restoreLayout(snapshot);
      return next;
    });
  };

  // Persist the Rentman equipment mapping on every edit (unlike this file's
  // other localStorage keys, which are one-shot handoffs read once then
  // removed - see equipmentMapping.ts).
  useEffect(() => {
    saveEquipmentMapping(equipmentMapping);
  }, [equipmentMapping]);

  useEffect(() => {
    setPanelsPerPowerOutlet((prev) => {
      const defaultVal = PANEL_TYPES[panelType].defaults.powerPanelsPerOutlet;
      return Math.min(Math.max(prev || defaultVal, 1), 21);
    });
    setPanelsPerSignalPort((prev) => {
      const defaultVal = PANEL_TYPES[panelType].defaults.signalPanelsPerPort;
      return Math.min(Math.max(prev || defaultVal, 1), defaultVal);
    });
  }, [panelType]);

  useEffect(() => {
    setActivePowerPort((prev) => clampActivePort(prev, powerPorts.length));
    setGrid((prev) =>
      prev.map((cell) => {
        if (cell.assignedPowerPort && cell.assignedPowerPort > powerPorts.length) {
          return {
            ...cell,
            assignedPowerPort: null,
            powerSequence: null,
            powerManual: false,
          };
        }
        return { ...cell };
      }),
    );
  }, [powerPorts.length]);

  useEffect(() => {
    const stop = () => {
      setIsDragging(false);
      setIsSelectingPanels(false);
      setSelectionStart(null);
      setSelectionEnd(null);
      setDragVisited(new Set());
      // Releasing outside the workspace cancels an in-flight move (the
      // workspace's own mouseup commits it first when released inside).
      setMoveDrag(null);
      setSnapGuide(null);
    };
    window.addEventListener("mouseup", stop);
    return () => window.removeEventListener("mouseup", stop);
  }, []);

  useEffect(() => {
    const clearPatchTarget = (event: MouseEvent) => {
      const target = event.target as HTMLElement | null;
      if (!target) return;
      if (target.closest("[data-patch-picker]") || target.closest("[data-panel-layout]")) return;
      setActivePort(0);
      setActivePowerPort(0);
    };
    window.addEventListener("click", clearPatchTarget);
    return () => window.removeEventListener("click", clearPatchTarget);
  }, []);

  // Recording progress ticker (display only - the actual stop is a setTimeout
  // inside downloadMovingTestPatternVideo).
  useEffect(() => {
    if (!isRecordingVideo) return;
    const id = window.setInterval(() => setVideoRecordSeconds((s) => Math.min(LOOP_SECONDS, s + 0.25)), 250);
    return () => window.clearInterval(id);
  }, [isRecordingVideo]);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      const target = event.target as HTMLElement | null;
      if (target && ["INPUT", "SELECT", "TEXTAREA"].includes(target.tagName)) return;
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "z") {
        event.preventDefault();
        if (event.shiftKey) redoLayout();
        else undoLayout();
        return;
      }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "y") {
        event.preventDefault();
        redoLayout();
        return;
      }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
        event.preventDefault();
        copySelectedPanels();
        return;
      }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "v") {
        event.preventDefault();
        startPaste();
        return;
      }
      if (event.key === "Escape") {
        if (isPasting) {
          cancelPaste();
          return;
        }
        setSelectedId(null);
        setSelectedCells(new Set());
        setEditMode("patch");
        return;
      }
      if (event.key === "Delete" || event.key === "Backspace") {
        event.preventDefault();
        deleteSelectedPanel();
        return;
      }
      if (event.key.toLowerCase() === "r") {
        event.preventDefault();
        rotateSelectedPanels();
        return;
      }
      if (event.key.toLowerCase() === "c") {
        event.preventDefault();
        clearSelectedPanelPatching();
        return;
      }
      // Mode shortcuts.
      if (event.key.toLowerCase() === "s") {
        event.preventDefault();
        setEditMode("select");
        return;
      }
      if (event.key.toLowerCase() === "m") {
        event.preventDefault();
        setEditMode("move");
        return;
      }
      if (event.key.toLowerCase() === "p") {
        event.preventDefault();
        setEditMode("patch");
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  });

  // Pick up a grid handed off from the standalone Quick Panel Layout tab. If
  // this tab's canvas is empty (including the app's own empty-on-load
  // default) there's nothing to lose, so apply it immediately; otherwise ask
  // via pendingQuickLayoutTransfer -> QuickLayoutTransferModal. The key is
  // removed as soon as it's read so a refresh never re-triggers this, even
  // under StrictMode's double-invoke in dev.
  useEffect(() => {
    let payload: QuickLayoutTransfer | null = null;
    try {
      const raw = localStorage.getItem(QUICK_LAYOUT_TRANSFER_KEY);
      if (raw) payload = JSON.parse(raw) as QuickLayoutTransfer;
    } catch (err) {
      console.error("Quick Panel Layout transfer payload was invalid", err);
    }
    if (!payload) return;
    localStorage.removeItem(QUICK_LAYOUT_TRANSFER_KEY);
    if (grid.length === 0) {
      pushUndoSnapshot();
      setCols(payload.cols);
      setRows(payload.rows);
      setDraftCols(String(payload.cols));
      setDraftRows(String(payload.rows));
      setPanelType(payload.panelType);
      setGrid(makeGridPanels(payload.cols, payload.rows, payload.panelType));
      setSelectedId(null);
      setSelectedCells(new Set());
      if (payload.projectName) setProjectName(payload.projectName);
    } else {
      setPendingQuickLayoutTransfer(payload);
    }
    // Mount-only: this reads a one-shot hand-off, not something to react to.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const applyQuickLayoutTransfer = (mode: "replace" | "add") => {
    const payload = pendingQuickLayoutTransfer;
    if (!payload) return;
    pushUndoSnapshot();
    if (payload.projectName) setProjectName(payload.projectName);
    if (mode === "replace") {
      setCols(payload.cols);
      setRows(payload.rows);
      setDraftCols(String(payload.cols));
      setDraftRows(String(payload.rows));
      setPanelType(payload.panelType);
      setGrid(makeGridPanels(payload.cols, payload.rows, payload.panelType));
      // Replace wipes the whole project's panels - any existing sub-screens
      // no longer have valid members, so start clean (same as importing).
      setSubScreens([]);
      setActiveSubScreenId(null);
      setSelectedId(null);
      setSelectedCells(new Set());
    } else {
      // Add keeps the existing panels (and their own cols/rows/type stay
      // whatever they already were), but the Panel Type selector should
      // still switch to match the batch that was just added, same as
      // picking a new type for "+ Add Panel" would.
      setPanelType(payload.panelType);
      const bbox = activeBBox(activePanels.map(cellRect));
      const GAP_MM = 500;
      const offsetX = bbox.w > 0 ? bbox.x + bbox.w + GAP_MM : 0;
      const offsetY = bbox.w > 0 ? bbox.y : 0;
      const added = makeGridPanels(payload.cols, payload.rows, payload.panelType, resolvedActiveSubScreenId).map((cell) => ({
        ...cell,
        x: cell.x + offsetX,
        y: cell.y + offsetY,
      }));
      setGrid((prev) => [...prev, ...added]);
    }
    setPendingQuickLayoutTransfer(null);
  };

  const maxAllowedPowerPanels = 21;
  const safePanelsPerPowerOutlet = Math.min(Math.max(panelsPerPowerOutlet, 1), maxAllowedPowerPanels);
  const safePanelsPerSignalPort = Math.min(Math.max(panelsPerSignalPort, 1), panel.defaults.signalPanelsPerPort);

  const powerOutletWatts = safePanelsPerPowerOutlet * powerSpec.maxW;
  const powerOutletAmps = safePanelsPerPowerOutlet * powerSpec.maxA;
  const powerOutletPercent = (powerOutletAmps / MAX_OUTLET_AMPS) * 100;

  const panelPixels = panel.pixW * panel.pixH;
  const signalPortPixels = safePanelsPerSignalPort * panelPixels;
  const signalPortPercent = (signalPortPixels / MAX_PIXELS_PER_PORT) * 100;

  // Scope: which sub-screen's panels are "live" for editing/calc purposes.
  // Canvas View (resolvedActiveSubScreenId === null) or "no sub-screens
  // exist" (legacy projects, or projects that never use the feature) means
  // the scope is the whole grid - i.e. today's exact behaviour.
  const scopedGrid = useMemo(() => {
    if (!subScreens.length || resolvedActiveSubScreenId === null) return grid;
    return grid.filter((cell) => cell.subScreenId === resolvedActiveSubScreenId);
  }, [grid, subScreens.length, resolvedActiveSubScreenId]);

  const activeCells = useMemo(() => scopedGrid.filter((cell) => !cell.isRemoved), [scopedGrid]);
  const activePanels = activeCells;
  const totalPanels = activePanels.length;
  // Per-sub-screen bounding boxes (workspace mm), for the boundary/label
  // overlay drawn in the live workspace. Always derived from the FULL grid
  // (not scopedGrid) so every sub-screen's outline is visible regardless of
  // which one is currently being edited.
  const subScreenBBoxes = useMemo(() => {
    const map = new Map<string, RectMm>();
    subScreens.forEach((screen) => {
      const bbox = subScreenBBoxOf(grid, screen.id, cellRect);
      if (bbox.w > 0 && bbox.h > 0) map.set(screen.id, bbox);
    });
    return map;
  }, [grid, subScreens]);
  // Bbox of the FULL active grid (not scope-filtered), used as the "local mm
  // origin" for panels not assigned to any sub-screen when computing their
  // final output-canvas position against the whole-layout canvas placement.
  const fullGridActiveBBox = useMemo(() => activeBBox(grid.filter((c) => !c.isRemoved).map(cellRect)), [grid]);
  // A panel's final position on the output canvas: its own sub-screen's
  // canvas X/Y plus its mm offset within that sub-screen (converted to
  // canvas pixels), or the whole-layout canvas position for panels that
  // aren't assigned to any sub-screen. Never writes back into panel x/y -
  // purely a derived, exported value (see canvasModel.ts).
  const getFinalCanvasPositionOf = (cell: Cell) => {
    if (cell.subScreenId) {
      const screen = subScreens.find((s) => s.id === cell.subScreenId);
      const bbox = subScreenBBoxes.get(cell.subScreenId);
      if (screen && bbox) return finalCanvasPositionOf(cell, bbox, screen.canvasX, screen.canvasY);
    }
    return finalCanvasPositionOf(cell, fullGridActiveBBox, wholeLayoutCanvasX, wholeLayoutCanvasY);
  };
  // Wall size = bounding box of all active panels (free layouts included).
  const wallBBox = useMemo(() => activeBBox(activePanels.map(cellRect)), [activePanels]);
  const wallWidthM = wallBBox.w / 1000;
  const wallHeightM = wallBBox.h / 1000;
  // True outer bounds of the whole layout, INCLUDING any panel rotated to a
  // non-cardinal angle (wallBBox/cellRect only ever swap w/h at 90/270, so a
  // panel spun to e.g. 30deg would otherwise poke outside wallBBox unnoticed).
  // Rotates each panel's own footprint rect around its centre by its full
  // stored rotation - the same box+angle the renderer itself draws - and
  // takes the union of every corner. Drives the vertical centre indicator.
  const trueOuterBBox = useMemo(() => {
    let minX = Infinity;
    let minY = Infinity;
    let maxX = -Infinity;
    let maxY = -Infinity;
    activePanels.forEach((cell) => {
      const rect = cellRect(cell);
      const rotation = cell.rotation ?? 0;
      const cx = rect.x + rect.w / 2;
      const cy = rect.y + rect.h / 2;
      const rad = (rotation * Math.PI) / 180;
      const cos = Math.cos(rad);
      const sin = Math.sin(rad);
      const hw = rect.w / 2;
      const hh = rect.h / 2;
      [[-hw, -hh], [hw, -hh], [hw, hh], [-hw, hh]].forEach(([lx, ly]) => {
        const x = cx + lx * cos - ly * sin;
        const y = cy + lx * sin + ly * cos;
        minX = Math.min(minX, x);
        minY = Math.min(minY, y);
        maxX = Math.max(maxX, x);
        maxY = Math.max(maxY, y);
      });
    });
    if (!Number.isFinite(minX)) return { x: 0, y: 0, w: 0, h: 0 };
    return { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
  }, [activePanels]);
  // Visual row bands (top->bottom, left->right) drive snake order, pixel maths,
  // the PNG test pattern, and row labels for non-uniform layouts.
  const panelBands = useMemo(() => bandPanels(activePanels, cellRect) as Cell[][], [activePanels]);
  const bandIndexById = useMemo(() => {
    const map = new Map<string, number>();
    panelBands.forEach((band, index) => band.forEach((cell) => map.set(cell.id, index)));
    return map;
  }, [panelBands]);
  const panelTypeCounts = useMemo(() => {
    const counts = { MG9: 0, MT: 0 } as Record<PanelTypeKey, number>;
    activePanels.forEach((cell) => {
      counts[cellPanelType(cell)] += 1;
    });
    return counts;
  }, [activePanels]);
  // Pixel resolution uses each panel's native pixels. Because MG9 (168x168) and
  // MT (256x64) have different pitches, a mixed wall isn't a single clean raster:
  // width is the widest band's pixels, height sums each band's tallest panel.
  const wallPixels = useMemo(() => {
    let pixelW = 0;
    let pixelH = 0;
    panelBands.forEach((band) => {
      let rowPixelW = 0;
      let rowPixelH = 0;
      band.forEach((cell) => {
        const p = PANEL_TYPES[cellPanelType(cell)];
        rowPixelW += p.pixW;
        rowPixelH = Math.max(rowPixelH, p.pixH);
      });
      pixelW = Math.max(pixelW, rowPixelW);
      pixelH += rowPixelH;
    });
    return { pixelW, pixelH };
  }, [panelBands]);
  const wallPixelW = wallPixels.pixelW;
  const wallPixelH = wallPixels.pixelH;
  const panelVariantCounts = useMemo(() => {
    const counts = Object.fromEntries(Object.keys(PANEL_VARIANTS).map((key) => [key, 0])) as Record<PanelVariantKey, number>;
    activePanels.forEach((cell) => {
      if (cellPanelType(cell) !== "MG9") return;
      counts[cell.panelVariant ?? "STANDARD"] += 1;
    });
    return counts;
  }, [activePanels]);
  // Shaped panels split by the orientation their rotation puts them in
  // (LU/LD/RU/RD) - each orientation is a separate physical stock item.
  const shapedOrientationCounts = useMemo(() => {
    const zero = () => ({ LU: 0, LD: 0, RU: 0, RD: 0 }) as Record<ShapeOrientationKey, number>;
    const counts = { TRIANGLE: zero(), CURVED: zero() };
    activePanels.forEach((cell) => {
      if (cellPanelType(cell) !== "MG9") return;
      const variant = cell.panelVariant ?? "STANDARD";
      if (variant !== "TRIANGLE" && variant !== "CURVED") return;
      const orientation = getShapeOrientation(variant, cell.rotation);
      if (orientation) counts[variant][orientation] += 1;
    });
    return counts;
  }, [activePanels]);
  // Spare-panel breakdown by surface (each sub-screen, plus "Unassigned" if
  // any panels aren't in one, or just "Whole Layout" when there are no
  // sub-screens) and by panel type bucket - always re-derived from the FULL
  // grid (not the scoped activePanels, which only ever reflects one
  // sub-screen - or none - at a time), so every surface shows at once
  // regardless of which one is currently being edited. Mirrors how the PDF's
  // sub-screens table is built.
  const sparePanelSurfaces = useMemo(() => {
    const zeroBuckets = () => Object.fromEntries(SPARE_BUCKETS.map((b) => [b.key, 0])) as Record<SpareBucketKey, number>;
    const activeGrid = grid.filter((cell) => !cell.isRemoved);
    const surfaces: Array<{ id: string; name: string; buckets: Record<SpareBucketKey, number>; total: number }> = [];

    const tally = (cells: Cell[]) => {
      const buckets = zeroBuckets();
      cells.forEach((cell) => {
        buckets[spareBucketOfCell(cell)] += 1;
      });
      return buckets;
    };

    if (subScreens.length > 0) {
      subScreens.forEach((screen) => {
        const cells = activeGrid.filter((cell) => cell.subScreenId === screen.id);
        if (cells.length === 0) return;
        surfaces.push({ id: screen.id, name: screen.name, buckets: tally(cells), total: cells.length });
      });
      const unassigned = activeGrid.filter((cell) => cell.subScreenId === null);
      if (unassigned.length > 0) {
        surfaces.push({ id: "unassigned", name: "Unassigned", buckets: tally(unassigned), total: unassigned.length });
      }
    } else if (activeGrid.length > 0) {
      surfaces.push({ id: "whole", name: "Whole Layout", buckets: tally(activeGrid), total: activeGrid.length });
    }

    return surfaces;
  }, [grid, subScreens]);
  // Per-surface bucket tallies -> renderable rows with spare/rounded already
  // computed, plus each surface's own subtotal and a project-wide grand
  // total - shared by the on-screen breakdown and the PDF report.
  const sparePanelSummary = useMemo(() => {
    const surfaceRows = sparePanelSurfaces.map((surface) => {
      const bucketRows = SPARE_BUCKETS.map((b) => {
        const used = surface.buckets[b.key];
        if (used === 0) return null;
        const { spare, rounded } = spareForBucket(used, b.key);
        return { label: b.label, used, spare, rounded };
      }).filter((row): row is { label: string; used: number; spare: number; rounded: number } => row !== null);
      const subtotal = bucketRows.reduce(
        (acc, row) => ({ used: acc.used + row.used, spare: acc.spare + row.spare, rounded: acc.rounded + row.rounded }),
        { used: 0, spare: 0, rounded: 0 },
      );
      return { name: surface.name, bucketRows, subtotal };
    });
    const grandTotal = surfaceRows.reduce(
      (acc, s) => ({ used: acc.used + s.subtotal.used, spare: acc.spare + s.subtotal.spare, rounded: acc.rounded + s.subtotal.rounded }),
      { used: 0, spare: 0, rounded: 0 },
    );
    return { surfaceRows, grandTotal, multiSurface: surfaceRows.length > 1 };
  }, [sparePanelSurfaces]);
  // Occupied 0.5m module columns/rows across the wall bbox - used by the
  // frame/floor deployment stock formulas (rectangle-oriented hardware).
  const activeColsCount = useMemo(() => {
    const occupied = new Set<number>();
    activePanels.forEach((cell) => {
      const r = cellRect(cell);
      const first = Math.floor((r.x - wallBBox.x) / MODULE_MM);
      const last = Math.ceil((r.x + r.w - wallBBox.x) / MODULE_MM) - 1;
      for (let i = first; i <= last; i += 1) occupied.add(i);
    });
    return occupied.size;
  }, [activePanels, wallBBox]);
  const activeRowsCount = panelBands.length;
  const activeWallWidthM = wallBBox.w / 1000;
  const activeWallHeightM = wallBBox.h / 1000;
  // Per-type totals: each panel contributes its own weight and power draw.
  const panelTotals = useMemo(() => {
    const totals = { weight: 0, maxW: 0, maxA: 0, avgW: 0, avgA: 0 };
    activePanels.forEach((cell) => {
      const p = PANEL_TYPES[cellPanelType(cell)];
      totals.weight += p.weight;
      totals.maxW += p.power.maxW;
      totals.maxA += p.power.maxA;
      totals.avgW += p.power.avgW;
      totals.avgA += p.power.avgA;
    });
    return totals;
  }, [activePanels]);
  const panelOnlyWeight = panelTotals.weight;
  const decimalRatio = wallPixelH === 0 ? 0 : wallPixelW / wallPixelH;
  const aspectRatio = wallPixelH === 0 ? "0.00" : `${decimalRatio.toFixed(3)}:1`;
  const ratioLabel = useMemo(() => {
    if (wallPixelW <= 0 || wallPixelH <= 0) return "-";
    const g = gcd(wallPixelW, wallPixelH);
    return `${wallPixelW / g}:${wallPixelH / g}`;
  }, [wallPixelW, wallPixelH]);

  // MT is a transparent panel missing every second LED row, so its vertical
  // pixel pitch is twice its horizontal pitch - wallPixelW/H (its native LED
  // grid) isn't a square-pixel raster and its own aspect ratio (above)
  // doesn't match the wall's true physical proportions. Only meaningful for
  // a wall built entirely from one such panel type - a mixed MG9+MT wall's
  // pixel grid is already an approximation (see the wallPixels comment
  // above), so it keeps today's plain resolution/aspect display instead of
  // inventing a blended "content resolution" for it.
  const isMtOnlyWall = totalPanels > 0 && panelTypeCounts.MT === totalPanels;
  const mtContentScaleY = (PANEL_TYPES.MT.h * 1000 / PANEL_TYPES.MT.pixH) / (PANEL_TYPES.MT.w * 1000 / PANEL_TYPES.MT.pixW);
  const contentPixelW = wallPixelW;
  const contentPixelH = isMtOnlyWall ? Math.round(wallPixelH * mtContentScaleY) : wallPixelH;
  // Physical Aspect Ratio - derived from the wall's true physical size (mm,
  // exact gcd reduction), not the raw LED pixel grid.
  const physicalRatioLabel = useMemo(() => {
    if (wallBBox.w <= 0 || wallBBox.h <= 0) return "-";
    const g = gcd(wallBBox.w, wallBBox.h);
    return `${wallBBox.w / g}:${wallBBox.h / g}`;
  }, [wallBBox.w, wallBBox.h]);
  const wallSizeLabel = isMtOnlyWall ? "Physical Size" : "Size";
  // Shared by both the Wall Summary card and the PDF export (see below) -
  // plain "x" separators to match the PDF's existing text style; the JSX
  // Wall Summary card renders its own "×" version directly instead of
  // reusing this array, to match that card's existing style.
  const wallResolutionSummaryLines = isMtOnlyWall
    ? [
        `LED Wall Resolution: ${wallPixelW} x ${wallPixelH}`,
        `Recommended Content Resolution: ${contentPixelW} x ${contentPixelH}`,
        `Physical Aspect Ratio: ${physicalRatioLabel}`,
      ]
    : [`Resolution: ${wallPixelW} x ${wallPixelH}`, `Aspect ratio: ${aspectRatio}`, `Reduced ratio: ${ratioLabel}`];

  const signalPortStats = useMemo(() => {
    const stats: Record<number, SignalPortStat> = Object.fromEntries(
      signalPorts.map((port) => [port.id, { panels: 0, path: [], firstKey: null, lastKey: null }]),
    );

    for (const cell of scopedGrid) {
      if (!isActiveCell(cell)) continue;
      if (!cell.assignedPort || !stats[cell.assignedPort]) continue;
      stats[cell.assignedPort].panels += 1;
      stats[cell.assignedPort].path.push(cell);
    }

    signalPorts.forEach((port) => {
      const stat = stats[port.id];
      stat.path.sort((a, b) => (a.sequence ?? 0) - (b.sequence ?? 0));
      const first = stat.path[0];
      const last = stat.path[stat.path.length - 1];
      stat.firstKey = first ? first.id : null;
      stat.lastKey = last ? last.id : null;
    });

    return stats;
  }, [scopedGrid, signalPorts]);

  const powerPortStats = useMemo(() => {
    const stats: Record<number, PowerPortStat> = Object.fromEntries(
      powerPorts.map((port) => [
        port.id,
        {
          panels: 0,
          maxWatts: 0,
          maxAmps: 0,
          avgWatts: 0,
          avgAmps: 0,
          utilisation: 0,
          phase: port.phase,
          manualPanels: 0,
          path: [],
          firstKey: null,
          lastKey: null,
        },
      ]),
    );

    for (const cell of scopedGrid) {
      if (!isActiveCell(cell)) continue;
      if (!cell.assignedPowerPort || !stats[cell.assignedPowerPort]) continue;
      const stat = stats[cell.assignedPowerPort];
      const cellPower = PANEL_TYPES[cellPanelType(cell)].power;
      stat.panels += 1;
      stat.maxWatts += cellPower.maxW;
      stat.maxAmps += cellPower.maxA;
      stat.avgWatts += cellPower.avgW;
      stat.avgAmps += cellPower.avgA;
      stat.path.push(cell);
      if (cell.powerManual) stat.manualPanels += 1;
    }

    Object.values(stats).forEach((stat) => {
      stat.utilisation = MAX_OUTLET_AMPS > 0 ? (stat.maxAmps / MAX_OUTLET_AMPS) * 100 : 0;
      stat.path.sort((a, b) => (a.powerSequence ?? 0) - (b.powerSequence ?? 0));
      const first = stat.path[0];
      const last = stat.path[stat.path.length - 1];
      stat.firstKey = first ? first.id : null;
      stat.lastKey = last ? last.id : null;
    });

    return stats;
  }, [scopedGrid, powerPorts, powerSpec.maxW, powerSpec.maxA, powerSpec.avgW, powerSpec.avgA]);

  // Chain-start port-number badges for a panel, shared by the live layout and
  // every export. Signal: the chain's first panel gets its primary port
  // number; when the backup signal loop is enabled, the chain's LAST panel
  // also gets a badge with the backup port number (primary port +
  // primarySignalPortCount) - see the totalSignalPorts/primarySignalPortCount
  // comment above for the numbering scheme. Power: the chain's first panel
  // gets its power port number.
  const getPanelIndicators = (cell: Cell) => {
    const key = cell.id;
    const sStat = cell.assignedPort ? signalPortStats[cell.assignedPort] : null;
    const signalBadges: number[] = [];
    if (sStat && cell.assignedPort) {
      if (sStat.firstKey === key) signalBadges.push(cell.assignedPort);
      if (backupSignalLoop && sStat.lastKey === key) signalBadges.push(cell.assignedPort + primarySignalPortCount);
    }
    const pStat = cell.assignedPowerPort ? powerPortStats[cell.assignedPowerPort] : null;
    const powerBadge = pStat && pStat.firstKey === key ? cell.assignedPowerPort : null;
    return { signalBadges, powerBadge };
  };

  const powerPortsUsed = useMemo(() => Object.values(powerPortStats).filter((stat) => stat.panels > 0).length, [powerPortStats]);
  const signalPortsUsed = useMemo(() => Object.values(signalPortStats).filter((stat) => stat.panels > 0).length, [signalPortStats]);
  const effectiveSignalPortsUsed = backupSignalLoop ? signalPortsUsed * 2 : signalPortsUsed;
  // Hanging/fly bars attach along the top row: one MG9 bar per top-row MG9 panel
  // and one MT bar per top-row MT panel (each type uses its own bar hardware).
  const topRowBars = useMemo(() => {
    // Hanging bars attach along the top edge of the wall: count panels whose
    // top edge sits on the bbox top (within half a module for near-misses).
    let mg9 = 0;
    let mt = 0;
    activePanels.forEach((cell) => {
      if (Math.abs(cellRect(cell).y - wallBBox.y) > MODULE_MM / 2) return;
      if (cellPanelType(cell) === "MT") mt += 1;
      else mg9 += 1;
    });
    return { mg9, mt };
  }, [activePanels, wallBBox]);
  const flyBarWeight = topRowBars.mg9 * PANEL_TYPES.MG9.defaults.flyBarWeight + topRowBars.mt * PANEL_TYPES.MT.defaults.flyBarWeight;
  const slingWeight = (topRowBars.mg9 + topRowBars.mt) * PANEL_TYPES.MG9.defaults.slingWeight;
  const powerCableWeight = powerPortsUsed * 3;
  const signalCableWeight = effectiveSignalPortsUsed * 1;
  const additionalWeight =
    (includeFlyBar ? flyBarWeight : 0) +
    (includeSling ? slingWeight : 0) +
    (includePowerCable ? powerCableWeight : 0) +
    (includeSignalCable ? signalCableWeight : 0) +
    (includeCustomWeight ? Number(customWeight || 0) : 0);
  const totalWeight = panelOnlyWeight + additionalWeight;

  const phaseStats = useMemo(() => {
    const phases = {
      P1: { maxWatts: 0, maxAmps: 0, avgWatts: 0, avgAmps: 0, utilisation: 0 },
      P2: { maxWatts: 0, maxAmps: 0, avgWatts: 0, avgAmps: 0, utilisation: 0 },
      P3: { maxWatts: 0, maxAmps: 0, avgWatts: 0, avgAmps: 0, utilisation: 0 },
    };

    powerPorts.forEach((port) => {
      const stat = powerPortStats[port.id];
      if (!stat) return;
      phases[port.phase as keyof typeof phases].maxWatts += stat.maxWatts;
      phases[port.phase as keyof typeof phases].maxAmps += stat.maxAmps;
      phases[port.phase as keyof typeof phases].avgWatts += stat.avgWatts;
      phases[port.phase as keyof typeof phases].avgAmps += stat.avgAmps;
    });

    Object.values(phases).forEach((phase) => {
      phase.utilisation = distro.safePhaseWatts > 0 ? (phase.maxWatts / distro.safePhaseWatts) * 100 : 0;
    });

    return phases;
  }, [powerPorts, powerPortStats, distro.safePhaseWatts]);

  const totalPowerMaxW = panelTotals.maxW;
  const totalPowerMaxA = panelTotals.maxA;
  const totalPowerAvgW = panelTotals.avgW;
  const totalPowerAvgA = panelTotals.avgA;
  const unassignedPowerPanels = activePanels.filter((cell) => !cell.assignedPowerPort).length;

  // Spares and boxes are per type (different spare ratios and box sizes).
  const mg9Count = panelTypeCounts.MG9;
  const mtCount = panelTypeCounts.MT;
  const mg9Defaults = PANEL_TYPES.MG9.defaults;
  const mtDefaults = PANEL_TYPES.MT.defaults;
  const mg9Spare = Math.ceil(mg9Count * mg9Defaults.spareRatio);
  const mtSpare = Math.ceil(mtCount * mtDefaults.spareRatio);
  const mg9Boxes = mg9Count > 0 ? Math.ceil((mg9Count + mg9Spare) / mg9Defaults.panelsPerBox) : 0;
  const mtBoxes = mtCount > 0 ? Math.ceil((mtCount + mtSpare) / mtDefaults.panelsPerBox) : 0;
  const sparePanels = mg9Spare + mtSpare;
  const totalPanelsWithSpare = totalPanels + sparePanels;
  const boxCount = mg9Boxes + mtBoxes;
  const boxSparePanels = mg9Boxes * mg9Defaults.panelsPerBox + mtBoxes * mtDefaults.panelsPerBox - totalPanelsWithSpare;
  const vx1000Percent = (wallPixelW * wallPixelH / 6500000) * 100;
  const vx2000Percent = (wallPixelW * wallPixelH / 13000000) * 100;
  const circuitsUsedMax = Math.ceil(totalPanels / Math.max(safePanelsPerPowerOutlet, 1));
  const powerPerCircuitMaxW = safePanelsPerPowerOutlet * powerSpec.maxW;
  const powerPerCircuitMaxA = safePanelsPerPowerOutlet * powerSpec.maxA;

  const resolutionOptions = [
    [640, 480], [800, 600], [1024, 768], [1280, 720], [1280, 800], [1280, 1024], [1366, 768], [1440, 900], [1600, 900],
    [1600, 1200], [1680, 1050], [1920, 1080], [1920, 1200], [2048, 1080], [2560, 1440], [2560, 1600], [3440, 1440], [3840, 2160], [4096, 2160], [5120, 2880], [6016, 3384],
  ];

  const bestResolution = useMemo(() => {
    const valid = resolutionOptions.filter(([w, h]) => w >= wallPixelW && h >= wallPixelH);
    if (!valid.length) return null;
    return valid.sort((a, b) => a[0] * a[1] - b[0] * b[1])[0];
  }, [wallPixelW, wallPixelH]);

  const signalCableBaseRequired = signalPortsUsed;
  const signalCableWithBackupRequired = backupSignalLoop ? signalCableBaseRequired * 2 : signalCableBaseRequired;
  const signalCableSpare = Math.ceil(signalCableWithBackupRequired * panel.defaults.signalSpareRatio);
  const powerCableSpare = Math.ceil(circuitsUsedMax * panel.defaults.powerSpareRatio);
  const distroRequired = Math.max(1, Math.ceil(powerPortsUsed / distro.portCount));

  const cornerJoinStats = useMemo(() => {
    // Corner-panel joins = flush shared edges with neighbours (position-based,
    // works for free layouts too). Corner-to-corner pairs counted once.
    let cornerToFlat = 0;
    let cornerToCorner = 0;
    const corners = activeCells.filter((cell) => cell.panelVariant === "CORNER");
    corners.forEach((cell) => {
      const geom = cellGeom(cell);
      activeCells.forEach((other) => {
        if (other.id === cell.id) return;
        if (!panelsAnchorJoined(geom, cellGeom(other))) return;
        if (other.panelVariant === "CORNER") {
          if (other.id > cell.id) cornerToCorner += 1;
        } else {
          cornerToFlat += 1;
        }
      });
    });

    return { cornerToFlat, cornerToCorner };
  }, [activeCells, grid]);

  const deploymentWarning = useMemo(() => {
    if ((deploymentType === DEPLOYMENT_TYPES.GROUND || deploymentType === DEPLOYMENT_TYPES.FLOOR) && mtCount > 0) {
      return `${deploymentType} deployment hardware is available for MG9 only - only the MG9 panels are included in the frame/floor stock.`;
    }
    if (deploymentType === DEPLOYMENT_TYPES.FLOOR && ((activeWallWidthM % 1 !== 0) || (activeWallHeightM % 1 !== 0))) {
      return "Floor deployment uses full 1m frame sections only. This wall size is not an exact ground-frame build.";
    }
    return "";
  }, [activeWallHeightM, activeWallWidthM, deploymentType, panelType]);

  const stockRows = useMemo(() => {
    // Panel-specific stock lives in each type's catalog; shared items (distro,
    // cables, prod case, joiners) are tracked in the MG9 catalog.
    const mg9StockCat = PANEL_TYPES.MG9.stock as Record<string, number>;
    const mtStockCat = PANEL_TYPES.MT.stock as Record<string, number>;
    const stock = mg9StockCat;
    const rowsOut: StockRow[] = [];
    // `required` is always the raw quantity needed to build the wall - `spare`
    // and `rounded` (defaulting to required + spare) carry the real order/pull
    // quantity, and `net` (shortfall) is checked against THAT, not the bare
    // required count.
    const pushBaseRow = (code: string, name: string, required: number, stockQty: number, method: string, spare = 0, rounded = required + spare) => {
      rowsOut.push({ code, name, required, spare, rounded, stock: stockQty, net: stockQty - rounded, method });
    };

    if (mg9Count > 0) {
      const standardCount = panelVariantCounts.STANDARD;
      const standardSpare = Math.ceil(standardCount * mg9Defaults.spareRatio);
      const standardRounded = roundUpToBox(standardCount + standardSpare, mg9Defaults.panelsPerBox);
      pushBaseRow(
        "12224",
        "MG9 LED Panel",
        standardCount,
        mg9StockCat.panels ?? 0,
        `${standardCount} + ${standardSpare} spare, rounded to box of ${mg9Defaults.panelsPerBox}`,
        standardSpare,
        standardRounded,
      );

      // Shaped panels (triangle / quarter circle) are one-way physical pieces:
      // each rotation orientation (LU/LD/RU/RD) is its own stock line, checked
      // against the per-orientation shelf quantity - same as the layout tool.
      (["TRIANGLE", "CURVED"] as const).forEach((variantKey) => {
        const variant = PANEL_VARIANTS[variantKey];
        const item = variant.stockItem;
        if (!item || panelVariantCounts[variantKey] <= 0) return;
        (Object.keys(SHAPE_ORIENTATIONS) as ShapeOrientationKey[]).forEach((orientationKey) => {
          const count = shapedOrientationCounts[variantKey][orientationKey];
          if (count <= 0) return;
          const orientation = SHAPE_ORIENTATIONS[orientationKey];
          const spare = Math.ceil(count * mg9Defaults.spareRatio);
          const stockQty = SHAPED_STOCK_PER_ORIENTATION[variantKey];
          pushBaseRow(
            `${item.code}-${orientationKey}`,
            `${variant.label} ${orientation.icon} ${orientation.label}`,
            count,
            stockQty,
            `${count} placed at this orientation + ${spare} spare`,
            spare,
          );
        });
      });

      // Corner panels are orientation-free; keep the original single line.
      {
        const item = PANEL_VARIANTS.CORNER.stockItem;
        const count = panelVariantCounts.CORNER;
        if (item && count > 0) {
          const spare = Math.ceil(count * mg9Defaults.spareRatio);
          const rounded = roundUpToBox(count + spare, mg9Defaults.panelsPerBox);
          rowsOut.push(makeStockRow(item, count, `${count} selected + ${spare} spare, rounded to box of ${mg9Defaults.panelsPerBox}`, spare, rounded));
        }
      }
    }

    if (mtCount > 0) {
      pushBaseRow("12223", "MT Mesh Panel", mtCount, mtStockCat.panels ?? 0, `${mtCount} + ${mtSpare} spare`, mtSpare);
    }

    rowsOut.push(makeStockRow(STOCK_CATALOG.prodCase, 1, "always 1 per project"));
    if (mg9Boxes > 0) {
      rowsOut.push({ code: "BOX-MG9", name: "Boxes required (MG9)", required: mg9Boxes, stock: mg9Boxes, net: 0, method: `ceil(${mg9Count + mg9Spare}/${mg9Defaults.panelsPerBox})` });
    }
    if (mtBoxes > 0) {
      rowsOut.push({ code: "BOX-MT", name: "Boxes required (MT)", required: mtBoxes, stock: mtBoxes, net: 0, method: `ceil(${mtCount + mtSpare}/${mtDefaults.panelsPerBox})` });
    }

    if (deploymentType === DEPLOYMENT_TYPES.FLOWN) {
      if (topRowBars.mg9 > 0) {
        rowsOut.push({
          code: "12257",
          name: "MG9 Floor / Hanging Bar",
          required: topRowBars.mg9,
          stock: mg9StockCat.hangingBar ?? 0,
          net: (mg9StockCat.hangingBar ?? 0) - topRowBars.mg9,
          method: "1 per top-row MG9 panel",
        });
      }
      if (topRowBars.mt > 0) {
        rowsOut.push({
          code: "12262",
          name: "MT Floor / Hanging Bar",
          required: topRowBars.mt,
          stock: mtStockCat.hangingBar ?? 0,
          net: (mtStockCat.hangingBar ?? 0) - topRowBars.mt,
          method: "1 per top-row MT panel",
        });
      }
    }

    rowsOut.push({
      code: powerDistro === "32A" ? "12245" : "12246",
      name: powerDistro === "32A" ? "32A 3-phase Power Distro" : "63A 3-phase Power Distro",
      required: distroRequired,
      stock: powerDistro === "32A" ? stock.distro32 ?? 0 : stock.distro63 ?? 0,
      net: (powerDistro === "32A" ? stock.distro32 ?? 0 : stock.distro63 ?? 0) - distroRequired,
      method: "selected distro",
    });

    pushBaseRow("12254", "15m PowerCON Cable", circuitsUsedMax, stock.powerCable15m ?? 0, `${circuitsUsedMax} + ${powerCableSpare} spare`, powerCableSpare);
    pushBaseRow(
      "12263",
      "15m Signal Cable",
      signalCableWithBackupRequired,
      stock.signalCable15m ?? 0,
      `${signalCableWithBackupRequired}${backupSignalLoop ? ` (${signalCableBaseRequired} x 2 backup loop)` : ""} + ${signalCableSpare} spare`,
      signalCableSpare,
    );

    if (backupSignalLoop) {
      const joinerRequired = signalPortsUsed;
      const joinerOverflow = Math.max(0, joinerRequired - STOCK_CATALOG.signalJoiner.stock);
      rowsOut.push(makeStockRow(STOCK_CATALOG.signalJoiner, joinerRequired, "1 per signal port for backup loop"));
      if (joinerOverflow > 0) {
        rowsOut.push(makeStockRow(STOCK_CATALOG.signalJoinerCable, joinerOverflow, `joiner stock exhausted, overflow ${joinerOverflow}`));
      } else {
        rowsOut.push(makeStockRow(STOCK_CATALOG.signalJoinerCable, 0, `fallback only if ${STOCK_CATALOG.signalJoiner.name} stock is exhausted`));
      }
    }

    if (mg9Count > 0) {
      const flatConnectorRequired = cornerJoinStats.cornerToFlat * 3;
      const cornerConnectorRequired = cornerJoinStats.cornerToCorner * 3;
      if (flatConnectorRequired > 0) {
        rowsOut.push(makeStockRow(STOCK_CATALOG.cornerFlatConnector, flatConnectorRequired, `3 per corner-to-flat join across ${cornerJoinStats.cornerToFlat} joins`));
      }
      if (cornerConnectorRequired > 0) {
        rowsOut.push(makeStockRow(STOCK_CATALOG.cornerCornerConnector, cornerConnectorRequired, `3 per corner-to-corner join across ${cornerJoinStats.cornerToCorner} joins`));
      }
    }

    if (mg9Count > 0 && includeReinforcementPlate) {
      pushBaseRow("12264", "MG9 Reinforcement Plate", Math.ceil(mg9Count * 0.86), stock.reinforcementPlate ?? 0, "sheet-style factor (MG9 panels)");
      pushBaseRow("12265", "MG9 Reinforcement Screw", Math.ceil(mg9Count * 3.42), stock.reinforcementScrew ?? 0, "sheet-style factor (MG9 panels)");
    }

    if (mg9Count > 0 && deploymentType === DEPLOYMENT_TYPES.GROUND) {
      const widthUnits = Math.floor(activeWallWidthM);
      const verticalSupports = Math.ceil(activeColsCount / 2);
      const verticalFrameHeightCount = Math.ceil(activeRowsCount / 3);
      const backBraces = verticalSupports;
      const horizontalFramePieces = Math.max(verticalSupports - 1, 0) * (verticalFrameHeightCount + 1);
      const verticalFrames = verticalSupports * verticalFrameHeightCount;
      const modularFrameCount = verticalFrames + backBraces + horizontalFramePieces;
      const verticalJoinCount = verticalSupports * Math.max(verticalFrameHeightCount - 1, 0);
      const verticalScrewCount = verticalSupports * Math.max(verticalFrameHeightCount, 0) * 2;
      const horizontalScrewCount = horizontalFramePieces * 4;
      rowsOut.push(makeStockRow(STOCK_CATALOG.modularFrame950, modularFrameCount, `${verticalFrames} vertical + ${backBraces} back brace + ${horizontalFramePieces} horizontal`));
      rowsOut.push(makeStockRow(STOCK_CATALOG.bottomBeam1m, widthUnits, `${widthUnits} full 1m bottom beams`));
      rowsOut.push(makeStockRow(STOCK_CATALOG.modularFrameScrew, verticalScrewCount + horizontalScrewCount, `${verticalScrewCount} vertical/back brace + ${horizontalScrewCount} horizontal screws`));
      rowsOut.push(makeStockRow(STOCK_CATALOG.modularFrameUCoupler, verticalFrames * 2, `2 per vertical frame across ${verticalFrames} frames`));
      rowsOut.push(makeStockRow(STOCK_CATALOG.connectingJoint, verticalJoinCount * 2, `2 per vertical join across ${verticalJoinCount} joins`));
    }

    if (mg9Count > 0 && deploymentType === DEPLOYMENT_TYPES.FLOOR) {
      const feet = Math.ceil(mg9Count / 2);
      const perimeterSegments = activeColsCount * 2 + activeRowsCount * 2;
      rowsOut.push(makeStockRow(STOCK_CATALOG.danceFloorFeet, feet, "1 per 2 panels"));
      rowsOut.push(makeStockRow(STOCK_CATALOG.temperedGlass, mg9Count, "1 per MG9 panel"));
      rowsOut.push(makeStockRow(STOCK_CATALOG.floorReinforcementBar, feet, "1 per foot"));
      rowsOut.push(makeStockRow(STOCK_CATALOG.floorTaperPin, feet * 4, "4 per foot"));
      rowsOut.push(makeStockRow(STOCK_CATALOG.danceFloorRamp, perimeterSegments, `${perimeterSegments} external 500mm edge segments`));
      rowsOut.push(makeStockRow(STOCK_CATALOG.danceFloorRampCorner, 4, "1 per corner"));
    }

    return rowsOut;
  }, [activeColsCount, activeRowsCount, activeWallWidthM, backupSignalLoop, circuitsUsedMax, cornerJoinStats, deploymentType, distroRequired, includeReinforcementPlate, panelVariantCounts, shapedOrientationCounts, powerCableSpare, powerDistro, signalCableBaseRequired, signalCableSpare, signalCableWithBackupRequired, signalPortsUsed, powerPortsUsed, distro.portCount, mg9Count, mtCount, mg9Spare, mtSpare, mg9Boxes, mtBoxes, mg9Defaults, mtDefaults, topRowBars]);

  // The on-screen table, PDF table and CSV export all list order/pull
  // quantities, not raw internal line items - a row whose real order
  // quantity (rounded, spare included) comes out to 0 is just noise there.
  // Live Rentman stock/availability (see src/rentman/) is overlaid HERE,
  // inside this one useMemo, rather than as a separate variable each
  // consumer has to remember to switch to - every real consumer (the
  // on-screen table, CSV export, PDF table, and shortfallRows right below)
  // already reads visibleStockRows, so they all become Rentman-aware for
  // free. Empty liveStock/liveAvailable maps make applyLiveRentmanData a
  // full no-op, so this is a zero-behaviour-change default.
  const visibleStockRows = useMemo(
    () => applyLiveRentmanData(stockRows.filter((row) => (row.rounded ?? row.required) > 0), liveStock, liveAvailable),
    [stockRows, liveStock, liveAvailable],
  );
  const shortfallRows = visibleStockRows.filter((row) => row.net < 0);
  const hasLiveAvailableColumn = visibleStockRows.some((row) => typeof row.available === "number");
  // Every real (orientation-normalized, synthetic-box-excluded) stock code
  // currently on the wall, for the equipment-mapping UI - one entry per
  // distinct base code, first-seen name wins.
  const mappableRentmanItems = useMemo<MappableItem[]>(() => {
    const seen = new Map<string, string>();
    stockRows.filter(isMappableStockRow).forEach((row) => {
      const code = baseCodeOf(row.code);
      if (!seen.has(code)) seen.set(code, row.name);
    });
    return Array.from(seen, ([code, name]) => ({ code, name }));
  }, [stockRows]);

  const refreshRentmanStock = async () => {
    // equipmentMapping is keyed by OUR local stock code (e.g. "12224"), but
    // its value's own .code is Rentman's equipment code (e.g. "877") - the
    // Worker/Rentman only know about the latter, so the fetch has to go out
    // keyed by ref.code, then get translated back to our local codes before
    // being stored (applyLiveRentmanData looks liveStock up by local code).
    const mappedEntries = mappableRentmanItems
      .map((item) => ({ localCode: item.code, ref: equipmentMapping[item.code] }))
      .filter((entry): entry is { localCode: string; ref: RentmanEquipmentRef } => Boolean(entry.ref));
    if (!mappedEntries.length) return;
    const rentmanCodes = mappedEntries.map((entry) => entry.ref.code);
    setRentmanRefreshing(true);
    setRentmanRefreshError(null);
    try {
      const stockResult = await fetchEquipmentStock(rentmanCodes);
      const nextStock: Record<string, LiveStockEntry> = {};
      mappedEntries.forEach(({ localCode, ref }) => {
        const entry = stockResult[ref.code];
        if (entry) nextStock[localCode] = entry;
      });
      setLiveStock(nextStock);

      if (rentmanDateFrom && rentmanDateTo) {
        const availabilityResult = await fetchEquipmentAvailability(rentmanCodes, rentmanDateFrom, rentmanDateTo);
        const nextAvailable: Record<string, number> = {};
        mappedEntries.forEach(({ localCode, ref }) => {
          const qty = availabilityResult[ref.code];
          if (typeof qty === "number") nextAvailable[localCode] = qty;
        });
        setLiveAvailable(nextAvailable);
      } else {
        setLiveAvailable({});
      }
      setRentmanLastRefreshedAt(new Date());
    } catch (err) {
      setRentmanRefreshError(err instanceof Error ? err.message : "Rentman refresh failed");
    } finally {
      setRentmanRefreshing(false);
    }
  };

  const setRentmanEquipmentMapping = (code: string, ref: RentmanEquipmentRef | null) => {
    setEquipmentMapping((prev) => {
      if (ref) return { ...prev, [code]: ref };
      const next = { ...prev };
      delete next[code];
      return next;
    });
  };
  const safeProjectName = projectName.trim() || "Untitled Project";
  const fileSafeProjectName = safeProjectName.replace(/[<>:"/\\|?*\x00-\x1F]/g, "-").replace(/\s+/g, "-");
  // Describe the panel mix for exports and headings.
  const panelTypeSummary =
    mg9Count > 0 && mtCount > 0
      ? `Mixed (${mg9Count} MG9 + ${mtCount} MT)`
      : mtCount > 0
        ? "MT"
        : "MG9";
  const fileSafePanelType = mg9Count > 0 && mtCount > 0 ? "MIX" : mtCount > 0 ? "MT" : "MG9";

  // Draw every signal + power cable route onto a canvas using the shared
  // orthogonal router, so the layout view, PDF and PNG all match. dispRectPx
  // maps a panel to its px rect in that canvas's coordinate space.
  // Cable LINES only, with a thin black outline (drawn under the colour). Meant
  // to be painted BEHIND the panels.
  const drawCanvasCableLines = (ctx: CanvasRenderingContext2D, dispRectPx: (cell: Cell) => RectMm) => {
    const drawPath = (path: Cell[] | undefined, color: string, offset: number) => {
      if (!path || path.length < 2) return;
      ctx.lineCap = "round";
      ctx.lineJoin = "round";
      for (let idx = 1; idx < path.length; idx += 1) {
        const pts = routeCablePx(dispRectPx(path[idx - 1]), dispRectPx(path[idx]), offset);
        const trace = () => {
          ctx.beginPath();
          ctx.moveTo(pts[0].x, pts[0].y);
          for (let k = 1; k < pts.length; k += 1) ctx.lineTo(pts[k].x, pts[k].y);
        };
        ctx.strokeStyle = "#000000";
        ctx.lineWidth = 6;
        trace();
        ctx.stroke();
        ctx.strokeStyle = color;
        ctx.lineWidth = 4;
        trace();
        ctx.stroke();
      }
    };
    Object.entries(signalPortStats).forEach(([portId, stat]) => {
      drawPath(stat.path, PORT_COLORS[(Number(portId) - 1) % PORT_COLORS.length], -4);
    });
    powerPorts.forEach((port) => {
      drawPath(powerPortStats[port.id]?.path, POWER_COLOR, 4);
    });
  };

  // Cable ARROWHEADS only (black-outlined), painted IN FRONT of the panels so the
  // signal/power direction stays visible.
  const drawCanvasCableArrows = (ctx: CanvasRenderingContext2D, dispRectPx: (cell: Cell) => RectMm) => {
    const drawPath = (path: Cell[] | undefined, color: string, offset: number) => {
      if (!path || path.length < 2) return;
      for (let idx = 1; idx < path.length; idx += 1) {
        const ra = dispRectPx(path[idx - 1]);
        const rb = dispRectPx(path[idx]);
        const pts = routeCablePx(ra, rb, offset);
        const last = pts[pts.length - 1];
        let prev = pts[pts.length - 2];
        // Touching panels collapse the last segment to a point; fall back to the
        // source->destination centre direction so the arrow still points the
        // right way (up/down/left/right).
        if (Math.hypot(last.x - prev.x, last.y - prev.y) < 1) {
          prev = { x: last.x - ((rb.x + rb.w / 2) - (ra.x + ra.w / 2)), y: last.y - ((rb.y + rb.h / 2) - (ra.y + ra.h / 2)) };
        }
        drawCanvasArrowHead(ctx, prev.x, prev.y, last.x, last.y, color);
      }
    };
    Object.entries(signalPortStats).forEach(([portId, stat]) => {
      drawPath(stat.path, PORT_COLORS[(Number(portId) - 1) % PORT_COLORS.length], -4);
    });
    powerPorts.forEach((port) => {
      drawPath(powerPortStats[port.id]?.path, POWER_COLOR, 4);
    });
  };

  const buildLayoutCanvas = (flipped = false, viewLabel = "Back View") => {
    const px = CELL_SIZE / MODULE_MM; // export scale, independent of on-screen zoom
    const margin = 52;
    const wallW = Math.max(1, Math.round(wallBBox.w * px));
    const wallH = Math.max(1, Math.round(wallBBox.h * px));
    const contentW = wallW + margin * 2;
    const contentH = wallH + margin * 2 + 20;

    // How big this image will actually print, so we can render it at just
    // enough pixel density to hit PDF_LAYOUT_IMAGE_DPI there - see the
    // constant's own comment for why this can't be a flat multiplier.
    const contentRatio = contentW / contentH;
    let printWidthMm = PDF_LAYOUT_USABLE_WIDTH_MM;
    let printHeightMm = printWidthMm / contentRatio;
    if (printHeightMm > PDF_LAYOUT_USABLE_HEIGHT_MM) {
      printHeightMm = PDF_LAYOUT_USABLE_HEIGHT_MM;
      printWidthMm = printHeightMm * contentRatio;
    }
    const scale = (PDF_LAYOUT_IMAGE_DPI / 25.4) * (printWidthMm / contentW);

    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(contentW * scale));
    canvas.height = Math.max(1, Math.round(contentH * scale));
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Canvas context unavailable");
    ctx.scale(scale, scale);

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, wallW + margin * 2, wallH + margin * 2 + 20);
    ctx.fillStyle = "#0f172a";
    ctx.font = "bold 18px Arial";
    ctx.textAlign = "left";
    ctx.fillText(viewLabel, 16, 24);

    ctx.save();
    ctx.translate(margin, margin + 20);

    // Panel rect in export px, mirrored for the front view.
    const dispRectPx = (cell: Cell): RectMm => {
      const raw = cellRect(cell);
      const d = flipped ? mirrorRectX(raw, wallBBox) : raw;
      return { x: (d.x - wallBBox.x) * px, y: (d.y - wallBBox.y) * px, w: d.w * px, h: d.h * px };
    };

    // Metre ruler along the top and left edges.
    ctx.strokeStyle = "#94a3b8";
    ctx.fillStyle = "#475569";
    ctx.font = "11px Arial";
    ctx.textAlign = "center";
    ctx.lineWidth = 1;
    for (let m = 0; m * 1000 <= wallBBox.w + 1; m += 0.5) {
      const x = m * 1000 * px;
      ctx.beginPath();
      ctx.moveTo(x, -4);
      ctx.lineTo(x, m % 1 === 0 ? -12 : -8);
      ctx.stroke();
      if (m % 1 === 0) ctx.fillText(`${m}m`, x, -16);
    }
    ctx.textAlign = "right";
    // Height ruler reads bottom-up (0m at the wall's base), matching the live
    // workspace - only the printed label flips, not the tick positions.
    const maxHeightM = Math.floor(wallBBox.h / 1000);
    for (let m = 0; m * 1000 <= wallBBox.h + 1; m += 0.5) {
      const y = m * 1000 * px;
      ctx.beginPath();
      ctx.moveTo(-4, y);
      ctx.lineTo(m % 1 === 0 ? -12 : -8, y);
      ctx.stroke();
      if (m % 1 === 0) ctx.fillText(`${maxHeightM - m}m`, -16, y + 4);
    }

    // Cable lines first so they sit behind the panels.
    drawCanvasCableLines(ctx, dispRectPx);

    activePanels.forEach((cell) => {
      if (!isPanelHead(cell)) return;
      const r = dispRectPx(cell);
      const fill = cell.assignedPort ? PORT_COLORS[(cell.assignedPort - 1) % PORT_COLORS.length] : "#1e293b";
      const { signalBadges, powerBadge } = getPanelIndicators(cell);
      drawPanelShape(ctx, r.x, r.y, r.w, r.h, cell, fill, "#0f172a", 2, { signalBadges, powerBadge, mirrorX: flipped });
    });

    // Only clutter the diagram with canvas-position labels once the user has
    // actually engaged with sub-screens/output-canvas positioning.
    const showCanvasLabels = subScreens.length > 0 || wholeLayoutCanvasX !== 0 || wholeLayoutCanvasY !== 0;
    activePanels.forEach((cell) => {
      if (!isPanelHead(cell)) return;
      const r = dispRectPx(cell);
      const cx = r.x + r.w / 2;
      ctx.fillStyle = "#020617";
      ctx.font = "bold 10px Arial";
      ctx.textAlign = "center";
      // Stack all per-panel info text from the BOTTOM of the panel upward,
      // leaving the top corners clear for the signal/power port-number
      // badges (drawn by drawPanelShape above, in the same top-left/top-right
      // spots as the live workspace).
      let by = r.y + r.h - 6;
      const variantSymbol = getPanelSymbol(cell);
      if (variantSymbol) {
        ctx.fillText(variantSymbol, cx, by);
        by -= 14;
      }
      if (showCanvasLabels) {
        const finalPos = getFinalCanvasPositionOf(cell);
        ctx.fillText(`CX ${finalPos.x} CY ${finalPos.y}`, cx, by);
        by -= 16;
      }
      if (cell.assignedPowerPort) {
        ctx.fillText(`⚡ Plug ${cell.assignedPowerPort}`, cx, by);
        by -= 16;
      }
      if (cell.assignedPort) {
        ctx.fillText(`🔌 P${cell.assignedPort} (${cell.sequence ?? "-"})`, cx, by);
        by -= 16;
      }
      ctx.fillText(`↓ ${panelRowLabel(cell)} → ${panelColLabel(cell)}${cellPanelType(cell) === "MT" ? " (MT)" : ""}`, cx, by);
    });

    // Arrowheads last so the signal/power direction stays visible in front.
    drawCanvasCableArrows(ctx, dispRectPx);

    // Vertical centre indicator - opt-in only (Include Centre Line checkbox),
    // mirrors the same trueOuterBBox-based calculation as the live workspace.
    if (includeCentreLineInExport && showCentreLine && trueOuterBBox.w > 0) {
      const centreTrueX = trueOuterBBox.x + trueOuterBBox.w / 2;
      const centreDisplayTrueX = flipped ? 2 * wallBBox.x + wallBBox.w - centreTrueX : centreTrueX;
      const lineX = (centreDisplayTrueX - wallBBox.x) * px;
      const yTop = (trueOuterBBox.y - wallBBox.y) * px;
      const yBottom = (trueOuterBBox.y + trueOuterBBox.h - wallBBox.y) * px;
      ctx.save();
      ctx.strokeStyle = "#eab308";
      ctx.lineWidth = 1.5;
      ctx.setLineDash([6, 4]);
      ctx.globalAlpha = 0.85;
      ctx.beginPath();
      ctx.moveTo(lineX, yTop);
      ctx.lineTo(lineX, yBottom);
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.globalAlpha = 1;
      ctx.fillStyle = "#eab308";
      ctx.fillRect(lineX - 20, Math.max(0, yTop - 16), 40, 14);
      ctx.fillStyle = "#1e293b";
      ctx.font = "bold 10px Arial";
      ctx.textAlign = "center";
      ctx.fillText("Centre", lineX, Math.max(10, yTop - 6));
      ctx.restore();
    }

    ctx.restore();
    return canvas;
  };

  // --- NovaStar processor configuration export ------------------------------
  // One entry per canvas entry (sub-screens, or a single WHOLE_LAYOUT_KEY
  // entry when none exist yet) - matches OutputCanvasPanel's own "whole
  // layout vs. sub-screens" branching exactly.
  const canvasInputsList: CanvasEntryInput[] = useMemo(() => {
    const keys = subScreens.length ? subScreens.map((s) => s.id) : [WHOLE_LAYOUT_KEY];
    return keys.map((key) => ({
      key,
      name: key === WHOLE_LAYOUT_KEY ? "Whole Layout" : subScreens.find((s) => s.id === key)?.name ?? key,
      interfacePk: canvasInputs[key] ?? null,
    }));
  }, [subScreens, canvasInputs]);

  const novaStarValidation = useMemo(() => {
    if (!processorModel) return null;
    return buildExportSummaryAndCabinets({
      processorModel,
      projectName: safeProjectName,
      surfaceName,
      outputCanvasW,
      outputCanvasH,
      wholeLayoutCanvasX,
      wholeLayoutCanvasY,
      grid,
      subScreens,
      inputMode,
      wholeCanvasInputId,
      canvasInputs: canvasInputsList,
    });
  }, [
    processorModel,
    safeProjectName,
    surfaceName,
    outputCanvasW,
    outputCanvasH,
    wholeLayoutCanvasX,
    wholeLayoutCanvasY,
    grid,
    subScreens,
    inputMode,
    wholeCanvasInputId,
    canvasInputsList,
  ]);

  const downloadNovaStarConfig = async () => {
    if (!processorModel) return;
    setIsGeneratingNovaStarFile(true);
    try {
      const result = await buildNovaStarExport({
        processorModel,
        projectName: safeProjectName,
        surfaceName,
        outputCanvasW,
        outputCanvasH,
        wholeLayoutCanvasX,
        wholeLayoutCanvasY,
        grid,
        subScreens,
        inputMode,
        wholeCanvasInputId,
        canvasInputs: canvasInputsList,
      });
      if (!result.ok || !result.blob || !result.fileName) {
        window.alert("NovaStar export blocked by validation errors - see the NovaStar Processor Configuration section.");
        return;
      }
      const url = window.URL.createObjectURL(result.blob);
      const link = document.createElement("a");
      link.href = url;
      link.setAttribute("download", result.fileName);
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => window.URL.revokeObjectURL(url), 1000);
    } catch (err) {
      console.error("NovaStar config download failed", err);
      window.alert("NovaStar config download failed - check console");
    } finally {
      setIsGeneratingNovaStarFile(false);
    }
  };

const exportJson = () => {
  try {
    // formatVersion 6: adds the Rentman availability date range (v1-v5
    // files still open, see openJson - the number bump itself is purely
    // documentation, there's no branching logic tied to it anywhere).
    // stockRows here is always the plain catalog-based numbers (this file
    // snapshots the theoretical required/spare/rounded math, not a live
    // Rentman read that would just go stale the moment the file is
    // reopened) - the equipment mapping that drives live data is account-
    // wide and lives in localStorage instead, not in this per-project file.
    const payload = {
      formatVersion: 6,
      appVersion: APP_VERSION,
      projectName: safeProjectName,
      surfaceName,
      panelType,
      powerDistro,
      backupSignalLoop,
      includeReinforcementPlate,
      deploymentType,
      wall: { cols, rows, widthM: wallWidthM, heightM: wallHeightM, pixelW: wallPixelW, pixelH: wallPixelH },
      panels: grid,
      patching: { signalPortsUsed, powerPortsUsed },
      stockRows,
      subScreens,
      outputCanvas: { w: outputCanvasW, h: outputCanvasH },
      wholeLayoutCanvasPos: { x: wholeLayoutCanvasX, y: wholeLayoutCanvasY },
      processorModel,
      canvasInputs,
      inputMode,
      wholeCanvasInputId,
      rentmanDateFrom,
      rentmanDateTo,
    };

    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" });
    const url = window.URL.createObjectURL(blob);

    const link = document.createElement("a");
    link.href = url;
    link.setAttribute("download", `${fileSafeProjectName}-${panelType}-${cols}x${rows}-settings.json`);

    document.body.appendChild(link);
    link.click();
    link.remove();

    setTimeout(() => window.URL.revokeObjectURL(url), 1000);
  } catch (err) {
    console.error("JSON download failed", err);
    alert("Settings download failed - check console");
  }
};

  const exportStockCsv = () => {
    try {
      const lines = ["Code,Order Qty", ...visibleStockRows.map((row) => `${row.code},${row.rounded ?? row.required}`)];
      const blob = new Blob([lines.join("\r\n")], { type: "text/csv;charset=utf-8" });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.setAttribute("download", `${fileSafeProjectName}-${panelType}-${cols}x${rows}-stock.csv`);
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => window.URL.revokeObjectURL(url), 1000);
    } catch (err) {
      console.error("Stock CSV download failed", err);
      alert("Stock CSV download failed - check console");
    }
  };

  const exportTestPatternPng = () => {
    try {
      // Shares computeTestPatternLayout with the video/live test pattern
      // (drawTestPattern.ts) instead of keeping a separate duplicate
      // position/label computation - a previous duplicate here silently
      // reintroduced the same "gaps collapse, front-view labels wrong"
      // bugs the shared version had already been fixed for.
      const layout = computeTestPatternLayout({ projectName: safeProjectName, surfaceName, panelType, panels: activePanels });
      const W = Math.max(1, layout.W);
      const H = Math.max(1, layout.H);
      const canvas = document.createElement("canvas");
      canvas.width = W;
      canvas.height = H;
      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("Canvas context unavailable");
      ctx.imageSmoothingEnabled = false;

      ctx.fillStyle = "#000000";
      ctx.fillRect(0, 0, W, H);

      // The PNG ALWAYS renders the front view (what an observer sees standing in
      // front of the finished wall): mirror each panel's native-pixel rect
      // horizontally within the total wall width (mirrorX below mirrors the
      // panel's own shape to match), independent of the on-screen Front/Back
      // toggle.
      const dispRectPx = (cell: Cell): RectMm => {
        const r = layout.panelPixelRects.get(cell.id);
        if (!r) return { x: 0, y: 0, w: 0, h: 0 };
        return { x: W - r.x - r.w, y: r.y, w: r.w, h: r.h };
      };

      layout.activePanels.forEach((cell) => {
        const r = dispRectPx(cell);
        const fill = cell.assignedPort ? PORT_COLORS[(cell.assignedPort - 1) % PORT_COLORS.length] : "#1e293b";
        // No signal/power chain-start ring indicators in the PNG - it's a
        // clean per-panel pixel map, not a patching diagram.
        drawPanelShape(ctx, r.x, r.y, r.w, r.h, cell, fill, "#ffffff", 1, { hatchStep: 24, mirrorX: true });

        const cx = r.x + r.w / 2;
        ctx.fillStyle = "#020617";
        ctx.textAlign = "center";
        ctx.font = `bold ${Math.max(12, Math.floor(r.h * 0.085))}px Arial`;
        ctx.fillText(`↓ ${layout.rowLabel(cell)} → ${layout.colLabel(cell)}`, cx, r.y + r.h * 0.4);
        // Shape symbol only (△/◜/Corner) - no rotate icon, no signal/power port info.
        const variantSymbol = PANEL_VARIANTS[cell.panelVariant ?? "STANDARD"].symbol;
        if (variantSymbol) {
          ctx.font = `bold ${Math.max(14, Math.floor(r.h * 0.12))}px Arial`;
          ctx.fillText(variantSymbol, cx, r.y + r.h - 8);
        }
      });

      // No signal/power cable-routing lines or arrowheads in the PNG: it is a
      // clean front-view pixel map of the wall for the observer / processor.

      const link = document.createElement("a");
      link.href = canvas.toDataURL("image/png");
      link.setAttribute("download", `${fileSafeProjectName}-${fileSafePanelType}-front-test-pattern.png`);
      document.body.appendChild(link);
      link.click();
      link.remove();
    } catch (err) {
      console.error("PNG test pattern failed", err);
      alert("PNG test pattern failed - check console");
    }
  };

  // Open the full-screen, canvas-only live test pattern in its own tab.
  // There's no router, so the project is handed off through localStorage and
  // the new tab (booted with ?testpattern=1, see main.jsx) reads it back and
  // renders TestPatternView - just the LED canvas, no page chrome.
  const openMovingTestPatternTab = () => {
    try {
      const payload = { formatVersion: 1, projectName: safeProjectName, surfaceName, panelType, panels: activePanels };
      localStorage.setItem("ledCablingTestPattern:v1", JSON.stringify(payload));
      window.open(`${location.pathname}?testpattern=1`, "_blank");
    } catch (err) {
      console.error("Moving test pattern failed", err);
      alert("Could not open the moving test pattern - check console");
    }
  };

  // Standalone panel-count calculator, opened with no hand-off data so it
  // always starts at its own neutral 1x1 MG9 default (see QuickLayoutView).
  const openQuickPanelLayoutTab = () => {
    window.open(`${location.pathname}?quicklayout=1`, "_blank");
  };

  const pickVideoMimeType = (): string | null => {
    if (typeof MediaRecorder === "undefined") return null;
    const candidates = ["video/webm;codecs=vp9", "video/webm;codecs=vp8", "video/webm"];
    for (const candidate of candidates) {
      if (MediaRecorder.isTypeSupported && MediaRecorder.isTypeSupported(candidate)) return candidate;
    }
    return null;
  };

  const downloadBlob = (blob: Blob, filename: string) => {
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.setAttribute("download", filename);
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => window.URL.revokeObjectURL(url), 1000);
  };

  // Records exactly one loop of the animated test pattern to a WebM Blob,
  // without ever showing a tab/window - draws into a detached canvas (never
  // added to the DOM) and captures it directly. A generous resolution-scaled
  // bitrate avoids the blocky compression artifacts a codec's low default
  // bitrate would produce on this pattern's large flat colour fields and
  // sharp edges/text. Shared by both the WebM and MP4 downloads - MP4 just
  // pipes this same recording through encodeWebmToMp4 afterwards.
  const recordMovingTestPatternWebm = (): Promise<Blob> | null => {
    if (isRecordingVideo) return null;
    const mimeType = pickVideoMimeType();
    if (!mimeType) {
      alert("This browser can't record video (no WebM/MediaRecorder support). Try Chrome, Edge or Firefox.");
      return null;
    }
    const project: TestPatternProject = { projectName: safeProjectName, surfaceName, panelType, panels: activePanels };
    const layout = computeTestPatternLayout(project);
    if (layout.W <= 0 || layout.H <= 0) {
      alert("No active panels to render a test pattern from.");
      return null;
    }
    const canvas = document.createElement("canvas");
    canvas.width = layout.W;
    canvas.height = layout.H;
    const ctx = canvas.getContext("2d");
    if (!ctx) {
      alert("Could not create a recording canvas - check console");
      return null;
    }

    const loopStart = performance.now();
    const drawId = window.setInterval(() => {
      drawTestPatternFrame(ctx, layout, (performance.now() - loopStart) / 1000);
    }, 1000 / DRAW_FPS);

    // ~6 bits/pixel of total resolution, floor 8Mbps / cap 80Mbps: MediaRecorder's
    // default bitrate is far too low for this pattern's sharp edges and text,
    // producing visible VP9 blocking - this scales generously with wall size
    // instead of leaving every export at one low fixed rate.
    const videoBitsPerSecond = Math.min(80_000_000, Math.max(8_000_000, Math.round(layout.W * layout.H * 6)));
    const stream = canvas.captureStream(DRAW_FPS);
    const recorder = new MediaRecorder(stream, { mimeType, videoBitsPerSecond });
    const chunks: BlobPart[] = [];
    recorder.ondataavailable = (event) => {
      if (event.data.size > 0) chunks.push(event.data);
    };
    setVideoRecordSeconds(0);
    setIsRecordingVideo(true);
    const result = new Promise<Blob>((resolve) => {
      recorder.onstop = () => {
        window.clearInterval(drawId);
        setIsRecordingVideo(false);
        resolve(new Blob(chunks, { type: mimeType }));
      };
    });
    recorder.start();
    setTimeout(() => recorder.stop(), LOOP_SECONDS * 1000);
    return result;
  };

  const downloadMovingTestPatternVideo = () => {
    const recording = recordMovingTestPatternWebm();
    if (!recording) return;
    recording.then((blob) => downloadBlob(blob, `${fileSafeProjectName}-front-test-pattern.webm`));
  };

  const downloadMovingTestPatternMp4 = () => {
    if (isEncodingMp4) return;
    const recording = recordMovingTestPatternWebm();
    if (!recording) return;
    recording.then(async (blob) => {
      setIsEncodingMp4(true);
      setMp4EncodeProgress(0);
      try {
        const { encodeWebmToMp4 } = await import("./testPattern/mp4Encode");
        const mp4Blob = await encodeWebmToMp4(blob, setMp4EncodeProgress);
        downloadBlob(mp4Blob, `${fileSafeProjectName}-front-test-pattern.mp4`);
      } catch (err) {
        console.error("MP4 encode failed", err);
        alert("MP4 encoding failed - check console. You can still use Download Moving Test Pattern (WebM).");
      } finally {
        setIsEncodingMp4(false);
      }
    });
  };

  const openJson = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(String(ev.target?.result || "{}")) as OpenJsonPayload;
        const nextCols = Math.max(1, Number(data.wall?.cols || cols));
        const nextRows = Math.max(1, Number(data.wall?.rows || rows));
        const rawGrid = data.patching?.grid;
        let nextPanels: Cell[];
        if (Array.isArray(data.panels)) {
          // formatVersion 2: free mm panel list.
          nextPanels = normalizePanels(data.panels);
        } else if (Array.isArray(rawGrid)) {
          // formatVersion 1 grid. Pre-panelType files were one type per wall
          // (MT cells there are a full 1m wide); typed grids carry mtTail pairs
          // which gridCellsToPanels absorbs into single MT records.
          const legacyAllType = isLegacyUntypedGrid(rawGrid) && data.panelType === "MT" ? "MT" : null;
          nextPanels = gridCellsToPanels(rawGrid, legacyAllType);
        } else {
          nextPanels = makeGridPanels(nextCols, nextRows);
        }

        if (data.projectName) setProjectName(data.projectName);
        setSurfaceName(data.surfaceName ?? "");
        if (data.panelType && PANEL_TYPES[data.panelType]) setPanelType(data.panelType);
        if (data.powerDistro && POWER_DISTROS[data.powerDistro]) setPowerDistro(data.powerDistro);
        setBackupSignalLoop(data.backupSignalLoop ?? true);
        setIncludeReinforcementPlate(data.includeReinforcementPlate ?? false);
        setDeploymentType(data.deploymentType ?? "");

        setCols(nextCols);
        setRows(nextRows);
        setDraftCols(String(nextCols));
        setDraftRows(String(nextRows));
        setGrid(nextPanels);
        setSubScreens(Array.isArray(data.subScreens) ? normalizeSubScreens(data.subScreens) : []);
        // Never resume mid-edit of a stale sub-screen from a previous session.
        setActiveSubScreenId(null);
        setOutputCanvasW(Number(data.outputCanvas?.w) || 1920);
        setOutputCanvasH(Number(data.outputCanvas?.h) || 1080);
        setWholeLayoutCanvasX(Number(data.wholeLayoutCanvasPos?.x) || 0);
        setWholeLayoutCanvasY(Number(data.wholeLayoutCanvasPos?.y) || 0);
        // formatVersion 4: older projects have neither field - default to
        // "no processor selected" / no input assignments rather than
        // guessing, so nothing is silently exported for a wall the user
        // never configured a processor for.
        setProcessorModel(data.processorModel && PROCESSOR_SPECS[data.processorModel] ? data.processorModel : "");
        setCanvasInputs(data.canvasInputs && typeof data.canvasInputs === "object" ? data.canvasInputs : {});
        // formatVersion 5: older projects have neither field - default to
        // the original "perEntry" behavior / no whole-canvas input.
        setInputMode(data.inputMode === "whole" ? "whole" : "perEntry");
        setWholeCanvasInputId(typeof data.wholeCanvasInputId === "number" ? data.wholeCanvasInputId : null);
        // formatVersion 6: older projects have neither field - default to
        // no date range set (Rentman availability just stays off).
        setRentmanDateFrom(typeof data.rentmanDateFrom === "string" ? data.rentmanDateFrom : "");
        setRentmanDateTo(typeof data.rentmanDateTo === "string" ? data.rentmanDateTo : "");
        setSelectedId(null);
        setSelectedCells(new Set());
        setUndoStack([]);
        setRedoStack([]);
      } catch {
        window.alert("Invalid JSON file");
      } finally {
        if (fileInputRef.current) fileInputRef.current.value = "";
      }
    };
    reader.readAsText(file);
  };

  // --- Import from the YES TECH Layout Tool --------------------------------
  // Read + validate the file, then show a preview modal before touching the
  // current project (the user can cancel, replace, or open a new project).
  const handleImportFile = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const result = parseYesTechLayout(String(ev.target?.result || ""));
      setImportPreview(result);
      if (importInputRef.current) importInputRef.current.value = "";
    };
    reader.readAsText(file);
  };

  const hasUnsavedWork = grid.some((cell) => isActiveCell(cell) && (cell.assignedPort || cell.assignedPowerPort));

  const applyImport = (result: ImportResult, mode: "replace" | "new") => {
    // Both modes replace the on-screen project; "new" also resets the name to
    // the imported one. The original source file is never modified.
    //
    // The Creative Layout Tool designs what the audience sees (the FRONT of the
    // wall). This app's stored layout is the back/working (wiring) view and the
    // Front View toggle mirrors it horizontally. So we store the horizontal
    // mirror of the imported design and switch on Front View: the Front View
    // then reproduces the original Creative layout exactly (positions, shapes
    // and rotations), while the back view shows the correct wiring mirror.
    const IMPORT_W = 500; // MG9-family footprint (mm); the only types imported.
    let minX = Infinity;
    let maxX = -Infinity;
    result.panels.forEach((p) => {
      minX = Math.min(minX, p.x);
      maxX = Math.max(maxX, p.x + IMPORT_W);
    });
    // Rotation under a horizontal mirror, per shape, so the mirrored (front)
    // render matches the source orientation exactly, at any angle (not just
    // multiples of 90). Kept as a fractional degree - rounding here would
    // silently snap custom import angles back to whole numbers.
    const mirrorRotation = (variant: string, rotation: number) => {
      const r = ((rotation % 360) + 360) % 360;
      if (variant === "TRIANGLE") return (270 - r + 360) % 360;
      if (variant === "CURVED") return (90 - r + 360) % 360;
      return (360 - r) % 360; // STANDARD: base square is vertical-axis symmetric, so mirroring negates the angle.
    };
    const panels: Cell[] = result.panels.map((p) => ({
      id: newCellId(),
      x: minX + maxX - (p.x + IMPORT_W),
      y: p.y,
      assignedPort: null,
      sequence: null,
      assignedPowerPort: null,
      powerSequence: null,
      powerManual: false,
      isRemoved: false,
      panelVariant: p.panelVariant,
      rotation: mirrorRotation(p.panelVariant, p.rotation),
      panelType: p.panelType,
      subScreenId: null,
    }));
    setProjectName(mode === "new" ? result.projectName : result.projectName || projectName);
    setPanelType("MG9");
    setGrid(panels);
    // Import replaces the whole project's panels wholesale - any existing
    // sub-screens no longer have valid members, so start clean rather than
    // leaving stale/empty sub-screens behind.
    setSubScreens([]);
    setActiveSubScreenId(null);
    setSelectedId(null);
    setSelectedCells(new Set());
    setUndoStack([]);
    setRedoStack([]);
    setEditMode("patch");
    setPatchMode("signal");
    setIsFlippedView(true);
    setOverlapNotice(null);
    setImportPreview(null);
  };

  const generatePdf = async () => {
  try {
    const jsPDF = (await import("jspdf")).default;
    const pdf = new jsPDF({ unit: "mm", format: "a4", orientation: "landscape", compress: true });
    const printedAt = new Date().toLocaleString();
    const usedSignalPorts = signalPorts.filter((port) => signalPortStats[port.id].panels > 0);
    const usedPowerPorts = powerPorts.filter((port) => powerPortStats[port.id].panels > 0);

    const addPdfFooters = () => {
      const totalPages = pdf.getNumberOfPages();
      for (let pageNo = 1; pageNo <= totalPages; pageNo += 1) {
        pdf.setPage(pageNo);
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(9);
        pdf.setTextColor(71, 85, 105);
        pdf.text(`Printed ${printedAt}`, 10, pageHeight - 6);
        pdf.text(`Page ${pageNo} of ${totalPages}`, pageWidth - 10, pageHeight - 6, { align: "right" });
        pdf.setTextColor(0, 0, 0);
      }
    };

    const drawInfoBox = (title: string, lines: string[], x: number, y: number, w: number, h: number) => {
      pdf.setDrawColor(148, 163, 184);
      pdf.setFillColor(248, 250, 252);
      pdf.roundedRect(x, y, w, h, 2, 2, "FD");
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(11);
      pdf.text(title, x + 3, y + 6);
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(9);
      let lineY = y + 12;
      lines.forEach((line) => {
        const wrapped = pdf.splitTextToSize(String(line), w - 6);
        wrapped.forEach((entry: string) => {
          if (lineY <= y + h - 3) pdf.text(entry, x + 3, lineY);
          lineY += 4.2;
        });
      });
    };

    const drawLayoutPage = (canvas: HTMLCanvasElement, viewLabel: string) => {
      pdf.addPage("a4", "landscape");
      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(18);
      pdf.text(`${safeProjectName} - Panel Layout - ${viewLabel}`, 10, 12);

      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(10);
      pdf.text(`Project name: ${safeProjectName}`, 10, 20);
      pdf.text(`Panel type: ${panelTypeSummary}`, 10, 26);
      pdf.text(`Power distro: ${distro.label}`, 10, 32);
      pdf.text(`Panels: ${totalPanels} active across ${panelBands.length} row band${panelBands.length === 1 ? "" : "s"}`, 10, 38);

      pdf.text(`${wallSizeLabel}: ${formatMeters(wallWidthM)}m x ${formatMeters(wallHeightM)}m`, 105, 20);
      pdf.text(`Total weight: ${totalWeight.toFixed(1)} kg`, 105, 26);
      pdf.text(wallResolutionSummaryLines[0], 105, 32);
      pdf.text(wallResolutionSummaryLines[1], 105, 38);
      pdf.text(wallResolutionSummaryLines[2], 105, 44);

      const usableWidth = pageWidth - 20;
      const usableHeight = pageHeight - 58;
      const layoutRatio = canvas.width / canvas.height;
      let drawWidth = usableWidth;
      let drawHeight = drawWidth / layoutRatio;
      if (drawHeight > usableHeight) {
        drawHeight = usableHeight;
        drawWidth = drawHeight * layoutRatio;
      }
      pdf.addImage(
        canvas.toDataURL("image/png"),
        "PNG",
        10 + (usableWidth - drawWidth) / 2,
        50 + (usableHeight - drawHeight) / 2,
        drawWidth,
        drawHeight,
        undefined,
        "SLOW",
      );
    };

    // Two column layouts: the plain 5-numeric-column one used whenever
    // Rentman availability isn't in play (byte-identical to before this
    // feature existed), and a 6-column one (Item's wrap narrowed to make
    // room) only when hasLiveAvailableColumn is actually true - so a
    // project that isn't using Rentman renders this table exactly as it
    // always has.
    const stockCols = hasLiveAvailableColumn
      ? { itemWrap: 104, required: 160, spare: 182, rounded: 208, stock: 232, net: 256, available: 284 }
      : { itemWrap: 128, required: 174, spare: 198, rounded: 226, stock: 252, net: 282, available: null };
    const drawStockTable = (startIndex: number, startY: number, maxY: number) => {
      let y = startY;
      const drawHeader = () => {
        pdf.setFillColor(226, 232, 240);
        pdf.rect(10, y - 5, 274, 7, "F");
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(8);
        pdf.text("Code", 12, y);
        pdf.text("Item", 34, y);
        pdf.text("Required", stockCols.required, y, { align: "right" });
        pdf.text("Spare", stockCols.spare, y, { align: "right" });
        pdf.text("Rounded + Spare", stockCols.rounded, y, { align: "right" });
        pdf.text("Stock", stockCols.stock, y, { align: "right" });
        pdf.text("Net", stockCols.net, y, { align: "right" });
        if (stockCols.available) pdf.text("Available", stockCols.available, y, { align: "right" });
        y += 6;
        pdf.setFont("helvetica", "normal");
      };
      drawHeader();
      for (let index = startIndex; index < visibleStockRows.length; index += 1) {
        const row = visibleStockRows[index];
        if (y > maxY) return index;
        if (row.net < 0) {
          pdf.setFillColor(254, 226, 226);
          pdf.rect(10, y - 4.5, 274, 6.2, "F");
        }
        pdf.text(String(row.code), 12, y);
        pdf.text(pdf.splitTextToSize(row.name, stockCols.itemWrap)[0], 34, y);
        pdf.text(formatNumber(row.required), stockCols.required, y, { align: "right" });
        pdf.text(formatNumber(row.spare ?? 0), stockCols.spare, y, { align: "right" });
        pdf.text(formatNumber(row.rounded ?? row.required), stockCols.rounded, y, { align: "right" });
        pdf.text(formatNumber(row.stock), stockCols.stock, y, { align: "right" });
        pdf.text(formatNumber(row.net), stockCols.net, y, { align: "right" });
        if (stockCols.available) pdf.text(typeof row.available === "number" ? formatNumber(row.available) : "-", stockCols.available, y, { align: "right" });
        y += 6;
      }
      return visibleStockRows.length;
    };

    const drawStockPage = (startIndex = 0) => {
      pdf.addPage("a4", "landscape");
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(16);
      pdf.text(`${safeProjectName} - Stock Summary`, 10, 12);
      let nextIndex = drawStockTable(startIndex, 22, 190);
      while (nextIndex < visibleStockRows.length) {
        pdf.addPage("a4", "landscape");
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(16);
        pdf.text(`${safeProjectName} - Stock Summary continued`, 10, 12);
        nextIndex = drawStockTable(nextIndex, 22, 190);
      }
    };

    // Per-sub-screen summary: always re-derives each sub-screen's own stats
    // from the FULL grid (not the live scoped activePanels, which only ever
    // reflects one sub-screen - or none - at a time), so the page covers
    // every sub-screen regardless of which one is currently being edited.
    const drawSubScreensTable = (startIndex: number, startY: number, maxY: number) => {
      let y = startY;
      const drawHeader = () => {
        pdf.setFillColor(226, 232, 240);
        pdf.rect(10, y - 5, 274, 7, "F");
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(8);
        pdf.text("Name", 12, y);
        pdf.text("Resolution", 90, y);
        pdf.text("Size (m)", 130, y);
        pdf.text("Canvas X,Y", 168, y);
        pdf.text("Right,Bottom", 210, y);
        pdf.text("Panels", 255, y, { align: "right" });
        pdf.text("Signal Ports", 282, y, { align: "right" });
        y += 6;
        pdf.setFont("helvetica", "normal");
      };
      drawHeader();
      for (let index = startIndex; index < subScreens.length; index += 1) {
        const screen = subScreens[index];
        if (y > maxY) return index;
        const bbox = subScreenBBoxOf(grid, screen.id, cellRect);
        const resolution = subScreenResolutionOf(grid, screen.id);
        const panelCount = subScreenPanelCount(grid, screen.id);
        const portsUsed = new Set(
          grid
            .filter((c) => c.subScreenId === screen.id && !c.isRemoved && c.assignedPort)
            .map((c) => c.assignedPort),
        ).size;
        pdf.text(pdf.splitTextToSize(screen.name, 74)[0], 12, y);
        pdf.text(`${resolution.w} x ${resolution.h}`, 90, y);
        pdf.text(`${(bbox.w / 1000).toFixed(2)} x ${(bbox.h / 1000).toFixed(2)}`, 130, y);
        pdf.text(`${screen.canvasX}, ${screen.canvasY}`, 168, y);
        pdf.text(`${screen.canvasX + resolution.w}, ${screen.canvasY + resolution.h}`, 210, y);
        pdf.text(formatNumber(panelCount), 255, y, { align: "right" });
        pdf.text(formatNumber(portsUsed), 282, y, { align: "right" });
        y += 6;
      }
      return subScreens.length;
    };

    const drawSubScreensSummaryPage = () => {
      pdf.addPage("a4", "landscape");
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(16);
      pdf.text(`${safeProjectName} - Sub-Screens`, 10, 12);
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(10);
      pdf.text(`Output canvas: ${outputCanvasW} x ${outputCanvasH}px`, 10, 18);
      let nextIndex = drawSubScreensTable(0, 28, 190);
      while (nextIndex < subScreens.length) {
        pdf.addPage("a4", "landscape");
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(16);
        pdf.text(`${safeProjectName} - Sub-Screens continued`, 10, 12);
        nextIndex = drawSubScreensTable(nextIndex, 22, 190);
      }
    };

    // Full per-port breakdown, as two side-by-side tables. A single fixed
    // page (no pagination) is always enough: signal ports are hard-capped at
    // 20 (the largest supported NovaStar processor) and power outputs at 18
    // (the 63A distro's port count) - both comfortably fit in one column of
    // rows well within the page height.
    const drawPortsInUsePage = () => {
      if (!usedSignalPorts.length && !usedPowerPorts.length) return;
      pdf.addPage("a4", "landscape");
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(16);
      pdf.text(`${safeProjectName} - Signal & Power Ports In Use`, 10, 12);

      const colW = 133;
      type PortColumn = { label: string; x: number; align?: "right" };
      const drawPortTable = (x: number, title: string, columns: PortColumn[], rows: string[][], emptyLabel: string) => {
        let y = 22;
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(12);
        pdf.text(title, x, y);
        y += 6;
        pdf.setFillColor(226, 232, 240);
        pdf.rect(x, y - 5, colW, 7, "F");
        pdf.setFontSize(8);
        columns.forEach((col) => pdf.text(col.label, x + col.x, y, col.align ? { align: col.align } : undefined));
        y += 6;
        pdf.setFont("helvetica", "normal");
        if (!rows.length) {
          pdf.text(emptyLabel, x, y);
          return;
        }
        rows.forEach((cells) => {
          columns.forEach((col, i) => pdf.text(cells[i], x + col.x, y, col.align ? { align: col.align } : undefined));
          y += 6;
        });
      };

      drawPortTable(
        10,
        `Signal Ports (${usedSignalPorts.length} of ${signalPorts.length} in use)`,
        [
          { label: "Port", x: 0 },
          { label: "Panels", x: 40, align: "right" },
          { label: "Chain (first -> last)", x: 48 },
        ],
        usedSignalPorts.map((port) => {
          const stat = signalPortStats[port.id];
          return [port.name, formatNumber(stat.panels), stat.firstKey ? `${stat.firstKey} -> ${stat.lastKey}` : "-"];
        }),
        "No signal ports in use.",
      );

      drawPortTable(
        154,
        `Power Outputs (${usedPowerPorts.length} of ${powerPorts.length} in use)`,
        [
          { label: "Plug", x: 0 },
          { label: "Panels", x: 40, align: "right" },
          { label: "Max W / A", x: 75, align: "right" },
          { label: "Phase", x: 110 },
        ],
        usedPowerPorts.map((port) => {
          const stat = powerPortStats[port.id];
          return [port.name, formatNumber(stat.panels), `${formatNumber(stat.maxWatts)}W / ${formatNumber(stat.maxAmps, 2)}A`, stat.phase || "-"];
        }),
        "No power outputs in use.",
      );
    };

    // Spare-panel breakdown: one table per surface (sub-screen, "Unassigned"
    // if any panels aren't in one, or just "Whole Layout" with no
    // sub-screens), each type bucket its own row (see sparePanelSummary) -
    // paginates itself since the number of surfaces is open-ended.
    const drawSparePanelsPage = () => {
      if (!sparePanelSummary.surfaceRows.length) return;
      const colX = { type: 10, used: 140, spare: 178, rounded: 225 };
      let y = 0;
      const startPage = (continued: boolean) => {
        pdf.addPage("a4", "landscape");
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(16);
        pdf.text(`${safeProjectName} - Spare Panels by Surface${continued ? " (continued)" : ""}`, 10, 12);
        y = 24;
        if (!continued) {
          pdf.setFont("helvetica", "normal");
          pdf.setFontSize(9);
          pdf.setTextColor(100, 116, 139);
          pdf.text(
            `Spare = ${formatNumber(PANEL_TYPES.MG9.defaults.spareRatio * 100, 0)}% of panels used, per panel type, rounded up to full boxes where that type is boxed (shaped panels are one-way pieces bought individually, not boxed).`,
            10,
            18,
          );
          pdf.setTextColor(15, 23, 42);
        }
      };
      const ensureRoom = (rowsNeeded: number) => {
        if (y === 0 || y + rowsNeeded * 5.5 > 195) startPage(y !== 0);
      };

      sparePanelSummary.surfaceRows.forEach((surface) => {
        ensureRoom(surface.bucketRows.length + 3);
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(12);
        pdf.text(surface.name, colX.type, y);
        y += 6;
        pdf.setFillColor(226, 232, 240);
        pdf.rect(10, y - 5, 274, 7, "F");
        pdf.setFontSize(8);
        pdf.text("Panel Type", colX.type, y);
        pdf.text("Used", colX.used, y, { align: "right" });
        pdf.text(`Spare (${formatNumber(PANEL_TYPES.MG9.defaults.spareRatio * 100, 0)}%)`, colX.spare, y, { align: "right" });
        pdf.text("Rounded + Spare", colX.rounded, y, { align: "right" });
        y += 6;
        pdf.setFont("helvetica", "normal");
        surface.bucketRows.forEach((row) => {
          pdf.text(row.label, colX.type, y);
          pdf.text(formatNumber(row.used), colX.used, y, { align: "right" });
          pdf.text(formatNumber(row.spare), colX.spare, y, { align: "right" });
          pdf.text(formatNumber(row.rounded), colX.rounded, y, { align: "right" });
          y += 5.5;
        });
        pdf.setFont("helvetica", "bold");
        pdf.text("Subtotal", colX.type, y);
        pdf.text(formatNumber(surface.subtotal.used), colX.used, y, { align: "right" });
        pdf.text(formatNumber(surface.subtotal.spare), colX.spare, y, { align: "right" });
        pdf.text(formatNumber(surface.subtotal.rounded), colX.rounded, y, { align: "right" });
        pdf.setFont("helvetica", "normal");
        y += 10;
      });

      if (sparePanelSummary.multiSurface) {
        ensureRoom(2);
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(11);
        pdf.setTextColor(3, 105, 161);
        pdf.text(
          `Grand total: ${formatNumber(sparePanelSummary.grandTotal.used)} used, ${formatNumber(sparePanelSummary.grandTotal.spare)} spare, ${formatNumber(sparePanelSummary.grandTotal.rounded)} incl. spare`,
          colX.type,
          y,
        );
        pdf.setTextColor(15, 23, 42);
        pdf.setFont("helvetica", "normal");
      }
    };

    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(18);
    pdf.text(safeProjectName, 10, 12);
    pdf.setFont("helvetica", "normal");
    pdf.setFontSize(10);
    pdf.text(`Printed ${printedAt}`, 10, 18);

    drawInfoBox("Wall", [
      `Panel type: ${panelTypeSummary}`,
      `Power distro: ${distro.label}`,
      `Panels: ${totalPanels} active across ${panelBands.length} row band${panelBands.length === 1 ? "" : "s"}`,
      `${wallSizeLabel}: ${formatMeters(wallWidthM)}m x ${formatMeters(wallHeightM)}m`,
      ...wallResolutionSummaryLines,
    ], 10, 24, 66, 48);

    drawInfoBox("Power", [
      `Max draw: ${formatNumber(totalPowerMaxW)} W / ${formatNumber(totalPowerMaxA, 2)} A`,
      `Average draw: ${formatNumber(totalPowerAvgW)} W / ${formatNumber(totalPowerAvgA, 2)} A`,
      `Circuits used: ${circuitsUsedMax}`,
      `Per outlet: ${formatNumber(powerPerCircuitMaxW)} W / ${formatNumber(powerPerCircuitMaxA, 2)} A`,
      `Outlet limit: ${safePanelsPerPowerOutlet} panels`,
      `Unassigned power panels: ${unassignedPowerPanels}`,
    ], 80, 24, 66, 48);

    drawInfoBox("Weight + Output", [
      `Total weight: ${totalWeight.toFixed(1)} kg`,
      `VX1000 use: ${formatNumber(vx1000Percent, 1)}%`,
      `VX2000 use: ${formatNumber(vx2000Percent, 1)}%`,
      `Best output: ${bestResolution ? `${bestResolution[0]} x ${bestResolution[1]}` : "None in preset list"}`,
      `Signal limit: ${safePanelsPerSignalPort} panels / ${formatNumber(signalPortPixels)} px`,
      `Active support span: ${activeColsCount} cols x ${activeRowsCount} rows`,
    ], 150, 24, 66, 48);

    drawInfoBox("Deployment + Stock", [
      `Spare panels: ${sparePanels}`,
      `Panels incl. spare: ${totalPanelsWithSpare}`,
      `Boxes: ${boxCount} (${boxSparePanels} additional spare)`,
      `Backup signal loop: ${backupSignalLoop ? `Yes, effective signal ports ${effectiveSignalPortsUsed}` : "No"}`,
      `Reinforcement plate: ${includeReinforcementPlate ? "Yes" : "No"}`,
      `Deployment type: ${deploymentType || "Not selected"}`,
      ...(deploymentWarning ? [`Warning: ${deploymentWarning}`] : []),
    ], 220, 24, 66, 48);

    drawInfoBox("Phase Load", Object.entries(phaseStats).map(([phase, stat]) =>
      `Phase ${phase.replace("P", "")}: ${formatNumber(stat.maxWatts)} W / ${formatNumber(stat.maxAmps, 2)} A (${formatNumber(stat.utilisation, 1)}%)`
    ), 10, 78, 92, 44);

    // Full per-port detail lives on its own page (drawPortsInUsePage below) -
    // these boxes are just a compact count, since the old approach (one line
    // per port crammed into a small fixed-height box) silently dropped any
    // ports past ~7 with no indication once a project used more than that
    // (easy to hit - VX2000 Pro alone offers up to 20 signal ports).
    drawInfoBox("Signal Ports In Use", [
      `${usedSignalPorts.length} of ${signalPorts.length} ports in use`,
      "See Signal & Power Ports page for full detail.",
    ], 106, 78, 88, 44);

    drawInfoBox("Power Outputs In Use", [
      `${usedPowerPorts.length} of ${powerPorts.length} outputs in use`,
      "See Signal & Power Ports page for full detail.",
    ], 198, 78, 88, 44);

    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(12);
    pdf.text("Stock Summary", 10, 128);
    const nextStockIndex = drawStockTable(0, 138, 190);

    const backLayoutCanvas = buildLayoutCanvas(false, "Back View");
    const frontLayoutCanvas = buildLayoutCanvas(true, "Front View");
    if (nextStockIndex < visibleStockRows.length) drawStockPage(nextStockIndex);
    drawSparePanelsPage();
    drawPortsInUsePage();
    if (subScreens.length > 0) drawSubScreensSummaryPage();
    drawLayoutPage(backLayoutCanvas, "Back View");
    drawLayoutPage(frontLayoutCanvas, "Front View");
    addPdfFooters();
    pdf.save(`${fileSafeProjectName}-${fileSafePanelType}-${cols}x${rows}.pdf`);
  } catch (err) {
    console.error("PDF failed", err);
    alert("PDF failed - check console");
  }
};

  const performApplyGridSize = () => {
    const nextCols = Number.parseInt(draftCols, 10);
    const nextRows = Number.parseInt(draftRows, 10);
    if (!Number.isFinite(nextCols) || !Number.isFinite(nextRows) || nextCols < 1 || nextRows < 1) return;

    pushUndoSnapshot();
    setCols(nextCols);
    setRows(nextRows);
    setGrid(makeGridPanels(nextCols, nextRows, panelType));
    // Regenerating the grid replaces every panel - any existing sub-screens
    // no longer have valid members, so start clean.
    setSubScreens([]);
    setActiveSubScreenId(null);
    setSelectedId(null);
    setSelectedCells(new Set());
    setDragVisited(new Set());
    setIsDragging(false);
  };

  const applyGridSize = () => {
    const nextCols = Number.parseInt(draftCols, 10);
    const nextRows = Number.parseInt(draftRows, 10);
    if (!Number.isFinite(nextCols) || !Number.isFinite(nextRows) || nextCols < 1 || nextRows < 1) return;

    // Regenerating the grid discards every existing panel - warn first
    // rather than silently wiping out a layout the user has already built.
    if (grid.some(isActiveCell)) {
      setShowGridSizeConfirm(true);
      return;
    }
    performApplyGridSize();
  };

  // --- Workspace display geometry ------------------------------------------
  // Display space = workspace mm, mirrored horizontally inside the wall bbox
  // when the front view is shown. The workspace origin is the bbox corner
  // minus padding and stays fixed during a drag gesture.
  const WORKSPACE_PAD_MM = 300;
  const workspaceOrigin = { x: wallBBox.x - WORKSPACE_PAD_MM, y: wallBBox.y - WORKSPACE_PAD_MM };
  const workspaceSizeMm = { w: wallBBox.w + WORKSPACE_PAD_MM * 2, h: wallBBox.h + WORKSPACE_PAD_MM * 2 };
  const mmToPx = (mm: number) => mm * pxPerMm;
  // Sets zoom so the FULL workspace (every active panel, including any
  // imported far outside the default view) fits inside the scrollable
  // viewport's current visible size - the direct fix for "some panels are
  // outside the accessible layout area" after importing a wide/tall project.
  const fitToView = () => {
    const viewport = workspaceViewportRef.current;
    if (!viewport || workspaceSizeMm.w <= 0 || workspaceSizeMm.h <= 0) return;
    // Viewport padding (p-4 = 16px each side) eats into the usable area.
    const availW = Math.max(1, viewport.clientWidth - 32);
    const availH = Math.max(1, viewport.clientHeight - 32);
    const pxPerMmAtZoom1 = CELL_SIZE / MODULE_MM;
    const fitZoom = Math.min(availW / (workspaceSizeMm.w * pxPerMmAtZoom1), availH / (workspaceSizeMm.h * pxPerMmAtZoom1));
    setZoom(Math.min(2, Math.max(0.02, fitZoom)));
    // Reset scroll to the origin so the whole (now-resized) workspace is
    // actually in view, rather than leaving a stale scroll position that
    // could still clip part of it after the content shrinks/grows.
    viewport.scrollLeft = 0;
    viewport.scrollTop = 0;
  };
  const displayRectOf = (cell: Cell): RectMm => {
    const rect = isFlippedView ? mirrorRectX(cellRect(cell), wallBBox) : cellRect(cell);
    if (moveDrag && moveDrag.ids.includes(cell.id)) {
      return { ...rect, x: rect.x + moveDrag.dx, y: rect.y + moveDrag.dy };
    }
    return rect;
  };
  const rectToPx = (rect: RectMm) => ({
    x: mmToPx(rect.x - workspaceOrigin.x),
    y: mmToPx(rect.y - workspaceOrigin.y),
    w: mmToPx(rect.w),
    h: mmToPx(rect.h),
  });
  const eventToDisplayMm = (event: React.MouseEvent): { x: number; y: number } | null => {
    const host = workspaceRef.current;
    if (!host) return null;
    const bounds = host.getBoundingClientRect();
    return {
      x: (event.clientX - bounds.left) / pxPerMm + workspaceOrigin.x,
      y: (event.clientY - bounds.top) / pxPerMm + workspaceOrigin.y,
    };
  };

  const rectsIntersect = (a: RectMm, b: RectMm) =>
    a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;

  const updateMarqueeSelection = (a: { x: number; y: number }, b: { x: number; y: number }) => {
    const marquee: RectMm = {
      x: Math.min(a.x, b.x),
      y: Math.min(a.y, b.y),
      w: Math.abs(a.x - b.x),
      h: Math.abs(a.y - b.y),
    };
    const ids = new Set<string>();
    grid.forEach((cell) => {
      if (isPanelDimmed(cell)) return;
      const rect = isFlippedView ? mirrorRectX(cellRect(cell), wallBBox) : cellRect(cell);
      if (rectsIntersect(marquee, rect)) ids.add(cell.id);
    });
    setSelectedCells(ids);
  };

  // Resolve a drag (display-space dx/dy) into the final true-space snapped delta,
  // reused by the live guide preview and the committed move.
  const resolveMoveSnap = (dragDx: number, dragDy: number, ids: string[]) => {
    const trueDx = isFlippedView ? -dragDx : dragDx;
    const trueDy = dragDy;
    const movingIds = new Set(ids);
    const moving = grid.filter((p) => movingIds.has(p.id) && !p.isRemoved);
    // Anchor-based snap (ported from the layout tool): shift the moved panels so
    // a connector anchor meets a stationary panel's anchor. Shaped panels only
    // snap on their real edges, so incompatible edges never join.
    const firstRect = moving[0] ? { ...cellRect(moving[0]), x: cellRect(moving[0]).x + trueDx, y: cellRect(moving[0]).y + trueDy } : null;
    const movingGeoms = moving.map((p) => {
      const g = cellGeom(p);
      return { ...g, cx: g.cx + trueDx, cy: g.cy + trueDy };
    });
    const otherPanels = grid.filter((p) => !movingIds.has(p.id) && !p.isRemoved);
    const otherGeoms = otherPanels.map(cellGeom);
    const snap = computeAnchorSnapDelta(movingGeoms, otherGeoms, snapEnabled, firstRect);
    return { movingIds, trueDx: trueDx + snap.dx, trueDy: trueDy + snap.dy, snappedTo: snap.snappedTo, otherPanels };
  };

  const commitMoveDrag = () => {
    const drag = moveDrag;
    setMoveDrag(null);
    setSnapGuide(null);
    if (!drag) return;
    if (Math.abs(drag.dx) < 1 && Math.abs(drag.dy) < 1) return;
    const { movingIds, trueDx: dx, trueDy: dy } = resolveMoveSnap(drag.dx, drag.dy, drag.ids);
    const nextPanels = grid.map((p) => (movingIds.has(p.id) ? { ...p, x: p.x + dx, y: p.y + dy } : { ...p }));
    const overlaps = findOverlaps(nextPanels, cellRect);
    if (overlaps.length && !allowOverlaps) {
      setOverlapNotice(
        `Move cancelled: it would overlap ${overlaps.length} panel pair${overlaps.length === 1 ? "" : "s"}. Enable "Allow overlaps" to override.`,
      );
      return;
    }
    setOverlapNotice(
      overlaps.length ? `${overlaps.length} overlapping panel pair${overlaps.length === 1 ? "" : "s"} kept by override.` : null,
    );
    commitGridUpdate(() => nextPanels);
  };

  const onWorkspaceMouseMove = (event: React.MouseEvent) => {
    if (moveDrag) {
      const mm = eventToDisplayMm(event);
      if (!mm) return;
      const dx = mm.x - moveDrag.startX;
      const dy = mm.y - moveDrag.startY;
      setMoveDrag((prev) => (prev ? { ...prev, dx, dy } : prev));
      // Live snap/join guide: outline where the panels will land and highlight
      // the edges they will join along (only when a real edge-snap is found).
      const { movingIds, trueDx, trueDy, snappedTo, otherPanels } = resolveMoveSnap(dx, dy, moveDrag.ids);
      if (snappedTo === "panel") {
        const ghostTrue = grid
          .filter((p) => movingIds.has(p.id) && !p.isRemoved)
          .map((p) => {
            const r = cellRect(p);
            return { ...r, x: r.x + trueDx, y: r.y + trueDy };
          });
        const otherRects = otherPanels.map(cellRect);
        const edges: { x1: number; y1: number; x2: number; y2: number }[] = [];
        ghostTrue.forEach((g) => {
          otherRects.forEach((o) => {
            if (!rectsJoined(g, o)) return;
            const gd = rectToPx(isFlippedView ? mirrorRectX(g, wallBBox) : g);
            const od = rectToPx(isFlippedView ? mirrorRectX(o, wallBBox) : o);
            // Shared vertical edge?
            const shX = Math.min(gd.x + gd.w, od.x + od.w) - Math.max(gd.x, od.x);
            if (Math.abs(gd.x - (od.x + od.w)) < 3 || Math.abs(od.x - (gd.x + gd.w)) < 3) {
              const ex = Math.abs(gd.x - (od.x + od.w)) < 3 ? gd.x : gd.x + gd.w;
              const y1 = Math.max(gd.y, od.y);
              const y2 = Math.min(gd.y + gd.h, od.y + od.h);
              edges.push({ x1: ex, y1, x2: ex, y2 });
            } else if (shX > 0) {
              const ey = Math.abs(gd.y - (od.y + od.h)) < 3 ? gd.y : gd.y + gd.h;
              const x1 = Math.max(gd.x, od.x);
              const x2 = Math.min(gd.x + gd.w, od.x + od.w);
              edges.push({ x1, y1: ey, x2, y2: ey });
            }
          });
        });
        setSnapGuide({
          ghosts: ghostTrue.map((g) => rectToPx(isFlippedView ? mirrorRectX(g, wallBBox) : g)),
          edges,
        });
      } else {
        setSnapGuide(null);
      }
      return;
    }
    if (editMode === "select" && isSelectingPanels && selectionStart) {
      const mm = eventToDisplayMm(event);
      if (!mm) return;
      setSelectionEnd(mm);
      updateMarqueeSelection(selectionStart, mm);
    }
  };

  const onWorkspaceMouseDown = (event: React.MouseEvent) => {
    // Marquee start on empty workspace (panel handlers stop propagation).
    if (editMode !== "select") return;
    const mm = eventToDisplayMm(event);
    if (!mm) return;
    setSelectionStart(mm);
    setSelectionEnd(mm);
    setIsSelectingPanels(true);
    if (!event.shiftKey) {
      setSelectedCells(new Set());
      setSelectedId(null);
    }
  };

  const onWorkspaceMouseUp = () => {
    if (moveDrag) commitMoveDrag();
    setSnapGuide(null);
    setIsSelectingPanels(false);
    setSelectionStart(null);
    setSelectionEnd(null);
  };

  // --- Copy / paste ---------------------------------------------------------
  const copySelectedPanels = () => {
    const keys = getSelectedIds(selectedCells, selectedId);
    if (!keys.size) return;
    const cells = grid.filter((c) => keys.has(c.id) && isActiveCell(c));
    if (!cells.length) return;
    const minX = Math.min(...cells.map((c) => c.x));
    const minY = Math.min(...cells.map((c) => c.y));
    const maxX = Math.max(...cells.map((c) => cellRect(c).x + cellRect(c).w));
    const maxY = Math.max(...cells.map((c) => cellRect(c).y + cellRect(c).h));
    setClipboard({
      panels: cells.map((c) => ({
        dx: c.x - minX,
        dy: c.y - minY,
        panelType: c.panelType,
        panelVariant: c.panelVariant,
        rotation: c.rotation ?? 0,
      })),
      w: maxX - minX,
      h: maxY - minY,
    });
  };

  const cancelPaste = () => {
    setIsPasting(false);
    setPasteAnchor(null);
  };

  const startPaste = () => {
    if (!clipboard) return;
    setIsPasting(true);
    // Default preview position (before the mouse moves over the workspace):
    // centred on the current wall bounds.
    setPasteAnchor({ x: wallBBox.x + wallBBox.w / 2 - clipboard.w / 2, y: wallBBox.y + wallBBox.h / 2 - clipboard.h / 2 });
  };

  // Builds real Cell records for the clipboard at a given true-space anchor
  // (top-left of the copied selection's own bounding box).
  const buildPastedCells = (anchorX: number, anchorY: number, subScreenId: string | null): Cell[] =>
    (clipboard?.panels ?? []).map((p) => ({
      id: newCellId(),
      x: anchorX + p.dx,
      y: anchorY + p.dy,
      assignedPort: null,
      sequence: null,
      assignedPowerPort: null,
      powerSequence: null,
      powerManual: false,
      isRemoved: false,
      panelVariant: p.panelVariant,
      rotation: p.rotation,
      panelType: p.panelType,
      subScreenId,
    }));

  // Snaps the paste group's anchor the same way a live panel move snaps -
  // against the canvas grid pitch and nearby existing panels' edges.
  const resolvePasteSnapAnchor = (anchorX: number, anchorY: number) => {
    if (!clipboard) return { x: anchorX, y: anchorY };
    const shadow = buildPastedCells(anchorX, anchorY, null);
    const movingGeoms = shadow.map(cellGeom);
    const firstRect = shadow[0] ? cellRect(shadow[0]) : null;
    const otherGeoms = grid.filter(isActiveCell).map(cellGeom);
    const snap = computeAnchorSnapDelta(movingGeoms, otherGeoms, snapEnabled, firstRect);
    return { x: anchorX + snap.dx, y: anchorY + snap.dy };
  };

  const updatePastePreviewFromEvent = (event: React.MouseEvent) => {
    if (!clipboard) return;
    const mm = eventToDisplayMm(event);
    if (!mm) return;
    const trueX = isFlippedView ? wallBBox.x * 2 + wallBBox.w - mm.x : mm.x;
    const trueY = mm.y;
    setPasteAnchor(resolvePasteSnapAnchor(trueX - clipboard.w / 2, trueY - clipboard.h / 2));
  };

  const commitPaste = () => {
    if (!clipboard || !pasteAnchor) return;
    const newCells = buildPastedCells(pasteAnchor.x, pasteAnchor.y, resolvedActiveSubScreenId);
    const nextPanels = [...grid, ...newCells];
    const overlaps = findOverlaps(nextPanels, cellRect);
    if (overlaps.length && !allowOverlaps) {
      setOverlapNotice(
        `Paste cancelled: it would overlap ${overlaps.length} panel pair${overlaps.length === 1 ? "" : "s"}. Enable "Allow overlaps" to override.`,
      );
      return;
    }
    setOverlapNotice(
      overlaps.length ? `${overlaps.length} overlapping panel pair${overlaps.length === 1 ? "" : "s"} kept by override.` : null,
    );
    commitGridUpdate(() => nextPanels);
    setSelectedCells(new Set(newCells.map((c) => c.id)));
    setSelectedId(null);
    setIsPasting(false);
    setPasteAnchor(null);
  };

  const assignSignalPanel = (target: Cell) => {
    if (activePort < 1) return;
    if (dragVisited.has(target.id)) return;

    commitGridUpdate((prev) => {
      const current = findCellById(prev, target.id);
      if (!current || !isActiveCell(current)) return prev;

      const currentCount = getPortPanelCount(prev, "assignedPort", activePort);
      const isAlreadySamePort = current.assignedPort === activePort;
      if (!isAlreadySamePort && currentCount >= safePanelsPerSignalPort) return prev;

      const next = cloneGrid(prev);
      const cell = findCellById(next, target.id)!;
      cell.assignedPort = activePort;
      if (!isAlreadySamePort) {
        cell.sequence = getNextSequence(next, "assignedPort", "sequence", activePort);
      }
      return next;
    });

    setDragVisited((prev) => new Set(prev).add(target.id));
  };

  const assignPowerPanel = (target: Cell) => {
    if (activePowerPort < 1) return;
    if (dragVisited.has(target.id)) return;

    commitGridUpdate((prev) => {
      const current = findCellById(prev, target.id);
      if (!current || !isActiveCell(current)) return prev;

      const currentPanels = getPortPanelCount(prev, "assignedPowerPort", activePowerPort);
      const isAlreadySamePort = current.assignedPowerPort === activePowerPort;
      if (!isAlreadySamePort && currentPanels >= safePanelsPerPowerOutlet) return prev;

      const cellWatts = PANEL_TYPES[cellPanelType(current)].power.maxW;
      const currentPortLoad = getPowerPortLoadWatts(prev, activePowerPort, 0, current.id);
      if (!isAlreadySamePort && currentPortLoad + cellWatts > MAX_OUTLET_AMPS * VOLTAGE) return prev;

      const next = cloneGrid(prev);
      const cell = findCellById(next, target.id)!;
      cell.assignedPowerPort = activePowerPort;
      cell.powerManual = true;
      if (!isAlreadySamePort) {
        cell.powerSequence = getNextSequence(next, "assignedPowerPort", "powerSequence", activePowerPort);
      }
      return next;
    });

    setDragVisited((prev) => new Set(prev).add(target.id));
  };

  // --- Workspace pointer interactions -------------------------------------
  // Patch mode: press/drag over panels assigns the active port.
  // Select mode: click selects, drag draws a marquee (workspace mm space).
  // Move mode: drag repositions the pressed panel, the multi-selection it
  // belongs to, or its joined group; snap + overlap checks run on release.

  const onPanelMouseDown = (cell: Cell, event: React.MouseEvent) => {
    if (isPanelDimmed(cell)) return;
    if (editMode === "move") {
      if (!isActiveCell(cell)) return;
      event.preventDefault();
      const mm = eventToDisplayMm(event);
      if (!mm) return;
      let ids: string[];
      if (activeSelectedKeys.has(cell.id) && selectedCount > 1) {
        ids = [...activeSelectedKeys].filter((id) => isActiveCell(findCellById(grid, id)));
      } else if (moveJoinedGroup) {
        ids = [...joinedGroupIdsByGeom(grid, cellGeom, new Set([cell.id]))];
      } else {
        ids = [cell.id];
      }
      if (!activeSelectedKeys.has(cell.id)) {
        setSelectedId(cell.id);
        setSelectedCells(new Set([cell.id]));
      }
      setOverlapNotice(null);
      setMoveDrag({ ids, startX: mm.x, startY: mm.y, dx: 0, dy: 0 });
      return;
    }
    if (editMode === "select") {
      const mm = eventToDisplayMm(event);
      setSelectionStart(mm);
      setSelectionEnd(mm);
      setIsSelectingPanels(true);
      if (event.shiftKey) {
        setSelectedCells((prev) => {
          const next = new Set(prev);
          if (next.has(cell.id)) next.delete(cell.id);
          else next.add(cell.id);
          return next;
        });
        setSelectedId(cell.id);
      } else {
        setSelectedId(cell.id);
        setSelectedCells(new Set([cell.id]));
      }
      return;
    }
    if (!isActiveCell(cell)) return;
    setDragVisited(new Set());
    setIsDragging(true);
    if (patchMode === "signal") assignSignalPanel(cell);
    else assignPowerPanel(cell);
  };

  const onPanelMouseEnter = (cell: Cell) => {
    if (isPanelDimmed(cell)) return;
    if (editMode !== "patch" || !isDragging) return;
    if (!isActiveCell(cell)) return;
    if (patchMode === "signal") assignSignalPanel(cell);
    else assignPowerPanel(cell);
  };

  const applyManualSignalPatch = (value: string) => {
    if (!selectedId) return;
    const nextPort = value === "" ? null : Number.parseInt(value, 10);
    if (nextPort !== null && (!Number.isFinite(nextPort) || nextPort < 1 || nextPort > primarySignalPortCount)) return;

    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      const target = findCellById(next, selectedId);
      if (!target) return prev;
      if (!isActiveCell(target)) return prev;

      if (nextPort === null) {
        target.assignedPort = null;
        target.sequence = null;
        return next;
      }

      const currentCount = getPortPanelCount(prev, "assignedPort", nextPort);
      const isAlreadySamePort = target.assignedPort === nextPort;
      if (!isAlreadySamePort && currentCount >= safePanelsPerSignalPort) return prev;

      target.assignedPort = nextPort;
      if (!isAlreadySamePort) {
        target.sequence = getNextSequence(next, "assignedPort", "sequence", nextPort);
      }
      return next;
    });
  };

  const applyManualPowerPatch = (value: string) => {
    if (!selectedId) return;
    const nextPort = value === "" ? null : Number.parseInt(value, 10);
    if (nextPort !== null && (!Number.isFinite(nextPort) || nextPort < 1 || nextPort > powerPorts.length)) return;

    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      const target = findCellById(next, selectedId);
      if (!target) return prev;
      if (!isActiveCell(target)) return prev;

      if (nextPort === null) {
        target.assignedPowerPort = null;
        target.powerSequence = null;
        target.powerManual = false;
        return next;
      }

      const currentPanels = getPortPanelCount(prev, "assignedPowerPort", nextPort);
      const isAlreadySamePort = target.assignedPowerPort === nextPort;
      if (!isAlreadySamePort && currentPanels >= safePanelsPerPowerOutlet) return prev;

      const currentPortLoad = getPowerPortLoadWatts(prev, nextPort, powerSpec.maxW, selectedId);
      if (!isAlreadySamePort && currentPortLoad + powerSpec.maxW > MAX_OUTLET_AMPS * VOLTAGE) return prev;

      target.assignedPowerPort = nextPort;
      target.powerManual = true;
      if (!isAlreadySamePort) {
        target.powerSequence = getNextSequence(next, "assignedPowerPort", "powerSequence", nextPort);
      }
      return next;
    });
  };

  const snakePatch = () => {
    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      // Scope: when a sub-screen is being edited, auto-patch must only
      // touch/reorder that sub-screen's panels - other sub-screens' existing
      // assignments are left completely alone, but still count against the
      // shared ports' capacity (ports are physically shared hardware).
      const scopeIds = currentScopeIds(); // null = whole grid, today's behaviour when unscoped.

      // Reading order over the free layout. LETTERS = one segment per connected
      // letter (bottom-up, branch-aware); LOOP_TOGETHER = one segment per loop;
      // otherwise a single reading-order segment over row/column bands.
      const letterMode = snakeDirection === "LETTERS";
      const scopedForOrdering = scopeIds ? next.filter((c) => scopeIds.has(c.id)) : next;
      const segments = letterMode
        ? orderPanelsForLetters(scopedForOrdering)
        : orderPanelsForSnake(scopedForOrdering, snakeDirection, snakeAlternates);

      if (patchMode === "signal") {
        for (const cell of next) {
          if (scopeIds && !scopeIds.has(cell.id)) continue;
          cell.assignedPort = null;
          cell.sequence = null;
        }

        // Live capacity-aware walk (mirrors the power loop below): unlike the
        // old blind port/seq counters, this queries actual port occupancy so
        // it correctly skips ports another sub-screen has already filled,
        // instead of assuming every port starts empty. Bounded by
        // primarySignalPortCount (not the full port list) so auto-patching
        // never assigns into the range reserved for the backup signal loop.
        let portIndex = 0;
        const advanceToPortWithCapacity = () => {
          while (portIndex < primarySignalPortCount && getPortPanelCount(next, "assignedPort", portIndex + 1) >= safePanelsPerSignalPort) {
            portIndex += 1;
          }
        };
        const assignToSignalPort = (cell: Cell) => {
          advanceToPortWithCapacity();
          if (portIndex >= primarySignalPortCount) return;
          const port = portIndex + 1;
          cell.assignedPort = port;
          cell.sequence = getNextSequence(next, "assignedPort", "sequence", port);
        };

        segments.forEach((segment) => {
          // Letter mode: don't split a letter across ports - advance to a fresh
          // port first if the whole letter won't fit in the current port's
          // remaining capacity (unless the letter is larger than a full port,
          // in which case it must split).
          if (letterMode) {
            advanceToPortWithCapacity();
            if (portIndex < primarySignalPortCount) {
              const currentCount = getPortPanelCount(next, "assignedPort", portIndex + 1);
              const fitsFullPort = segment.length <= safePanelsPerSignalPort;
              const fitsRemaining = currentCount + segment.length <= safePanelsPerSignalPort;
              if (currentCount > 0 && fitsFullPort && !fitsRemaining) portIndex += 1;
            }
          }
          segment.forEach(assignToSignalPort);
          // Each loop-together segment starts on a fresh port.
          if (!letterMode && segments.length > 1) {
            advanceToPortWithCapacity();
            if (portIndex < primarySignalPortCount && getPortPanelCount(next, "assignedPort", portIndex + 1) > 0) portIndex += 1;
          }
        });
      }

      if (patchMode === "power") {
        for (const cell of next) {
          if (scopeIds && !scopeIds.has(cell.id)) continue;
          cell.assignedPowerPort = null;
          cell.powerSequence = null;
          cell.powerManual = false;
        }

        let portIndex = 0;
        const assignToPlug = (cell: Cell) => {
          const cellWatts = PANEL_TYPES[cellPanelType(cell)].power.maxW;
          while (portIndex < powerPorts.length) {
            const port = powerPorts[portIndex];
            const currentLoad = getPowerPortLoadWatts(next, port.id, 0);
            const currentPanels = getPortPanelCount(next, "assignedPowerPort", port.id);
            if (currentPanels >= safePanelsPerPowerOutlet) {
              portIndex += 1;
              continue;
            }
            if (currentLoad + cellWatts <= MAX_OUTLET_AMPS * VOLTAGE) {
              cell.assignedPowerPort = port.id;
              cell.powerSequence = getNextSequence(next, "assignedPowerPort", "powerSequence", port.id);
              cell.powerManual = false;
              return;
            }
            portIndex += 1;
          }
        };

        segments.forEach((segment) => {
          // Letter mode: keep a letter on one plug where it fits (advance first
          // if the whole letter won't fit the current plug's count/amp headroom).
          if (letterMode && portIndex < powerPorts.length) {
            const plug = powerPorts[portIndex];
            const load = getPowerPortLoadWatts(next, plug.id, 0);
            const count = getPortPanelCount(next, "assignedPowerPort", plug.id);
            const letterWatts = segment.reduce((s, c) => s + PANEL_TYPES[cellPanelType(c)].power.maxW, 0);
            const fitsFullPlug = segment.length <= safePanelsPerPowerOutlet && letterWatts <= MAX_OUTLET_AMPS * VOLTAGE;
            const fitsRemaining = count + segment.length <= safePanelsPerPowerOutlet && load + letterWatts <= MAX_OUTLET_AMPS * VOLTAGE;
            if (count > 0 && fitsFullPlug && !fitsRemaining) portIndex += 1;
          }
          segment.forEach(assignToPlug);
        });
      }

      return next;
    });

    setSelectedId(null);
    setSelectedCells(new Set());
    setDragVisited(new Set());
    setIsDragging(false);
  };

  // Remove every panel, leaving an empty workspace (undoable).
  const clearAllPanels = () => {
    const scopeIds = currentScopeIds();
    const scopedLabel = scopeIds ? "this sub-screen" : "the layout";
    const affectedCount = scopeIds ? grid.filter((c) => scopeIds.has(c.id)).length : grid.length;
    if (affectedCount && !window.confirm(`Remove all panels from ${scopedLabel}? This can be undone.`)) return;
    commitGridUpdate((prev) => (scopeIds ? prev.filter((c) => !scopeIds.has(c.id)) : []));
    setSelectedId(null);
    setSelectedCells(new Set());
    setDragVisited(new Set());
    setIsDragging(false);
    setOverlapNotice(null);
  };

  // Patch power to follow the existing signal patch: walk panels in signal order
  // (signal port, then sequence) and fill power plugs, starting a fresh plug for
  // each signal port so power plugs line up with the signal ports. Respects the
  // power panel-count and amp limits, and stops when the plugs run out.
  const matchPowerToSignal = () => {
    // Scope: only check/follow the active sub-screen's own signal patching -
    // otherwise "patch signal first" could fire (or not) based on unrelated
    // sub-screens, which would be confusing.
    const scopeIdsForCheck = currentScopeIds();
    const hasSignal = grid.some(
      (cell) => isActiveCell(cell) && cell.assignedPort && (!scopeIdsForCheck || scopeIdsForCheck.has(cell.id)),
    );
    if (!hasSignal) {
      alert("Patch the signal ports first - power will follow the same pattern.");
      return;
    }

    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      const scopeIds = scopeIdsForCheck;

      for (const cell of next) {
        if (scopeIds && !scopeIds.has(cell.id)) continue;
        cell.assignedPowerPort = null;
        cell.powerSequence = null;
        cell.powerManual = false;
      }

      const byPort = new Map<number, Cell[]>();
      next.forEach((cell) => {
        if (!isActiveCell(cell) || !cell.assignedPort) return;
        if (scopeIds && !scopeIds.has(cell.id)) return;
        const list = byPort.get(cell.assignedPort) ?? [];
        list.push(cell);
        byPort.set(cell.assignedPort, list);
      });
      const orderedSignalPorts = [...byPort.keys()].sort((a, b) => a - b);

      let plugIndex = 0;
      const plugLeft = () => plugIndex < powerPorts.length;

      for (const sigPort of orderedSignalPorts) {
        if (!plugLeft()) break;
        // Align power plugs to signal ports: each new signal port starts on a fresh plug.
        if (getPortPanelCount(next, "assignedPowerPort", powerPorts[plugIndex].id) > 0) {
          plugIndex += 1;
        }

        const cells = byPort.get(sigPort)!.sort((a, b) => (a.sequence ?? 0) - (b.sequence ?? 0));
        for (const cell of cells) {
          const cellWatts = PANEL_TYPES[cellPanelType(cell)].power.maxW;
          let placed = false;
          while (plugLeft()) {
            const plug = powerPorts[plugIndex];
            const currentPanels = getPortPanelCount(next, "assignedPowerPort", plug.id);
            const currentLoad = getPowerPortLoadWatts(next, plug.id, 0);
            if (currentPanels >= safePanelsPerPowerOutlet) {
              plugIndex += 1;
              continue;
            }
            if (currentLoad + cellWatts > MAX_OUTLET_AMPS * VOLTAGE) {
              plugIndex += 1;
              continue;
            }
            cell.assignedPowerPort = plug.id;
            cell.powerSequence = getNextSequence(next, "assignedPowerPort", "powerSequence", plug.id);
            cell.powerManual = false;
            placed = true;
            break;
          }
          if (!placed) break;
        }
        if (!plugLeft()) break;
      }

      return next;
    });

    setPatchMode("power");
    setSelectedId(null);
    setSelectedCells(new Set());
    setDragVisited(new Set());
    setIsDragging(false);
  };

  const clearSignalCabling = () => {
    commitGridUpdate((prev) => clearSignalOnGrid(prev, currentScopeIds()));
    setSelectedId(null);
    setSelectedCells(new Set());
    setDragVisited(new Set());
    setIsDragging(false);
  };

  const clearPowerAssignments = () => {
    commitGridUpdate((prev) => clearPowerOnGrid(prev, currentScopeIds()));
    setSelectedId(null);
    setSelectedCells(new Set());
  };

  const clearSelectedPanelPatching = () => {
    const keys = getSelectedIds(selectedCells, selectedId);
    if (!keys.size) return;
    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      keys.forEach((key) => {
        const target = findCellById(next, key);
        if (!target || !isActiveCell(target)) return;
        target.assignedPort = null;
        target.sequence = null;
        target.assignedPowerPort = null;
        target.powerSequence = null;
        target.powerManual = false;
      });
      return next;
    });
  };

  // --- Sub-screen CRUD / assignment ----------------------------------------
  const selectSubScreen = (id: string | null) => {
    setActiveSubScreenId(id);
    // The old selection almost certainly doesn't belong to the new scope;
    // clearing avoids leaving a dimmed/invisible panel "selected".
    setSelectedId(null);
    setSelectedCells(new Set());
  };

  const createSubScreen = (name: string) => {
    const snapshot = captureLayout();
    const screen = makeSubScreen(name, Date.now() + subScreens.length);
    setSubScreens((prev) => [...prev, screen]);
    // Deliberately stay on whatever view the user was already on (usually
    // Canvas View) instead of jumping into the brand-new, empty sub-screen -
    // switching there immediately would scope the workspace down to zero
    // panels and dim/lock everything else, which looks like the whole
    // layout vanished. The user assigns panels to it first, then switches in.
    pushUndoSnapshot(snapshot);
  };

  const renameSubScreen = (id: string, name: string) => {
    commitSubScreensUpdate((prev) => prev.map((s) => (s.id === id ? { ...s, name } : s)));
  };

  // Deleting a sub-screen unassigns its panels (they become "unassigned",
  // not deleted) and falls back to Canvas View if it was the active one.
  const deleteSubScreen = (id: string) => {
    const snapshot = captureLayout();
    setGrid((prev) => prev.map((cell) => (cell.subScreenId === id ? { ...cell, subScreenId: null } : cell)));
    setSubScreens((prev) => prev.filter((s) => s.id !== id));
    if (resolvedActiveSubScreenId === id) setActiveSubScreenId(null);
    setSelectedId(null);
    setSelectedCells(new Set());
    pushUndoSnapshot(snapshot);
  };

  const assignSelectedToSubScreen = (id: string) => {
    const keys = getSelectedIds(selectedCells, selectedId);
    if (!keys.size) return;
    commitGridUpdate((prev) => prev.map((cell) => (keys.has(cell.id) ? { ...cell, subScreenId: id } : cell)));
  };

  const removeSelectedFromSubScreen = () => {
    const keys = getSelectedIds(selectedCells, selectedId);
    if (!keys.size) return;
    commitGridUpdate((prev) => prev.map((cell) => (keys.has(cell.id) ? { ...cell, subScreenId: null } : cell)));
  };

  const selectAllInSubScreen = (id: string) => {
    const ids = grid.filter((cell) => !cell.isRemoved && cell.subScreenId === id).map((cell) => cell.id);
    setSelectedCells(new Set(ids));
    setSelectedId(null);
  };

  // --- Output canvas positioning --------------------------------------------
  // id === null updates the whole-layout position (used only when no
  // sub-screens exist); otherwise updates that sub-screen's canvas position.
  // Kept as a plain project-data mutation through commitCanvasUpdate, so
  // repositioning participates in the same single undo stack as everything
  // else - dragging a sub-screen never touches any panel's mm x/y.
  const updateCanvasPosition = (id: string | null, x: number, y: number) => {
    commitCanvasUpdate(() => {
      if (id === null) {
        setWholeLayoutCanvasX(x);
        setWholeLayoutCanvasY(y);
      } else {
        setSubScreens((prev) => prev.map((s) => (s.id === id ? { ...s, canvasX: x, canvasY: y } : s)));
      }
    });
  };

  const updateOutputCanvasResolution = (w: number, h: number) => {
    commitCanvasUpdate(() => {
      setOutputCanvasW(w);
      setOutputCanvasH(h);
    });
  };

  /** key is a sub-screen id, or WHOLE_LAYOUT_KEY when no sub-screens exist. */
  const updateCanvasInput = (key: string, interfacePk: number | null) => {
    setCanvasInputs((prev) => ({ ...prev, [key]: interfacePk }));
  };

  // Delete now prompts (Remove / Mark Inactive / Cancel); the button and the
  // Delete key just open the confirmation.
  const deleteSelectedPanel = () => {
    const keys = getSelectedIds(selectedCells, selectedId);
    if (!keys.size) return;
    setShowDeleteConfirm(true);
  };

  // Permanently remove the selected panels from the layout.
  const removeSelectedPanels = () => {
    const keys = getSelectedIds(selectedCells, selectedId);
    if (!keys.size) return;
    commitGridUpdate((prev) => cloneGrid(prev).filter((cell) => !keys.has(cell.id)));
    setSelectedId(null);
    setSelectedCells(new Set());
    setShowDeleteConfirm(false);
  };

  // Keep the selected panels in place but mark them inactive (excluded from
  // totals, patching and outputs).
  const markSelectedInactive = () => {
    const keys = getSelectedIds(selectedCells, selectedId);
    if (!keys.size) return;
    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      keys.forEach((key) => {
        const target = findCellById(next, key);
        if (!target || target.isRemoved) return;
        target.assignedPort = null;
        target.sequence = null;
        target.assignedPowerPort = null;
        target.powerSequence = null;
        target.powerManual = false;
        target.isRemoved = true;
      });
      return next;
    });
    setShowDeleteConfirm(false);
  };

  const restoreSelectedPanel = () => {
    const keys = getSelectedIds(selectedCells, selectedId);
    if (!keys.size) return;
    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      keys.forEach((key) => {
        const target = findCellById(next, key);
        if (!target || !target.isRemoved) return;
        target.isRemoved = false;
        target.assignedPort = null;
        target.sequence = null;
        target.assignedPowerPort = null;
        target.powerSequence = null;
        target.powerManual = false;
        target.panelVariant = "STANDARD";
        target.rotation = 0;
      });
      return next;
    });
  };

  const applySelectedPanelVariant = (variant: PanelVariantKey) => {
    const keys = getSelectedIds(selectedCells, selectedId);
    if (!keys.size) return;
    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      keys.forEach((key) => {
        const target = findCellById(next, key);
        if (!target || !isActiveCell(target)) return;
        // Variants (triangle/curve/corner) are MG9-only.
        if (cellPanelType(target) !== "MG9") return;
        target.panelVariant = variant;
      });
      return next;
    });
  };

  const applySelectedPanelType = (type: PanelTypeKey) => {
    const keys = getSelectedIds(selectedCells, selectedId);
    if (!keys.size) return;
    commitGridUpdate((prev) => {
      let next = cloneGrid(prev);
      keys.forEach((key) => {
        next = convertPanelTypeInList(next, key, type);
      });
      return next;
    });
  };

  // Rotates every selected panel in place by deltaDeg (any angle, not just a
  // multiple of 90). Each panel spins around its own centre - positions never
  // move - so a multi-selected group's arrangement and spacing relative to
  // each other is preserved automatically.
  const rotateSelectedPanels = (deltaDeg: number = 90) => {
    const keys = getSelectedIds(selectedCells, selectedId);
    if (!keys.size) return;
    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      keys.forEach((key) => {
        const target = findCellById(next, key);
        if (!target || !isActiveCell(target)) return;
        target.rotation = (((target.rotation ?? 0) + deltaDeg) % 360 + 360) % 360;
      });
      return next;
    });
  };

  const clearSelectedPortPatching = () => {
    if ((patchMode === "signal" && activePort < 1) || (patchMode === "power" && activePowerPort < 1)) return;
    const scopeIds = currentScopeIds();
    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      next.forEach((cell) => {
        if (!isActiveCell(cell)) return;
        if (scopeIds && !scopeIds.has(cell.id)) return;
        if (patchMode === "signal" && cell.assignedPort === activePort) {
          cell.assignedPort = null;
          cell.sequence = null;
        }
        if (patchMode === "power" && cell.assignedPowerPort === activePowerPort) {
          cell.assignedPowerPort = null;
          cell.powerSequence = null;
          cell.powerManual = false;
        }
      });
      return next;
    });
  };

  // Workspace pixel size (bbox + padding at the current zoom).
  const svgW = Math.max(1, Math.round(mmToPx(workspaceSizeMm.w)));
  const svgH = Math.max(1, Math.round(mmToPx(workspaceSizeMm.h)));
  // Human-friendly row/column labels for a panel (grid-ish when aligned).
  const panelRowLabel = (cell: Cell) => (bandIndexById.get(cell.id) ?? 0) + 1;
  const panelColLabel = (cell: Cell) => {
    const col = (cellRect(cell).x - wallBBox.x) / MODULE_MM + 1;
    return Number.isInteger(col) ? String(col) : col.toFixed(1);
  };

  // Cable hops (one per adjacent panel pair) in display pixels, shared by the
  // behind-panels line layer and the in-front arrowhead layer so both stay in
  // sync. Signal hops are offset -4px, power hops +4px so the two runs separate.
  type CableHop = { key: string; pts: Array<{ x: number; y: number }>; color: string; dir: { x: number; y: number } };
  const cableHops: CableHop[] = [];
  const centerOf = (r: RectMm) => ({ x: r.x + r.w / 2, y: r.y + r.h / 2 });
  Object.entries(signalPortStats).forEach(([portId, stat]) => {
    if (!stat.path || stat.path.length < 2) return;
    const color = PORT_COLORS[(Number(portId) - 1) % PORT_COLORS.length];
    stat.path.forEach((cell, idx) => {
      if (idx === 0) return;
      const a = rectToPx(displayRectOf(stat.path[idx - 1]));
      const b = rectToPx(displayRectOf(cell));
      const ca = centerOf(a);
      const cb = centerOf(b);
      cableHops.push({ key: `sig-${portId}-${idx}`, pts: routeCablePx(a, b, -4), color, dir: { x: cb.x - ca.x, y: cb.y - ca.y } });
    });
  });
  powerPorts.forEach((port) => {
    const stat = powerPortStats[port.id];
    const path = stat?.path ?? [];
    if (path.length < 2) return;
    path.forEach((cell, idx) => {
      if (idx === 0) return;
      const a = rectToPx(displayRectOf(path[idx - 1]));
      const b = rectToPx(displayRectOf(cell));
      const ca = centerOf(a);
      const cb = centerOf(b);
      cableHops.push({ key: `pow-${port.id}-${idx}`, pts: routeCablePx(a, b, 4), color: POWER_COLOR, dir: { x: cb.x - ca.x, y: cb.y - ca.y } });
    });
  });
  // Arrowhead polygon at the destination end of a hop, pointing along the last
  // route segment. When adjacent panels touch, that segment collapses to a point,
  // so fall back to the source->destination centre direction (e.g. a panel wired
  // to the one above it points the arrow upward).
  const cableArrowHead = (hop: CableHop, size = 9) => {
    const pts = hop.pts;
    const p2 = pts[pts.length - 1];
    const p1 = pts[pts.length - 2];
    const segDx = p2.x - p1.x;
    const segDy = p2.y - p1.y;
    const ang = Math.hypot(segDx, segDy) < 1 ? Math.atan2(hop.dir.y, hop.dir.x) : Math.atan2(segDy, segDx);
    const x1 = p2.x - size * Math.cos(ang - Math.PI / 6);
    const y1 = p2.y - size * Math.sin(ang - Math.PI / 6);
    const x2 = p2.x - size * Math.cos(ang + Math.PI / 6);
    const y2 = p2.y - size * Math.sin(ang + Math.PI / 6);
    return (
      <polygon
        key={`ah-${hop.key}`}
        points={`${p2.x},${p2.y} ${x1},${y1} ${x2},${y2}`}
        fill={hop.color}
        stroke="black"
        strokeWidth={1}
        strokeLinejoin="round"
      />
    );
  };

  return (
    <div className="min-h-screen bg-[#0f172a] p-6 text-white print-container">
      {showHelp ? <HelpModal onClose={() => setShowHelp(false)} /> : null}
      {showDeleteConfirm ? (
        <DeleteConfirmModal
          count={selectedCount}
          onRemove={removeSelectedPanels}
          onMarkInactive={markSelectedInactive}
          onCancel={() => setShowDeleteConfirm(false)}
        />
      ) : null}
      {showGridSizeConfirm ? (
        <GridSizeConfirmModal
          onConfirm={() => {
            setShowGridSizeConfirm(false);
            performApplyGridSize();
          }}
          onCancel={() => setShowGridSizeConfirm(false)}
        />
      ) : null}
      {showDownloadFormatModal ? (
        <DownloadFormatModal
          format={downloadFormat}
          onFormatChange={setDownloadFormat}
          onCancel={() => setShowDownloadFormatModal(false)}
          onDownload={() => {
            setShowDownloadFormatModal(false);
            if (downloadFormat === "mp4") downloadMovingTestPatternMp4();
            else downloadMovingTestPatternVideo();
          }}
        />
      ) : null}
      {importPreview ? (
        <ImportPreviewModal
          result={importPreview}
          hasUnsavedWork={hasUnsavedWork}
          onCancel={() => setImportPreview(null)}
          onApply={applyImport}
        />
      ) : null}
      {pendingQuickLayoutTransfer ? (
        <QuickLayoutTransferModal
          payload={pendingQuickLayoutTransfer}
          onCancel={() => setPendingQuickLayoutTransfer(null)}
          onReplace={() => applyQuickLayoutTransfer("replace")}
          onAdd={() => applyQuickLayoutTransfer("add")}
        />
      ) : null}
      <style>{`
        @media print {
          @page { size: landscape; margin: 12mm; }
          body { background: white !important; color: black !important; }
          .no-print { display: none !important; }
          .print-container { padding: 0 !important; background: white !important; }
          .print-card { background: white !important; color: black !important; border-color: #d1d5db !important; box-shadow: none !important; }
          .print-card * { color: black !important; text-shadow: none !important; }
        }
      `}</style>

      <div className="mx-auto max-w-[1900px] space-y-6">
        <div className="flex flex-wrap items-center justify-between gap-3 no-print">
          <div>
            <div className="text-sm uppercase tracking-[0.2em] text-sky-300">LED cabling planner</div>
            <div className="flex items-center gap-3">
              <h1 className="text-3xl font-semibold text-white [text-shadow:0_0_2px_black]">LED Port Mapper</h1>
              <a
                className="rounded-full border border-slate-500 bg-slate-800 px-3 py-1 text-xs font-semibold text-slate-200 hover:bg-slate-700"
                href="https://github.com/underdog1234/LED-Cabling-Web-App#recent-changes-in-v0202"
                target="_blank"
                rel="noreferrer"
              >
                v{APP_VERSION}
              </a>
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-3">
            <div className="flex items-center gap-2 rounded-lg border border-slate-700/70 bg-slate-900/40 p-1.5">
              <Button intent="primary" onClick={generatePdf}>
                <FileText className="h-4 w-4" />Generate PDF
              </Button>
              <Button intent="primary" onClick={exportTestPatternPng}>
                <ImageDown className="h-4 w-4" />Test Pattern
              </Button>
              <Button intent="primary" onClick={openMovingTestPatternTab}>
                <Video className="h-4 w-4" />Moving Test Pattern
              </Button>
              <Button
                intent="primary"
                onClick={() => setShowDownloadFormatModal(true)}
                disabled={isRecordingVideo || isEncodingMp4}
              >
                <Download className="h-4 w-4" />
                {isEncodingMp4
                  ? `Encoding MP4… ${Math.round(mp4EncodeProgress * 100)}%`
                  : isRecordingVideo
                    ? `Recording… ${videoRecordSeconds.toFixed(0)}/${LOOP_SECONDS}s`
                    : "Download Moving Test Pattern"}
              </Button>
              <label className="flex items-center gap-1 px-1 text-xs text-slate-200" title="Include the vertical centre indicator in the PDF's Panel Layout pages">
                <input type="checkbox" checked={includeCentreLineInExport} onChange={() => setIncludeCentreLineInExport((prev) => !prev)} />
                Include Centre Line
              </label>
            </div>
            <div className="flex items-center gap-2 rounded-lg border border-slate-700/70 bg-slate-900/40 p-1.5">
              <Button intent="secondary" onClick={exportJson}>
                <Download className="h-4 w-4" />Save
              </Button>
              <Button intent="secondary" onClick={() => fileInputRef.current?.click()}>
                <Upload className="h-4 w-4" />Open
              </Button>
              <Button intent="secondary" onClick={() => importInputRef.current?.click()} title="Import a project from the Creative Layout Tool">
                <Upload className="h-4 w-4" />Import Project from Creative Layout Tool
              </Button>
              <Button intent="secondary" onClick={openQuickPanelLayoutTab} title="Open a standalone panel-count calculator in a new tab">
                <LayoutGrid className="h-4 w-4" />Quick Panel Layout
              </Button>
              <Button intent="ghost" onClick={() => setShowHelp(true)}>
                <HelpCircle className="h-4 w-4" />Help
              </Button>
            </div>
            <input ref={fileInputRef} type="file" accept="application/json" className="hidden" onChange={openJson} />
            <input ref={importInputRef} type="file" accept="application/json,.json" className="hidden" onChange={handleImportFile} />
          </div>
        </div>

        <div className="grid gap-4 xl:grid-cols-[1.1fr_1.2fr]">
          <Card className="border-slate-700 bg-slate-800 print-card" collapsible>
            <CardHeader>
            <CardTitle className="text-white [text-shadow:0_0_2px_black]">LED Wall Setup</CardTitle>
          </CardHeader>
          <CardContent className="space-y-4 text-white [text-shadow:0_0_2px_black]">
              <div className="grid gap-4 md:grid-cols-2">
                <div className="space-y-1 md:col-span-2">
                  <label className="text-xs text-slate-300">Project Name</label>
                  <Input className="bg-white text-black" type="text" value={projectName} onChange={(e) => setProjectName(e.target.value)} placeholder="Enter project name" />
                </div>
                <div className="space-y-1">
                  <label className="text-xs text-slate-300">Columns →</label>
                  <Input className="bg-white text-black" type="number" min="1" step="1" value={draftCols} onChange={(e) => setDraftCols(e.target.value)} />
                </div>
                <div className="space-y-1">
                  <label className="text-xs text-slate-300">Rows ↓</label>
                  <Input className="bg-white text-black" type="number" min="1" step="1" value={draftRows} onChange={(e) => setDraftRows(e.target.value)} />
                </div>
              </div>

              <div className="grid gap-4 md:grid-cols-2">
                <div className="space-y-1">
                  <label className="text-xs text-slate-300">Panel Type</label>
                  <select className="w-full rounded bg-white p-2 text-black" value={panelType} onChange={(e) => setPanelType(e.target.value as PanelTypeKey)}>
                    {Object.entries(PANEL_TYPES).map(([key, value]) => (
                      <option key={key} value={key}>{value.name}</option>
                    ))}
                  </select>
                </div>
                <div className="space-y-1">
                  <label className="text-xs text-slate-300">Power Distro</label>
                  <select className="w-full rounded bg-white p-2 text-black" value={powerDistro} onChange={(e) => setPowerDistro(e.target.value as PowerDistroKey)}>
                    {Object.values(POWER_DISTROS).map((option) => (
                      <option key={option.id} value={option.id}>{option.label}</option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="grid gap-4 md:grid-cols-2">
                <div className="space-y-1">
                  <label className="text-xs text-slate-300">Processor Model</label>
                  <select
                    className="w-full rounded bg-white p-2 text-black"
                    value={processorModel}
                    onChange={(e) => setProcessorModel(e.target.value as ProcessorModelId | "")}
                  >
                    <option value="">None selected</option>
                    {PROCESSOR_MODEL_IDS.map((id) => (
                      <option key={id} value={id}>{PROCESSOR_SPECS[id].label}</option>
                    ))}
                  </select>
                </div>
                <div className="space-y-1">
                  <label className="text-xs text-slate-300">Processor Capacity</label>
                  <div className="rounded border border-slate-700 bg-slate-900 p-2 text-xs">
                    {processorModel && novaStarValidation ? (
                      <>
                        <div>
                          {novaStarValidation.summary.outputPixelLoads.reduce((sum, o) => sum + o.pixels, 0).toLocaleString()} /{" "}
                          {PROCESSOR_SPECS[processorModel].maxTotalPixels.toLocaleString()} px total
                        </div>
                        <div>
                          {novaStarValidation.summary.ethernetOutputsUsed} / {PROCESSOR_SPECS[processorModel].ethernetOutputCount} outputs used
                        </div>
                      </>
                    ) : (
                      <span className="text-slate-400">Select a processor to see capacity usage.</span>
                    )}
                  </div>
                </div>
              </div>

              <div className="flex flex-wrap gap-2 no-print">
                <Button onClick={applyGridSize}>Apply Grid Size</Button>
                <Button intent="danger" onClick={clearAllPanels}>Clear All Panels</Button>
              </div>

              <div className="grid gap-4 md:grid-cols-2">
                <div className="space-y-2 rounded border border-slate-700 bg-slate-900 p-3">
                  <label className="text-sm font-semibold">Panels per Power Outlet</label>
                  <Input
                    className="bg-white text-black"
                    type="number"
                    min="1"
                    max="21"
                    value={safePanelsPerPowerOutlet}
                    onChange={(e) => {
                      const raw = Number.parseInt(e.target.value || "0", 10);
                      const next = Math.min(Math.max(raw || 1, 1), 21);
                      setPanelsPerPowerOutlet(next);
                    }}
                  />
                  <div className="text-xs">{formatNumber(powerOutletWatts)} W</div>
                  <div className="text-xs">{formatNumber(powerOutletAmps, 2)} A</div>
                  <UtilBar percent={powerOutletPercent} />
                  <div className="text-xs">{formatNumber(powerOutletPercent, 1)}% of 16A</div>
                </div>

                <div className="space-y-2 rounded border border-slate-700 bg-slate-900 p-3">
                  <label className="text-sm font-semibold">Panels per Signal Port</label>
                  <Input
                    className="bg-white text-black"
                    type="number"
                    min="1"
                    max={panel.defaults.signalPanelsPerPort}
                    value={safePanelsPerSignalPort}
                    onChange={(e) => {
                      const raw = Number.parseInt(e.target.value || "0", 10);
                      const next = Math.min(Math.max(raw || 1, 1), panel.defaults.signalPanelsPerPort);
                      setPanelsPerSignalPort(next);
                    }}
                  />
                  <div className="text-xs">{formatNumber(signalPortPixels)} pixels</div>
                  <UtilBar percent={signalPortPercent} />
                  <div className="text-xs">{formatNumber(signalPortPercent, 1)}% of 650,000</div>
                </div>
              </div>

              <ControlGroup label="Patch mode" className="no-print">
                <Button active={patchMode === "signal"} activeAccent="sky" intent="secondary" onClick={() => setPatchMode("signal")}>
                  <Zap className="h-4 w-4" />Signal Patch Mode
                </Button>
                <Button active={patchMode === "power"} activeAccent="amber" intent="secondary" onClick={() => setPatchMode("power")}>
                  <Zap className="h-4 w-4" />Power Patch Mode
                </Button>
                <Button
                  intent="secondary"
                  onClick={matchPowerToSignal}
                  title="Patch power plugs to follow the signal patch order, aligned to the signal ports"
                >
                  <Wand2 className="h-4 w-4" />Match Power To Signal Pattern
                </Button>
                <StatusChip tone={patchMode === "signal" ? "sky" : "amber"}>
                  {patchMode === "signal"
                    ? activePort > 0 ? `Signal patching · port ${activePort}` : "Signal patching · no port selected"
                    : activePowerPort > 0 ? `Power patching · plug ${activePowerPort}` : "Power patching · no plug selected"}
                </StatusChip>
              </ControlGroup>

              <ControlGroup label="Auto patching" className="no-print">
                <select className="rounded-lg border border-slate-500 bg-white p-2 text-sm text-black" value={snakeDirection} onChange={(e) => setSnakeDirection(e.target.value as typeof snakeDirection)}>
                  <option value="LR">Left to Right</option>
                  <option value="RL">Right to Left</option>
                  <option value="LRB">Left to Right from the Bottom</option>
                  <option value="RLB">Right to Left from the Bottom</option>
                  <option value="TB">Top to Bottom</option>
                  <option value="BT">Bottom to Top</option>
                  <option value="LOOP_TOGETHER">Loop together</option>
                  <option value="LETTERS">Letter patching (bottom-up)</option>
                </select>
                <label className="flex items-center gap-2 rounded-lg border border-slate-600 bg-slate-900 px-3 py-2 text-sm text-white">
                  <input type="checkbox" checked={snakeAlternates} onChange={() => setSnakeAlternates((prev) => !prev)} />
                  <span>Snake / alternate</span>
                </label>
                <Button intent="primary" onClick={snakePatch}><Wand2 className="h-4 w-4" />Auto Snake</Button>
                <Button intent="secondary" onClick={clearSelectedPortPatching}>
                  Clear Selected {patchMode === "signal" ? (activePort > 0 ? `Port ${activePort}` : "Port") : (activePowerPort > 0 ? `Plug ${activePowerPort}` : "Plug")}
                </Button>
                <Button intent="danger" onClick={clearSignalCabling}>Clear Signal</Button>
                <Button intent="danger" onClick={clearPowerAssignments}>Clear Power</Button>
              </ControlGroup>
            </CardContent>
          </Card>

          <div className="space-y-4">
          <Card className="border-slate-700 bg-slate-800 print-card" collapsible>
            <CardHeader>
              <CardTitle className="text-white [text-shadow:0_0_2px_black]">Wall Summary</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4 text-sm text-white [text-shadow:0_0_2px_black]">
              <div className="grid gap-4 md:grid-cols-3">
                <div className="rounded border border-slate-700 bg-slate-900 p-3">
                  <div className="mb-2 font-bold">Wall Details</div>
                  <div>Panels: {totalPanels} active across {panelBands.length} row band{panelBands.length === 1 ? "" : "s"}</div>
                  <div>{wallSizeLabel}: {formatMeters(wallWidthM)}m × {formatMeters(wallHeightM)}m</div>
                  <div>Area: {formatNumber(wallWidthM * wallHeightM, 1)} m²</div>
                  {isMtOnlyWall ? (
                    <>
                      <div>LED Wall Resolution: {wallPixelW} × {wallPixelH}</div>
                      <div>Recommended Content Resolution: {contentPixelW} × {contentPixelH}</div>
                      <div>Physical Aspect Ratio: {physicalRatioLabel}</div>
                    </>
                  ) : (
                    <>
                      <div>Resolution: {wallPixelW} × {wallPixelH}</div>
                      <div>Aspect: {aspectRatio}</div>
                      <div>Ratio: {ratioLabel}</div>
                    </>
                  )}
                </div>
                <div className="rounded border border-slate-700 bg-slate-900 p-3">
                  <div className="mb-2 font-bold">{panel.name} Guts</div>
                  <div>Panel size: {panel.w}m × {panel.h}m</div>
                  <div>Pixels per panel: {formatNumber(panelPixels)}</div>
                  <div>Weight per panel: {panel.weight} kg</div>
                  <div>Max power: {powerSpec.maxW} W / {powerSpec.maxA} A</div>
                  <div>Avg power: {powerSpec.avgW} W / {powerSpec.avgA} A</div>
                </div>
                <div className="rounded border border-slate-700 bg-slate-900 p-3">
                  <div className="mb-2 font-bold">Signal + Output</div>
                  <div>Ports used: {effectiveSignalPortsUsed}{backupSignalLoop ? ` (${signalPortsUsed} main + ${signalPortsUsed} backup)` : ""}</div>
                  <div>Pixels per port: {formatNumber(panelPixels)}</div>
                  <div>Port capacity use: {formatNumber((wallPixelW * wallPixelH) / Math.max(signalPortsUsed, 1), 0)} px avg</div>
                  <div>VX1000 max use: {formatNumber(vx1000Percent, 1)}%</div>
                  <div>VX2000 max use: {formatNumber(vx2000Percent, 1)}%</div>
                </div>
              </div>

              <div className="grid gap-4 md:grid-cols-3">
                <div className="rounded border border-slate-700 bg-slate-900 p-3">
                  <div className="mb-2 font-bold">Weight</div>
                  <div>Panel weight: {panelOnlyWeight.toFixed(1)} kg</div>
                  <div>Additional subtotal: {additionalWeight.toFixed(1)} kg</div>
                  <div className="font-semibold">Total weight: {totalWeight.toFixed(1)} kg</div>
                </div>
                <div className="rounded border border-slate-700 bg-slate-900 p-3">
                  <div className="mb-2 font-bold">Power</div>
                  <div>Max: {formatNumber(totalPowerMaxW, 0)} W / {formatNumber(totalPowerMaxA, 2)} A</div>
                  <div>Avg: {formatNumber(totalPowerAvgW, 0)} W / {formatNumber(totalPowerAvgA, 2)} A</div>
                  <div>Circuits used (max): {circuitsUsedMax}</div>
                  <div>Per outlet: {formatNumber(powerPerCircuitMaxW, 0)} W / {formatNumber(powerPerCircuitMaxA, 2)} A</div>
                  <div>Active support span: {activeColsCount} cols × {activeRowsCount} rows</div>
                </div>
                <div className="rounded border border-slate-700 bg-slate-900 p-3">
                  <div className="mb-2 font-bold">Best Standard Output</div>
                  {bestResolution ? (
                    <>
                      <div>{bestResolution[0]} × {bestResolution[1]}</div>
                      <div>Wall uses {formatNumber(((wallPixelW * wallPixelH) / (bestResolution[0] * bestResolution[1])) * 100, 1)}%</div>
                      <div>Spare output: {formatNumber(100 - ((wallPixelW * wallPixelH) / (bestResolution[0] * bestResolution[1])) * 100, 1)}%</div>
                    </>
                  ) : (
                    <div>No standard size in preset list fits this wall.</div>
                  )}
                </div>
              </div>

              <div className="grid gap-4 border-t border-slate-700 pt-3 no-print lg:grid-cols-2">
                <div className="space-y-2">
                  <div className="font-bold">Additional Weights</div>

                  <label className="flex items-center gap-2">
                    <input type="checkbox" checked={includeFlyBar} onChange={() => setIncludeFlyBar(!includeFlyBar)} />
                    <span>Fly Bar (per top-row panel: MG9 {PANEL_TYPES.MG9.defaults.flyBarWeight}kg / MT {PANEL_TYPES.MT.defaults.flyBarWeight}kg) → {flyBarWeight.toFixed(1)} kg</span>
                  </label>

                  <label className="flex items-center gap-2">
                    <input type="checkbox" checked={includeSling} onChange={() => setIncludeSling(!includeSling)} />
                    <span>Sling &amp; Shackle ({PANEL_TYPES.MG9.defaults.slingWeight}kg per top-row panel) → {slingWeight.toFixed(1)} kg</span>
                  </label>

                  <label className="flex items-center gap-2">
                    <input type="checkbox" checked={includePowerCable} onChange={() => setIncludePowerCable(!includePowerCable)} />
                    <span>Power cables (3kg per outlet used) → {powerCableWeight.toFixed(1)} kg</span>
                  </label>

                  <label className="flex items-center gap-2">
                    <input type="checkbox" checked={includeSignalCable} onChange={() => setIncludeSignalCable(!includeSignalCable)} />
                    <span>Signal cables (1kg per signal port used) → {signalCableWeight.toFixed(1)} kg</span>
                  </label>

                  <div className="flex items-center gap-2">
                    <input type="checkbox" checked={includeCustomWeight} onChange={() => setIncludeCustomWeight(!includeCustomWeight)} />
                    <span>Custom Weight</span>
                    <input type="number" className="w-24 rounded bg-white p-1 text-black" value={customWeight} onChange={(e) => setCustomWeight(Number(e.target.value))} />
                    <span>kg</span>
                  </div>
                </div>

                <div className="space-y-3 rounded border border-slate-700 bg-slate-900 p-3">
                  <div className="font-bold">LED Wall Deployment Settings</div>
                  <label className="flex items-center gap-2">
                    <input type="checkbox" checked={backupSignalLoop} onChange={() => setBackupSignalLoop((prev) => !prev)} />
                    <span>Do backup signal loop</span>
                  </label>
                  <label className="flex items-center gap-2">
                    <input type="checkbox" checked={includeReinforcementPlate} onChange={() => setIncludeReinforcementPlate((prev) => !prev)} />
                    <span>Reinforcement Plate</span>
                  </label>
                  <div className="space-y-1">
                    <div className="text-xs text-slate-300">Type of deployment</div>
                    <select className="w-full rounded bg-white p-2 text-black" value={deploymentType} onChange={(e) => setDeploymentType(e.target.value as DeploymentType | "")}>
                      <option value="">Select deployment type</option>
                      <option value={DEPLOYMENT_TYPES.FLOWN}>{DEPLOYMENT_TYPES.FLOWN}</option>
                      <option value={DEPLOYMENT_TYPES.GROUND}>{DEPLOYMENT_TYPES.GROUND}</option>
                      <option value={DEPLOYMENT_TYPES.NO_SUPPORT}>{DEPLOYMENT_TYPES.NO_SUPPORT}</option>
                      <option value={DEPLOYMENT_TYPES.FLOOR}>{DEPLOYMENT_TYPES.FLOOR}</option>
                    </select>
                  </div>
                  {deploymentWarning ? (
                    <div className="rounded border border-amber-500/40 bg-amber-500/10 p-2 text-xs text-amber-200">
                      {deploymentWarning}
                    </div>
                  ) : null}
                </div>
              </div>

              <div className="grid gap-2 pt-2 md:grid-cols-3">
                {Object.entries(phaseStats).map(([phase, stat]) => (
                  <div key={phase} className="rounded border border-slate-700 bg-slate-900 p-2 text-xs text-white [text-shadow:0_0_2px_black]">
                    <div className="font-medium">{`Phase ${phase.replace("P", "")}`}</div>
                    <div>{formatNumber(stat.maxWatts, 0)} W / {formatNumber(stat.maxAmps, 2)} A</div>
                    <div>Avg {formatNumber(stat.avgWatts, 0)} W / {formatNumber(stat.avgAmps, 2)} A</div>
                    <UtilBar percent={stat.utilisation} />
                    <div>Safe phase limit: {formatNumber(distro.safePhaseWatts, 0)} W</div>
                  </div>
                ))}
              </div>
            </CardContent>
          </Card>

          <SubScreenPanel
            subScreens={subScreens}
            activeSubScreenId={resolvedActiveSubScreenId}
            grid={grid}
            onSelectScreen={selectSubScreen}
            onCreate={createSubScreen}
            onRename={renameSubScreen}
            onDelete={deleteSubScreen}
            onSelectAllInSubScreen={selectAllInSubScreen}
          />
          </div>
        </div>

        <Card className="border-slate-700 bg-slate-800 print-card" data-panel-layout collapsible>
          <CardHeader>
            <div className="flex flex-wrap items-center justify-between gap-3">
              <CardTitle className="text-white [text-shadow:0_0_2px_black]">Panel Layout ({formatMeters(wallWidthM)}m x {formatMeters(wallHeightM)}m) - {patchMode === "signal" ? "Signal" : "Power"} patching</CardTitle>
            </div>
            {subScreens.length > 0 ? (
              <div className="mt-2 flex flex-wrap items-center gap-2 text-xs no-print">
                {resolvedActiveSubScreenId !== null ? (
                  <>
                    <StatusChip tone="sky">
                      Editing sub-screen: {subScreens.find((s) => s.id === resolvedActiveSubScreenId)?.name ?? "Unknown"}
                    </StatusChip>
                    <span className="text-slate-300">Other panels are hidden while a sub-screen is active.</span>
                    <Button intent="ghost" size="sm" onClick={(e) => { e.stopPropagation(); selectSubScreen(null); }}>
                      All Screens
                    </Button>
                  </>
                ) : (
                  <StatusChip tone="emerald">All Screens - showing the complete layout</StatusChip>
                )}
              </div>
            ) : null}
          </CardHeader>
          <CardContent>
            <div className="mb-3 flex flex-wrap items-start gap-2 text-xs text-white [text-shadow:0_0_2px_black] no-print">
              <ControlGroup label="Panel tools">
                <Button
                  intent="secondary"
                  size="sm"
                  active={editMode === "patch"}
                  activeAccent="sky"
                  onClick={() => setEditMode("patch")}
                  title="Click or drag panels to patch the active signal port / power plug"
                >
                  Patch
                </Button>
                <Button
                  intent="secondary"
                  size="sm"
                  active={editMode === "select"}
                  activeAccent="emerald"
                  onClick={() => {
                    setEditMode((prev) => {
                      if (prev === "select") {
                        setSelectedId(null);
                        setSelectedCells(new Set());
                        return "patch";
                      }
                      return "select";
                    });
                  }}
                  title="Click panels or drag a box to select (Shift adds)"
                >
                  Select
                </Button>
                <Button
                  intent="secondary"
                  size="sm"
                  active={editMode === "move"}
                  activeAccent="amber"
                  onClick={() => setEditMode((prev) => (prev === "move" ? "patch" : "move"))}
                  title="Drag panels to reposition them freely; edges snap together"
                >
                  Move
                </Button>
                <Button intent="secondary" size="sm" onClick={clearSelectedPanelPatching} disabled={selectedCount === 0}>Clear Patching</Button>
                <StatusChip tone="emerald">{selectedCount ? `${selectedCount} selected` : "None selected"}</StatusChip>
                {editMode === "move" ? (
                  <>
                    <label className="flex items-center gap-1 rounded border border-slate-600 bg-slate-800 px-2 py-1">
                      <input type="checkbox" checked={snapEnabled} onChange={() => setSnapEnabled((prev) => !prev)} />
                      <span>Snap</span>
                    </label>
                    <label className="flex items-center gap-1 rounded border border-slate-600 bg-slate-800 px-2 py-1">
                      <input type="checkbox" checked={moveJoinedGroup} onChange={() => setMoveJoinedGroup((prev) => !prev)} />
                      <span>Move joined group</span>
                    </label>
                    <label className="flex items-center gap-1 rounded border border-slate-600 bg-slate-800 px-2 py-1" title="Permit intentional panel overlaps">
                      <input type="checkbox" checked={allowOverlaps} onChange={() => setAllowOverlaps((prev) => !prev)} />
                      <span>Allow overlaps</span>
                    </label>
                  </>
                ) : null}
              </ControlGroup>

              <ControlGroup label="Selection & editing">
                <Button
                  intent="secondary"
                  size="sm"
                  onClick={() => {
                    commitGridUpdate((prev) => [
                      ...prev,
                      makePanelAt(wallBBox.x, wallBBox.y + wallBBox.h + MODULE_MM, panelType, resolvedActiveSubScreenId),
                    ]);
                    setEditMode("move");
                  }}
                  title="Add a new panel below the wall, ready to move into place"
                >
                  + Add Panel
                </Button>
                <select
                  className="rounded-lg border border-slate-500 bg-white p-2 text-sm text-black disabled:opacity-60"
                  disabled={selectedCount === 0}
                  title="Set the panel type for the selected panels (MT spans two 0.5m modules)"
                  value={selectedPanel ? cellPanelType(selectedPanel) : "MG9"}
                  onChange={(e) => applySelectedPanelType(e.target.value as PanelTypeKey)}
                >
                  {(Object.keys(PANEL_TYPES) as PanelTypeKey[]).map((key) => (
                    <option key={key} value={key}>{PANEL_TYPES[key].name} panel</option>
                  ))}
                </select>
                <select
                  className="rounded-lg border border-slate-500 bg-white p-2 text-sm text-black disabled:opacity-60"
                  disabled={selectedCount === 0 || (selectedPanel ? cellPanelType(selectedPanel) !== "MG9" : false)}
                  value={selectedPanel?.panelVariant ?? "STANDARD"}
                  onChange={(e) => applySelectedPanelVariant(e.target.value as PanelVariantKey)}
                >
                  {(Object.keys(PANEL_VARIANTS) as PanelVariantKey[]).map((key) => (
                    <option key={key} value={key}>{PANEL_VARIANTS[key].label}</option>
                  ))}
                </select>
                <Button intent="secondary" size="sm" onClick={copySelectedPanels} disabled={selectedCount === 0} title="Copy selected panels (Ctrl/Cmd+C)">Copy</Button>
                <Button
                  intent={isPasting ? "primary" : "secondary"}
                  size="sm"
                  onClick={() => (isPasting ? cancelPaste() : startPaste())}
                  disabled={!clipboard}
                  title={isPasting ? "Click on the layout to place, Esc/right-click to cancel" : "Paste copied panels (Ctrl/Cmd+V)"}
                >
                  {isPasting ? "Click to Place…" : "Paste"}
                </Button>
                <Button intent="danger" size="sm" onClick={deleteSelectedPanel} disabled={selectedCount === 0}>Delete</Button>
                <Button intent="success" size="sm" onClick={restoreSelectedPanel} disabled={selectedCount === 0}>Restore</Button>
                <Button intent="ghost" size="sm" onClick={undoLayout} disabled={!undoStack.length}><Undo2 className="h-4 w-4" />Undo</Button>
                <Button intent="ghost" size="sm" onClick={redoLayout} disabled={!redoStack.length}><Redo2 className="h-4 w-4" />Redo</Button>
                {subScreens.length > 0 ? (
                  <>
                    <select
                      className="rounded-lg border border-slate-500 bg-white p-2 text-sm text-black disabled:opacity-60"
                      value={assignTargetSubScreenId}
                      onChange={(e) => setAssignTargetSubScreenId(e.target.value)}
                      disabled={selectedCount === 0}
                      title="Choose which sub-screen to assign the selected panels to"
                    >
                      <option value="">Choose a sub-screen...</option>
                      {subScreens.map((screen) => (
                        <option key={screen.id} value={screen.id}>
                          {screen.name}
                        </option>
                      ))}
                    </select>
                    <Button
                      intent="secondary"
                      size="sm"
                      onClick={() => {
                        if (assignTargetSubScreenId) assignSelectedToSubScreen(assignTargetSubScreenId);
                      }}
                      disabled={selectedCount === 0 || !assignTargetSubScreenId}
                    >
                      Assign Selected ({selectedCount})
                    </Button>
                    <Button intent="ghost" size="sm" onClick={removeSelectedFromSubScreen} disabled={selectedCount === 0}>
                      Remove from Sub-Screen
                    </Button>
                  </>
                ) : null}
              </ControlGroup>

              <ControlGroup label="Rotation & transforms">
                <Button intent="secondary" size="sm" onClick={() => rotateSelectedPanels(45)} disabled={selectedCount === 0} title="Rotate selected panels 45° clockwise">Rotate 45° 🔄</Button>
                <Button intent="secondary" size="sm" onClick={() => rotateSelectedPanels(90)} disabled={selectedCount === 0} title="Rotate selected panels 90° clockwise">Rotate 90° 🔄</Button>
                <div className="flex items-center gap-1 rounded-lg border border-slate-500 bg-white p-1">
                  <input
                    type="number"
                    className="w-16 rounded border border-slate-300 p-1 text-sm text-black"
                    value={customRotationDeg}
                    onChange={(e) => setCustomRotationDeg(e.target.value)}
                    title="Custom rotation angle in degrees"
                    disabled={selectedCount === 0}
                  />
                  <Button
                    intent="secondary"
                    size="sm"
                    onClick={() => {
                      const deg = Number.parseFloat(customRotationDeg);
                      if (Number.isFinite(deg)) rotateSelectedPanels(deg);
                    }}
                    disabled={selectedCount === 0 || !Number.isFinite(Number.parseFloat(customRotationDeg))}
                    title="Rotate selected panels by the entered angle"
                  >
                    Rotate °
                  </Button>
                </div>
              </ControlGroup>

              <ControlGroup label="View, zoom & navigation">
                <StatusChip tone={isFlippedView ? "amber" : "sky"}>{isFlippedView ? "Front View" : "Back View"}</StatusChip>
                <Button intent="secondary" size="sm" onClick={() => setIsFlippedView((prev) => !prev)}>
                  {isFlippedView ? "Show Back View" : "Show Front View"}
                </Button>
                <select
                  className="rounded-lg border border-slate-500 bg-white p-1.5 text-xs text-black"
                  value={String(zoom)}
                  onChange={(e) => setZoom(Number(e.target.value))}
                  title="Workspace zoom"
                >
                  <option value="0.5">50%</option>
                  <option value="0.75">75%</option>
                  <option value="1">100%</option>
                  <option value="1.5">150%</option>
                </select>
                <Button
                  intent="secondary"
                  size="sm"
                  onClick={fitToView}
                  title="Zoom so the entire layout - including any wide/tall imported project - fits in the visible workspace"
                >
                  Fit to View
                </Button>
              </ControlGroup>

              <ControlGroup label="Overlays & display">
                <Button
                  intent={showCentreLine ? "primary" : "secondary"}
                  size="sm"
                  onClick={() => setShowCentreLine((prev) => !prev)}
                  title="Toggle the vertical centre indicator"
                >
                  {showCentreLine ? "Hide Centre Line" : "Show Centre Line"}
                </Button>
              </ControlGroup>
            </div>
            {overlapNotice ? (
              <div className="mb-3 rounded-lg border border-amber-400 bg-amber-500/15 px-3 py-2 text-sm text-amber-200 no-print">
                ⚠ {overlapNotice}
              </div>
            ) : null}
            <div ref={workspaceViewportRef} className="w-full overflow-auto rounded-xl bg-white/5 p-4 select-none">
              <div
                ref={workspaceRef}
                className="relative"
                style={{
                  width: svgW,
                  height: svgH,
                  cursor: moveDrag ? "grabbing" : editMode === "move" ? "grab" : editMode === "select" ? "crosshair" : "pointer",
                }}
                onMouseDown={onWorkspaceMouseDown}
                onMouseMove={onWorkspaceMouseMove}
                onMouseUp={onWorkspaceMouseUp}
              >
                {/* Metre grid + ruler labels. Lines are anchored to the wall origin so
                    the 1m (major, dashed) and 0.5m (minor, fainter dashed) lines line up
                    exactly with the metre markings. No solid outer border is drawn. */}
                <svg className="absolute inset-0 z-0 pointer-events-none" width={svgW} height={svgH}>
                  {(() => {
                    const lines: React.ReactNode[] = [];
                    const kxStart = Math.floor((workspaceOrigin.x - wallBBox.x) / MODULE_MM);
                    const kxEnd = Math.ceil((workspaceOrigin.x + workspaceSizeMm.w - wallBBox.x) / MODULE_MM);
                    for (let k = kxStart; k <= kxEnd; k += 1) {
                      const x = mmToPx(wallBBox.x + k * MODULE_MM - workspaceOrigin.x);
                      const major = k % 2 === 0;
                      lines.push(
                        <line
                          key={`gv-${k}`}
                          x1={x}
                          y1={0}
                          x2={x}
                          y2={svgH}
                          stroke={major ? "rgba(148,163,184,0.38)" : "rgba(148,163,184,0.16)"}
                          strokeWidth={major ? 1.4 : 1}
                          strokeDasharray={major ? "6 4" : "2 6"}
                        />,
                      );
                    }
                    const kyStart = Math.floor((workspaceOrigin.y - wallBBox.y) / MODULE_MM);
                    const kyEnd = Math.ceil((workspaceOrigin.y + workspaceSizeMm.h - wallBBox.y) / MODULE_MM);
                    for (let k = kyStart; k <= kyEnd; k += 1) {
                      const y = mmToPx(wallBBox.y + k * MODULE_MM - workspaceOrigin.y);
                      const major = k % 2 === 0;
                      lines.push(
                        <line
                          key={`gh-${k}`}
                          x1={0}
                          y1={y}
                          x2={svgW}
                          y2={y}
                          stroke={major ? "rgba(148,163,184,0.38)" : "rgba(148,163,184,0.16)"}
                          strokeWidth={major ? 1.4 : 1}
                          strokeDasharray={major ? "6 4" : "2 6"}
                        />,
                      );
                    }
                    return lines;
                  })()}
                  {Array.from({ length: Math.floor(wallBBox.w / 1000) + 1 }).map((_, m) => (
                    <text
                      key={`rx-${m}`}
                      x={mmToPx(wallBBox.x + m * 1000 - workspaceOrigin.x)}
                      y={mmToPx(wallBBox.y - workspaceOrigin.y) - 8}
                      fill="#94a3b8"
                      fontSize="10"
                      textAnchor="middle"
                    >
                      {m}m
                    </text>
                  ))}
                  {/* Height ruler reads bottom-up (0m at the wall's base, increasing
                      upward) to match how a physical wall is measured/built - the line
                      positions themselves are unchanged, only the printed label. */}
                  {Array.from({ length: Math.floor(wallBBox.h / 1000) + 1 }).map((_, m) => (
                    <text
                      key={`ry-${m}`}
                      x={mmToPx(wallBBox.x - workspaceOrigin.x) - 10}
                      y={mmToPx(wallBBox.y + m * 1000 - workspaceOrigin.y) + 3}
                      fill="#94a3b8"
                      fontSize="10"
                      textAnchor="end"
                    >
                      {Math.floor(wallBBox.h / 1000) - m}m
                    </text>
                  ))}
                </svg>

                {/* Sub-screen boundary outlines + name labels. Purely a visual aid -
                    derived from panel positions, never obscures panel content since it
                    sits below the panel layer (z-10+). The actively-edited sub-screen (if
                    any) is drawn solid/bright; others are faint, matching the panel
                    dimming treatment so the visual language stays consistent. */}
                {subScreens.length ? (
                  <svg className="absolute inset-0 z-[2] pointer-events-none" width={svgW} height={svgH}>
                    {subScreens.map((screen, index) => {
                      const bbox = subScreenBBoxes.get(screen.id);
                      if (!bbox) return null;
                      const displayBBox = isFlippedView ? mirrorRectX(bbox, wallBBox) : bbox;
                      const r = rectToPx(displayBBox);
                      const isActive = resolvedActiveSubScreenId === screen.id;
                      const isOtherActive = resolvedActiveSubScreenId !== null && !isActive;
                      const color = PORT_COLORS[index % PORT_COLORS.length];
                      return (
                        <g key={screen.id} opacity={isOtherActive ? 0.35 : 1}>
                          <rect
                            x={r.x - 6}
                            y={r.y - 6}
                            width={r.w + 12}
                            height={r.h + 12}
                            fill="none"
                            stroke={color}
                            strokeWidth={isActive ? 2.5 : 1.5}
                            strokeDasharray={isActive ? undefined : "6 5"}
                            rx={6}
                          />
                          <text x={r.x - 4} y={r.y - 12} fill={color} fontSize="12" fontWeight="bold">
                            {screen.name}
                          </text>
                        </g>
                      );
                    })}
                  </svg>
                ) : null}

                {/* Cable LINES sit behind the panels (z-[1], panels are z-10+) so they
                    never obscure panel labels. Each line is drawn twice: a wider black
                    stroke underneath for a thin outline, then the coloured stroke on top.
                    Arrowheads are drawn separately in front of the panels (below). */}
                <svg className="absolute inset-0 z-[1] pointer-events-none" width={svgW} height={svgH}>
                  {cableHops.map((hop) => {
                    const pointStr = hop.pts.map((p) => `${p.x},${p.y}`).join(" ");
                    return (
                      <g key={`line-${hop.key}`}>
                        <polyline
                          points={pointStr}
                          fill="none"
                          stroke="black"
                          strokeWidth="6"
                          strokeLinejoin="round"
                          strokeLinecap="round"
                        />
                        <polyline
                          points={pointStr}
                          fill="none"
                          stroke={hop.color}
                          strokeWidth="4"
                          strokeLinejoin="round"
                          strokeLinecap="round"
                        />
                      </g>
                    );
                  })}
                </svg>

                {grid.map((cell) => {
                  // Sub-screen isolation: while a sub-screen is active, every panel not
                  // assigned to it (including unassigned panels - reassignment is done
                  // from All Screens via the "Assign Selected" dropdown, not while the
                  // target screen is active) is fully hidden, not just dimmed, so it
                  // can't be selected/moved/removed/patched by accident.
                  const isDimmed = isPanelDimmed(cell);
                  if (isDimmed) return null;
                  const rect = rectToPx(displayRectOf(cell));
                  const isMoving = !!moveDrag && moveDrag.ids.includes(cell.id);
                  const signalStat = cell.assignedPort ? signalPortStats[cell.assignedPort] : null;
                  const isEdge = signalStat?.firstKey === cell.id || signalStat?.lastKey === cell.id;
                  const { signalBadges, powerBadge } = getPanelIndicators(cell);
                  const isSelected = selectedCells.has(cell.id) || selectedId === cell.id;
                  const isRemoved = cell.isRemoved;
                  const displayColor = isRemoved ? "transparent" : cell.assignedPort ? PORT_COLORS[(cell.assignedPort - 1) % PORT_COLORS.length] : "#1e293b";
                  const variant = PANEL_VARIANTS[cell.panelVariant ?? "STANDARD"];
                  // Match the canvas/PDF base shapes (and the YES TECH layout
                  // tool): triangle = right angle at bottom-left at rotation 0;
                  // curve = quarter disc centred on the bottom-right corner.
                  const shapeClipPath =
                    variant.shape === "triangle"
                      ? "polygon(0 0, 100% 100%, 0 100%)"
                      : variant.shape === "curve"
                        ? "circle(farthest-side at 100% 100%)"
                        : undefined;
                  const hatch =
                    variant.shape === "corner"
                      ? `repeating-linear-gradient(135deg, transparent 0 6px, rgba(15,23,42,0.35) 6px 8px), ${displayColor}`
                      : displayColor;

                  return (
                    <div
                      key={cell.id}
                      onMouseDown={(event) => {
                        event.stopPropagation();
                        onPanelMouseDown(cell, event);
                      }}
                      onMouseEnter={() => onPanelMouseEnter(cell)}
                      style={{
                        position: "absolute",
                        left: rect.x,
                        top: rect.y,
                        width: rect.w,
                        height: rect.h,
                        zIndex: isMoving ? 30 : isSelected ? 25 : 10,
                        opacity: isRemoved ? undefined : isMoving ? 0.85 : 1,
                        background: "transparent",
                        border: `2px ${isRemoved ? "dashed" : "solid"} ${isMoving ? "#fbbf24" : isSelected ? "#ffffff" : isRemoved ? "#64748b" : "transparent"}`,
                        boxShadow: "none",
                        color: isRemoved ? "#94a3b8" : "#020617",
                      }}
                      className="flex cursor-pointer select-none flex-col items-center justify-end gap-[2px] p-1 text-[9px] font-semibold leading-tight tracking-tight"
                    >
                      {isRemoved ? (
                        null
                      ) : (
                        <>
                          <div
                            className="absolute inset-0"
                            style={{
                              background: hatch,
                              border: `2px solid ${isEdge ? "black" : "#334155"}`,
                              clipPath: shapeClipPath,
                              // Front view mirrors the whole wall horizontally: flip each
                              // panel's shape (scaleX -1) around the rotated shape, without
                              // changing its stored rotation. Labels stay un-mirrored.
                              transform: `${isFlippedView ? "scaleX(-1) " : ""}rotate(${cell.rotation ?? 0}deg)`,
                              transformOrigin: "center",
                            }}
                          />
                          {/* Chain-start indicator rings that follow the true panel shape
                              (triangle / curve / rect), mirrored with the panel in the front
                              view - shown together with the port-number badges below (both
                              requested). Blue = signal chain start / backup end; orange = power
                              chain start, drawn just inside so both stay visible together. */}
                          {signalBadges.length || powerBadge ? (
                            <svg
                              className="pointer-events-none absolute inset-0 z-[6]"
                              width="100%"
                              height="100%"
                              viewBox="0 0 100 100"
                              preserveAspectRatio="none"
                              style={{
                                overflow: "visible",
                                transform: `${isFlippedView ? "scaleX(-1) " : ""}rotate(${cell.rotation ?? 0}deg)`,
                                transformOrigin: "center",
                                printColorAdjust: "exact",
                                WebkitPrintColorAdjust: "exact",
                              }}
                            >
                              {powerBadge ? (
                                <path
                                  d={variantOutlineSvgPath(variant.shape)}
                                  fill="none"
                                  stroke={POWER_START_COLOR}
                                  strokeWidth={6}
                                  strokeLinejoin="round"
                                  transform={signalBadges.length ? "translate(50 50) scale(0.78) translate(-50 -50)" : undefined}
                                />
                              ) : null}
                              {signalBadges.length ? (
                                <path
                                  d={variantOutlineSvgPath(variant.shape)}
                                  fill="none"
                                  stroke={SIGNAL_START_COLOR}
                                  strokeWidth={6}
                                  strokeLinejoin="round"
                                />
                              ) : null}
                            </svg>
                          ) : null}
                          {/* Port-number badges: small filled circles with the port number, all
                              in the panel's own top-left corner as actually displayed (front or
                              back view - rect/left/top already reflect whichever is showing, so
                              no extra mirroring here) - signal (blue) first, then power (orange),
                              side by side in one neatly-spaced, non-overlapping row. Deliberately
                              NOT rotated with the panel (unlike the shape-fill div above) so the
                              digit stays upright and legible on a rotated panel. A chain's first
                              panel gets its primary port number; when the backup signal loop is
                              enabled, the chain's last panel also gets a badge with the backup
                              port number (see getPanelIndicators) - a single-panel chain shows
                              both signal badges plus the power badge, all in the same row. */}
                          {signalBadges.length || powerBadge ? (() => {
                            const badgeD = Math.max(12, Math.round(Math.min(rect.w, rect.h) * 0.3));
                            const pad = Math.max(2, Math.round(badgeD * 0.18));
                            const badgeStyle: React.CSSProperties = {
                              position: "absolute",
                              top: pad,
                              width: badgeD,
                              height: badgeD,
                              borderRadius: "50%",
                              border: "1px solid #0f172a",
                              color: "#ffffff",
                              fontSize: Math.round(badgeD * 0.55),
                              fontWeight: 700,
                              display: "flex",
                              alignItems: "center",
                              justifyContent: "center",
                              lineHeight: 1,
                              zIndex: 20,
                              printColorAdjust: "exact",
                              WebkitPrintColorAdjust: "exact",
                            };
                            const cornerBadges: Array<{ color: string; text: number }> = signalBadges.map((portNum) => ({ color: SIGNAL_START_COLOR, text: portNum }));
                            if (powerBadge) cornerBadges.push({ color: POWER_START_COLOR, text: powerBadge });
                            return cornerBadges.map((b, i) => (
                              <div key={`cb-${i}`} style={{ ...badgeStyle, left: pad + i * (badgeD + pad), background: b.color }}>
                                {b.text}
                              </div>
                            ));
                          })() : null}
                          <div className="relative z-10">{`↓ ${panelRowLabel(cell)} → ${panelColLabel(cell)}`}</div>
                          {cell.assignedPort ? <div className="relative z-10 whitespace-nowrap">{`🔌 P${cell.assignedPort} (${cell.sequence ?? "-"})`}</div> : null}
                          {cell.assignedPowerPort ? <div className="relative z-10 whitespace-nowrap">{`⚡ Plug ${cell.assignedPowerPort}`}</div> : null}
                          {getPanelSymbol(cell) ? <div className="relative z-10 text-[11px]">{getPanelSymbol(cell)}</div> : null}
                        </>
                      )}
                    </div>
                  );
                })}

                {/* Cable ARROWHEADS sit in front of the panels (z-[35]) so the signal /
                    power direction stays visible even over a selected panel. */}
                <svg className="pointer-events-none absolute inset-0 z-[35]" width={svgW} height={svgH}>
                  {cableHops.map((hop) => cableArrowHead(hop))}
                </svg>

                {/* Vertical centre indicator: marks the horizontal centre of the whole
                    layout's TRUE outer bounds (trueOuterBBox - includes any panel
                    rotated to a non-cardinal angle, not just the axis-aligned wallBBox).
                    Sits above the panels (thin/dashed/translucent) so it's always
                    visible without covering panel text. */}
                {showCentreLine && trueOuterBBox.w > 0 ? (() => {
                  const centreTrueX = trueOuterBBox.x + trueOuterBBox.w / 2;
                  const centreDisplayTrueX = isFlippedView ? 2 * wallBBox.x + wallBBox.w - centreTrueX : centreTrueX;
                  const lineX = mmToPx(centreDisplayTrueX - workspaceOrigin.x);
                  const yTop = mmToPx(trueOuterBBox.y - workspaceOrigin.y);
                  const yBottom = mmToPx(trueOuterBBox.y + trueOuterBBox.h - workspaceOrigin.y);
                  return (
                    <svg className="pointer-events-none absolute inset-0 z-[36]" width={svgW} height={svgH}>
                      <line
                        x1={lineX}
                        y1={yTop}
                        x2={lineX}
                        y2={yBottom}
                        stroke="#facc15"
                        strokeWidth={1.5}
                        strokeDasharray="6 4"
                        strokeOpacity={0.75}
                      />
                      <rect x={lineX - 20} y={Math.max(0, yTop - 16)} width={40} height={14} rx={3} fill="#facc15" opacity={0.9} />
                      <text x={lineX} y={Math.max(11, yTop - 5)} textAnchor="middle" fontSize={10} fontWeight="bold" fill="#1e293b">
                        Centre
                      </text>
                    </svg>
                  );
                })() : null}

                {/* Live snap/join guide: ghost of the snapped landing position + join edges. */}
                {snapGuide ? (
                  <>
                    {snapGuide.ghosts.map((g, i) => (
                      <div
                        key={`ghost-${i}`}
                        className="pointer-events-none absolute z-40 rounded-sm border-2 border-dashed border-emerald-300"
                        style={{ left: g.x, top: g.y, width: g.w, height: g.h, background: "rgba(52,211,153,0.12)" }}
                      />
                    ))}
                    <svg className="pointer-events-none absolute inset-0 z-40" width={svgW} height={svgH}>
                      {snapGuide.edges.map((e, i) => (
                        <line key={`edge-${i}`} x1={e.x1} y1={e.y1} x2={e.x2} y2={e.y2} stroke="#34d399" strokeWidth="4" strokeLinecap="round" />
                      ))}
                    </svg>
                  </>
                ) : null}

                {/* Marquee rectangle while box-selecting. */}
                {isSelectingPanels && selectionStart && selectionEnd ? (
                  (() => {
                    const marquee = rectToPx({
                      x: Math.min(selectionStart.x, selectionEnd.x),
                      y: Math.min(selectionStart.y, selectionEnd.y),
                      w: Math.abs(selectionStart.x - selectionEnd.x),
                      h: Math.abs(selectionStart.y - selectionEnd.y),
                    });
                    return (
                      <div
                        className="pointer-events-none absolute z-40 border-2 border-emerald-300 bg-emerald-300/10"
                        style={{ left: marquee.x, top: marquee.y, width: marquee.w, height: marquee.h }}
                      />
                    );
                  })()
                ) : null}

                {/* Paste-placement mode: an invisible hit-layer captures the click that
                    commits the paste (so it works even over existing panels, which
                    normally stop propagation of their own mousedown), plus a dashed
                    preview of the copied panels following the cursor. */}
                {isPasting && clipboard ? (
                  <>
                    <div
                      className="absolute inset-0 z-[45]"
                      style={{ cursor: "copy" }}
                      onMouseMove={updatePastePreviewFromEvent}
                      onMouseDown={(event) => {
                        if (event.button !== 0) return;
                        event.preventDefault();
                        commitPaste();
                      }}
                      onContextMenu={(event) => {
                        event.preventDefault();
                        cancelPaste();
                      }}
                    />
                    {pasteAnchor
                      ? buildPastedCells(pasteAnchor.x, pasteAnchor.y, null).map((cell, i) => {
                          const rect = isFlippedView ? mirrorRectX(cellRect(cell), wallBBox) : cellRect(cell);
                          const px = rectToPx(rect);
                          return (
                            <div
                              key={`paste-preview-${i}`}
                              className="pointer-events-none absolute z-[45] rounded-sm border-2 border-dashed border-sky-300"
                              style={{
                                left: px.x,
                                top: px.y,
                                width: px.w,
                                height: px.h,
                                background: "rgba(56,189,248,0.18)",
                                transform: `rotate(${cell.rotation ?? 0}deg)`,
                                transformOrigin: "center",
                              }}
                            />
                          );
                        })
                      : null}
                  </>
                ) : null}
              </div>
            </div>

          </CardContent>
        </Card>

        <Card className="border-slate-700 bg-slate-800 print-card no-print" data-patch-picker collapsible>
          <CardHeader>
            <CardTitle className="text-white [text-shadow:0_0_2px_black]">Signal Patching</CardTitle>
            <div className="mt-1 text-xs text-slate-300">Manual assignment follows the current Panels per Signal Port maximum.</div>
          </CardHeader>
          <CardContent className="grid grid-cols-2 gap-2 md:grid-cols-5 xl:grid-cols-10">
            {signalPorts.map((port) => {
              const stat = signalPortStats[port.id];
              const loadPercent = safePanelsPerSignalPort > 0 ? (stat.panels / safePanelsPerSignalPort) * 100 : 0;
              const indicator = getStatusColor(loadPercent);
              // Reserved for the backup signal loop (the second half of the
              // port range, only when the loop is enabled) - not selectable
              // as a primary patch target; hatched to make that clear at a
              // glance, matching the port-number badge each backup panel
              // shows (see getPanelIndicators/drawPanelShape).
              const isBackupPort = backupSignalLoop && port.id > primarySignalPortCount;
              const baseBg = activePort === port.id && patchMode === "signal" ? port.color : "#1e293b";
              return (
                <div
                  key={port.id}
                  onClick={() => {
                    if (isBackupPort) return;
                    setPatchMode("signal");
                    setActivePort(port.id);
                  }}
                  className={`rounded border p-3 ${isBackupPort ? "cursor-not-allowed" : "cursor-pointer"}`}
                  style={{
                    background: isBackupPort ? `repeating-linear-gradient(135deg, transparent 0 6px, rgba(15,23,42,0.5) 6px 8px), ${baseBg}` : baseBg,
                    borderColor: port.color,
                    opacity: isBackupPort ? 0.7 : 1,
                  }}
                  title={isBackupPort ? `Reserved: backup for S${port.id - primarySignalPortCount}` : undefined}
                >
                  <div className="flex justify-between text-sm text-white [text-shadow:0_0_2px_black]">
                    <span>{`S${port.id}`}</span>
                    <span>{`${stat.panels}`}</span>
                  </div>
                  {isBackupPort ? (
                    <div className="text-[10px] text-slate-300">{`Backup for S${port.id - primarySignalPortCount}`}</div>
                  ) : null}
                  <div className="mt-2 h-2 rounded border border-white/30 bg-black/30">
                    <div style={{ width: `${Math.min(loadPercent, 100)}%`, background: indicator, height: "100%" }} />
                  </div>
                </div>
              );
            })}
          </CardContent>
        </Card>

        <Card className="border-slate-700 bg-slate-800 print-card no-print" data-patch-picker collapsible>
          <CardHeader>
            <CardTitle className="text-white [text-shadow:0_0_2px_black]">Power Outputs</CardTitle>
            <div className="mt-1 text-xs text-slate-300">Manual assignment follows the current Panels per Power Outlet maximum.</div>
          </CardHeader>
          <CardContent className="space-y-4 text-white [text-shadow:0_0_2px_black]">
            <div className="grid grid-cols-2 gap-2 md:grid-cols-4 xl:grid-cols-9">
              {powerPorts.map((port) => {
                const stat = powerPortStats[port.id];
                const indicator = getStatusColor(stat.utilisation);
                const barWidth = Math.min(stat.utilisation, 100);
                return (
                  <div
                    key={port.id}
                    onClick={() => {
                      setPatchMode("power");
                      setActivePowerPort(port.id);
                    }}
                    className="cursor-pointer rounded border p-3"
                    style={{ background: activePowerPort === port.id && patchMode === "power" ? POWER_COLOR : "#1e293b", borderColor: POWER_COLOR }}
                  >
                    <div className="flex justify-between text-sm text-white [text-shadow:0_0_2px_black]">
                      <span>{port.name}</span>
                      <span>{`${stat.panels}`}</span>
                    </div>
                    <div className="mt-1 text-[11px]">{`Phase ${port.phase.replace("P", "")}`}</div>
                    <div className="mt-1 text-[11px]">{`${formatNumber(stat.maxWatts, 0)} W / ${formatNumber(stat.maxAmps, 2)} A`}</div>
                    <div className="mt-2 h-2 rounded border border-white/30 bg-black/30">
                      <div style={{ width: `${barWidth}%`, background: indicator, height: "100%" }} />
                    </div>
                  </div>
                );
              })}
            </div>

            <div className="rounded border border-slate-700 bg-slate-900 p-3 text-sm text-white [text-shadow:0_0_2px_black]">
              <div className="font-medium">Phase Load</div>
              <div className="mt-3 grid gap-3 md:grid-cols-3">
                {Object.entries(phaseStats).map(([phase, stat]) => {
                  const indicator = getStatusColor(stat.utilisation);
                  const barWidth = Math.min(stat.utilisation, 100);
                  return (
                    <div key={phase} className="rounded border border-slate-700 p-3">
                      <div className="flex items-center justify-between">
                        <span>{`Phase ${phase.replace("P", "")}`}</span>
                        <span style={{ color: indicator }}>{formatNumber(stat.utilisation, 1)}%</span>
                      </div>
                      <div className="mt-2 text-xs">{formatNumber(stat.maxWatts, 0)} W / {formatNumber(stat.maxAmps, 2)} A</div>
                      <div className="text-[11px]">Avg {formatNumber(stat.avgWatts, 0)} W / {formatNumber(stat.avgAmps, 2)} A</div>
                      <div className="mt-2 h-2 rounded border border-white/30 bg-black/30">
                        <div style={{ width: `${barWidth}%`, background: indicator, height: "100%" }} />
                      </div>
                      <div className="mt-1 text-[11px]">Safe phase limit: {formatNumber(distro.safePhaseWatts, 0)} W</div>
                    </div>
                  );
                })}
              </div>

              {unassignedPowerPanels > 0 ? (
                <div className="mt-3 rounded border border-red-500/40 bg-red-500/10 p-2 text-xs text-red-200">
                  {`${unassignedPowerPanels} panel${unassignedPowerPanels === 1 ? "" : "s"} could not be assigned within the current power limits.`}
                </div>
              ) : null}
            </div>
          </CardContent>
        </Card>

        <OutputCanvasPanel
          outputCanvasW={outputCanvasW}
          outputCanvasH={outputCanvasH}
          onResolutionChange={updateOutputCanvasResolution}
          subScreens={subScreens}
          grid={grid}
          wholeLayoutCanvasX={wholeLayoutCanvasX}
          wholeLayoutCanvasY={wholeLayoutCanvasY}
          snapEnabled={canvasSnapEnabled}
          onToggleSnap={() => setCanvasSnapEnabled((prev) => !prev)}
          onPositionChange={updateCanvasPosition}
          processorInputs={processorModel ? PROCESSOR_SPECS[processorModel].inputs : []}
          canvasInputs={canvasInputs}
          onInputChange={updateCanvasInput}
          inputMode={inputMode}
          onInputModeChange={setInputMode}
          wholeCanvasInputId={wholeCanvasInputId}
          onWholeCanvasInputChange={setWholeCanvasInputId}
        />

        <RentmanPanel
          dateFrom={rentmanDateFrom}
          dateTo={rentmanDateTo}
          onDateFromChange={setRentmanDateFrom}
          onDateToChange={setRentmanDateTo}
          mappableItems={mappableRentmanItems}
          mapping={equipmentMapping}
          onMap={setRentmanEquipmentMapping}
          onRefresh={refreshRentmanStock}
          refreshing={rentmanRefreshing}
          refreshError={rentmanRefreshError}
          lastRefreshedAt={rentmanLastRefreshedAt}
        />

        <Card className="border-slate-700 bg-slate-800 print-card" collapsible defaultOpen={false}>
          <CardHeader>
            <div className="flex flex-wrap items-center justify-between gap-3">
              <CardTitle className="text-white [text-shadow:0_0_2px_black]">Stock Calculations</CardTitle>
              <Button variant="outline" className="no-print" onClick={(e) => { e.stopPropagation(); exportStockCsv(); }}>
                <Download className="mr-2 h-4 w-4" />Download CSV
              </Button>
            </div>
          </CardHeader>
          <CardContent className="space-y-4 text-white [text-shadow:0_0_2px_black]">
            <div className="grid gap-3 md:grid-cols-3">
              <div className="rounded border border-slate-700 bg-slate-900 p-3">
                <div className="text-xs text-slate-300">Spare ratio</div>
                <div className="text-lg font-semibold">{formatNumber(PANEL_TYPES.MG9.defaults.spareRatio * 100, 1)}%</div>
              </div>
              <div className="rounded border border-slate-700 bg-slate-900 p-3">
                <div className="text-xs text-slate-300">Total spare panels</div>
                <div className="text-lg font-semibold">{sparePanelSummary.grandTotal.spare}</div>
              </div>
              <div className="rounded border border-slate-700 bg-slate-900 p-3">
                <div className="text-xs text-slate-300">Total incl. spare</div>
                <div className="text-lg font-semibold">{sparePanelSummary.grandTotal.rounded}</div>
              </div>
            </div>

            <div className="space-y-3">
              <div className="text-xs font-semibold uppercase tracking-wide text-slate-400">
                Spare Panels by Surface{sparePanelSummary.multiSurface ? " & Type" : ""}
              </div>
              {sparePanelSummary.surfaceRows.length === 0 ? (
                <div className="text-sm text-slate-400">No panels placed yet.</div>
              ) : (
                sparePanelSummary.surfaceRows.map((surface) => (
                  <div key={surface.name} className="overflow-x-auto rounded border border-slate-700">
                    <table className="min-w-full table-fixed text-left text-sm">
                      <thead className="bg-slate-900">
                        {sparePanelSummary.multiSurface ? (
                          <tr>
                            <th className="px-3 py-1.5 text-xs font-semibold uppercase tracking-wide text-slate-300" colSpan={4}>{surface.name}</th>
                          </tr>
                        ) : null}
                        <tr>
                          <th className="w-40 px-3 py-2">Panel Type</th>
                          <th className="w-24 px-3 py-2 text-right">Used</th>
                          <th className="w-28 px-3 py-2 text-right">Spare ({formatNumber(PANEL_TYPES.MG9.defaults.spareRatio * 100, 0)}%)</th>
                          <th className="w-32 px-3 py-2 text-right">Rounded + Spare</th>
                        </tr>
                      </thead>
                      <tbody>
                        {surface.bucketRows.map((row) => (
                          <tr key={row.label} className="border-t border-slate-700">
                            <td className="px-3 py-1.5">{row.label}</td>
                            <td className="px-3 py-1.5 text-right">{formatNumber(row.used)}</td>
                            <td className="px-3 py-1.5 text-right">{formatNumber(row.spare)}</td>
                            <td className="px-3 py-1.5 text-right font-semibold">{formatNumber(row.rounded)}</td>
                          </tr>
                        ))}
                        <tr className="border-t border-slate-600 bg-slate-900/60 font-semibold">
                          <td className="px-3 py-1.5">Subtotal</td>
                          <td className="px-3 py-1.5 text-right">{formatNumber(surface.subtotal.used)}</td>
                          <td className="px-3 py-1.5 text-right">{formatNumber(surface.subtotal.spare)}</td>
                          <td className="px-3 py-1.5 text-right">{formatNumber(surface.subtotal.rounded)}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                ))
              )}
              {sparePanelSummary.multiSurface ? (
                <div className="rounded border border-sky-700/60 bg-sky-900/20 p-3 text-sm">
                  <span className="font-semibold">Grand total:</span> {formatNumber(sparePanelSummary.grandTotal.used)} used, {formatNumber(sparePanelSummary.grandTotal.spare)} spare, {formatNumber(sparePanelSummary.grandTotal.rounded)} incl. spare
                </div>
              ) : null}
            </div>

            <div className="overflow-x-auto rounded border border-slate-700">
              <table className="min-w-full table-fixed text-left text-sm">
                <thead className="bg-slate-900">
                  <tr>
                    <th className="w-24 px-3 py-2">Code</th>
                    <th className="px-3 py-2">Item</th>
                    <th className="w-24 px-3 py-2 text-right">Required</th>
                    <th className="w-20 px-3 py-2 text-right">Spare</th>
                    <th className="w-32 px-3 py-2 text-right">Rounded + Spare</th>
                    <th className="w-20 px-3 py-2 text-right">Stock</th>
                    <th className="w-24 px-3 py-2 text-right">Net</th>
                    {hasLiveAvailableColumn ? <th className="w-28 px-3 py-2 text-right">Available (range)</th> : null}
                  </tr>
                </thead>
                <tbody>
                  {visibleStockRows.map((row) => (
                    <tr key={`${row.code}-${row.name}`} className={`border-t border-slate-700 ${row.net < 0 ? "bg-red-500/10" : ""}`}>
                      <td className={`px-3 py-2 whitespace-nowrap ${row.net < 0 ? "text-red-200" : ""}`}>{row.code}</td>
                      <td className="px-3 py-2 truncate">{row.name}</td>
                      <td className="px-3 py-2 text-right">{formatNumber(row.required)}</td>
                      <td className="px-3 py-2 text-right">{formatNumber(row.spare ?? 0)}</td>
                      <td className="px-3 py-2 text-right">{formatNumber(row.rounded ?? row.required)}</td>
                      <td className="px-3 py-2 text-right">{formatNumber(row.stock)}</td>
                      <td className={`px-3 py-2 text-right font-semibold ${row.net < 0 ? "text-red-300" : "text-emerald-300"}`}>{formatNumber(row.net)}</td>
                      {hasLiveAvailableColumn ? (
                        <td className="px-3 py-2 text-right">{typeof row.available === "number" ? formatNumber(row.available) : "-"}</td>
                      ) : null}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </CardContent>
        </Card>

        <Card className="border-slate-700 bg-slate-800 print-card" collapsible>
          <CardHeader>
            <CardTitle className="text-white [text-shadow:0_0_2px_black]">Relevant Stock / Shortfalls</CardTitle>
          </CardHeader>
          <CardContent className="space-y-3 text-white [text-shadow:0_0_2px_black]">
            {shortfallRows.length ? (
              <div className="space-y-2">
                {shortfallRows.map((row) => (
                  <div key={`short-${row.code}-${row.name}`} className="rounded border border-red-500/40 bg-red-500/10 p-3">
                    <div className="font-semibold">{row.name}</div>
                    <div className="text-sm">Need {formatNumber(row.rounded ?? row.required)}, stock {formatNumber(row.stock)}, short by {formatNumber(Math.abs(row.net))}</div>
                  </div>
                ))}
              </div>
            ) : (
              <div className="rounded border border-emerald-500/40 bg-emerald-500/10 p-3 text-emerald-200">
                No stock shortfalls detected from the current spreadsheet-style calculations.
              </div>
            )}
          </CardContent>
        </Card>

        <NovaStarExportPanel
          hasProcessorModel={Boolean(processorModel)}
          summary={novaStarValidation?.summary ?? null}
          errors={novaStarValidation?.errors ?? []}
          warnings={novaStarValidation?.warnings ?? []}
          onDownload={downloadNovaStarConfig}
          downloading={isGeneratingNovaStarFile}
        />
      </div>
    </div>
  );
}

// Basic sanity checks for core helpers
console.assert(gcd(4032, 1344) === 1344, "gcd should reduce 4032 and 1344 correctly");
console.assert(`${4032 / gcd(4032, 1344)}:${1344 / gcd(4032, 1344)}` === "3:1", "ratio reduction should produce 3:1");
console.assert(makeGridPanels(2, 3).length === 6, "makeGridPanels should build cols*rows panels");
console.assert(makeGridPanels(2, 3)[1].x === MODULE_MM, "grid panels should be on a 500mm pitch");

