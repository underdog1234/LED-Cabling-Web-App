import { Wand2, Zap, Download, Upload, FileText } from "lucide-react";
import React, { useEffect, useMemo, useRef, useState } from "react";
import { ImageDown } from "lucide-react";
import { HelpCircle, Redo2, Undo2 } from "lucide-react";
import { Button, Card, CardHeader, CardContent, CardTitle, Input, ControlGroup, StatusChip } from "./components/ui";
import {
  type RectMm,
  activeBBox,
  bandPanels,
  computeSnapDelta,
  findOverlaps,
  joinedGroupIds,
  rectsJoined,
  MODULE_MM,
} from "./model/panels";
import { parseYesTechLayout, type ImportResult } from "./import/yesTechLayout";

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
const APP_VERSION = "0.15.0";

const PANEL_TYPES = {
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

const POWER_DISTROS = {
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

const PANEL_VARIANTS = {
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

const PORT_COLORS = [
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

type PanelTypeKey = keyof typeof PANEL_TYPES;
type PanelVariantKey = keyof typeof PANEL_VARIANTS;
type PowerDistroKey = keyof typeof POWER_DISTROS;
type DeploymentType = (typeof DEPLOYMENT_TYPES)[keyof typeof DEPLOYMENT_TYPES];

type StockRow = {
  code: string;
  name: string;
  required: number;
  stock: number;
  net: number;
  method: string;
  spare?: number;
  rounded?: number;
};

// A panel in the free workspace. x/y are the TOP-LEFT corner in workspace
// millimetres - panels are no longer bound to a rows x cols grid (the grid
// generator just emits panels on a 500mm pitch). MT is a plain 1000x500mm
// record; the old head/tail module pairing exists only in legacy migration.
type Cell = {
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
};

type LayoutSnapshot = {
  panels: Cell[];
};

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
};

type OpenJsonPayload = {
  formatVersion?: number;
  projectName?: string;
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
};

const gcd = (a: number, b: number): number => {
  const absA = Math.abs(a);
  const absB = Math.abs(b);
  if (absB === 0) return absA;
  return gcd(absB, absA % absB);
};

const makeSignalPorts = () =>
  Array.from({ length: SIGNAL_PORT_COUNT }, (_, i) => ({
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

const makePanelAt = (xMm: number, yMm: number, panelType: PanelTypeKey = "MG9"): Cell => ({
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
});

// Grid generator: cols x rows of the given type on its own pitch (MG9 500mm,
// MT 1000mm wide). After generation every panel is freely movable.
const makeGridPanels = (cols: number, rows: number, panelType: PanelTypeKey = "MG9"): Cell[] => {
  const wMm = PANEL_TYPES[panelType].w * 1000;
  const hMm = PANEL_TYPES[panelType].h * 1000;
  const panels: Cell[] = [];
  for (let y = 0; y < rows; y += 1) {
    for (let x = 0; x < cols; x += 1) {
      panels.push(makePanelAt(x * wMm, y * hMm, panelType));
    }
  }
  return panels;
};

const cellPanelType = (cell: Cell): PanelTypeKey => cell.panelType ?? "MG9";

// Footprint in workspace mm, honouring rotation (90/270 swaps width/height).
const cellSizeMm = (cell: Cell) => {
  const spec = PANEL_TYPES[cellPanelType(cell)];
  const rot = ((Math.round((cell.rotation ?? 0) / 90) * 90) % 360 + 360) % 360;
  const wMm = spec.w * 1000;
  const hMm = spec.h * 1000;
  return rot === 90 || rot === 270 ? { wMm: hMm, hMm: wMm } : { wMm, hMm };
};

const cellRect = (cell: Cell): RectMm => {
  const { wMm, hMm } = cellSizeMm(cell);
  return { x: cell.x, y: cell.y, w: wMm, h: hMm };
};

// The old grid model called real panels "heads" (vs MT tail modules). In the
// free model every active record is a panel; keep the name for call sites.
const isPanelHead = (cell: Cell | null | undefined): cell is Cell => isActiveCell(cell);

const cloneGrid = (panels: Cell[]): Cell[] => panels.map((cell) => ({ ...cell }));

// Validate/repair a v2 panel list from a file.
const normalizePanels = (raw: unknown): Cell[] => {
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
    });
  });
  return panels;
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
  const filler = makePanelAt(target.x + 500, target.y, "MG9");
  return [...panels, filler];
};

const formatNumber = (value: number, digits = 0) =>
  Number(value || 0).toLocaleString(undefined, {
    maximumFractionDigits: digits,
    minimumFractionDigits: digits,
  });

const getStatusColor = (percent: number) => {
  if (percent > 100) return "#ef4444";
  if (percent >= 80) return "#f59e0b";
  return "#22c55e";
};

const clampActivePort = (value: number, max: number) => Math.min(Math.max(value, 1), max);

const clearSignalOnGrid = (panels: Cell[]) =>
  panels.map((cell) => ({
    ...cell,
    assignedPort: null,
    sequence: null,
  }));

const clearPowerOnGrid = (panels: Cell[]) =>
  panels.map((cell) => ({
    ...cell,
    assignedPowerPort: null,
    powerSequence: null,
    powerManual: false,
  }));

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
const mirrorRectX = (rect: RectMm, bbox: RectMm): RectMm => ({
  ...rect,
  x: 2 * bbox.x + bbox.w - rect.x - rect.w,
});

const makeStockRow = (
  item: { code: string; name: string; stock: number },
  required: number,
  method: string,
  spare = 0,
  rounded = required,
): StockRow => ({
  code: item.code,
  name: item.name,
  required,
  stock: item.stock,
  net: item.stock - required,
  method,
  spare,
  rounded,
});

const roundUpToBox = (value: number, boxSize = 10) => Math.ceil(Math.max(value, 0) / boxSize) * boxSize;

const getSelectedIds = (selectedCells: Set<string>, selectedId: string | null) => {
  if (selectedCells.size > 0) return selectedCells;
  return selectedId ? new Set([selectedId]) : new Set<string>();
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
  options: { hatchStep?: number; curveStyle?: "test-pattern"; signalRing?: boolean; powerRing?: boolean } = {},
) => {
  const variant = PANEL_VARIANTS[cell.panelVariant ?? "STANDARD"];
  ctx.save();
  ctx.translate(x + w / 2, y + h / 2);
  ctx.rotate(((cell.rotation ?? 0) * Math.PI) / 180);
  ctx.translate(-w / 2, -h / 2);
  ctx.fillStyle = fill;
  ctx.strokeStyle = stroke;
  ctx.lineWidth = lineWidth;

  if (variant.shape === "triangle") {
    ctx.beginPath();
    ctx.moveTo(0, 0);
    ctx.lineTo(w, h);
    ctx.lineTo(0, h);
    ctx.closePath();
    ctx.fill();
    ctx.stroke();
  } else if (variant.shape === "curve") {
    ctx.beginPath();
    if (options.curveStyle === "test-pattern") {
      ctx.moveTo(0, 0);
      ctx.lineTo(w, 0);
      ctx.quadraticCurveTo(w, h, 0, h);
    } else {
      ctx.moveTo(w, 0);
      ctx.lineTo(w, h);
      ctx.lineTo(0, h);
      ctx.quadraticCurveTo(0, 0, w, 0);
    }
    ctx.closePath();
    ctx.fill();
    ctx.stroke();
  } else {
    ctx.fillRect(0, 0, w, h);
    ctx.strokeRect(0, 0, w, h);
  }

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

  // Chain-start indicator rings, drawn on top of (and inside) the panel border so
  // they never replace it. Blue = signal chain start/backup end, orange = power
  // chain start. When both apply they nest concentrically and stay distinct.
  const rings: string[] = [];
  if (options.signalRing) rings.push(SIGNAL_START_COLOR);
  if (options.powerRing) rings.push(POWER_START_COLOR);
  if (rings.length) {
    ctx.restore();
    ctx.save();
    ctx.translate(x + w / 2, y + h / 2);
    ctx.rotate(((cell.rotation ?? 0) * Math.PI) / 180);
    ctx.translate(-w / 2, -h / 2);
    const ringW = Math.max(2, Math.round(Math.min(w, h) * 0.06));
    rings.forEach((color, i) => {
      const inset = lineWidth + ringW / 2 + i * ringW;
      ctx.strokeStyle = color;
      ctx.lineWidth = ringW;
      ctx.strokeRect(inset, inset, w - inset * 2, h - inset * 2);
    });
  }
  ctx.restore();
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
            <div><b>Ctrl+Z</b>: Undo</div>
            <div><b>Ctrl+Y</b> or <b>Ctrl+Shift+Z</b>: Redo</div>
            <div><b>Delete</b>: Delete selected panels</div>
            <div><b>R</b>: Rotate selected panels</div>
            <div><b>C</b>: Clear selected panel patching</div>
            <div><b>Escape</b>: Clear selection or leave the current mode</div>
          </div>
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

export default function App() {
  const signalPorts = useMemo(() => makeSignalPorts(), []);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const importInputRef = useRef<HTMLInputElement | null>(null);
  const workspaceRef = useRef<HTMLDivElement | null>(null);
  const [importPreview, setImportPreview] = useState<ImportResult | null>(null);

  const [projectName, setProjectName] = useState("Untitled Project");
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
  const [grid, setGrid] = useState<Cell[]>(() => makeGridPanels(24, 8));
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
  const [snapEnabled, setSnapEnabled] = useState(true);
  const [allowOverlaps, setAllowOverlaps] = useState(false);
  const [moveJoinedGroup, setMoveJoinedGroup] = useState(false);
  const [zoom, setZoom] = useState(1);
  const [overlapNotice, setOverlapNotice] = useState<string | null>(null);
  const [undoStack, setUndoStack] = useState<LayoutSnapshot[]>([]);
  const [redoStack, setRedoStack] = useState<LayoutSnapshot[]>([]);
  const [showHelp, setShowHelp] = useState(false);
  const [snakeDirection, setSnakeDirection] = useState<"LR" | "RL" | "LRB" | "RLB" | "TB" | "BT" | "LOOP_TOGETHER">("LR");
  const [snakeAlternates, setSnakeAlternates] = useState(true);
  const [isFlippedView, setIsFlippedView] = useState(false);
  const [backupSignalLoop, setBackupSignalLoop] = useState(true);
  const [includeReinforcementPlate, setIncludeReinforcementPlate] = useState(false);
  const [deploymentType, setDeploymentType] = useState<DeploymentType | "">("");

  const panel = PANEL_TYPES[panelType];
  // Workspace scale: CELL_SIZE px per 0.5m module at zoom 1.
  const pxPerMm = (CELL_SIZE / MODULE_MM) * zoom;
  const panelSelectMode = editMode === "select";
  const powerSpec = panel.power;
  const distro = POWER_DISTROS[powerDistro];
  const powerPorts = useMemo(() => makePowerPorts(distro.portCount), [distro.portCount]);

  const [panelsPerPowerOutlet, setPanelsPerPowerOutlet] = useState<number>(panel.defaults.powerPanelsPerOutlet);
  const [panelsPerSignalPort, setPanelsPerSignalPort] = useState<number>(panel.defaults.signalPanelsPerPort);

  const selectedPanel = findCellById(grid, selectedId);
  const activeSelectedKeys = getSelectedIds(selectedCells, selectedId);
  const selectedCount = activeSelectedKeys.size;
  const isPatchTargetActive = patchMode === "signal" ? activePort > 0 : activePowerPort > 0;

  const captureLayout = (): LayoutSnapshot => ({ panels: cloneGrid(grid) });
  const restoreLayout = (snapshot: LayoutSnapshot) => {
    setGrid(cloneGrid(snapshot.panels));
    setSelectedId(null);
    setSelectedCells(new Set());
    setDragVisited(new Set());
    setIsDragging(false);
    setIsSelectingPanels(false);
    setMoveDrag(null);
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
      if (event.key === "Escape") {
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
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  });

  const maxAllowedPowerPanels = 21;
  const safePanelsPerPowerOutlet = Math.min(Math.max(panelsPerPowerOutlet, 1), maxAllowedPowerPanels);
  const safePanelsPerSignalPort = Math.min(Math.max(panelsPerSignalPort, 1), panel.defaults.signalPanelsPerPort);

  const powerOutletWatts = safePanelsPerPowerOutlet * powerSpec.maxW;
  const powerOutletAmps = safePanelsPerPowerOutlet * powerSpec.maxA;
  const powerOutletPercent = (powerOutletAmps / MAX_OUTLET_AMPS) * 100;

  const panelPixels = panel.pixW * panel.pixH;
  const signalPortPixels = safePanelsPerSignalPort * panelPixels;
  const signalPortPercent = (signalPortPixels / MAX_PIXELS_PER_PORT) * 100;

  const activeCells = useMemo(() => grid.filter((cell) => !cell.isRemoved), [grid]);
  const activePanels = activeCells;
  const totalPanels = activePanels.length;
  // Wall size = bounding box of all active panels (free layouts included).
  const wallBBox = useMemo(() => activeBBox(activePanels.map(cellRect)), [activePanels]);
  const wallWidthM = wallBBox.w / 1000;
  const wallHeightM = wallBBox.h / 1000;
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

  const signalPortStats = useMemo(() => {
    const stats: Record<number, SignalPortStat> = Object.fromEntries(
      signalPorts.map((port) => [port.id, { panels: 0, path: [], firstKey: null, lastKey: null }]),
    );

    for (const cell of grid) {
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
  }, [grid, signalPorts]);

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

    for (const cell of grid) {
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
  }, [grid, powerPorts, powerSpec.maxW, powerSpec.maxA, powerSpec.avgW, powerSpec.avgA]);

  // Chain-start indicators for a panel, shared by the live layout and every export.
  // Blue ring: first panel of its signal chain (and the last panel too when the
  // backup signal loop is enabled). Orange ring: first panel of its power chain.
  const getPanelIndicators = (cell: Cell) => {
    const key = cell.id;
    const sStat = cell.assignedPort ? signalPortStats[cell.assignedPort] : null;
    const signalRing = !!sStat && (sStat.firstKey === key || (backupSignalLoop && sStat.lastKey === key));
    const pStat = cell.assignedPowerPort ? powerPortStats[cell.assignedPowerPort] : null;
    const powerRing = !!pStat && pStat.firstKey === key;
    return { signalRing, powerRing };
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
  const signalCableTotalRequired = signalCableWithBackupRequired + signalCableSpare;
  const powerCableTotalRequired = circuitsUsedMax + Math.ceil(circuitsUsedMax * panel.defaults.powerSpareRatio);
  const distroRequired = Math.max(1, Math.ceil(powerPortsUsed / distro.portCount));

  const cornerJoinStats = useMemo(() => {
    // Corner-panel joins = flush shared edges with neighbours (position-based,
    // works for free layouts too). Corner-to-corner pairs counted once.
    let cornerToFlat = 0;
    let cornerToCorner = 0;
    const corners = activeCells.filter((cell) => cell.panelVariant === "CORNER");
    corners.forEach((cell) => {
      const rect = cellRect(cell);
      activeCells.forEach((other) => {
        if (other.id === cell.id) return;
        if (!rectsJoined(rect, cellRect(other))) return;
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
    const pushBaseRow = (code: string, name: string, required: number, stockQty: number, method: string) => {
      rowsOut.push({ code, name, required, stock: stockQty, net: stockQty - required, method });
    };

    if (mg9Count > 0) {
      const standardCount = panelVariantCounts.STANDARD;
      const standardSpare = Math.ceil(standardCount * mg9Defaults.spareRatio);
      const standardRounded = roundUpToBox(standardCount + standardSpare, mg9Defaults.panelsPerBox);
      pushBaseRow("12224", "MG9 LED Panel", standardRounded, mg9StockCat.panels ?? 0, `${standardCount} + ${standardSpare} spare, rounded to box of ${mg9Defaults.panelsPerBox}`);
      rowsOut[rowsOut.length - 1].spare = standardSpare;
      rowsOut[rowsOut.length - 1].rounded = standardRounded;

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
            count + spare,
            stockQty,
            `${count} placed at this orientation + ${spare} spare`,
          );
          rowsOut[rowsOut.length - 1].spare = spare;
        });
      });

      // Corner panels are orientation-free; keep the original single line.
      {
        const item = PANEL_VARIANTS.CORNER.stockItem;
        const count = panelVariantCounts.CORNER;
        if (item && count > 0) {
          const spare = Math.ceil(count * mg9Defaults.spareRatio);
          const rounded = roundUpToBox(count + spare, mg9Defaults.panelsPerBox);
          rowsOut.push(makeStockRow(item, rounded, `${count} selected + ${spare} spare, rounded to box of ${mg9Defaults.panelsPerBox}`, spare, rounded));
        }
      }
    }

    if (mtCount > 0) {
      const mtWithSpare = mtCount + mtSpare;
      pushBaseRow("12223", "MT Mesh Panel", mtWithSpare, mtStockCat.panels ?? 0, `${mtCount} + ${mtSpare} spare`);
      rowsOut[rowsOut.length - 1].spare = mtSpare;
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

    pushBaseRow("12254", "15m PowerCON Cable", powerCableTotalRequired, stock.powerCable15m ?? 0, `${circuitsUsedMax} + ${Math.ceil(circuitsUsedMax * panel.defaults.powerSpareRatio)} spare`);
    rowsOut[rowsOut.length - 1].spare = Math.ceil(circuitsUsedMax * panel.defaults.powerSpareRatio);
    pushBaseRow("12263", "15m Signal Cable", signalCableTotalRequired, stock.signalCable15m ?? 0, `${signalCableWithBackupRequired}${backupSignalLoop ? ` (${signalCableBaseRequired} x 2 backup loop)` : ""} + ${signalCableSpare} spare`);
    rowsOut[rowsOut.length - 1].spare = signalCableSpare;

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
  }, [activeColsCount, activeRowsCount, activeWallWidthM, backupSignalLoop, circuitsUsedMax, cornerJoinStats, deploymentType, distroRequired, includeReinforcementPlate, panelVariantCounts, shapedOrientationCounts, powerCableTotalRequired, powerDistro, signalCableBaseRequired, signalCableSpare, signalCableTotalRequired, signalCableWithBackupRequired, signalPortsUsed, powerPortsUsed, distro.portCount, mg9Count, mtCount, mg9Spare, mtSpare, mg9Boxes, mtBoxes, mg9Defaults, mtDefaults, topRowBars]);

  const shortfallRows = stockRows.filter((row) => row.required > 0 && row.net < 0);
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

  const buildLayoutCanvas = (flipped = false, viewLabel = "Back View") => {
    const scale = 2;
    const px = CELL_SIZE / MODULE_MM; // export scale, independent of on-screen zoom
    const margin = 52;
    const wallW = Math.max(1, Math.round(wallBBox.w * px));
    const wallH = Math.max(1, Math.round(wallBBox.h * px));
    const canvas = document.createElement("canvas");
    canvas.width = (wallW + margin * 2) * scale;
    canvas.height = (wallH + margin * 2 + 20) * scale;
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
    for (let m = 0; m * 1000 <= wallBBox.h + 1; m += 0.5) {
      const y = m * 1000 * px;
      ctx.beginPath();
      ctx.moveTo(-4, y);
      ctx.lineTo(m % 1 === 0 ? -12 : -8, y);
      ctx.stroke();
      if (m % 1 === 0) ctx.fillText(`${m}m`, -16, y + 4);
    }

    grid.forEach((cell) => {
      if (!isPanelHead(cell)) return;
      const r = dispRectPx(cell);
      const fill = cell.assignedPort ? PORT_COLORS[(cell.assignedPort - 1) % PORT_COLORS.length] : "#1e293b";
      const { signalRing, powerRing } = getPanelIndicators(cell);
      drawPanelShape(ctx, r.x, r.y, r.w, r.h, cell, fill, "#0f172a", 2, { signalRing, powerRing });
    });

    Object.entries(signalPortStats).forEach(([portId, stat]) => {
      if (!stat.path || stat.path.length < 2) return;
      const color = PORT_COLORS[(Number(portId) - 1) % PORT_COLORS.length];
      ctx.strokeStyle = color;
      ctx.lineWidth = 4;
      ctx.lineCap = "round";
      stat.path.forEach((cell, idx) => {
        if (idx === 0) return;
        const { x1, y1, x2, y2 } = getLineEndpointsPx(dispRectPx(stat.path[idx - 1]), dispRectPx(cell), -4);
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.stroke();
        drawCanvasArrowHead(ctx, x1, y1, x2, y2, color);
      });
    });

    powerPorts.forEach((port) => {
      const stat = powerPortStats[port.id];
      const path = stat?.path ?? [];
      if (path.length < 2) return;
      ctx.strokeStyle = POWER_COLOR;
      ctx.lineWidth = 4;
      ctx.lineCap = "round";
      path.forEach((cell, idx) => {
        if (idx === 0) return;
        const { x1, y1, x2, y2 } = getLineEndpointsPx(dispRectPx(path[idx - 1]), dispRectPx(cell), 4);
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.stroke();
        drawCanvasArrowHead(ctx, x1, y1, x2, y2, POWER_COLOR);
      });
    });

    grid.forEach((cell) => {
      if (!isPanelHead(cell)) return;
      const r = dispRectPx(cell);
      const cx = r.x + r.w / 2;
      ctx.fillStyle = "#020617";
      ctx.font = "bold 10px Arial";
      ctx.textAlign = "center";
      ctx.fillText(`↓ ${panelRowLabel(cell)} → ${panelColLabel(cell)}${cellPanelType(cell) === "MT" ? " (MT)" : ""}`, cx, r.y + 18);
      if (cell.assignedPort) ctx.fillText(`🔌 P${cell.assignedPort} (${cell.sequence ?? "-"})`, cx, r.y + 34);
      if (cell.assignedPowerPort) ctx.fillText(`⚡ Plug ${cell.assignedPowerPort}`, cx, r.y + 50);
      const variantSymbol = getPanelSymbol(cell);
      if (variantSymbol) ctx.fillText(variantSymbol, cx, r.y + r.h - 6);
    });

    ctx.restore();
    return canvas;
  };

const exportJson = () => {
  try {
    // formatVersion 2: free mm-positioned panel list (v1 grids still open).
    const payload = {
      formatVersion: 2,
      appVersion: APP_VERSION,
      projectName: safeProjectName,
      panelType,
      powerDistro,
      backupSignalLoop,
      includeReinforcementPlate,
      deploymentType,
      wall: { cols, rows, widthM: wallWidthM, heightM: wallHeightM, pixelW: wallPixelW, pixelH: wallPixelH },
      panels: grid,
      patching: { signalPortsUsed, powerPortsUsed },
      stockRows,
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
      const lines = ["Code,Required", ...stockRows.map((row) => `${row.code},${row.required}`)];
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
      // Front-view pixel map. Each panel uses its own native pixel size (MG9
      // 168x168, MT 256x64) laid out left-to-right per row band; bands are
      // top-aligned and as tall as the tallest panel in the band (mixed pitches
      // aren't a single clean raster). Bands come from the panels' mm positions
      // and are mirrored for the front view.
      const rowsHeads = panelBands.map((band) =>
        band
          .map((cell) => ({ cell, leftX: mirrorRectX(cellRect(cell), wallBBox).x }))
          .sort((a, b) => a.leftX - b.leftX),
      );
      const rowHeights = rowsHeads.map((heads) => heads.reduce((m, h) => Math.max(m, PANEL_TYPES[cellPanelType(h.cell)].pixH), 0));
      const rowWidths = rowsHeads.map((heads) => heads.reduce((s, h) => s + PANEL_TYPES[cellPanelType(h.cell)].pixW, 0));

      const canvas = document.createElement("canvas");
      canvas.width = Math.max(1, ...rowWidths);
      canvas.height = Math.max(1, rowHeights.reduce((a, b) => a + b, 0));
      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("Canvas context unavailable");
      ctx.imageSmoothingEnabled = false;

      ctx.fillStyle = "#000000";
      ctx.fillRect(0, 0, canvas.width, canvas.height);

      let rowTop = 0;
      rowsHeads.forEach((heads, y) => {
        let xCursor = 0;
        heads.forEach(({ cell }) => {
          const p = PANEL_TYPES[cellPanelType(cell)];
          const x = xCursor;
          const yy = rowTop;
          const fill = cell.assignedPort ? PORT_COLORS[(cell.assignedPort - 1) % PORT_COLORS.length] : "#1e293b";
          const { signalRing, powerRing } = getPanelIndicators(cell);
          drawPanelShape(ctx, x, yy, p.pixW, p.pixH, cell, fill, "#ffffff", 1, { hatchStep: 24, curveStyle: "test-pattern", signalRing, powerRing });

          ctx.fillStyle = "#020617";
          ctx.textAlign = "center";
          ctx.font = `bold ${Math.max(12, Math.floor(p.pixH * 0.085))}px Arial`;
          ctx.fillText(`↓ ${panelRowLabel(cell)} → ${panelColLabel(cell)}`, x + p.pixW / 2, yy + p.pixH * 0.28);
          if (cell.assignedPort) ctx.fillText(`🔌 P${cell.assignedPort} (${cell.sequence ?? "-"})`, x + p.pixW / 2, yy + p.pixH * 0.5);
          if (cell.assignedPowerPort) ctx.fillText(`⚡ Plug ${cell.assignedPowerPort}`, x + p.pixW / 2, yy + p.pixH * 0.72);
          const variantSymbol = getPanelSymbol(cell);
          if (variantSymbol) {
            ctx.font = `bold ${Math.max(14, Math.floor(p.pixH * 0.12))}px Arial`;
            ctx.fillText(variantSymbol, x + p.pixW / 2, yy + p.pixH - 8);
          }
          xCursor += p.pixW;
        });
        rowTop += rowHeights[y];
      });

      const link = document.createElement("a");
      link.href = canvas.toDataURL("image/png");
      link.setAttribute("download", `${fileSafeProjectName}-${fileSafePanelType}-${cols}x${rows}-front-test-pattern.png`);
      document.body.appendChild(link);
      link.click();
      link.remove();
    } catch (err) {
      console.error("PNG test pattern failed", err);
      alert("PNG test pattern failed - check console");
    }
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
    const panels: Cell[] = result.panels.map((p) => ({
      id: newCellId(),
      x: p.x,
      y: p.y,
      assignedPort: null,
      sequence: null,
      assignedPowerPort: null,
      powerSequence: null,
      powerManual: false,
      isRemoved: false,
      panelVariant: p.panelVariant,
      rotation: p.rotation,
      panelType: p.panelType,
    }));
    setProjectName(mode === "new" ? result.projectName : result.projectName || projectName);
    setPanelType("MG9");
    setGrid(panels);
    setSelectedId(null);
    setSelectedCells(new Set());
    setUndoStack([]);
    setRedoStack([]);
    setEditMode("patch");
    setPatchMode("signal");
    setOverlapNotice(null);
    setImportPreview(null);
  };

  const generatePdf = async () => {
  try {
    const jsPDF = (await import("jspdf")).default;
    const pdf = new jsPDF({ unit: "mm", format: "a4", orientation: "landscape" });
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

      pdf.text(`Size: ${wallWidthM}m x ${wallHeightM}m`, 105, 20);
      pdf.text(`Total weight: ${totalWeight.toFixed(1)} kg`, 105, 26);
      pdf.text(`Resolution: ${wallPixelW} x ${wallPixelH}`, 105, 32);
      pdf.text(`Aspect ratio: ${aspectRatio}`, 105, 38);
      pdf.text(`Reduced ratio: ${ratioLabel}`, 105, 44);

      const usableWidth = pageWidth - 20;
      const usableHeight = pageHeight - 58;
      const layoutRatio = canvas.width / canvas.height;
      let drawWidth = usableWidth;
      let drawHeight = drawWidth / layoutRatio;
      if (drawHeight > usableHeight) {
        drawHeight = usableHeight;
        drawWidth = drawHeight * layoutRatio;
      }
      pdf.addImage(canvas.toDataURL("image/png"), "PNG", 10 + (usableWidth - drawWidth) / 2, 50 + (usableHeight - drawHeight) / 2, drawWidth, drawHeight);
    };

    const drawStockTable = (startIndex: number, startY: number, maxY: number) => {
      let y = startY;
      const drawHeader = () => {
        pdf.setFillColor(226, 232, 240);
        pdf.rect(10, y - 5, 274, 7, "F");
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(8);
        pdf.text("Code", 12, y);
        pdf.text("Item", 34, y);
        pdf.text("Required", 174, y, { align: "right" });
        pdf.text("Spare", 198, y, { align: "right" });
        pdf.text("Rounded", 226, y, { align: "right" });
        pdf.text("Stock", 252, y, { align: "right" });
        pdf.text("Net", 282, y, { align: "right" });
        y += 6;
        pdf.setFont("helvetica", "normal");
      };
      drawHeader();
      for (let index = startIndex; index < stockRows.length; index += 1) {
        const row = stockRows[index];
        if (y > maxY) return index;
        if (row.net < 0) {
          pdf.setFillColor(254, 226, 226);
          pdf.rect(10, y - 4.5, 274, 6.2, "F");
        }
        pdf.text(String(row.code), 12, y);
        pdf.text(pdf.splitTextToSize(row.name, 128)[0], 34, y);
        pdf.text(formatNumber(row.required), 174, y, { align: "right" });
        pdf.text(formatNumber(row.spare ?? 0), 198, y, { align: "right" });
        pdf.text(formatNumber(row.rounded ?? row.required), 226, y, { align: "right" });
        pdf.text(formatNumber(row.stock), 252, y, { align: "right" });
        pdf.text(formatNumber(row.net), 282, y, { align: "right" });
        y += 6;
      }
      return stockRows.length;
    };

    const drawStockPage = (startIndex = 0) => {
      pdf.addPage("a4", "landscape");
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(16);
      pdf.text(`${safeProjectName} - Stock Summary`, 10, 12);
      let nextIndex = drawStockTable(startIndex, 22, 190);
      while (nextIndex < stockRows.length) {
        pdf.addPage("a4", "landscape");
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(16);
        pdf.text(`${safeProjectName} - Stock Summary continued`, 10, 12);
        nextIndex = drawStockTable(nextIndex, 22, 190);
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
      `Size: ${wallWidthM}m x ${wallHeightM}m`,
      `Resolution: ${wallPixelW} x ${wallPixelH}`,
      `Aspect ratio: ${aspectRatio}`,
      `Reduced ratio: ${ratioLabel}`,
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
      `Boxes: ${boxCount} (${boxSparePanels} spare in boxes)`,
      `Backup signal loop: ${backupSignalLoop ? `Yes, effective signal ports ${effectiveSignalPortsUsed}` : "No"}`,
      `Reinforcement plate: ${includeReinforcementPlate ? "Yes" : "No"}`,
      `Deployment type: ${deploymentType || "Not selected"}`,
      ...(deploymentWarning ? [`Warning: ${deploymentWarning}`] : []),
    ], 220, 24, 66, 48);

    drawInfoBox("Phase Load", Object.entries(phaseStats).map(([phase, stat]) =>
      `Phase ${phase.replace("P", "")}: ${formatNumber(stat.maxWatts)} W / ${formatNumber(stat.maxAmps, 2)} A (${formatNumber(stat.utilisation, 1)}%)`
    ), 10, 78, 92, 44);

    drawInfoBox("Signal Ports In Use", usedSignalPorts.length
      ? usedSignalPorts.map((port) => {
          const stat = signalPortStats[port.id];
          return `${port.name}: ${stat.panels} panels${stat.firstKey ? `, ${stat.firstKey} -> ${stat.lastKey}` : ""}`;
        })
      : ["No signal ports in use"], 106, 78, 88, 44);

    drawInfoBox("Power Outputs In Use", usedPowerPorts.length
      ? usedPowerPorts.map((port) => {
          const stat = powerPortStats[port.id];
          return `${port.name}: ${stat.panels} panels, ${formatNumber(stat.maxWatts)} W / ${formatNumber(stat.maxAmps, 2)} A`;
        })
      : ["No power outputs in use"], 198, 78, 88, 44);

    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(12);
    pdf.text("Stock Summary", 10, 128);
    const nextStockIndex = drawStockTable(0, 138, 190);

    const backLayoutCanvas = buildLayoutCanvas(false, "Back View");
    const frontLayoutCanvas = buildLayoutCanvas(true, "Front View");
    if (nextStockIndex < stockRows.length) drawStockPage(nextStockIndex);
    drawLayoutPage(backLayoutCanvas, "Back View");
    drawLayoutPage(frontLayoutCanvas, "Front View");
    addPdfFooters();
    pdf.save(`${fileSafeProjectName}-${fileSafePanelType}-${cols}x${rows}.pdf`);
  } catch (err) {
    console.error("PDF failed", err);
    alert("PDF failed - check console");
  }
};

  const applyGridSize = () => {
    const nextCols = Number.parseInt(draftCols, 10);
    const nextRows = Number.parseInt(draftRows, 10);
    if (!Number.isFinite(nextCols) || !Number.isFinite(nextRows) || nextCols < 1 || nextRows < 1) return;

    pushUndoSnapshot();
    setCols(nextCols);
    setRows(nextRows);
    setGrid(makeGridPanels(nextCols, nextRows, panelType));
    setSelectedId(null);
    setSelectedCells(new Set());
    setDragVisited(new Set());
    setIsDragging(false);
  };

  // --- Workspace display geometry ------------------------------------------
  // Display space = workspace mm, mirrored horizontally inside the wall bbox
  // when the front view is shown. The workspace origin is the bbox corner
  // minus padding and stays fixed during a drag gesture.
  const WORKSPACE_PAD_MM = 300;
  const workspaceOrigin = { x: wallBBox.x - WORKSPACE_PAD_MM, y: wallBBox.y - WORKSPACE_PAD_MM };
  const workspaceSizeMm = { w: wallBBox.w + WORKSPACE_PAD_MM * 2, h: wallBBox.h + WORKSPACE_PAD_MM * 2 };
  const mmToPx = (mm: number) => mm * pxPerMm;
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
      const rect = isFlippedView ? mirrorRectX(cellRect(cell), wallBBox) : cellRect(cell);
      if (rectsIntersect(marquee, rect)) ids.add(cell.id);
    });
    setSelectedCells(ids);
  };

  const commitMoveDrag = () => {
    const drag = moveDrag;
    setMoveDrag(null);
    if (!drag) return;
    if (Math.abs(drag.dx) < 1 && Math.abs(drag.dy) < 1) return;
    // Display-space delta -> true mm delta (front view mirrors x).
    const trueDx = isFlippedView ? -drag.dx : drag.dx;
    const trueDy = drag.dy;
    const movingIds = new Set(drag.ids);
    const movingRects = grid
      .filter((p) => movingIds.has(p.id) && !p.isRemoved)
      .map((p) => {
        const r = cellRect(p);
        return { ...r, x: r.x + trueDx, y: r.y + trueDy };
      });
    const otherRects = grid.filter((p) => !movingIds.has(p.id) && !p.isRemoved).map(cellRect);
    const snap = computeSnapDelta(movingRects, otherRects, snapEnabled);
    const dx = trueDx + snap.dx;
    const dy = trueDy + snap.dy;
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
      setMoveDrag((prev) => (prev ? { ...prev, dx: mm.x - prev.startX, dy: mm.y - prev.startY } : prev));
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
    setIsSelectingPanels(false);
    setSelectionStart(null);
    setSelectionEnd(null);
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
    if (editMode === "move") {
      if (!isActiveCell(cell)) return;
      event.preventDefault();
      const mm = eventToDisplayMm(event);
      if (!mm) return;
      let ids: string[];
      if (activeSelectedKeys.has(cell.id) && selectedCount > 1) {
        ids = [...activeSelectedKeys].filter((id) => isActiveCell(findCellById(grid, id)));
      } else if (moveJoinedGroup) {
        ids = [...joinedGroupIds(grid, cellRect, new Set([cell.id]))];
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
    if (editMode !== "patch" || !isDragging) return;
    if (!isActiveCell(cell)) return;
    if (patchMode === "signal") assignSignalPanel(cell);
    else assignPowerPanel(cell);
  };

  const applyManualSignalPatch = (value: string) => {
    if (!selectedId) return;
    const nextPort = value === "" ? null : Number.parseInt(value, 10);
    if (nextPort !== null && (!Number.isFinite(nextPort) || nextPort < 1 || nextPort > SIGNAL_PORT_COUNT)) return;

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
      // Reading order over the free layout: row/column bands with alternation;
      // LOOP_TOGETHER returns one segment per loop, each starting a new port.
      const segments = orderPanelsForSnake(next, snakeDirection, snakeAlternates);

      if (patchMode === "signal") {
        for (const cell of next) {
          cell.assignedPort = null;
          cell.sequence = null;
        }

        let port = 1;
        let seq = 1;
        segments.forEach((segment) => {
          segment.forEach((cell) => {
            if (port > SIGNAL_PORT_COUNT) return;
            cell.assignedPort = port;
            cell.sequence = seq;
            seq += 1;
            if (seq > safePanelsPerSignalPort) {
              port += 1;
              seq = 1;
            }
          });
          // Each loop-together segment starts on a fresh port.
          if (segments.length > 1 && seq !== 1) {
            port += 1;
            seq = 1;
          }
        });
      }

      if (patchMode === "power") {
        for (const cell of next) {
          cell.assignedPowerPort = null;
          cell.powerSequence = null;
          cell.powerManual = false;
        }

        let portIndex = 0;
        segments.flat().forEach((cell) => {
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
        });
      }

      return next;
    });

    setSelectedId(null);
    setSelectedCells(new Set());
    setDragVisited(new Set());
    setIsDragging(false);
  };

  // Patch power to follow the existing signal patch: walk panels in signal order
  // (signal port, then sequence) and fill power plugs, starting a fresh plug for
  // each signal port so power plugs line up with the signal ports. Respects the
  // power panel-count and amp limits, and stops when the plugs run out.
  const matchPowerToSignal = () => {
    const hasSignal = grid.some((cell) => isActiveCell(cell) && cell.assignedPort);
    if (!hasSignal) {
      alert("Patch the signal ports first - power will follow the same pattern.");
      return;
    }

    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);

      for (const cell of next) {
        cell.assignedPowerPort = null;
        cell.powerSequence = null;
        cell.powerManual = false;
      }

      const byPort = new Map<number, Cell[]>();
      next.forEach((cell) => {
        if (!isActiveCell(cell) || !cell.assignedPort) return;
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
    commitGridUpdate((prev) => clearSignalOnGrid(prev));
    setSelectedId(null);
    setSelectedCells(new Set());
    setDragVisited(new Set());
    setIsDragging(false);
  };

  const clearPowerAssignments = () => {
    commitGridUpdate((prev) => clearPowerOnGrid(prev));
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

  const deleteSelectedPanel = () => {
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

  const rotateSelectedPanels = () => {
    const keys = getSelectedIds(selectedCells, selectedId);
    if (!keys.size) return;
    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      keys.forEach((key) => {
        const target = findCellById(next, key);
        if (!target || !isActiveCell(target)) return;
        target.rotation = ((target.rotation ?? 0) + 90) % 360;
      });
      return next;
    });
  };

  const clearSelectedPortPatching = () => {
    if ((patchMode === "signal" && activePort < 1) || (patchMode === "power" && activePowerPort < 1)) return;
    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      next.forEach((cell) => {
        if (!isActiveCell(cell)) return;
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

  return (
    <div className="min-h-screen bg-[#0f172a] p-6 text-white print-container">
      {showHelp ? <HelpModal onClose={() => setShowHelp(false)} /> : null}
      {importPreview ? (
        <ImportPreviewModal
          result={importPreview}
          hasUnsavedWork={hasUnsavedWork}
          onCancel={() => setImportPreview(null)}
          onApply={applyImport}
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
                href="https://github.com/underdog1234/LED-Cabling-Web-App#recent-changes-in-v0150"
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
                <ImageDown className="h-4 w-4" />PNG Test Pattern
              </Button>
            </div>
            <div className="flex items-center gap-2 rounded-lg border border-slate-700/70 bg-slate-900/40 p-1.5">
              <Button intent="secondary" onClick={exportJson}>
                <Download className="h-4 w-4" />Save
              </Button>
              <Button intent="secondary" onClick={() => fileInputRef.current?.click()}>
                <Upload className="h-4 w-4" />Open
              </Button>
              <Button intent="secondary" onClick={() => importInputRef.current?.click()} title="Import a project from the YES TECH Layout Tool">
                <Upload className="h-4 w-4" />Import Project
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
          <Card className="border-slate-700 bg-slate-800 print-card">
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
                  <label className="text-xs text-slate-300">Columns</label>
                  <Input className="bg-white text-black" type="number" min="1" step="1" value={draftCols} onChange={(e) => setDraftCols(e.target.value)} />
                </div>
                <div className="space-y-1">
                  <label className="text-xs text-slate-300">Rows</label>
                  <Input className="bg-white text-black" type="number" min="1" step="1" value={draftRows} onChange={(e) => setDraftRows(e.target.value)} />
                </div>
              </div>

              <div className="flex flex-wrap gap-2 no-print">
                <Button onClick={applyGridSize}>Apply Grid Size</Button>
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

          <Card className="border-slate-700 bg-slate-800 print-card">
            <CardHeader>
              <CardTitle className="text-white [text-shadow:0_0_2px_black]">Wall Summary</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4 text-sm text-white [text-shadow:0_0_2px_black]">
              <div className="grid gap-4 md:grid-cols-3">
                <div className="rounded border border-slate-700 bg-slate-900 p-3">
                  <div className="mb-2 font-bold">Wall Details</div>
                  <div>Panels: {totalPanels} active across {panelBands.length} row band{panelBands.length === 1 ? "" : "s"}</div>
                  <div>Size: {wallWidthM}m × {wallHeightM}m</div>
                  <div>Area: {formatNumber(wallWidthM * wallHeightM, 1)} m²</div>
                  <div>Resolution: {wallPixelW} × {wallPixelH}</div>
                  <div>Aspect: {aspectRatio}</div>
                  <div>Ratio: {ratioLabel}</div>
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
        </div>

        <Card className="border-slate-700 bg-slate-800 print-card" data-panel-layout>
          <CardHeader>
            <div className="flex flex-wrap items-center justify-between gap-3">
              <CardTitle className="text-white [text-shadow:0_0_2px_black]">Panel Layout ({wallWidthM}m x {wallHeightM}m) - {patchMode === "signal" ? "Signal" : "Power"} patching</CardTitle>
              <div className="flex items-center gap-2 no-print">
                <StatusChip tone={isFlippedView ? "amber" : "sky"}>{isFlippedView ? "Front View" : "Back View"}</StatusChip>
                <Button intent="secondary" size="sm" onClick={() => setIsFlippedView((prev) => !prev)}>
                  {isFlippedView ? "Show Back View" : "Show Front View"}
                </Button>
              </div>
            </div>
          </CardHeader>
          <CardContent>
            <div className="mb-3 flex flex-wrap items-center gap-2 rounded-lg border border-slate-700 bg-slate-900 p-2 text-xs text-white [text-shadow:0_0_2px_black] no-print">
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
                onClick={() => {
                  commitGridUpdate((prev) => [
                    ...prev,
                    makePanelAt(wallBBox.x, wallBBox.y + wallBBox.h + MODULE_MM, panelType),
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
              <Button intent="secondary" size="sm" onClick={rotateSelectedPanels} disabled={selectedCount === 0}>Rotate 🔄</Button>
              <Button intent="secondary" size="sm" onClick={clearSelectedPanelPatching} disabled={selectedCount === 0}>Clear Patching</Button>
              <Button intent="danger" size="sm" onClick={deleteSelectedPanel} disabled={selectedCount === 0}>Delete</Button>
              <Button intent="success" size="sm" onClick={restoreSelectedPanel} disabled={selectedCount === 0}>Restore</Button>
              <Button intent="ghost" size="sm" onClick={undoLayout} disabled={!undoStack.length}><Undo2 className="h-4 w-4" />Undo</Button>
              <Button intent="ghost" size="sm" onClick={redoLayout} disabled={!redoStack.length}><Redo2 className="h-4 w-4" />Redo</Button>
            </div>
            {overlapNotice ? (
              <div className="mb-3 rounded-lg border border-amber-400 bg-amber-500/15 px-3 py-2 text-sm text-amber-200 no-print">
                ⚠ {overlapNotice}
              </div>
            ) : null}
            <div className="w-full overflow-auto rounded-xl bg-white/5 p-4 select-none">
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
                {/* Metre grid + ruler labels (0.5m lines, metre numbers on the wall edges). */}
                <svg className="absolute inset-0 z-0 pointer-events-none" width={svgW} height={svgH}>
                  {Array.from({ length: Math.floor(workspaceSizeMm.w / MODULE_MM) + 1 }).map((_, i) => {
                    const x = mmToPx(i * MODULE_MM);
                    return <line key={`gv-${i}`} x1={x} y1={0} x2={x} y2={svgH} stroke="rgba(148,163,184,0.12)" strokeWidth="1" />;
                  })}
                  {Array.from({ length: Math.floor(workspaceSizeMm.h / MODULE_MM) + 1 }).map((_, i) => {
                    const y = mmToPx(i * MODULE_MM);
                    return <line key={`gh-${i}`} x1={0} y1={y} x2={svgW} y2={y} stroke="rgba(148,163,184,0.12)" strokeWidth="1" />;
                  })}
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
                  {Array.from({ length: Math.floor(wallBBox.h / 1000) + 1 }).map((_, m) => (
                    <text
                      key={`ry-${m}`}
                      x={mmToPx(wallBBox.x - workspaceOrigin.x) - 10}
                      y={mmToPx(wallBBox.y + m * 1000 - workspaceOrigin.y) + 3}
                      fill="#94a3b8"
                      fontSize="10"
                      textAnchor="end"
                    >
                      {m}m
                    </text>
                  ))}
                </svg>

                <svg className="absolute inset-0 z-20 pointer-events-none" width={svgW} height={svgH}>
                  <defs>
                    <marker id="arrow" markerWidth="4" markerHeight="4" refX="2" refY="2" orient="auto" markerUnits="strokeWidth">
                      <polygon points="0 0, 4 2, 0 4" fill="context-stroke" stroke="black" strokeWidth="0.4" />
                    </marker>
                  </defs>
                  {Object.entries(signalPortStats).map(([portId, stat]) => {
                    if (!stat.path || stat.path.length < 2) return null;
                    const color = PORT_COLORS[(Number(portId) - 1) % PORT_COLORS.length];
                    return stat.path.map((cell, idx) => {
                      if (idx === 0) return null;
                      const a = rectToPx(displayRectOf(stat.path[idx - 1]));
                      const b = rectToPx(displayRectOf(cell));
                      const { x1, y1, x2, y2 } = getLineEndpointsPx(a, b, -4);
                      return <line key={`sig-${portId}-${idx}`} x1={x1} y1={y1} x2={x2} y2={y2} stroke={color} style={{ color }} strokeWidth="4" markerEnd="url(#arrow)" />;
                    });
                  })}

                  {powerPorts.map((port) => {
                    const stat = powerPortStats[port.id];
                    const path = stat?.path ?? [];
                    if (path.length < 2) return null;
                    return path.map((cell, idx) => {
                      if (idx === 0) return null;
                      const a = rectToPx(displayRectOf(path[idx - 1]));
                      const b = rectToPx(displayRectOf(cell));
                      const { x1, y1, x2, y2 } = getLineEndpointsPx(a, b, 4);
                      return <line key={`pow-${port.id}-${idx}`} x1={x1} y1={y1} x2={x2} y2={y2} stroke={POWER_COLOR} style={{ color: POWER_COLOR }} strokeWidth="4" markerEnd="url(#arrow)" />;
                    });
                  })}
                </svg>

                {grid.map((cell) => {
                  const rect = rectToPx(displayRectOf(cell));
                  const isMoving = !!moveDrag && moveDrag.ids.includes(cell.id);
                  const signalStat = cell.assignedPort ? signalPortStats[cell.assignedPort] : null;
                  const isEdge = signalStat?.firstKey === cell.id || signalStat?.lastKey === cell.id;
                  const { signalRing, powerRing } = getPanelIndicators(cell);
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
                        zIndex: isMoving ? 30 : 10,
                        opacity: isMoving ? 0.85 : 1,
                        background: "transparent",
                        border: `2px ${isRemoved ? "dashed" : "solid"} ${isMoving ? "#fbbf24" : isSelected ? "#ffffff" : isRemoved ? "#64748b" : "transparent"}`,
                        boxShadow: "none",
                        color: isRemoved ? "#94a3b8" : "#020617",
                      }}
                      className="flex cursor-pointer select-none flex-col items-center justify-center gap-[2px] p-1 text-[9px] font-semibold leading-tight tracking-tight"
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
                              transform: `rotate(${cell.rotation ?? 0}deg)`,
                              transformOrigin: "center",
                            }}
                          />
                          {/* Chain-start indicators, drawn on top of the panel fill/border without replacing it. */}
                          {signalRing ? (
                            <div
                              className="pointer-events-none absolute inset-0 z-[5]"
                              style={{ border: `3px solid ${SIGNAL_START_COLOR}`, printColorAdjust: "exact", WebkitPrintColorAdjust: "exact" }}
                            />
                          ) : null}
                          {powerRing ? (
                            <div
                              className="pointer-events-none absolute z-[6]"
                              style={{ inset: 3, border: `3px solid ${POWER_START_COLOR}`, printColorAdjust: "exact", WebkitPrintColorAdjust: "exact" }}
                            />
                          ) : null}
                          <div className="relative z-10">{`↓ ${panelRowLabel(cell)} → ${panelColLabel(cell)}`}</div>
                          {cell.assignedPort ? <div className="relative z-10 whitespace-nowrap">{`🔌 P${cell.assignedPort} (${cell.sequence ?? "-"})`}</div> : null}
                          {cell.assignedPowerPort ? <div className="relative z-10 whitespace-nowrap">{`⚡ Plug ${cell.assignedPowerPort}`}</div> : null}
                          {getPanelSymbol(cell) ? <div className="relative z-10 text-[11px]">{getPanelSymbol(cell)}</div> : null}
                        </>
                      )}
                    </div>
                  );
                })}

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
              </div>
            </div>

          </CardContent>
        </Card>

        <Card className="border-slate-700 bg-slate-800 print-card no-print" data-patch-picker>
          <CardHeader>
            <CardTitle className="text-white [text-shadow:0_0_2px_black]">Signal Patching</CardTitle>
            <div className="mt-1 text-xs text-slate-300">Manual assignment follows the current Panels per Signal Port maximum.</div>
          </CardHeader>
          <CardContent className="grid grid-cols-2 gap-2 md:grid-cols-5 xl:grid-cols-10">
            {signalPorts.map((port) => {
              const stat = signalPortStats[port.id];
              const loadPercent = safePanelsPerSignalPort > 0 ? (stat.panels / safePanelsPerSignalPort) * 100 : 0;
              const indicator = getStatusColor(loadPercent);
              return (
                <div
                  key={port.id}
                  onClick={() => {
                    setPatchMode("signal");
                    setActivePort(port.id);
                  }}
                  className="cursor-pointer rounded border p-3"
                  style={{ background: activePort === port.id && patchMode === "signal" ? port.color : "#1e293b", borderColor: port.color }}
                >
                  <div className="flex justify-between text-sm text-white [text-shadow:0_0_2px_black]">
                    <span>{`S${port.id}`}</span>
                    <span>{`${stat.panels}`}</span>
                  </div>
                  <div className="mt-2 h-2 rounded border border-white/30 bg-black/30">
                    <div style={{ width: `${Math.min(loadPercent, 100)}%`, background: indicator, height: "100%" }} />
                  </div>
                </div>
              );
            })}
          </CardContent>
        </Card>

        <Card className="border-slate-700 bg-slate-800 print-card no-print" data-patch-picker>
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

        <Card className="border-slate-700 bg-slate-800 print-card">
          <CardHeader>
            <div className="flex flex-wrap items-center justify-between gap-3">
              <CardTitle className="text-white [text-shadow:0_0_2px_black]">Stock Calculations</CardTitle>
              <Button variant="outline" className="no-print" onClick={exportStockCsv}>
                <Download className="mr-2 h-4 w-4" />Download CSV
              </Button>
            </div>
          </CardHeader>
          <CardContent className="space-y-4 text-white [text-shadow:0_0_2px_black]">
            <div className="grid gap-3 md:grid-cols-4">
              <div className="rounded border border-slate-700 bg-slate-900 p-3">
                <div className="text-xs text-slate-300">Spare ratio</div>
                <div className="text-lg font-semibold">{formatNumber(panel.defaults.spareRatio * 100, 1)}%</div>
              </div>
              <div className="rounded border border-slate-700 bg-slate-900 p-3">
                <div className="text-xs text-slate-300">Spare panels</div>
                <div className="text-lg font-semibold">{sparePanels}</div>
              </div>
              <div className="rounded border border-slate-700 bg-slate-900 p-3">
                <div className="text-xs text-slate-300">Panels incl. spare</div>
                <div className="text-lg font-semibold">{totalPanelsWithSpare}</div>
              </div>
              <div className="rounded border border-slate-700 bg-slate-900 p-3">
                <div className="text-xs text-slate-300">Boxes</div>
                <div className="text-lg font-semibold">{boxCount}</div>
                <div className="text-xs text-slate-400">Box spare panels: {boxSparePanels}</div>
              </div>
            </div>

            <div className="overflow-x-auto rounded border border-slate-700">
              <table className="min-w-full table-fixed text-left text-sm">
                <thead className="bg-slate-900">
                  <tr>
                    <th className="w-24 px-3 py-2">Code</th>
                    <th className="w-28 px-3 py-2 text-right">Required</th>
                    <th className="px-3 py-2">Item</th>
                    <th className="w-28 px-3 py-2 text-right">Net Stock</th>
                  </tr>
                </thead>
                <tbody>
                  {stockRows.map((row) => (
                    <tr key={`${row.code}-${row.name}`} className={`border-t border-slate-700 ${row.net < 0 ? "bg-red-500/10" : ""}`}>
                      <td className={`px-3 py-2 whitespace-nowrap ${row.net < 0 ? "text-red-200" : ""}`}>{row.code}</td>
                      <td className="px-3 py-2 text-right">{formatNumber(row.required)}</td>
                      <td className="px-3 py-2 truncate">{row.name}</td>
                      <td className={`px-3 py-2 text-right font-semibold ${row.net < 0 ? "text-red-300" : "text-emerald-300"}`}>{formatNumber(row.net)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </CardContent>
        </Card>

        <Card className="border-slate-700 bg-slate-800 print-card">
          <CardHeader>
            <CardTitle className="text-white [text-shadow:0_0_2px_black]">Relevant Stock / Shortfalls</CardTitle>
          </CardHeader>
          <CardContent className="space-y-3 text-white [text-shadow:0_0_2px_black]">
            {shortfallRows.length ? (
              <div className="space-y-2">
                {shortfallRows.map((row) => (
                  <div key={`short-${row.code}-${row.name}`} className="rounded border border-red-500/40 bg-red-500/10 p-3">
                    <div className="font-semibold">{row.name}</div>
                    <div className="text-sm">Need {formatNumber(row.required)}, stock {formatNumber(row.stock)}, short by {formatNumber(Math.abs(row.net))}</div>
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
      </div>
    </div>
  );
}

// Basic sanity checks for core helpers
console.assert(gcd(4032, 1344) === 1344, "gcd should reduce 4032 and 1344 correctly");
console.assert(`${4032 / gcd(4032, 1344)}:${1344 / gcd(4032, 1344)}` === "3:1", "ratio reduction should produce 3:1");
console.assert(makeGridPanels(2, 3).length === 6, "makeGridPanels should build cols*rows panels");
console.assert(makeGridPanels(2, 3)[1].x === MODULE_MM, "grid panels should be on a 500mm pitch");

