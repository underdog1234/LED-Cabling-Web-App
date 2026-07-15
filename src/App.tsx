import { Wand2, Zap, Download, Upload, FileText } from "lucide-react";
import React, { useEffect, useMemo, useRef, useState } from "react";
import { ImageDown } from "lucide-react";
import { HelpCircle, Redo2, Undo2 } from "lucide-react";
import { Button, Card, CardHeader, CardContent, CardTitle, Input, ControlGroup, StatusChip } from "./components/ui";

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

type Cell = {
  x: number;
  y: number;
  assignedPort: number | null;
  sequence: number | null;
  assignedPowerPort: number | null;
  powerSequence: number | null;
  powerManual: boolean;
  isRemoved: boolean;
  panelVariant: PanelVariantKey;
  rotation: number;
  // Per-cell panel profile. The grid is a 0.5m module grid: MG9 fills one module,
  // MT fills two side-by-side modules (a "head" module plus the module to its right
  // marked as mtTail). Tail modules are drawn and patched as part of their head.
  panelType: PanelTypeKey;
  mtTail: boolean;
};

type LayoutSnapshot = {
  grid: Cell[][];
  cols: number;
  rows: number;
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

type OpenJsonPayload = {
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
  patching?: {
    grid?: Cell[][];
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

const makeGrid = (w: number, h: number): Cell[][] =>
  Array.from({ length: h }, (_, y) =>
    Array.from({ length: w }, (_, x) => ({
      x,
      y,
      assignedPort: null,
      sequence: null,
      assignedPowerPort: null,
      powerSequence: null,
      powerManual: false,
      isRemoved: false,
      panelVariant: "STANDARD",
      rotation: 0,
      panelType: "MG9",
      mtTail: false,
    })),
  );

// MG9 fills one 0.5m module. MT spans two: a head module and the module to its
// right (mtTail). These helpers keep panel-level logic readable across the app.
const cellPanelType = (cell: Cell): PanelTypeKey => cell.panelType ?? "MG9";
const cellSpanX = (cell: Cell): number => (cellPanelType(cell) === "MT" ? 2 : 1);
// A "panel head" is a real panel to count/patch/render: an active, non-tail cell.
const isPanelHead = (cell: Cell | null | undefined): cell is Cell => isActiveCell(cell) && !cell?.mtTail;

const cloneGrid = (grid: Cell[][]): Cell[][] => grid.map((row) => row.map((cell) => ({ ...cell })));

const normalizeGrid = (grid: Cell[][], cols: number, rows: number): Cell[][] =>
  Array.from({ length: rows }, (_, y) =>
    Array.from({ length: cols }, (_, x) => {
      const cell = grid?.[y]?.[x];
      return {
        x,
        y,
        assignedPort: cell?.assignedPort ?? null,
        sequence: cell?.sequence ?? null,
        assignedPowerPort: cell?.assignedPowerPort ?? null,
        powerSequence: cell?.powerSequence ?? null,
        powerManual: cell?.powerManual ?? false,
        isRemoved: cell?.isRemoved ?? false,
        panelVariant: cell?.panelVariant && PANEL_VARIANTS[cell.panelVariant] ? cell.panelVariant : "STANDARD",
        rotation: Number.isFinite(cell?.rotation) ? ((Number(cell?.rotation) % 360) + 360) % 360 : 0,
        panelType: cell?.panelType && PANEL_TYPES[cell.panelType] ? cell.panelType : "MG9",
        mtTail: Boolean(cell?.mtTail),
      };
    }),
  );

// A settings file saved before per-cell panel types existed has cells with no
// `panelType` field.
const isLegacyGrid = (rawGrid: unknown): rawGrid is Cell[][] =>
  Array.isArray(rawGrid) && rawGrid.some((row) => Array.isArray(row) && row.some((cell) => cell && (cell as Cell).panelType === undefined));

// A legacy all-MT project stored one MT panel per grid cell (1m wide). Expand it
// onto the 0.5m module grid: double the columns and pair each cell into an MT
// head + tail, carrying the panel's patching onto the head.
const expandLegacyMtGrid = (rawGrid: Cell[][], oldCols: number, rows: number) => {
  const newCols = oldCols * 2;
  const grid = makeGrid(newCols, rows);
  for (let y = 0; y < rows; y += 1) {
    for (let x = 0; x < oldCols; x += 1) {
      const src = rawGrid?.[y]?.[x];
      const head = grid[y][x * 2];
      const tail = grid[y][x * 2 + 1];
      const removed = Boolean(src?.isRemoved);
      head.assignedPort = src?.assignedPort ?? null;
      head.sequence = src?.sequence ?? null;
      head.assignedPowerPort = src?.assignedPowerPort ?? null;
      head.powerSequence = src?.powerSequence ?? null;
      head.powerManual = Boolean(src?.powerManual);
      head.isRemoved = removed;
      head.rotation = Number.isFinite(src?.rotation) ? ((Number(src?.rotation) % 360) + 360) % 360 : 0;
      head.panelType = "MT";
      head.mtTail = false;
      head.panelVariant = "STANDARD";
      tail.panelType = "MT";
      tail.mtTail = true;
      tail.isRemoved = removed;
    }
  }
  return { grid, cols: newCols };
};

const isActiveCell = (cell: Cell | null | undefined) => Boolean(cell && !cell.isRemoved);

// Convert a module at (x,y) in a cloned grid to a single MG9 panel, releasing any
// MT pairing it took part in (its own tail, or the head it was a tail of).
const setModuleToMG9 = (grid: Cell[][], x: number, y: number) => {
  const cell = grid[y]?.[x];
  if (!cell) return;
  if (cellPanelType(cell) === "MT" && !cell.mtTail) {
    const tail = grid[y]?.[x + 1];
    if (tail && tail.mtTail) {
      tail.mtTail = false;
      tail.panelType = "MG9";
    }
  }
  if (cell.mtTail) {
    const head = grid[y]?.[x - 1];
    if (head && cellPanelType(head) === "MT" && !head.mtTail) head.panelType = "MG9";
  }
  cell.panelType = "MG9";
  cell.mtTail = false;
};

// Convert the module at (x,y) into an MT head that consumes the module to its
// right as a tail. Returns false if there is no free module to the right.
const setModuleToMT = (grid: Cell[][], x: number, y: number): boolean => {
  const head = grid[y]?.[x];
  const right = grid[y]?.[x + 1];
  if (!head || !right) return false;
  if (right.isRemoved) return false;
  // Normalise both modules first so any prior MT pairing is released cleanly.
  setModuleToMG9(grid, x, y);
  setModuleToMG9(grid, x + 1, y);
  head.panelType = "MT";
  head.mtTail = false;
  right.panelType = "MT";
  right.mtTail = true;
  // The tail carries no independent patching or shape; it belongs to the head.
  right.assignedPort = null;
  right.sequence = null;
  right.assignedPowerPort = null;
  right.powerSequence = null;
  right.powerManual = false;
  right.panelVariant = "STANDARD";
  right.rotation = head.rotation ?? 0;
  return true;
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

const clearSignalOnGrid = (grid: Cell[][]) =>
  grid.map((row) =>
    row.map((cell) => ({
      ...cell,
      assignedPort: null,
      sequence: null,
    })),
  );

const clearPowerOnGrid = (grid: Cell[][]) =>
  grid.map((row) =>
    row.map((cell) => ({
      ...cell,
      assignedPowerPort: null,
      powerSequence: null,
      powerManual: false,
    })),
  );

const getNextSequence = (
  grid: Cell[][],
  portField: "assignedPort" | "assignedPowerPort",
  sequenceField: "sequence" | "powerSequence",
  portId: number,
) => {
  let max = 0;
  for (const row of grid) {
    for (const cell of row) {
      if (!isActiveCell(cell)) continue;
      if (cell[portField] === portId && (cell[sequenceField] ?? 0) > max) {
        max = cell[sequenceField] ?? 0;
      }
    }
  }
  return max + 1;
};

const getPowerPortLoadWatts = (
  grid: Cell[][],
  portId: number,
  _legacyMaxW: number,
  excludeCell: { x: number; y: number } | null = null,
) => {
  // Each assigned panel draws its own type's max watts (MG9 vs MT differ).
  let watts = 0;
  for (const row of grid) {
    for (const cell of row) {
      if (!isActiveCell(cell)) continue;
      if (excludeCell && cell.x === excludeCell.x && cell.y === excludeCell.y) continue;
      if (cell.assignedPowerPort === portId) watts += PANEL_TYPES[cellPanelType(cell)].power.maxW;
    }
  }
  return watts;
};

const getPortPanelCount = (grid: Cell[][], portField: "assignedPort" | "assignedPowerPort", portId: number) =>
  grid.flat().filter((cell) => isActiveCell(cell) && cell[portField] === portId).length;

const getSnakeOrder = (cols: number, rows: number, snakeDirection: string, snakeAlternates = true) => {
  const ordered: Array<{ x: number; y: number }> = [];

  if (snakeDirection === "LR" || snakeDirection === "RL" || snakeDirection === "LRB" || snakeDirection === "RLB") {
    const startFromBottom = snakeDirection === "LRB" || snakeDirection === "RLB";
    const horizontalDirection = snakeDirection === "RL" || snakeDirection === "RLB" ? "RL" : "LR";
    const yIndexes = [...Array(rows).keys()];
    if (startFromBottom) yIndexes.reverse();

    yIndexes.forEach((y, rowIndex) => {
      let row = [...Array(cols).keys()];
      if (horizontalDirection === "RL") row.reverse();
      if (snakeAlternates && rowIndex % 2 === 1) row.reverse();
      row.forEach((x) => ordered.push({ x, y }));
    });
  } else {
    for (let x = 0; x < cols; x += 1) {
      let col = [...Array(rows).keys()];
      if (snakeDirection === "BT") col.reverse();
      if (snakeAlternates && x % 2 === 1) col.reverse();
      col.forEach((y) => ordered.push({ x, y }));
    }
  }

  return ordered;
};

const getLoopTogetherSegments = (cols: number, rows: number) => {
  const segments: Array<Array<{ x: number; y: number }>> = [];
  const leftCount = Math.floor(cols / 2);
  const rightStart = leftCount;

  for (let pairStart = 0; pairStart < rows; pairStart += 2) {
    const topY = pairStart;
    const bottomY = pairStart + 1 < rows ? pairStart + 1 : null;

    const leftSegment: Array<{ x: number; y: number }> = [];
    for (let x = leftCount - 1; x >= 0; x -= 1) leftSegment.push({ x, y: topY });
    if (bottomY !== null) {
      for (let x = 0; x < leftCount; x += 1) leftSegment.push({ x, y: bottomY });
    }
    if (leftSegment.length) segments.push(leftSegment);

    const rightSegment: Array<{ x: number; y: number }> = [];
    for (let x = rightStart; x < cols; x += 1) rightSegment.push({ x, y: topY });
    if (bottomY !== null) {
      for (let x = cols - 1; x >= rightStart; x -= 1) rightSegment.push({ x, y: bottomY });
    }
    if (rightSegment.length) segments.push(rightSegment);
  }

  return segments;
};

const getVerticalStartOrder = (cols: number, rows: number) => {
  const ordered: Array<{ x: number; y: number }> = [];
  for (let x = 0; x < cols; x += 1) {
    for (let y = 0; y < rows; y += 1) {
      ordered.push({ x, y });
    }
  }
  return ordered;
};

const flipX = (x: number, cols: number) => cols - 1 - x;

const getDisplayCell = (cell: Cell, cols: number, isFlippedView: boolean): Cell => ({
  ...cell,
  x: isFlippedView ? flipX(cell.x, cols) : cell.x,
});

// For cabling geometry we want the panel's LEFT-edge display column. An MT panel
// occupies modules x and x+1; when the view is flipped, its left display edge is
// flipX(x) - (span - 1). getLineEndpoints then adds the panel width from its span.
const displayPanelForCabling = (cell: Cell, cols: number, isFlippedView: boolean): Cell => {
  const span = cellSpanX(cell);
  return { ...cell, x: isFlippedView ? flipX(cell.x, cols) - (span - 1) : cell.x };
};

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

const getSelectedKeys = (selectedCells: Set<string>, selectedCell: { x: number; y: number } | null) => {
  if (selectedCells.size > 0) return selectedCells;
  return selectedCell ? new Set([`${selectedCell.x}-${selectedCell.y}`]) : new Set<string>();
};

const getPanelSymbol = (cell: Cell) => {
  const variant = PANEL_VARIANTS[cell.panelVariant ?? "STANDARD"];
  const parts = [];
  if (variant.symbol) parts.push(variant.symbol);
  if (cell.rotation) parts.push("🔄");
  return parts.join(" ");
};

// Cabling endpoints. Cells passed here carry a LEFT-edge display x (see
// displayPanelForCabling), so MT panels (span 2) connect at their true outer
// edge / visual centre instead of a single module's centre.
const getLineEndpoints = (prev: Cell, cell: Cell, offsetY = 0, cellW = CELL_SIZE, cellH = CELL_SIZE) => {
  const stepX = cellW + GRID_GAP;
  const stepY = cellH + GRID_GAP;
  const gapInset = GRID_GAP / 2;
  const widthOf = (c: Cell) => cellSpanX(c) * cellW + (cellSpanX(c) - 1) * GRID_GAP;
  const wPrev = widthOf(prev);
  const wCell = widthOf(cell);
  const leftPrev = prev.x * stepX;
  const leftCell = cell.x * stepX;

  let x1 = leftPrev + wPrev / 2;
  let y1 = prev.y * stepY + cellH / 2 + offsetY;
  let x2 = leftCell + wCell / 2;
  let y2 = cell.y * stepY + cellH / 2 + offsetY;

  if (prev.y === cell.y) {
    if (cell.x > prev.x) {
      x1 = leftPrev + wPrev + gapInset * 0.3;
      x2 = leftCell - gapInset * 0.3;
    } else {
      x1 = leftPrev - gapInset * 0.3;
      x2 = leftCell + wCell + gapInset * 0.3;
    }
  } else if (prev.x === cell.x) {
    if (cell.y > prev.y) {
      y1 = prev.y * stepY + cellH + gapInset * 0.3 + offsetY;
      y2 = cell.y * stepY - gapInset * 0.3 + offsetY;
    } else {
      y1 = prev.y * stepY - gapInset * 0.3 + offsetY;
      y2 = cell.y * stepY + cellH + gapInset * 0.3 + offsetY;
    }
  }

  return { x1, y1, x2, y2 };
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
            <div><b>Patching Mode</b>: click or drag panels to patch the selected signal port or power plug.</div>
            <div><b>Select Mode</b>: drag a box around panels, then change type, rotate, clear, delete, or restore.</div>
            <div>Click away from the signal/power patching cards to clear the active patch target.</div>
          </div>
          <div className="space-y-2">
            <div className="font-semibold text-sky-200">Shortcuts</div>
            <div><b>Ctrl+Z</b>: Undo</div>
            <div><b>Ctrl+Y</b> or <b>Ctrl+Shift+Z</b>: Redo</div>
            <div><b>Delete</b>: Delete selected panels</div>
            <div><b>R</b>: Rotate selected panels</div>
            <div><b>C</b>: Clear selected panel patching</div>
            <div><b>Escape</b>: Clear selection or leave Select Mode</div>
          </div>
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const signalPorts = useMemo(() => makeSignalPorts(), []);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

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
  const [grid, setGrid] = useState<Cell[][]>(() => makeGrid(24, 8));
  const [activePort, setActivePort] = useState(1);
  const [activePowerPort, setActivePowerPort] = useState(1);
  const [patchMode, setPatchMode] = useState<"signal" | "power">("signal");
  const [powerDistro, setPowerDistro] = useState<PowerDistroKey>("32A");
  const [isDragging, setIsDragging] = useState(false);
  const [dragVisited, setDragVisited] = useState<Set<string>>(() => new Set());
  const [selectedCell, setSelectedCell] = useState<{ x: number; y: number } | null>(null);
  const [selectedCells, setSelectedCells] = useState<Set<string>>(() => new Set());
  const [panelSelectMode, setPanelSelectMode] = useState(false);
  const [isSelectingPanels, setIsSelectingPanels] = useState(false);
  const [selectionStart, setSelectionStart] = useState<{ x: number; y: number } | null>(null);
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
  // The grid is a 0.5m square module grid. A single module is CELL_SIZE x CELL_SIZE;
  // MT panels are drawn spanning two modules (see mtTail). MG9 fills one module.
  const cellH = CELL_SIZE;
  const cellW = CELL_SIZE;
  const powerSpec = panel.power;
  const distro = POWER_DISTROS[powerDistro];
  const powerPorts = useMemo(() => makePowerPorts(distro.portCount), [distro.portCount]);

  const [panelsPerPowerOutlet, setPanelsPerPowerOutlet] = useState<number>(panel.defaults.powerPanelsPerOutlet);
  const [panelsPerSignalPort, setPanelsPerSignalPort] = useState<number>(panel.defaults.signalPanelsPerPort);

  const selectedPanel = selectedCell ? grid[selectedCell.y]?.[selectedCell.x] ?? null : null;
  const activeSelectedKeys = getSelectedKeys(selectedCells, selectedCell);
  const selectedCount = activeSelectedKeys.size;
  const isPatchTargetActive = patchMode === "signal" ? activePort > 0 : activePowerPort > 0;

  const captureLayout = (): LayoutSnapshot => ({ grid: cloneGrid(grid), cols, rows });
  const restoreLayout = (snapshot: LayoutSnapshot) => {
    setGrid(cloneGrid(snapshot.grid));
    setCols(snapshot.cols);
    setRows(snapshot.rows);
    setDraftCols(String(snapshot.cols));
    setDraftRows(String(snapshot.rows));
    setSelectedCell(null);
    setSelectedCells(new Set());
    setDragVisited(new Set());
    setIsDragging(false);
    setIsSelectingPanels(false);
  };
  const pushUndoSnapshot = (snapshot = captureLayout()) => {
    setUndoStack((prev) => [...prev.slice(-49), snapshot]);
    setRedoStack([]);
  };
  const commitGridUpdate = (updater: (prev: Cell[][]) => Cell[][]) => {
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
      prev.map((row) =>
        row.map((cell) => {
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
      ),
    );
  }, [powerPorts.length]);

  useEffect(() => {
    const stop = () => {
      setIsDragging(false);
      setIsSelectingPanels(false);
      setSelectionStart(null);
      setDragVisited(new Set());
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
        setSelectedCell(null);
        setSelectedCells(new Set());
        setPanelSelectMode(false);
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

  // The grid is a 0.5m module grid, so physical size is module-count based.
  const wallWidthM = cols * 0.5;
  const wallHeightM = rows * 0.5;
  const activeCells = useMemo(() => grid.flat().filter((cell) => !cell.isRemoved), [grid]);
  // Panels to count/patch: active modules that are not MT tails (MG9 + MT heads).
  const activePanels = useMemo(() => activeCells.filter((cell) => !cell.mtTail), [activeCells]);
  const totalPanels = activePanels.length;
  const panelTypeCounts = useMemo(() => {
    const counts = { MG9: 0, MT: 0 } as Record<PanelTypeKey, number>;
    activePanels.forEach((cell) => {
      counts[cellPanelType(cell)] += 1;
    });
    return counts;
  }, [activePanels]);
  // Pixel resolution uses each panel's native pixels. Because MG9 (168x168) and
  // MT (256x64) have different pitches, a mixed wall isn't a single clean raster:
  // width is the widest row's pixels, height sums each row's tallest panel.
  const wallPixels = useMemo(() => {
    let pixelW = 0;
    let pixelH = 0;
    grid.forEach((row) => {
      let rowPixelW = 0;
      let rowPixelH = 0;
      let rowActive = false;
      row.forEach((cell) => {
        if (!isPanelHead(cell)) return;
        rowActive = true;
        const p = PANEL_TYPES[cellPanelType(cell)];
        rowPixelW += p.pixW;
        rowPixelH = Math.max(rowPixelH, p.pixH);
      });
      pixelW = Math.max(pixelW, rowPixelW);
      if (rowActive) pixelH += rowPixelH;
    });
    return { pixelW, pixelH };
  }, [grid]);
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
  const activeColumns = useMemo(
    () => Array.from({ length: cols }, (_, x) => x).filter((x) => activeCells.some((cell) => cell.x === x)),
    [activeCells, cols],
  );
  const activeRows = useMemo(
    () => Array.from({ length: rows }, (_, y) => y).filter((y) => activeCells.some((cell) => cell.y === y)),
    [activeCells, rows],
  );
  const activeColsCount = activeColumns.length;
  const activeRowsCount = activeRows.length;
  // Module grid: each active module column/row is 0.5m.
  const activeWallWidthM = activeColsCount * 0.5;
  const activeWallHeightM = activeRowsCount * 0.5;
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

    for (const row of grid) {
      for (const cell of row) {
        if (!isActiveCell(cell)) continue;
        if (!cell.assignedPort || !stats[cell.assignedPort]) continue;
        stats[cell.assignedPort].panels += 1;
        stats[cell.assignedPort].path.push(cell);
      }
    }

    signalPorts.forEach((port) => {
      const stat = stats[port.id];
      stat.path.sort((a, b) => (a.sequence ?? 0) - (b.sequence ?? 0));
      const first = stat.path[0];
      const last = stat.path[stat.path.length - 1];
      stat.firstKey = first ? `${first.x}-${first.y}` : null;
      stat.lastKey = last ? `${last.x}-${last.y}` : null;
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

    for (const row of grid) {
      for (const cell of row) {
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
    }

    Object.values(stats).forEach((stat) => {
      stat.utilisation = MAX_OUTLET_AMPS > 0 ? (stat.maxAmps / MAX_OUTLET_AMPS) * 100 : 0;
      stat.path.sort((a, b) => (a.powerSequence ?? 0) - (b.powerSequence ?? 0));
      const first = stat.path[0];
      const last = stat.path[stat.path.length - 1];
      stat.firstKey = first ? `${first.x}-${first.y}` : null;
      stat.lastKey = last ? `${last.x}-${last.y}` : null;
    });

    return stats;
  }, [grid, powerPorts, powerSpec.maxW, powerSpec.maxA, powerSpec.avgW, powerSpec.avgA]);

  // Chain-start indicators for a panel, shared by the live layout and every export.
  // Blue ring: first panel of its signal chain (and the last panel too when the
  // backup signal loop is enabled). Orange ring: first panel of its power chain.
  const getPanelIndicators = (cell: Cell) => {
    const key = `${cell.x}-${cell.y}`;
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
    let mg9 = 0;
    let mt = 0;
    const topRow = grid.find((row) => row.some((cell) => isPanelHead(cell)));
    topRow?.forEach((cell) => {
      if (!isPanelHead(cell)) return;
      if (cellPanelType(cell) === "MT") mt += 1;
      else mg9 += 1;
    });
    return { mg9, mt };
  }, [grid]);
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
    let cornerToFlat = 0;
    let cornerToCorner = 0;
    const directions = [
      [1, 0],
      [-1, 0],
      [0, 1],
      [0, -1],
    ];

    activeCells.forEach((cell) => {
      if (cell.panelVariant !== "CORNER") return;
      directions.forEach(([dx, dy]) => {
        const neighbor = grid[cell.y + dy]?.[cell.x + dx];
        if (!isActiveCell(neighbor)) return;
        if (neighbor.panelVariant === "CORNER") {
          if (neighbor.y > cell.y || (neighbor.y === cell.y && neighbor.x > cell.x)) cornerToCorner += 1;
        }
        else cornerToFlat += 1;
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

      (Object.keys(PANEL_VARIANTS) as PanelVariantKey[]).forEach((variantKey) => {
        if (variantKey === "STANDARD") return;
        const variant = PANEL_VARIANTS[variantKey];
        const item = variant.stockItem;
        const count = panelVariantCounts[variantKey];
        if (!item || count <= 0) return;
        const spare = Math.ceil(count * mg9Defaults.spareRatio);
        const rounded = roundUpToBox(count + spare, mg9Defaults.panelsPerBox);
        rowsOut.push(makeStockRow(item, rounded, `${count} selected + ${spare} spare, rounded to box of ${mg9Defaults.panelsPerBox}`, spare, rounded));
      });
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
  }, [activeColsCount, activeRowsCount, activeWallWidthM, backupSignalLoop, circuitsUsedMax, cornerJoinStats, deploymentType, distroRequired, includeReinforcementPlate, panelVariantCounts, powerCableTotalRequired, powerDistro, signalCableBaseRequired, signalCableSpare, signalCableTotalRequired, signalCableWithBackupRequired, signalPortsUsed, powerPortsUsed, distro.portCount, mg9Count, mtCount, mg9Spare, mtSpare, mg9Boxes, mtBoxes, mg9Defaults, mtDefaults, topRowBars]);

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
    const canvas = document.createElement("canvas");
    canvas.width = (svgW + 96) * scale;
    canvas.height = (svgH + 96) * scale;
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Canvas context unavailable");
    ctx.scale(scale, scale);

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, svgW + 96, svgH + 96);
    ctx.fillStyle = "#0f172a";
    ctx.font = "bold 18px Arial";
    ctx.textAlign = "left";
    ctx.fillText(viewLabel, 20, 24);

    const offsetX = 56;
    const offsetY = 40;
    ctx.save();
    ctx.translate(offsetX, offsetY);

    ctx.fillStyle = "#0f172a";
    ctx.font = "16px Arial";
    ctx.textAlign = "center";
    for (let i = 0; i < cols; i += 1) {
      const displayIndex = flipped ? cols - i : i + 1;
      ctx.fillText(String(displayIndex), i * (cellW + GRID_GAP) + cellW / 2, -10);
    }

    ctx.textAlign = "left";
    for (let i = 0; i < rows; i += 1) {
      ctx.fillText(String(i + 1), -28, i * (cellH + GRID_GAP) + cellH / 2 + 6);
    }

    grid.flat().forEach((cell) => {
      if (!isPanelHead(cell)) return;
      const disp = displayPanelForCabling(cell, cols, flipped);
      const span = cellSpanX(cell);
      const w = span * cellW + (span - 1) * GRID_GAP;
      const x = disp.x * (cellW + GRID_GAP);
      const y = cell.y * (cellH + GRID_GAP);
      const fill = cell.assignedPort ? PORT_COLORS[(cell.assignedPort - 1) % PORT_COLORS.length] : "#1e293b";
      const { signalRing, powerRing } = getPanelIndicators(cell);
      drawPanelShape(ctx, x, y, w, cellH, cell, fill, "#0f172a", 2, { signalRing, powerRing });
    });

    Object.entries(signalPortStats).forEach(([portId, stat]) => {
      if (!stat.path || stat.path.length < 2) return;
      const color = PORT_COLORS[(Number(portId) - 1) % PORT_COLORS.length];
      ctx.strokeStyle = color;
      ctx.lineWidth = 4;
      ctx.lineCap = "round";
      stat.path.forEach((cell, idx) => {
        if (idx === 0) return;
        const prev = displayPanelForCabling(stat.path[idx - 1], cols, flipped);
        const current = displayPanelForCabling(cell, cols, flipped);
        let { x1, y1, x2, y2 } = getLineEndpoints(prev, current, 0, cellW, cellH);
        if (current.y !== prev.y) {
          const sideOffset = GRID_GAP * 0.5;
          x1 += flipped ? sideOffset : -sideOffset;
          x2 += flipped ? sideOffset : -sideOffset;
        }
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
        const prev = displayPanelForCabling(path[idx - 1], cols, flipped);
        const current = displayPanelForCabling(cell, cols, flipped);
        let { x1, y1, x2, y2 } = getLineEndpoints(prev, current, 4, cellW, cellH);
        if (current.y !== prev.y) {
          const sideOffset = GRID_GAP * 0.5;
          x1 += flipped ? -sideOffset : sideOffset;
          x2 += flipped ? -sideOffset : sideOffset;
        }
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.stroke();
        drawCanvasArrowHead(ctx, x1, y1, x2, y2, POWER_COLOR);
      });
    });

    grid.flat().forEach((cell) => {
      if (!isPanelHead(cell)) return;
      const disp = displayPanelForCabling(cell, cols, flipped);
      const span = cellSpanX(cell);
      const w = span * cellW + (span - 1) * GRID_GAP;
      const x = disp.x * (cellW + GRID_GAP);
      const y = cell.y * (cellH + GRID_GAP);
      const headDisplayX = getDisplayCell(cell, cols, flipped).x;
      const cx = x + w / 2;
      ctx.fillStyle = "#020617";
      ctx.font = "bold 10px Arial";
      ctx.textAlign = "center";
      ctx.fillText(`↓ ${cell.y + 1} → ${headDisplayX + 1}${cellPanelType(cell) === "MT" ? " (MT)" : ""}`, cx, y + 18);
      if (cell.assignedPort) ctx.fillText(`🔌 P${cell.assignedPort} (${cell.sequence ?? "-"})`, cx, y + 34);
      if (cell.assignedPowerPort) ctx.fillText(`⚡ Plug ${cell.assignedPowerPort}`, cx, y + 50);
      const variantSymbol = getPanelSymbol(cell);
      if (variantSymbol) ctx.fillText(variantSymbol, cx, y + cellH - 6);
    });

    ctx.restore();
    return canvas;
  };

const exportJson = () => {
  try {
    const payload = {
      projectName: safeProjectName,
      panelType,
      powerDistro,
      backupSignalLoop,
      includeReinforcementPlate,
      deploymentType,
      wall: { cols, rows, widthM: wallWidthM, heightM: wallHeightM, pixelW: wallPixelW, pixelH: wallPixelH },
      patching: { grid, signalPortsUsed, powerPortsUsed },
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
      // 168x168, MT 256x64) laid out left-to-right per row; rows are top-aligned
      // and as tall as the tallest panel in that row (mixed pitches aren't a
      // single clean raster). Panels within a row are ordered by front-view x.
      const rowsHeads = grid.map((row) =>
        row
          .filter((cell) => isPanelHead(cell))
          .map((cell) => ({ cell, leftX: flipX(cell.x, cols) - (cellSpanX(cell) - 1) }))
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
          const headDisplayX = flipX(cell.x, cols);
          const fill = cell.assignedPort ? PORT_COLORS[(cell.assignedPort - 1) % PORT_COLORS.length] : "#1e293b";
          const { signalRing, powerRing } = getPanelIndicators(cell);
          drawPanelShape(ctx, x, yy, p.pixW, p.pixH, cell, fill, "#ffffff", 1, { hatchStep: 24, curveStyle: "test-pattern", signalRing, powerRing });

          ctx.fillStyle = "#020617";
          ctx.textAlign = "center";
          ctx.font = `bold ${Math.max(12, Math.floor(p.pixH * 0.085))}px Arial`;
          ctx.fillText(`↓ ${cell.y + 1} → ${headDisplayX + 1}`, x + p.pixW / 2, yy + p.pixH * 0.28);
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
        let nextCols = Math.max(1, Number(data.wall?.cols || cols));
        const nextRows = Math.max(1, Number(data.wall?.rows || rows));
        const rawGrid = data.patching?.grid;
        let nextGrid: Cell[][];
        if (data.panelType === "MT" && isLegacyGrid(rawGrid)) {
          // Legacy all-MT project: expand onto the 0.5m module grid (columns x2).
          const migrated = expandLegacyMtGrid(rawGrid, nextCols, nextRows);
          nextGrid = migrated.grid;
          nextCols = migrated.cols;
        } else {
          nextGrid = Array.isArray(rawGrid) ? normalizeGrid(rawGrid, nextCols, nextRows) : makeGrid(nextCols, nextRows);
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
        setGrid(nextGrid);
        setSelectedCell(null);
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
      pdf.text(`Panels: ${cols} x ${rows} module grid, ${totalPanels} active`, 10, 38);

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
      `Panels: ${cols} x ${rows} module grid, ${totalPanels} active`,
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
    setGrid(makeGrid(nextCols, nextRows));
    setSelectedCell(null);
    setSelectedCells(new Set());
    setDragVisited(new Set());
    setIsDragging(false);
  };

  const assignSignalCell = (x: number, y: number) => {
    if (activePort < 1) return;
    const key = `${x}-${y}`;
    if (dragVisited.has(key)) return;

    commitGridUpdate((prev) => {
      const current = prev[y]?.[x];
      if (!current) return prev;
      if (!isActiveCell(current) || current.mtTail) return prev;

      const currentCount = getPortPanelCount(prev, "assignedPort", activePort);
      const isAlreadySamePort = current.assignedPort === activePort;
      if (!isAlreadySamePort && currentCount >= safePanelsPerSignalPort) return prev;

      const next = cloneGrid(prev);
      next[y][x].assignedPort = activePort;
      if (!isAlreadySamePort) {
        next[y][x].sequence = getNextSequence(next, "assignedPort", "sequence", activePort);
      }
      return next;
    });

    setDragVisited((prev) => new Set(prev).add(key));
  };

  const assignPowerCell = (x: number, y: number) => {
    if (activePowerPort < 1) return;
    const key = `${x}-${y}`;
    if (dragVisited.has(key)) return;

    commitGridUpdate((prev) => {
      const current = prev[y]?.[x];
      if (!current) return prev;
      if (!isActiveCell(current) || current.mtTail) return prev;

      const currentPanels = getPortPanelCount(prev, "assignedPowerPort", activePowerPort);
      const isAlreadySamePort = current.assignedPowerPort === activePowerPort;
      if (!isAlreadySamePort && currentPanels >= safePanelsPerPowerOutlet) return prev;

      const cellWatts = PANEL_TYPES[cellPanelType(current)].power.maxW;
      const currentPortLoad = getPowerPortLoadWatts(prev, activePowerPort, 0, { x, y });
      if (!isAlreadySamePort && currentPortLoad + cellWatts > MAX_OUTLET_AMPS * VOLTAGE) return prev;

      const next = cloneGrid(prev);
      next[y][x].assignedPowerPort = activePowerPort;
      next[y][x].powerManual = true;
      if (!isAlreadySamePort) {
        next[y][x].powerSequence = getNextSequence(next, "assignedPowerPort", "powerSequence", activePowerPort);
      }
      return next;
    });

    setDragVisited((prev) => new Set(prev).add(key));
  };

  const startDrag = (x: number, y: number) => {
    const target = grid[y]?.[x];
    if (!target) return;
    if (panelSelectMode) {
      const key = `${x}-${y}`;
      setSelectionStart({ x, y });
      setIsSelectingPanels(true);
      setSelectedCell({ x, y });
      setSelectedCells(new Set([key]));
      return;
    }
    if (!isActiveCell(target)) return;
    setDragVisited(new Set());
    setIsDragging(true);
    if (patchMode === "signal") assignSignalCell(x, y);
    else assignPowerCell(x, y);
  };

  const continueDrag = (x: number, y: number) => {
    if (panelSelectMode && isSelectingPanels && selectionStart) {
      const minX = Math.min(selectionStart.x, x);
      const maxX = Math.max(selectionStart.x, x);
      const minY = Math.min(selectionStart.y, y);
      const maxY = Math.max(selectionStart.y, y);
      const next = new Set<string>();
      for (let yy = minY; yy <= maxY; yy += 1) {
        for (let xx = minX; xx <= maxX; xx += 1) {
          if (grid[yy]?.[xx]) next.add(`${xx}-${yy}`);
        }
      }
      setSelectedCells(next);
      setSelectedCell({ x, y });
      return;
    }
    if (!isDragging) return;
    const target = grid[y]?.[x];
    if (!isActiveCell(target)) return;
    if (patchMode === "signal") assignSignalCell(x, y);
    else assignPowerCell(x, y);
  };

  const applyManualSignalPatch = (value: string) => {
    if (!selectedCell) return;
    const nextPort = value === "" ? null : Number.parseInt(value, 10);
    if (nextPort !== null && (!Number.isFinite(nextPort) || nextPort < 1 || nextPort > SIGNAL_PORT_COUNT)) return;

    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      const target = next[selectedCell.y]?.[selectedCell.x];
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
    if (!selectedCell) return;
    const nextPort = value === "" ? null : Number.parseInt(value, 10);
    if (nextPort !== null && (!Number.isFinite(nextPort) || nextPort < 1 || nextPort > powerPorts.length)) return;

    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      const target = next[selectedCell.y]?.[selectedCell.x];
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

      const currentPortLoad = getPowerPortLoadWatts(prev, nextPort, powerSpec.maxW, selectedCell);
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
    const ordered = snakeDirection === "LOOP_TOGETHER" ? [] : getSnakeOrder(cols, rows, snakeDirection, snakeAlternates);
    const loopTogetherSegments = snakeDirection === "LOOP_TOGETHER" ? getLoopTogetherSegments(cols, rows) : [];
    const useVerticalLoopTogether = snakeDirection === "LOOP_TOGETHER" && Math.max(cols, rows) > 23;
    const loopTogetherVertical = useVerticalLoopTogether ? getVerticalStartOrder(cols, rows) : [];

    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);

      if (patchMode === "signal") {
        for (const row of next) {
          for (const cell of row) {
            cell.assignedPort = null;
            cell.sequence = null;
          }
        }

        if (snakeDirection === "LOOP_TOGETHER") {
          let port = 1;
          let seq = 1;
          const applyCell = ({ x, y }: { x: number; y: number }) => {
            if (port > SIGNAL_PORT_COUNT) return;
            if (!isPanelHead(next[y]?.[x])) return;
            next[y][x].assignedPort = port;
            next[y][x].sequence = seq;
            seq += 1;
            if (seq > safePanelsPerSignalPort) {
              port += 1;
              seq = 1;
            }
          };

          if (useVerticalLoopTogether) {
            loopTogetherVertical.forEach(applyCell);
          } else {
            loopTogetherSegments.forEach((segment) => {
              segment.forEach(applyCell);
              if (seq !== 1) {
                port += 1;
                seq = 1;
              }
            });
          }
        } else {
          let port = 1;
          let seq = 1;
          ordered.forEach(({ x, y }) => {
            if (port > SIGNAL_PORT_COUNT) return;
            if (!isPanelHead(next[y]?.[x])) return;
            next[y][x].assignedPort = port;
            next[y][x].sequence = seq;
            seq += 1;
            if (seq > safePanelsPerSignalPort) {
              port += 1;
              seq = 1;
            }
          });
        }
      }

      if (patchMode === "power") {
        for (const row of next) {
          for (const cell of row) {
            cell.assignedPowerPort = null;
            cell.powerSequence = null;
            cell.powerManual = false;
          }
        }

        let portIndex = 0;
        ordered.forEach(({ x, y }) => {
          if (!isPanelHead(next[y]?.[x])) return;
          const cellWatts = PANEL_TYPES[cellPanelType(next[y][x])].power.maxW;
          while (portIndex < powerPorts.length) {
            const port = powerPorts[portIndex];
            const currentLoad = getPowerPortLoadWatts(next, port.id, 0);
            const currentPanels = getPortPanelCount(next, "assignedPowerPort", port.id);

            if (currentPanels >= safePanelsPerPowerOutlet) {
              portIndex += 1;
              continue;
            }

            if (currentLoad + cellWatts <= MAX_OUTLET_AMPS * VOLTAGE) {
              next[y][x].assignedPowerPort = port.id;
              next[y][x].powerSequence = getNextSequence(next, "assignedPowerPort", "powerSequence", port.id);
              next[y][x].powerManual = false;
              return;
            }

            portIndex += 1;
          }
        });
      }

      return next;
    });

    setSelectedCell(null);
    setSelectedCells(new Set());
    setDragVisited(new Set());
    setIsDragging(false);
  };

  // Patch power to follow the existing signal patch: walk panels in signal order
  // (signal port, then sequence) and fill power plugs, starting a fresh plug for
  // each signal port so power plugs line up with the signal ports. Respects the
  // power panel-count and amp limits, and stops when the plugs run out.
  const matchPowerToSignal = () => {
    const hasSignal = grid.flat().some((cell) => isActiveCell(cell) && cell.assignedPort);
    if (!hasSignal) {
      alert("Patch the signal ports first - power will follow the same pattern.");
      return;
    }

    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);

      for (const row of next) {
        for (const cell of row) {
          cell.assignedPowerPort = null;
          cell.powerSequence = null;
          cell.powerManual = false;
        }
      }

      const byPort = new Map<number, { x: number; y: number; seq: number }[]>();
      next.forEach((row, y) =>
        row.forEach((cell, x) => {
          if (!isActiveCell(cell) || !cell.assignedPort) return;
          const list = byPort.get(cell.assignedPort) ?? [];
          list.push({ x, y, seq: cell.sequence ?? 0 });
          byPort.set(cell.assignedPort, list);
        }),
      );
      const orderedSignalPorts = [...byPort.keys()].sort((a, b) => a - b);

      let plugIndex = 0;
      const plugLeft = () => plugIndex < powerPorts.length;

      for (const sigPort of orderedSignalPorts) {
        if (!plugLeft()) break;
        // Align power plugs to signal ports: each new signal port starts on a fresh plug.
        if (getPortPanelCount(next, "assignedPowerPort", powerPorts[plugIndex].id) > 0) {
          plugIndex += 1;
        }

        const cells = byPort.get(sigPort)!.sort((a, b) => a.seq - b.seq);
        for (const { x, y } of cells) {
          const cellWatts = PANEL_TYPES[cellPanelType(next[y][x])].power.maxW;
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
            next[y][x].assignedPowerPort = plug.id;
            next[y][x].powerSequence = getNextSequence(next, "assignedPowerPort", "powerSequence", plug.id);
            next[y][x].powerManual = false;
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
    setSelectedCell(null);
    setSelectedCells(new Set());
    setDragVisited(new Set());
    setIsDragging(false);
  };

  const clearSignalCabling = () => {
    commitGridUpdate((prev) => clearSignalOnGrid(prev));
    setSelectedCell(null);
    setSelectedCells(new Set());
    setDragVisited(new Set());
    setIsDragging(false);
  };

  const clearPowerAssignments = () => {
    commitGridUpdate((prev) => clearPowerOnGrid(prev));
    setSelectedCell(null);
    setSelectedCells(new Set());
  };

  const clearSelectedPanelPatching = () => {
    const keys = getSelectedKeys(selectedCells, selectedCell);
    if (!keys.size) return;
    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      keys.forEach((key) => {
        const [x, y] = key.split("-").map(Number);
        const target = next[y]?.[x];
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
    const keys = getSelectedKeys(selectedCells, selectedCell);
    if (!keys.size) return;
    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      keys.forEach((key) => {
        const [x, y] = key.split("-").map(Number);
        const target = next[y]?.[x];
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
    const keys = getSelectedKeys(selectedCells, selectedCell);
    if (!keys.size) return;
    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      keys.forEach((key) => {
        const [x, y] = key.split("-").map(Number);
        const target = next[y]?.[x];
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
    const keys = getSelectedKeys(selectedCells, selectedCell);
    if (!keys.size) return;
    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      keys.forEach((key) => {
        const [x, y] = key.split("-").map(Number);
        const target = next[y]?.[x];
        if (!target || !isActiveCell(target)) return;
        // Variants (triangle/curve/corner) are MG9-only.
        if (cellPanelType(target) !== "MG9") return;
        target.panelVariant = variant;
      });
      return next;
    });
  };

  const applySelectedPanelType = (type: PanelTypeKey) => {
    const keys = getSelectedKeys(selectedCells, selectedCell);
    if (!keys.size) return;
    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      keys.forEach((key) => {
        const [x, y] = key.split("-").map(Number);
        const target = next[y]?.[x];
        // Only convert real panels (active, non-tail heads).
        if (!isPanelHead(target)) return;
        if (type === "MT") {
          if (setModuleToMT(next, x, y)) target.panelVariant = "STANDARD";
        } else {
          setModuleToMG9(next, x, y);
        }
      });
      return next;
    });
  };

  const rotateSelectedPanels = () => {
    const keys = getSelectedKeys(selectedCells, selectedCell);
    if (!keys.size) return;
    commitGridUpdate((prev) => {
      const next = cloneGrid(prev);
      keys.forEach((key) => {
        const [x, y] = key.split("-").map(Number);
        const target = next[y]?.[x];
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
      next.forEach((row) =>
        row.forEach((cell) => {
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
        }),
      );
      return next;
    });
  };

  const toDisplayX = (x: number) => (isFlippedView ? flipX(x, cols) : x);
  const fromDisplayX = (displayX: number) => (isFlippedView ? flipX(displayX, cols) : displayX);

  const svgW = cols * cellW + (cols - 1) * GRID_GAP;
  const svgH = rows * cellH + (rows - 1) * GRID_GAP;

  return (
    <div className="min-h-screen bg-[#0f172a] p-6 text-white print-container">
      {showHelp ? <HelpModal onClose={() => setShowHelp(false)} /> : null}
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
              <Button intent="ghost" onClick={() => setShowHelp(true)}>
                <HelpCircle className="h-4 w-4" />Help
              </Button>
            </div>
            <input ref={fileInputRef} type="file" accept="application/json" className="hidden" onChange={openJson} />
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
                  <div>Panels: {cols} × {rows} grid, {totalPanels} active panels</div>
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
                active={panelSelectMode}
                activeAccent="emerald"
                onClick={() => {
                  setPanelSelectMode((prev) => {
                    if (prev) {
                      setSelectedCell(null);
                      setSelectedCells(new Set());
                    }
                    return !prev;
                  });
                }}
              >
                {panelSelectMode ? "Select Mode (on)" : "Enable Select Mode"}
              </Button>
              <StatusChip tone="emerald">{selectedCount ? `${selectedCount} selected` : "None selected"}</StatusChip>
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
            <div className="w-full overflow-auto rounded-xl bg-white/5 p-4 pt-6 pl-8 select-none">
              <div className="relative" style={{ width: svgW, height: svgH }}>
                <div className="absolute left-0 top-[-20px] grid text-xs text-white [text-shadow:0_0_2px_black]" style={{ gridTemplateColumns: `repeat(${cols}, ${cellW}px)`, gap: GRID_GAP }}>
                  {Array.from({ length: cols }).map((_, index) => <div key={`col-${index}`} className="text-center">{isFlippedView ? cols - index : index + 1}</div>)}
                </div>

                <div className="absolute left-[-30px] top-0 grid text-xs text-white [text-shadow:0_0_2px_black]" style={{ gridTemplateRows: `repeat(${rows}, ${cellH}px)`, gap: GRID_GAP }}>
                  {Array.from({ length: rows }).map((_, index) => <div key={`row-${index}`} className="flex items-center">{index + 1}</div>)}
                </div>

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
                      const prev = displayPanelForCabling(stat.path[idx - 1], cols, isFlippedView);
                      const current = displayPanelForCabling(cell, cols, isFlippedView);
                      let { x1, y1, x2, y2 } = getLineEndpoints(prev, current, 0, cellW, cellH);

                      if (current.y !== prev.y) {
                        const sideOffset = GRID_GAP * 0.5;
                        x1 += isFlippedView ? sideOffset : -sideOffset;
                        x2 += isFlippedView ? sideOffset : -sideOffset;
                      }

                      return <line key={`sig-${portId}-${idx}`} x1={x1} y1={y1} x2={x2} y2={y2} stroke={color} style={{ color }} strokeWidth="4" markerEnd="url(#arrow)" />;
                    });
                  })}

                  {powerPorts.map((port) => {
                    const stat = powerPortStats[port.id];
                    const path = stat?.path ?? [];
                    if (path.length < 2) return null;

                    return path.map((cell, idx) => {
                      if (idx === 0) return null;
                      const prev = displayPanelForCabling(path[idx - 1], cols, isFlippedView);
                      const current = displayPanelForCabling(cell, cols, isFlippedView);
                      let { x1, y1, x2, y2 } = getLineEndpoints(prev, current, 4, cellW, cellH);

                      if (current.y !== prev.y) {
                        const sideOffset = GRID_GAP * 0.5;
                        x1 += isFlippedView ? -sideOffset : sideOffset;
                        x2 += isFlippedView ? -sideOffset : sideOffset;
                      }

                      return <line key={`pow-${port.id}-${idx}`} x1={x1} y1={y1} x2={x2} y2={y2} stroke={POWER_COLOR} style={{ color: POWER_COLOR }} strokeWidth="4" markerEnd="url(#arrow)" />;
                    });
                  })}
                </svg>

                <div className="absolute inset-0 z-10 grid" style={{ gridTemplateColumns: `repeat(${cols}, ${cellW}px)`, gap: GRID_GAP }}>
                  {grid.flat().map((cell) => {
                    // MT tail modules are drawn as part of their head panel.
                    if (cell.mtTail) return null;
                    const span = cellSpanX(cell);
                    const displayX = toDisplayX(cell.x);
                    const key = `${displayX}-${cell.y}`;
                    const originalKey = `${cell.x}-${cell.y}`;
                    const signalStat = cell.assignedPort ? signalPortStats[cell.assignedPort] : null;
                    const isEdge = signalStat?.firstKey === originalKey || signalStat?.lastKey === originalKey;
                    const { signalRing, powerRing } = getPanelIndicators(cell);
                    const isSelected = selectedCells.has(originalKey) || (selectedCell?.x === cell.x && selectedCell?.y === cell.y);
                    const isRemoved = cell.isRemoved;
                    const displayColor = isRemoved ? "transparent" : cell.assignedPort ? PORT_COLORS[(cell.assignedPort - 1) % PORT_COLORS.length] : "#1e293b";
                    const variant = PANEL_VARIANTS[cell.panelVariant ?? "STANDARD"];
                    const shapeClipPath =
                      variant.shape === "triangle"
                        ? "polygon(50% 0, 100% 100%, 0 100%)"
                        : variant.shape === "curve"
                          ? "polygon(0 0, 100% 0, 100% 70%, 72% 100%, 0 100%)"
                          : undefined;
                    const hatch =
                      variant.shape === "corner"
                        ? `repeating-linear-gradient(135deg, transparent 0 6px, rgba(15,23,42,0.35) 6px 8px), ${displayColor}`
                        : displayColor;

                    return (
                      <div
                        key={key}
                        onMouseDown={() => startDrag(fromDisplayX(displayX), cell.y)}
                        onMouseEnter={() => continueDrag(fromDisplayX(displayX), cell.y)}
                        onClick={() => {
                          if (isSelectingPanels) return;
                          if (!panelSelectMode) return;
                          setSelectedCell({ x: cell.x, y: cell.y });
                          setSelectedCells(new Set([originalKey]));
                        }}
                        style={{
                          width: span * cellW + (span - 1) * GRID_GAP,
                          height: cellH,
                          background: "transparent",
                          border: `2px ${isRemoved ? "dashed" : "solid"} ${isSelected ? "#ffffff" : isRemoved ? "#64748b" : "transparent"}`,
                          boxShadow: "none",
                          color: isRemoved ? "#94a3b8" : "#020617",
                          gridColumnStart: (isFlippedView ? displayX - (span - 1) : displayX) + 1,
                          gridColumnEnd: `span ${span}`,
                          gridRowStart: cell.y + 1,
                        }}
                        className="relative flex cursor-pointer select-none flex-col items-center justify-center gap-[2px] p-1 text-[9px] font-semibold leading-tight tracking-tight"
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
                            <div className="relative z-10">{`↓ ${cell.y + 1} → ${displayX + 1}`}</div>
                            {cell.assignedPort ? <div className="relative z-10 whitespace-nowrap">{`🔌 P${cell.assignedPort} (${cell.sequence ?? "-"})`}</div> : null}
                            {cell.assignedPowerPort ? <div className="relative z-10 whitespace-nowrap">{`⚡ Plug ${cell.assignedPowerPort}`}</div> : null}
                            {getPanelSymbol(cell) ? <div className="relative z-10 text-[11px]">{getPanelSymbol(cell)}</div> : null}
                          </>
                        )}
                      </div>
                    );
                  })}
                </div>
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
console.assert(makeGrid(2, 3).length === 3 && makeGrid(2, 3)[0].length === 2, "makeGrid should build correct dimensions");

