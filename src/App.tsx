import { Wand2, Zap, Download, Upload, FileText } from "lucide-react";
import React, { useEffect, useMemo, useRef, useState } from "react";

// Simple local UI components (replacing shadcn)
const Button = ({ children, className = "", variant = "solid", type = "button", ...props }: any) => (
  <button
    className={`inline-flex items-center justify-center rounded-lg border px-3 py-2 font-medium transition-colors ${
      variant === "outline"
        ? "border-slate-500 bg-slate-700 text-white hover:bg-slate-600"
        : "border-sky-400 bg-sky-600 text-white hover:bg-sky-500"
    } ${className}`}
    type={type}
    {...props}
  >
    {children}
  </button>
);

const Card = ({ children, className = "" }: any) => (
  <div className={`rounded border p-3 ${className}`}>{children}</div>
);

const CardHeader = ({ children, className = "" }: any) => <div className={`mb-2 font-bold ${className}`}>{children}</div>;
const CardContent = ({ children, className = "" }: any) => <div className={className}>{children}</div>;
const CardTitle = ({ children, className = "" }: any) => <div className={className}>{children}</div>;

const Input = ({ className = "", ...props }: any) => (
  <input className={`w-full rounded border p-2 text-black ${className}`} {...props} />
);

const SIGNAL_PORT_COUNT = 20;
const CELL_SIZE = 64;
const GRID_GAP = 8;
const MAX_PIXELS_PER_PORT = 650000;
const VOLTAGE = 230;
const MAX_OUTLET_AMPS = 16;
const POWER_COLOR = "#f97316";
const APP_VERSION = "0.9.0";

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
type PowerDistroKey = keyof typeof POWER_DISTROS;
type DeploymentType = (typeof DEPLOYMENT_TYPES)[keyof typeof DEPLOYMENT_TYPES];

type StockRow = {
  code: string;
  name: string;
  required: number;
  stock: number;
  net: number;
  method: string;
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
    })),
  );

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
      };
    }),
  );

const isActiveCell = (cell: Cell | null | undefined) => Boolean(cell && !cell.isRemoved);

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
  panelPowerMaxW: number,
  excludeCell: { x: number; y: number } | null = null,
) => {
  let watts = 0;
  for (const row of grid) {
    for (const cell of row) {
      if (!isActiveCell(cell)) continue;
      if (excludeCell && cell.x === excludeCell.x && cell.y === excludeCell.y) continue;
      if (cell.assignedPowerPort === portId) watts += panelPowerMaxW;
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

const makeStockRow = (
  item: { code: string; name: string; stock: number },
  required: number,
  method: string,
): StockRow => ({
  code: item.code,
  name: item.name,
  required,
  stock: item.stock,
  net: item.stock - required,
  method,
});

const getLineEndpoints = (prev: Cell, cell: Cell, offsetY = 0) => {
  const center = CELL_SIZE / 2;
  const panelSpan = CELL_SIZE + GRID_GAP;
  const gapInset = GRID_GAP / 2;

  let x1 = prev.x * panelSpan + center;
  let y1 = prev.y * panelSpan + center + offsetY;
  let x2 = cell.x * panelSpan + center;
  let y2 = cell.y * panelSpan + center + offsetY;

  if (prev.y === cell.y) {
    if (cell.x > prev.x) {
      x1 = prev.x * panelSpan + CELL_SIZE + gapInset * 0.3;
      x2 = cell.x * panelSpan - gapInset * 0.3;
    } else {
      x1 = prev.x * panelSpan - gapInset * 0.3;
      x2 = cell.x * panelSpan + CELL_SIZE + gapInset * 0.3;
    }
  } else if (prev.x === cell.x) {
    if (cell.y > prev.y) {
      y1 = prev.y * panelSpan + CELL_SIZE + gapInset * 0.3 + offsetY;
      y2 = cell.y * panelSpan - gapInset * 0.3 + offsetY;
    } else {
      y1 = prev.y * panelSpan - gapInset * 0.3 + offsetY;
      y2 = cell.y * panelSpan + CELL_SIZE + gapInset * 0.3 + offsetY;
    }
  }

  return { x1, y1, x2, y2 };
};

function UtilBar({ percent }: { percent: number }) {
  const color = getStatusColor(percent);
  return (
    <div className="h-2 w-full rounded border border-white/30 bg-black/30">
      <div className="h-2 rounded" style={{ width: `${Math.min(percent, 100)}%`, background: color }} />
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
  const [patchMode, setPatchMode] = useState("signal");
  const [powerDistro, setPowerDistro] = useState<PowerDistroKey>("32A");
  const [isDragging, setIsDragging] = useState(false);
  const [dragVisited, setDragVisited] = useState<Set<string>>(() => new Set());
  const [selectedCell, setSelectedCell] = useState<{ x: number; y: number } | null>(null);
  const [snakeDirection, setSnakeDirection] = useState("LR");
  const [snakeAlternates, setSnakeAlternates] = useState(true);
  const [isFlippedView, setIsFlippedView] = useState(false);
  const [backupSignalLoop, setBackupSignalLoop] = useState(true);
  const [includeReinforcementPlate, setIncludeReinforcementPlate] = useState(false);
  const [deploymentType, setDeploymentType] = useState<DeploymentType | "">("");

  const panel = PANEL_TYPES[panelType];
  const powerSpec = panel.power;
  const distro = POWER_DISTROS[powerDistro];
  const powerPorts = useMemo(() => makePowerPorts(distro.portCount), [distro.portCount]);

  const [panelsPerPowerOutlet, setPanelsPerPowerOutlet] = useState(panel.defaults.powerPanelsPerOutlet);
  const [panelsPerSignalPort, setPanelsPerSignalPort] = useState(panel.defaults.signalPanelsPerPort);

  const selectedPanel = selectedCell ? grid[selectedCell.y]?.[selectedCell.x] ?? null : null;
  const selectedDisplayCell = selectedPanel ? getDisplayCell(selectedPanel, cols, isFlippedView) : null;

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
      setDragVisited(new Set());
    };
    window.addEventListener("mouseup", stop);
    return () => window.removeEventListener("mouseup", stop);
  }, []);

  const maxAllowedPowerPanels = 21;
  const safePanelsPerPowerOutlet = Math.min(Math.max(panelsPerPowerOutlet, 1), maxAllowedPowerPanels);
  const safePanelsPerSignalPort = Math.min(Math.max(panelsPerSignalPort, 1), panel.defaults.signalPanelsPerPort);

  const powerOutletWatts = safePanelsPerPowerOutlet * powerSpec.maxW;
  const powerOutletAmps = safePanelsPerPowerOutlet * powerSpec.maxA;
  const powerOutletPercent = (powerOutletAmps / MAX_OUTLET_AMPS) * 100;

  const panelPixels = panel.pixW * panel.pixH;
  const signalPortPixels = safePanelsPerSignalPort * panelPixels;
  const signalPortPercent = (signalPortPixels / MAX_PIXELS_PER_PORT) * 100;

  const wallWidthM = cols * panel.w;
  const wallHeightM = rows * panel.h;
  const wallPixelW = cols * panel.pixW;
  const wallPixelH = rows * panel.pixH;
  const activeCells = useMemo(() => grid.flat().filter((cell) => !cell.isRemoved), [grid]);
  const totalPanels = activeCells.length;
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
  const activeWallWidthM = activeColsCount * panel.w;
  const activeWallHeightM = activeRowsCount * panel.h;
  const panelOnlyWeight = totalPanels * panel.weight;
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
        },
      ]),
    );

    for (const row of grid) {
      for (const cell of row) {
        if (!isActiveCell(cell)) continue;
        if (!cell.assignedPowerPort || !stats[cell.assignedPowerPort]) continue;
        const stat = stats[cell.assignedPowerPort];
        stat.panels += 1;
        stat.maxWatts += powerSpec.maxW;
        stat.maxAmps += powerSpec.maxA;
        stat.avgWatts += powerSpec.avgW;
        stat.avgAmps += powerSpec.avgA;
        stat.path.push(cell);
        if (cell.powerManual) stat.manualPanels += 1;
      }
    }

    Object.values(stats).forEach((stat) => {
      stat.utilisation = MAX_OUTLET_AMPS > 0 ? (stat.maxAmps / MAX_OUTLET_AMPS) * 100 : 0;
      stat.path.sort((a, b) => (a.powerSequence ?? 0) - (b.powerSequence ?? 0));
    });

    return stats;
  }, [grid, powerPorts, powerSpec.maxW, powerSpec.maxA, powerSpec.avgW, powerSpec.avgA]);

  const powerPortsUsed = useMemo(() => Object.values(powerPortStats).filter((stat) => stat.panels > 0).length, [powerPortStats]);
  const signalPortsUsed = useMemo(() => Object.values(signalPortStats).filter((stat) => stat.panels > 0).length, [signalPortStats]);
  const powerStartKeys = useMemo(
    () =>
      new Set(
        Object.values(powerPortStats)
          .map((stat) => stat.path[0])
          .filter(Boolean)
          .map((cell) => `${cell!.x}-${cell!.y}`),
      ),
    [powerPortStats],
  );

  const flyBarWeight = activeColsCount * panel.defaults.flyBarWeight;
  const slingWeight = activeColsCount * panel.defaults.slingWeight;
  const powerCableWeight = powerPortsUsed * 3;
  const signalCableWeight = signalPortsUsed * 1;
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

  const totalPowerMaxW = totalPanels * powerSpec.maxW;
  const totalPowerMaxA = totalPanels * powerSpec.maxA;
  const totalPowerAvgW = totalPanels * powerSpec.avgW;
  const totalPowerAvgA = totalPanels * powerSpec.avgA;
  const unassignedPowerPanels = activeCells.filter((cell) => !cell.assignedPowerPort).length;

  const sparePanels = Math.ceil(totalPanels * panel.defaults.spareRatio);
  const totalPanelsWithSpare = totalPanels + sparePanels;
  const boxCount = Math.ceil(totalPanelsWithSpare / panel.defaults.panelsPerBox);
  const boxSparePanels = boxCount * panel.defaults.panelsPerBox - totalPanelsWithSpare;
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

  const deploymentWarning = useMemo(() => {
    if ((deploymentType === DEPLOYMENT_TYPES.GROUND || deploymentType === DEPLOYMENT_TYPES.FLOOR) && panelType !== "MG9") {
      return `${deploymentType} deployment hardware is currently available for MG9 only.`;
    }
    if (deploymentType === DEPLOYMENT_TYPES.FLOOR && ((activeWallWidthM % 1 !== 0) || (activeWallHeightM % 1 !== 0))) {
      return "Floor deployment uses full 1m frame sections only. This wall size is not an exact ground-frame build.";
    }
    return "";
  }, [activeWallHeightM, activeWallWidthM, deploymentType, panelType]);

  const stockRows = useMemo(() => {
    const stock = panel.stock as Record<string, number>;
    const rowsOut: StockRow[] = [];
    const pushBaseRow = (code: string, name: string, required: number, stockQty: number, method: string) => {
      rowsOut.push({ code, name, required, stock: stockQty, net: stockQty - required, method });
    };

    if (panelType === "MG9") {
      pushBaseRow("12224", "MG9 LED Panel", totalPanelsWithSpare, stock.panels ?? 0, `${totalPanels} + ${sparePanels} spare`);
    } else {
      pushBaseRow("12223", "MT Mesh Panel", totalPanelsWithSpare, stock.panels ?? 0, `${totalPanels} + ${sparePanels} spare`);
    }

    rowsOut.push(makeStockRow(STOCK_CATALOG.prodCase, 1, "always 1 per project"));
    rowsOut.push({ code: "BOX", name: "Boxes required", required: boxCount, stock: boxCount, net: 0, method: `ceil(${totalPanelsWithSpare}/${panel.defaults.panelsPerBox})` });

    if (deploymentType === DEPLOYMENT_TYPES.FLOWN) {
      rowsOut.push({
        code: panelType === "MG9" ? "12257" : "12262",
        name: panelType === "MG9" ? "MG9 Floor / Hanging Bar" : "MT Floor / Hanging Bar",
        required: activeColsCount,
        stock: stock.hangingBar ?? 0,
        net: (stock.hangingBar ?? 0) - activeColsCount,
        method: "1 per active column",
      });
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
    pushBaseRow("12263", "15m Signal Cable", signalCableTotalRequired, stock.signalCable15m ?? 0, `${signalCableWithBackupRequired}${backupSignalLoop ? ` (${signalCableBaseRequired} x 2 backup loop)` : ""} + ${signalCableSpare} spare`);

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

    if (panelType === "MG9" && includeReinforcementPlate) {
      pushBaseRow("12264", "MG9 Reinforcement Plate", Math.ceil(totalPanels * 0.86), stock.reinforcementPlate ?? 0, "sheet-style factor");
      pushBaseRow("12265", "MG9 Reinforcement Screw", Math.ceil(totalPanels * 3.42), stock.reinforcementScrew ?? 0, "sheet-style factor");
    }

    if (panelType === "MG9" && deploymentType === DEPLOYMENT_TYPES.GROUND) {
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

    if (panelType === "MG9" && deploymentType === DEPLOYMENT_TYPES.FLOOR) {
      const feet = Math.ceil(totalPanels / 2);
      const perimeterSegments = activeColsCount * 2 + activeRowsCount * 2;
      rowsOut.push(makeStockRow(STOCK_CATALOG.danceFloorFeet, feet, "1 per 2 panels"));
      rowsOut.push(makeStockRow(STOCK_CATALOG.temperedGlass, totalPanels, "1 per panel"));
      rowsOut.push(makeStockRow(STOCK_CATALOG.floorReinforcementBar, feet, "1 per foot"));
      rowsOut.push(makeStockRow(STOCK_CATALOG.floorTaperPin, feet * 4, "4 per foot"));
      rowsOut.push(makeStockRow(STOCK_CATALOG.danceFloorRamp, perimeterSegments, `${perimeterSegments} external 500mm edge segments`));
      rowsOut.push(makeStockRow(STOCK_CATALOG.danceFloorRampCorner, 4, "1 per corner"));
    }

    return rowsOut;
  }, [activeColsCount, activeRowsCount, activeWallWidthM, backupSignalLoop, boxCount, circuitsUsedMax, deploymentType, distroRequired, includeReinforcementPlate, panel, panelType, powerCableTotalRequired, powerDistro, signalCableBaseRequired, signalCableSpare, signalCableTotalRequired, signalCableWithBackupRequired, signalPortsUsed, sparePanels, totalPanels, totalPanelsWithSpare, powerPortsUsed, distro.portCount]);

  const shortfallRows = stockRows.filter((row) => row.required > 0 && row.net < 0);
  const safeProjectName = projectName.trim() || "Untitled Project";
  const fileSafeProjectName = safeProjectName.replace(/[<>:"/\\|?*\x00-\x1F]/g, "-").replace(/\s+/g, "-");

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

    const drawArrowHead = (x1: number, y1: number, x2: number, y2: number, color: string) => {
      const angle = Math.atan2(y2 - y1, x2 - x1);
      const size = 16;
      const baseX1 = x2 - size * Math.cos(angle - Math.PI / 6);
      const baseY1 = y2 - size * Math.sin(angle - Math.PI / 6);
      const baseX2 = x2 - size * Math.cos(angle + Math.PI / 6);
      const baseY2 = y2 - size * Math.sin(angle + Math.PI / 6);
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
    };

    ctx.fillStyle = "#0f172a";
    ctx.font = "16px Arial";
    ctx.textAlign = "center";
    for (let i = 0; i < cols; i += 1) {
      const displayIndex = flipped ? cols - i : i + 1;
      ctx.fillText(String(displayIndex), i * (CELL_SIZE + GRID_GAP) + CELL_SIZE / 2, -10);
    }

    ctx.textAlign = "left";
    for (let i = 0; i < rows; i += 1) {
      ctx.fillText(String(i + 1), -28, i * (CELL_SIZE + GRID_GAP) + CELL_SIZE / 2 + 6);
    }

    grid.flat().forEach((cell) => {
      if (!isActiveCell(cell)) return;
      const displayCell = getDisplayCell(cell, cols, flipped);
      const x = displayCell.x * (CELL_SIZE + GRID_GAP);
      const y = cell.y * (CELL_SIZE + GRID_GAP);
      const fill = cell.assignedPort ? PORT_COLORS[(cell.assignedPort - 1) % PORT_COLORS.length] : "#1e293b";
      ctx.fillStyle = fill;
      ctx.fillRect(x, y, CELL_SIZE, CELL_SIZE);
      ctx.strokeStyle = "#0f172a";
      ctx.lineWidth = 2;
      ctx.strokeRect(x, y, CELL_SIZE, CELL_SIZE);
      if (powerStartKeys.has(`${cell.x}-${cell.y}`)) {
        ctx.strokeStyle = POWER_COLOR;
        ctx.lineWidth = 4;
        ctx.strokeRect(x + 2, y + 2, CELL_SIZE - 4, CELL_SIZE - 4);
      }
    });

    Object.entries(signalPortStats).forEach(([portId, stat]) => {
      if (!stat.path || stat.path.length < 2) return;
      const color = PORT_COLORS[(Number(portId) - 1) % PORT_COLORS.length];
      ctx.strokeStyle = color;
      ctx.lineWidth = 4;
      ctx.lineCap = "round";
      stat.path.forEach((cell, idx) => {
        if (idx === 0) return;
        const prev = getDisplayCell(stat.path[idx - 1], cols, flipped);
        const current = getDisplayCell(cell, cols, flipped);
        let { x1, y1, x2, y2 } = getLineEndpoints(prev, current, 0);
        if (current.y !== prev.y) {
          const sideOffset = GRID_GAP * 0.5;
          x1 += flipped ? sideOffset : -sideOffset;
          x2 += flipped ? sideOffset : -sideOffset;
        }
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.stroke();
        drawArrowHead(x1, y1, x2, y2, color);
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
        const prev = getDisplayCell(path[idx - 1], cols, flipped);
        const current = getDisplayCell(cell, cols, flipped);
        let { x1, y1, x2, y2 } = getLineEndpoints(prev, current, 4);
        if (current.y !== prev.y) {
          const sideOffset = GRID_GAP * 0.5;
          x1 += flipped ? -sideOffset : sideOffset;
          x2 += flipped ? -sideOffset : sideOffset;
        }
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.stroke();
        drawArrowHead(x1, y1, x2, y2, POWER_COLOR);
      });
    });

    grid.flat().forEach((cell) => {
      if (!isActiveCell(cell)) return;
      const displayCell = getDisplayCell(cell, cols, flipped);
      const x = displayCell.x * (CELL_SIZE + GRID_GAP);
      const y = cell.y * (CELL_SIZE + GRID_GAP);
      ctx.fillStyle = "#020617";
      ctx.font = "bold 10px Arial";
      ctx.textAlign = "center";
      ctx.fillText(`↓ ${cell.y + 1} → ${displayCell.x + 1}`, x + CELL_SIZE / 2, y + 18);
      if (cell.assignedPort) ctx.fillText(`🔌 P${cell.assignedPort} (${cell.sequence ?? "-"})`, x + CELL_SIZE / 2, y + 34);
      if (cell.assignedPowerPort) ctx.fillText(`⚡ Plug ${cell.assignedPowerPort}`, x + CELL_SIZE / 2, y + 50);
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


  const openJson = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(String(ev.target?.result || "{}")) as OpenJsonPayload;
        const nextCols = Math.max(1, Number(data.wall?.cols || cols));
        const nextRows = Math.max(1, Number(data.wall?.rows || rows));
        const nextGrid = Array.isArray(data.patching?.grid) ? normalizeGrid(data.patching?.grid, nextCols, nextRows) : makeGrid(nextCols, nextRows);

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
      pdf.text(`Panel type: ${panel.name}`, 10, 26);
      pdf.text(`Power distro: ${distro.label}`, 10, 32);
      pdf.text(`Panels: ${cols} x ${rows} grid, ${totalPanels} active`, 10, 38);

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

    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(18);
    pdf.text(safeProjectName, 10, 12);
    pdf.setFont("helvetica", "normal");
    pdf.setFontSize(10);
    pdf.text(`Printed ${printedAt}`, 10, 18);

    drawInfoBox("Wall", [
      `Panel type: ${panel.name}`,
      `Power distro: ${distro.label}`,
      `Panels: ${cols} x ${rows} grid, ${totalPanels} active`,
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
      `Backup signal loop: ${backupSignalLoop ? "Yes" : "No"}`,
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

    drawInfoBox("Stock Summary", stockRows.slice(0, 10).map((row) =>
      `${row.code}: req ${row.required}, net ${row.net} ${row.net < 0 ? "(short)" : ""}`
    ), 10, 128, 184, 56);

    drawInfoBox("Shortfalls", shortfallRows.length
      ? shortfallRows.map((row) => `${row.code}: short by ${Math.abs(row.net)}`)
      : ["No stock shortfalls detected"], 198, 128, 88, 56);

    const backLayoutCanvas = buildLayoutCanvas(false, "Back View");
    const frontLayoutCanvas = buildLayoutCanvas(true, "Front View");
    drawLayoutPage(backLayoutCanvas, "Back View");
    drawLayoutPage(frontLayoutCanvas, "Front View");
    addPdfFooters();
    pdf.save(`${fileSafeProjectName}-${panelType}-${cols}x${rows}.pdf`);
  } catch (err) {
    console.error("PDF failed", err);
    alert("PDF failed - check console");
  }
};

  const applyGridSize = () => {
    const nextCols = Number.parseInt(draftCols, 10);
    const nextRows = Number.parseInt(draftRows, 10);
    if (!Number.isFinite(nextCols) || !Number.isFinite(nextRows) || nextCols < 1 || nextRows < 1) return;

    setCols(nextCols);
    setRows(nextRows);
    setGrid(makeGrid(nextCols, nextRows));
    setSelectedCell(null);
    setDragVisited(new Set());
    setIsDragging(false);
  };

  const assignSignalCell = (x: number, y: number) => {
    const key = `${x}-${y}`;
    if (dragVisited.has(key)) return;

    setGrid((prev) => {
      const current = prev[y]?.[x];
      if (!current) return prev;
      if (!isActiveCell(current)) return prev;

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
    const key = `${x}-${y}`;
    if (dragVisited.has(key)) return;

    setGrid((prev) => {
      const current = prev[y]?.[x];
      if (!current) return prev;
      if (!isActiveCell(current)) return prev;

      const currentPanels = getPortPanelCount(prev, "assignedPowerPort", activePowerPort);
      const isAlreadySamePort = current.assignedPowerPort === activePowerPort;
      if (!isAlreadySamePort && currentPanels >= safePanelsPerPowerOutlet) return prev;

      const currentPortLoad = getPowerPortLoadWatts(prev, activePowerPort, powerSpec.maxW, { x, y });
      if (!isAlreadySamePort && currentPortLoad + powerSpec.maxW > MAX_OUTLET_AMPS * VOLTAGE) return prev;

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
    if (!isActiveCell(target)) return;
    setSelectedCell(null);
    setDragVisited(new Set());
    setIsDragging(true);
    if (patchMode === "signal") assignSignalCell(x, y);
    else assignPowerCell(x, y);
  };

  const continueDrag = (x: number, y: number) => {
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

    setGrid((prev) => {
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

    setGrid((prev) => {
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

    setGrid((prev) => {
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
            if (!isActiveCell(next[y]?.[x])) return;
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
            if (!isActiveCell(next[y]?.[x])) return;
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
          if (!isActiveCell(next[y]?.[x])) return;
          while (portIndex < powerPorts.length) {
            const port = powerPorts[portIndex];
            const currentLoad = getPowerPortLoadWatts(next, port.id, powerSpec.maxW);
            const currentPanels = getPortPanelCount(next, "assignedPowerPort", port.id);

            if (currentPanels >= safePanelsPerPowerOutlet) {
              portIndex += 1;
              continue;
            }

            if (currentLoad + powerSpec.maxW <= MAX_OUTLET_AMPS * VOLTAGE) {
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
    setDragVisited(new Set());
    setIsDragging(false);
  };

  const clearSignalCabling = () => {
    setGrid((prev) => clearSignalOnGrid(prev));
    setSelectedCell(null);
    setDragVisited(new Set());
    setIsDragging(false);
  };

  const clearPowerAssignments = () => {
    setGrid((prev) => clearPowerOnGrid(prev));
    setSelectedCell(null);
  };

  const clearSelectedPanelPatching = () => {
    if (!selectedCell) return;
    setGrid((prev) => {
      const next = cloneGrid(prev);
      const target = next[selectedCell.y]?.[selectedCell.x];
      if (!target || !isActiveCell(target)) return prev;
      target.assignedPort = null;
      target.sequence = null;
      target.assignedPowerPort = null;
      target.powerSequence = null;
      target.powerManual = false;
      return next;
    });
  };

  const deleteSelectedPanel = () => {
    if (!selectedCell) return;
    setGrid((prev) => {
      const next = cloneGrid(prev);
      const target = next[selectedCell.y]?.[selectedCell.x];
      if (!target || target.isRemoved) return prev;
      target.assignedPort = null;
      target.sequence = null;
      target.assignedPowerPort = null;
      target.powerSequence = null;
      target.powerManual = false;
      target.isRemoved = true;
      return next;
    });
  };

  const restoreSelectedPanel = () => {
    if (!selectedCell) return;
    setGrid((prev) => {
      const next = cloneGrid(prev);
      const target = next[selectedCell.y]?.[selectedCell.x];
      if (!target || !target.isRemoved) return prev;
      target.isRemoved = false;
      target.assignedPort = null;
      target.sequence = null;
      target.assignedPowerPort = null;
      target.powerSequence = null;
      target.powerManual = false;
      return next;
    });
  };

  const toDisplayX = (x: number) => (isFlippedView ? flipX(x, cols) : x);
  const fromDisplayX = (displayX: number) => (isFlippedView ? flipX(displayX, cols) : displayX);

  const svgW = cols * CELL_SIZE + (cols - 1) * GRID_GAP;
  const svgH = rows * CELL_SIZE + (rows - 1) * GRID_GAP;

  return (
    <div className="min-h-screen bg-[#0f172a] p-6 text-white print-container">
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
              <span className="rounded-full border border-slate-500 bg-slate-800 px-3 py-1 text-xs font-semibold text-slate-200">v{APP_VERSION}</span>
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <Button className="border-cyan-400 bg-cyan-600 hover:bg-cyan-500" onClick={generatePdf}>
              <FileText className="mr-2 h-4 w-4" />Generate PDF
            </Button>
            <Button variant="outline" onClick={exportJson}>
              <Download className="mr-2 h-4 w-4" />Download Settings
            </Button>
            <Button variant="outline" onClick={() => fileInputRef.current?.click()}>
              <Upload className="mr-2 h-4 w-4" />Open Settings
            </Button>
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

              <div className="flex flex-wrap gap-2 no-print">
                <Button
                  variant="outline"
                  className={patchMode === "signal" ? "border-sky-300 bg-sky-200 font-semibold text-slate-950 hover:bg-sky-100" : ""}
                  onClick={() => setPatchMode("signal")}
                >
                  Signal Patch Mode
                </Button>
                <Button
                  variant="outline"
                  className={patchMode === "power" ? "border-amber-300 bg-amber-200 font-semibold text-slate-950 hover:bg-amber-100" : ""}
                  onClick={() => setPatchMode("power")}
                >
                  <Zap className="mr-2 h-4 w-4" />Power Patch Mode
                </Button>
              </div>

              <div className={`rounded-lg border px-4 py-3 text-sm font-medium [text-shadow:none] no-print ${
                patchMode === "signal" ? "border-sky-300 bg-sky-100 text-slate-950" : "border-amber-300 bg-amber-100 text-slate-950"
              }`}>
                Current mode: {patchMode === "signal" ? `Signal patching on port ${activePort}` : `Power patching on plug ${activePowerPort}`}
              </div>

              <div className="flex flex-wrap items-center gap-2 no-print">
                <select className="rounded bg-white p-2 text-black" value={snakeDirection} onChange={(e) => setSnakeDirection(e.target.value)}>
                  <option value="LR">Left to Right</option>
                  <option value="RL">Right to Left</option>
                  <option value="LRB">Left to Right from the Bottom</option>
                  <option value="RLB">Right to Left from the Bottom</option>
                  <option value="TB">Top to Bottom</option>
                  <option value="BT">Bottom to Top</option>
                  <option value="LOOP_TOGETHER">Loop together</option>
                </select>
                <label className="flex items-center gap-2 rounded border border-slate-600 bg-slate-900 px-3 py-2 text-sm text-white">
                  <input type="checkbox" checked={snakeAlternates} onChange={() => setSnakeAlternates((prev) => !prev)} />
                  <span>Snake / alternate direction</span>
                </label>
                <Button className="border-violet-400 bg-violet-600 hover:bg-violet-500" onClick={snakePatch}><Wand2 className="mr-2 h-4 w-4" />Auto Snake</Button>
                <Button className="border-rose-400 bg-rose-600 hover:bg-rose-500" onClick={clearSignalCabling}>Clear Signal</Button>
                <Button className="border-orange-400 bg-orange-600 hover:bg-orange-500" onClick={clearPowerAssignments}>Clear Power</Button>
              </div>
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
                  <div>Ports used: {signalPortsUsed}</div>
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
                    <span>Fly Bar ({panel.defaults.flyBarWeight}kg per column) → {flyBarWeight.toFixed(1)} kg</span>
                  </label>

                  <label className="flex items-center gap-2">
                    <input type="checkbox" checked={includeSling} onChange={() => setIncludeSling(!includeSling)} />
                    <span>Sling &amp; Shackle ({panel.defaults.slingWeight}kg per column) → {slingWeight.toFixed(1)} kg</span>
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

        <Card className="border-slate-700 bg-slate-800 print-card">
          <CardHeader>
            <div className="flex flex-wrap items-center justify-between gap-3">
              <CardTitle className="text-white [text-shadow:0_0_2px_black]">Panel Layout ({wallWidthM}m x {wallHeightM}m) - {patchMode === "signal" ? "Signal" : "Power"} patching</CardTitle>
              <div className="flex items-center gap-2 no-print">
                <div className={`rounded-full border px-3 py-1 text-xs font-semibold ${isFlippedView ? "border-amber-300 bg-amber-100 text-slate-950" : "border-sky-300 bg-sky-100 text-slate-950"}`}>
                  {isFlippedView ? "Current: Front View" : "Current: Back View"}
                </div>
                <Button variant="outline" className="text-sm" onClick={() => setIsFlippedView((prev) => !prev)}>
                  {isFlippedView ? "Show Back View" : "Show Front View"}
                </Button>
              </div>
            </div>
          </CardHeader>
          <CardContent>
            <div className="w-full overflow-auto rounded-xl bg-white/5 p-4 pt-6 pl-8 select-none">
              <div className="relative" style={{ width: svgW, height: svgH }}>
                <div className="absolute left-0 top-[-20px] grid text-xs text-white [text-shadow:0_0_2px_black]" style={{ gridTemplateColumns: `repeat(${cols}, ${CELL_SIZE}px)`, gap: GRID_GAP }}>
                  {Array.from({ length: cols }).map((_, index) => <div key={`col-${index}`} className="text-center">{isFlippedView ? cols - index : index + 1}</div>)}
                </div>

                <div className="absolute left-[-30px] top-0 grid text-xs text-white [text-shadow:0_0_2px_black]" style={{ gridTemplateRows: `repeat(${rows}, ${CELL_SIZE}px)`, gap: GRID_GAP }}>
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
                      const prev = getDisplayCell(stat.path[idx - 1], cols, isFlippedView);
                      const current = getDisplayCell(cell, cols, isFlippedView);
                      let { x1, y1, x2, y2 } = getLineEndpoints(prev, current, 0);

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
                      const prev = getDisplayCell(path[idx - 1], cols, isFlippedView);
                      const current = getDisplayCell(cell, cols, isFlippedView);
                      let { x1, y1, x2, y2 } = getLineEndpoints(prev, current, 4);

                      if (current.y !== prev.y) {
                        const sideOffset = GRID_GAP * 0.5;
                        x1 += isFlippedView ? -sideOffset : sideOffset;
                        x2 += isFlippedView ? -sideOffset : sideOffset;
                      }

                      return <line key={`pow-${port.id}-${idx}`} x1={x1} y1={y1} x2={x2} y2={y2} stroke={POWER_COLOR} style={{ color: POWER_COLOR }} strokeWidth="4" markerEnd="url(#arrow)" />;
                    });
                  })}
                </svg>

                <div className="absolute inset-0 z-10 grid" style={{ gridTemplateColumns: `repeat(${cols}, ${CELL_SIZE}px)`, gap: GRID_GAP }}>
                  {grid.flat().map((cell) => {
                    const displayX = toDisplayX(cell.x);
                    const key = `${displayX}-${cell.y}`;
                    const originalKey = `${cell.x}-${cell.y}`;
                    const signalStat = cell.assignedPort ? signalPortStats[cell.assignedPort] : null;
                    const isEdge = signalStat?.firstKey === originalKey || signalStat?.lastKey === originalKey;
                    const isSelected = selectedCell?.x === cell.x && selectedCell?.y === cell.y;
                    const isPowerStart = powerStartKeys.has(originalKey);
                    const isRemoved = cell.isRemoved;
                    const displayColor = isRemoved ? "transparent" : cell.assignedPort ? PORT_COLORS[(cell.assignedPort - 1) % PORT_COLORS.length] : "#1e293b";

                    return (
                      <div
                        key={key}
                        onMouseDown={() => startDrag(fromDisplayX(displayX), cell.y)}
                        onMouseEnter={() => continueDrag(fromDisplayX(displayX), cell.y)}
                        onClick={() => setSelectedCell({ x: cell.x, y: cell.y })}
                        style={{
                          width: CELL_SIZE,
                          height: CELL_SIZE,
                          background: displayColor,
                          border: `2px ${isRemoved ? "dashed" : "solid"} ${isSelected ? "#ffffff" : isRemoved ? "#64748b" : isEdge ? "black" : "#334155"}`,
                          boxShadow: !isRemoved && isPowerStart ? `0 0 0 3px ${POWER_COLOR}` : "none",
                          color: isRemoved ? "#94a3b8" : "#020617",
                          gridColumnStart: displayX + 1,
                          gridRowStart: cell.y + 1,
                        }}
                        className="flex cursor-pointer select-none flex-col items-center justify-center gap-[2px] p-1 text-[9px] font-semibold leading-tight tracking-tight"
                      >
                        {isRemoved ? (
                          <div className="text-center text-[10px] font-bold uppercase tracking-wide">Removed</div>
                        ) : (
                          <>
                            <div>{`↓ ${cell.y + 1} → ${displayX + 1}`}</div>
                            {cell.assignedPort ? <div className="whitespace-nowrap">{`🔌 P${cell.assignedPort} (${cell.sequence ?? "-"})`}</div> : null}
                            {cell.assignedPowerPort ? <div className="whitespace-nowrap">{`⚡ Plug ${cell.assignedPowerPort}`}</div> : null}
                          </>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>

            {selectedPanel ? (
              <div className="mt-4 rounded bg-slate-900 p-3 text-white [text-shadow:0_0_2px_black] print-card no-print">
                <div className="mb-2">{selectedPanel.isRemoved ? "Removed Panel Position" : "Panel"} ({selectedDisplayCell!.x + 1}, {selectedDisplayCell!.y + 1})</div>
                <div className="mb-2 grid gap-2 text-xs md:grid-cols-2">
                  <div>
                    <div>Size: {panel.w}m x {panel.h}m</div>
                    <div>Pixels: {panel.pixW} x {panel.pixH}</div>
                    <div>Weight: {selectedPanel.isRemoved ? "Removed" : `${panel.weight} kg`}</div>
                  </div>
                  <div>
                    <div>Max power: {powerSpec.maxW} W / {powerSpec.maxA.toFixed(2)} A</div>
                    <div>Average power: {powerSpec.avgW} W / {powerSpec.avgA.toFixed(2)} A</div>
                  </div>
                </div>
                {!selectedPanel.isRemoved ? (
                  <>
                    <div className="grid gap-2 md:grid-cols-2">
                      <select className="rounded bg-white p-2 text-black" onChange={(e) => applyManualSignalPatch(e.target.value)} value={selectedPanel.assignedPort ?? ""}>
                        <option value="">Unpatch signal</option>
                        {signalPorts.map((port) => <option key={port.id} value={port.id}>{port.name}</option>)}
                      </select>
                      <select className="rounded bg-white p-2 text-black" onChange={(e) => applyManualPowerPatch(e.target.value)} value={selectedPanel.assignedPowerPort ?? ""}>
                        <option value="">Unpatch power</option>
                        {powerPorts.map((port) => <option key={port.id} value={port.id}>{port.name}</option>)}
                      </select>
                    </div>
                    <div className="mt-3 flex flex-wrap gap-2">
                      <Button className="border-rose-400 bg-rose-600 hover:bg-rose-500" onClick={clearSelectedPanelPatching}>Clear Power And Signal Patching</Button>
                      <Button className="border-slate-400 bg-slate-700 hover:bg-slate-600" onClick={deleteSelectedPanel}>Delete Panel</Button>
                    </div>
                  </>
                ) : (
                  <div className="mt-3 flex flex-wrap gap-2">
                    <Button className="border-emerald-400 bg-emerald-600 hover:bg-emerald-500" onClick={restoreSelectedPanel}>Restore Panel</Button>
                  </div>
                )}
              </div>
            ) : null}
          </CardContent>
        </Card>

        <Card className="border-slate-700 bg-slate-800 print-card no-print">
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

        <Card className="border-slate-700 bg-slate-800 print-card no-print">
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

