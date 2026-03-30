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

const PORT_COLORS = [
  "#d946ef",
  "#2563eb",
  "#dc2626",
  "#0891b2",
  "#7c3aed",
  "#ea580c",
  "#be123c",
  "#0f766e",
  "#4338ca",
  "#b45309",
  "#1d4ed8",
  "#9d174d",
  "#0369a1",
  "#6d28d9",
  "#c2410c",
  "#1f2937",
  "#0f766e",
  "#4f46e5",
  "#b91c1c",
  "#7e22ce",
];

type PanelTypeKey = keyof typeof PANEL_TYPES;
type PowerDistroKey = keyof typeof POWER_DISTROS;

type Cell = {
  x: number;
  y: number;
  assignedPort: number | null;
  sequence: number | null;
  assignedPowerPort: number | null;
  powerSequence: number | null;
  powerManual: boolean;
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
    })),
  );

const cloneGrid = (grid: Cell[][]): Cell[][] => grid.map((row) => row.map((cell) => ({ ...cell })));

const formatNumber = (value: number, digits = 0) =>
  Number(value || 0).toLocaleString(undefined, {
    maximumFractionDigits: digits,
    minimumFractionDigits: digits,
  });

const getStatusColor = (percent: number) => {
  if (percent >= 90) return "#ef4444";
  if (percent >= 70) return "#f59e0b";
  return "#22c55e";
};

const clampActivePort = (value: number, max: number) => Math.min(Math.max(value, 1), max);

const clearSignalOnGrid = (grid: Cell[][]) =>
  grid.map((row) => row.map((cell) => ({ ...cell, assignedPort: null, sequence: null })));

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
      if (excludeCell && cell.x === excludeCell.x && cell.y === excludeCell.y) continue;
      if (cell.assignedPowerPort === portId) watts += panelPowerMaxW;
    }
  }
  return watts;
};

const getPortPanelCount = (grid: Cell[][], portField: "assignedPort" | "assignedPowerPort", portId: number) =>
  grid.flat().filter((cell) => cell[portField] === portId).length;

const getSnakeOrder = (cols: number, rows: number, snakeDirection: string) => {
  const ordered: Array<{ x: number; y: number }> = [];

  if (snakeDirection === "LR" || snakeDirection === "RL") {
    for (let y = 0; y < rows; y += 1) {
      let row = [...Array(cols).keys()];
      if (snakeDirection === "RL") row.reverse();
      if (y % 2 === 1) row.reverse();
      row.forEach((x) => ordered.push({ x, y }));
    }
  } else {
    for (let x = 0; x < cols; x += 1) {
      let col = [...Array(rows).keys()];
      if (snakeDirection === "BT") col.reverse();
      if (x % 2 === 1) col.reverse();
      col.forEach((y) => ordered.push({ x, y }));
    }
  }

  return ordered;
};

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
    <div className="h-2 w-full rounded bg-black/30">
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

  const panel = PANEL_TYPES[panelType];
  const powerSpec = panel.power;
  const distro = POWER_DISTROS[powerDistro];
  const powerPorts = useMemo(() => makePowerPorts(distro.portCount), [distro.portCount]);

  const [panelsPerPowerOutlet, setPanelsPerPowerOutlet] = useState(panel.defaults.powerPanelsPerOutlet);
  const [panelsPerSignalPort, setPanelsPerSignalPort] = useState(panel.defaults.signalPanelsPerPort);

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
  const totalPanels = rows * cols;
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

  const flyBarWeight = cols * panel.defaults.flyBarWeight;
  const slingWeight = cols * panel.defaults.slingWeight;
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
  const unassignedPowerPanels = grid.flat().filter((cell) => !cell.assignedPowerPort).length;
  const selectedPanel = selectedCell ? grid[selectedCell.y]?.[selectedCell.x] ?? null : null;

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

  const stockRows = useMemo(() => {
    const stock = panel.stock as Record<string, number>;
    if (panelType === "MG9") {
      return [
        { code: "12224", name: "MG9 LED Panel", required: totalPanelsWithSpare, stock: stock.panels ?? 0, net: (stock.panels ?? 0) - totalPanelsWithSpare, method: `${totalPanels} + ${sparePanels} spare` },
        { code: "BOX", name: "Boxes required", required: boxCount, stock: boxCount, net: 0, method: `ceil(${totalPanelsWithSpare}/${panel.defaults.panelsPerBox})` },
        { code: "12257", name: "MG9 Floor / Hanging Bar", required: cols, stock: stock.hangingBar ?? 0, net: (stock.hangingBar ?? 0) - cols, method: `1 per column` },
        { code: powerDistro === "32A" ? "12245" : "12246", name: powerDistro === "32A" ? "32A 3Φ Power Distro" : "63A 3Φ Power Distro", required: Math.max(1, Math.ceil(powerPortsUsed / distro.portCount)), stock: powerDistro === "32A" ? stock.distro32 ?? 0 : stock.distro63 ?? 0, net: (powerDistro === "32A" ? stock.distro32 ?? 0 : stock.distro63 ?? 0) - Math.max(1, Math.ceil(powerPortsUsed / distro.portCount)), method: `selected distro` },
        { code: "12254", name: "15m PowerCON Cable", required: circuitsUsedMax + Math.ceil(circuitsUsedMax * panel.defaults.powerSpareRatio), stock: stock.powerCable15m ?? 0, net: (stock.powerCable15m ?? 0) - (circuitsUsedMax + Math.ceil(circuitsUsedMax * panel.defaults.powerSpareRatio)), method: `${circuitsUsedMax} + ${Math.ceil(circuitsUsedMax * panel.defaults.powerSpareRatio)} spare` },
        { code: "12263", name: "15m Signal Cable", required: signalPortsUsed + Math.ceil(signalPortsUsed * panel.defaults.signalSpareRatio), stock: stock.signalCable15m ?? 0, net: (stock.signalCable15m ?? 0) - (signalPortsUsed + Math.ceil(signalPortsUsed * panel.defaults.signalSpareRatio)), method: `${signalPortsUsed} + ${Math.ceil(signalPortsUsed * panel.defaults.signalSpareRatio)} spare` },
        { code: "12264", name: "MG9 Reinforcement Plate", required: Math.ceil(totalPanels * 0.86), stock: stock.reinforcementPlate ?? 0, net: (stock.reinforcementPlate ?? 0) - Math.ceil(totalPanels * 0.86), method: `sheet-style factor` },
        { code: "12265", name: "MG9 Reinforcement Screw", required: Math.ceil(totalPanels * 3.42), stock: stock.reinforcementScrew ?? 0, net: (stock.reinforcementScrew ?? 0) - Math.ceil(totalPanels * 3.42), method: `sheet-style factor` },
      ];
    }

    return [
      { code: "12223", name: "MT Mesh Panel", required: totalPanelsWithSpare, stock: stock.panels ?? 0, net: (stock.panels ?? 0) - totalPanelsWithSpare, method: `${totalPanels} + ${sparePanels} spare` },
      { code: "BOX", name: "Boxes required", required: boxCount, stock: boxCount, net: 0, method: `ceil(${totalPanelsWithSpare}/${panel.defaults.panelsPerBox})` },
      { code: "12262", name: "MT Floor / Hanging Bar", required: cols, stock: stock.hangingBar ?? 0, net: (stock.hangingBar ?? 0) - cols, method: `1 per column` },
      { code: powerDistro === "32A" ? "12245" : "12246", name: powerDistro === "32A" ? "32A 3Φ Power Distro" : "63A 3Φ Power Distro", required: Math.max(1, Math.ceil(powerPortsUsed / distro.portCount)), stock: powerDistro === "32A" ? stock.distro32 ?? 0 : stock.distro63 ?? 0, net: (powerDistro === "32A" ? stock.distro32 ?? 0 : stock.distro63 ?? 0) - Math.max(1, Math.ceil(powerPortsUsed / distro.portCount)), method: `selected distro` },
    ];
  }, [panel, panelType, totalPanelsWithSpare, totalPanels, sparePanels, boxCount, cols, powerDistro, powerPortsUsed, distro.portCount, circuitsUsedMax, signalPortsUsed]);

  const shortfallRows = stockRows.filter((row) => row.required > 0 && row.net < 0);
  const safeProjectName = projectName.trim() || "Untitled Project";
  const fileSafeProjectName = safeProjectName.replace(/[<>:"/\\|?*\x00-\x1F]/g, "-").replace(/\s+/g, "-");

  const ensurePdfSpace = (pdf: any, y: number, needed = 10) => {
    const pageHeight = pdf.internal.pageSize.getHeight();
    if (y + needed <= pageHeight - 12) return y;
    pdf.addPage("a4", "portrait");
    return 18;
  };

  const addPdfLine = (pdf: any, label: string, value: string, y: number) => {
    const nextY = ensurePdfSpace(pdf, y, 8);
    pdf.setFont("helvetica", "bold");
    pdf.text(label, 14, nextY);
    pdf.setFont("helvetica", "normal");
    pdf.text(value, 68, nextY);
    return nextY + 6;
  };

  const addPdfSectionTitle = (pdf: any, title: string, y: number) => {
    const nextY = ensurePdfSpace(pdf, y, 12);
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(14);
    pdf.text(title, 14, nextY);
    pdf.setDrawColor(203, 213, 225);
    pdf.line(14, nextY + 2, 196, nextY + 2);
    pdf.setFontSize(11);
    return nextY + 10;
  };

  const buildLayoutCanvas = () => {
    const canvas = document.createElement("canvas");
    canvas.width = svgW + 96;
    canvas.height = svgH + 96;
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Canvas context unavailable");

    ctx.fillStyle = "#0f172a";
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    const offsetX = 56;
    const offsetY = 40;
    ctx.save();
    ctx.translate(offsetX, offsetY);

    ctx.fillStyle = "#e2e8f0";
    ctx.font = "16px Arial";
    ctx.textAlign = "center";
    for (let i = 0; i < cols; i += 1) {
      ctx.fillText(String(i + 1), i * (CELL_SIZE + GRID_GAP) + CELL_SIZE / 2, -10);
    }

    ctx.textAlign = "left";
    for (let i = 0; i < rows; i += 1) {
      ctx.fillText(String(i + 1), -28, i * (CELL_SIZE + GRID_GAP) + CELL_SIZE / 2 + 6);
    }

    Object.entries(signalPortStats).forEach(([portId, stat]) => {
      if (!stat.path || stat.path.length < 2) return;
      const color = PORT_COLORS[(Number(portId) - 1) % PORT_COLORS.length];
      ctx.strokeStyle = color;
      ctx.lineWidth = 4;
      stat.path.forEach((cell, idx) => {
        if (idx === 0) return;
        const prev = stat.path[idx - 1];
        let { x1, y1, x2, y2 } = getLineEndpoints(prev, cell, 0);
        if (cell.y !== prev.y) {
          const sideOffset = GRID_GAP * 0.5;
          x1 -= sideOffset;
          x2 -= sideOffset;
        }
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.stroke();
      });
    });

    powerPorts.forEach((port) => {
      const stat = powerPortStats[port.id];
      const path = stat?.path ?? [];
      if (path.length < 2) return;
      ctx.strokeStyle = POWER_COLOR;
      ctx.lineWidth = 4;
      path.forEach((cell, idx) => {
        if (idx === 0) return;
        const prev = path[idx - 1];
        let { x1, y1, x2, y2 } = getLineEndpoints(prev, cell, 4);
        if (cell.y !== prev.y) {
          const sideOffset = GRID_GAP * 0.5;
          x1 += sideOffset;
          x2 += sideOffset;
        }
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.stroke();
      });
    });

    grid.flat().forEach((cell) => {
      const x = cell.x * (CELL_SIZE + GRID_GAP);
      const y = cell.y * (CELL_SIZE + GRID_GAP);
      const fill = cell.assignedPort ? PORT_COLORS[(cell.assignedPort - 1) % PORT_COLORS.length] : "#1e293b";
      ctx.fillStyle = fill;
      ctx.fillRect(x, y, CELL_SIZE, CELL_SIZE);
      ctx.strokeStyle = "#e2e8f0";
      ctx.lineWidth = 2;
      ctx.strokeRect(x, y, CELL_SIZE, CELL_SIZE);

      ctx.fillStyle = "#ffffff";
      ctx.font = "bold 10px Arial";
      ctx.textAlign = "center";
      ctx.fillText(`↓ ${cell.y + 1} → ${cell.x + 1}`, x + CELL_SIZE / 2, y + 18);
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


  const openJson = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(String(ev.target?.result || "{}")) as OpenJsonPayload;
        const nextCols = Math.max(1, Number(data.wall?.cols || cols));
        const nextRows = Math.max(1, Number(data.wall?.rows || rows));
        const nextGrid = Array.isArray(data.patching?.grid) ? data.patching?.grid : makeGrid(nextCols, nextRows);

        if (data.projectName) setProjectName(data.projectName);
        if (data.panelType && PANEL_TYPES[data.panelType]) setPanelType(data.panelType);
        if (data.powerDistro && POWER_DISTROS[data.powerDistro]) setPowerDistro(data.powerDistro);

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
    const pdf = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
    let y = 18;

    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(18);
    pdf.text(safeProjectName, 14, y);
    y += 8;
    pdf.setFont("helvetica", "normal");
    pdf.setFontSize(10);
    pdf.text(`Generated ${new Date().toLocaleString()}`, 14, y);
    y += 10;
    pdf.setFontSize(11);

    y = addPdfSectionTitle(pdf, "Wall", y);
    y = addPdfLine(pdf, "Panel type", panel.name, y);
    y = addPdfLine(pdf, "Power distro", distro.label, y);
    y = addPdfLine(pdf, "Panels", `${cols} x ${rows} = ${totalPanels}`, y);
    y = addPdfLine(pdf, "Size", `${wallWidthM}m x ${wallHeightM}m`, y);
    y = addPdfLine(pdf, "Resolution", `${wallPixelW} x ${wallPixelH}`, y);
    y = addPdfLine(pdf, "Aspect ratio", aspectRatio, y);
    y = addPdfLine(pdf, "Reduced ratio", ratioLabel, y);

    y = addPdfSectionTitle(pdf, "Power", y);
    y = addPdfLine(pdf, "Max draw", `${formatNumber(totalPowerMaxW)} W / ${formatNumber(totalPowerMaxA, 2)} A`, y);
    y = addPdfLine(pdf, "Average draw", `${formatNumber(totalPowerAvgW)} W / ${formatNumber(totalPowerAvgA, 2)} A`, y);
    y = addPdfLine(pdf, "Circuits used", String(circuitsUsedMax), y);
    y = addPdfLine(pdf, "Per outlet", `${formatNumber(powerPerCircuitMaxW)} W / ${formatNumber(powerPerCircuitMaxA, 2)} A`, y);
    y = addPdfLine(pdf, "Signal ports used", String(signalPortsUsed), y);
    y = addPdfLine(pdf, "Power ports used", String(powerPortsUsed), y);

    y = addPdfSectionTitle(pdf, "Weight + Output", y);
    y = addPdfLine(pdf, "Total weight", `${totalWeight.toFixed(1)} kg`, y);
    y = addPdfLine(pdf, "VX1000 use", `${formatNumber(vx1000Percent, 1)}%`, y);
    y = addPdfLine(pdf, "VX2000 use", `${formatNumber(vx2000Percent, 1)}%`, y);
    y = addPdfLine(pdf, "Best standard output", bestResolution ? `${bestResolution[0]} x ${bestResolution[1]}` : "None in preset list", y);

    y = addPdfSectionTitle(pdf, "Phase Load", y);
    Object.entries(phaseStats).forEach(([phase, stat]) => {
      y = addPdfLine(pdf, `Phase ${phase.replace("P", "")}`, `${formatNumber(stat.maxWatts)} W / ${formatNumber(stat.maxAmps, 2)} A, avg ${formatNumber(stat.avgWatts)} W / ${formatNumber(stat.avgAmps, 2)} A, ${formatNumber(stat.utilisation, 1)}%`, y);
    });

    y = addPdfSectionTitle(pdf, "Stock Summary", y);
    stockRows.forEach((row) => {
      y = addPdfLine(pdf, row.name, `Req ${row.required}, Stock ${row.stock}, Net ${row.net}, ${row.method}`, y);
    });

    y = addPdfSectionTitle(pdf, "Additional Details", y);
    y = addPdfLine(pdf, "Fly bar included", includeFlyBar ? `${flyBarWeight.toFixed(1)} kg` : "No", y);
    y = addPdfLine(pdf, "Sling included", includeSling ? `${slingWeight.toFixed(1)} kg` : "No", y);
    y = addPdfLine(pdf, "Power cable included", includePowerCable ? `${powerCableWeight.toFixed(1)} kg` : "No", y);
    y = addPdfLine(pdf, "Signal cable included", includeSignalCable ? `${signalCableWeight.toFixed(1)} kg` : "No", y);
    y = addPdfLine(pdf, "Custom weight", includeCustomWeight ? `${customWeight} kg` : "No", y);
    y = addPdfLine(pdf, "Spare panels", String(sparePanels), y);
    y = addPdfLine(pdf, "Panels incl. spare", String(totalPanelsWithSpare), y);
    y = addPdfLine(pdf, "Boxes", `${boxCount} (box spare panels ${boxSparePanels})`, y);

    const layoutCanvas = buildLayoutCanvas();
    pdf.addPage("a4", "landscape");
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(18);
    pdf.text(`${safeProjectName} - Panel Layout`, 10, 12);
    const usableWidth = pageWidth - 20;
    const usableHeight = pageHeight - 22;
    const layoutRatio = layoutCanvas.width / layoutCanvas.height;
    let drawWidth = usableWidth;
    let drawHeight = drawWidth / layoutRatio;
    if (drawHeight > usableHeight) {
      drawHeight = usableHeight;
      drawWidth = drawHeight * layoutRatio;
    }
    pdf.addImage(layoutCanvas.toDataURL("image/png"), "PNG", 10 + (usableWidth - drawWidth) / 2, 16 + (usableHeight - drawHeight) / 2, drawWidth, drawHeight);
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
    setSelectedCell(null);
    setDragVisited(new Set());
    setIsDragging(true);
    if (patchMode === "signal") assignSignalCell(x, y);
    else assignPowerCell(x, y);
  };

  const continueDrag = (x: number, y: number) => {
    if (!isDragging) return;
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
    const ordered = getSnakeOrder(cols, rows, snakeDirection);

    setGrid((prev) => {
      const next = cloneGrid(prev);

      if (patchMode === "signal") {
        for (const row of next) {
          for (const cell of row) {
            cell.assignedPort = null;
            cell.sequence = null;
          }
        }

        let port = 1;
        let seq = 1;
        ordered.forEach(({ x, y }) => {
          if (port > SIGNAL_PORT_COUNT) return;
          next[y][x].assignedPort = port;
          next[y][x].sequence = seq;
          seq += 1;
          if (seq > safePanelsPerSignalPort) {
            port += 1;
            seq = 1;
          }
        });
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
            <h1 className="text-3xl font-semibold text-white [text-shadow:0_0_2px_black]">LED Port Mapper</h1>
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
                  <option value="TB">Top to Bottom</option>
                  <option value="BT">Bottom to Top</option>
                </select>
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
                  <div>Panels: {cols} × {rows} = {totalPanels}</div>
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

              <div className="border-t border-slate-700 pt-3 space-y-2 no-print">
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
            <CardTitle className="text-white [text-shadow:0_0_2px_black]">Panel Layout ({wallWidthM}m x {wallHeightM}m) - {patchMode === "signal" ? "Signal" : "Power"} patching</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="w-full overflow-auto rounded-xl bg-white/5 p-4 pt-6 pl-8">
              <div className="relative" style={{ width: svgW, height: svgH }}>
                <div className="absolute left-0 top-[-20px] grid text-xs text-white [text-shadow:0_0_2px_black]" style={{ gridTemplateColumns: `repeat(${cols}, ${CELL_SIZE}px)`, gap: GRID_GAP }}>
                  {Array.from({ length: cols }).map((_, index) => <div key={`col-${index}`} className="text-center">{index + 1}</div>)}
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
                      const prev = stat.path[idx - 1];
                      let { x1, y1, x2, y2 } = getLineEndpoints(prev, cell, 0);

                      if (cell.y !== prev.y) {
                        const sideOffset = GRID_GAP * 0.5;
                        x1 -= sideOffset;
                        x2 -= sideOffset;
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
                      const prev = path[idx - 1];
                      let { x1, y1, x2, y2 } = getLineEndpoints(prev, cell, 4);

                      if (cell.y !== prev.y) {
                        const sideOffset = GRID_GAP * 0.5;
                        x1 += sideOffset;
                        x2 += sideOffset;
                      }

                      return <line key={`pow-${port.id}-${idx}`} x1={x1} y1={y1} x2={x2} y2={y2} stroke={POWER_COLOR} style={{ color: POWER_COLOR }} strokeWidth="4" markerEnd="url(#arrow)" />;
                    });
                  })}
                </svg>

                <div className="absolute inset-0 z-10 grid" style={{ gridTemplateColumns: `repeat(${cols}, ${CELL_SIZE}px)`, gap: GRID_GAP }}>
                  {grid.flat().map((cell) => {
                    const key = `${cell.x}-${cell.y}`;
                    const signalStat = cell.assignedPort ? signalPortStats[cell.assignedPort] : null;
                    const isEdge = signalStat?.firstKey === key || signalStat?.lastKey === key;
                    const isSelected = selectedCell?.x === cell.x && selectedCell?.y === cell.y;
                    const displayColor = cell.assignedPort ? PORT_COLORS[(cell.assignedPort - 1) % PORT_COLORS.length] : "#1e293b";

                    return (
                      <div
                        key={key}
                        onMouseDown={() => startDrag(cell.x, cell.y)}
                        onMouseEnter={() => continueDrag(cell.x, cell.y)}
                        onClick={() => setSelectedCell({ x: cell.x, y: cell.y })}
                        style={{
                          width: CELL_SIZE,
                          height: CELL_SIZE,
                          background: displayColor,
                          border: `2px solid ${isSelected ? "#ffffff" : isEdge ? "black" : "#334155"}`,
                          color: "white",
                        }}
                        className="flex cursor-pointer flex-col items-center justify-center gap-[2px] p-1 text-[9px] font-semibold leading-tight tracking-tight"
                      >
                        <div>{`↓ ${cell.y + 1} → ${cell.x + 1}`}</div>
                        {cell.assignedPort ? <div className="whitespace-nowrap">{`🔌 P${cell.assignedPort} (${cell.sequence ?? "-"})`}</div> : null}
                        {cell.assignedPowerPort ? <div className="whitespace-nowrap">{`⚡ Plug ${cell.assignedPowerPort}`}</div> : null}
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>

            {selectedPanel ? (
              <div className="mt-4 rounded bg-slate-900 p-3 text-white [text-shadow:0_0_2px_black] print-card no-print">
                <div className="mb-2">Panel ({selectedCell!.x + 1}, {selectedCell!.y + 1})</div>
                <div className="mb-2 grid gap-2 text-xs md:grid-cols-2">
                  <div>
                    <div>Size: {panel.w}m x {panel.h}m</div>
                    <div>Pixels: {panel.pixW} x {panel.pixH}</div>
                    <div>Weight: {panel.weight} kg</div>
                  </div>
                  <div>
                    <div>Max power: {powerSpec.maxW} W / {powerSpec.maxA.toFixed(2)} A</div>
                    <div>Average power: {powerSpec.avgW} W / {powerSpec.avgA.toFixed(2)} A</div>
                  </div>
                </div>
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
              const loadPercent = safePanelsPerSignalPort > 0 ? Math.min(100, (stat.panels / safePanelsPerSignalPort) * 100) : 0;
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
                  <div className="mt-2 h-2 bg-black/30">
                    <div style={{ width: `${loadPercent}%`, background: port.color, height: "100%" }} />
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
                    <div className="mt-2 h-2 bg-black/30">
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
                      <div className="mt-2 h-2 bg-black/30">
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
            <CardTitle className="text-white [text-shadow:0_0_2px_black]">Stock Calculations</CardTitle>
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
              <table className="min-w-full text-left text-sm">
                <thead className="bg-slate-900">
                  <tr>
                    <th className="px-3 py-2">Code</th>
                    <th className="px-3 py-2">Item</th>
                    <th className="px-3 py-2 text-right">Required</th>
                    <th className="px-3 py-2 text-right">Stock</th>
                    <th className="px-3 py-2 text-right">Net</th>
                    <th className="px-3 py-2">Method</th>
                  </tr>
                </thead>
                <tbody>
                  {stockRows.map((row) => (
                    <tr key={`${row.code}-${row.name}`} className="border-t border-slate-700">
                      <td className="px-3 py-2">{row.code}</td>
                      <td className="px-3 py-2">{row.name}</td>
                      <td className="px-3 py-2 text-right">{formatNumber(row.required)}</td>
                      <td className="px-3 py-2 text-right">{formatNumber(row.stock)}</td>
                      <td className={`px-3 py-2 text-right ${row.net < 0 ? "text-red-300" : "text-emerald-300"}`}>{formatNumber(row.net)}</td>
                      <td className="px-3 py-2 text-xs text-slate-300">{row.method}</td>
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
