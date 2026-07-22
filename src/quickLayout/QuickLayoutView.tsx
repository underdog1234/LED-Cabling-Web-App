import React, { useState } from "react";
import { ChevronUp, ChevronDown, FileText } from "lucide-react";
import { PANEL_TYPES, type PanelTypeKey } from "../App";
import { Button, Card, CardHeader, CardContent, CardTitle, Input, Select } from "../components/ui";

// Must match QUICK_LAYOUT_TRANSFER_KEY in App.tsx.
const QUICK_LAYOUT_TRANSFER_KEY = "ledCablingQuickLayoutTransfer:v1";

const MIN_CELLS = 1;
const MAX_CELLS = 100;
const clampCells = (n: number) => Math.min(MAX_CELLS, Math.max(MIN_CELLS, Math.round(n) || MIN_CELLS));

const gcd = (a: number, b: number): number => (b === 0 ? a : gcd(b, a % b));

const formatM = (n: number) => `${n.toFixed(2)} m`;

// How far apart to space ruler marks along an axis so a large wall doesn't
// end up with dozens of overlapping labels - grows through "nice" round
// metre steps until at most ~12 marks are needed.
const RULER_STEPS = [0.5, 1, 2, 5, 10, 20, 25, 50, 100, 200, 500];
const pickRulerStep = (totalM: number) => RULER_STEPS.find((step) => totalM / step <= 12) ?? RULER_STEPS[RULER_STEPS.length - 1];
const rulerMarks = (totalM: number) => {
  if (totalM <= 0) return [0];
  const step = pickRulerStep(totalM);
  const marks: number[] = [];
  for (let m = 0; m <= totalM + 1e-6; m += step) marks.push(Math.round(m * 100) / 100);
  const last = marks[marks.length - 1];
  if (Math.abs(last - totalM) > 1e-6) marks.push(Math.round(totalM * 100) / 100);
  return marks;
};

// How far the centred 16:9 box can be nudged up/down per click.
const FIT_SHIFT_STEP = 0.15;

// Standalone panel-count calculator: opened in its own tab (see App.tsx's
// "Quick Panel Layout" button and main.jsx's ?quicklayout=1 route), with no
// dependency on the main app ever having been mounted - it always starts at
// a neutral 1x1 MG9 default and works from a bookmarked/typed URL alone.
export default function QuickLayoutView() {
  const [panelType, setPanelType] = useState<PanelTypeKey>("MG9");
  const [cols, setCols] = useState(1);
  const [rows, setRows] = useState(1);
  // -1 (top) .. 0 (centred) .. 1 (bottom); only has visible effect when the
  // wall has vertical slack around the 16:9 box (see fitSlackY below).
  const [fitShift, setFitShift] = useState(0);

  const panel = PANEL_TYPES[panelType];
  const wallWidthM = cols * panel.w;
  const wallHeightM = rows * panel.h;
  const pixelW = cols * panel.pixW;
  const pixelH = rows * panel.pixH;
  const totalPanels = cols * rows;
  const totalPixels = pixelW * pixelH;

  const ratioDivisor = gcd(pixelW, pixelH) || 1;
  const ratioLabel = `${pixelW / ratioDivisor}:${pixelH / ratioDivisor}`;

  // Largest 16:9 rect that fits inside the wall's pixel resolution, nudged
  // up/down within whatever vertical slack is available (fitShift).
  const fitsWide = pixelW / pixelH > 16 / 9;
  const fitW = fitsWide ? (pixelH * 16) / 9 : pixelW;
  const fitH = fitsWide ? pixelH : (pixelW * 9) / 16;
  const fitOffsetX = (pixelW - fitW) / 2;
  const fitSlackY = pixelH - fitH;
  const fitOffsetY = (fitSlackY / 2) * (1 + fitShift);

  const wallBelowFullHd = pixelW < 1920 || pixelH < 1080;
  const contentBelowFullHd = fitW < 1920 || fitH < 1080;

  const setWidthM = (valueM: number) => setCols(clampCells(Math.round(valueM / panel.w)));
  const setHeightM = (valueM: number) => setRows(clampCells(Math.round(valueM / panel.h)));

  const clearAll = () => {
    setPanelType("MG9");
    setCols(1);
    setRows(1);
    setFitShift(0);
  };

  const sendToMainTool = () => {
    localStorage.setItem(QUICK_LAYOUT_TRANSFER_KEY, JSON.stringify({ panelType, cols, rows }));
    window.location.href = window.location.pathname;
  };

  // One-page summary: stats on the left, a to-scale wall diagram (grid lines,
  // ruler marks and the 16:9 overlay) on the right - a PDF twin of the
  // on-screen Preview card, not the main app's full per-panel report.
  const exportPdf = async () => {
    try {
      const jsPDF = (await import("jspdf")).default;
      const pdf = new jsPDF({ unit: "mm", format: "a4", orientation: "landscape" });
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();

      pdf.setFontSize(18);
      pdf.setTextColor(15, 23, 42);
      pdf.text("Quick Panel Layout", 10, 14);
      pdf.setFontSize(9);
      pdf.setTextColor(100, 116, 139);
      pdf.text(`Printed ${new Date().toLocaleString()}`, pageW - 10, 14, { align: "right" });

      let statsY = 28;
      const stat = (label: string, value: string) => {
        pdf.setFontSize(9);
        pdf.setTextColor(100, 116, 139);
        pdf.text(label, 10, statsY);
        pdf.setFontSize(12);
        pdf.setTextColor(15, 23, 42);
        pdf.text(value, 10, statsY + 6);
        statsY += 14;
      };
      stat("Panel Type", panelType === "MT" ? "MT (1m x 0.5m)" : "MG9 (0.5m x 0.5m)");
      stat("Grid", `${cols} columns x ${rows} rows (${totalPanels} panels)`);
      stat("Wall Size", `${formatM(wallWidthM)} x ${formatM(wallHeightM)}`);
      stat("Resolution", `${pixelW} x ${pixelH} px (${totalPixels.toLocaleString()} px total)`);
      stat("Aspect Ratio", ratioLabel);
      stat("16:9 Content Area", `${Math.round(fitW)} x ${Math.round(fitH)} px`);

      const warnings: string[] = [];
      if (wallBelowFullHd) warnings.push("Wall resolution is below 1920x1080 (Full HD).");
      if (!wallBelowFullHd && contentBelowFullHd) warnings.push("The 16:9 content area is below 1920x1080 (Full HD).");
      if (warnings.length) {
        pdf.setFontSize(9);
        pdf.setTextColor(180, 83, 9);
        warnings.forEach((line, i) => pdf.text(`⚠ ${line}`, 10, statsY + i * 6));
        pdf.setTextColor(15, 23, 42);
      }

      // Diagram, scaled to fit its reserved area (small top/left margin for
      // the ruler labels) while preserving the wall's true aspect ratio.
      const diagAreaX = 110;
      const diagAreaY = 34;
      const diagAreaW = pageW - diagAreaX - 12;
      const diagAreaH = pageH - diagAreaY - 16;
      const scale = Math.min(diagAreaW / pixelW, diagAreaH / pixelH);
      const boxX = diagAreaX;
      const boxY = diagAreaY;
      const boxW = pixelW * scale;
      const boxH = pixelH * scale;

      pdf.setDrawColor(100, 116, 139);
      pdf.setLineWidth(0.3);
      pdf.rect(boxX, boxY, boxW, boxH);

      pdf.setDrawColor(203, 213, 225);
      pdf.setLineWidth(0.1);
      for (let c = 1; c < cols; c += 1) {
        const x = boxX + (c / cols) * boxW;
        pdf.line(x, boxY, x, boxY + boxH);
      }
      for (let r = 1; r < rows; r += 1) {
        const y = boxY + (r / rows) * boxH;
        pdf.line(boxX, y, boxX + boxW, y);
      }

      const rectX = boxX + (fitOffsetX / pixelW) * boxW;
      const rectY = boxY + (fitOffsetY / pixelH) * boxH;
      const rectW = (fitW / pixelW) * boxW;
      const rectH = (fitH / pixelH) * boxH;
      pdf.setDrawColor(245, 158, 11);
      pdf.setLineWidth(0.5);
      pdf.setLineDashPattern([1.5, 1], 0);
      pdf.rect(rectX, rectY, rectW, rectH);
      pdf.setLineDashPattern([], 0);

      pdf.setFontSize(7);
      pdf.setTextColor(100, 116, 139);
      rulerMarks(wallWidthM).forEach((m) => {
        const x = boxX + (wallWidthM > 0 ? (m / wallWidthM) * boxW : 0);
        pdf.text(`${m}m`, x, boxY - 2, { align: "center" });
      });
      rulerMarks(wallHeightM).forEach((m) => {
        const y = boxY + (wallHeightM > 0 ? (m / wallHeightM) * boxH : 0);
        pdf.text(`${m}m`, boxX - 2, y + 1, { align: "right" });
      });

      pdf.save(`quick-panel-layout-${panelType}-${cols}x${rows}.pdf`);
    } catch (err) {
      console.error("Quick Panel Layout PDF export failed", err);
      alert("PDF export failed - check console");
    }
  };

  const topMarks = rulerMarks(wallWidthM);
  const sideMarks = rulerMarks(wallHeightM);

  return (
    <div className="min-h-screen bg-[#0f172a] p-6 text-white">
      <div className="mx-auto max-w-5xl">
        <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
          <div>
            <div className="text-xl font-bold">Quick Panel Layout</div>
            <div className="text-sm text-slate-400">
              A standalone calculator for panel counts, resolution and aspect ratio.{" "}
              <a href={location.pathname} className="text-sky-400 hover:underline">
                Back to LED Cabling Planner
              </a>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <Button intent="secondary" onClick={clearAll}>Clear</Button>
            <Button intent="secondary" onClick={exportPdf}>
              <FileText className="h-4 w-4" />Export PDF
            </Button>
            <Button intent="primary" onClick={sendToMainTool}>Send to Main Layout Tool</Button>
          </div>
        </div>

        <div className="grid gap-4 md:grid-cols-[1fr_1.3fr]">
          <Card>
            <CardHeader>
              <CardTitle>Grid</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              <div>
                <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">Panel Type</div>
                <Select value={panelType} onChange={(e: React.ChangeEvent<HTMLSelectElement>) => setPanelType(e.target.value as PanelTypeKey)}>
                  <option value="MG9">MG9 (0.5m × 0.5m)</option>
                  <option value="MT">MT (1m × 0.5m)</option>
                </Select>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">Width (m)</div>
                  <Input
                    type="number"
                    min={panel.w}
                    step={panel.w}
                    value={wallWidthM}
                    onChange={(e: React.ChangeEvent<HTMLInputElement>) => setWidthM(Number(e.target.value))}
                    className="text-center"
                  />
                </div>
                <div>
                  <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">Height (m)</div>
                  <Input
                    type="number"
                    min={panel.h}
                    step={panel.h}
                    value={wallHeightM}
                    onChange={(e: React.ChangeEvent<HTMLInputElement>) => setHeightM(Number(e.target.value))}
                    className="text-center"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">Columns</div>
                  <div className="flex items-center gap-1">
                    <Button size="sm" intent="secondary" onClick={() => setCols((c) => clampCells(c - 1))}>-</Button>
                    <Input
                      type="number"
                      min={MIN_CELLS}
                      max={MAX_CELLS}
                      value={cols}
                      onChange={(e: React.ChangeEvent<HTMLInputElement>) => setCols(clampCells(Number(e.target.value)))}
                      className="text-center"
                    />
                    <Button size="sm" intent="secondary" onClick={() => setCols((c) => clampCells(c + 1))}>+</Button>
                  </div>
                </div>
                <div>
                  <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">Rows</div>
                  <div className="flex items-center gap-1">
                    <Button size="sm" intent="secondary" onClick={() => setRows((r) => clampCells(r - 1))}>-</Button>
                    <Input
                      type="number"
                      min={MIN_CELLS}
                      max={MAX_CELLS}
                      value={rows}
                      onChange={(e: React.ChangeEvent<HTMLInputElement>) => setRows(clampCells(Number(e.target.value)))}
                      className="text-center"
                    />
                    <Button size="sm" intent="secondary" onClick={() => setRows((r) => clampCells(r + 1))}>+</Button>
                  </div>
                </div>
              </div>

              <div className="rounded-lg border border-slate-700/70 bg-slate-900/40 p-3 text-sm">
                <dl className="grid grid-cols-2 gap-x-4 gap-y-1.5">
                  <dt className="text-slate-400">Panel count</dt>
                  <dd>{totalPanels}</dd>
                  <dt className="text-slate-400">Wall size</dt>
                  <dd>{formatM(wallWidthM)} × {formatM(wallHeightM)}</dd>
                  <dt className="text-slate-400">Resolution</dt>
                  <dd>{pixelW} × {pixelH} px ({totalPixels.toLocaleString()} px total)</dd>
                  <dt className="text-slate-400">Aspect ratio</dt>
                  <dd>{ratioLabel}</dd>
                  <dt className="text-slate-400">16:9 content area</dt>
                  <dd>{Math.round(fitW)} × {Math.round(fitH)} px</dd>
                </dl>
              </div>

              {wallBelowFullHd ? (
                <div className="rounded-lg border border-amber-500 bg-amber-500/15 p-2 text-xs text-amber-200">
                  ⚠ Wall resolution is below 1920×1080 (Full HD).
                </div>
              ) : null}
              {!wallBelowFullHd && contentBelowFullHd ? (
                <div className="rounded-lg border border-amber-500 bg-amber-500/15 p-2 text-xs text-amber-200">
                  ⚠ The 16:9 content area is below 1920×1080 (Full HD).
                </div>
              ) : null}
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle>Preview</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="mx-auto grid" style={{ maxWidth: 640, gridTemplateColumns: "28px 1fr", gridTemplateRows: "20px 1fr" }}>
                <div />
                <div className="relative">
                  {topMarks.map((m) => (
                    <span
                      key={`t-${m}`}
                      className="absolute -translate-x-1/2 whitespace-nowrap text-[10px] text-slate-400"
                      style={{ left: `${wallWidthM > 0 ? (m / wallWidthM) * 100 : 0}%` }}
                    >
                      {m}m
                    </span>
                  ))}
                </div>
                <div className="relative">
                  {sideMarks.map((m) => (
                    <span
                      key={`s-${m}`}
                      className="absolute -translate-y-1/2 whitespace-nowrap text-[10px] text-slate-400"
                      style={{ top: `${wallHeightM > 0 ? (m / wallHeightM) * 100 : 0}%` }}
                    >
                      {m}m
                    </span>
                  ))}
                </div>
                <div
                  className="relative overflow-hidden rounded-lg border border-slate-600 bg-slate-950"
                  style={{
                    aspectRatio: `${pixelW} / ${pixelH}`,
                    backgroundImage:
                      "linear-gradient(to right, rgba(148,163,184,0.45) 1px, transparent 1px), linear-gradient(to bottom, rgba(148,163,184,0.45) 1px, transparent 1px)",
                    backgroundSize: `${100 / cols}% ${100 / rows}%`,
                  }}
                >
                  <div
                    className="absolute border-2 border-dashed border-amber-400/90 bg-amber-400/10"
                    style={{
                      left: `${(fitOffsetX / pixelW) * 100}%`,
                      top: `${(fitOffsetY / pixelH) * 100}%`,
                      width: `${(fitW / pixelW) * 100}%`,
                      height: `${(fitH / pixelH) * 100}%`,
                    }}
                  />
                </div>
              </div>
              <div className="mt-2 flex items-center justify-center gap-2 text-xs text-slate-400">
                <span>Dashed box = 16:9</span>
                <Button
                  size="sm"
                  intent="secondary"
                  disabled={fitSlackY <= 0 || fitShift <= -1}
                  onClick={() => setFitShift((f) => Math.max(-1, f - FIT_SHIFT_STEP))}
                  title="Move 16:9 area up"
                >
                  <ChevronUp className="h-3.5 w-3.5" />
                </Button>
                <Button
                  size="sm"
                  intent="secondary"
                  disabled={fitSlackY <= 0 || fitShift >= 1}
                  onClick={() => setFitShift((f) => Math.min(1, f + FIT_SHIFT_STEP))}
                  title="Move 16:9 area down"
                >
                  <ChevronDown className="h-3.5 w-3.5" />
                </Button>
              </div>
            </CardContent>
          </Card>
        </div>
      </div>
    </div>
  );
}
