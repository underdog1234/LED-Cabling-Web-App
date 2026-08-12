import React, { useState } from "react";
import { ChevronUp, ChevronDown, FileText } from "lucide-react";
import { PANEL_TYPES, POWER_DISTROS, type PanelTypeKey, type PowerDistroKey } from "../App";
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
  const [showFitBox, setShowFitBox] = useState(true);
  // Optional project name - shown as a header on the PDF export and carried
  // forward to become the main tool's project name via Send to Main Layout Tool.
  const [projectName, setProjectName] = useState("");

  const panel = PANEL_TYPES[panelType];
  const wallWidthM = cols * panel.w;
  const wallHeightM = rows * panel.h;
  // LED Wall Pixel Count - the panel's own native pixel grid (e.g. MT is
  // 256x64 per panel), not a physical-size scaling of anything.
  const pixelW = cols * panel.pixW;
  const pixelH = rows * panel.pixH;
  const totalPanels = cols * rows;
  const totalPixels = pixelW * pixelH;

  // MT is a transparent panel missing every second LED row, so its vertical
  // pixel pitch (7.8mm) is twice its horizontal pitch (3.9mm) - the raw
  // pixW x pixH grid isn't a square-pixel raster and doesn't match the
  // panel's true physical aspect ratio. Recommended Content Resolution scales
  // the LED pixel height back up so content authored at that resolution
  // (square pixels) matches the wall's real proportions. Always exactly 1x
  // for panels like MG9 where the pitch already matches on both axes, so
  // everything below collapses back to today's plain single-resolution
  // behaviour for them.
  const pitchXmm = (panel.w * 1000) / panel.pixW;
  const pitchYmm = (panel.h * 1000) / panel.pixH;
  const hasSquarePixels = Math.abs(pitchYmm / pitchXmm - 1) < 1e-6;
  const contentPixelW = pixelW;
  const contentPixelH = hasSquarePixels ? pixelH : Math.round(pixelH * (pitchYmm / pitchXmm));

  // Physical Aspect Ratio - derived from the wall's true physical size (mm,
  // so the gcd reduction is exact), not the raw LED pixel grid, which for
  // non-square-pixel panels like MT gives a different (wrong) ratio.
  const wallWidthMm = Math.round(wallWidthM * 1000);
  const wallHeightMm = Math.round(wallHeightM * 1000);
  const ratioDivisor = gcd(wallWidthMm, wallHeightMm) || 1;
  const ratioLabel = wallHeightMm > 0 ? `${wallWidthMm / ratioDivisor}:${wallHeightMm / ratioDivisor}` : "-";

  // Power draw and distro sizing - uses the same per-panel power spec and
  // safe-panels-per-outlet default as the main Layout Tool (no per-panel
  // patching here, so this is an aggregate/sizing estimate, not a real
  // phase-balanced plan).
  const powerSpec = panel.power;
  const totalMaxW = totalPanels * powerSpec.maxW;
  const totalMaxA = totalPanels * powerSpec.maxA;
  const totalAvgW = totalPanels * powerSpec.avgW;
  const totalAvgA = totalPanels * powerSpec.avgA;
  const safePanelsPerOutlet = panel.defaults.powerPanelsPerOutlet;
  const summarizeDistro = (key: PowerDistroKey) => {
    const distro = POWER_DISTROS[key];
    const circuits = totalPanels > 0 ? Math.ceil(totalPanels / Math.max(safePanelsPerOutlet, 1)) : 0;
    const units = circuits > 0 ? Math.ceil(circuits / distro.portCount) : 0;
    // Each distro unit has 3 phases, each rated to distro.safePhaseWatts.
    const capacityW = units * distro.safePhaseWatts * 3;
    const utilisationPct = capacityW > 0 ? (totalMaxW / capacityW) * 100 : 0;
    return { distro, circuits, units, capacityW, utilisationPct };
  };
  const distro32 = summarizeDistro("32A");
  const distro63 = summarizeDistro("63A");

  // Weight estimate - fully automatic (deployment assumed Flown, no manual
  // rigging/cable input needed), reusing the same per-panel constants as the
  // main Layout Tool's own weight breakdown. Indicative planning figure
  // only, not a certified rigging calculation.
  const panelWeight = totalPanels * panel.weight;
  // Uniform grid, so every column has exactly one panel in the top row.
  const topRowPanelCount = totalPanels > 0 ? cols : 0;
  const flyBarWeight = topRowPanelCount * panel.defaults.flyBarWeight;
  const slingShackleWeight = topRowPanelCount * panel.defaults.slingWeight;
  const flyingHardwareWeight = flyBarWeight + slingShackleWeight;
  // Cable weight assumes cables snake left-to-right, alternating direction
  // each row (the same pattern Auto Snake uses) - at the port-count level
  // this only affects how many ports/outlets that pattern needs, not the
  // exact route, so it's the same ceil() math as the power/signal port
  // sizing above.
  const signalPortsUsedForCable = totalPanels > 0 ? Math.ceil(totalPanels / Math.max(panel.defaults.signalPanelsPerPort, 1)) : 0;
  const powerPortsUsedForCable = totalPanels > 0 ? Math.ceil(totalPanels / Math.max(safePanelsPerOutlet, 1)) : 0;
  const powerCableWeight = powerPortsUsedForCable * 3;
  const signalCableWeight = signalPortsUsedForCable * 1;
  const cableWeight = powerCableWeight + signalCableWeight;
  const totalFlownWeight = panelWeight + flyingHardwareWeight + cableWeight;

  // Largest 16:9 rect that fits inside the wall's RECOMMENDED CONTENT
  // resolution (square-pixel space, not the raw LED grid - see
  // contentPixelW/H above), nudged up/down within whatever vertical slack is
  // available (fitShift). Using the raw LED pixel grid here would produce a
  // box (and, for the preview/diagram shapes below, a wall outline) that's
  // visibly the wrong shape for non-square-pixel panels like MT.
  const fitsWide = contentPixelW / contentPixelH > 16 / 9;
  const fitW = fitsWide ? (contentPixelH * 16) / 9 : contentPixelW;
  const fitH = fitsWide ? contentPixelH : (contentPixelW * 9) / 16;
  const fitOffsetX = (contentPixelW - fitW) / 2;
  const fitSlackY = contentPixelH - fitH;
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
    setShowFitBox(true);
    setProjectName("");
  };

  const sendToMainTool = () => {
    const trimmedName = projectName.trim();
    localStorage.setItem(
      QUICK_LAYOUT_TRANSFER_KEY,
      JSON.stringify({ panelType, cols, rows, projectName: trimmedName || undefined }),
    );
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

      const trimmedProjectName = projectName.trim();
      if (trimmedProjectName) {
        pdf.setFontSize(12);
        pdf.setFont("helvetica", "bold");
        pdf.setTextColor(51, 65, 85);
        pdf.text(trimmedProjectName, 10, 21);
        pdf.setFont("helvetica", "normal");
        pdf.setTextColor(15, 23, 42);
      }

      // Stats are grouped into headed sections (Panel / Power / Weight) laid
      // out in a compact grid so the whole block - plus warnings - comfortably
      // fits above the diagram's left edge (diagAreaX below) within one page.
      let statsY = trimmedProjectName ? 30 : 26;
      const statsX = 10;
      const statsColW = 92;

      const sectionHeader = (title: string) => {
        pdf.setFontSize(10);
        pdf.setFont("helvetica", "bold");
        pdf.setTextColor(37, 99, 235);
        pdf.text(title.toUpperCase(), statsX, statsY);
        pdf.setDrawColor(37, 99, 235);
        pdf.setLineWidth(0.4);
        pdf.line(statsX, statsY + 1.3, statsX + statsColW, statsY + 1.3);
        pdf.setFont("helvetica", "normal");
        pdf.setTextColor(15, 23, 42);
        statsY += 7;
      };
      const statGrid = (items: Array<[string, string]>, columns: 1 | 2) => {
        const colW = statsColW / columns;
        const rowH = columns === 1 ? 10.5 : 11.5;
        const startY = statsY;
        items.forEach(([label, value], i) => {
          const col = i % columns;
          const row = Math.floor(i / columns);
          const x = statsX + col * colW;
          const y = startY + row * rowH;
          pdf.setFontSize(7.5);
          pdf.setTextColor(100, 116, 139);
          pdf.text(label, x, y);
          pdf.setFontSize(10);
          pdf.setTextColor(15, 23, 42);
          pdf.text(value, x, y + 4.5);
        });
        const rowCount = Math.ceil(items.length / columns);
        statsY = startY + rowCount * rowH + 5;
      };

      const panelStats: Array<[string, string]> = [
        ["Panel Type", panelType === "MT" ? "MT (1m x 0.5m)" : "MG9 (0.5m x 0.5m)"],
        ["Grid", `${cols} x ${rows} (${totalPanels} panels)`],
        [hasSquarePixels ? "Wall Size" : "Physical Size", `${formatM(wallWidthM)} x ${formatM(wallHeightM)}`],
      ];
      if (hasSquarePixels) {
        panelStats.push(["Resolution", `${pixelW} x ${pixelH} px`], ["Aspect Ratio", ratioLabel]);
      } else {
        panelStats.push(
          ["LED Wall Resolution", `${pixelW} x ${pixelH} px`],
          ["Recommended Content Resolution", `${contentPixelW} x ${contentPixelH} px`],
          ["Physical Aspect Ratio", ratioLabel],
        );
      }
      if (showFitBox) panelStats.push(["16:9 Content Area", `${Math.round(fitW)} x ${Math.round(fitH)} px`]);
      sectionHeader("Panel");
      statGrid(panelStats, 2);

      sectionHeader("Power");
      statGrid(
        [
          ["Total Power Draw", `Max ${totalMaxW.toLocaleString()} W / ${totalMaxA.toFixed(2)} A (Avg ${totalAvgW.toLocaleString()} W)`],
          ["32A Distro", `${distro32.circuits} circuits, ${distro32.units} unit(s), ${distro32.utilisationPct.toFixed(1)}% used`],
          ["63A Distro", `${distro63.circuits} circuits, ${distro63.units} unit(s), ${distro63.utilisationPct.toFixed(1)}% used`],
        ],
        1,
      );

      sectionHeader("Weight (Flown Estimate)");
      statGrid(
        [
          ["Panel Weight", `${panelWeight.toFixed(1)} kg`],
          ["Flying Hardware", `${flyingHardwareWeight.toFixed(1)} kg`],
          ["Cable Weight", `${cableWeight.toFixed(1)} kg`],
          ["Total Flown Weight", `${totalFlownWeight.toFixed(1)} kg`],
        ],
        2,
      );

      const warnings: string[] = [];
      if (wallBelowFullHd) warnings.push("Wall resolution is below 1920x1080 (Full HD).");
      if (showFitBox && !wallBelowFullHd && contentBelowFullHd) warnings.push("The 16:9 content area is below 1920x1080 (Full HD).");
      if (warnings.length) {
        pdf.setFontSize(9);
        pdf.setTextColor(180, 83, 9);
        warnings.forEach((line, i) => pdf.text(`⚠ ${line}`, 10, statsY + i * 6));
        pdf.setTextColor(15, 23, 42);
      }

      // Diagram, scaled to fit its reserved area (small top/left margin for
      // the ruler labels) while preserving the wall's true PHYSICAL aspect
      // ratio (contentPixelW/H, not the raw LED pixel grid - see above).
      const diagAreaX = 110;
      const diagAreaY = 34;
      const diagAreaW = pageW - diagAreaX - 12;
      const diagAreaH = pageH - diagAreaY - 16;
      const scale = Math.min(diagAreaW / contentPixelW, diagAreaH / contentPixelH);
      const boxX = diagAreaX;
      const boxY = diagAreaY;
      const boxW = contentPixelW * scale;
      const boxH = contentPixelH * scale;

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

      if (showFitBox) {
        const rectX = boxX + (fitOffsetX / contentPixelW) * boxW;
        const rectY = boxY + (fitOffsetY / contentPixelH) * boxH;
        const rectW = (fitW / contentPixelW) * boxW;
        const rectH = (fitH / contentPixelH) * boxH;
        pdf.setDrawColor(245, 158, 11);
        pdf.setLineWidth(0.5);
        pdf.setLineDashPattern([1.5, 1], 0);
        pdf.rect(rectX, rectY, rectW, rectH);
        pdf.setLineDashPattern([], 0);

        pdf.setFontSize(7.5);
        pdf.setTextColor(146, 64, 14);
        pdf.text(
          "Dashed box = largest 16:9 area that fits within the wall - recommended safe area for 16:9 content.",
          diagAreaX,
          pageH - 8,
        );
        pdf.setTextColor(15, 23, 42);
      }

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
                <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">Project Name (optional)</div>
                <Input
                  type="text"
                  value={projectName}
                  onChange={(e: React.ChangeEvent<HTMLInputElement>) => setProjectName(e.target.value)}
                  placeholder="Used as the PDF header - carries to Main Layout Tool"
                />
              </div>

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
                  <div className="flex items-center gap-1">
                    <Button size="sm" intent="secondary" onClick={() => setWidthM(wallWidthM - panel.w)}>-</Button>
                    <Input
                      type="number"
                      min={panel.w}
                      step={panel.w}
                      value={wallWidthM}
                      onChange={(e: React.ChangeEvent<HTMLInputElement>) => setWidthM(Number(e.target.value))}
                      className="text-center"
                    />
                    <Button size="sm" intent="secondary" onClick={() => setWidthM(wallWidthM + panel.w)}>+</Button>
                  </div>
                </div>
                <div>
                  <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">Height (m)</div>
                  <div className="flex items-center gap-1">
                    <Button size="sm" intent="secondary" onClick={() => setHeightM(wallHeightM - panel.h)}>-</Button>
                    <Input
                      type="number"
                      min={panel.h}
                      step={panel.h}
                      value={wallHeightM}
                      onChange={(e: React.ChangeEvent<HTMLInputElement>) => setHeightM(Number(e.target.value))}
                      className="text-center"
                    />
                    <Button size="sm" intent="secondary" onClick={() => setHeightM(wallHeightM + panel.h)}>+</Button>
                  </div>
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
                  <dt className="text-slate-400">{hasSquarePixels ? "Wall size" : "Physical size"}</dt>
                  <dd>{formatM(wallWidthM)} × {formatM(wallHeightM)}</dd>
                  {hasSquarePixels ? (
                    <>
                      <dt className="text-slate-400">Resolution</dt>
                      <dd>{pixelW} × {pixelH} px ({totalPixels.toLocaleString()} px total)</dd>
                      <dt className="text-slate-400">Aspect ratio</dt>
                      <dd>{ratioLabel}</dd>
                    </>
                  ) : (
                    <>
                      <dt className="text-slate-400">LED wall resolution</dt>
                      <dd>{pixelW} × {pixelH} px</dd>
                      <dt className="text-slate-400">Recommended content resolution</dt>
                      <dd>{contentPixelW} × {contentPixelH} px</dd>
                      <dt className="text-slate-400">Physical aspect ratio</dt>
                      <dd>{ratioLabel}</dd>
                    </>
                  )}
                  {showFitBox ? (
                    <>
                      <dt className="text-slate-400">16:9 content area</dt>
                      <dd>{Math.round(fitW)} × {Math.round(fitH)} px</dd>
                    </>
                  ) : null}
                </dl>
              </div>

              {wallBelowFullHd ? (
                <div className="rounded-lg border border-amber-500 bg-amber-500/15 p-2 text-xs text-amber-200">
                  ⚠ Wall resolution is below 1920×1080 (Full HD).
                </div>
              ) : null}
              {showFitBox && !wallBelowFullHd && contentBelowFullHd ? (
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
                    // Shaped by the wall's true PHYSICAL aspect ratio
                    // (contentPixelW/H), not the raw LED pixel grid - for
                    // MT those differ, and the raw grid's shape would look
                    // wrong here (see contentPixelH above).
                    aspectRatio: `${contentPixelW} / ${contentPixelH}`,
                    backgroundImage:
                      "linear-gradient(to right, rgba(148,163,184,0.45) 1px, transparent 1px), linear-gradient(to bottom, rgba(148,163,184,0.45) 1px, transparent 1px)",
                    backgroundSize: `${100 / cols}% ${100 / rows}%`,
                  }}
                >
                  {showFitBox ? (
                    <div
                      className="absolute border-2 border-dashed border-amber-400/90 bg-amber-400/10"
                      style={{
                        left: `${(fitOffsetX / contentPixelW) * 100}%`,
                        top: `${(fitOffsetY / contentPixelH) * 100}%`,
                        width: `${(fitW / contentPixelW) * 100}%`,
                        height: `${(fitH / contentPixelH) * 100}%`,
                      }}
                    />
                  ) : null}
                </div>
              </div>
              <div className="mt-2 flex flex-wrap items-center justify-center gap-2 text-xs text-slate-400">
                <label className="flex cursor-pointer items-center gap-1.5">
                  <input
                    type="checkbox"
                    checked={showFitBox}
                    onChange={(e: React.ChangeEvent<HTMLInputElement>) => setShowFitBox(e.target.checked)}
                    className="h-3.5 w-3.5 accent-amber-400"
                  />
                  Show 16:9 content area
                </label>
                {showFitBox ? (
                  <>
                    <span className="text-slate-600">|</span>
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
                  </>
                ) : null}
              </div>
            </CardContent>
          </Card>
        </div>

        <Card className="mt-4">
          <CardHeader>
            <CardTitle>Power</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="grid gap-3 md:grid-cols-3">
              <div className="rounded-lg border border-slate-700/70 bg-slate-900/40 p-3 text-sm">
                <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">Total power draw</div>
                <div>Max: {totalMaxW.toLocaleString()} W / {totalMaxA.toFixed(2)} A</div>
                <div>Avg: {totalAvgW.toLocaleString()} W / {totalAvgA.toFixed(2)} A</div>
              </div>
              {([["32A", distro32], ["63A", distro63]] as const).map(([key, summary]) => (
                <div key={key} className="rounded-lg border border-slate-700/70 bg-slate-900/40 p-3 text-sm">
                  <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">{summary.distro.label}</div>
                  <div>Circuits needed: {summary.circuits}</div>
                  <div>Distro units needed: {summary.units || 0}</div>
                  <div>Capacity used: {summary.utilisationPct.toFixed(1)}% of {summary.capacityW.toLocaleString()} W</div>
                  {summary.utilisationPct > 100 ? (
                    <div className="mt-1 text-amber-400">⚠ Over safe capacity - add another {key} distro</div>
                  ) : null}
                </div>
              ))}
            </div>
            <div className="mt-2 text-xs text-slate-500">
              Estimated sizing only - assumes {safePanelsPerOutlet} panels per outlet and even load across all phases (no per-panel patching in this tool).
            </div>
          </CardContent>
        </Card>

        <Card className="mt-4">
          <CardHeader>
            <CardTitle>Weight (Flown Estimate)</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="grid gap-3 md:grid-cols-4">
              <div className="rounded-lg border border-slate-700/70 bg-slate-900/40 p-3 text-sm">
                <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">Panel weight</div>
                <div className="text-lg font-semibold">{panelWeight.toFixed(1)} kg</div>
              </div>
              <div className="rounded-lg border border-slate-700/70 bg-slate-900/40 p-3 text-sm">
                <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">Flying hardware</div>
                <div className="text-lg font-semibold">{flyingHardwareWeight.toFixed(1)} kg</div>
                <div className="mt-1 text-xs text-slate-400">Fly bars + slings/shackles, {topRowPanelCount} top-row panel(s)</div>
              </div>
              <div className="rounded-lg border border-slate-700/70 bg-slate-900/40 p-3 text-sm">
                <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">Cable weight</div>
                <div className="text-lg font-semibold">{cableWeight.toFixed(1)} kg</div>
                <div className="mt-1 text-xs text-slate-400">{powerPortsUsedForCable} power + {signalPortsUsedForCable} signal run(s)</div>
              </div>
              <div className="rounded-lg border border-sky-700/70 bg-sky-900/20 p-3 text-sm">
                <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">Total flown weight</div>
                <div className="text-lg font-semibold">{totalFlownWeight.toFixed(1)} kg</div>
              </div>
            </div>
            <div className="mt-2 text-xs text-slate-500">
              Indicative estimate only, not a certified rigging calculation. Assumes a Flown deployment and cables snaking left-to-right, alternating direction each row - always confirm rigging with a qualified rigger before flying a screen.
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
