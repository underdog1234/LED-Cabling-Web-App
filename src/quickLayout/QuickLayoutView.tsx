import React, { useState } from "react";
import { PANEL_TYPES, type PanelTypeKey } from "../App";
import { Button, Card, CardHeader, CardContent, CardTitle, Input, Select } from "../components/ui";

// Must match QUICK_LAYOUT_TRANSFER_KEY in App.tsx.
const QUICK_LAYOUT_TRANSFER_KEY = "ledCablingQuickLayoutTransfer:v1";

const MIN_CELLS = 1;
const MAX_CELLS = 100;
const clampCells = (n: number) => Math.min(MAX_CELLS, Math.max(MIN_CELLS, Math.round(n) || MIN_CELLS));

const gcd = (a: number, b: number): number => (b === 0 ? a : gcd(b, a % b));

const formatM = (n: number) => `${n.toFixed(2)} m`;

// Standalone panel-count calculator: opened in its own tab (see App.tsx's
// "Quick Panel Layout" button and main.jsx's ?quicklayout=1 route), with no
// dependency on the main app ever having been mounted - it always starts at
// a neutral 1x1 MG9 default and works from a bookmarked/typed URL alone.
export default function QuickLayoutView() {
  const [panelType, setPanelType] = useState<PanelTypeKey>("MG9");
  const [cols, setCols] = useState(1);
  const [rows, setRows] = useState(1);

  const panel = PANEL_TYPES[panelType];
  const wallWidthM = cols * panel.w;
  const wallHeightM = rows * panel.h;
  const pixelW = cols * panel.pixW;
  const pixelH = rows * panel.pixH;
  const totalPanels = cols * rows;
  const totalPixels = pixelW * pixelH;

  const ratioDivisor = gcd(pixelW, pixelH) || 1;
  const ratioLabel = `${pixelW / ratioDivisor}:${pixelH / ratioDivisor}`;

  // Largest 16:9 rect centred inside the wall's pixel resolution.
  const fitsWide = pixelW / pixelH > 16 / 9;
  const fitW = fitsWide ? (pixelH * 16) / 9 : pixelW;
  const fitH = fitsWide ? pixelH : (pixelW * 9) / 16;
  const fitOffsetX = (pixelW - fitW) / 2;
  const fitOffsetY = (pixelH - fitH) / 2;

  const wallBelowFullHd = pixelW < 1920 || pixelH < 1080;
  const contentBelowFullHd = fitW < 1920 || fitH < 1080;

  const sendToMainTool = () => {
    localStorage.setItem(QUICK_LAYOUT_TRANSFER_KEY, JSON.stringify({ panelType, cols, rows }));
    window.location.href = window.location.pathname;
  };

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
          <Button intent="primary" onClick={sendToMainTool}>
            Send to Main Layout Tool
          </Button>
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
                  ⚠ The centred 16:9 content area is below 1920×1080 (Full HD).
                </div>
              ) : null}
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle>Preview</CardTitle>
            </CardHeader>
            <CardContent>
              <div
                className="relative mx-auto w-full overflow-hidden rounded-lg border border-slate-600 bg-slate-950"
                style={{
                  maxWidth: 640,
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
              <div className="mt-2 text-center text-xs text-slate-400">Dashed box = centred 16:9 content area</div>
            </CardContent>
          </Card>
        </div>
      </div>
    </div>
  );
}
