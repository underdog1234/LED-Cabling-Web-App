import React, { useMemo, useState } from "react";
import { type RectMm, activeBBox } from "../model/panels";
import { Button, Card, CardContent, CardHeader, CardTitle, Input, Select } from "../components/ui";
import { OUTPUT_CANVAS_PRESETS, type Cell, type SubScreen, PORT_COLORS, cellRect } from "../App";
import { subScreenBBoxOf } from "../subScreens/subScreenModel";
import { resolutionOf, subScreenResolutionOf } from "./canvasModel";

type ScreenEntry = {
  /** null = the whole layout (used only when no sub-screens exist). */
  id: string | null;
  name: string;
  canvasX: number;
  canvasY: number;
  bboxMm: RectMm;
  resolution: { w: number; h: number };
};

type Props = {
  outputCanvasW: number;
  outputCanvasH: number;
  onResolutionChange: (w: number, h: number) => void;
  subScreens: SubScreen[];
  grid: Cell[];
  wholeLayoutCanvasX: number;
  wholeLayoutCanvasY: number;
  snapEnabled: boolean;
  onToggleSnap: () => void;
  /** id === null updates the whole-layout position. */
  onPositionChange: (id: string | null, x: number, y: number) => void;
};

const SNAP_PX = 10;
const PREVIEW_MAX_W = 820;
const PREVIEW_MAX_H = 460;

const rectsOverlap = (a: RectMm, b: RectMm) => a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;

export default function OutputCanvasPanel({
  outputCanvasW,
  outputCanvasH,
  onResolutionChange,
  subScreens,
  grid,
  wholeLayoutCanvasX,
  wholeLayoutCanvasY,
  snapEnabled,
  onToggleSnap,
  onPositionChange,
}: Props) {
  const [customW, setCustomW] = useState(String(outputCanvasW));
  const [customH, setCustomH] = useState(String(outputCanvasH));
  const [selectedKey, setSelectedKey] = useState<string | null>(null);
  const [drag, setDrag] = useState<{ key: string; id: string | null; startClientX: number; startClientY: number; startX: number; startY: number } | null>(null);

  const entries: ScreenEntry[] = useMemo(() => {
    if (subScreens.length) {
      return subScreens.map((screen) => ({
        id: screen.id,
        name: screen.name,
        canvasX: screen.canvasX,
        canvasY: screen.canvasY,
        bboxMm: subScreenBBoxOf(grid, screen.id, cellRect),
        resolution: subScreenResolutionOf(grid, screen.id),
      }));
    }
    const activePanels = grid.filter((cell) => !cell.isRemoved);
    return [
      {
        id: null,
        name: "Whole Layout",
        canvasX: wholeLayoutCanvasX,
        canvasY: wholeLayoutCanvasY,
        bboxMm: activeBBox(activePanels.map(cellRect)),
        resolution: resolutionOf(activePanels),
      },
    ];
  }, [subScreens, grid, wholeLayoutCanvasX, wholeLayoutCanvasY]);

  const keyOf = (id: string | null) => id ?? "__whole__";

  const scale = Math.min(PREVIEW_MAX_W / Math.max(outputCanvasW, 1), PREVIEW_MAX_H / Math.max(outputCanvasH, 1), 1);
  const previewW = Math.max(1, Math.round(outputCanvasW * scale));
  const previewH = Math.max(1, Math.round(outputCanvasH * scale));

  const overlaps = useMemo(() => {
    const pairs: [string, string][] = [];
    for (let i = 0; i < entries.length; i += 1) {
      for (let j = i + 1; j < entries.length; j += 1) {
        const a = entries[i];
        const b = entries[j];
        if (a.resolution.w <= 0 || a.resolution.h <= 0 || b.resolution.w <= 0 || b.resolution.h <= 0) continue;
        const rectA = { x: a.canvasX, y: a.canvasY, w: a.resolution.w, h: a.resolution.h };
        const rectB = { x: b.canvasX, y: b.canvasY, w: b.resolution.w, h: b.resolution.h };
        if (rectsOverlap(rectA, rectB)) pairs.push([a.name, b.name]);
      }
    }
    return pairs;
  }, [entries]);

  const warnings: string[] = [];
  entries.forEach((entry) => {
    if (entry.canvasX < 0 || entry.canvasY < 0) {
      warnings.push(`${entry.name}: negative canvas position (X ${entry.canvasX}, Y ${entry.canvasY}).`);
    }
    if (entry.canvasX + entry.resolution.w > outputCanvasW || entry.canvasY + entry.resolution.h > outputCanvasH) {
      warnings.push(`${entry.name}: extends beyond the ${outputCanvasW}×${outputCanvasH} output canvas.`);
    }
  });
  overlaps.forEach(([a, b]) => warnings.push(`${a} and ${b} overlap on the output canvas.`));
  const layoutBBoxAll = activeBBox(grid.filter((c) => !c.isRemoved).map(cellRect));
  const layoutRes = resolutionOf(grid);
  if (entries.length > 1) {
    const minX = Math.min(...entries.map((e) => e.canvasX));
    const minY = Math.min(...entries.map((e) => e.canvasY));
    const maxX = Math.max(...entries.map((e) => e.canvasX + e.resolution.w));
    const maxY = Math.max(...entries.map((e) => e.canvasY + e.resolution.h));
    if (maxX - minX > outputCanvasW || maxY - minY > outputCanvasH) {
      warnings.push("The complete layout exceeds the configured canvas resolution.");
    }
  }

  const applyPreset = (value: string) => {
    const preset = OUTPUT_CANVAS_PRESETS.find((p) => `${p.w}x${p.h}` === value);
    if (preset) {
      onResolutionChange(preset.w, preset.h);
      setCustomW(String(preset.w));
      setCustomH(String(preset.h));
    }
  };

  const applyCustomResolution = () => {
    const w = Math.max(1, Number.parseInt(customW, 10) || outputCanvasW);
    const h = Math.max(1, Number.parseInt(customH, 10) || outputCanvasH);
    onResolutionChange(w, h);
  };

  const snapValue = (value: number, edges: number[]) => {
    if (!snapEnabled) return value;
    const hit = edges.find((edge) => Math.abs(edge - value) <= SNAP_PX / scale);
    return hit !== undefined ? hit : value;
  };

  const moveEntry = (entry: ScreenEntry, x: number, y: number) => {
    // Snap to canvas edges (0, and outputCanvas - own size) and to other
    // entries' edges - disabled entirely via the snap toggle for precise
    // placement.
    const otherEdgesX = [0, outputCanvasW - entry.resolution.w];
    const otherEdgesY = [0, outputCanvasH - entry.resolution.h];
    entries
      .filter((e) => e.id !== entry.id)
      .forEach((e) => {
        otherEdgesX.push(e.canvasX, e.canvasX + e.resolution.w - entry.resolution.w, e.canvasX + e.resolution.w);
        otherEdgesY.push(e.canvasY, e.canvasY + e.resolution.h - entry.resolution.h, e.canvasY + e.resolution.h);
      });
    const snappedX = Math.round(snapValue(x, otherEdgesX));
    const snappedY = Math.round(snapValue(y, otherEdgesY));
    onPositionChange(entry.id, snappedX, snappedY);
  };

  const onPreviewMouseDown = (entry: ScreenEntry, event: React.MouseEvent) => {
    event.stopPropagation();
    setSelectedKey(keyOf(entry.id));
    setDrag({
      key: keyOf(entry.id),
      id: entry.id,
      startClientX: event.clientX,
      startClientY: event.clientY,
      startX: entry.canvasX,
      startY: entry.canvasY,
    });
  };

  const onPreviewMouseMove = (event: React.MouseEvent) => {
    if (!drag) return;
    const entry = entries.find((e) => keyOf(e.id) === drag.key);
    if (!entry) return;
    const dxCanvas = (event.clientX - drag.startClientX) / scale;
    const dyCanvas = (event.clientY - drag.startClientY) / scale;
    moveEntry(entry, drag.startX + dxCanvas, drag.startY + dyCanvas);
  };

  const onPreviewMouseUp = () => setDrag(null);

  const onPreviewKeyDown = (event: React.KeyboardEvent) => {
    if (!selectedKey) return;
    const entry = entries.find((e) => keyOf(e.id) === selectedKey);
    if (!entry) return;
    const step = event.shiftKey ? 10 : 1;
    if (event.key === "ArrowLeft") { moveEntry(entry, entry.canvasX - step, entry.canvasY); event.preventDefault(); }
    else if (event.key === "ArrowRight") { moveEntry(entry, entry.canvasX + step, entry.canvasY); event.preventDefault(); }
    else if (event.key === "ArrowUp") { moveEntry(entry, entry.canvasX, entry.canvasY - step); event.preventDefault(); }
    else if (event.key === "ArrowDown") { moveEntry(entry, entry.canvasX, entry.canvasY + step); event.preventDefault(); }
  };

  return (
    <Card className="border-slate-700 bg-slate-800 print-card no-print" collapsible>
      <CardHeader>
        <CardTitle className="text-white [text-shadow:0_0_2px_black]">Output Canvas</CardTitle>
      </CardHeader>
      <CardContent className="space-y-4 text-sm text-white [text-shadow:0_0_2px_black]">
        <div className="flex flex-wrap items-end gap-2">
          <div className="space-y-1">
            <label className="text-xs text-slate-300">Preset</label>
            <Select onChange={(e: React.ChangeEvent<HTMLSelectElement>) => applyPreset(e.target.value)} defaultValue="">
              <option value="">Custom...</option>
              {OUTPUT_CANVAS_PRESETS.map((p) => (
                <option key={`${p.w}x${p.h}`} value={`${p.w}x${p.h}`}>
                  {p.w} × {p.h}
                </option>
              ))}
            </Select>
          </div>
          <div className="space-y-1">
            <label className="text-xs text-slate-300">Width (px)</label>
            <Input className="bg-white text-black" type="number" min="1" value={customW} onChange={(e) => setCustomW(e.target.value)} onBlur={applyCustomResolution} />
          </div>
          <div className="space-y-1">
            <label className="text-xs text-slate-300">Height (px)</label>
            <Input className="bg-white text-black" type="number" min="1" value={customH} onChange={(e) => setCustomH(e.target.value)} onBlur={applyCustomResolution} />
          </div>
          <Button intent="secondary" size="sm" onClick={applyCustomResolution}>
            Apply
          </Button>
          <label className="flex items-center gap-2 rounded-lg border border-slate-600 bg-slate-900 px-3 py-2">
            <input type="checkbox" checked={snapEnabled} onChange={onToggleSnap} />
            <span>Snap to edges / other screens</span>
          </label>
        </div>

        <div className="space-y-2">
          {entries.map((entry, index) => {
            const color = PORT_COLORS[index % PORT_COLORS.length];
            return (
              <div key={keyOf(entry.id)} className="flex flex-wrap items-center gap-2 rounded-lg border border-slate-700 bg-slate-900/50 p-2">
                <span className="w-32 shrink-0 truncate font-semibold" style={{ color }}>
                  {entry.name}
                </span>
                <span className="text-xs text-slate-400 w-28 shrink-0">
                  {entry.resolution.w}×{entry.resolution.h}px
                </span>
                <span className="text-xs text-slate-400 w-28 shrink-0">
                  {(entry.bboxMm.w / 1000).toFixed(2)}m × {(entry.bboxMm.h / 1000).toFixed(2)}m
                </span>
                <label className="flex items-center gap-1 text-xs text-slate-300">
                  X
                  <Input
                    className="w-20 bg-white text-black"
                    type="number"
                    value={entry.canvasX}
                    onChange={(e) => onPositionChange(entry.id, Number(e.target.value) || 0, entry.canvasY)}
                  />
                </label>
                <label className="flex items-center gap-1 text-xs text-slate-300">
                  Y
                  <Input
                    className="w-20 bg-white text-black"
                    type="number"
                    value={entry.canvasY}
                    onChange={(e) => onPositionChange(entry.id, entry.canvasX, Number(e.target.value) || 0)}
                  />
                </label>
                <span className="text-xs text-slate-500">
                  right {entry.canvasX + entry.resolution.w}, bottom {entry.canvasY + entry.resolution.h}
                </span>
                <div className="flex flex-wrap gap-1">
                  <Button intent="ghost" size="sm" onClick={() => onPositionChange(entry.id, 0, entry.canvasY)}>⇤ Left</Button>
                  <Button intent="ghost" size="sm" onClick={() => onPositionChange(entry.id, outputCanvasW - entry.resolution.w, entry.canvasY)}>Right ⇥</Button>
                  <Button intent="ghost" size="sm" onClick={() => onPositionChange(entry.id, entry.canvasX, 0)}>⇡ Top</Button>
                  <Button intent="ghost" size="sm" onClick={() => onPositionChange(entry.id, entry.canvasX, outputCanvasH - entry.resolution.h)}>Bottom ⇣</Button>
                  <Button intent="ghost" size="sm" onClick={() => onPositionChange(entry.id, Math.round((outputCanvasW - entry.resolution.w) / 2), entry.canvasY)}>Centre H</Button>
                  <Button intent="ghost" size="sm" onClick={() => onPositionChange(entry.id, entry.canvasX, Math.round((outputCanvasH - entry.resolution.h) / 2))}>Centre V</Button>
                </div>
              </div>
            );
          })}
        </div>

        {warnings.length ? (
          <div className="rounded-lg border border-amber-400 bg-amber-500/15 p-2 text-xs text-amber-200">
            {warnings.map((w, i) => (
              <div key={i}>⚠ {w}</div>
            ))}
          </div>
        ) : null}

        <div className="text-xs text-slate-400">
          Complete layout resolution: {layoutRes.w}×{layoutRes.h}px · {(layoutBBoxAll.w / 1000).toFixed(2)}m × {(layoutBBoxAll.h / 1000).toFixed(2)}m
        </div>

        <div
          className="relative overflow-hidden rounded-lg border-2 border-dashed border-slate-500 bg-slate-950 outline-none"
          style={{ width: previewW, height: previewH }}
          tabIndex={0}
          onMouseMove={onPreviewMouseMove}
          onMouseUp={onPreviewMouseUp}
          onMouseLeave={onPreviewMouseUp}
          onKeyDown={onPreviewKeyDown}
          onMouseDown={() => setSelectedKey(null)}
        >
          {entries.map((entry, index) => {
            const color = PORT_COLORS[index % PORT_COLORS.length];
            const outOfBounds =
              entry.canvasX < 0 ||
              entry.canvasY < 0 ||
              entry.canvasX + entry.resolution.w > outputCanvasW ||
              entry.canvasY + entry.resolution.h > outputCanvasH;
            const isSelected = selectedKey === keyOf(entry.id);
            return (
              <div
                key={keyOf(entry.id)}
                onMouseDown={(event) => onPreviewMouseDown(entry, event)}
                className="absolute flex cursor-move select-none items-start justify-start overflow-hidden text-[10px] font-semibold"
                style={{
                  left: entry.canvasX * scale,
                  top: entry.canvasY * scale,
                  width: Math.max(2, entry.resolution.w * scale),
                  height: Math.max(2, entry.resolution.h * scale),
                  background: `${color}33`,
                  border: `2px ${outOfBounds ? "dashed #f87171" : isSelected ? "solid #ffffff" : "solid " + color}`,
                  color: "#e2e8f0",
                }}
              >
                <span className="m-1 rounded bg-black/60 px-1 py-0.5">
                  {entry.name} · {entry.resolution.w}×{entry.resolution.h} · X{entry.canvasX} Y{entry.canvasY}
                </span>
              </div>
            );
          })}
        </div>
      </CardContent>
    </Card>
  );
}
