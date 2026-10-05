// The main canvas, live, with every sub-screen selectable, draggable and
// resizable on top of it. The picture underneath is the real render (the
// same one the outputs and exports use), scaled to fit; the handles, guides
// and outlines are DOM on top of it and never reach an output.
import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  boundsOf,
  resizeRect,
  screenRect,
  snapEdge,
  snapMove,
  snapTargetsFor,
  type GeneratorConfig,
  type Rect,
  type ResizeHandle,
  type SnapGuide,
} from "./model";
import { renderComposite } from "./render";

type Props = {
  config: GeneratorConfig;
  /** Read every frame: what to draw and at what point in the loop. */
  frameSource: React.MutableRefObject<() => { config: GeneratorConfig; time: number }>;
  selectedIds: string[];
  onSelect: (ids: string[]) => void;
  /** Called once when a drag starts, so the whole drag is one undo step. */
  onBeginEdit: () => void;
  onRects: (rects: Record<string, Rect>) => void;
  maxHeight: number;
};

type Drag =
  | { kind: "move"; startX: number; startY: number; ids: string[]; start: Record<string, Rect>; moved: boolean }
  | { kind: "resize"; startX: number; startY: number; id: string; handle: ResizeHandle; start: Rect }
  | { kind: "marquee"; startX: number; startY: number; base: string[]; x0: number; y0: number; x1: number; y1: number };

const HANDLES: Array<{ h: ResizeHandle; x: number; y: number; cursor: string }> = [
  { h: "nw", x: 0, y: 0, cursor: "nwse-resize" },
  { h: "n", x: 0.5, y: 0, cursor: "ns-resize" },
  { h: "ne", x: 1, y: 0, cursor: "nesw-resize" },
  { h: "e", x: 1, y: 0.5, cursor: "ew-resize" },
  { h: "se", x: 1, y: 1, cursor: "nwse-resize" },
  { h: "s", x: 0.5, y: 1, cursor: "ns-resize" },
  { h: "sw", x: 0, y: 1, cursor: "nesw-resize" },
  { h: "w", x: 0, y: 0.5, cursor: "ew-resize" },
];

export default function CanvasEditor({ config, frameSource, selectedIds, onSelect, onBeginEdit, onRects, maxHeight }: Props) {
  const wrapRef = useRef<HTMLDivElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [boxW, setBoxW] = useState(800);
  const [guides, setGuides] = useState<SnapGuide[]>([]);
  const [marquee, setMarquee] = useState<Rect | null>(null);
  const [hud, setHud] = useState<string | null>(null);
  const dragRef = useRef<Drag | null>(null);
  const { w: W, h: H } = config.canvas;

  useEffect(() => {
    const el = wrapRef.current;
    if (!el) return;
    const measure = () => setBoxW(Math.max(200, el.clientWidth));
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  const scale = Math.min(boxW / Math.max(1, W), maxHeight / Math.max(1, H));
  const previewW = Math.max(1, Math.round(W * scale));
  const previewH = Math.max(1, Math.round(H * scale));

  // Live render loop. Skips frames that would be identical (paused, nothing
  // changed) so an idle editor costs nothing.
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.max(1, Math.round(previewW * dpr));
    canvas.height = Math.max(1, Math.round(previewH * dpr));
    const ctx = canvas.getContext("2d")!;
    const cache = new Map<string, HTMLCanvasElement>();
    let raf = 0;
    let last = 0;
    let lastConfig: GeneratorConfig | null = null;
    let lastTime = -1;
    const tick = (ts: number) => {
      raf = requestAnimationFrame(tick);
      if (ts - last < 1000 / 30) return;
      last = ts;
      const { config: c, time } = frameSource.current();
      if (c === lastConfig && time === lastTime) return;
      lastConfig = c;
      lastTime = time;
      renderComposite(ctx, c, time, cache, scale * dpr);
    };
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
  }, [previewW, previewH, scale, frameSource]);

  const visible = useMemo(() => config.screens.filter((s) => s.visible), [config.screens]);
  const selected = new Set(selectedIds);
  const single = selectedIds.length === 1 ? config.screens.find((s) => s.id === selectedIds[0]) : undefined;

  const toCanvas = (clientX: number, clientY: number) => {
    const r = wrapRef.current!.querySelector("[data-stage]")!.getBoundingClientRect();
    return { x: (clientX - r.left) / scale, y: (clientY - r.top) / scale };
  };

  const thresholdPx = config.snap.threshold / scale;
  const grid = config.snap.grid ? config.snap.gridSize : null;

  const onScreenDown = (e: React.PointerEvent, id: string) => {
    if (e.button !== 0) return;
    e.stopPropagation();
    (e.currentTarget as HTMLElement).setPointerCapture?.(e.pointerId);
    let ids = selectedIds;
    if (e.shiftKey || e.ctrlKey || e.metaKey) {
      ids = selected.has(id) ? selectedIds.filter((x) => x !== id) : [...selectedIds, id];
      onSelect(ids);
      if (!ids.includes(id)) return;
    } else if (!selected.has(id)) {
      ids = [id];
      onSelect(ids);
    }
    const movable = config.screens.filter((s) => ids.includes(s.id) && !s.locked);
    if (!movable.length) return;
    dragRef.current = {
      kind: "move",
      startX: e.clientX,
      startY: e.clientY,
      ids: movable.map((s) => s.id),
      start: Object.fromEntries(movable.map((s) => [s.id, screenRect(s)])),
      moved: false,
    };
  };

  const onHandleDown = (e: React.PointerEvent, handle: ResizeHandle) => {
    if (!single || single.locked) return;
    e.stopPropagation();
    (e.currentTarget as HTMLElement).setPointerCapture?.(e.pointerId);
    onBeginEdit();
    dragRef.current = { kind: "resize", startX: e.clientX, startY: e.clientY, id: single.id, handle, start: screenRect(single) };
  };

  const onStageDown = (e: React.PointerEvent) => {
    if (e.button !== 0) return;
    (e.currentTarget as HTMLElement).setPointerCapture?.(e.pointerId);
    const p = toCanvas(e.clientX, e.clientY);
    const base = e.shiftKey || e.ctrlKey || e.metaKey ? selectedIds : [];
    if (!base.length) onSelect([]);
    dragRef.current = { kind: "marquee", startX: e.clientX, startY: e.clientY, base, x0: p.x, y0: p.y, x1: p.x, y1: p.y };
  };

  const onMove = (e: React.PointerEvent) => {
    const d = dragRef.current;
    if (!d) return;
    const dx = (e.clientX - d.startX) / scale;
    const dy = (e.clientY - d.startY) / scale;
    const snapOn = config.snap.enabled && !e.altKey;
    if (d.kind === "move") {
      if (!d.moved) {
        if (Math.abs(e.clientX - d.startX) + Math.abs(e.clientY - d.startY) < 3) return;
        d.moved = true;
        onBeginEdit();
      }
      const startBounds = boundsOf(Object.values(d.start));
      const moved = { ...startBounds, x: startBounds.x + dx, y: startBounds.y + dy };
      let ox = Math.round(dx);
      let oy = Math.round(dy);
      let g: SnapGuide[] = [];
      if (snapOn) {
        const others = visible.filter((s) => !d.ids.includes(s.id)).map(screenRect);
        const res = snapMove(moved, snapTargetsFor(config.canvas, others, config.snap), thresholdPx, grid);
        ox = res.x - startBounds.x;
        oy = res.y - startBounds.y;
        g = res.guides;
      }
      setGuides(g);
      const next: Record<string, Rect> = {};
      d.ids.forEach((id) => {
        const s = d.start[id];
        next[id] = { ...s, x: s.x + ox, y: s.y + oy };
      });
      onRects(next);
      const b = boundsOf(Object.values(next));
      setHud(`X ${b.x}  Y ${b.y}${d.ids.length > 1 ? `  ·  ${d.ids.length} screens` : ""}`);
    } else if (d.kind === "resize") {
      const screen = config.screens.find((s) => s.id === d.id);
      const lock = (screen?.aspectLock ?? false) || e.shiftKey;
      let r = resizeRect(d.start, d.handle, dx, dy, lock);
      const g: SnapGuide[] = [];
      if (snapOn && !lock) {
        const others = visible.filter((s) => s.id !== d.id).map(screenRect);
        const t = snapTargetsFor(config.canvas, others, config.snap);
        if (d.handle.includes("e")) {
          const s = snapEdge(r.x + r.w, t.xs, thresholdPx, grid);
          r = { ...r, w: Math.max(1, s.value - r.x) };
          if (s.guide !== null) g.push({ axis: "x", at: s.guide, from: 0, to: 0 });
        }
        if (d.handle.includes("w")) {
          const right = r.x + r.w;
          const s = snapEdge(r.x, t.xs, thresholdPx, grid);
          const x = Math.min(right - 1, s.value);
          r = { ...r, x, w: right - x };
          if (s.guide !== null) g.push({ axis: "x", at: s.guide, from: 0, to: 0 });
        }
        if (d.handle.includes("s")) {
          const s = snapEdge(r.y + r.h, t.ys, thresholdPx, grid);
          r = { ...r, h: Math.max(1, s.value - r.y) };
          if (s.guide !== null) g.push({ axis: "y", at: s.guide, from: 0, to: 0 });
        }
        if (d.handle.includes("n")) {
          const bottom = r.y + r.h;
          const s = snapEdge(r.y, t.ys, thresholdPx, grid);
          const y = Math.min(bottom - 1, s.value);
          r = { ...r, y, h: bottom - y };
          if (s.guide !== null) g.push({ axis: "y", at: s.guide, from: 0, to: 0 });
        }
      }
      setGuides(g);
      onRects({ [d.id]: r });
      setHud(`${r.w} × ${r.h}  ·  X ${r.x}  Y ${r.y}`);
    } else {
      const p = toCanvas(e.clientX, e.clientY);
      d.x1 = p.x;
      d.y1 = p.y;
      const m = { x: Math.min(d.x0, d.x1), y: Math.min(d.y0, d.y1), w: Math.abs(d.x1 - d.x0), h: Math.abs(d.y1 - d.y0) };
      setMarquee(m);
      const hit = visible.filter((s) => s.x < m.x + m.w && m.x < s.x + s.w && s.y < m.y + m.h && m.y < s.y + s.h).map((s) => s.id);
      onSelect([...new Set([...d.base, ...hit])]);
    }
  };

  const onUp = () => {
    dragRef.current = null;
    setGuides([]);
    setMarquee(null);
    setHud(null);
  };

  const outOfBounds = (s: Rect) => s.x < 0 || s.y < 0 || s.x + s.w > W || s.y + s.h > H;

  return (
    <div ref={wrapRef} className="w-full select-none">
      <div className="relative mx-auto" style={{ width: previewW, height: previewH }}>
        <canvas ref={canvasRef} className="absolute inset-0 block" style={{ width: previewW, height: previewH }} />
        <div
          data-stage
          className="absolute inset-0 cursor-crosshair outline outline-1 outline-slate-500"
          onPointerDown={onStageDown}
          onPointerMove={onMove}
          onPointerUp={onUp}
          onPointerCancel={onUp}
        >
          {config.snap.enabled && config.snap.grid && config.snap.gridSize * scale >= 6 ? (
            <div
              className="pointer-events-none absolute inset-0"
              style={{
                backgroundImage: "linear-gradient(to right, rgba(148,163,184,0.18) 1px, transparent 1px), linear-gradient(to bottom, rgba(148,163,184,0.18) 1px, transparent 1px)",
                backgroundSize: `${config.snap.gridSize * scale}px ${config.snap.gridSize * scale}px`,
              }}
            />
          ) : null}
          {visible.map((s) => {
            const isSel = selected.has(s.id);
            const oob = outOfBounds(s);
            return (
              <div
                key={s.id}
                data-screen-id={s.id}
                onPointerDown={(e) => onScreenDown(e, s.id)}
                className={`absolute ${s.locked ? "cursor-not-allowed" : "cursor-move"}`}
                style={{
                  left: s.x * scale,
                  top: s.y * scale,
                  width: Math.max(2, s.w * scale),
                  height: Math.max(2, s.h * scale),
                  outline: isSel ? "2px solid #ffffff" : `1px ${oob ? "dashed #f87171" : s.locked ? "dashed " + s.color : "solid " + s.color}`,
                  outlineOffset: isSel ? 0 : -1,
                  boxShadow: isSel ? "0 0 0 3px rgba(56,189,248,0.6)" : undefined,
                }}
                title={`${s.name} · ${s.w}×${s.h} · X ${s.x} Y ${s.y}${s.locked ? " · locked" : ""}`}
              >
                <span className="pointer-events-none absolute left-0 top-0 max-w-full truncate rounded-br bg-black/70 px-1 text-[10px] font-semibold leading-4 text-white">
                  {s.locked ? "🔒 " : ""}
                  {s.name}
                </span>
              </div>
            );
          })}
          {single && single.visible && !single.locked
            ? HANDLES.map(({ h, x, y, cursor }) => (
                <div
                  key={h}
                  data-handle={h}
                  onPointerDown={(e) => onHandleDown(e, h)}
                  className="absolute h-3 w-3 rounded-sm border border-slate-900 bg-white"
                  style={{ left: (single.x + single.w * x) * scale - 6, top: (single.y + single.h * y) * scale - 6, cursor }}
                />
              ))
            : null}
          {guides.map((g, i) =>
            g.axis === "x" ? (
              <div key={i} className="pointer-events-none absolute top-0 h-full w-px bg-fuchsia-400" style={{ left: g.at * scale }} />
            ) : (
              <div key={i} className="pointer-events-none absolute left-0 h-px w-full bg-fuchsia-400" style={{ top: g.at * scale }} />
            ),
          )}
          {marquee ? (
            <div
              className="pointer-events-none absolute border border-sky-300 bg-sky-400/10"
              style={{ left: marquee.x * scale, top: marquee.y * scale, width: marquee.w * scale, height: marquee.h * scale }}
            />
          ) : null}
          {hud ? <div className="pointer-events-none absolute bottom-1 right-1 rounded bg-black/80 px-2 py-0.5 font-mono text-[11px] text-white">{hud}</div> : null}
        </div>
      </div>
      <div className="mt-1 text-center text-[11px] text-slate-500">
        Preview at {(scale * 100).toFixed(scale < 0.1 ? 1 : 0)}% · drag to move · handles resize (Shift keeps the ratio) · Alt while dragging skips snapping · arrow keys nudge
      </div>
    </div>
  );
}
