// Side panels of the generator: the sub-screen list, the inspector for the
// selected screen, arrange/align tools, canvas settings and the playlist.
import React, { useState } from "react";
import { PROCESSOR_MODEL_IDS, PROCESSOR_SPECS } from "../novastar/processorModels";
import { Check, ColorField, Label, NumField, Panel, SelectField, SmallButton, TextField } from "./controls";
import {
  LED_PANEL_PRESETS,
  MOTION_DIRECTIONS,
  RESOLUTION_PRESETS,
  aspectLabel,
  clampDimension,
  resizeWithAspect,
  type AlignMode,
  type ArrangeMode,
  type GeneratorConfig,
  type LayerMove,
  type MotionDirection,
  type PatternSettings,
  type PlaylistStep,
  type ResolutionPreset,
  type SubScreen,
} from "./model";
import { LED_LAYOUT_PATTERN_ID, PATTERNS, PATTERN_BY_ID, patternName, resolveSettings, type ParamDef } from "./patterns";

// ---------------------------------------------------------------------------
// Resolution picker - shared by the canvas and every screen.
// ---------------------------------------------------------------------------

export function ResolutionEditor({
  w,
  h,
  aspectLock,
  onSize,
  onAspectLock,
  customPresets,
  onSavePreset,
}: {
  w: number;
  h: number;
  aspectLock: boolean;
  onSize: (w: number, h: number) => void;
  onAspectLock: (v: boolean) => void;
  customPresets: ResolutionPreset[];
  onSavePreset: (p: ResolutionPreset) => void;
}) {
  const [presetName, setPresetName] = useState("");
  const all = [...RESOLUTION_PRESETS, ...customPresets];
  const match = all.find((p) => p.w === w && p.h === h);
  return (
    <div className="space-y-2">
      <SelectField
        label="Preset"
        value={match ? `${match.w}x${match.h}:${match.name}` : ""}
        onChange={(v) => {
          const p = all.find((x) => `${x.w}x${x.h}:${x.name}` === v);
          if (p) onSize(p.w, p.h);
        }}
      >
        <option value="">Custom - {w} × {h}</option>
        <optgroup label="Common">
          {RESOLUTION_PRESETS.map((p) => (
            <option key={p.name} value={`${p.w}x${p.h}:${p.name}`}>{p.name}</option>
          ))}
        </optgroup>
        {customPresets.length ? (
          <optgroup label="Saved presets">
            {customPresets.map((p) => (
              <option key={`c-${p.name}`} value={`${p.w}x${p.h}:${p.name}`}>{p.name} - {p.w} × {p.h}</option>
            ))}
          </optgroup>
        ) : null}
      </SelectField>
      <div className="grid grid-cols-[1fr_auto_1fr] items-end gap-2">
        <NumField label="Width (px)" value={w} min={1} max={16384} onCommit={(v) => { const r = resizeWithAspect({ w, h }, "w", v, aspectLock); onSize(r.w, r.h); }} />
        <SmallButton className="mb-0.5" title="Swap width and height (portrait / landscape)" onClick={() => onSize(h, w)}>⇄</SmallButton>
        <NumField label="Height (px)" value={h} min={1} max={16384} onCommit={(v) => { const r = resizeWithAspect({ w, h }, "h", v, aspectLock); onSize(r.w, r.h); }} />
      </div>
      <div className="flex flex-wrap items-center justify-between gap-2">
        <Check checked={aspectLock} onChange={onAspectLock}>Lock aspect ratio</Check>
        <span className="text-xs text-slate-400">{aspectLabel(w, h)} · {(w * h / 1e6).toFixed(2)} MP</span>
      </div>
      <div className="flex gap-2">
        <TextField value={presetName} onCommit={setPresetName} placeholder="Preset name" />
        <SmallButton
          disabled={!presetName.trim()}
          onClick={() => {
            onSavePreset({ name: presetName.trim(), w, h });
            setPresetName("");
          }}
        >
          Save preset
        </SmallButton>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Sub-screen list.
// ---------------------------------------------------------------------------

export function ScreenList({
  config,
  selectedIds,
  onSelect,
  onAdd,
  onDuplicate,
  onDelete,
  onPatch,
  onLayer,
}: {
  config: GeneratorConfig;
  selectedIds: string[];
  onSelect: (ids: string[]) => void;
  onAdd: () => void;
  onDuplicate: () => void;
  onDelete: () => void;
  onPatch: (id: string, patch: Partial<SubScreen>) => void;
  onLayer: (move: LayerMove) => void;
}) {
  const [renaming, setRenaming] = useState<string | null>(null);
  const sel = new Set(selectedIds);
  const ordered = [...config.screens].reverse(); // top layer first
  const click = (e: React.MouseEvent, id: string) => {
    if (e.shiftKey && selectedIds.length) {
      // Range in list order.
      const ids = ordered.map((s) => s.id);
      const a = ids.indexOf(selectedIds[selectedIds.length - 1]);
      const b = ids.indexOf(id);
      const range = ids.slice(Math.min(a, b), Math.max(a, b) + 1);
      onSelect([...new Set([...selectedIds, ...range])]);
    } else if (e.ctrlKey || e.metaKey) {
      onSelect(sel.has(id) ? selectedIds.filter((x) => x !== id) : [...selectedIds, id]);
    } else onSelect([id]);
  };
  const none = !selectedIds.length;
  return (
    <Panel
      title={`Sub-screens (${config.screens.length})`}
      actions={<SmallButton tone="primary" onClick={onAdd} title="Add a sub-screen">+ Add</SmallButton>}
    >
      <div className="flex flex-wrap gap-1">
        <SmallButton onClick={onDuplicate} disabled={none} title="Duplicate (Ctrl+D)">Duplicate</SmallButton>
        <SmallButton onClick={onDelete} disabled={none} tone="danger" title="Delete (Del)">Delete</SmallButton>
        <SmallButton onClick={() => onSelect(config.screens.map((s) => s.id))} title="Select all (Ctrl+A)">All</SmallButton>
        <SmallButton onClick={() => onSelect([])} disabled={none}>None</SmallButton>
      </div>
      <div className="flex flex-wrap items-center gap-1">
        <Label>Layer</Label>
        <SmallButton onClick={() => onLayer("front")} disabled={none} title="Bring to front">⤒ Front</SmallButton>
        <SmallButton onClick={() => onLayer("forward")} disabled={none} title="Bring forward">↑</SmallButton>
        <SmallButton onClick={() => onLayer("backward")} disabled={none} title="Send backward">↓</SmallButton>
        <SmallButton onClick={() => onLayer("back")} disabled={none} title="Send to back">⤓ Back</SmallButton>
      </div>
      <div className="max-h-80 space-y-1 overflow-auto pr-1" role="listbox" aria-multiselectable>
        {ordered.map((s) => (
          <div
            key={s.id}
            role="option"
            aria-selected={sel.has(s.id)}
            data-testid="screen-row"
            onClick={(e) => click(e, s.id)}
            onDoubleClick={() => setRenaming(s.id)}
            className={`flex cursor-pointer items-center gap-2 rounded-md border px-2 py-1.5 text-sm ${sel.has(s.id) ? "border-sky-400 bg-sky-500/15" : "border-slate-700 bg-slate-800/60 hover:bg-slate-800"} ${s.visible ? "" : "opacity-50"}`}
          >
            <span className="h-3 w-3 shrink-0 rounded-sm" style={{ background: s.color }} />
            <div className="min-w-0 flex-1">
              {renaming === s.id ? (
                <div onClick={(e) => e.stopPropagation()}>
                  <TextField
                    value={s.name}
                    onCommit={(v) => {
                      onPatch(s.id, { name: v.trim() || s.name });
                      setRenaming(null);
                    }}
                  />
                </div>
              ) : (
                <div className="truncate font-medium text-white">{s.name}</div>
              )}
              <div className="truncate text-[11px] text-slate-400">
                {s.w}×{s.h} · X{s.x} Y{s.y} · {patternName(s.pattern)}
              </div>
            </div>
            <button type="button" className="text-xs text-slate-300 hover:text-white" title={s.visible ? "Hide" : "Show"} onClick={(e) => { e.stopPropagation(); onPatch(s.id, { visible: !s.visible }); }}>
              {s.visible ? "👁" : "◌"}
            </button>
            <button type="button" className="text-xs text-slate-300 hover:text-white" title={s.locked ? "Unlock" : "Lock"} onClick={(e) => { e.stopPropagation(); onPatch(s.id, { locked: !s.locked }); }}>
              {s.locked ? "🔒" : "🔓"}
            </button>
          </div>
        ))}
        {!config.screens.length ? <div className="text-xs text-slate-400">No sub-screens yet. The canvas shows its background colour.</div> : null}
      </div>
      <div className="text-[11px] text-slate-500">Top of the list is the top layer. Click to select, Ctrl/⌘-click to add, Shift-click for a range, double-click to rename.</div>
    </Panel>
  );
}

// ---------------------------------------------------------------------------
// Pattern settings.
// ---------------------------------------------------------------------------

const ParamControl = ({ def, value, onChange }: { def: ParamDef; value: number | string | boolean; onChange: (v: number | string | boolean) => void }) => {
  if (def.type === "number") return <NumField label={def.label} value={Number(value)} min={def.min} max={def.max} step={def.step} onCommit={onChange} />;
  if (def.type === "color") return <div className="flex flex-col gap-1"><Label>{def.label}</Label><ColorField value={String(value)} onChange={onChange} /></div>;
  if (def.type === "boolean") return <Check checked={Boolean(value)} onChange={onChange}>{def.label}</Check>;
  return (
    <SelectField label={def.label} value={String(value)} onChange={onChange}>
      {def.options.map((o) => (
        <option key={o.value} value={o.value}>{o.label}</option>
      ))}
    </SelectField>
  );
};

export function PatternPicker({ value, onChange, allowLedLayout }: { value: string; onChange: (id: string) => void; allowLedLayout: boolean }) {
  const groups = [...new Set(PATTERNS.map((p) => p.category))];
  return (
    <SelectField label="Test pattern" value={value} onChange={onChange}>
      {groups.map((g) => (
        <optgroup key={g} label={g}>
          {PATTERNS.filter((p) => p.category === g && (p.id !== LED_LAYOUT_PATTERN_ID || allowLedLayout)).map((p) => (
            <option key={p.id} value={p.id}>{p.name}</option>
          ))}
        </optgroup>
      ))}
    </SelectField>
  );
}

// ---------------------------------------------------------------------------
// Inspector for the selected screen.
// ---------------------------------------------------------------------------

export function Inspector({
  config,
  screen,
  selectionCount,
  onPatch,
  onApplyPatternToSelection,
  onSavePreset,
}: {
  config: GeneratorConfig;
  screen: SubScreen;
  selectionCount: number;
  onPatch: (patch: Partial<SubScreen>) => void;
  onApplyPatternToSelection: () => void;
  onSavePreset: (p: ResolutionPreset) => void;
}) {
  const def = PATTERN_BY_ID.get(screen.pattern) ?? PATTERN_BY_ID.get("info")!;
  const settings = resolveSettings(def, screen.settings[def.id]);
  const setSetting = (key: string, v: number | string | boolean) =>
    onPatch({ settings: { ...screen.settings, [def.id]: { ...(screen.settings[def.id] ?? {}), [key]: v } as PatternSettings } });
  const W = config.canvas.w;
  const H = config.canvas.h;
  const place = (x: number, y: number) => onPatch({ x: Math.round(x), y: Math.round(y) });
  const led = screen.led;
  const [panelPreset, setPanelPreset] = useState("");
  const spec = config.processor.model ? PROCESSOR_SPECS[config.processor.model] : null;
  return (
    <Panel title={<span>Screen: <span style={{ color: screen.color }}>{screen.name}</span>{selectionCount > 1 ? <span className="text-slate-400"> (+{selectionCount - 1} more selected)</span> : null}</span>}>
      <div className="grid grid-cols-[1fr_auto] items-end gap-2">
        <TextField label="Name" value={screen.name} onCommit={(v) => onPatch({ name: v.trim() || screen.name })} />
        <ColorField value={screen.color} onChange={(v) => onPatch({ color: v })} />
      </div>
      <div className="flex flex-wrap gap-3">
        <Check checked={screen.visible} onChange={(v) => onPatch({ visible: v })}>Visible</Check>
        <Check checked={screen.locked} onChange={(v) => onPatch({ locked: v })}>Locked</Check>
      </div>

      <div className="rounded-lg border border-slate-700 p-2">
        <ResolutionEditor
          w={screen.w}
          h={screen.h}
          aspectLock={screen.aspectLock}
          onSize={(w, h) => !screen.locked && onPatch({ w, h })}
          onAspectLock={(v) => onPatch({ aspectLock: v })}
          customPresets={config.customPresets}
          onSavePreset={onSavePreset}
        />
      </div>

      <div className="rounded-lg border border-slate-700 p-2">
        <div className="grid grid-cols-2 gap-2">
          <NumField label="X (px)" value={screen.x} min={-16384} max={32768} disabled={screen.locked} onCommit={(v) => onPatch({ x: v })} />
          <NumField label="Y (px)" value={screen.y} min={-16384} max={32768} disabled={screen.locked} onCommit={(v) => onPatch({ y: v })} />
        </div>
        <div className="mt-1 text-[11px] text-slate-400">Right {screen.x + screen.w} · Bottom {screen.y + screen.h}</div>
        <div className="mt-2 flex flex-wrap gap-1">
          <SmallButton disabled={screen.locked} onClick={() => place(0, screen.y)}>⇤ Left</SmallButton>
          <SmallButton disabled={screen.locked} onClick={() => place((W - screen.w) / 2, screen.y)}>Centre H</SmallButton>
          <SmallButton disabled={screen.locked} onClick={() => place(W - screen.w, screen.y)}>Right ⇥</SmallButton>
          <SmallButton disabled={screen.locked} onClick={() => place(screen.x, 0)}>⇡ Top</SmallButton>
          <SmallButton disabled={screen.locked} onClick={() => place(screen.x, (H - screen.h) / 2)}>Centre V</SmallButton>
          <SmallButton disabled={screen.locked} onClick={() => place(screen.x, H - screen.h)}>Bottom ⇣</SmallButton>
          <SmallButton disabled={screen.locked} onClick={() => onPatch({ x: 0, y: 0, w: W, h: H })} title="Fill the whole canvas">Fill canvas</SmallButton>
        </div>
      </div>

      <div className="space-y-2 rounded-lg border border-slate-700 p-2">
        <PatternPicker value={screen.pattern} onChange={(id) => onPatch({ pattern: id })} allowLedLayout={!!screen.ledLayout} />
        {screen.pattern === LED_LAYOUT_PATTERN_ID ? (
          <div className="text-xs text-slate-400">
            The LED Cabling Planner&apos;s own Moving Test Pattern for the {screen.ledLayout?.panels.length ?? 0} panels imported with this screen
            {screen.ledLayout?.projectName ? ` from "${screen.ledLayout.projectName}"` : ""}.
          </div>
        ) : null}
        <div className="grid grid-cols-2 gap-2">
          {def.params.map((p) => (
            <div key={p.key} className={p.type === "boolean" ? "col-span-2" : ""}>
              <ParamControl def={p} value={settings[p.key]} onChange={(v) => setSetting(p.key, v)} />
            </div>
          ))}
        </div>
        <div className="grid grid-cols-2 gap-2">
          <SelectField label="Movement" value={screen.motion.direction} onChange={(v) => onPatch({ motion: { ...screen.motion, direction: v as MotionDirection } })}>
            {MOTION_DIRECTIONS.map((d) => (
              <option key={d.value} value={d.value}>{d.label}</option>
            ))}
          </SelectField>
          <NumField
            label="Passes per loop"
            value={screen.motion.passes}
            min={1}
            max={100}
            disabled={screen.motion.direction === "none"}
            title="How many times the pattern scrolls all the way across in one loop"
            onCommit={(v) => onPatch({ motion: { ...screen.motion, passes: v } })}
          />
        </div>
        {def.animated || screen.motion.direction !== "none" ? (
          <NumField label="Animation offset (seconds)" value={screen.phase} integer={false} step={0.1} min={-3600} max={3600} onCommit={(v) => onPatch({ phase: v })} />
        ) : null}
        {selectionCount > 1 ? <SmallButton onClick={onApplyPatternToSelection}>Use this pattern and settings on all {selectionCount} selected</SmallButton> : null}
      </div>

      <div className="rounded-lg border border-slate-700 p-2">
        <Label>Overlays (exported with the screen)</Label>
        <div className="mt-1 grid grid-cols-2 gap-1">
          <Check checked={screen.overlays.border} onChange={(v) => onPatch({ overlays: { ...screen.overlays, border: v } })}>Border</Check>
          <Check checked={screen.overlays.label} onChange={(v) => onPatch({ overlays: { ...screen.overlays, label: v } })}>Name label</Check>
          <Check checked={screen.overlays.resolution} onChange={(v) => onPatch({ overlays: { ...screen.overlays, resolution: v } })}>Resolution & position</Check>
          <Check checked={screen.overlays.crosshair} onChange={(v) => onPatch({ overlays: { ...screen.overlays, crosshair: v } })}>Centre crosshair</Check>
          <Check checked={screen.overlays.clock} onChange={(v) => onPatch({ overlays: { ...screen.overlays, clock: v } })} title="Analogue clock with a seconds hand and the time in digits, top right">Clock</Check>
        </div>
      </div>

      <div className="space-y-2 rounded-lg border border-slate-700 p-2">
        <Check checked={!!led} onChange={(v) => onPatch({ led: v ? { cols: Math.max(1, Math.round(screen.w / 168)), rows: Math.max(1, Math.round(screen.h / 168)), panelPxW: 168, panelPxH: 168, panelName: LED_PANEL_PRESETS[0].name } : null })}>
          LED panel grid (optional)
        </Check>
        {led ? (
          <>
            <SelectField
              label="Panel"
              value={panelPreset}
              onChange={(v) => {
                setPanelPreset(v);
                const p = LED_PANEL_PRESETS.find((x) => x.name === v);
                if (p) onPatch({ led: { ...led, panelPxW: p.pxW, panelPxH: p.pxH, panelName: p.name } });
              }}
            >
              <option value="">{led.panelName ?? "Custom"} - {led.panelPxW} × {led.panelPxH} px</option>
              {LED_PANEL_PRESETS.map((p) => (
                <option key={p.name} value={p.name}>{p.name} - {p.pxW} × {p.pxH} px</option>
              ))}
            </SelectField>
            <div className="grid grid-cols-2 gap-2">
              <NumField label="Columns" value={led.cols} min={1} max={500} onCommit={(v) => onPatch({ led: { ...led, cols: v } })} />
              <NumField label="Rows" value={led.rows} min={1} max={500} onCommit={(v) => onPatch({ led: { ...led, rows: v } })} />
              <NumField label="Panel px wide" value={led.panelPxW} min={1} max={4096} onCommit={(v) => onPatch({ led: { ...led, panelPxW: v, panelName: "Custom" } })} />
              <NumField label="Panel px high" value={led.panelPxH} min={1} max={4096} onCommit={(v) => onPatch({ led: { ...led, panelPxH: v, panelName: "Custom" } })} />
            </div>
            <div className="text-xs text-slate-400">
              {led.cols} × {led.rows} panels = {led.cols * led.panelPxW} × {led.rows * led.panelPxH} px
              {led.cols * led.panelPxW !== screen.w || led.rows * led.panelPxH !== screen.h ? " - not this screen's resolution yet" : " - matches this screen"}
            </div>
            <div className="flex flex-wrap gap-1">
              <SmallButton disabled={screen.locked} onClick={() => onPatch({ w: clampDimension(led.cols * led.panelPxW), h: clampDimension(led.rows * led.panelPxH) })}>Set resolution from panels</SmallButton>
              <SmallButton onClick={() => onPatch({ pattern: "led-moving" })}>Use LED Moving Pattern</SmallButton>
            </div>
          </>
        ) : null}
      </div>

      {spec && config.processor.inputMode === "perEntry" ? (
        <SelectField label={`${spec.label} input`} value={screen.input === null ? "" : String(screen.input)} onChange={(v) => onPatch({ input: v ? Number(v) : null })}>
          <option value="">Unassigned</option>
          {spec.inputs.map((i) => (
            <option key={i.interfacePk} value={i.interfacePk}>{i.label}</option>
          ))}
        </SelectField>
      ) : null}
    </Panel>
  );
}

// ---------------------------------------------------------------------------
// Align, distribute, arrange.
// ---------------------------------------------------------------------------

export type AlignRef = "canvas" | "selection" | string;

export function ArrangePanel({
  config,
  selectedIds,
  onAlign,
  onDistribute,
  onSpace,
  onArrange,
  onMatchSize,
  onNudge,
}: {
  config: GeneratorConfig;
  selectedIds: string[];
  onAlign: (mode: AlignMode, ref: AlignRef) => void;
  onDistribute: (axis: "horizontal" | "vertical") => void;
  onSpace: (axis: "horizontal" | "vertical", gap: number) => void;
  onArrange: (mode: ArrangeMode, opts: { columns: number; gapX: number; gapY: number; fromCanvasOrigin: boolean }) => void;
  onMatchSize: (refId: string, dim: "w" | "h" | "both") => void;
  onNudge: (step: number, shiftStep: number) => void;
}) {
  const [ref, setRef] = useState<AlignRef>("canvas");
  const [gap, setGap] = useState(0);
  const [columns, setColumns] = useState(2);
  const [gapX, setGapX] = useState(0);
  const [gapY, setGapY] = useState(0);
  const [fromOrigin, setFromOrigin] = useState(false);
  const selected = config.screens.filter((s) => selectedIds.includes(s.id));
  const n = selected.length;
  const refValid = ref === "canvas" || ref === "selection" || config.screens.some((s) => s.id === ref);
  const activeRef = refValid ? ref : "canvas";
  const refScreen = config.screens.find((s) => s.id === activeRef);
  const btn = (mode: AlignMode, label: string, title: string) => (
    <SmallButton key={mode} disabled={!n || (activeRef === "selection" && n < 2)} onClick={() => onAlign(mode, activeRef)} title={title}>{label}</SmallButton>
  );
  return (
    <Panel title={`Align & arrange${n ? ` (${n} selected)` : ""}`}>
      {!n ? <div className="text-xs text-slate-400">Select one or more sub-screens on the canvas or in the list.</div> : null}
      <SelectField label="Align relative to" value={activeRef} onChange={setRef}>
        <option value="canvas">Main canvas ({config.canvas.w} × {config.canvas.h})</option>
        <option value="selection">Selection bounds</option>
        <optgroup label="Reference sub-screen">
          {config.screens.map((s) => (
            <option key={s.id} value={s.id}>{s.name}</option>
          ))}
        </optgroup>
      </SelectField>
      <div className="grid grid-cols-3 gap-1">
        {btn("left", "⇤ Left", "Align left edges")}
        {btn("centre", "Centre", "Align horizontal centres")}
        {btn("right", "Right ⇥", "Align right edges")}
        {btn("top", "⇡ Top", "Align top edges")}
        {btn("middle", "Middle", "Align vertical centres")}
        {btn("bottom", "Bottom ⇣", "Align bottom edges")}
      </div>
      {refScreen ? (
        <div className="flex flex-wrap gap-1">
          <Label>Match {refScreen.name}&apos;s</Label>
          <SmallButton disabled={!n} onClick={() => onMatchSize(refScreen.id, "w")}>Width</SmallButton>
          <SmallButton disabled={!n} onClick={() => onMatchSize(refScreen.id, "h")}>Height</SmallButton>
          <SmallButton disabled={!n} onClick={() => onMatchSize(refScreen.id, "both")}>Size</SmallButton>
        </div>
      ) : null}
      <div className="space-y-1">
        <Label>Distribute evenly (3 or more)</Label>
        <div className="flex flex-wrap gap-1">
          <SmallButton disabled={n < 3} onClick={() => onDistribute("horizontal")}>↔ Horizontally</SmallButton>
          <SmallButton disabled={n < 3} onClick={() => onDistribute("vertical")}>↕ Vertically</SmallButton>
        </div>
      </div>
      <div className="space-y-1">
        <Label>Space with a fixed gap</Label>
        <div className="flex items-end gap-1">
          <div className="w-20"><NumField value={gap} min={-16384} max={16384} onCommit={setGap} title="Gap in pixels" /></div>
          <SmallButton disabled={n < 2} onClick={() => onSpace("horizontal", gap)}>↔</SmallButton>
          <SmallButton disabled={n < 2} onClick={() => onSpace("vertical", gap)}>↕</SmallButton>
          <span className="pb-1 text-[11px] text-slate-400">px</span>
        </div>
      </div>
      <div className="space-y-2 rounded-lg border border-slate-700 p-2">
        <Label>Arrange</Label>
        <div className="grid grid-cols-3 gap-2">
          <NumField label="Grid columns" value={columns} min={1} max={100} onCommit={setColumns} />
          <NumField label="Gap X (px)" value={gapX} min={-16384} max={16384} onCommit={setGapX} />
          <NumField label="Gap Y (px)" value={gapY} min={-16384} max={16384} onCommit={setGapY} />
        </div>
        <Check checked={fromOrigin} onChange={setFromOrigin}>Start at canvas 0, 0 (otherwise the selection&apos;s top-left)</Check>
        <div className="flex flex-wrap gap-1">
          <SmallButton disabled={!n} onClick={() => onArrange("row", { columns, gapX, gapY, fromCanvasOrigin: fromOrigin })}>In a row</SmallButton>
          <SmallButton disabled={!n} onClick={() => onArrange("column", { columns, gapX, gapY, fromCanvasOrigin: fromOrigin })}>In a column</SmallButton>
          <SmallButton disabled={!n} onClick={() => onArrange("grid", { columns, gapX, gapY, fromCanvasOrigin: fromOrigin })}>In a grid</SmallButton>
        </div>
      </div>
      <div className="grid grid-cols-2 gap-2">
        <NumField label="Arrow key step (px)" value={config.nudge.step} min={1} max={4096} onCommit={(v) => onNudge(v, config.nudge.shiftStep)} />
        <NumField label="Shift + arrow (px)" value={config.nudge.shiftStep} min={1} max={4096} onCommit={(v) => onNudge(config.nudge.step, v)} />
      </div>
      <div className="text-[11px] text-slate-500">Locked screens never move; they can still be the reference.</div>
    </Panel>
  );
}

// ---------------------------------------------------------------------------
// Canvas settings: size, background, snapping, LED calculator, processor.
// ---------------------------------------------------------------------------

export function CanvasSettings({
  config,
  update,
  onSavePreset,
  onDeletePreset,
  warnings,
  hasSelection,
  onLedApply,
}: {
  config: GeneratorConfig;
  update: (fn: (c: GeneratorConfig) => GeneratorConfig) => void;
  onSavePreset: (p: ResolutionPreset) => void;
  onDeletePreset: (name: string) => void;
  warnings: string[];
  hasSelection: boolean;
  onLedApply: (target: "canvas" | "selected" | "new", w: number, h: number, led: { cols: number; rows: number; panelPxW: number; panelPxH: number; panelName: string }) => void;
}) {
  const c = config.canvas;
  const setCanvas = (patch: Partial<GeneratorConfig["canvas"]>) => update((cfg) => ({ ...cfg, canvas: { ...cfg.canvas, ...patch } }));
  const setSnap = (patch: Partial<GeneratorConfig["snap"]>) => update((cfg) => ({ ...cfg, snap: { ...cfg.snap, ...patch } }));
  const [ledPanel, setLedPanel] = useState(LED_PANEL_PRESETS[0]);
  const [ledCols, setLedCols] = useState(10);
  const [ledRows, setLedRows] = useState(5);
  const [customPx, setCustomPx] = useState({ w: 168, h: 168 });
  const [useCustom, setUseCustom] = useState(false);
  const pw = useCustom ? customPx.w : ledPanel.pxW;
  const ph = useCustom ? customPx.h : ledPanel.pxH;
  const ledW = ledCols * pw;
  const ledH = ledRows * ph;
  const ledInfo = { cols: ledCols, rows: ledRows, panelPxW: pw, panelPxH: ph, panelName: useCustom ? "Custom" : ledPanel.name };
  const proc = config.processor;
  const spec = proc.model ? PROCESSOR_SPECS[proc.model] : null;
  return (
    <>
      <Panel title={`Main canvas - ${c.w} × ${c.h}`}>
        <TextField label="Output name" value={config.name} onCommit={(v) => update((cfg) => ({ ...cfg, name: v.trim() || cfg.name }))} />
        <ResolutionEditor
          w={c.w}
          h={c.h}
          aspectLock={c.aspectLock}
          onSize={(w, h) => setCanvas({ w, h })}
          onAspectLock={(v) => setCanvas({ aspectLock: v })}
          customPresets={config.customPresets}
          onSavePreset={onSavePreset}
        />
        {config.customPresets.length ? (
          <div className="space-y-1">
            <Label>Saved presets</Label>
            {config.customPresets.map((p) => (
              <div key={p.name} className="flex items-center justify-between gap-2 text-xs">
                <span className="truncate">{p.name} - {p.w} × {p.h}</span>
                <SmallButton tone="danger" onClick={() => onDeletePreset(p.name)}>Remove</SmallButton>
              </div>
            ))}
          </div>
        ) : null}
        <div className="flex flex-col gap-1">
          <Label>Background (unused canvas area)</Label>
          <ColorField value={c.background} onChange={(v) => setCanvas({ background: v })} />
        </div>
        <NumField label="Loop length (seconds) - every animated pattern and video" value={config.loopSeconds} integer={false} min={1} max={600} onCommit={(v) => update((cfg) => ({ ...cfg, loopSeconds: v }))} />
        {warnings.length ? (
          <div className="rounded-lg border border-amber-400 bg-amber-500/15 p-2 text-xs text-amber-200" data-testid="warnings">
            {warnings.map((w, i) => (
              <div key={i}>⚠ {w}</div>
            ))}
          </div>
        ) : null}
      </Panel>

      <Panel title="Snapping" defaultOpen={false}>
        <Check checked={config.snap.enabled} onChange={(v) => setSnap({ enabled: v })}>Snap while dragging</Check>
        <div className="grid grid-cols-2 gap-1 pl-5">
          <Check checked={config.snap.canvasEdges} onChange={(v) => setSnap({ canvasEdges: v })}>Canvas edges</Check>
          <Check checked={config.snap.canvasCentre} onChange={(v) => setSnap({ canvasCentre: v })}>Canvas centre</Check>
          <Check checked={config.snap.screens} onChange={(v) => setSnap({ screens: v })}>Other screens</Check>
          <Check checked={config.snap.grid} onChange={(v) => setSnap({ grid: v })}>Grid</Check>
        </div>
        <div className="grid grid-cols-2 gap-2">
          <NumField label="Grid size (px)" value={config.snap.gridSize} min={1} max={4096} onCommit={(v) => setSnap({ gridSize: v })} />
          <NumField label="Snap distance (screen px)" value={config.snap.threshold} min={1} max={64} onCommit={(v) => setSnap({ threshold: v })} />
        </div>
      </Panel>

      <Panel title="LED wall resolution calculator" defaultOpen={false}>
        <div className="text-xs text-slate-400">Work out a resolution from panel rows and columns. Optional - every other part of the generator works with plain pixel sizes.</div>
        <SelectField
          label="Panel"
          value={useCustom ? "__custom" : ledPanel.name}
          onChange={(v) => {
            if (v === "__custom") setUseCustom(true);
            else {
              setUseCustom(false);
              setLedPanel(LED_PANEL_PRESETS.find((p) => p.name === v) ?? LED_PANEL_PRESETS[0]);
            }
          }}
        >
          {LED_PANEL_PRESETS.map((p) => (
            <option key={p.name} value={p.name}>{p.name} - {p.pxW} × {p.pxH} px</option>
          ))}
          <option value="__custom">Custom panel size…</option>
        </SelectField>
        {useCustom ? (
          <div className="grid grid-cols-2 gap-2">
            <NumField label="Panel px wide" value={customPx.w} min={1} max={4096} onCommit={(v) => setCustomPx((p) => ({ ...p, w: v }))} />
            <NumField label="Panel px high" value={customPx.h} min={1} max={4096} onCommit={(v) => setCustomPx((p) => ({ ...p, h: v }))} />
          </div>
        ) : null}
        <div className="grid grid-cols-2 gap-2">
          <NumField label="Columns" value={ledCols} min={1} max={500} onCommit={setLedCols} />
          <NumField label="Rows" value={ledRows} min={1} max={500} onCommit={setLedRows} />
        </div>
        <div className="text-sm font-semibold text-white" data-testid="led-result">
          {ledCols} × {ledRows} panels of {pw} × {ph} = {ledW} × {ledH} px <span className="font-normal text-slate-400">({aspectLabel(ledW, ledH)})</span>
        </div>
        <div className="flex flex-wrap gap-1">
          <SmallButton onClick={() => onLedApply("canvas", ledW, ledH, ledInfo)}>Set canvas</SmallButton>
          <SmallButton disabled={!hasSelection} onClick={() => onLedApply("selected", ledW, ledH, ledInfo)}>Set selected screen</SmallButton>
          <SmallButton tone="primary" onClick={() => onLedApply("new", ledW, ledH, ledInfo)}>Add as new screen</SmallButton>
        </div>
      </Panel>

      <Panel title="Processor inputs" defaultOpen={false}>
        <SelectField label="NovaStar processor" value={proc.model} onChange={(v) => update((cfg) => ({ ...cfg, processor: { ...cfg.processor, model: v as GeneratorConfig["processor"]["model"] } }))}>
          <option value="">None</option>
          {PROCESSOR_MODEL_IDS.map((id) => (
            <option key={id} value={id}>{PROCESSOR_SPECS[id].label}</option>
          ))}
        </SelectField>
        {spec ? (
          <>
            <div className="flex gap-1">
              <SmallButton active={proc.inputMode === "whole"} onClick={() => update((cfg) => ({ ...cfg, processor: { ...cfg.processor, inputMode: "whole" } }))}>Whole canvas</SmallButton>
              <SmallButton active={proc.inputMode === "perEntry"} onClick={() => update((cfg) => ({ ...cfg, processor: { ...cfg.processor, inputMode: "perEntry" } }))}>Per sub-screen</SmallButton>
            </div>
            {proc.inputMode === "whole" ? (
              <SelectField label="Canvas input" value={proc.wholeInput === null ? "" : String(proc.wholeInput)} onChange={(v) => update((cfg) => ({ ...cfg, processor: { ...cfg.processor, wholeInput: v ? Number(v) : null } }))}>
                <option value="">Unassigned</option>
                {spec.inputs.map((i) => (
                  <option key={i.interfacePk} value={i.interfacePk}>{i.label}</option>
                ))}
              </SelectField>
            ) : (
              <div className="space-y-1 text-xs">
                {config.screens.map((s) => (
                  <div key={s.id} className="flex justify-between gap-2">
                    <span className="truncate">{s.name}</span>
                    <span className="text-slate-400">{spec.inputs.find((i) => i.interfacePk === s.input)?.label ?? "Unassigned"}</span>
                  </div>
                ))}
                <div className="text-slate-500">Set each screen&apos;s input in its inspector.</div>
              </div>
            )}
            <div className="text-xs text-slate-400">
              Canvas {c.w} × {c.h} of max {spec.maxCanvasWidth} × {spec.maxCanvasHeight}
              {c.w > spec.maxCanvasWidth || c.h > spec.maxCanvasHeight ? " - too big for this processor" : ""}
            </div>
          </>
        ) : null}
      </Panel>
    </>
  );
}

// ---------------------------------------------------------------------------
// Playlist.
// ---------------------------------------------------------------------------

export function PlaylistPanel({
  config,
  activeIndex,
  playing,
  loop,
  onCapture,
  onUpdateStep,
  onRecapture,
  onDelete,
  onMove,
  onLoad,
  onPlay,
  onStop,
  onLoopChange,
}: {
  config: GeneratorConfig;
  activeIndex: number | null;
  playing: boolean;
  loop: boolean;
  onCapture: () => void;
  onUpdateStep: (id: string, patch: Partial<PlaylistStep>) => void;
  onRecapture: (id: string) => void;
  onDelete: (id: string) => void;
  onMove: (id: string, dir: -1 | 1) => void;
  onLoad: (id: string) => void;
  onPlay: () => void;
  onStop: () => void;
  onLoopChange: (v: boolean) => void;
}) {
  const total = config.playlist.reduce((s, p) => s + p.duration, 0);
  return (
    <Panel title={`Playlist (${config.playlist.length} steps${total ? `, ${total}s` : ""})`} defaultOpen={false}>
      <div className="text-xs text-slate-400">Each step remembers the pattern and settings of every screen. Playing steps through them on the editor and every output window.</div>
      <div className="flex flex-wrap gap-1">
        <SmallButton tone="primary" onClick={onCapture}>+ Add step from current patterns</SmallButton>
        {playing ? <SmallButton onClick={onStop}>■ Stop playlist</SmallButton> : <SmallButton disabled={!config.playlist.length} onClick={onPlay}>▶ Play playlist</SmallButton>}
        <Check checked={loop} onChange={onLoopChange}>Loop</Check>
      </div>
      <div className="space-y-1">
        {config.playlist.map((step, i) => (
          <div key={step.id} className={`space-y-1 rounded-md border p-2 ${activeIndex === i ? "border-emerald-400 bg-emerald-500/10" : "border-slate-700 bg-slate-800/60"}`}>
            <div className="grid grid-cols-[1fr_5rem] gap-2">
              <TextField value={step.name} onCommit={(v) => onUpdateStep(step.id, { name: v || step.name })} />
              <NumField value={step.duration} min={1} max={3600} integer={false} onCommit={(v) => onUpdateStep(step.id, { duration: v })} title="Seconds" />
            </div>
            <div className="text-[11px] text-slate-400">
              {config.screens.map((s) => `${s.name}: ${patternName(step.assignments[s.id]?.pattern ?? s.pattern)}`).join(" · ")}
            </div>
            <div className="flex flex-wrap gap-1">
              <SmallButton onClick={() => onLoad(step.id)} title="Put these patterns on the screens in the editor">Load</SmallButton>
              <SmallButton onClick={() => onRecapture(step.id)} title="Replace with the patterns on the screens now">Update</SmallButton>
              <SmallButton disabled={i === 0} onClick={() => onMove(step.id, -1)}>↑</SmallButton>
              <SmallButton disabled={i === config.playlist.length - 1} onClick={() => onMove(step.id, 1)}>↓</SmallButton>
              <SmallButton tone="danger" onClick={() => onDelete(step.id)}>Delete</SmallButton>
            </div>
          </div>
        ))}
      </div>
    </Panel>
  );
}
