// ---------------------------------------------------------------------------
// Test Pattern Generator - the editor.
//
// A standalone page built from the LED Cabling Planner's Output Canvas and
// Moving Test Pattern, for any output: video, projection, monitors or LED. A
// main canvas of any resolution holds any number of independent sub-screens,
// each with its own resolution, position and running test pattern.
// ---------------------------------------------------------------------------

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import ScreenPickerModal from "../testPattern/ScreenPickerModal";
import { isMultiScreenLikely, openWindowOnScreen, requestScreenDetails } from "../testPattern/screenPlacement";
import CanvasEditor from "./CanvasEditor";
import { Check, NumField, SelectField, SmallButton, TextField } from "./controls";
import ExportDialog from "./ExportDialog";
import { importPlannerProject, looksLikePlannerProject } from "./ledImport";
import {
  alignRects,
  applyPlaylistStep,
  arrangeRects,
  boundsOf,
  defaultConfig,
  distributeRects,
  distributeWithGap,
  fileSafe,
  layoutWarnings,
  makeScreen,
  newId,
  normalizeConfig,
  reorderLayers,
  screenRect,
  type AlignMode,
  type ArrangeMode,
  type GeneratorConfig,
  type LayerMove,
  type PlaylistStep,
  type Rect,
  type ResolutionPreset,
  type SubScreen,
} from "./model";
import OutputView from "./OutputView";
import { ArrangePanel, CanvasSettings, Inspector, PlaylistPanel, ScreenList, type AlignRef } from "./panels";
import { loadLedModule } from "./render";
import {
  defaultPlayback,
  effectiveConfig,
  loadAutosave,
  loadCustomPresets,
  loadSavedConfigs,
  openChannel,
  playbackTime,
  publishOutputState,
  saveAutosave,
  storeCustomPresets,
  storeSavedConfigs,
  type OutputState,
  type Playback,
  type SavedConfig,
} from "./storage";
import { downloadBlob } from "./exporters";
import { useSyncSound } from "./useSyncSound";

const HISTORY_LIMIT = 100;

const initialConfig = (): GeneratorConfig => {
  const saved = loadAutosave();
  const base = saved ?? defaultConfig();
  // Presets saved in this browser are offered in every setup.
  const presets = loadCustomPresets();
  const names = new Set(base.customPresets.map((p) => p.name));
  return { ...base, customPresets: [...base.customPresets, ...presets.filter((p) => !names.has(p.name))] };
};

export default function GeneratorApp() {
  const [config, setConfigState] = useState<GeneratorConfig>(initialConfig);
  const configRef = useRef(config);
  const past = useRef<GeneratorConfig[]>([]);
  const future = useRef<GeneratorConfig[]>([]);
  const [, setHistoryVersion] = useState(0);
  const [selectedIds, setSelectedIds] = useState<string[]>(() => (config.screens[0] ? [config.screens[0].id] : []));
  const [playback, setPlayback] = useState<Playback>(defaultPlayback);
  const playbackRef = useRef(playback);
  const [now, setNow] = useState(Date.now());
  const [showExport, setShowExport] = useState(false);
  const [showConfigs, setShowConfigs] = useState(false);
  const [inlineOutput, setInlineOutput] = useState<{ screenId: string | null } | null>(null);
  const [outputTarget, setOutputTarget] = useState<string>("__canvas");
  const [screenPicker, setScreenPicker] = useState<{ screens: ScreenDetailed[]; url: string } | null>(null);
  const [toast, setToast] = useState<string | null>(null);
  const fileRef = useRef<HTMLInputElement | null>(null);
  const plannerRef = useRef<HTMLInputElement | null>(null);
  const channelRef = useRef<BroadcastChannel | null>(null);

  const commit = useCallback((next: GeneratorConfig, history: boolean) => {
    if (next === configRef.current) return;
    if (history) {
      past.current.push(configRef.current);
      if (past.current.length > HISTORY_LIMIT) past.current.shift();
      future.current = [];
      setHistoryVersion((v) => v + 1);
    }
    configRef.current = next;
    setConfigState(next);
  }, []);

  const update = useCallback((fn: (c: GeneratorConfig) => GeneratorConfig, history = true) => commit(fn(configRef.current), history), [commit]);

  const beginEdit = useCallback(() => {
    past.current.push(configRef.current);
    if (past.current.length > HISTORY_LIMIT) past.current.shift();
    future.current = [];
    setHistoryVersion((v) => v + 1);
  }, []);

  const undo = useCallback(() => {
    const prev = past.current.pop();
    if (!prev) return;
    future.current.push(configRef.current);
    configRef.current = prev;
    setConfigState(prev);
    setHistoryVersion((v) => v + 1);
  }, []);

  const redo = useCallback(() => {
    const next = future.current.pop();
    if (!next) return;
    past.current.push(configRef.current);
    configRef.current = next;
    setConfigState(next);
    setHistoryVersion((v) => v + 1);
  }, []);

  const flash = (message: string) => {
    setToast(message);
    window.setTimeout(() => setToast((t) => (t === message ? null : t)), 3500);
  };

  // Keep the selection to screens that still exist.
  useEffect(() => {
    const ids = new Set(config.screens.map((s) => s.id));
    setSelectedIds((prev) => (prev.every((id) => ids.has(id)) ? prev : prev.filter((id) => ids.has(id))));
  }, [config.screens]);

  // Autosave.
  useEffect(() => {
    const id = window.setTimeout(() => saveAutosave(config), 300);
    return () => window.clearTimeout(id);
  }, [config]);

  useEffect(() => {
    if (config.screens.some((s) => s.ledLayout)) void loadLedModule();
  }, [config.screens]);

  // Output window link.
  useEffect(() => {
    channelRef.current = openChannel();
    return () => channelRef.current?.close();
  }, []);
  const outputStateRef = useRef<OutputState>({ config, playback });
  useEffect(() => {
    playbackRef.current = playback;
    outputStateRef.current = { config, playback };
    const id = window.setTimeout(() => publishOutputState(outputStateRef.current, channelRef.current), 30);
    return () => window.clearTimeout(id);
  }, [config, playback]);

  // A coarse clock for the playback readout and playlist highlight.
  useEffect(() => {
    const id = window.setInterval(() => setNow(Date.now()), 250);
    return () => window.clearInterval(id);
  }, []);

  // What the editor preview draws each frame.
  const frameSource = useRef(() => {
    const p = playbackRef.current;
    const t = Date.now();
    return { config: effectiveConfig(configRef.current, p, t).config, time: playbackTime(p, t) };
  });
  // Effective config caches per (config, playlist step, auto-cycle step) so the preview can skip identical frames.
  const effCache = useRef<{ base: GeneratorConfig | null; step: string | null; result: GeneratorConfig | null }>({ base: null, step: null, result: null });
  frameSource.current = () => {
    const p = playbackRef.current;
    const t = Date.now();
    const eff = effectiveConfig(configRef.current, p, t);
    const c = effCache.current;
    const step = `${eff.stepIndex}|${eff.cycleStep}`;
    if (c.base !== configRef.current || c.step !== step) {
      c.base = configRef.current;
      c.step = step;
      c.result = eff.config;
    }
    return { config: c.result!, time: playbackTime(p, t) };
  };

  const inlineOutputRef = useRef(inlineOutput);
  inlineOutputRef.current = inlineOutput;
  const sound = useSyncSound(
    () => {
      // The inline fullscreen output plays its own tones.
      if (inlineOutputRef.current) return null;
      const f = frameSource.current();
      return { config: f.config, screens: f.config.screens, time: f.time, playing: playbackRef.current.playing };
    },
    "testPatternGenerator:sound:editor",
    true,
  );
  const currentTime = playbackTime(playback, now);
  const activeStep = effectiveConfig(config, playback, now).stepIndex;

  // --- screen operations -----------------------------------------------------

  const patchScreen = (id: string, patch: Partial<SubScreen>, history = true) =>
    update((c) => ({ ...c, screens: c.screens.map((s) => (s.id === id ? { ...s, ...patch } : s)) }), history);

  const setRects = (rects: Record<string, Rect>) =>
    update((c) => ({ ...c, screens: c.screens.map((s) => (rects[s.id] ? { ...s, ...rects[s.id] } : s)) }), false);

  const setPositions = (positions: Record<string, { x: number; y: number }>) =>
    update((c) => ({ ...c, screens: c.screens.map((s) => (positions[s.id] && !s.locked ? { ...s, x: positions[s.id].x, y: positions[s.id].y } : s)) }));

  const addScreen = (partial: Partial<SubScreen> = {}) => {
    const c = configRef.current;
    const w = Math.min(1920, c.canvas.w);
    const h = Math.min(1080, c.canvas.h);
    // First free spot scanning the canvas, else the top-left.
    let pos = { x: 0, y: 0 };
    const occupied = c.screens.filter((s) => s.visible).map(screenRect);
    outer: for (let y = 0; y + h <= c.canvas.h; y += Math.max(1, Math.round(h / 2))) {
      for (let x = 0; x + w <= c.canvas.w; x += Math.max(1, Math.round(w / 2))) {
        if (!occupied.some((r) => x < r.x + r.w && r.x < x + w && y < r.y + r.h && r.y < y + h)) {
          pos = { x, y };
          break outer;
        }
      }
    }
    const used = new Set(c.screens.map((s) => s.name));
    let n = c.screens.length + 1;
    while (used.has(`Screen ${n}`)) n += 1;
    const screen = makeScreen(c.screens.length, { name: `Screen ${n}`, w, h, ...pos, ...partial });
    update((cfg) => ({ ...cfg, screens: [...cfg.screens, screen] }));
    setSelectedIds([screen.id]);
  };

  const duplicate = () => {
    const c = configRef.current;
    const picked = c.screens.filter((s) => selectedIds.includes(s.id));
    if (!picked.length) return;
    const copies = picked.map((s) => ({ ...s, id: newId(), name: `${s.name} copy`, x: s.x + 40, y: s.y + 40, locked: false, settings: JSON.parse(JSON.stringify(s.settings)) }));
    update((cfg) => ({ ...cfg, screens: [...cfg.screens, ...copies] }));
    setSelectedIds(copies.map((s) => s.id));
  };

  const deleteSelected = () => {
    if (!selectedIds.length) return;
    const ids = new Set(selectedIds);
    update((c) => ({
      ...c,
      screens: c.screens.filter((s) => !ids.has(s.id)),
      playlist: c.playlist.map((p) => ({ ...p, assignments: Object.fromEntries(Object.entries(p.assignments).filter(([id]) => !ids.has(id))) })),
    }));
    setSelectedIds([]);
  };

  const layer = (move: LayerMove) => update((c) => ({ ...c, screens: reorderLayers(c.screens, new Set(selectedIds), move) }));

  const nudge = (dx: number, dy: number) => {
    const ids = new Set(selectedIds);
    update((c) => ({ ...c, screens: c.screens.map((s) => (ids.has(s.id) && !s.locked ? { ...s, x: s.x + dx, y: s.y + dy } : s)) }));
  };

  const selectedScreens = () => configRef.current.screens.filter((s) => selectedIds.includes(s.id));

  const align = (mode: AlignMode, ref: AlignRef) => {
    const c = configRef.current;
    const sel = selectedScreens();
    const movable = sel.filter((s) => !s.locked && s.id !== ref);
    let refRect: Rect;
    if (ref === "canvas") refRect = { x: 0, y: 0, w: c.canvas.w, h: c.canvas.h };
    else if (ref === "selection") refRect = boundsOf(sel.map(screenRect));
    else {
      const r = c.screens.find((s) => s.id === ref);
      if (!r) return;
      refRect = screenRect(r);
    }
    const out = alignRects(movable.map(screenRect), mode, refRect);
    setPositions(Object.fromEntries(movable.map((s, i) => [s.id, out[i]])));
  };

  const distribute = (axis: "horizontal" | "vertical") => {
    const movable = selectedScreens();
    // Locked screens hold their place; they still count as fixed points.
    const out = distributeRects(movable.map(screenRect), axis);
    setPositions(Object.fromEntries(movable.map((s, i) => [s.id, out[i]])));
  };

  const space = (axis: "horizontal" | "vertical", gap: number) => {
    const movable = selectedScreens().filter((s) => !s.locked);
    const out = distributeWithGap(movable.map(screenRect), axis, gap);
    setPositions(Object.fromEntries(movable.map((s, i) => [s.id, out[i]])));
  };

  const arrange = (mode: ArrangeMode, opts: { columns: number; gapX: number; gapY: number; fromCanvasOrigin: boolean }) => {
    const movable = selectedScreens().filter((s) => !s.locked);
    if (!movable.length) return;
    const b = boundsOf(movable.map(screenRect));
    const out = arrangeRects(movable.map(screenRect), mode, { ...opts, origin: opts.fromCanvasOrigin ? { x: 0, y: 0 } : { x: b.x, y: b.y } });
    setPositions(Object.fromEntries(movable.map((s, i) => [s.id, out[i]])));
  };

  const matchSize = (refId: string, dim: "w" | "h" | "both") => {
    const r = configRef.current.screens.find((s) => s.id === refId);
    if (!r) return;
    const ids = new Set(selectedIds);
    update((c) => ({
      ...c,
      screens: c.screens.map((s) => (ids.has(s.id) && !s.locked && s.id !== refId ? { ...s, w: dim === "h" ? s.w : r.w, h: dim === "w" ? s.h : r.h } : s)),
    }));
  };

  const savePreset = (p: ResolutionPreset) => {
    update((c) => ({ ...c, customPresets: [...c.customPresets.filter((x) => x.name !== p.name), p] }));
    const stored = loadCustomPresets().filter((x) => x.name !== p.name);
    storeCustomPresets([...stored, p]);
    flash(`Saved preset "${p.name}" (${p.w} × ${p.h}).`);
  };

  const deletePreset = (name: string) => {
    update((c) => ({ ...c, customPresets: c.customPresets.filter((x) => x.name !== name) }));
    storeCustomPresets(loadCustomPresets().filter((x) => x.name !== name));
  };

  const applyLed = (target: "canvas" | "selected" | "new", w: number, h: number, led: SubScreen["led"]) => {
    if (target === "canvas") update((c) => ({ ...c, canvas: { ...c.canvas, w, h } }));
    else if (target === "selected") {
      const ids = new Set(selectedIds);
      update((c) => ({ ...c, screens: c.screens.map((s) => (ids.has(s.id) && !s.locked ? { ...s, w, h, led } : s)) }));
    } else addScreen({ w, h, led, pattern: "led-moving", name: `LED wall ${led ? `${led.cols}×${led.rows}` : ""}`.trim() });
  };

  // --- playback & playlist ------------------------------------------------------

  const setPlaying = (playing: boolean) =>
    setPlayback((p) => {
      const t = playbackTime(p);
      return { ...p, playing, anchor: Date.now(), anchorTime: t };
    });
  const restart = () => setPlayback((p) => ({ ...p, anchor: Date.now(), anchorTime: 0, playlist: { ...p.playlist, startTime: 0 } }));

  const captureAssignments = (c: GeneratorConfig): PlaylistStep["assignments"] =>
    Object.fromEntries(c.screens.map((s) => [s.id, { pattern: s.pattern, settings: { ...(s.settings[s.pattern] ?? {}) } }]));

  const captureStep = () =>
    update((c) => ({
      ...c,
      playlist: [...c.playlist, { id: newId(), name: `Step ${c.playlist.length + 1}`, duration: 10, assignments: captureAssignments(c) }],
    }));

  const playPlaylist = () =>
    setPlayback((p) => {
      const t = playbackTime(p);
      return { ...p, playing: true, anchor: Date.now(), anchorTime: t, playlist: { ...p.playlist, active: true, startTime: t } };
    });

  // --- files ----------------------------------------------------------------------

  const saveFile = () => {
    const blob = new Blob([JSON.stringify(configRef.current, null, 2)], { type: "application/json" });
    downloadBlob(blob, `${fileSafe(configRef.current.name)}.testpattern.json`);
  };

  const openFile = (file: File, planner: boolean) => {
    const reader = new FileReader();
    reader.onload = async () => {
      try {
        const data = JSON.parse(String(reader.result || "{}"));
        if (planner || looksLikePlannerProject(data)) {
          const next = await importPlannerProject(data, configRef.current);
          commit(next, true);
          setSelectedIds(next.screens.map((s) => s.id));
          flash(`Imported ${next.screens.length} screen${next.screens.length === 1 ? "" : "s"} from the LED Cabling Planner project.`);
        } else {
          const next = normalizeConfig(data);
          commit(next, true);
          setSelectedIds(next.screens[0] ? [next.screens[0].id] : []);
          flash(`Opened "${next.name}".`);
        }
      } catch (e) {
        window.alert(e instanceof Error ? e.message : "That file could not be read.");
      }
    };
    reader.readAsText(file);
  };

  // --- output -----------------------------------------------------------------------

  const openOutputWindow = async () => {
    publishOutputState({ config: configRef.current, playback: playbackRef.current }, channelRef.current);
    const params = new URLSearchParams({ output: "1" });
    if (outputTarget !== "__canvas") params.set("screen", outputTarget);
    const url = `${location.pathname}?${params.toString()}`;
    if (!isMultiScreenLikely()) {
      window.open(url, "_blank");
      return;
    }
    const result = await requestScreenDetails();
    if (!result.ok || result.details.screens.length <= 1) {
      window.open(url, "_blank");
      return;
    }
    setScreenPicker({ screens: result.details.screens, url });
  };

  // --- keyboard ---------------------------------------------------------------------

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (inlineOutput || showExport || showConfigs) return;
      const target = e.target as HTMLElement | null;
      if (target && (target.tagName === "INPUT" || target.tagName === "SELECT" || target.tagName === "TEXTAREA" || target.isContentEditable)) return;
      const mod = e.ctrlKey || e.metaKey;
      const step = e.shiftKey ? configRef.current.nudge.shiftStep : configRef.current.nudge.step;
      if (e.key.startsWith("Arrow") && selectedIds.length) {
        e.preventDefault();
        if (e.key === "ArrowLeft") nudge(-step, 0);
        if (e.key === "ArrowRight") nudge(step, 0);
        if (e.key === "ArrowUp") nudge(0, -step);
        if (e.key === "ArrowDown") nudge(0, step);
      } else if ((e.key === "Delete" || e.key === "Backspace") && selectedIds.length) {
        e.preventDefault();
        deleteSelected();
      } else if (mod && e.key.toLowerCase() === "z") {
        e.preventDefault();
        if (e.shiftKey) redo();
        else undo();
      } else if (mod && e.key.toLowerCase() === "y") {
        e.preventDefault();
        redo();
      } else if (mod && e.key.toLowerCase() === "d") {
        e.preventDefault();
        duplicate();
      } else if (mod && e.key.toLowerCase() === "a") {
        e.preventDefault();
        setSelectedIds(configRef.current.screens.map((s) => s.id));
      } else if (e.key === "Escape") {
        setSelectedIds([]);
      } else if (e.key === " " && !mod) {
        e.preventDefault();
        setPlaying(!playbackRef.current.playing);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  });

  const warnings = useMemo(() => layoutWarnings(config), [config]);
  const primary = config.screens.find((s) => s.id === selectedIds[selectedIds.length - 1]) ?? null;
  const editorMaxHeight = Math.max(260, Math.round(window.innerHeight * 0.62));

  return (
    <div className="min-h-screen text-slate-100">
      <header className="sticky top-0 z-30 flex flex-wrap items-center gap-2 border-b border-slate-800 bg-slate-950/95 px-4 py-2 backdrop-blur">
        <div className="mr-2">
          <div className="text-base font-bold text-white">Test Pattern Generator</div>
          <div className="text-[11px] text-slate-400">{config.name} · {config.canvas.w} × {config.canvas.h} · {config.screens.length} sub-screens</div>
        </div>
        <SmallButton onClick={undo} disabled={!past.current.length} title="Undo (Ctrl+Z)">↶ Undo</SmallButton>
        <SmallButton onClick={redo} disabled={!future.current.length} title="Redo (Ctrl+Y)">↷ Redo</SmallButton>
        <span className="mx-1 h-5 w-px bg-slate-700" />
        <SmallButton onClick={() => setShowConfigs(true)}>Configurations</SmallButton>
        <SmallButton onClick={saveFile}>Save file</SmallButton>
        <SmallButton onClick={() => fileRef.current?.click()}>Open file</SmallButton>
        <SmallButton onClick={() => plannerRef.current?.click()} title="Open a project saved by the LED Cabling Planner: its Output Canvas and sub-screens come across">Import LED Planner project</SmallButton>
        <input ref={fileRef} type="file" accept=".json,application/json" className="hidden" onChange={(e) => { const f = e.target.files?.[0]; if (f) openFile(f, false); e.target.value = ""; }} />
        <input ref={plannerRef} type="file" accept=".json,application/json" className="hidden" onChange={(e) => { const f = e.target.files?.[0]; if (f) openFile(f, true); e.target.value = ""; }} />
        <span className="mx-1 h-5 w-px bg-slate-700" />
        <div className="w-48">
          <SelectField value={outputTarget} onChange={setOutputTarget}>
            <option value="__canvas">Output: entire canvas</option>
            {config.screens.map((s) => (
              <option key={s.id} value={s.id}>Output: {s.name}</option>
            ))}
          </SelectField>
        </div>
        <SmallButton onClick={() => void openOutputWindow()} title="Open the output in its own window, on any display">Output window</SmallButton>
        <SmallButton
          onClick={() => {
            setInlineOutput({ screenId: outputTarget === "__canvas" ? null : outputTarget });
            document.documentElement.requestFullscreen?.().catch(() => {});
          }}
        >
          Fullscreen
        </SmallButton>
        <SmallButton tone="primary" onClick={() => setShowExport(true)}>Download…</SmallButton>
        <a className="ml-auto text-xs text-sky-300 hover:underline" href={location.pathname.replace(/generator\.html$/, "")}>LED Cabling Planner ↗</a>
      </header>

      <div className="grid gap-4 p-4 xl:grid-cols-[minmax(0,1fr)_380px]">
        <main className="min-w-0 space-y-3">
          <div className="flex flex-wrap items-center gap-2 rounded-xl border border-slate-700 bg-slate-900/80 px-3 py-2">
            <SmallButton active={playback.playing} onClick={() => setPlaying(!playback.playing)} title="Play / pause (Space)">{playback.playing ? "❚❚ Pause" : "▶ Play"}</SmallButton>
            <SmallButton onClick={restart} title="Back to the start of the loop">⏮ Restart</SmallButton>
            <span className="font-mono text-xs text-slate-300">
              {(currentTime % config.loopSeconds).toFixed(1)}s / {config.loopSeconds}s loop
            </span>
            <span className="mx-1 h-5 w-px bg-slate-700" />
            <Check checked={sound.enabled} onChange={sound.setEnabled} title="Play the timecode and AV sync beeps through this computer, in step with the flashes">
              🔊 Sync beeps
            </Check>
            {sound.enabled ? (
              <>
                <NumField
                  className="w-20"
                  value={sound.offsetMs}
                  min={-1000}
                  max={1000}
                  title="Audio offset in milliseconds: positive delays the beep, to match a display that shows the picture late"
                  onCommit={sound.setOffsetMs}
                />
                <span className="text-xs text-slate-400">ms</span>
                {sound.blocked && sound.active ? <span className="text-xs text-amber-300">Click anywhere to allow sound</span> : null}
              </>
            ) : null}
            <span className="mx-1 h-5 w-px bg-slate-700" />
            <Check
              checked={config.autoCycle.enabled}
              onChange={(v) => update((c) => ({ ...c, autoCycle: { ...c.autoCycle, enabled: v } }))}
              title="Step every screen through all the test patterns in turn (a running playlist takes over while it plays)"
            >
              Auto cycle all patterns
            </Check>
            {config.autoCycle.enabled ? (
              <>
                <span className="text-xs text-slate-400">every</span>
                <NumField className="w-16" value={config.autoCycle.seconds} min={1} max={3600} onCommit={(v) => update((c) => ({ ...c, autoCycle: { ...c.autoCycle, seconds: v } }))} />
                <span className="text-xs text-slate-400">s</span>
                <Check checked={config.autoCycle.stagger} onChange={(v) => update((c) => ({ ...c, autoCycle: { ...c.autoCycle, stagger: v } }))} title="Each screen starts one pattern further on, so they all show something different">
                  Different per screen
                </Check>
              </>
            ) : null}
            {playback.playlist.active && activeStep !== null ? (
              <span className="rounded bg-emerald-500/20 px-2 py-0.5 text-xs text-emerald-200">Playlist: {config.playlist[activeStep]?.name}</span>
            ) : null}
            <span className="ml-auto text-xs text-slate-400">{selectedIds.length ? `${selectedIds.length} selected` : "Nothing selected"}</span>
          </div>
          <div className="rounded-xl border border-slate-700 bg-slate-900/60 p-3">
            <CanvasEditor
              config={config}
              frameSource={frameSource}
              selectedIds={selectedIds}
              onSelect={setSelectedIds}
              onBeginEdit={beginEdit}
              onRects={setRects}
              maxHeight={editorMaxHeight}
            />
          </div>
          {warnings.length ? (
            <div className="rounded-lg border border-amber-400/70 bg-amber-500/10 p-2 text-xs text-amber-200">
              {warnings.map((w, i) => (
                <div key={i}>⚠ {w}</div>
              ))}
            </div>
          ) : null}
          <div className="grid gap-3 lg:grid-cols-2">
            <CanvasSettings
              config={config}
              update={(fn) => update(fn)}
              onSavePreset={savePreset}
              onDeletePreset={deletePreset}
              warnings={[]}
              hasSelection={selectedIds.length > 0}
              onLedApply={applyLed}
            />
          </div>
        </main>

        <aside className="space-y-3">
          <ScreenList
            config={config}
            selectedIds={selectedIds}
            onSelect={setSelectedIds}
            onAdd={() => addScreen()}
            onDuplicate={duplicate}
            onDelete={deleteSelected}
            onPatch={(id, patch) => patchScreen(id, patch)}
            onLayer={layer}
          />
          <ArrangePanel
            config={config}
            selectedIds={selectedIds}
            onAlign={align}
            onDistribute={distribute}
            onSpace={space}
            onArrange={arrange}
            onMatchSize={matchSize}
            onNudge={(step, shiftStep) => update((c) => ({ ...c, nudge: { step, shiftStep } }))}
          />
          {primary ? (
            <Inspector
              key={primary.id}
              config={config}
              screen={primary}
              selectionCount={selectedIds.length}
              onPatch={(patch) => patchScreen(primary.id, patch)}
              onApplyPatternToSelection={() => {
                const ids = new Set(selectedIds);
                update((c) => ({
                  ...c,
                  screens: c.screens.map((s) =>
                    ids.has(s.id) && s.id !== primary.id && (primary.pattern !== "led-layout" || s.ledLayout)
                      ? { ...s, pattern: primary.pattern, settings: { ...s.settings, [primary.pattern]: { ...(primary.settings[primary.pattern] ?? {}) } } }
                      : s,
                  ),
                }));
              }}
              onSavePreset={savePreset}
            />
          ) : null}
          <PlaylistPanel
            config={config}
            activeIndex={playback.playlist.active ? activeStep : null}
            playing={playback.playlist.active}
            loop={playback.playlist.loop}
            onCapture={captureStep}
            onUpdateStep={(id, patch) => update((c) => ({ ...c, playlist: c.playlist.map((p) => (p.id === id ? { ...p, ...patch } : p)) }))}
            onRecapture={(id) => update((c) => ({ ...c, playlist: c.playlist.map((p) => (p.id === id ? { ...p, assignments: captureAssignments(c) } : p)) }))}
            onDelete={(id) => update((c) => ({ ...c, playlist: c.playlist.filter((p) => p.id !== id) }))}
            onMove={(id, dir) =>
              update((c) => {
                const i = c.playlist.findIndex((p) => p.id === id);
                const j = i + dir;
                if (i < 0 || j < 0 || j >= c.playlist.length) return c;
                const list = [...c.playlist];
                [list[i], list[j]] = [list[j], list[i]];
                return { ...c, playlist: list };
              })
            }
            onLoad={(id) => {
              const step = configRef.current.playlist.find((p) => p.id === id);
              if (!step) return;
              setPlayback((p) => ({ ...p, playlist: { ...p.playlist, active: false } }));
              update((c) => applyPlaylistStep(c, step));
            }}
            onPlay={playPlaylist}
            onStop={() => setPlayback((p) => ({ ...p, playlist: { ...p.playlist, active: false } }))}
            onLoopChange={(v) => setPlayback((p) => ({ ...p, playlist: { ...p.playlist, loop: v } }))}
          />
        </aside>
      </div>

      {showExport ? <ExportDialog config={config} selectedIds={selectedIds} currentTime={currentTime} onClose={() => setShowExport(false)} /> : null}
      {showConfigs ? (
        <ConfigsDialog
          config={config}
          onLoad={(c) => {
            commit(c, true);
            setSelectedIds(c.screens[0] ? [c.screens[0].id] : []);
            setShowConfigs(false);
            flash(`Loaded "${c.name}".`);
          }}
          onNew={() => {
            commit({ ...defaultConfig(), screens: [], customPresets: configRef.current.customPresets, name: "Untitled output" }, true);
            setSelectedIds([]);
            setShowConfigs(false);
          }}
          onExample={() => {
            const c = { ...defaultConfig(), customPresets: configRef.current.customPresets };
            commit(c, true);
            setSelectedIds([c.screens[0].id]);
            setShowConfigs(false);
          }}
          onClose={() => setShowConfigs(false)}
        />
      ) : null}
      {inlineOutput ? (
        <OutputView
          stateRef={outputStateRef}
          screenId={inlineOutput.screenId}
          onClose={() => {
            setInlineOutput(null);
            if (document.fullscreenElement) document.exitFullscreen?.().catch(() => {});
          }}
        />
      ) : null}
      {screenPicker ? (
        <ScreenPickerModal
          screens={screenPicker.screens}
          onSelect={(screen) => {
            openWindowOnScreen(screenPicker.url, screen);
            setScreenPicker(null);
          }}
          onCancel={() => setScreenPicker(null)}
        />
      ) : null}
      {toast ? <div className="fixed bottom-4 left-1/2 z-50 -translate-x-1/2 rounded-lg border border-slate-600 bg-slate-900 px-4 py-2 text-sm text-white shadow-xl">{toast}</div> : null}
    </div>
  );
}

function ConfigsDialog({
  config,
  onLoad,
  onNew,
  onExample,
  onClose,
}: {
  config: GeneratorConfig;
  onLoad: (c: GeneratorConfig) => void;
  onNew: () => void;
  onExample: () => void;
  onClose: () => void;
}) {
  const [list, setList] = useState<SavedConfig[]>(loadSavedConfigs);
  const [name, setName] = useState(config.name);
  const [replace, setReplace] = useState(true);
  const save = () => {
    const n = name.trim() || config.name;
    const entry: SavedConfig = { name: n, savedAt: new Date().toISOString(), config: { ...config, name: n } };
    const next = replace ? [...list.filter((x) => x.name !== n), entry] : [...list, entry];
    if (!storeSavedConfigs(next)) {
      window.alert("This browser's storage is full - use Save file instead.");
      return;
    }
    setList(next);
  };
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4" onMouseDown={onClose}>
      <div className="max-h-[90vh] w-full max-w-lg overflow-auto rounded-xl border border-slate-600 bg-slate-900 p-5 text-white shadow-2xl" onMouseDown={(e) => e.stopPropagation()}>
        <div className="mb-1 text-lg font-bold">Saved configurations</div>
        <div className="mb-3 text-xs text-slate-400">
          Each one keeps the whole setup: the canvas, every sub-screen&apos;s position, resolution, pattern, settings and layer order, the playlist and your presets. They live in this browser - use Save file to move one to another computer.
        </div>
        <div className="mb-3 flex items-end gap-2">
          <div className="flex-1"><TextField label="Save current as" value={name} onCommit={setName} /></div>
          <SmallButton tone="primary" onClick={save}>Save</SmallButton>
        </div>
        <Check checked={replace} onChange={setReplace}>Replace a saved configuration with the same name</Check>
        <div className="mt-3 space-y-1">
          {list.length ? (
            list.map((entry, i) => (
              <div key={`${entry.name}-${i}`} className="flex items-center gap-2 rounded-md border border-slate-700 bg-slate-800/60 px-2 py-1.5 text-sm">
                <div className="min-w-0 flex-1">
                  <div className="truncate font-medium">{entry.name}</div>
                  <div className="text-[11px] text-slate-400">
                    {entry.config.canvas.w} × {entry.config.canvas.h} · {entry.config.screens.length} screens{entry.savedAt ? ` · ${new Date(entry.savedAt).toLocaleString()}` : ""}
                  </div>
                </div>
                <SmallButton onClick={() => onLoad(entry.config)}>Load</SmallButton>
                <SmallButton
                  tone="danger"
                  onClick={() => {
                    const next = list.filter((_, j) => j !== i);
                    storeSavedConfigs(next);
                    setList(next);
                  }}
                >
                  Delete
                </SmallButton>
              </div>
            ))
          ) : (
            <div className="text-xs text-slate-400">Nothing saved in this browser yet.</div>
          )}
        </div>
        <div className="mt-4 flex flex-wrap justify-between gap-2">
          <div className="flex gap-2">
            <SmallButton onClick={onNew}>New empty canvas</SmallButton>
            <SmallButton onClick={onExample}>Load example (4 screens)</SmallButton>
          </div>
          <SmallButton onClick={onClose}>Close</SmallButton>
        </div>
      </div>
    </div>
  );
}
