// Download dialog: what to export (the whole canvas, one sub-screen, or
// several sub-screens as separate files in a ZIP), in which format, and the
// exact file names that will come out - before anything is rendered.
import { useMemo, useState } from "react";
import { MP4_PROFILE, h264LevelFor, keyframeIntervalFor } from "../testPattern/mp4Encode";
import { Check, NumField, SelectField, SmallButton } from "./controls";
import { downloadBlob, itemsFor, runExport, type ExportFormat, type ExportProgress, type ExportScope } from "./exporters";
import { fileSafe, type GeneratorConfig } from "./model";
import { PATTERN_BY_ID, patternName } from "./patterns";

type Props = {
  config: GeneratorConfig;
  selectedIds: string[];
  currentTime: number;
  onClose: () => void;
};

export default function ExportDialog({ config, selectedIds, currentTime, onClose }: Props) {
  const firstSelected = selectedIds.find((id) => config.screens.some((s) => s.id === id)) ?? config.screens[0]?.id ?? null;
  const [scope, setScope] = useState<ExportScope>(selectedIds.length > 1 ? "selected" : selectedIds.length === 1 ? "screen" : "canvas");
  const [screenId, setScreenId] = useState<string | null>(firstSelected);
  const [picked, setPicked] = useState<string[]>(selectedIds.length ? selectedIds : config.screens.map((s) => s.id));
  const [format, setFormat] = useState<ExportFormat>("png");
  const [frameTime, setFrameTime] = useState(Math.round(currentTime * 100) / 100 % config.loopSeconds);
  const [fps, setFps] = useState<number>(MP4_PROFILE.defaultFps);
  const [targetMbps, setTargetMbps] = useState<number>(MP4_PROFILE.defaultTargetMbps);
  const [maxMbps, setMaxMbps] = useState<number>(MP4_PROFILE.defaultMaxMbps);
  const [busy, setBusy] = useState<ExportProgress | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [done, setDone] = useState<string[] | null>(null);

  const items = useMemo(() => itemsFor(config, scope, screenId, picked), [config, scope, screenId, picked]);
  const names = items.map((i) => i.fileBase(format));
  const zipName = `${fileSafe(config.name)}_Selected-Screens_${format.toUpperCase()}.zip`;
  const largest = items.reduce((m, i) => (i.w * i.h > m.w * m.h ? i : m), { w: 0, h: 0 } as { w: number; h: number });
  const level = h264LevelFor(largest.w, largest.h, fps);
  const anyAnimated = (scope === "canvas" ? config.screens.filter((s) => s.visible) : config.screens.filter((s) => (scope === "screen" ? s.id === screenId : picked.includes(s.id)))).some(
    (s) => PATTERN_BY_ID.get(s.pattern)?.animated ?? true,
  );

  const start = async () => {
    setError(null);
    setDone(null);
    setBusy({ label: "Starting", ratio: 0 });
    try {
      const result = await runExport(items, format, { time: frameTime, loopSeconds: config.loopSeconds, video: { fps, targetMbps, maxMbps }, zipName }, setBusy);
      downloadBlob(result.blob, result.filename);
      setDone(result.files.length > 1 ? [result.filename, ...result.files.map((f) => `  ${f}`)] : [result.filename]);
    } catch (e) {
      console.error(e);
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(null);
    }
  };

  const radio = (value: ExportScope, title: string, detail: string) => (
    <label className={`flex cursor-pointer items-start gap-2 rounded-lg border p-2 text-sm ${scope === value ? "border-sky-400 bg-sky-500/10" : "border-slate-700 bg-slate-800"}`}>
      <input type="radio" className="mt-1" checked={scope === value} onChange={() => setScope(value)} />
      <span>
        <span className="font-semibold text-white">{title}</span>
        <span className="block text-xs text-slate-400">{detail}</span>
      </span>
    </label>
  );

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4" onMouseDown={busy ? undefined : onClose}>
      <div className="max-h-[92vh] w-full max-w-2xl overflow-auto rounded-xl border border-slate-600 bg-slate-900 p-5 text-white shadow-2xl" onMouseDown={(e) => e.stopPropagation()}>
        <div className="mb-3 text-lg font-bold">Download test pattern</div>

        <div className="mb-3 space-y-2">
          <div className="text-[11px] font-semibold uppercase tracking-[0.14em] text-slate-400">Output</div>
          {radio("canvas", `Entire canvas - ${config.canvas.w} × ${config.canvas.h}`, "The complete output at its exact resolution, every visible sub-screen in place.")}
          {radio("screen", "One sub-screen", "Just that screen at its own native resolution - no surrounding canvas, nothing from screens that overlap it.")}
          {scope === "screen" ? (
            <div className="pl-7">
              <SelectField value={screenId ?? ""} onChange={setScreenId}>
                {config.screens.map((s) => (
                  <option key={s.id} value={s.id}>
                    {s.name} - {s.w} × {s.h} - {patternName(s.pattern)}
                  </option>
                ))}
              </SelectField>
            </div>
          ) : null}
          {radio("selected", "Selected sub-screens", "Each one as its own file at its own resolution, packed into a ZIP.")}
          {scope === "selected" ? (
            <div className="space-y-1 pl-7">
              <div className="flex gap-2">
                <SmallButton onClick={() => setPicked(config.screens.map((s) => s.id))}>All</SmallButton>
                <SmallButton onClick={() => setPicked([])}>None</SmallButton>
              </div>
              {config.screens.map((s) => (
                <Check key={s.id} checked={picked.includes(s.id)} onChange={(v) => setPicked((prev) => (v ? [...prev, s.id] : prev.filter((x) => x !== s.id)))}>
                  {s.name} <span className="text-slate-400">- {s.w} × {s.h} - {patternName(s.pattern)}</span>
                </Check>
              ))}
            </div>
          ) : null}
        </div>

        <div className="mb-3 space-y-2">
          <div className="text-[11px] font-semibold uppercase tracking-[0.14em] text-slate-400">Format</div>
          <div className="flex flex-wrap gap-2">
            {(["png", "webm", "mp4"] as ExportFormat[]).map((f) => (
              <SmallButton key={f} active={format === f} onClick={() => setFormat(f)}>
                {f === "png" ? "PNG still" : f === "webm" ? "WebM loop" : "MP4 loop"}
              </SmallButton>
            ))}
          </div>
          {format === "png" ? (
            <div className="flex flex-wrap items-end gap-3">
              <div className="w-40">
                <NumField label="Frame at (seconds)" value={frameTime} integer={false} min={0} max={config.loopSeconds} step={0.1} onCommit={setFrameTime} />
              </div>
              <div className="pb-1 text-xs text-slate-400">{anyAnimated ? "Animated patterns are captured at this point in their loop." : "These patterns are static."}</div>
            </div>
          ) : (
            <div className="space-y-2 rounded-lg border border-slate-700 bg-slate-800/60 p-3">
              <div className="flex flex-wrap items-end gap-3">
                <div className="w-28">
                  <SelectField label="Frame rate" value={String(fps)} onChange={(v) => setFps(Number(v))}>
                    {[24, 25, 30, 50, 60].map((v) => (
                      <option key={v} value={v}>{v} fps</option>
                    ))}
                  </SelectField>
                </div>
                {format === "mp4" ? (
                  <>
                    <div className="w-24"><NumField label="Target Mbps" value={targetMbps} min={1} max={500} integer={false} onCommit={setTargetMbps} /></div>
                    <div className="w-24"><NumField label="Max Mbps" value={maxMbps} min={1} max={500} integer={false} onCommit={setMaxMbps} /></div>
                  </>
                ) : null}
              </div>
              <div className="text-xs text-slate-400">
                Exactly one loop ({config.loopSeconds}s) per file, recorded in real time{items.length > 1 ? `, one file after another (${items.length} files, about ${Math.ceil((items.length * config.loopSeconds) / 60)} min)` : ""}.
              </div>
              {format === "mp4" ? (
                <div className="text-xs text-slate-400">
                  {MP4_PROFILE.container} · {MP4_PROFILE.profile} / {level}
                  {level !== MP4_PROFILE.preferredLevel ? ` (${MP4_PROFILE.preferredLevel} can't carry ${largest.w} × ${largest.h})` : ""} · constant {fps} fps · keyframe every {keyframeIntervalFor(fps)} frames · {MP4_PROFILE.bFrames} B-frames · {MP4_PROFILE.pixelFormat} · {MP4_PROFILE.colour}. MP4 adds an encoding pass after recording.
                </div>
              ) : null}
            </div>
          )}
        </div>

        <div className="mb-3 rounded-lg border border-slate-700 bg-slate-950 p-2 font-mono text-xs text-slate-300">
          <div className="mb-1 font-sans text-[11px] font-semibold uppercase tracking-[0.14em] text-slate-500">{items.length > 1 ? `${zipName} will contain` : "File"}</div>
          {names.length ? names.map((n) => <div key={n}>{n}</div>) : <div className="text-amber-300">Choose at least one screen.</div>}
        </div>

        {busy ? (
          <div className="mb-3">
            <div className="mb-1 text-xs text-slate-300">{busy.label}</div>
            <div className="h-2 overflow-hidden rounded bg-slate-800">
              <div className="h-full bg-sky-500 transition-all" style={{ width: `${Math.round(busy.ratio * 100)}%` }} />
            </div>
          </div>
        ) : null}
        {error ? <div className="mb-3 rounded border border-rose-500 bg-rose-500/15 p-2 text-xs text-rose-200">{error}</div> : null}
        {done ? (
          <div className="mb-3 rounded border border-emerald-500 bg-emerald-500/10 p-2 font-mono text-xs text-emerald-200">
            Downloaded:
            {done.map((d) => (
              <div key={d} className="whitespace-pre">{d}</div>
            ))}
          </div>
        ) : null}

        <div className="flex justify-end gap-2">
          <SmallButton onClick={onClose} disabled={!!busy}>Close</SmallButton>
          <SmallButton tone="primary" onClick={start} disabled={!!busy || !items.length}>
            {busy ? "Working…" : "Download"}
          </SmallButton>
        </div>
      </div>
    </div>
  );
}
