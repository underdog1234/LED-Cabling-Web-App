// Small form controls shared across the generator's panels. Numbers are typed
// into a local draft and committed on Enter or blur, so a half-typed value is
// never applied (and never rounded away while you type it).
import React, { useEffect, useState } from "react";

export const Label = ({ children }: { children: React.ReactNode }) => (
  <span className="text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-400">{children}</span>
);

export function NumField({
  value,
  onCommit,
  min,
  max,
  step = 1,
  integer = true,
  className = "",
  label,
  disabled,
  title,
}: {
  value: number;
  onCommit: (v: number) => void;
  min?: number;
  max?: number;
  step?: number;
  integer?: boolean;
  className?: string;
  label?: React.ReactNode;
  disabled?: boolean;
  title?: string;
}) {
  const [draft, setDraft] = useState(String(value));
  const [focused, setFocused] = useState(false);
  useEffect(() => {
    if (!focused) setDraft(String(value));
  }, [value, focused]);
  const commit = () => {
    const n = Number(draft);
    if (draft.trim() === "" || !Number.isFinite(n)) {
      setDraft(String(value));
      return;
    }
    let v = integer ? Math.round(n) : n;
    if (min !== undefined) v = Math.max(min, v);
    if (max !== undefined) v = Math.min(max, v);
    setDraft(String(v));
    if (v !== value) onCommit(v);
  };
  const input = (
    <input
      type="number"
      inputMode={integer ? "numeric" : "decimal"}
      className={`w-full min-w-0 rounded-md border border-slate-600 bg-slate-950 px-2 py-1 text-sm text-white focus:border-sky-400 focus:outline-none disabled:opacity-50 ${className}`}
      value={draft}
      min={min}
      max={max}
      step={step}
      disabled={disabled}
      title={title}
      onFocus={() => setFocused(true)}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={() => {
        setFocused(false);
        commit();
      }}
      onKeyDown={(e) => {
        if (e.key === "Enter") {
          commit();
          (e.target as HTMLInputElement).blur();
        }
        if (e.key === "Escape") {
          setDraft(String(value));
          (e.target as HTMLInputElement).blur();
        }
        e.stopPropagation();
      }}
    />
  );
  if (!label) return input;
  return (
    <label className="flex min-w-0 flex-col gap-1">
      <Label>{label}</Label>
      {input}
    </label>
  );
}

export function TextField({ value, onCommit, label, className = "", placeholder }: { value: string; onCommit: (v: string) => void; label?: React.ReactNode; className?: string; placeholder?: string }) {
  const [draft, setDraft] = useState(value);
  const [focused, setFocused] = useState(false);
  useEffect(() => {
    if (!focused) setDraft(value);
  }, [value, focused]);
  const input = (
    <input
      className={`w-full min-w-0 rounded-md border border-slate-600 bg-slate-950 px-2 py-1 text-sm text-white focus:border-sky-400 focus:outline-none ${className}`}
      value={draft}
      placeholder={placeholder}
      onFocus={() => setFocused(true)}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={() => {
        setFocused(false);
        if (draft !== value) onCommit(draft);
      }}
      onKeyDown={(e) => {
        if (e.key === "Enter") (e.target as HTMLInputElement).blur();
        e.stopPropagation();
      }}
    />
  );
  if (!label) return input;
  return (
    <label className="flex min-w-0 flex-col gap-1">
      <Label>{label}</Label>
      {input}
    </label>
  );
}

export function ColorField({ value, onChange, label }: { value: string; onChange: (v: string) => void; label?: React.ReactNode }) {
  return (
    <label className="flex items-center gap-2 text-sm text-slate-200">
      <input type="color" className="h-7 w-10 cursor-pointer rounded border border-slate-600 bg-slate-950" value={value} onChange={(e) => onChange(e.target.value)} />
      {label ? <span>{label}</span> : null}
      <span className="font-mono text-xs text-slate-500">{value}</span>
    </label>
  );
}

export function Check({ checked, onChange, children, title }: { checked: boolean; onChange: (v: boolean) => void; children: React.ReactNode; title?: string }) {
  return (
    <label className="flex cursor-pointer items-center gap-2 text-sm text-slate-200" title={title}>
      <input type="checkbox" className="h-4 w-4 accent-sky-500" checked={checked} onChange={(e) => onChange(e.target.checked)} />
      <span>{children}</span>
    </label>
  );
}

export function SelectField({
  value,
  onChange,
  children,
  label,
  className = "",
  disabled,
}: {
  value: string;
  onChange: (v: string) => void;
  children: React.ReactNode;
  label?: React.ReactNode;
  className?: string;
  disabled?: boolean;
}) {
  const select = (
    <select
      className={`w-full min-w-0 rounded-md border border-slate-600 bg-slate-950 px-2 py-1 text-sm text-white focus:border-sky-400 focus:outline-none disabled:opacity-50 ${className}`}
      value={value}
      disabled={disabled}
      onChange={(e) => onChange(e.target.value)}
    >
      {children}
    </select>
  );
  if (!label) return select;
  return (
    <label className="flex min-w-0 flex-col gap-1">
      <Label>{label}</Label>
      {select}
    </label>
  );
}

export function Panel({ title, children, actions, defaultOpen = true }: { title: React.ReactNode; children: React.ReactNode; actions?: React.ReactNode; defaultOpen?: boolean }) {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <section className="rounded-xl border border-slate-700 bg-slate-900/80 shadow-sm">
      <header className="flex items-center gap-2 border-b border-slate-800 px-3 py-2">
        <button type="button" className="flex flex-1 items-center gap-2 text-left text-sm font-semibold text-white" onClick={() => setOpen((v) => !v)} aria-expanded={open}>
          <span className={`inline-block text-slate-400 transition-transform ${open ? "rotate-90" : ""}`}>▸</span>
          {title}
        </button>
        {actions}
      </header>
      {open ? <div className="space-y-3 p-3">{children}</div> : null}
    </section>
  );
}

export function SmallButton({
  children,
  onClick,
  title,
  active,
  disabled,
  tone = "default",
  className = "",
}: {
  children: React.ReactNode;
  onClick?: () => void;
  title?: string;
  active?: boolean;
  disabled?: boolean;
  tone?: "default" | "primary" | "danger";
  className?: string;
}) {
  const tones = {
    default: "border-slate-600 bg-slate-800 text-slate-100 hover:bg-slate-700",
    primary: "border-sky-500/60 bg-sky-600 text-white hover:bg-sky-500",
    danger: "border-rose-500/60 bg-rose-700 text-white hover:bg-rose-600",
  };
  return (
    <button
      type="button"
      title={title}
      disabled={disabled}
      aria-pressed={active || undefined}
      onClick={onClick}
      className={`inline-flex items-center justify-center gap-1 whitespace-nowrap rounded-md border px-2 py-1 text-xs font-medium transition disabled:cursor-not-allowed disabled:opacity-40 ${
        active ? "border-sky-200 bg-sky-400 text-slate-950" : tones[tone]
      } ${className}`}
    >
      {children}
    </button>
  );
}
