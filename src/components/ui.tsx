import React from "react";

// ---------------------------------------------------------------------------
// Shared design-system UI primitives.
// A single Button with clear intents and an unmistakable active/selected state
// (filled accent + ring), so tools and modes read as "on" at a glance rather
// than relying on a subtle border change.
// ---------------------------------------------------------------------------

export type ButtonIntent = "primary" | "secondary" | "ghost" | "danger" | "success";
export type ActiveAccent = "sky" | "amber" | "emerald" | "violet";
export type ButtonSize = "sm" | "md";

const INTENT_CLASSES: Record<ButtonIntent, string> = {
  primary: "border-sky-500/60 bg-sky-600 text-white hover:bg-sky-500",
  secondary: "border-slate-600 bg-slate-700 text-slate-100 hover:bg-slate-600",
  ghost: "border-transparent bg-slate-800/40 text-slate-200 hover:bg-slate-700/70",
  danger: "border-rose-500/60 bg-rose-600 text-white hover:bg-rose-500",
  success: "border-emerald-500/60 bg-emerald-600 text-white hover:bg-emerald-500",
};

// Active/selected = bright fill + ring, dark text for contrast. Same treatment
// for every toggle/mode button so active state is consistent everywhere.
const ACTIVE_CLASSES: Record<ActiveAccent, string> = {
  sky: "border-sky-200 bg-sky-400 text-slate-950 ring-2 ring-sky-300/80 shadow-[0_0_0_1px_rgba(255,255,255,0.15)]",
  amber: "border-amber-200 bg-amber-400 text-slate-950 ring-2 ring-amber-300/80 shadow-[0_0_0_1px_rgba(255,255,255,0.15)]",
  emerald: "border-emerald-200 bg-emerald-400 text-slate-950 ring-2 ring-emerald-300/80 shadow-[0_0_0_1px_rgba(255,255,255,0.15)]",
  violet: "border-violet-200 bg-violet-400 text-slate-950 ring-2 ring-violet-300/80 shadow-[0_0_0_1px_rgba(255,255,255,0.15)]",
};

const SIZE_CLASSES: Record<ButtonSize, string> = {
  sm: "px-2.5 py-1.5 text-xs",
  md: "px-3 py-2 text-sm",
};

type ButtonProps = React.ButtonHTMLAttributes<HTMLButtonElement> & {
  intent?: ButtonIntent;
  size?: ButtonSize;
  active?: boolean;
  activeAccent?: ActiveAccent;
  /** Backward-compat with old call sites that pass variant="outline"/"solid". */
  variant?: "outline" | "solid";
};

export const Button = ({
  children,
  className = "",
  intent,
  size = "md",
  active = false,
  activeAccent = "sky",
  variant,
  type = "button",
  ...props
}: ButtonProps) => {
  const resolvedIntent: ButtonIntent = intent ?? (variant === "outline" ? "secondary" : "primary");
  const stateClasses = active ? ACTIVE_CLASSES[activeAccent] : INTENT_CLASSES[resolvedIntent];
  return (
    <button
      className={`inline-flex items-center justify-center gap-2 whitespace-nowrap rounded-lg border font-medium transition-all disabled:cursor-not-allowed disabled:opacity-50 ${SIZE_CLASSES[size]} ${stateClasses} ${className}`}
      type={type}
      aria-pressed={active || undefined}
      {...props}
    >
      {children}
    </button>
  );
};

// ---------------------------------------------------------------------------
// Cards & sections
// ---------------------------------------------------------------------------

export const Card = ({ children, className = "", ...props }: any) => (
  <div className={`rounded-xl border border-slate-700 bg-slate-800/80 p-4 shadow-sm ${className}`} {...props}>
    {children}
  </div>
);

export const CardHeader = ({ children, className = "" }: any) => (
  <div className={`mb-3 ${className}`}>{children}</div>
);
export const CardContent = ({ children, className = "" }: any) => <div className={className}>{children}</div>;
export const CardTitle = ({ children, className = "" }: any) => (
  <div className={`text-base font-semibold text-white [text-shadow:0_0_2px_black] ${className}`}>{children}</div>
);

export const Input = ({ className = "", ...props }: any) => (
  <input className={`w-full rounded-lg border border-slate-500 bg-white p-2 text-sm text-black focus:border-sky-400 focus:outline-none focus:ring-2 focus:ring-sky-400/50 ${className}`} {...props} />
);

export const Select = ({ className = "", ...props }: any) => (
  <select className={`rounded-lg border border-slate-500 bg-white p-2 text-sm text-black focus:border-sky-400 focus:outline-none focus:ring-2 focus:ring-sky-400/50 disabled:opacity-60 ${className}`} {...props} />
);

/** A labelled sub-group inside a card, for visually grouping related controls. */
export const ControlGroup = ({ label, children, className = "" }: any) => (
  <div className={`rounded-lg border border-slate-700/70 bg-slate-900/40 p-3 ${className}`}>
    {label ? (
      <div className="mb-2 text-[11px] font-semibold uppercase tracking-[0.14em] text-slate-400">{label}</div>
    ) : null}
    <div className="flex flex-wrap items-center gap-2">{children}</div>
  </div>
);

/** Small pill showing the current active mode/status. */
export const StatusChip = ({ tone = "sky", children }: { tone?: ActiveAccent; children: React.ReactNode }) => {
  const tones: Record<ActiveAccent, string> = {
    sky: "border-sky-300 bg-sky-100 text-slate-950",
    amber: "border-amber-300 bg-amber-100 text-slate-950",
    emerald: "border-emerald-300 bg-emerald-100 text-slate-950",
    violet: "border-violet-300 bg-violet-100 text-slate-950",
  };
  return (
    <span className={`inline-flex items-center gap-1.5 rounded-full border px-3 py-1 text-xs font-semibold ${tones[tone]}`}>
      {children}
    </span>
  );
};
