import { useState } from "react";
import { Button } from "../components/ui";

export type ExportSection = {
  key: string;
  label: string;
  /** Optional one-line explanation shown under the label. */
  hint?: string;
};

type Props = {
  title: string;
  intro: string;
  sections: ExportSection[];
  confirmLabel: string;
  /**
   * "multi" (default): tick-boxes, everything on by default, any combination.
   * "single": radio buttons, exactly one choice, the first section preselected -
   * for outputs that can only ever render one thing at a time (the live Moving
   * Test Pattern shows one surface on one display; the video records one).
   */
  mode?: "multi" | "single";
  onConfirm: (selected: Set<string>) => void;
  onClose: () => void;
};

/**
 * "Which parts do you actually want?" picker, shown before an export runs.
 *
 * In "multi" mode everything starts checked, so the default is always the
 * complete export - deselecting is an explicit choice, never something that
 * happens by accident because a box defaulted to off. In "single" mode the
 * first option starts selected, so confirming without touching anything gives
 * the whole wall rather than nothing.
 *
 * Shared by the PDF report, the test-pattern PNG package and the Moving Test
 * Pattern so they all feel the same and none grows its own near-identical
 * modal.
 */
export default function ExportSectionsModal({ title, intro, sections, confirmLabel, mode = "multi", onConfirm, onClose }: Props) {
  const single = mode === "single";
  const [selected, setSelected] = useState<Set<string>>(() =>
    single ? new Set(sections.length ? [sections[0].key] : []) : new Set(sections.map((section) => section.key)),
  );

  const toggle = (key: string) => {
    if (single) {
      setSelected(new Set([key]));
      return;
    }
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  const allSelected = selected.size === sections.length;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 no-print" onMouseDown={onClose}>
      <div
        className="max-h-[85vh] w-full max-w-lg overflow-auto rounded-xl border border-slate-600 bg-slate-900 p-5 text-white shadow-2xl"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <div className="mb-2 text-lg font-bold">{title}</div>
        <p className="mb-4 text-sm text-slate-300">{intro}</p>

        <div className="space-y-1.5">
          {sections.map((section) => (
            <label
              key={section.key}
              className="flex cursor-pointer items-start gap-2 rounded-lg border border-slate-700 bg-slate-800/60 px-3 py-2 hover:bg-slate-800"
            >
              <input
                type={single ? "radio" : "checkbox"}
                name={single ? "export-section-choice" : undefined}
                className="mt-1"
                checked={selected.has(section.key)}
                onChange={() => toggle(section.key)}
              />
              <span className="min-w-0">
                <span className="block text-sm font-medium">{section.label}</span>
                {section.hint ? <span className="block text-xs text-slate-400">{section.hint}</span> : null}
              </span>
            </label>
          ))}
        </div>

        <div className="mt-4 flex flex-wrap items-center justify-between gap-2">
          {single ? (
            <span />
          ) : (
            <Button
              intent="ghost"
              size="sm"
              onClick={() => setSelected(allSelected ? new Set() : new Set(sections.map((section) => section.key)))}
            >
              {allSelected ? "Deselect all" : "Select all"}
            </Button>
          )}
          <div className="flex gap-2">
            <Button variant="outline" onClick={onClose}>Cancel</Button>
            <Button intent="primary" onClick={() => onConfirm(selected)} disabled={selected.size === 0}>
              {confirmLabel}
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
}
