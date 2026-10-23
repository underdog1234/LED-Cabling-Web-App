import React from "react";
import { Button } from "../components/ui";

export type ExportChoice = {
  key: string;
  label: string;
  hint: string;
  /** Disabled choices still show, with the reason, rather than vanishing - a button that is simply gone looks like a bug. */
  disabledReason?: string;
};

type Props = {
  title: string;
  intro: string;
  choices: ExportChoice[];
  onChoose: (key: string) => void;
  onClose: () => void;
};

/**
 * "Which of these did you mean?" - the first step of an export button that
 * covers more than one output.
 *
 * Each choice goes straight on to its own existing options dialog rather than
 * trying to merge every option into one screen: the test pattern picks
 * surfaces, the LED map picks a background, and jamming both into a single
 * form would only make each of them harder to read.
 */
export default function ExportChoiceModal({ title, intro, choices, onChoose, onClose }: Props) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4">
      <div className="w-full max-w-lg space-y-4 rounded-xl border border-slate-700 bg-slate-900 p-5 text-sm text-white">
        <div>
          <div className="text-lg font-semibold">{title}</div>
          <div className="mt-1 text-xs text-slate-400">{intro}</div>
        </div>

        <div className="space-y-2">
          {choices.map((choice) => {
            const disabled = Boolean(choice.disabledReason);
            return (
              <button
                key={choice.key}
                type="button"
                disabled={disabled}
                onClick={() => onChoose(choice.key)}
                className={`w-full rounded-lg border p-3 text-left transition ${
                  disabled
                    ? "cursor-not-allowed border-slate-800 bg-slate-900 opacity-60"
                    : "border-slate-600 bg-slate-800 hover:border-sky-400 hover:bg-slate-700"
                }`}
              >
                <div className="font-semibold">{choice.label}</div>
                <div className="mt-0.5 text-xs text-slate-400">{choice.disabledReason ?? choice.hint}</div>
              </button>
            );
          })}
        </div>

        <div className="flex justify-end pt-1">
          <Button intent="secondary" onClick={onClose}>Cancel</Button>
        </div>
      </div>
    </div>
  );
}
