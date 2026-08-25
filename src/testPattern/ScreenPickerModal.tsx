// Shown when launching the Moving Test Pattern and more than one secondary
// display is available (see screenPlacement.ts) - lets the user pick which
// one to open the output window on. Modal chrome matches the convention
// already used throughout App.tsx (fixed inset-0 backdrop, centered card).
import { Button } from "../components/ui";

type Props = {
  screens: ScreenDetailed[];
  onSelect: (screen: ScreenDetailed) => void;
  onCancel: () => void;
};

export default function ScreenPickerModal({ screens, onSelect, onCancel }: Props) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 no-print" onMouseDown={onCancel}>
      <div
        className="w-full max-w-md rounded-xl border border-slate-600 bg-slate-900 p-5 text-white shadow-2xl"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <div className="mb-2 text-lg font-bold">Choose an output display</div>
        <p className="mb-4 text-sm text-slate-300">
          More than one secondary display is connected. Pick which one the Moving Test Pattern should open on - this choice is remembered for next time.
        </p>
        <div className="flex flex-col gap-2">
          {screens.map((screen, index) => (
            <button
              key={`${screen.label}-${screen.left}-${screen.top}-${index}`}
              type="button"
              className="rounded-lg border border-slate-600 bg-slate-800 px-4 py-3 text-left hover:bg-slate-700"
              onClick={() => onSelect(screen)}
            >
              <div className="font-semibold">{screen.label || `Display ${index + 1}`}</div>
              <div className="text-xs text-slate-400">
                {screen.width} x {screen.height}
                {screen.isPrimary ? " - primary" : ""}
              </div>
            </button>
          ))}
        </div>
        <div className="mt-4 flex justify-end">
          <Button variant="outline" onClick={onCancel}>Cancel</Button>
        </div>
      </div>
    </div>
  );
}
