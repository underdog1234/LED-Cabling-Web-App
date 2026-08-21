import { Button } from "../components/ui";
import type { StockComparisonRow } from "./stockOverrides";

type Props = {
  rows: StockComparisonRow[];
  onApply: (overrides: Record<string, number>) => void;
  onClose: () => void;
};

export default function StockComparisonModal({ rows, onApply, onClose }: Props) {
  const changedCount = rows.filter((row) => row.newQuantity !== null && row.newQuantity !== row.oldQuantity).length;
  const applicableCount = rows.filter((row) => row.newQuantity !== null).length;

  const handleApply = () => {
    const overrides: Record<string, number> = {};
    rows.forEach((row) => {
      if (row.newQuantity !== null) overrides[row.code] = row.newQuantity;
    });
    onApply(overrides);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 no-print" onMouseDown={onClose}>
      <div
        className="max-h-[85vh] w-full max-w-3xl overflow-auto rounded-xl border border-slate-600 bg-slate-900 p-5 text-white shadow-2xl"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <div className="mb-3 flex items-center justify-between gap-4">
          <div className="text-lg font-bold">Current Stock from Rentman</div>
          <Button variant="outline" onClick={onClose}>Close</Button>
        </div>
        <p className="mb-4 text-sm text-slate-300">
          {changedCount === 0
            ? "Everything matches what's currently stored - nothing to update."
            : `${changedCount} of ${rows.length} item${rows.length === 1 ? "" : "s"} differ from what's currently stored. Review below, then apply if this looks right.`}
        </p>

        <div className="overflow-x-auto rounded border border-slate-700">
          <table className="min-w-full table-fixed text-left text-sm">
            <thead className="bg-slate-800">
              <tr>
                <th className="w-24 px-3 py-2">Code</th>
                <th className="px-3 py-2">Item</th>
                <th className="px-3 py-2">Rentman Item</th>
                <th className="w-20 px-3 py-2 text-right">Old Qty</th>
                <th className="w-20 px-3 py-2 text-right">New Qty</th>
                <th className="w-24 px-3 py-2 text-right">Difference</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row) => {
                const notFound = row.newQuantity === null;
                const changed = !notFound && row.newQuantity !== row.oldQuantity;
                const diff = notFound ? null : row.newQuantity! - row.oldQuantity;
                return (
                  <tr key={row.code} className={`border-t border-slate-700 ${changed ? "bg-amber-500/15" : ""}`}>
                    <td className="px-3 py-2 whitespace-nowrap text-slate-400">{row.code}</td>
                    <td className="px-3 py-2 truncate">{row.localName}</td>
                    <td className={`px-3 py-2 truncate ${notFound ? "text-slate-500 italic" : "text-slate-300"}`}>
                      {row.rentmanName ?? "Not found in Rentman"}
                    </td>
                    <td className="px-3 py-2 text-right">{row.oldQuantity}</td>
                    <td className={`px-3 py-2 text-right ${changed ? "font-semibold" : ""}`}>{notFound ? "-" : row.newQuantity}</td>
                    <td className={`px-3 py-2 text-right font-semibold ${diff && diff > 0 ? "text-emerald-300" : diff && diff < 0 ? "text-red-300" : ""}`}>
                      {diff === null ? "-" : diff > 0 ? `+${diff}` : diff}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        <div className="mt-4 flex flex-wrap justify-end gap-2">
          <Button variant="outline" onClick={onClose}>Cancel</Button>
          <Button intent="primary" onClick={handleApply} disabled={applicableCount === 0}>Apply All Changes</Button>
        </div>
      </div>
    </div>
  );
}
