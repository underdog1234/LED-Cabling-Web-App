import { Fragment, useState } from "react";
import { Button } from "../components/ui";
import type { EquipmentAvailabilityProject } from "./rentmanClient";

export type AvailabilityRow = {
  code: string;
  name: string;
  totalStock: number;
  totalRequired: number;
  remaining: number;
  projects: EquipmentAvailabilityProject[];
};

type Props = {
  rows: AvailabilityRow[];
  dateFrom: string;
  dateTo: string;
  onClose: () => void;
};

const formatDate = (iso: string) => {
  const date = new Date(iso);
  return Number.isNaN(date.getTime()) ? iso : date.toLocaleDateString();
};

export default function AvailabilityModal({ rows, dateFrom, dateTo, onClose }: Props) {
  const [expandedCode, setExpandedCode] = useState<string | null>(null);
  const shortfallCount = rows.filter((row) => row.remaining < 0).length;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 no-print" onMouseDown={onClose}>
      <div
        className="max-h-[85vh] w-full max-w-4xl overflow-auto rounded-xl border border-slate-600 bg-slate-900 p-5 text-white shadow-2xl"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <div className="mb-3 flex items-center justify-between gap-4">
          <div className="text-lg font-bold">Stock Availability, {dateFrom} to {dateTo}</div>
          <Button variant="outline" onClick={onClose}>Close</Button>
        </div>
        <p className="mb-4 text-sm text-slate-300">
          {shortfallCount === 0
            ? "Every item has enough stock left over for this date range after other projects' requirements."
            : `${shortfallCount} item${shortfallCount === 1 ? "" : "s"} may run short in this date range - click a row to see which projects are using it.`}
        </p>

        <div className="overflow-x-auto rounded border border-slate-700">
          <table className="min-w-full table-fixed text-left text-sm">
            <thead className="bg-slate-800">
              <tr>
                <th className="w-24 px-3 py-2">Code</th>
                <th className="px-3 py-2">Item</th>
                <th className="w-28 px-3 py-2 text-right">Total Stock</th>
                <th className="w-32 px-3 py-2 text-right">Required (range)</th>
                <th className="w-24 px-3 py-2 text-right">Remaining</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row) => {
                const isShort = row.remaining < 0;
                const isExpanded = expandedCode === row.code;
                return (
                  <Fragment key={row.code}>
                    <tr
                      className={`cursor-pointer border-t border-slate-700 ${isShort ? "bg-red-500/10" : ""} ${isExpanded ? "bg-slate-800/60" : ""}`}
                      role="button"
                      tabIndex={0}
                      aria-expanded={isExpanded}
                      onClick={() => setExpandedCode(isExpanded ? null : row.code)}
                      onKeyDown={(event) => {
                        if (event.key === "Enter" || event.key === " ") {
                          event.preventDefault();
                          setExpandedCode(isExpanded ? null : row.code);
                        }
                      }}
                    >
                      <td className={`px-3 py-2 whitespace-nowrap ${isShort ? "text-red-200" : ""}`}>
                        <span className="mr-1 inline-block w-3 text-slate-500">{isExpanded ? "▾" : "▸"}</span>
                        {row.code}
                      </td>
                      <td className="px-3 py-2 truncate">{row.name}</td>
                      <td className="px-3 py-2 text-right">{row.totalStock}</td>
                      <td className="px-3 py-2 text-right">{row.totalRequired}</td>
                      <td className={`px-3 py-2 text-right font-semibold ${isShort ? "text-red-300" : "text-emerald-300"}`}>{row.remaining}</td>
                    </tr>
                    {isExpanded ? (
                      <tr className="border-t border-slate-800 bg-slate-950/60">
                        <td colSpan={5} className="px-3 py-2">
                          {row.projects.length === 0 ? (
                            <div className="text-xs text-slate-400">No other projects need this item in this date range.</div>
                          ) : (
                            <table className="min-w-full text-left text-xs">
                              <thead className="text-slate-400">
                                <tr>
                                  <th className="py-1 pr-3">Project</th>
                                  <th className="py-1 pr-3">Number</th>
                                  <th className="py-1 pr-3">Status</th>
                                  <th className="py-1 pr-3 text-right">Qty</th>
                                  <th className="py-1 pr-3">Dates</th>
                                </tr>
                              </thead>
                              <tbody>
                                {row.projects.map((project, index) => (
                                  <tr key={index} className="border-t border-slate-800">
                                    <td className="py-1 pr-3">{project.projectName}</td>
                                    <td className="py-1 pr-3">{project.projectNumber}</td>
                                    <td className="py-1 pr-3">{project.status ?? "-"}</td>
                                    <td className="py-1 pr-3 text-right">{project.quantity}</td>
                                    <td className="py-1 pr-3 whitespace-nowrap">
                                      {formatDate(project.planPeriodStart)} - {formatDate(project.planPeriodEnd)}
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          )}
                        </td>
                      </tr>
                    ) : null}
                  </Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
