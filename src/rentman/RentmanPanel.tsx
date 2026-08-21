import React, { useEffect, useState } from "react";
import { Search, X } from "lucide-react";
import { Button, Card, CardContent, CardHeader, CardTitle, Input } from "../components/ui";
import { searchEquipment, isRentmanProxyConfigured, type EquipmentSearchResult } from "./rentmanClient";
import type { EquipmentMapping, RentmanEquipmentRef } from "./equipmentMapping";

export type MappableItem = { code: string; name: string };

type Props = {
  dateFrom: string;
  dateTo: string;
  onDateFromChange: (value: string) => void;
  onDateToChange: (value: string) => void;
  mappableItems: MappableItem[];
  mapping: EquipmentMapping;
  onMap: (code: string, ref: RentmanEquipmentRef | null) => void;
  onRefresh: () => void;
  refreshing: boolean;
  refreshError: string | null;
  lastRefreshedAt: Date | null;
};

// Small debounced search-as-you-type box, purely local/ephemeral UI state -
// everything that outlives this interaction (the resulting mapping) is
// reported up via onMap, matching the fully prop-driven convention used by
// NovaStarExportPanel.tsx / OutputCanvasPanel.tsx for everything else here.
function EquipmentPicker({ onPick, onCancel }: { onPick: (ref: RentmanEquipmentRef) => void; onCancel: () => void }) {
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<EquipmentSearchResult[]>([]);
  const [searching, setSearching] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const handle = setTimeout(() => {
      setSearching(true);
      setError(null);
      searchEquipment(query)
        .then(setResults)
        .catch((err) => setError(err instanceof Error ? err.message : "Search failed"))
        .finally(() => setSearching(false));
    }, 250);
    return () => clearTimeout(handle);
  }, [query]);

  return (
    <div className="mt-1 space-y-1 rounded border border-slate-600 bg-slate-950 p-2">
      <div className="flex items-center gap-1">
        <Search className="h-3.5 w-3.5 shrink-0 text-slate-400" />
        <Input
          autoFocus
          type="text"
          value={query}
          onChange={(e: React.ChangeEvent<HTMLInputElement>) => setQuery(e.target.value)}
          placeholder="Search Rentman equipment by name or code..."
          className="text-xs"
        />
        <Button size="sm" intent="secondary" onClick={onCancel} title="Cancel">
          <X className="h-3.5 w-3.5" />
        </Button>
      </div>
      {error ? <div className="text-xs text-red-400">{error}</div> : null}
      {searching ? <div className="text-xs text-slate-400">Searching...</div> : null}
      {!searching && !error && results.length === 0 ? <div className="text-xs text-slate-500">No matches.</div> : null}
      <div className="max-h-40 overflow-y-auto">
        {results.map((r) => (
          <button
            key={r.id}
            type="button"
            className="block w-full rounded px-2 py-1 text-left text-xs text-white hover:bg-slate-800"
            onClick={() => onPick({ id: r.id, code: r.code, name: r.name })}
          >
            <span className="text-slate-400">{r.code}</span> {r.name}
          </button>
        ))}
      </div>
    </div>
  );
}

export default function RentmanPanel({
  dateFrom,
  dateTo,
  onDateFromChange,
  onDateToChange,
  mappableItems,
  mapping,
  onMap,
  onRefresh,
  refreshing,
  refreshError,
  lastRefreshedAt,
}: Props) {
  const [editingCode, setEditingCode] = useState<string | null>(null);
  const configured = isRentmanProxyConfigured();
  const mappedCount = mappableItems.filter((item) => mapping[item.code]).length;

  return (
    <Card className="border-slate-700 bg-slate-800 print-card no-print" collapsible defaultOpen={false}>
      <CardHeader>
        <CardTitle className="text-white [text-shadow:0_0_2px_black]">Rentman Integration</CardTitle>
      </CardHeader>
      <CardContent className="space-y-4 text-sm text-white [text-shadow:0_0_2px_black]">
        {!configured ? (
          <div className="rounded-lg border border-amber-500 bg-amber-500/15 p-3 text-xs text-amber-200">
            Not configured - set <code>VITE_RENTMAN_PROXY_URL</code> and rebuild (see <code>.env.example</code> and{" "}
            <code>rentman-proxy/README.md</code>). Stock Calculations works as normal using the built-in numbers until then.
          </div>
        ) : null}

        <div className="grid gap-3 sm:grid-cols-2">
          <div>
            <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">Date From</div>
            <Input type="date" value={dateFrom} onChange={(e: React.ChangeEvent<HTMLInputElement>) => onDateFromChange(e.target.value)} disabled={!configured} />
          </div>
          <div>
            <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">Date To</div>
            <Input type="date" value={dateTo} onChange={(e: React.ChangeEvent<HTMLInputElement>) => onDateToChange(e.target.value)} disabled={!configured} />
          </div>
        </div>
        <div className="text-xs text-slate-400">
          Sets the "Available (range)" column in Stock Calculations - stock minus what's already booked on other Rentman projects overlapping these dates.
          Leave blank to skip availability and only pull live on-hand stock counts.
        </div>

        <div>
          <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
            <div className="text-xs font-semibold uppercase tracking-wide text-slate-400">
              Equipment Mapping ({mappedCount}/{mappableItems.length} mapped)
            </div>
            <Button intent="primary" size="sm" onClick={onRefresh} disabled={!configured || refreshing || mappedCount === 0}>
              {refreshing ? "Refreshing..." : "Refresh Stock"}
            </Button>
          </div>
          {refreshError ? <div className="mb-2 rounded-lg border border-red-500 bg-red-500/15 p-2 text-xs text-red-200">{refreshError}</div> : null}
          {lastRefreshedAt ? <div className="mb-2 text-xs text-slate-500">Last refreshed {lastRefreshedAt.toLocaleString()}</div> : null}

          <div className="space-y-1.5">
            {mappableItems.map((item) => {
              const mapped = mapping[item.code];
              return (
                <div key={item.code} className="rounded border border-slate-700 bg-slate-900/40 p-2">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <div className="text-xs">
                      <span className="text-slate-400">{item.code}</span> {item.name}
                    </div>
                    {mapped && editingCode !== item.code ? (
                      <div className="flex items-center gap-2 text-xs">
                        <span className="text-emerald-300">
                          {mapped.code} {mapped.name}
                        </span>
                        <Button size="sm" intent="secondary" onClick={() => setEditingCode(item.code)}>
                          Change
                        </Button>
                        <Button size="sm" intent="danger" onClick={() => onMap(item.code, null)}>
                          Clear
                        </Button>
                      </div>
                    ) : editingCode !== item.code ? (
                      <Button size="sm" intent="secondary" onClick={() => setEditingCode(item.code)} disabled={!configured}>
                        Map to Rentman equipment...
                      </Button>
                    ) : null}
                  </div>
                  {editingCode === item.code ? (
                    <EquipmentPicker
                      onCancel={() => setEditingCode(null)}
                      onPick={(ref) => {
                        onMap(item.code, ref);
                        setEditingCode(null);
                      }}
                    />
                  ) : null}
                </div>
              );
            })}
            {mappableItems.length === 0 ? <div className="text-xs text-slate-500">No stock items to map yet - build a wall first.</div> : null}
          </div>
        </div>
      </CardContent>
    </Card>
  );
}
