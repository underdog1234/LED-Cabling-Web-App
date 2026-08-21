import React from "react";
import { Button, Card, CardContent, CardHeader, CardTitle, Input } from "../components/ui";
import { isRentmanProxyConfigured } from "./rentmanClient";

type Props = {
  dateFrom: string;
  dateTo: string;
  onDateFromChange: (value: string) => void;
  onDateToChange: (value: string) => void;
  onCheckStock: () => void;
  stockChecking: boolean;
  stockCheckError: string | null;
  lastStockCheckedAt: Date | null;
  onCheckAvailability: () => void;
  availabilityChecking: boolean;
  availabilityError: string | null;
};

export default function RentmanPanel({
  dateFrom,
  dateTo,
  onDateFromChange,
  onDateToChange,
  onCheckStock,
  stockChecking,
  stockCheckError,
  lastStockCheckedAt,
  onCheckAvailability,
  availabilityChecking,
  availabilityError,
}: Props) {
  const configured = isRentmanProxyConfigured();

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

        <div>
          <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
            <div className="text-xs font-semibold uppercase tracking-wide text-slate-400">Current Stock</div>
            <Button intent="primary" size="sm" onClick={onCheckStock} disabled={!configured || stockChecking}>
              {stockChecking ? "Checking..." : "Get Current Stock"}
            </Button>
          </div>
          <div className="text-xs text-slate-400">
            Looks up every stock item's current quantity in Rentman by its equipment code and shows a side-by-side comparison to review before updating anything here.
          </div>
          {stockCheckError ? <div className="mt-2 rounded-lg border border-red-500 bg-red-500/15 p-2 text-xs text-red-200">{stockCheckError}</div> : null}
          {lastStockCheckedAt ? <div className="mt-2 text-xs text-slate-500">Last checked {lastStockCheckedAt.toLocaleString()}</div> : null}
        </div>

        <div className="border-t border-slate-700 pt-4">
          <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-400">Stock Availability by Date Range</div>
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
          <div className="mt-2 mb-2 text-xs text-slate-400">
            For each item: total Rentman stock, how much other projects need in this window, which projects those are, and what's left over - so a shortage is easy to spot before you commit to it.
          </div>
          <Button
            intent="primary"
            size="sm"
            onClick={onCheckAvailability}
            disabled={!configured || availabilityChecking || !dateFrom || !dateTo}
          >
            {availabilityChecking ? "Checking..." : "Check Availability"}
          </Button>
          {availabilityError ? <div className="mt-2 rounded-lg border border-red-500 bg-red-500/15 p-2 text-xs text-red-200">{availabilityError}</div> : null}
        </div>
      </CardContent>
    </Card>
  );
}
