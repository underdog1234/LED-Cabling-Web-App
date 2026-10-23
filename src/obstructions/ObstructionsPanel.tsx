import React, { useState } from "react";
import { Button, Card, CardContent, CardHeader, CardTitle, Input } from "../components/ui";
import { type Obstruction, MIN_OBSTRUCTION_MM } from "./obstructionModel";

type Props = {
  obstructions: Obstruction[];
  onAdd: (name: string) => void;
  onChange: (id: string, patch: Partial<Obstruction>) => void;
  onDelete: (id: string) => void;
};

// Everything here is METRES, because that is how an obstruction gets measured
// on site. The model stores mm, so each field converts on the way in and out.
// The draft string is kept separate from the committed number so a half-typed
// "1." doesn't momentarily parse as 1 and snap the box to a different size
// while someone is still typing.
const toMetres = (mm: number) => (mm / 1000).toFixed(2);

export default function ObstructionsPanel({ obstructions, onAdd, onChange, onDelete }: Props) {
  const [draftName, setDraftName] = useState("");
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);
  const [drafts, setDrafts] = useState<Record<string, string>>({});

  const add = () => {
    onAdd(draftName.trim() || `Obstruction ${obstructions.length + 1}`);
    setDraftName("");
  };

  const fieldKey = (id: string, field: string) => `${id}:${field}`;

  const metreField = (obstruction: Obstruction, field: "wMm" | "hMm" | "xMm" | "yMm", label: string) => {
    const key = fieldKey(obstruction.id, field);
    const draft = drafts[key];
    const isSize = field === "wMm" || field === "hMm";
    return (
      <label className="flex flex-col gap-1">
        <span className="text-[11px] uppercase tracking-wide text-slate-400">{label}</span>
        <Input
          className="bg-white text-black"
          type="number"
          step="0.05"
          min={isSize ? MIN_OBSTRUCTION_MM / 1000 : undefined}
          value={draft ?? toMetres(obstruction[field])}
          onChange={(e: React.ChangeEvent<HTMLInputElement>) => {
            setDrafts((prev) => ({ ...prev, [key]: e.target.value }));
            const metres = Number.parseFloat(e.target.value);
            if (!Number.isFinite(metres)) return;
            const mm = metres * 1000;
            onChange(obstruction.id, { [field]: isSize ? Math.max(MIN_OBSTRUCTION_MM, mm) : mm });
          }}
          onBlur={() => setDrafts((prev) => {
            const next = { ...prev };
            delete next[key];
            return next;
          })}
        />
      </label>
    );
  };

  return (
    <Card className="border-slate-700 bg-slate-800 print-card no-print" collapsible>
      <CardHeader>
        <CardTitle className="text-white [text-shadow:0_0_2px_black]">Obstructions</CardTitle>
      </CardHeader>
      <CardContent className="space-y-3 text-sm text-white [text-shadow:0_0_2px_black]">
        <div className="space-y-1 text-xs text-slate-400">
          <div>
            Anything standing in front of the LED. Drag one around the Panel Layout to position it. Obstructions appear on the
            layout, in the PDF, on every test pattern and on the Basic LED Map, labelled in pixels and metres - they never affect
            panel counts, power, weight or patching.
          </div>
          <div>
            <span className="text-slate-300">X 0, Y 0</span> is centred on the wall with the bottom edge level with the bottom of
            the layout. X positive moves right, Y positive moves down. Measured from the wall, so adding panels does not leave an
            obstruction behind.
          </div>
        </div>

        <div className="flex gap-2">
          <Input
            className="bg-white text-black"
            placeholder="Name (e.g. PA stack left)"
            value={draftName}
            onChange={(e: React.ChangeEvent<HTMLInputElement>) => setDraftName(e.target.value)}
            onKeyDown={(e: React.KeyboardEvent<HTMLInputElement>) => {
              if (e.key === "Enter") add();
            }}
          />
          <Button size="sm" onClick={add}>Add</Button>
        </div>

        {obstructions.length === 0 ? (
          <div className="rounded border border-slate-700 bg-slate-900 p-3 text-xs text-slate-400">
            No obstructions. Add one to show it on the test patterns and the Basic LED Map.
          </div>
        ) : (
          <div className="space-y-2">
            {obstructions.map((obstruction) => (
              <div key={obstruction.id} className="space-y-2 rounded border border-slate-700 bg-slate-900 p-3">
                <div className="flex items-center gap-2">
                  <Input
                    className="bg-white text-black"
                    value={obstruction.name}
                    onChange={(e: React.ChangeEvent<HTMLInputElement>) => onChange(obstruction.id, { name: e.target.value })}
                  />
                  {confirmDeleteId === obstruction.id ? (
                    <>
                      <Button
                        intent="danger"
                        size="sm"
                        onClick={() => {
                          onDelete(obstruction.id);
                          setConfirmDeleteId(null);
                        }}
                      >
                        Confirm
                      </Button>
                      <Button intent="ghost" size="sm" onClick={() => setConfirmDeleteId(null)}>Cancel</Button>
                    </>
                  ) : (
                    <Button intent="danger" size="sm" onClick={() => setConfirmDeleteId(obstruction.id)}>Delete</Button>
                  )}
                </div>
                <div className="grid grid-cols-2 gap-2 md:grid-cols-4">
                  {metreField(obstruction, "wMm", "Width (m)")}
                  {metreField(obstruction, "hMm", "Height (m)")}
                  {metreField(obstruction, "xMm", "X offset (m)")}
                  {metreField(obstruction, "yMm", "Y offset (m)")}
                </div>
              </div>
            ))}
          </div>
        )}
      </CardContent>
    </Card>
  );
}
