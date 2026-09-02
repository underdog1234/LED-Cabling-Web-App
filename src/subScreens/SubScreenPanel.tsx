import React, { useState } from "react";
import { Button, Card, CardContent, CardHeader, CardTitle, Input } from "../components/ui";
import { SUB_SCREEN_COLORS, normalizeSubScreenColor, type Cell, type SubScreen } from "../App";
import { subScreenPanelCount } from "./subScreenModel";

type Props = {
  subScreens: SubScreen[];
  /** Resolved active sub-screen id - null means Canvas View. */
  activeSubScreenId: string | null;
  grid: Cell[];
  onSelectScreen: (id: string | null) => void;
  onCreate: (name: string) => void;
  onRename: (id: string, name: string) => void;
  onRecolor: (id: string, color: string) => void;
  onDelete: (id: string) => void;
  onSelectAllInSubScreen: (id: string) => void;
};

export default function SubScreenPanel({
  subScreens,
  activeSubScreenId,
  grid,
  onSelectScreen,
  onCreate,
  onRename,
  onRecolor,
  onDelete,
  onSelectAllInSubScreen,
}: Props) {
  const [draftName, setDraftName] = useState("");
  const [renamingId, setRenamingId] = useState<string | null>(null);
  const [renameDraft, setRenameDraft] = useState("");
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);

  const sorted = [...subScreens].sort((a, b) => a.createdAt - b.createdAt);
  const unassignedCount = grid.filter((cell) => !cell.isRemoved && cell.subScreenId === null).length;

  const startCreate = () => {
    const name = draftName.trim();
    if (!name) return;
    onCreate(name);
    setDraftName("");
  };

  return (
    <Card className="border-slate-700 bg-slate-800 print-card no-print" collapsible>
      <CardHeader>
        <CardTitle className="text-white [text-shadow:0_0_2px_black]">Sub-Screens</CardTitle>
      </CardHeader>
      <CardContent className="space-y-3 text-sm text-white [text-shadow:0_0_2px_black]">
        <div className="flex gap-2">
          <Input
            className="bg-white text-black"
            type="text"
            value={draftName}
            onChange={(e) => setDraftName(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") startCreate();
            }}
            placeholder="e.g. Centre Screen, Stage Left Tower"
          />
          <Button intent="primary" size="sm" onClick={startCreate} disabled={!draftName.trim()}>
            Create
          </Button>
        </div>

        <div className="space-y-1">
          <button
            type="button"
            onClick={() => onSelectScreen(null)}
            className={`flex w-full items-center justify-between rounded-lg border px-3 py-2 text-left transition-colors ${
              activeSubScreenId === null ? "border-sky-300 bg-sky-500/20" : "border-slate-600 bg-slate-900 hover:bg-slate-700/60"
            }`}
          >
            <span className="font-semibold">All Screens</span>
            <span className="text-xs text-slate-300">whole layout{unassignedCount ? ` · ${unassignedCount} unassigned` : ""}</span>
          </button>

          {sorted.map((screen, index) => {
            const count = subScreenPanelCount(grid, screen.id);
            const isActive = activeSubScreenId === screen.id;
            const color = normalizeSubScreenColor(screen.color, index);
            return (
              <div
                key={screen.id}
                className={`rounded-lg border px-3 py-2 ${isActive ? "border-sky-300 bg-sky-500/20" : "border-slate-600 bg-slate-900"}`}
              >
                {renamingId === screen.id ? (
                  <div className="flex gap-2">
                    <Input
                      className="bg-white text-black"
                      autoFocus
                      value={renameDraft}
                      onChange={(e) => setRenameDraft(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === "Enter") {
                          const name = renameDraft.trim();
                          if (name) onRename(screen.id, name);
                          setRenamingId(null);
                        } else if (e.key === "Escape") {
                          setRenamingId(null);
                        }
                      }}
                    />
                    <Button
                      intent="primary"
                      size="sm"
                      onClick={() => {
                        const name = renameDraft.trim();
                        if (name) onRename(screen.id, name);
                        setRenamingId(null);
                      }}
                    >
                      Save
                    </Button>
                    <Button intent="ghost" size="sm" onClick={() => setRenamingId(null)}>
                      Cancel
                    </Button>
                  </div>
                ) : (
                  <button type="button" onClick={() => onSelectScreen(screen.id)} className="flex w-full items-center justify-between text-left">
                    <span className="flex min-w-0 items-center gap-2 font-semibold">
                      <span className="h-3 w-3 shrink-0 rounded-full border border-slate-900/60" style={{ backgroundColor: color }} aria-hidden />
                      <span className="truncate">{screen.name}</span>
                    </span>
                    <span className="text-xs text-slate-300">{count} panel{count === 1 ? "" : "s"}</span>
                  </button>
                )}

                {isActive && renamingId !== screen.id ? (
                  <div className="mt-2 space-y-2">
                  <label className="flex items-center gap-2 text-xs text-slate-300">
                    <span className="w-14 shrink-0">Colour</span>
                    <input
                      type="color"
                      value={color}
                      onChange={(e) => onRecolor(screen.id, e.target.value)}
                      className="h-7 w-12 cursor-pointer rounded border border-slate-600 bg-slate-800 p-0.5"
                      title={`Outline colour for ${screen.name} in the Panel Layout`}
                    />
                    <div className="flex flex-wrap gap-1">
                      {SUB_SCREEN_COLORS.map((swatch) => (
                        <button
                          key={swatch}
                          type="button"
                          onClick={() => onRecolor(screen.id, swatch)}
                          aria-label={`Use ${swatch}`}
                          className={`h-5 w-5 rounded-full border ${color === swatch ? "border-white ring-2 ring-white/70" : "border-slate-600"}`}
                          style={{ backgroundColor: swatch }}
                        />
                      ))}
                    </div>
                  </label>
                  <div className="flex flex-wrap gap-1.5">
                    <Button
                      intent="secondary"
                      size="sm"
                      onClick={() => {
                        setRenamingId(screen.id);
                        setRenameDraft(screen.name);
                      }}
                    >
                      Rename
                    </Button>
                    <Button intent="secondary" size="sm" onClick={() => onSelectAllInSubScreen(screen.id)} disabled={count === 0}>
                      Select all
                    </Button>
                    {confirmDeleteId === screen.id ? (
                      <>
                        <Button intent="danger" size="sm" onClick={() => { onDelete(screen.id); setConfirmDeleteId(null); }}>
                          Confirm delete
                        </Button>
                        <Button intent="ghost" size="sm" onClick={() => setConfirmDeleteId(null)}>
                          Cancel
                        </Button>
                      </>
                    ) : (
                      <Button intent="danger" size="sm" onClick={() => setConfirmDeleteId(screen.id)}>
                        Delete
                      </Button>
                    )}
                  </div>
                  </div>
                ) : null}
              </div>
            );
          })}
        </div>

        <div className="text-xs text-slate-400">
          {sorted.length > 0
            ? "Select panels in the workspace, then use \"Assign Selected\" next to Undo/Redo to add them to a sub-screen."
            : "Create a sub-screen above, then select panels in the workspace to assign them to it."}
        </div>
      </CardContent>
    </Card>
  );
}
