import React, { useEffect, useRef, useState } from "react";
import { Button, Input } from "../components/ui";

export type BasicLedMapSettings = {
  backgroundColor: string;
  /** Decoded and ready to draw, or null for a flat colour. */
  backgroundImage: HTMLImageElement | null;
  backgroundImageName: string | null;
  showPieceOutlines: boolean;
};

type Props = {
  /** Shown so it is obvious what the map will cover before anything is generated. */
  summary: string;
  onConfirm: (settings: BasicLedMapSettings) => void;
  onClose: () => void;
};

const DEFAULT_COLOR = "#1d4ed8";

/**
 * Options for the Basic LED Map PNG: a flat colour or an uploaded image
 * stretched across the whole surface.
 *
 * The image is decoded HERE, before the export runs, so a file that turns out
 * not to be a usable image is reported while the dialog is still open rather
 * than failing silently halfway through a download.
 */
export default function BasicLedMapModal({ summary, onConfirm, onClose }: Props) {
  const [backgroundColor, setBackgroundColor] = useState(DEFAULT_COLOR);
  const [image, setImage] = useState<HTMLImageElement | null>(null);
  const [imageName, setImageName] = useState<string | null>(null);
  const [imageError, setImageError] = useState<string | null>(null);
  const [showPieceOutlines, setShowPieceOutlines] = useState(true);
  const objectUrlRef = useRef<string | null>(null);
  const fileRef = useRef<HTMLInputElement | null>(null);

  // The object URL stays alive while the dialog is open (the <img> keeps
  // reading from it), and is released when it closes.
  useEffect(
    () => () => {
      if (objectUrlRef.current) URL.revokeObjectURL(objectUrlRef.current);
    },
    [],
  );

  const pickImage = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setImageError(null);
    if (objectUrlRef.current) URL.revokeObjectURL(objectUrlRef.current);
    const url = URL.createObjectURL(file);
    objectUrlRef.current = url;
    const img = new Image();
    img.onload = () => {
      setImage(img);
      setImageName(`${file.name} (${img.naturalWidth} × ${img.naturalHeight} px)`);
    };
    img.onerror = () => {
      setImage(null);
      setImageName(null);
      setImageError(`${file.name} could not be read as an image.`);
    };
    img.src = url;
  };

  const clearImage = () => {
    if (objectUrlRef.current) URL.revokeObjectURL(objectUrlRef.current);
    objectUrlRef.current = null;
    setImage(null);
    setImageName(null);
    setImageError(null);
    if (fileRef.current) fileRef.current.value = "";
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4">
      <div className="w-full max-w-lg space-y-4 rounded-xl border border-slate-700 bg-slate-900 p-5 text-sm text-white">
        <div>
          <div className="text-lg font-semibold">Basic LED Map</div>
          <div className="mt-1 text-xs text-slate-400">
            A plain picture of the finished surface - every separate piece, with any obstructions drawn over it and labelled in
            pixels and metres. No patching, numbering or test pattern.
          </div>
        </div>

        <div className="rounded border border-slate-700 bg-slate-800 p-3 text-xs text-slate-300">{summary}</div>

        <label className="flex items-center gap-3">
          <span className="w-36 text-xs uppercase tracking-wide text-slate-400">Background colour</span>
          <input
            type="color"
            value={backgroundColor}
            onChange={(e) => setBackgroundColor(e.target.value)}
            className="h-9 w-16 cursor-pointer rounded border border-slate-600 bg-slate-800"
          />
          <Input
            className="w-28 bg-white text-black"
            value={backgroundColor}
            onChange={(e: React.ChangeEvent<HTMLInputElement>) => setBackgroundColor(e.target.value)}
          />
        </label>

        <div className="space-y-2">
          <div className="text-xs uppercase tracking-wide text-slate-400">Background image (optional)</div>
          <div className="text-xs text-slate-400">Stretched across the full surface, so each piece shows its own part of it.</div>
          <div className="flex items-center gap-2">
            <input
              ref={fileRef}
              type="file"
              accept="image/*"
              onChange={pickImage}
              className="block w-full text-xs file:mr-3 file:rounded file:border-0 file:bg-slate-700 file:px-3 file:py-1.5 file:text-white"
            />
            {image ? (
              <Button intent="secondary" size="sm" onClick={clearImage}>Clear</Button>
            ) : null}
          </div>
          {imageName ? <div className="text-xs text-emerald-300">{imageName}</div> : null}
          {imageError ? <div className="text-xs text-rose-300">⚠ {imageError}</div> : null}
        </div>

        <label className="flex items-center gap-2 text-xs">
          <input type="checkbox" checked={showPieceOutlines} onChange={() => setShowPieceOutlines((prev) => !prev)} />
          <span>Outline and label each separate piece</span>
        </label>

        <div className="flex justify-end gap-2 pt-1">
          <Button intent="secondary" onClick={onClose}>Cancel</Button>
          <Button
            onClick={() =>
              onConfirm({ backgroundColor, backgroundImage: image, backgroundImageName: imageName, showPieceOutlines })
            }
          >
            Download PNG
          </Button>
        </div>
      </div>
    </div>
  );
}
