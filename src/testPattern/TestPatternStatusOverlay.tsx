// Presentational status panel for the live Moving Test Pattern tab - reports
// exactly what's actually happening (measured, not assumed - see
// pixelMapping.ts) so the tool can be trusted for real LED testing. Two
// separate, explicitly-labeled mapping stats are shown on purpose: "Browser
// -> Content Canvas" (is the browser genuinely showing the canvas 1:1 on this
// display?) and "MT Vertical Content Mapping" (a static fact about MT's
// content-to-physical-LED relationship) - conflating the two into one
// "Pixel Mapping" figure would make either one easy to misread as the other.

export type TestPatternStatusInfo = {
  physicalW: number;
  physicalH: number;
  contentW: number;
  contentH: number;
  /** True only when contentH !== physicalH (an all-MT wall). */
  isMtContent: boolean;
  /** Null when the display's resolution genuinely can't be determined. */
  displayW: number | null;
  displayH: number | null;
  /** True when Display Resolution came from the Window Management API's ScreenDetailed.devicePixelRatio (unaffected by page zoom) rather than the window.devicePixelRatio fallback (which can be thrown off by zoom). */
  displayResolutionIsExact: boolean;
  canvasW: number;
  canvasH: number;
  devicePixelRatio: number;
  isFullscreen: boolean;
  browserToContentCanvas: "1:1" | "scaled";
};

type Props = {
  info: TestPatternStatusInfo;
  visible: boolean;
};

const Row = ({ label, value, warn }: { label: string; value: string; warn?: boolean }) => (
  <div className="flex justify-between gap-4">
    <span className="text-slate-400">{label}</span>
    <span className={warn ? "font-semibold text-red-300" : "font-semibold text-white"}>{value}</span>
  </div>
);

export default function TestPatternStatusOverlay({ info, visible }: Props) {
  if (!visible) return null;
  const {
    physicalW, physicalH, contentW, contentH, isMtContent,
    displayW, displayH, displayResolutionIsExact,
    canvasW, canvasH, devicePixelRatio, isFullscreen, browserToContentCanvas,
  } = info;
  const isScaled = browserToContentCanvas === "scaled";

  return (
    <div className="pointer-events-none fixed left-3 top-3 z-40 max-w-xs space-y-2 font-mono text-xs">
      {isScaled ? (
        <div className="rounded-lg border-2 border-red-500 bg-red-950/90 px-3 py-2 font-sans text-sm font-bold text-red-200 shadow-lg">
          ⚠ TEST PATTERN IS NOT BEING DISPLAYED 1:1
        </div>
      ) : null}
      <div className="space-y-1 rounded-lg border border-slate-600 bg-slate-950/85 px-3 py-2 text-slate-200 shadow-lg">
        <Row label="Physical LED Resolution:" value={`${physicalW} x ${physicalH}`} />
        {isMtContent ? <Row label="Recommended Content Resolution:" value={`${contentW} x ${contentH}`} /> : null}
        <Row label="Test Pattern Resolution:" value={`${contentW} x ${contentH}`} />
        <Row
          label="Display Resolution:"
          value={displayW && displayH ? `${displayW} x ${displayH}${displayResolutionIsExact ? "" : " (approx.)"}` : "Unknown"}
        />
        <Row label="Canvas Resolution:" value={`${canvasW} x ${canvasH}`} />
        <Row label="Device Pixel Ratio:" value={devicePixelRatio.toFixed(2)} />
        <Row label="Fullscreen:" value={isFullscreen ? "Yes" : "No"} />
        <Row label="Browser -> Content Canvas:" value={isScaled ? "Scaled" : "1:1"} warn={isScaled} />
        {isMtContent ? <Row label="MT Vertical Content Mapping:" value={`${(contentH / physicalH).toFixed(0)}:1`} /> : null}
      </div>
    </div>
  );
}
