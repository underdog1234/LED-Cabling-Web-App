// Pure helpers for the "is the test pattern genuinely shown 1:1?"
// verification and the devicePixelRatio-change watcher used by
// TestPatternView.tsx. Kept separate/pure so the actual comparison logic
// (not the DOM plumbing around it) is unit-testable.

export type PixelMappingStatus = "1:1" | "scaled";

// Tolerance for the inherent subpixel rounding that happens when a CSS box
// size isn't an exact integer number of device pixels (e.g. a fractional
// devicePixelRatio like 1.25/1.5) - a fraction of a physical pixel of
// mismatch is normal browser layout rounding, not a real scale.
const EPSILON_PHYSICAL_PX = 0.6;

/**
 * Compares the canvas's actual rendered CSS box (as measured via
 * getBoundingClientRect - never assumed from the numbers we set) against its
 * backing-store resolution, to decide whether 1 backing pixel really is
 * landing on 1 physical display pixel. Never report "1:1" from the setup
 * math alone - always from this measurement, so a browser/OS quirk that
 * silently scales the element is still caught and reported honestly.
 */
export function computePixelMappingStatus(
  canvasWidth: number,
  canvasHeight: number,
  renderedCssWidth: number,
  renderedCssHeight: number,
  devicePixelRatio: number,
): PixelMappingStatus {
  const renderedPhysicalWidth = renderedCssWidth * devicePixelRatio;
  const renderedPhysicalHeight = renderedCssHeight * devicePixelRatio;
  const widthMatches = Math.abs(renderedPhysicalWidth - canvasWidth) <= EPSILON_PHYSICAL_PX;
  const heightMatches = Math.abs(renderedPhysicalHeight - canvasHeight) <= EPSILON_PHYSICAL_PX;
  return widthMatches && heightMatches ? "1:1" : "scaled";
}

/**
 * Watches window.devicePixelRatio for changes (e.g. the window moving to a
 * display with a different scale factor, or the browser/OS zoom changing).
 * There's no native DPR-change event, so this uses the standard
 * `matchMedia('(resolution: Xdppx)')` re-registration idiom: a media query
 * tied to the CURRENT dpr fires 'change' the moment dpr crosses that exact
 * value, at which point a fresh query for the new dpr is registered so the
 * next change is caught too. Returns an unsubscribe function.
 */
export function watchDevicePixelRatio(onChange: (dpr: number) => void): () => void {
  let media: MediaQueryList | null = null;
  let disposed = false;

  const register = () => {
    if (disposed) return;
    const dpr = window.devicePixelRatio || 1;
    media = window.matchMedia(`(resolution: ${dpr}dppx)`);
    media.addEventListener("change", handleChange);
  };

  function handleChange() {
    media?.removeEventListener("change", handleChange);
    onChange(window.devicePixelRatio || 1);
    register();
  }

  register();
  return () => {
    disposed = true;
    media?.removeEventListener("change", handleChange);
  };
}
