// Window Management / Screen Details API helpers for automatically opening
// the Moving Test Pattern on a secondary display (see screenDetails.d.ts for
// why the types are hand-declared). Chromium-only, requires a secure context
// (HTTPS or localhost) and the "window-management" permission - every
// function here is written to fail closed (return an explicit "unavailable"
// result, never throw) so callers have one clean fallback path regardless of
// WHY the API isn't available (unsupported browser, insecure context, denied
// permission), per the app's explicit "explain to the user, don't silently
// fail" requirement.

export type ScreenDetailsResult = { ok: true; details: ScreenDetails } | { ok: false; reason: "unsupported" | "insecure-context" | "denied" | "error" };

/** Cheap, permission-free check for whether requesting full screen details is even worth attempting. */
export function isMultiScreenLikely(): boolean {
  if (typeof window === "undefined" || !window.isSecureContext) return false;
  return window.screen?.isExtended === true;
}

/**
 * Requests full screen enumeration. Must be called from a user-gesture
 * handler (e.g. a click) - the browser's permission prompt requires it.
 * Never throws; every failure mode (unsupported browser, insecure context,
 * denied permission, anything else) resolves to a typed `{ ok: false }`
 * result instead, so the caller can show ONE clear explanatory message
 * rather than a special case per reason.
 */
export async function requestScreenDetails(): Promise<ScreenDetailsResult> {
  if (typeof window === "undefined" || !window.isSecureContext) return { ok: false, reason: "insecure-context" };
  if (typeof window.getScreenDetails !== "function") return { ok: false, reason: "unsupported" };
  try {
    if (navigator.permissions?.query) {
      try {
        const status = await navigator.permissions.query({ name: "window-management" as PermissionName });
        if (status.state === "denied") return { ok: false, reason: "denied" };
      } catch {
        // Permission name not recognized by this browser version - fall through and try the real call anyway.
      }
    }
    const details = await window.getScreenDetails();
    return { ok: true, details };
  } catch (err) {
    console.error("getScreenDetails() failed or was denied", err);
    return { ok: false, reason: "denied" };
  }
}

// There is deliberately no "remember this display" here any more. The Moving
// Test Pattern asks which display to use on every launch: on a show floor the
// right screen changes between one launch and the next, and a remembered
// choice silently sending the pattern to yesterday's monitor is worse than
// one extra click. (The old choice was stored under
// "ledCablingTestPatternDisplay:v1"; nothing reads or writes it now.)

/**
 * Opens a new window positioned over the given screen (or the browser's
 * default placement if `screen` is null - the single-display/unsupported
 * fallback). Positioning alone doesn't fullscreen it - the opened window
 * still needs to call requestFullscreen() itself once it has loaded and is
 * actually showing on that screen (see TestPatternView.tsx).
 */
export function openWindowOnScreen(url: string, screen: ScreenDetailed | null): Window | null {
  if (!screen) return window.open(url, "_blank");
  const features = `left=${screen.left},top=${screen.top},width=${screen.width},height=${screen.height}`;
  return window.open(url, "_blank", features);
}
