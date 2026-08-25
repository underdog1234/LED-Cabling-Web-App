// Ambient types for the Window Management / Screen Details API
// (`window.getScreenDetails()`, `screen.isExtended`). Chromium-only,
// experimental, and requires a secure context - not part of any standardized
// WHATWG/W3C IDL, so TypeScript's bundled DOM lib (checked against the
// installed typescript@6.0.2) has no types for it at all. This file only
// declares the surface actually used by screenPlacement.ts/TestPatternView.tsx
// - always feature-detect at runtime (`typeof window.getScreenDetails ===
// "function"`) before calling any of this, since these types describe an API
// that may simply not exist in the running browser.
//
// No imports/exports in this file - that's what makes it a global ambient
// declaration file rather than a module, so `ScreenDetailed`/`ScreenDetails`
// can be referenced anywhere in src/ without an explicit import.

interface ScreenDetailed extends Screen {
  readonly availLeft: number;
  readonly availTop: number;
  readonly left: number;
  readonly top: number;
  readonly isPrimary: boolean;
  readonly isInternal: boolean;
  /** The physical display's own ratio - unaffected by page zoom, unlike window.devicePixelRatio. */
  readonly devicePixelRatio: number;
  readonly label: string;
}

interface ScreenDetails extends EventTarget {
  readonly screens: ScreenDetailed[];
  readonly currentScreen: ScreenDetailed;
  oncurrentscreenchange: ((this: ScreenDetails, ev: Event) => void) | null;
  onscreenschange: ((this: ScreenDetails, ev: Event) => void) | null;
}

interface Window {
  getScreenDetails?: () => Promise<ScreenDetails>;
}

interface Screen {
  /** True when more than one display is available - no permission needed to read this. */
  isExtended?: boolean;
}
