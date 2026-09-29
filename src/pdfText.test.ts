import { describe, it, expect } from "vitest";
import { pdfSafeText, STOCK_CATALOG } from "./App";

// jsPDF's built-in fonts are WinAnsi only. Hand one a character outside that
// and the line comes out as spaced-out nonsense - or the export throws and
// takes the whole report with it. Every string in the PDF crosses one
// boundary, and this is what guards it.

describe("pdfSafeText", () => {
  it("writes three-phase the way jsPDF can set it", () => {
    // The adaptor's catalogue name printed as "3 2 A  3 |  P D L ..." before
    // this existed.
    expect(pdfSafeText("32A 3Φ PDL - 32A 3Φ Ceeform Power Adaptor"))
      .toBe("32A 3Ph PDL - 32A 3Ph Ceeform Power Adaptor");
  });

  it("keeps every catalogue name settable", () => {
    Object.values(STOCK_CATALOG).forEach((item) => {
      const safe = pdfSafeText(item.name);
      expect({ code: item.code, outside: safe.match(/[^ -ÿ]/g) }).toEqual({ code: item.code, outside: null });
      expect(safe.length).toBeGreaterThan(0);
    });
  });

  it("leaves ordinary text, accents and symbols WinAnsi already covers alone", () => {
    expect(pdfSafeText("MG9 LED Panel")).toBe("MG9 LED Panel");
    expect(pdfSafeText("Café 200mm x 100mm")).toBe("Cafe 200mm x 100mm"); // combining mark dropped
    expect(pdfSafeText("Café °C ½")).toBe("Café °C ½");
  });

  it("turns the marks the app itself prints into something readable", () => {
    expect(pdfSafeText("↓ 1 → 2")).toBe("v 1 -> 2");
    expect(pdfSafeText("1920 × 1080")).toBe("1920 x 1080");
    expect(pdfSafeText("‘quoted’ — dash")).toBe("'quoted' - dash");
  });

  it("drops what it cannot render rather than corrupting the line around it", () => {
    // A screen name someone typed an emoji into still prints the words.
    expect(pdfSafeText("Stage Left \u{1f3a4} Tower")).toBe("Stage Left Tower");
    expect(pdfSafeText("\u{1f50c} P1 (3)")).toBe("P1 (3)");
  });
});
