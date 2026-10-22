import { describe, it, expect } from "vitest";
import { isCancelledStatus, peakUsage } from "./rentman/availabilityPeak";
import type { EquipmentAvailabilityProject } from "./rentman/rentmanClient";

// "Can I have 320 panels from the 4th to the 9th?" is a question about the
// worst DAY in that window, not about the sum of everything that touches it.
// These are the cases that were being answered wrongly.

const booking = (
  projectName: string,
  quantity: number,
  planPeriodStart: string,
  planPeriodEnd: string,
): EquipmentAvailabilityProject => ({
  projectName,
  projectNumber: projectName,
  status: "Confirmed",
  quantity,
  planPeriodStart,
  planPeriodEnd,
});

/** Same booking, on a job with a status other than Confirmed. */
const withStatus = (base: EquipmentAvailabilityProject, status: string | null): EquipmentAvailabilityProject => ({
  ...base,
  status,
});

describe("isCancelledStatus", () => {
  it("matches both spellings Rentman uses for the same status", () => {
    // The API's status `name` is "Canceled"; the UI shows the display name
    // "Cancelled". Matching either one exactly would miss half of them.
    expect(isCancelledStatus("Canceled")).toBe(true);
    expect(isCancelledStatus("Cancelled")).toBe(true);
    expect(isCancelledStatus("cancelled by client")).toBe(true);
  });

  it("leaves every other status in this account's list counting", () => {
    // The real list from GET /statuses - a job that is merely not confirmed
    // yet is still kit somebody expects to send out.
    ["Pending", "Confirmed", "Prepped", "On location", "Returned", "Inquiry", "Concept", null, ""].forEach((status) => {
      expect({ status, cancelled: isCancelledStatus(status) }).toEqual({ status, cancelled: false });
    });
  });
});

describe("peakUsage", () => {
  it("does not add up jobs that never clash", () => {
    // The bug: 100 out on the 4th and 100 out on the 8th was reported as 200
    // unavailable, though the first lot is back on the shelf days before the
    // second goes out.
    const result = peakUsage(
      [booking("Job A", 100, "2026-10-04", "2026-10-05"), booking("Job B", 100, "2026-10-08", "2026-10-09")],
      "2026-10-01",
      "2026-10-10",
    );
    expect(result.peak).toBe(100);
    expect(result.total).toBe(200);
    expect(result.overstated).toBe(true);
  });

  it("adds up jobs that do clash, on the day they clash", () => {
    const result = peakUsage(
      [booking("Job A", 100, "2026-10-04", "2026-10-07"), booking("Job B", 60, "2026-10-06", "2026-10-09")],
      "2026-10-01",
      "2026-10-10",
    );
    expect(result).toMatchObject({ peak: 160, peakStart: "2026-10-06", peakEnd: "2026-10-07", total: 160, overstated: false });
    expect(result.peakBookings.map((b) => b.projectName).sort()).toEqual(["Job A", "Job B"]);
  });

  it("counts two jobs sharing a single day as clashing", () => {
    // One derigs in the morning, the other rigs in the afternoon: the same
    // panels cannot do both, and a planner wants that flagged.
    const result = peakUsage(
      [booking("Out", 40, "2026-10-01", "2026-10-05"), booking("In", 40, "2026-10-05", "2026-10-08")],
      "2026-10-01",
      "2026-10-10",
    );
    expect(result).toMatchObject({ peak: 80, peakStart: "2026-10-05", peakEnd: "2026-10-05" });
  });

  it("only counts the days a booking shares with this job", () => {
    // A month-long sub-rental that happens to cover this window counts once,
    // and a booking outside the window does not count at all.
    const result = peakUsage(
      [booking("Long run", 50, "2026-09-01", "2026-10-31"), booking("Before", 999, "2026-08-01", "2026-08-09")],
      "2026-10-04",
      "2026-10-06",
    );
    expect(result).toMatchObject({ peak: 50, peakStart: "2026-10-04", peakEnd: "2026-10-06" });
    expect(result.peakBookings.map((b) => b.projectName)).toEqual(["Long run"]);
  });

  it("reports the whole run the peak holds for, not just its first day", () => {
    const result = peakUsage(
      [booking("A", 10, "2026-10-01", "2026-10-10"), booking("B", 10, "2026-10-03", "2026-10-07")],
      "2026-10-01",
      "2026-10-31",
    );
    expect(result).toMatchObject({ peak: 20, peakStart: "2026-10-03", peakEnd: "2026-10-07" });
  });

  it("holds the peak across a change that does not lower it", () => {
    // B leaves on the same day C arrives, so 20 runs straight through.
    const result = peakUsage(
      [
        booking("A", 10, "2026-10-01", "2026-10-10"),
        booking("B", 10, "2026-10-02", "2026-10-04"),
        booking("C", 10, "2026-10-05", "2026-10-06"),
      ],
      "2026-10-01",
      "2026-10-31",
    );
    expect(result).toMatchObject({ peak: 20, peakStart: "2026-10-02", peakEnd: "2026-10-06" });
  });

  it("says nothing is out when nothing overlaps the window", () => {
    expect(peakUsage([booking("Elsewhere", 80, "2026-01-01", "2026-01-05")], "2026-10-01", "2026-10-10"))
      .toMatchObject({ peak: 0, peakStart: null, peakEnd: null, overstated: false });
    expect(peakUsage([], "2026-10-01", "2026-10-10")).toMatchObject({ peak: 0, total: 0 });
  });

  it("survives dates it cannot read, and quantities it should ignore", () => {
    const result = peakUsage(
      [
        booking("No dates", 50, "", ""),
        booking("Zero", 0, "2026-10-04", "2026-10-05"),
        booking("Real", 25, "2026-10-04", "2026-10-05"),
      ],
      "2026-10-01",
      "2026-10-10",
    );
    expect(result.peak).toBe(25);
  });

  it("treats a booking that ends before it starts as a single day", () => {
    // Bad data should not quietly remove demand from the answer.
    const result = peakUsage([booking("Backwards", 30, "2026-10-05", "2026-10-01")], "2026-10-01", "2026-10-10");
    expect(result).toMatchObject({ peak: 30, peakStart: "2026-10-05", peakEnd: "2026-10-05" });
  });

  it("reads a plan period with a time on it", () => {
    const result = peakUsage(
      [booking("Morning out", 12, "2026-10-04T08:00:00Z", "2026-10-04T18:00:00Z")],
      "2026-10-01",
      "2026-10-10",
    );
    expect(result).toMatchObject({ peak: 12, peakStart: "2026-10-04", peakEnd: "2026-10-04" });
  });

  it("does not hold kit against a job that was cancelled", () => {
    // Rentman keeps a cancelled job's equipment lines, so they still come
    // back from the date-overlap query - and used to be counted, reporting a
    // shortage over kit sitting on the shelf.
    const result = peakUsage(
      [
        booking("Live job", 100, "2026-10-04", "2026-10-06"),
        withStatus(booking("Called off", 200, "2026-10-04", "2026-10-06"), "Canceled"),
      ],
      "2026-10-01",
      "2026-10-10",
    );
    expect(result.peak).toBe(100);
    expect(result.total).toBe(100);
    expect(result.cancelledQuantity).toBe(200);
    expect(result.cancelled.map((b) => b.projectName)).toEqual(["Called off"]);
    // Not in the peak either, so nothing in the UI marks it as making one up.
    expect(result.peakBookings.map((b) => b.projectName)).toEqual(["Live job"]);
  });

  it("reports nothing out when every job in the window is cancelled", () => {
    const result = peakUsage(
      [
        withStatus(booking("Called off", 200, "2026-10-04", "2026-10-06"), "Cancelled"),
        withStatus(booking("Also off", 50, "2026-10-05", "2026-10-07"), "Canceled"),
      ],
      "2026-10-01",
      "2026-10-10",
    );
    expect(result).toMatchObject({ peak: 0, peakStart: null, peakEnd: null, total: 0, cancelledQuantity: 250 });
    expect(result.peakBookings).toEqual([]);
    // Nothing is being overstated - there is no demand here at all.
    expect(result.overstated).toBe(false);
  });

  it("still counts a job that is only pending", () => {
    const result = peakUsage(
      [withStatus(booking("Not confirmed yet", 80, "2026-10-04", "2026-10-06"), "Pending")],
      "2026-10-01",
      "2026-10-10",
    );
    expect(result).toMatchObject({ peak: 80, total: 80, cancelledQuantity: 0 });
    expect(result.cancelled).toEqual([]);
  });
});
