import { describe, it, expect } from "vitest";
import { powerFollowOrder } from "./App";

// Match Power To Signal gives each signal port its own power plug - that is
// the point of it. It used to do that for BACKUP ports too, which are the
// same run looped back rather than another line, so each one burned a plug
// and ran the distro out partway down the wall. Measured on a 72-panel wall
// patched across 8 ports with a VX1000 (primary 1-5): 8 plugs before, 6
// after. On a 20-port setup patched across 15, it ran out of plugs entirely
// and left 27 of 72 panels with no power at all.
describe("powerFollowOrder", () => {
  it("gives every primary port its own plug", () => {
    expect(powerFollowOrder([1, 2, 3, 4, 5], true, 5)).toEqual({ freshPlug: [1, 2, 3, 4, 5], followOn: [] });
  });

  it("keeps backup ports off a plug of their own", () => {
    // VX1000 with the loop on: 1-5 are primary, 6-10 back them up.
    expect(powerFollowOrder([1, 2, 3, 4, 5, 6, 7, 8], true, 5)).toEqual({
      freshPlug: [1, 2, 3, 4, 5],
      followOn: [6, 7, 8],
    });
  });

  it("treats the whole range as primary when the loop is off", () => {
    // The same patch with no backup loop: every port is its own line.
    expect(powerFollowOrder([1, 2, 3, 4, 5, 6, 7, 8], false, 10)).toEqual({
      freshPlug: [1, 2, 3, 4, 5, 6, 7, 8],
      followOn: [],
    });
  });

  it("still feeds the backup ports' panels - they are real panels", () => {
    const { freshPlug, followOn } = powerFollowOrder([1, 2, 11, 12], true, 10);
    // Nothing is dropped: every port given is accounted for in one list or
    // the other, which is what stops a panel silently losing its power.
    expect([...freshPlug, ...followOn].sort((a, b) => a - b)).toEqual([1, 2, 11, 12]);
  });

  it("walks in port order whatever order it is handed", () => {
    expect(powerFollowOrder([7, 2, 11, 1], true, 10)).toEqual({ freshPlug: [1, 2, 7], followOn: [11] });
  });

  it("copes with no ports patched at all", () => {
    expect(powerFollowOrder([], true, 10)).toEqual({ freshPlug: [], followOn: [] });
  });
});
