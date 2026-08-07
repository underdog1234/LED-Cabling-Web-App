// Thin, typed access to exactly the SQLite tables this feature needs inside
// a NovaStar `Userver.db`. Everything not covered here is left completely
// untouched (copied byte-for-byte from the embedded blank template), per the
// requirement to preserve any information not directly related to the LED
// layout.
//
// The embedded template (src/novastar/assets/blank-template.uprj, a real
// blank single-processor NovaStar project) already ships with usable rows
// for every table below except t_cabinet_topology - this class updates the
// template's own rows in place wherever a row already exists (preserving
// every column it doesn't understand) and only INSERTs fresh rows where the
// template truly has none.
//
// Confirmed table meanings (found by diffing a blank template against a
// real, fully-cabled project the user supplied specifically to prove this
// data is stored here - see the plan / PR description for the full byte
// analysis):
//
//   t_canvas               - one row: the LED-screen pixel canvas
//                             (width/height, all_in_one_controller_*).
//   t_cabinet_topology     - one row per physical cabinet/panel: x/y/width/
//                             height/angle (position, size, rotation),
//                             net_port_id+net_port_index (which Ethernet
//                             output), cabinet_index (patching/daisy-chain
//                             order within that output - 0 = starting
//                             cabinet).
//   t_screen_splice_load   - how this canvas relates to a larger multi-
//                             processor wall. This tool always exports a
//                             single processor, so total==region and the
//                             offset is always (0,0) - splicing disabled.
//   t_logic_screen_general - the main "logic screen" (output canvas
//                             resolution feeding the wall). screen_pk 2 in
//                             the embedded template is the primary/PGM
//                             screen (pgm_edit=1, selected=1); screen_pk 1
//                             is an unrelated "MVR" monitoring screen and is
//                             never touched.
//   t_logic_screen_outputs - ties the main logic screen to a physical
//                             output connector; crop_width/crop_height here
//                             is kept equal to the canvas resolution.
//   t_layer_general        - a positioned/scaled input window on the main
//                             logic screen. The template ships exactly 12
//                             pre-built spare rows (layer_pk 1-12) on the
//                             main screen for this purpose; unused ones are
//                             forced disabled so no stray default layer
//                             survives into the export.
//   t_subcard / t_node /
//   t_project_file         - device/slot identity and project name.
//
// CONFIRMED (updated after comparing a real, working reference export): the
// exact numeric value of `net_port_id`/`net_port_index` is not a stable
// "port 1/2/3..." label - it's allocated from the target device's own
// ever-incrementing internal counters, and a fixed low/predictable scheme
// risks colliding with pre-existing port allocations already on the device
// a file gets imported into (manifesting as cabinets silently merging into
// unrelated existing ports instead of getting their own - i.e. "the port
// patching doesn't come through"). See replaceCabinetTopology below for the
// fresh-per-export allocation scheme this uses instead.

import initSqlJs, { type Database as SqlJsDatabase } from "sql.js";

let sqlJsPromise: ReturnType<typeof initSqlJs> | null = null;

/**
 * Loads sql.js once per process. In the browser bundle this fetches the
 * WASM binary Vite copies into the build output; under Vitest/Node it uses
 * sql.js's Node build, which locates its own WASM file on disk without any
 * extra configuration - kept as two branches instead of one clever
 * abstraction so neither environment depends on bundler-specific behaviour
 * the other doesn't have.
 */
function loadSqlJs() {
  if (!sqlJsPromise) {
    if (typeof window !== "undefined") {
      sqlJsPromise = import("sql.js/dist/sql-wasm.wasm?url").then(({ default: wasmUrl }) =>
        initSqlJs({ locateFile: () => wasmUrl }),
      );
    } else {
      sqlJsPromise = initSqlJs();
    }
  }
  return sqlJsPromise;
}

export type CabinetInput = {
  /** 0-based Ethernet output order (becomes net_port_index). */
  netPortIndex: number;
  /** 0-based position within that output's daisy chain (0 = starting cabinet). */
  cabinetIndex: number;
  x: number;
  y: number;
  width: number;
  height: number;
  angle: 0 | 90 | 180 | 270;
};

export type InputLayerAssignment = {
  name: string;
  /** FK into t_interface_baseinfo.interface_pk - see processorModels.ts. */
  interfacePk: number;
  x: number;
  y: number;
  width: number;
  height: number;
};

const MAIN_SCREEN_PK = 2;
const SPARE_LAYER_PKS = Array.from({ length: 12 }, (_, i) => i + 1);

export class NovaDb {
  private constructor(private db: SqlJsDatabase) {}

  /**
   * Opens any Userver.db - this tool's own embedded template, or an
   * arbitrary real-world file for read-only debug inspection (see
   * parseUprj.ts). Does not assert template shape by itself: callers that
   * are about to *write* to the embedded template should call
   * assertTemplateShape() explicitly first (see exportBuilder.ts) - a real,
   * human-edited project file is not expected to keep matching a pristine
   * template's row counts (e.g. its spare layer rows get consumed over
   * time), and read-only parsing shouldn't reject those.
   */
  static async open(bytes: Uint8Array): Promise<NovaDb> {
    const SQL = await loadSqlJs();
    return new NovaDb(new SQL.Database(bytes));
  }

  toBytes(): Uint8Array {
    return this.db.export();
  }

  close(): void {
    this.db.close();
  }

  private scalar(sql: string, params: unknown[] = []): number {
    const stmt = this.db.prepare(sql);
    try {
      stmt.bind(params as never);
      stmt.step();
      return Number(Object.values(stmt.getAsObject())[0] ?? 0);
    } finally {
      stmt.free();
    }
  }

  private all<T>(sql: string, params: unknown[] = []): T[] {
    const stmt = this.db.prepare(sql);
    const rows: T[] = [];
    try {
      stmt.bind(params as never);
      while (stmt.step()) rows.push(stmt.getAsObject() as T);
    } finally {
      stmt.free();
    }
    return rows;
  }

  /**
   * Defensive check that the embedded template still has the exact shape
   * every write method assumes. Throws a clear error instead of silently
   * writing wrong data if the template asset is ever swapped for a
   * differently-structured one. Call this before writing to a database
   * that is expected to be our own embedded template (see
   * exportBuilder.ts) - not appropriate for arbitrary real-world files
   * opened read-only (see parseUprj.ts), which won't necessarily match.
   */
  assertTemplateShape(): void {
    const expectations: [string, number][] = [
      ["SELECT COUNT(*) c FROM t_canvas", 1],
      ["SELECT COUNT(*) c FROM t_screen_splice_load", 1],
      [`SELECT COUNT(*) c FROM t_logic_screen_general WHERE screen_pk = ${MAIN_SCREEN_PK}`, 1],
      [`SELECT COUNT(*) c FROM t_logic_screen_outputs WHERE screen_pk = ${MAIN_SCREEN_PK}`, 1],
      [
        `SELECT COUNT(*) c FROM t_layer_general WHERE screen_pk = ${MAIN_SCREEN_PK} AND layer_pk IN (${SPARE_LAYER_PKS.join(",")})`,
        SPARE_LAYER_PKS.length,
      ],
    ];
    for (const [sql, expected] of expectations) {
      const actual = this.scalar(sql);
      if (actual !== expected) {
        throw new Error(
          `novaDb: embedded template shape mismatch - "${sql}" returned ${actual}, expected ${expected}. ` +
            "The blank-template.uprj asset may have changed shape; NovaDb's assumptions need updating.",
        );
      }
    }
  }

  // ---- t_canvas -----------------------------------------------------------

  setCanvasResolution(width: number, height: number): void {
    this.db.run(
      `UPDATE t_canvas SET width = ?, height = ?, all_in_one_controller_width = ?,
       all_in_one_controller_height = ?, is_custom_resolution = 1, updated_at = ?`,
      [width, height, width, height, new Date().toISOString()],
    );
  }

  readCanvasResolution(): { width: number; height: number } {
    const row = this.all<{ width: number; height: number }>("SELECT width, height FROM t_canvas")[0];
    return { width: row.width, height: row.height };
  }

  // ---- t_cabinet_topology (empty in the template - fresh INSERTs) --------

  replaceCabinetTopology(cabinets: CabinetInput[]): void {
    this.db.run("DELETE FROM t_cabinet_topology");
    const canvasId = this.scalar("SELECT canvas_id FROM t_canvas");
    const now = new Date().toISOString();

    // net_port_index/net_port_id/cabinet_id are NOT stable "port 1/2/3..."
    // labels - confirmed against a real device's export that a brand new
    // cabinet-topology session gets allocated fresh, ever-increasing values
    // from the device's own internal counters rather than restarting at 0
    // each time (net_port_id = <session base> + net_port_index in both an
    // old leftover 110-114 group and a new 1212165-1212169 group seen on the
    // same real device). A fixed low/predictable scheme (e.g. always
    // starting at 0) risks silently colliding with whatever port
    // allocations already exist on the device/project this file gets
    // imported into - which reads exactly like "the port patching didn't
    // come through" (the new cabinets get merged into old, unrelated
    // ports instead of getting their own). A fresh timestamp+random base
    // every export avoids that collision; the relative structure (cabinets
    // grouped consistently per output, in daisy-chain order via
    // cabinet_index) is what's confirmed to be preserved, not the
    // absolute numbers.
    const sessionSeed = Date.now() + Math.floor(Math.random() * 1_000_000);
    const netPortIndexBase = 1_000 + (sessionSeed % 100_000);
    const netPortIdBase = 2_000_000 + (sessionSeed % 1_000_000);
    const cabinetIdBase = BigInt(sessionSeed) * 1_000_000n;

    const usedIndices = Array.from(new Set(cabinets.map((c) => c.netPortIndex))).sort((a, b) => a - b);
    const portIndexMap = new Map(usedIndices.map((original, i) => [original, netPortIndexBase + i]));

    for (const cabinet of cabinets) {
      const netPortIndex = portIndexMap.get(cabinet.netPortIndex)!;
      const portOrdinal = netPortIndex - netPortIndexBase;
      const netPortId = netPortIdBase + portOrdinal;
      const cabinetId = cabinetIdBase + BigInt(portOrdinal) * 65536n + BigInt(cabinet.cabinetIndex);
      this.db.run(
        `INSERT INTO t_cabinet_topology
         (cabinet_id, canvas_id, net_port_id, net_port_index, cabinet_index, cabinet_group_id,
          lock_state, x, y, width, height, angle, created_at, updated_at)
         VALUES (?, ?, ?, ?, ?, 0, 0, ?, ?, ?, ?, ?, ?, ?)`,
        [
          cabinetId.toString(),
          canvasId,
          netPortId,
          netPortIndex,
          cabinet.cabinetIndex,
          cabinet.x,
          cabinet.y,
          cabinet.width,
          cabinet.height,
          cabinet.angle,
          now,
          now,
        ],
      );
    }
  }

  readCabinetTopology(): (CabinetInput & { cabinetId: string })[] {
    return this.all<{
      cabinet_id: string;
      net_port_index: number;
      cabinet_index: number;
      x: number;
      y: number;
      width: number;
      height: number;
      angle: number;
    }>(
      "SELECT cabinet_id, net_port_index, cabinet_index, x, y, width, height, angle FROM t_cabinet_topology ORDER BY net_port_index, cabinet_index",
    ).map((r) => ({
      cabinetId: r.cabinet_id,
      netPortIndex: r.net_port_index,
      cabinetIndex: r.cabinet_index,
      x: r.x,
      y: r.y,
      width: r.width,
      height: r.height,
      angle: r.angle as 0 | 90 | 180 | 270,
    }));
  }

  // ---- t_screen_splice_load (single processor: no active splicing) ------

  setScreenSpliceLoad(width: number, height: number): void {
    this.db.run(
      `UPDATE t_screen_splice_load SET enable = 0, total_width = ?, total_height = ?,
       region_width = ?, region_height = ?, region_start_x = 0, region_start_y = 0, updated_at = ?`,
      [width, height, width, height, new Date().toISOString()],
    );
  }

  // ---- t_logic_screen_general / t_logic_screen_outputs (main screen) ----

  setMainLogicScreen(width: number, height: number, name: string): void {
    this.db.run(`UPDATE t_logic_screen_general SET width = ?, height = ?, name = ?, updated_at = ? WHERE screen_pk = ?`, [
      width,
      height,
      name,
      new Date().toISOString(),
      MAIN_SCREEN_PK,
    ]);
    this.db.run(`UPDATE t_logic_screen_outputs SET crop_width = ?, crop_height = ?, updated_at = ? WHERE screen_pk = ?`, [
      width,
      height,
      new Date().toISOString(),
      MAIN_SCREEN_PK,
    ]);
  }

  readMainLogicScreen(): { width: number; height: number; name: string } {
    const row = this.all<{ width: number; height: number; name: string }>(
      `SELECT width, height, name FROM t_logic_screen_general WHERE screen_pk = ${MAIN_SCREEN_PK}`,
    )[0];
    return row;
  }

  // ---- t_layer_general (12 pre-built spare rows on the main screen) -----

  replaceInputLayers(assignments: InputLayerAssignment[]): void {
    if (assignments.length > SPARE_LAYER_PKS.length) {
      throw new Error(
        `novaDb: ${assignments.length} input assignments exceeds the template's ${SPARE_LAYER_PKS.length} spare layer rows`,
      );
    }
    const now = new Date().toISOString();
    SPARE_LAYER_PKS.forEach((layerPk, i) => {
      const assignment = assignments[i];
      if (assignment) {
        this.db.run(
          `UPDATE t_layer_general SET enable = 1, valid = 1, name = ?, source_id = ?,
           window_x = ?, window_y = ?, window_width = ?, window_height = ?,
           custom_window_x = ?, custom_window_y = ?, custom_window_width = ?, custom_window_height = ?,
           updated_at = ? WHERE layer_pk = ?`,
          [
            assignment.name,
            assignment.interfacePk,
            assignment.x,
            assignment.y,
            assignment.width,
            assignment.height,
            assignment.x,
            assignment.y,
            assignment.width,
            assignment.height,
            now,
            layerPk,
          ],
        );
      } else {
        this.db.run(`UPDATE t_layer_general SET enable = 0, updated_at = ? WHERE layer_pk = ?`, [now, layerPk]);
      }
    });
  }

  readInputLayers(): { name: string; interfacePk: number; x: number; y: number; width: number; height: number }[] {
    return this.all<{
      name: string;
      source_id: number;
      window_x: number;
      window_y: number;
      window_width: number;
      window_height: number;
    }>(
      `SELECT name, source_id, window_x, window_y, window_width, window_height FROM t_layer_general
       WHERE screen_pk = ${MAIN_SCREEN_PK} AND layer_pk IN (${SPARE_LAYER_PKS.join(",")}) AND enable = 1
       ORDER BY layer_pk`,
    ).map((r) => ({
      name: r.name,
      interfacePk: r.source_id,
      x: r.window_x,
      y: r.window_y,
      width: r.window_width,
      height: r.window_height,
    }));
  }

  // ---- identity: t_subcard / subcard.json mirror, t_node, t_project_file -

  setDeviceName(name: string): void {
    this.db.run(`UPDATE t_node SET name = ? WHERE node_pk = (SELECT node_pk FROM t_node LIMIT 1)`, [name]);
  }

  setProjectIdentity(projectId: string, projectName: string): void {
    this.db.run(`UPDATE t_project_file SET identify = ?, name = ? WHERE project_pk = (SELECT project_pk FROM t_project_file LIMIT 1)`, [
      projectId,
      projectName,
    ]);
  }

  /**
   * Mirrors t_subcard back into the plain-JSON shape NovaStar also ships as
   * subcard.json. t_subcard itself has no board_id column - subcard.json's
   * "boardId" is consistently 0 on every sample seen, so it's hardcoded here
   * rather than read from a nonexistent column.
   */
  readSubcardManifest(): { slotId: number; modelId: number; boardId: number; cardType: number; specialFunc: number }[] {
    return this.all<{ slot_id: number; model_id: number; card_type: number; special_func: number }>(
      "SELECT slot_id, model_id, card_type, special_func FROM t_subcard ORDER BY slot_pk",
    ).map((r) => ({
      slotId: r.slot_id,
      modelId: r.model_id,
      boardId: 0,
      cardType: r.card_type,
      specialFunc: r.special_func,
    }));
  }
}
