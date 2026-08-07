// sql.js's WASM binary is imported with Vite's explicit `?url` asset suffix
// (see novaDb.ts) so the browser bundle gets a fetchable URL instead of
// Vite trying to inline/parse it as JS. Vite's own client types don't cover
// `.wasm?url` out of the box, so this ambient declaration fills the gap.
declare module "sql.js/dist/sql-wasm.wasm?url" {
  const url: string;
  export default url;
}
