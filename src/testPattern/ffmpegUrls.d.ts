// ffmpeg.wasm's core JS/WASM binaries, imported with Vite's explicit `?url`
// asset suffix (see mp4Encode.ts) so the browser bundle gets a fetchable URL
// instead of Vite trying to inline/parse them as JS - same pattern as
// sql.js's WASM binary in novastar/wasm-url.d.ts.
declare module "@ffmpeg/core?url" {
  const url: string;
  export default url;
}
declare module "@ffmpeg/core/wasm?url" {
  const url: string;
  export default url;
}
