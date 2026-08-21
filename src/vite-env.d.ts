/// <reference types="vite/client" />

interface ImportMetaEnv {
  /** Base URL of the deployed rentman-proxy Cloudflare Worker - see src/rentman/rentmanClient.ts. */
  readonly VITE_RENTMAN_PROXY_URL?: string;
}
