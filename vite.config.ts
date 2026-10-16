import { defineConfig } from 'vite';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('.', import.meta.url));

/**
 * Office add-ins must load over HTTPS even in dev. The dev cert lives in
 * ~/.office-addin-dev-certs/ — the office-addin-dev-certs package
 * generates one with `npx office-addin-dev-certs install` (run once per
 * dev machine; gitignored).
 */
function readDevCerts(): { key?: Buffer; cert?: Buffer } {
  try {
    const home = process.env['HOME'] || process.env['USERPROFILE'] || '';
    return {
      key: readFileSync(resolve(home, '.office-addin-dev-certs/localhost.key')),
      cert: readFileSync(resolve(home, '.office-addin-dev-certs/localhost.crt')),
    };
  } catch {
    return {};
  }
}

const devCerts = readDevCerts();

/**
 * Same-origin engine proxy: https://localhost:3000/engine/* → NOVOMCP_ENGINE_URL.
 *
 * The taskpane is an HTTPS page, so Word's webview blocks plain-http fetches
 * to anything but localhost as mixed content — an engine on another machine
 * (e.g. a DGX Spark at http://spark-xxxx.local:8018) is unreachable directly.
 * Routing through the dev server makes those calls same-origin over HTTPS:
 *
 *   NOVOMCP_ENGINE_URL=http://spark-xxxx.local:8018 npm run dev
 *
 * then set the taskpane's Engine URL to https://localhost:3000/engine.
 */
const engineProxy = {
  '/engine': {
    target: process.env['NOVOMCP_ENGINE_URL'] || 'http://localhost:8018',
    changeOrigin: true,
    rewrite: (path: string) => path.replace(/^\/engine/, ''),
  },
};

export default defineConfig({
  root,
  publicDir: 'assets',
  build: {
    target: 'es2022',
    sourcemap: false,
    outDir: 'dist',
    rollupOptions: {
      input: {
        taskpane: resolve(root, 'taskpane.html'),
        commands: resolve(root, 'commands.html'),
      },
      output: { manualChunks: undefined },
    },
  },
  server: {
    port: 3000,
    strictPort: true,
    https: devCerts.cert ? devCerts : undefined,
    proxy: engineProxy,
  },
});
