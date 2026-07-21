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
  },
});
