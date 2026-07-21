/**
 * Manifest URL rewriter.
 *
 * The source manifest.xml hardcodes https://addin.novomcp.com URLs (the
 * production target). For local dev, Office requires HTTPS so we rewrite
 * to https://localhost:3000 in the dist copy. AppSource uploads use the
 * source manifest as-is.
 *
 * Usage:
 *   node scripts/build-manifest.mjs              → dist/manifest.xml (prod URLs)
 *   node scripts/build-manifest.mjs --dev        → dist/manifest.xml (localhost)
 */
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const isDev = process.argv.includes('--dev');

const PROD_HOST = 'https://addin.novomcp.com';
const DEV_HOST = 'https://localhost:3000';

const src = readFileSync(resolve(root, 'manifest.xml'), 'utf8');
const out = isDev ? src.replaceAll(PROD_HOST, DEV_HOST) : src;

mkdirSync(resolve(root, 'dist'), { recursive: true });
writeFileSync(resolve(root, 'dist/manifest.xml'), out);
console.log(`✓ manifest written to dist/manifest.xml (${isDev ? 'dev' : 'prod'} URLs)`);
