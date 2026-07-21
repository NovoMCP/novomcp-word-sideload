/**
 * Pre-start hook: remove any stale sideload manifest left in Word's
 * wef directory by a previous failed `office-addin-debugging start`.
 *
 * office-addin-debugging registers a sideload by creating a hardlink:
 *   <Word wef dir>/<addin GUID>.manifest.xml -> dist/manifest.xml
 *
 * If the previous run failed mid-cycle, the link survives and the next
 * run hits EEXIST. Removing the stale link before starting is safe —
 * `npm run start` will recreate it from the current dist/manifest.xml
 * the user has just rebuilt.
 */
import { readFileSync, unlinkSync, existsSync } from 'node:fs';
import { homedir } from 'node:os';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

// Pull the GUID out of the source manifest so we know which file to
// remove — the GUID is the filename Word uses inside its wef dir.
let guid;
try {
  const manifest = readFileSync(resolve(root, 'manifest.xml'), 'utf8');
  const m = manifest.match(/<Id>([^<]+)<\/Id>/);
  if (!m) {
    console.error('clean-sideload: could not find <Id> in manifest.xml');
    process.exit(0); // non-fatal; let `start` proceed and surface its own error
  }
  guid = m[1].toLowerCase();
} catch (e) {
  console.error('clean-sideload: could not read manifest.xml:', e.message);
  process.exit(0);
}

// Word's sideload directory differs by platform.
const wefDirs = [];
if (process.platform === 'darwin') {
  wefDirs.push(resolve(homedir(), 'Library/Containers/com.microsoft.Word/Data/Documents/wef'));
} else if (process.platform === 'win32') {
  // %USERPROFILE%/AppData/Roaming/Microsoft/Office/16.0/Wef on Windows
  const userProfile = process.env.USERPROFILE;
  if (userProfile) {
    wefDirs.push(resolve(userProfile, 'AppData/Roaming/Microsoft/Office/16.0/Wef'));
  }
}

let removed = 0;
for (const dir of wefDirs) {
  const candidate = resolve(dir, `${guid}.manifest.xml`);
  if (existsSync(candidate)) {
    try {
      unlinkSync(candidate);
      removed++;
      console.log(`clean-sideload: removed stale ${candidate}`);
    } catch (e) {
      console.warn(`clean-sideload: could not remove ${candidate}: ${e.message}`);
    }
  }
}

if (removed === 0) {
  // Nothing to clean — quiet success.
}
