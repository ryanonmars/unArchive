// Set the app version everywhere it is recorded, e.g. `npm run version:set -- 0.1.1`.
// The release workflow runs this with the version from the pushed tag.
import fs from 'node:fs';
import path from 'node:path';

const version = process.argv[2];
if (!/^\d+\.\d+\.\d+$/.test(version || '')) {
  console.error('Usage: node scripts/set-version.mjs <major.minor.patch>');
  process.exit(1);
}
const root = path.resolve(import.meta.dirname, '..');

function edit(file, change) {
  const target = path.join(root, file);
  const before = fs.readFileSync(target, 'utf8');
  const after = change(before);
  if (after === before && !before.includes(version)) throw new Error(`Could not find the version in ${file}.`);
  fs.writeFileSync(target, after);
}

// JSON files keep their formatting; only the top-level version values change.
edit('package.json', text => text.replace(/^(\s*"version":\s*")[^"]*(")/m, `$1${version}$2`));
edit('package-lock.json', text => text
  .replace(/^(\s{2}"version":\s*")[^"]*(")/m, `$1${version}$2`)
  .replace(/("packages":\s*\{\s*"":\s*\{[^}]*?"version":\s*")[^"]*(")/, `$1${version}$2`));
edit('src-tauri/tauri.conf.json', text => text.replace(/^(\s*"version":\s*")[^"]*(")/m, `$1${version}$2`));
edit('src-tauri/Cargo.toml', text => text.replace(/(\[package\][^[]*?\nversion\s*=\s*")[^"]*(")/, `$1${version}$2`));
edit('src-tauri/Cargo.lock', text => text.replace(/(\nname = "unarchive"\r?\nversion = ")[^"]*(")/, `$1${version}$2`));
console.log(`Version set to ${version}.`);
