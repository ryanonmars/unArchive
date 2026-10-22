// Prints the CHANGELOG.md section for one version, for the release workflow:
//   node scripts/release-notes.mjs 0.2.3 > release-notes.md
// Fails when the version has no section or the section is empty, so a release
// cannot be published without its notes.
import fs from 'node:fs';
import path from 'node:path';
import { releaseNotesFor } from '../release-notes.js';

const version = (process.argv[2] || '').replace(/^v/i, '');
if (!/^\d+\.\d+\.\d+$/.test(version)) {
  console.error('Usage: node scripts/release-notes.mjs <version>, for example 0.2.3');
  process.exit(2);
}
const changelog = fs.readFileSync(path.join(import.meta.dirname, '..', 'CHANGELOG.md'), 'utf8');
const release = releaseNotesFor(changelog, version);
if (!release || !release.notes) {
  console.error(`CHANGELOG.md has no notes for ${version}. Rename the Unreleased section to "## [${version}] - YYYY-MM-DD" before tagging.`);
  process.exit(1);
}
process.stdout.write(`${release.notes}\n`);
