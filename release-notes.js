// Release notes and version helpers shared by the app, the release workflow
// (scripts/release-notes.mjs), and the tests. CHANGELOG.md is the only place
// release notes are written; each release publishes its own section.

// Compares two versions such as 0.10.0 and v0.9.2 part by part. Returns a
// negative number, zero, or a positive number. Missing parts count as 0.
export function compareVersions(a, b) {
  const parts = value => String(value ?? '').trim().replace(/^v/i, '').split(/[.+-]/).slice(0, 3).map(part => Number.parseInt(part, 10) || 0);
  const [x, y] = [parts(a), parts(b)];
  for (let index = 0; index < 3; index++) {
    const difference = (x[index] || 0) - (y[index] || 0);
    if (difference) return difference;
  }
  return 0;
}

// Splits CHANGELOG.md into releases, newest first as written. Each has its
// version ('Unreleased' for work not yet released), date, and notes: the
// Markdown between its heading and the next one.
export function parseChangelog(markdown) {
  const releases = [];
  let current = null;
  for (const line of String(markdown ?? '').replace(/\r\n?/g, '\n').split('\n')) {
    const heading = /^##\s+\[?([^\]\s]+)\]?(?:\s*-\s*(\d{4}-\d{2}-\d{2}))?\s*$/.exec(line);
    if (heading) {
      current = { version: heading[1], date: heading[2] || '', lines: [] };
      releases.push(current);
    } else if (current) current.lines.push(line);
  }
  return releases.map(({ version, date, lines }) => ({ version, date, notes: lines.join('\n').trim() }));
}

export function releaseNotesFor(markdown, version) {
  return parseChangelog(markdown).find(release => release.version !== 'Unreleased' && compareVersions(release.version, version) === 0) || null;
}

// The installed version and every earlier release, newest first. Unreleased
// work and versions newer than the installed one are left out.
export function releaseHistory(markdown, installedVersion) {
  return parseChangelog(markdown)
    .filter(release => release.version !== 'Unreleased' && /^\d/.test(release.version) && (!installedVersion || compareVersions(release.version, installedVersion) <= 0))
    .sort((a, b) => compareVersions(b.version, a.version));
}

// Releases published before CHANGELOG.md existed carry only "unArchive
// vX.Y.Z" as their update notes, which says nothing about the release.
export function hasRealNotes(notes) {
  const text = String(notes ?? '').trim();
  return Boolean(text) && !/^unArchive\s+v?[\d.]+$/i.test(text);
}

// Turns release notes into plain blocks for display: headings ("### Added"),
// bullet items ("- ..."), and paragraphs. Markdown emphasis and links are
// shown as their text. Nothing is ever inserted as HTML.
export function noteBlocks(notes) {
  const plain = value => value.replace(/\*\*(.+?)\*\*/g, '$1').replace(/`([^`]+)`/g, '$1').replace(/\[([^\]]+)\]\([^)]*\)/g, '$1').trim();
  const blocks = [];
  for (const raw of String(notes ?? '').replace(/\r\n?/g, '\n').split('\n')) {
    const line = raw.trim();
    if (!line) continue;
    const heading = /^#{1,6}\s+(.*)$/.exec(line);
    const item = /^[-*]\s+(.*)$/.exec(line);
    if (heading) blocks.push({ kind: 'heading', text: plain(heading[1]) });
    else if (item) {
      const last = blocks[blocks.length - 1];
      if (last?.kind === 'list') last.items.push(plain(item[1])); else blocks.push({ kind: 'list', items: [plain(item[1])] });
    } else blocks.push({ kind: 'paragraph', text: plain(line) });
  }
  return blocks;
}
