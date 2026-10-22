// Loads app.js in a linkedom DOM with the Tauri dialog and file plugins
// stubbed, so a test can export a workbook from CPS XML and import it back.
// app.js runs once per process, so each test suite runs in its own process.
import fs from 'node:fs';
import { register } from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import * as XLSX from 'xlsx';
import { DOMParser, Event, parseHTML } from 'linkedom';

export const root = path.resolve(import.meta.dirname, '..');

// Vite bundles "file?raw" imports (the app's CHANGELOG.md) as text; teach
// Node the same, so app.js loads here unchanged.
register(`data:text/javascript,${encodeURIComponent(`
  export async function resolve(specifier, context, next) {
    if (specifier.endsWith('?raw')) return { url: new URL(specifier, context.parentURL).href, shortCircuit: true };
    return next(specifier, context);
  }
  export async function load(url, context, next) {
    if (!url.endsWith('?raw')) return next(url, context);
    const { readFile } = await import('node:fs/promises');
    const text = await readFile(new URL(url.slice(0, -'?raw'.length)), 'utf8');
    return { format: 'module', source: 'export default ' + JSON.stringify(text), shortCircuit: true };
  }
`)}`);
const { window, document } = parseHTML(fs.readFileSync(path.join(root, 'index.html'), 'utf8'));

// linkedom returns escaped attribute entities for XML nodes while a browser DOM
// returns their decoded value. Normalize the test DOM to browser behavior.
const nativeGetAttribute = window.Element.prototype.getAttribute;
window.Element.prototype.getAttribute = function getAttribute(name) {
  const value = nativeGetAttribute.call(this, name);
  return typeof value === 'string' ? value.replaceAll('&amp;', '&') : value;
};

document.implementation = {
  createDocument() {
    const doc = new DOMParser().parseFromString('<temporary/>', 'application/xml');
    doc.documentElement.remove();
    return doc;
  }
};

// The Tauri bridge. Every call is recorded in tauri.calls; a test answers a
// command by setting tauri.handlers[command] (throwing makes the call fail).
// The app runs as the desktop app: it reads its version from package.json,
// builds its menu, and makes its quiet startup update check, which finds no
// update unless a test says otherwise.
export const tauri = {
  calls: [],
  handlers: {
    'plugin:app|version': () => JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8')).version,
    'plugin:menu|new': () => [tauri.nextResource++, 'menu']
  },
  callbacks: new Map(), nextCallback: 1, nextResource: 1, channelIndexes: new Map()
};
// A test process can set how the startup update check and menu behave, since
// both run once when app.js loads: UNARCHIVE_TEST_STARTUP_UPDATE is "fail"
// (offline) or a version to offer, and UNARCHIVE_TEST_NO_MENU makes the menu
// unavailable.
const startupUpdate = process.env.UNARCHIVE_TEST_STARTUP_UPDATE;
if (startupUpdate === 'fail') tauri.handlers['plugin:updater|check'] = () => { throw new Error('offline'); };
else if (startupUpdate) tauri.handlers['plugin:updater|check'] = () => ({ rid: 900, currentVersion: tauri.handlers['plugin:app|version'](), version: startupUpdate, date: '', body: ['### Fixed', '- A startup test fix'].join('\n'), rawJson: {} });
if (process.env.UNARCHIVE_TEST_NO_MENU) tauri.handlers['plugin:menu|new'] = () => { throw new Error('menus are unavailable'); };
let exportedWorkbook;
window.__TAURI_INTERNALS__ = {
  async invoke(command, args) {
    tauri.calls.push({ command, args });
    if (command === 'plugin:dialog|save') return 'C:\\temporary\\unarchive-test.xlsx';
    if (command === 'plugin:fs|write_file') {
      exportedWorkbook = new Uint8Array(args);
      return null;
    }
    return tauri.handlers[command] ? tauri.handlers[command](args) : null;
  },
  transformCallback(callback) { const id = tauri.nextCallback++; tauri.callbacks.set(id, callback); return id; },
  unregisterCallback(id) { tauri.callbacks.delete(id); }
};
// Delivers a message on a Tauri channel, as the desktop app does when a menu
// item is clicked.
export function sendToChannel(channel, message) {
  const index = tauri.channelIndexes.get(channel.id) || 0;
  tauri.channelIndexes.set(channel.id, index + 1);
  tauri.callbacks.get(channel.id)({ index, message });
}

Object.assign(globalThis, {
  window,
  document,
  DOMParser,
  Event,
  Node: window.Node,
  XMLSerializer: class XMLSerializer {
    serializeToString(node) { return node.toString(); }
  },
  FileReader: class FileReader {
    readAsArrayBuffer(file) {
      this.result = file.data.buffer.slice(file.data.byteOffset, file.data.byteOffset + file.data.byteLength);
      queueMicrotask(() => this.onload?.());
    }
  }
});

await import(`${pathToFileURL(path.join(root, 'app.js'))}?safety-test=${Date.now()}`);

export { window, document, DOMParser, Event };
const status = document.getElementById('output-status');
const output = document.getElementById('xml-output');
const workbookPicker = document.getElementById('workbook-file');

export async function waitFor(predicate, description, timeout = 6000) {
  const started = Date.now();
  while (!predicate()) {
    if (Date.now() - started > timeout) throw new Error(`Timed out waiting for ${description}.`);
    await new Promise(resolve => setTimeout(resolve, 25));
  }
}

export async function exportWorkbookFrom(sourceXml) {
  exportedWorkbook = undefined;
  const input = document.getElementById('xml-input');
  input.value = sourceXml;
  input.dispatchEvent(new Event('change'));
  document.getElementById('export-workbook').dispatchEvent(new Event('click'));
  await waitFor(() => exportedWorkbook?.length, `workbook export (${document.getElementById('xml-status').textContent})`);
  return exportedWorkbook;
}

export function workbookFrom(bytes) { return XLSX.read(bytes, { type: 'array', raw: false }); }
export function rowsFor(workbook, sheet) { return XLSX.utils.sheet_to_json(workbook.Sheets[sheet], { header: 1, raw: false, defval: '' }); }
export function replaceRows(workbook, sheet, rows) { workbook.Sheets[sheet] = XLSX.utils.aoa_to_sheet(rows); }
export function appendRow(workbook, sheet, values) {
  const rows = rowsFor(workbook, sheet);
  rows.push(rows[0].map(header => values[header] ?? ''));
  replaceRows(workbook, sheet, rows);
}

// createAllNew sets the "Create all channels as new" option for this import
// only, so one test cannot leak it into the next.
export async function importWorkbook(workbook, { createAllNew = false } = {}) {
  status.textContent = '';
  const createAllNewBox = document.getElementById('import-as-new');
  if (createAllNewBox) createAllNewBox.checked = createAllNew;
  const data = new Uint8Array(XLSX.write(workbook, { type: 'array', bookType: 'xlsx' }));
  Object.defineProperty(workbookPicker, 'files', { configurable: true, value: [{ data }] });
  workbookPicker.dispatchEvent(new Event('change'));
  await waitFor(() => status.textContent, 'workbook import');
  if (createAllNewBox) createAllNewBox.checked = false;
  const summaryItems = [...(document.getElementById('import-summary')?.querySelectorAll('li') || [])];
  return {
    message: status.textContent, ok: status.className.includes('ok'), xml: output.value,
    summary: summaryItems.map(item => item.textContent), notes: summaryItems.filter(item => item.className === 'note').map(item => item.textContent)
  };
}

export function xmlDocument(xml) { return new DOMParser().parseFromString(xml, 'application/xml'); }
export function directElements(node, tag) { return [...node.children].filter(child => !tag || child.tagName === tag); }
export function category(doc, name) { return directElements(doc.documentElement, 'category').find(node => node.getAttribute('name') === name); }
export function directSet(parent, alias) { return directElements(parent, 'set').find(node => node.getAttribute('alias') === alias); }
export function activeCollections(contact) {
  return directElements(contact, 'collection').filter(collection => directElements(collection, 'set').length).map(collection => collection.getAttribute('name'));
}
export function allCollections(contact) { return directElements(contact, 'collection').map(collection => collection.getAttribute('name')); }
export function allChannels(doc) {
  return directElements(category(doc, 'Zone'), 'set').flatMap(zone =>
    directElements(zone, 'collection')
      .filter(collection => collection.getAttribute('name') === 'ZoneItems')
      .flatMap(collection => directElements(collection, 'set'))
  );
}
export function fieldOf(node, name) { return directElements(node, 'field').find(item => item.getAttribute('name') === name); }
export function channelType(channel) {
  return directElements(channel, 'field').find(field => field.getAttribute('name') === 'CP_PERSTYPE')?.textContent || '';
}
