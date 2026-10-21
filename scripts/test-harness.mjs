// Loads app.js in a linkedom DOM with the Tauri dialog and file plugins
// stubbed, so a test can export a workbook from CPS XML and import it back.
// app.js runs once per process, so each test suite runs in its own process.
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import * as XLSX from 'xlsx';
import { DOMParser, Event, parseHTML } from 'linkedom';

export const root = path.resolve(import.meta.dirname, '..');
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

let exportedWorkbook;
window.__TAURI_INTERNALS__ = {
  async invoke(command, args) {
    if (command === 'plugin:dialog|save') return 'C:\\temporary\\unarchive-test.xlsx';
    if (command === 'plugin:fs|write_file') {
      exportedWorkbook = new Uint8Array(args);
      return null;
    }
    return null;
  },
  transformCallback() { return 1; },
  unregisterCallback() {}
};

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

export async function importWorkbook(workbook) {
  status.textContent = '';
  const data = new Uint8Array(XLSX.write(workbook, { type: 'array', bookType: 'xlsx' }));
  Object.defineProperty(workbookPicker, 'files', { configurable: true, value: [{ data }] });
  workbookPicker.dispatchEvent(new Event('change'));
  await waitFor(() => status.textContent, 'workbook import');
  return { message: status.textContent, ok: status.className.includes('ok'), xml: output.value };
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
