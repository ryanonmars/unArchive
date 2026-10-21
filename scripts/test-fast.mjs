// Fast regression suite. It runs on a synthetic codeplug built from the
// committed blueprints, so it needs no customer data and is tracked in Git.
// The full suite against real CPS fixtures stays local (see README).
import assert from 'node:assert/strict';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { CPS_BLUEPRINTS } from '../blueprints.generated.js';
import {
  document, exportWorkbookFrom, workbookFrom, rowsFor, replaceRows, appendRow, importWorkbook as harnessImport,
  xmlDocument, directElements, category, allChannels, fieldOf, channelType
} from './test-harness.mjs';
import { syntheticCodeplug } from './test-fixture.mjs';

// The import summary is for people, so it never shows CPS's internal
// field names, record types, category names, or enum codes. Every import in
// this suite is checked.
const INTERNAL_TEXT = [
  /\b[A-Z][A-Z0-9]*_[A-Z0-9_]+\b/,
  /\b(ANLGCONV|DGTLCONV6PT25|CAPPLUSSYSCONV|LINKEDCAPPLUSSYSCONV|CAPPLUSVOICECONV|CHANNELPOOL|ScanItems|RoamList|PCRContacts|CapacityPlus\w*)\b/,
  /\b[A-Z0-9]{13,}\b/,
  /Channel — |Off Set|OffSet/
];
function assertPlainSummary(result) {
  for (const line of result.summary || []) for (const pattern of INTERNAL_TEXT) assert.doesNotMatch(line, pattern, `Summary shows internal text: ${line}`);
  return result;
}
const importWorkbook = async (...args) => assertPlainSummary(await harnessImport(...args));

const sourceXml = syntheticCodeplug();
const source = xmlDocument(sourceXml);
const bytes = await exportWorkbookFrom(sourceXml);

// Rebuilt records may list attributes in another order, which XML ignores.
const canonical = node => ({ tag: node.tagName, attrs: Object.fromEntries([...node.attributes].map(attr => [attr.name, attr.value]).sort()), children: node.children.length ? directElements(node).map(canonical) : node.textContent });
const shape = node => directElements(node).map(item => `${item.tagName}:${item.getAttribute('name')}`).join('|');
const blueprintShape = type => shape(xmlDocument(`<r>${CPS_BLUEPRINTS.channels[type]}</r>`).documentElement.firstElementChild);
const place = channel => `${channel.parentElement.parentElement.getAttribute('key') === 'CHANNELPOOL' ? 'POOL' : channel.parentElement.parentElement.getAttribute('alias')}|${channel.getAttribute('alias')}`;
const channelIn = (doc, alias) => allChannels(doc).find(item => item.getAttribute('alias') === alias);
const sourceChannel = alias => channelIn(source, alias);
const poolOf = doc => directElements(category(doc, 'Zone'), 'set').find(zone => zone.getAttribute('key') === 'CHANNELPOOL');
const poolChannels = doc => directElements(poolOf(doc), 'collection').flatMap(collection => directElements(collection, 'set'));
const value = (channel, name) => fieldOf(channel, name)?.textContent;
const sheetXml = async (workbookBytes, sheet) => {
  const workbook = workbookFrom(workbookBytes);
  return (await JSZip.loadAsync(workbookBytes)).file(`xl/worksheets/sheet${workbook.SheetNames.indexOf(sheet) + 1}.xml`).async('string');
};
const column = (headers, header) => XLSX.utils.encode_col(headers.indexOf(header));
// Changes cells by channel name: { 'DIG TALK': { 'Receive Only': '' } }.
function edit(workbook, sheet, changes) {
  const rows = rowsFor(workbook, sheet);
  const name = rows[0].indexOf('Name / Alias');
  for (const [alias, cells] of Object.entries(changes)) {
    const row = rows.find((item, index) => index && item[name] === alias);
    assert.ok(row, `${sheet} has no row ${alias}`);
    for (const [header, cell] of Object.entries(cells)) row[rows[0].indexOf(header)] = cell;
  }
  replaceRows(workbook, sheet, rows);
  return workbook;
}
async function imported(workbook) {
  const result = await importWorkbook(workbook);
  assert.equal(result.ok, true, result.message);
  return xmlDocument(result.xml);
}
async function rejected(workbook, pattern) {
  const result = await importWorkbook(workbook);
  assert.equal(result.ok, false, `Expected import to stop with ${pattern}`);
  assert.equal(result.xml, '');
  assert.match(result.message, pattern);
}
const fresh = () => workbookFrom(bytes);

// A receive-only channel needs no contact, as in CPS: new Digital, Capacity
// Plus, and Linked Capacity Plus channels with Receive Only True and no
// contact import; the same channels able to transmit still need one.
{
  for (const [type, extra] of [['Digital', { 'Receive Frequency': '461.4000', 'Transmit Frequency': '461.4000', 'Color Code': '1' }], ['Capacity Plus', { 'Voice List': 'VOICE A' }], ['Linked Capacity Plus', { 'Site List': 'SITES A' }]]) {
    const listening = fresh();
    appendRow(listening, 'Zone Channels', { 'Assigned Zone': 'ZONE A', 'Name / Alias': 'LISTEN ONLY', 'Channel Type': type, 'Receive Only': 'True', ...extra });
    const doc = await imported(listening);
    assert.deepEqual([value(channelIn(doc, 'LISTEN ONLY'), 'CP_RXONLYEN'), value(channelIn(doc, 'LISTEN ONLY'), 'CP_UKPPERS')], ['True', 'NONE'], `${type} receive-only without a contact`);
    const talking = fresh();
    appendRow(talking, 'Zone Channels', { 'Assigned Zone': 'ZONE A', 'Name / Alias': 'TALK NO CONTACT', 'Channel Type': type, ...extra });
    await rejected(talking, new RegExp(`TALK NO CONTACT: ${type} needs a Contact Name`));
  }
}

// "Create all channels as new" starts unticked, and says what it resets.
{
  const box = document.getElementById('import-as-new');
  assert.equal(box.type, 'checkbox');
  // No checked attribute, so the page loads with it off. (linkedom leaves the
  // checked property undefined until set; a browser reports false.)
  assert.equal(box.hasAttribute('checked'), false);
  assert.equal(Boolean(box.checked), false);
  assert.equal(box.closest('label').textContent.trim(), 'Create all channels as new');
  assert.match(box.closest('label').nextElementSibling.textContent, /every other channel setting, such as time slot, transmit power, RX group list, emergency system, and PL tones, resets to the CPS template default/);
}

// READ FIRST tells people to clear settings a changed channel type lacks.
assert.match(rowsFor(fresh(), 'READ FIRST').flat().join(' '), /Clear any setting the new type does not have, such as Privacy Enabled and the key aliases when a Digital channel becomes Analog/);

// An unchanged round trip rebuilds every channel, contact, list, and scan list
// record exactly, so no field is added, dropped, or emptied.
{
  const doc = await imported(fresh());
  for (const name of ['PCRContacts', 'CapacityPlusVoiceList', 'CapacityPlusSiteList', 'ScanItems', 'RoamList']) {
    assert.deepEqual(directElements(category(doc, name)).map(canonical), directElements(category(source, name)).map(canonical), `${name} round-trips exactly.`);
  }
  const rebuilt = new Map(allChannels(doc).map(channel => [place(channel), channel]));
  assert.equal(rebuilt.size, allChannels(source).length);
  for (const original of allChannels(source)) assert.deepEqual(canonical(rebuilt.get(place(original))), canonical(original), `${place(original)} round-trips exactly.`);
}

// Channel Pool drop-downs: Assigned Zone, Channel Type, Contact Name, Receive
// Only, and Scan List. They cover the whole column, so an empty pool keeps them.
for (const pool of [true, false]) {
  const workbookBytes = pool ? bytes : await exportWorkbookFrom(syntheticCodeplug({ pool: false }));
  const headers = rowsFor(workbookFrom(workbookBytes), 'Channel Pool')[0];
  assert.deepEqual(headers.slice(0, 4), ['Assigned Zone', 'Name / Alias', 'Contact Name', 'Channel Type']);
  if (!pool) assert.equal(rowsFor(workbookFrom(workbookBytes), 'Channel Pool').length, 1);
  const xml = await sheetXml(workbookBytes, 'Channel Pool');
  for (const [header, formula] of [['Assigned Zone', '&quot;Channel Pool&quot;'], ['Channel Type', '&quot;Capacity Plus Voice,Digital&quot;'], ['Receive Only', '&quot;True,False&quot;'], ['Contact Name', 'unArchiveContactNames'], ['Scan List', 'unArchiveScanListNames']]) {
    const col = column(headers, header);
    assert.match(xml, new RegExp(`sqref="${col}2:${col}5000"><formula1>${formula}</formula1>`), `Channel Pool ${header} drop-down (pool ${pool ? 'present' : 'empty'})`);
  }
}
{
  const xml = await sheetXml(bytes, 'Zone Channels');
  const headers = rowsFor(fresh(), 'Zone Channels')[0];
  for (const [header, formula] of [['Receive Only', '&quot;True,False&quot;'], ['Scan List', 'unArchiveScanListNames']]) {
    const col = column(headers, header);
    assert.match(xml, new RegExp(`sqref="${col}2:${col}5000"><formula1>${formula}</formula1>`), `Zone Channels ${header} drop-down`);
  }
  const zip = await JSZip.loadAsync(bytes);
  assert.match(await zip.file('xl/workbook.xml').async('string'), /name="unArchiveScanListNames">&apos;_CPS_CHOICES&apos;!\$E\$2:\$E\$4</);
  assert.deepEqual(rowsFor(fresh(), '_CPS_CHOICES').map(row => row[4]).filter(Boolean), ['Scan List Names', 'None', 'SCAN A', 'SCAN B']);
}

// Channel Pool rows import into the pool whatever Assigned Zone says, and a
// Digital pool channel keeps its contact, such as a Digital All Call.
{
  const pool = rowsFor(fresh(), 'Channel Pool');
  const cell = (alias, header) => pool.find(row => row[1] === alias)[pool[0].indexOf(header)];
  assert.equal(cell('POOL ALL CALL', 'Contact Name'), 'DIG ALL CALL');
  assert.equal(cell('VOICE 1', 'Contact Name'), '');

  const workbook = fresh();
  appendRow(workbook, 'Channel Pool', { 'Assigned Zone': 'Channel Pool', 'Name / Alias': 'NEW POOL DIG', 'Contact Name': 'DIG 1', 'Channel Type': 'Digital', 'Receive Frequency': '463.4000', 'Transmit Frequency': '468.4000', 'Color Code': '1' });
  appendRow(workbook, 'Channel Pool', { 'Name / Alias': 'NEW POOL NONE', 'Channel Type': 'Digital', 'Receive Frequency': '463.5000', 'Transmit Frequency': '468.5000', 'Color Code': '1' });
  appendRow(workbook, 'Channel Pool', { 'Name / Alias': 'NEW VOICE', 'Channel Type': 'Capacity Plus Voice', 'Receive Frequency': '463.6000', 'Transmit Frequency': '468.6000', 'Color Code': '1' });
  edit(workbook, 'Channel Pool', { 'POOL ALL CALL': { 'Contact Name': 'DIG 2' } });
  const doc = await imported(workbook);
  assert.deepEqual(poolChannels(doc).map(item => item.getAttribute('alias')), ['VOICE 1', 'VOICE 2', 'POOL ALL CALL', 'NEW POOL DIG', 'NEW POOL NONE', 'NEW VOICE']);
  assert.deepEqual([fieldOf(channelIn(doc, 'NEW POOL DIG'), 'CP_UKPPERS').textContent, fieldOf(channelIn(doc, 'NEW POOL DIG'), 'CP_UKPPERS').getAttribute('Name')], ['DIG 1', 'DIG 1']);
  assert.equal(value(channelIn(doc, 'NEW POOL NONE'), 'CP_UKPPERS'), 'NONE', 'A blank pool contact means none.');
  assert.equal(value(channelIn(doc, 'POOL ALL CALL'), 'CP_UKPPERS'), 'DIG 2');
  assert.equal(shape(channelIn(doc, 'NEW POOL DIG')), blueprintShape('DGTLCONV6PT25'));
  assert.equal(shape(channelIn(doc, 'NEW VOICE')), blueprintShape('CAPPLUSVOICECONV'));

  await rejected(edit(fresh(), 'Channel Pool', { 'POOL ALL CALL': { 'Contact Name': 'CP 1' } }), /POOL ALL CALL: Digital needs a Digital contact/);
  await rejected(edit(fresh(), 'Channel Pool', { 'VOICE 1': { 'Contact Name': 'DIG 1' } }), /VOICE 1: Capacity Plus Voice channels have no Contact Name in CPS/);
  const newVoice = fresh();
  appendRow(newVoice, 'Channel Pool', { 'Name / Alias': 'NEW VOICE', 'Contact Name': 'DIG 1', 'Channel Type': 'Capacity Plus Voice', 'Receive Frequency': '463.6000', 'Transmit Frequency': '468.6000', 'Color Code': '1' });
  await rejected(newVoice, /NEW VOICE: Capacity Plus Voice channels have no Contact Name in CPS/);
}

// Scan List is CPS's Scan/Roam List (CP_SCNROAMLISTIT): ScanItems/<name> or
// NONE, on Analog and Digital channels only.
{
  const zone = rowsFor(fresh(), 'Zone Channels');
  const cell = (alias, header) => zone.find(row => row[1] === alias)[zone[0].indexOf(header)];
  assert.deepEqual(['DIG RXONLY', 'DIG TALK', 'ANA RXONLY', 'ANA TALK', 'CP TALK', 'LCP TALK'].map(alias => cell(alias, 'Scan List')), ['SCAN A', 'None', 'SCAN B', 'None', '', '']);

  const workbook = edit(fresh(), 'Zone Channels', {
    'DIG RXONLY': { 'Scan List': 'None' }, 'DIG TALK': { 'Scan List': 'scan b' }, 'ANA RXONLY': { 'Scan List': '' }, 'ANA TALK': { 'Scan List': 'SCAN A' }
  });
  appendRow(workbook, 'Zone Channels', { 'Assigned Zone': 'ZONE A', 'Name / Alias': 'NEW DIG SCAN', 'Contact Name': 'DIG 1', 'Channel Type': 'Digital', 'Receive Frequency': '461.5000', 'Transmit Frequency': '461.5000', 'Color Code': '1', 'Scan List': 'SCAN B' });
  appendRow(workbook, 'Zone Channels', { 'Assigned Zone': 'ZONE A', 'Name / Alias': 'NEW ANA', 'Channel Type': 'Analog', 'Receive Frequency': '462.6000', 'Transmit Frequency': '462.6000' });
  const doc = await imported(workbook);
  const scan = alias => { const field = fieldOf(channelIn(doc, alias), 'CP_SCNROAMLISTIT'); return [field.textContent, field.getAttribute('Name')]; };
  assert.deepEqual(scan('DIG RXONLY'), ['NONE', 'None']);
  assert.deepEqual(scan('DIG TALK'), ['ScanItems/SCAN B', 'ScanItems/SCAN B']);
  assert.deepEqual(scan('ANA RXONLY'), ['NONE', 'None']);
  assert.deepEqual(scan('ANA TALK'), ['ScanItems/SCAN A', 'ScanItems/SCAN A']);
  assert.deepEqual(scan('NEW DIG SCAN'), ['ScanItems/SCAN B', 'ScanItems/SCAN B']);
  assert.deepEqual(scan('NEW ANA'), ['NONE', 'None']);
  for (const alias of ['DIG RXONLY', 'DIG TALK', 'ANA RXONLY', 'ANA TALK']) assert.equal(shape(channelIn(doc, alias)), shape(sourceChannel(alias)), `${alias} keeps its CPS fields.`);
  assert.deepEqual(directElements(category(doc, 'ScanItems')).map(canonical), directElements(category(source, 'ScanItems')).map(canonical), 'Scan list records are untouched.');

  await rejected(edit(fresh(), 'Zone Channels', { 'DIG TALK': { 'Scan List': 'SCAN Z' } }), /DIG TALK: Scan List "SCAN Z" is not a scan list in this codeplug/);
  await rejected(edit(fresh(), 'Zone Channels', { 'CP TALK': { 'Scan List': 'SCAN A' } }), /CP TALK: Capacity Plus channels have no Scan List in CPS/);
  await rejected(edit(fresh(), 'Zone Channels', { 'LCP TALK': { 'Scan List': 'SCAN B' } }), /LCP TALK: Linked Capacity Plus channels have no Scan List in CPS/);
  // None means no scan list, so it is accepted there and adds no field.
  const noneDoc = await imported(edit(fresh(), 'Zone Channels', { 'LCP TALK': { 'Scan List': 'None' } }));
  assert.equal(shape(channelIn(noneDoc, 'LCP TALK')), shape(sourceChannel('LCP TALK')));
  await rejected(edit(fresh(), 'Channel Pool', { 'VOICE 1': { 'Scan List': 'SCAN A' } }), /VOICE 1: Capacity Plus Voice channels have no Scan List in CPS/);
  await rejected(edit(fresh(), 'Zone Channels', { 'DIG RXONLY': { 'Channel Type': 'Capacity Plus', 'Contact Name': 'CP 1', 'Voice List': 'VOICE A', 'Receive Only': '' } }), /DIG RXONLY: Capacity Plus channels have no Scan List in CPS/);

  // A stored reference that is not a scan list, such as a roam list, is kept
  // while unchanged but cannot be typed in.
  const roamBytes = await exportWorkbookFrom(syntheticCodeplug({ roamListChannel: true }));
  const roamRows = rowsFor(workbookFrom(roamBytes), 'Zone Channels');
  assert.equal(roamRows.find(row => row[1] === 'DIG TALK')[roamRows[0].indexOf('Scan List')], 'RoamList/ROAM A');
  const roamDoc = await imported(workbookFrom(roamBytes));
  assert.deepEqual([value(channelIn(roamDoc, 'DIG TALK'), 'CP_SCNROAMLISTIT'), fieldOf(channelIn(roamDoc, 'DIG TALK'), 'CP_SCNROAMLISTIT').getAttribute('Name')], ['RoamList/ROAM A', 'RoamList/ROAM A']);
  await rejected(edit(workbookFrom(roamBytes), 'Zone Channels', { 'DIG RXONLY': { 'Scan List': 'RoamList/ROAM A' } }), /DIG RXONLY: Scan List "RoamList\/ROAM A" is not a scan list/);
}

// Receive Only (CP_RXONLYEN) is a True/False field on Analog, Digital, Capacity
// Plus, and Linked Capacity Plus channels. Capacity Plus Voice has none.
{
  const supported = { DGTLCONV6PT25: 'DIG', ANLGCONV: 'ANA', CAPPLUSSYSCONV: 'CP', LINKEDCAPPLUSSYSCONV: 'LCP' };
  for (const type of Object.keys(CPS_BLUEPRINTS.channels)) {
    assert.equal(blueprintShape(type).includes('field:CP_RXONLYEN'), type in supported, `${type} Receive Only support`);
  }
  const zone = rowsFor(fresh(), 'Zone Channels'), pool = rowsFor(fresh(), 'Channel Pool');
  const cell = (rows, alias) => rows.find(row => row[1] === alias)[rows[0].indexOf('Receive Only')];
  for (const prefix of Object.values(supported)) {
    assert.equal(cell(zone, `${prefix} RXONLY`), 'True');
    assert.equal(cell(zone, `${prefix} TALK`), 'False');
  }
  assert.equal(cell(pool, 'VOICE 1'), '', 'Capacity Plus Voice exports a blank Receive Only.');
  assert.equal(cell(pool, 'POOL ALL CALL'), 'True');

  // True and False, in any letter case, are written in CPS's own form.
  const flipped = edit(fresh(), 'Zone Channels', Object.fromEntries(Object.values(supported).flatMap(prefix => [[`${prefix} RXONLY`, { 'Receive Only': 'false' }], [`${prefix} TALK`, { 'Receive Only': 'TRUE' }]])));
  edit(flipped, 'Channel Pool', { 'POOL ALL CALL': { 'Receive Only': 'False' } });
  const flippedDoc = await imported(flipped);
  for (const prefix of Object.values(supported)) {
    assert.equal(value(channelIn(flippedDoc, `${prefix} RXONLY`), 'CP_RXONLYEN'), 'False');
    assert.equal(value(channelIn(flippedDoc, `${prefix} TALK`), 'CP_RXONLYEN'), 'True');
    for (const alias of [`${prefix} RXONLY`, `${prefix} TALK`]) assert.equal(shape(channelIn(flippedDoc, alias)), shape(sourceChannel(alias)), `${alias} keeps its CPS fields.`);
  }
  assert.equal(value(channelIn(flippedDoc, 'POOL ALL CALL'), 'CP_RXONLYEN'), 'False');

  // A blank cell never writes an empty Boolean: an existing channel keeps its
  // stored value, and a new channel or a changed type takes the CPS default.
  const blank = edit(fresh(), 'Zone Channels', Object.fromEntries(Object.values(supported).flatMap(prefix => [[`${prefix} RXONLY`, { 'Receive Only': '' }], [`${prefix} TALK`, { 'Receive Only': '' }]])));
  edit(blank, 'Channel Pool', { 'POOL ALL CALL': { 'Receive Only': '' } });
  appendRow(blank, 'Zone Channels', { 'Assigned Zone': 'ZONE A', 'Name / Alias': 'NEW DIG', 'Contact Name': 'DIG 1', 'Channel Type': 'Digital', 'Receive Frequency': '461.6000', 'Transmit Frequency': '461.6000', 'Color Code': '1' });
  appendRow(blank, 'Zone Channels', { 'Assigned Zone': 'ZONE A', 'Name / Alias': 'NEW ANA', 'Channel Type': 'Analog', 'Receive Frequency': '462.6000', 'Transmit Frequency': '462.6000' });
  appendRow(blank, 'Zone Channels', { 'Assigned Zone': 'ZONE A', 'Name / Alias': 'NEW CP', 'Contact Name': 'CP 1', 'Channel Type': 'Capacity Plus', 'Voice List': 'VOICE A' });
  appendRow(blank, 'Zone Channels', { 'Assigned Zone': 'ZONE A', 'Name / Alias': 'NEW LCP', 'Contact Name': 'CP 1', 'Channel Type': 'Linked Capacity Plus', 'Site List': 'SITES A' });
  const blankDoc = await imported(blank);
  for (const prefix of Object.values(supported)) {
    assert.equal(value(channelIn(blankDoc, `${prefix} RXONLY`), 'CP_RXONLYEN'), 'True', `${prefix} RXONLY keeps True when cleared.`);
    assert.equal(value(channelIn(blankDoc, `${prefix} TALK`), 'CP_RXONLYEN'), 'False', `${prefix} TALK keeps False when cleared.`);
  }
  assert.equal(value(channelIn(blankDoc, 'POOL ALL CALL'), 'CP_RXONLYEN'), 'True');
  for (const [alias, type] of [['NEW DIG', 'DGTLCONV6PT25'], ['NEW ANA', 'ANLGCONV'], ['NEW CP', 'CAPPLUSSYSCONV'], ['NEW LCP', 'LINKEDCAPPLUSSYSCONV']]) {
    assert.equal(value(channelIn(blankDoc, alias), 'CP_RXONLYEN'), 'False', `${alias} takes the CPS default.`);
    assert.equal(shape(channelIn(blankDoc, alias)), blueprintShape(type));
  }
  const changedType = await imported(edit(fresh(), 'Zone Channels', {
    'DIG RXONLY': { 'Channel Type': 'Analog', 'Contact Name': '', 'Scan List': '', 'Receive Only': '', 'Privacy Enabled': '' },
    'ANA RXONLY': { 'Channel Type': 'Digital', 'Contact Name': 'DIG 1', 'Color Code': '1', 'Scan List': '', 'Receive Only': 'True' }
  }));
  assert.equal(value(channelIn(changedType, 'DIG RXONLY'), 'CP_RXONLYEN'), 'False', 'A changed type takes the new type\'s default.');
  assert.equal(value(channelIn(changedType, 'ANA RXONLY'), 'CP_RXONLYEN'), 'True');
  assert.equal(shape(channelIn(changedType, 'DIG RXONLY')), blueprintShape('ANLGCONV'));
  assert.equal(shape(channelIn(changedType, 'ANA RXONLY')), blueprintShape('DGTLCONV6PT25'));

  // Anything other than True or False is refused rather than written.
  for (const entry of ['Yes', '1', 'on']) {
    await rejected(edit(fresh(), 'Zone Channels', { 'CP TALK': { 'Receive Only': entry } }), new RegExp(`CP TALK: Receive Only must be True or False, not "${entry}"`));
  }
  const newBad = fresh();
  appendRow(newBad, 'Zone Channels', { 'Assigned Zone': 'ZONE A', 'Name / Alias': 'NEW BAD', 'Channel Type': 'Analog', 'Receive Frequency': '462.6000', 'Transmit Frequency': '462.6000', 'Receive Only': 'Y' });
  await rejected(newBad, /NEW BAD: Receive Only must be True or False, not "Y"/);

  // Capacity Plus Voice: any value, even False, is refused for existing,
  // new, and converted channels; blank adds no field.
  for (const entry of ['False', 'True', 'false']) {
    await rejected(edit(fresh(), 'Channel Pool', { 'VOICE 1': { 'Receive Only': entry } }), /VOICE 1: Capacity Plus Voice channels have no Receive Only setting in CPS\. Leave Receive Only blank\./);
    const added = fresh();
    appendRow(added, 'Channel Pool', { 'Name / Alias': 'NEW VOICE', 'Channel Type': 'Capacity Plus Voice', 'Receive Frequency': '463.6000', 'Transmit Frequency': '468.6000', 'Color Code': '1', 'Receive Only': entry });
    await rejected(added, /NEW VOICE: Capacity Plus Voice channels have no Receive Only setting in CPS/);
    await rejected(edit(fresh(), 'Channel Pool', { 'POOL ALL CALL': { 'Channel Type': 'Capacity Plus Voice', 'Contact Name': '', 'Scan List': '', 'Privacy Enabled': '', 'Privacy Key Alias': '', 'RAS Key Alias': '', 'Receive Only': entry } }), /POOL ALL CALL: Capacity Plus Voice channels have no Receive Only setting in CPS/);
  }
  const converted = await imported(edit(fresh(), 'Channel Pool', { 'POOL ALL CALL': { 'Channel Type': 'Capacity Plus Voice', 'Contact Name': '', 'Scan List': '', 'Privacy Enabled': '', 'Privacy Key Alias': '', 'RAS Key Alias': '', 'Receive Only': '' } }));
  assert.equal(channelType(channelIn(converted, 'POOL ALL CALL')), 'CAPPLUSVOICECONV');
  assert.equal(fieldOf(channelIn(converted, 'POOL ALL CALL'), 'CP_RXONLYEN'), undefined);
  assert.equal(shape(channelIn(converted, 'POOL ALL CALL')), blueprintShape('CAPPLUSVOICECONV'), 'The converted channel has exactly the CPS template fields.');
  for (const alias of ['VOICE 1', 'VOICE 2']) assert.equal(fieldOf(channelIn(converted, alias), 'CP_RXONLYEN'), undefined);
}

// Privacy Enabled (CP_VOICEPRIVACYEN) and Auto Roam (CP_LCPALATRM) are True/False
// fields like Receive Only, on fewer channel types: Privacy Enabled on Digital,
// Capacity Plus, and Linked Capacity Plus (default False), and Auto Roam on
// Linked Capacity Plus only (default True).
{
  const settings = [
    { header: 'Privacy Enabled', field: 'CP_VOICEPRIVACYEN', types: ['DGTLCONV6PT25', 'CAPPLUSSYSCONV', 'LINKEDCAPPLUSSYSCONV'], fallback: 'False' },
    { header: 'Auto Roam', field: 'CP_LCPALATRM', types: ['LINKEDCAPPLUSSYSCONV'], fallback: 'True' }
  ];
  const typeLabels = { DGTLCONV6PT25: 'Digital', ANLGCONV: 'Analog', CAPPLUSSYSCONV: 'Capacity Plus', LINKEDCAPPLUSSYSCONV: 'Linked Capacity Plus', CAPPLUSVOICECONV: 'Capacity Plus Voice' };
  const prefixes = { DGTLCONV6PT25: 'DIG', ANLGCONV: 'ANA', CAPPLUSSYSCONV: 'CP', LINKEDCAPPLUSSYSCONV: 'LCP' };
  const newRow = {
    DGTLCONV6PT25: { 'Contact Name': 'DIG 1', 'Channel Type': 'Digital', 'Receive Frequency': '461.6000', 'Transmit Frequency': '461.6000', 'Color Code': '1' },
    ANLGCONV: { 'Channel Type': 'Analog', 'Receive Frequency': '462.6000', 'Transmit Frequency': '462.6000' },
    CAPPLUSSYSCONV: { 'Contact Name': 'CP 1', 'Channel Type': 'Capacity Plus', 'Voice List': 'VOICE A' },
    LINKEDCAPPLUSSYSCONV: { 'Contact Name': 'CP 1', 'Channel Type': 'Linked Capacity Plus', 'Site List': 'SITES A' }
  };
  const zone = rowsFor(fresh(), 'Zone Channels');
  const cell = (alias, header) => zone.find(row => row[1] === alias)[zone[0].indexOf(header)];
  assert.ok(!rowsFor(fresh(), 'Channel Pool')[0].includes('Auto Roam'), 'Auto Roam is not on the Channel Pool tab.');

  for (const { header, field, types, fallback } of settings) {
    for (const type of Object.keys(CPS_BLUEPRINTS.channels)) {
      assert.equal(blueprintShape(type).includes(`field:${field}`), types.includes(type), `${type} ${header} support`);
      if (types.includes(type)) assert.equal(value(xmlDocument(`<r>${CPS_BLUEPRINTS.channels[type]}</r>`).documentElement.firstElementChild, field), fallback, `${type} ${header} default`);
    }
    const headers = rowsFor(fresh(), 'Zone Channels')[0];
    assert.match(await sheetXml(bytes, 'Zone Channels'), new RegExp(`sqref="${column(headers, header)}2:${column(headers, header)}5000"><formula1>&quot;True,False&quot;</formula1>`), `${header} drop-down`);
    const supported = types.filter(type => type in prefixes).map(type => prefixes[type]);
    const unsupported = Object.keys(prefixes).filter(type => !types.includes(type));

    // Export shows the stored value on supported types and a blank elsewhere.
    for (const prefix of supported) assert.equal(cell(`${prefix} TALK`, header), value(sourceChannel(`${prefix} TALK`), field));
    for (const type of unsupported) assert.equal(cell(`${prefixes[type]} TALK`, header), '', `${typeLabels[type]} exports a blank ${header}.`);

    // Any letter case is written in CPS's form; a blank keeps the stored value.
    const opposite = stored => stored === 'True' ? 'False' : 'True';
    const flipped = await imported(edit(fresh(), 'Zone Channels', Object.fromEntries(supported.flatMap(prefix => [
      [`${prefix} TALK`, { [header]: opposite(value(sourceChannel(`${prefix} TALK`), field)).toLowerCase() }],
      [`${prefix} RXONLY`, { [header]: '' }]
    ]))));
    for (const prefix of supported) {
      assert.equal(value(channelIn(flipped, `${prefix} TALK`), field), opposite(value(sourceChannel(`${prefix} TALK`), field)), `${prefix} TALK ${header} in CPS form`);
      assert.equal(value(channelIn(flipped, `${prefix} RXONLY`), field), value(sourceChannel(`${prefix} RXONLY`), field), `${prefix} RXONLY keeps its ${header} when cleared`);
      for (const alias of [`${prefix} TALK`, `${prefix} RXONLY`]) assert.equal(shape(channelIn(flipped, alias)), shape(sourceChannel(alias)), `${alias} keeps its CPS fields.`);
    }

    // New channels: blank takes the CPS default, an entered value is kept.
    const added = fresh();
    for (const type of types) {
      appendRow(added, 'Zone Channels', { 'Assigned Zone': 'ZONE A', 'Name / Alias': `NEW BLANK ${prefixes[type]}`, ...newRow[type] });
      appendRow(added, 'Zone Channels', { 'Assigned Zone': 'ZONE A', 'Name / Alias': `NEW SET ${prefixes[type]}`, ...newRow[type], [header]: opposite(fallback) });
    }
    const addedDoc = await imported(added);
    for (const type of types) {
      assert.equal(value(channelIn(addedDoc, `NEW BLANK ${prefixes[type]}`), field), fallback);
      assert.equal(value(channelIn(addedDoc, `NEW SET ${prefixes[type]}`), field), opposite(fallback));
      for (const alias of [`NEW BLANK ${prefixes[type]}`, `NEW SET ${prefixes[type]}`]) assert.equal(shape(channelIn(addedDoc, alias)), blueprintShape(type));
    }

    // Values other than True or False are refused.
    await rejected(edit(fresh(), 'Zone Channels', { [`${supported[0]} TALK`]: { [header]: 'Yes' } }), new RegExp(`${supported[0]} TALK: ${header} must be True or False, not "Yes"`));

    // Types without the field take no value, even False: existing, new, and
    // converted channels alike, and a blank adds no field.
    const missing = new RegExp(`channels have no ${header} setting in CPS\\. Leave ${header} blank\\.`);
    for (const type of unsupported) {
      for (const entry of ['True', 'False', 'false']) {
        await rejected(edit(fresh(), 'Zone Channels', { [`${prefixes[type]} TALK`]: { [header]: entry } }), new RegExp(`${prefixes[type]} TALK: ${typeLabels[type]} ${missing.source}`));
      }
      const addedBad = fresh();
      appendRow(addedBad, 'Zone Channels', { 'Assigned Zone': 'ZONE A', 'Name / Alias': 'NEW BAD', ...newRow[type], [header]: 'False' });
      await rejected(addedBad, new RegExp(`NEW BAD: ${typeLabels[type]} ${missing.source}`));
      const blankDoc = await imported(edit(fresh(), 'Zone Channels', { [`${prefixes[type]} TALK`]: { [header]: '' } }));
      assert.equal(fieldOf(channelIn(blankDoc, `${prefixes[type]} TALK`), field), undefined);
    }
  }

  // Changing a Linked Capacity Plus channel to Digital: its exported Auto Roam
  // must be cleared, while Privacy Enabled carries over because Digital has it.
  const toDigital = { 'Channel Type': 'Digital', 'Contact Name': 'DIG 1', 'Site List': '', 'Receive Frequency': '461.8000', 'Transmit Frequency': '461.8000', 'Color Code': '1' };
  await rejected(edit(fresh(), 'Zone Channels', { 'LCP TALK': toDigital }), /LCP TALK: Digital channels have no Auto Roam setting in CPS/);
  const converted = await imported(edit(fresh(), 'Zone Channels', { 'LCP TALK': { ...toDigital, 'Auto Roam': '', 'Privacy Enabled': 'True' } }));
  assert.equal(fieldOf(channelIn(converted, 'LCP TALK'), 'CP_LCPALATRM'), undefined);
  assert.equal(value(channelIn(converted, 'LCP TALK'), 'CP_VOICEPRIVACYEN'), 'True');
  assert.equal(shape(channelIn(converted, 'LCP TALK')), blueprintShape('DGTLCONV6PT25'));
  // And back: a Digital channel changed to Linked Capacity Plus takes its default Auto Roam.
  const toLinked = await imported(edit(fresh(), 'Zone Channels', { 'DIG TALK': { 'Channel Type': 'Linked Capacity Plus', 'Contact Name': 'CP 1', 'Site List': 'SITES A', 'Scan List': '' } }));
  assert.equal(value(channelIn(toLinked, 'DIG TALK'), 'CP_LCPALATRM'), 'True');
  assert.equal(shape(channelIn(toLinked, 'DIG TALK')), blueprintShape('LINKEDCAPPLUSSYSCONV'));
}

// Privacy Key Alias (CP_ENHCPRVCYIT) and RAS Key Alias (CP_RASDATAITEM) are
// key references on Digital, Capacity Plus, and Linked Capacity Plus only.
// Changing between those types keeps them; a type without them takes only a
// blank or None, so a reference is never silently dropped.
{
  const keyedXml = syntheticCodeplug({ keys: true });
  const keyedSource = xmlDocument(keyedXml);
  const keyedBytes = await exportWorkbookFrom(keyedXml);
  const keyed = () => workbookFrom(keyedBytes);
  const fields = { 'Privacy Key Alias': 'CP_ENHCPRVCYIT', 'RAS Key Alias': 'CP_RASDATAITEM' };
  const keyedChannel = alias => allChannels(keyedSource).find(item => item.getAttribute('alias') === alias);
  for (const name of Object.values(fields)) {
    for (const type of Object.keys(CPS_BLUEPRINTS.channels)) {
      assert.equal(blueprintShape(type).includes(`field:${name}`), ['DGTLCONV6PT25', 'CAPPLUSSYSCONV', 'LINKEDCAPPLUSSYSCONV'].includes(type), `${type} ${name} support`);
    }
  }
  const zone = rowsFor(keyed(), 'Zone Channels');
  const cell = (alias, header) => zone.find(row => row[1] === alias)[zone[0].indexOf(header)];
  assert.deepEqual([cell('DIG TALK', 'Privacy Key Alias'), cell('DIG TALK', 'RAS Key Alias'), cell('ANA TALK', 'Privacy Key Alias'), cell('ANA TALK', 'RAS Key Alias')], ['PRIV A', 'RAS A', '', '']);
  assert.deepEqual(rowsFor(keyed(), 'Security').slice(1).map(row => row.slice(0, 2)), [['RAS A', 'PRIV A']]);

  // Unchanged import keeps every key reference and its label exactly.
  const same = await imported(keyed());
  for (const original of allChannels(keyedSource)) assert.deepEqual(canonical(channelIn(same, original.getAttribute('alias'))), canonical(original), `${original.getAttribute('alias')} round-trips exactly.`);

  // Changing between types that have the fields keeps both references.
  const moved = await imported(edit(keyed(), 'Zone Channels', {
    'DIG TALK': { 'Channel Type': 'Capacity Plus', 'Contact Name': 'CP 1', 'Voice List': 'VOICE A', 'Scan List': '' },
    'CP TALK': { 'Channel Type': 'Linked Capacity Plus', 'Voice List': '', 'Site List': 'SITES A' },
    'LCP TALK': { 'Channel Type': 'Digital', 'Contact Name': 'DIG 1', 'Site List': '', 'Auto Roam': '', 'Receive Frequency': '461.9000', 'Transmit Frequency': '461.9000', 'Color Code': '1' }
  }));
  for (const [alias, type] of [['DIG TALK', 'CAPPLUSSYSCONV'], ['CP TALK', 'LINKEDCAPPLUSSYSCONV'], ['LCP TALK', 'DGTLCONV6PT25']]) {
    const channel = channelIn(moved, alias);
    assert.equal(channelType(channel), type);
    for (const name of Object.values(fields)) assert.deepEqual([value(channel, name), fieldOf(channel, name).getAttribute('Name')], [value(keyedChannel(alias), name), fieldOf(keyedChannel(alias), name).getAttribute('Name')], `${alias} keeps ${name}`);
    assert.equal(shape(channel), blueprintShape(type));
  }

  // Changing to Analog with a key alias left in is refused, naming the field;
  // cleared or None, the channel becomes a plain Analog record.
  const toAnalog = { 'Channel Type': 'Analog', 'Contact Name': '', 'Color Code': '', 'Privacy Enabled': '', 'Scan List': '' };
  for (const alias of ['DIG TALK', 'LCP TALK']) {
    const extra = alias === 'LCP TALK' ? { 'Site List': '', 'Auto Roam': '', 'Receive Frequency': '462.8000', 'Transmit Frequency': '462.8000' } : {};
    await rejected(edit(keyed(), 'Zone Channels', { [alias]: { ...toAnalog, ...extra } }), new RegExp(`${alias}: Analog channels have no Privacy Key Alias in CPS\\. Leave Privacy Key Alias blank or None\\.`));
    await rejected(edit(keyed(), 'Zone Channels', { [alias]: { ...toAnalog, ...extra, 'Privacy Key Alias': '' } }), new RegExp(`${alias}: Analog channels have no RAS Key Alias in CPS`));
    for (const cleared of ['', 'NONE', 'None']) {
      const doc = await imported(edit(keyed(), 'Zone Channels', { [alias]: { ...toAnalog, ...extra, 'Privacy Key Alias': cleared, 'RAS Key Alias': cleared } }));
      assert.equal(channelType(channelIn(doc, alias)), 'ANLGCONV');
      assert.equal(shape(channelIn(doc, alias)), blueprintShape('ANLGCONV'), `${alias} as Analog has exactly the CPS template fields.`);
    }
  }

  // Existing and new Analog channels take no key alias; None adds no field.
  for (const header of Object.keys(fields)) {
    const alias = header === 'Privacy Key Alias' ? 'PRIV A' : 'RAS A';
    await rejected(edit(keyed(), 'Zone Channels', { 'ANA TALK': { [header]: alias } }), new RegExp(`ANA TALK: Analog channels have no ${header} in CPS`));
    const added = keyed();
    appendRow(added, 'Zone Channels', { 'Assigned Zone': 'ZONE A', 'Name / Alias': 'NEW ANA KEY', 'Channel Type': 'Analog', 'Receive Frequency': '462.6000', 'Transmit Frequency': '462.6000', [header]: alias });
    await rejected(added, new RegExp(`NEW ANA KEY: Analog channels have no ${header} in CPS`));
  }
  const noneDoc = await imported(edit(keyed(), 'Zone Channels', { 'ANA TALK': { 'Privacy Key Alias': 'NONE', 'RAS Key Alias': 'None' } }));
  assert.equal(shape(channelIn(noneDoc, 'ANA TALK')), shape(keyedChannel('ANA TALK')));

  // Supported types: clearing a key writes NONE with CPS's label, and an
  // alias missing from the Security sheet is still refused.
  const cleared = await imported(edit(keyed(), 'Zone Channels', { 'CP TALK': { 'Privacy Key Alias': '', 'RAS Key Alias': '' } }));
  for (const name of Object.values(fields)) assert.deepEqual([value(channelIn(cleared, 'CP TALK'), name), fieldOf(channelIn(cleared, 'CP TALK'), name).getAttribute('Name')], ['NONE', 'None']);
  await rejected(edit(keyed(), 'Zone Channels', { 'DIG TALK': { 'Privacy Key Alias': 'PRIV Z' } }), /DIG TALK: Privacy Key Alias "PRIV Z" is not listed on the Security sheet/);
  const fromAnalog = await imported(edit(keyed(), 'Zone Channels', { 'ANA TALK': { 'Channel Type': 'Digital', 'Contact Name': 'DIG 1', 'Color Code': '1', 'Scan List': '', 'Privacy Key Alias': 'PRIV A', 'RAS Key Alias': 'RAS A' } }));
  assert.deepEqual(Object.values(fields).map(name => value(channelIn(fromAnalog, 'ANA TALK'), name)), ['PRIV A', 'RAS A']);
  assert.equal(shape(channelIn(fromAnalog, 'ANA TALK')), blueprintShape('DGTLCONV6PT25'));
}

// The Channel Pool tab shows Privacy Enabled, Privacy Key Alias, and RAS Key
// Alias, which a Digital pool channel has exactly as a Digital zone channel
// does. Capacity Plus Voice has none of them.
{
  const SECURITY = ['Privacy Enabled', 'Privacy Key Alias', 'RAS Key Alias'];
  const FIELDS = { 'Privacy Enabled': 'CP_VOICEPRIVACYEN', 'Privacy Key Alias': 'CP_ENHCPRVCYIT', 'RAS Key Alias': 'CP_RASDATAITEM' };
  const poolXml = syntheticCodeplug({ poolKeys: true });
  const poolSource = xmlDocument(poolXml);
  const poolBytes = await exportWorkbookFrom(poolXml);
  const book = () => workbookFrom(poolBytes);
  const security = (doc, alias) => Object.values(FIELDS).map(name => value(channelIn(doc, alias), name));
  const labels = (doc, alias) => ['CP_ENHCPRVCYIT', 'CP_RASDATAITEM'].map(name => fieldOf(channelIn(doc, alias), name).getAttribute('Name'));
  const digitalTemplate = xmlDocument(`<r>${CPS_BLUEPRINTS.channels.DGTLCONV6PT25}</r>`).documentElement.firstElementChild;
  const templateSecurity = Object.values(FIELDS).map(name => value(digitalTemplate, name));
  assert.deepEqual(templateSecurity, ['False', 'NONE', 'NONE'], 'The CPS Digital template has privacy off and no keys.');

  // Columns sit after Channel Type with the Zone Channels drop-downs.
  const headers = rowsFor(book(), 'Channel Pool')[0];
  assert.deepEqual(headers.slice(0, 7), ['Assigned Zone', 'Name / Alias', 'Contact Name', 'Channel Type', ...SECURITY]);
  assert.equal(headers[1], 'Name / Alias', 'Voice List members still read pool names from column B.');
  const xml = await sheetXml(poolBytes, 'Channel Pool');
  for (const [header, formula, strict] of [['Privacy Enabled', '&quot;True,False&quot;', true], ['Privacy Key Alias', 'unArchivePrivacyKeyAliases', true], ['RAS Key Alias', 'unArchiveRASKeyAliases', true]]) {
    const col = column(headers, header);
    assert.match(xml, new RegExp(`showErrorMessage="${strict ? 1 : 0}" sqref="${col}2:${col}5000"><formula1>${formula}</formula1>`), `Channel Pool ${header} drop-down`);
  }

  // Export shows a pool channel's actual settings, including keys set in CPS
  // while privacy is off; Capacity Plus Voice rows stay blank.
  const pool = rowsFor(book(), 'Channel Pool');
  const cells = alias => SECURITY.map(header => pool.find(row => row[1] === alias)[pool[0].indexOf(header)]);
  assert.deepEqual(cells('POOL ALL CALL'), ['False', 'PRIV A', 'RAS A']);
  assert.deepEqual(cells('VOICE 1'), ['', '', '']);
  assert.deepEqual(rowsFor(book(), 'Security').slice(1).map(row => row.slice(0, 2)), [['RAS A', 'PRIV A']], 'Pool keys are listed on the Security sheet.');

  // An unchanged import keeps every pool and zone record exactly.
  const same = await imported(book());
  for (const original of allChannels(poolSource)) assert.deepEqual(canonical(channelIn(same, original.getAttribute('alias'))), canonical(original), `${original.getAttribute('alias')} round-trips exactly.`);

  // The CH 1 / CH 2 case: an existing pool channel renamed to CH 1 keeps its
  // record and so its keys, which the tab now shows; a new CH 2 with the same
  // visible settings gets the CPS template's privacy off and no keys.
  const twin = book();
  edit(twin, 'Channel Pool', { 'POOL ALL CALL': { 'Name / Alias': 'CH 1', 'Contact Name': 'DIG 1', 'Receive Frequency': '463.8000', 'Transmit Frequency': '468.8000', 'Color Code': '1', 'Receive Only': 'False', 'Scan List': 'None' } });
  appendRow(twin, 'Channel Pool', { 'Assigned Zone': 'Channel Pool', 'Name / Alias': 'CH 2', 'Contact Name': 'DIG 1', 'Channel Type': 'Digital', 'Receive Frequency': '463.8000', 'Transmit Frequency': '468.8000', 'Color Code': '1', 'Receive Only': 'False', 'Scan List': 'None' });
  const twinDoc = await imported(twin);
  assert.deepEqual(security(twinDoc, 'CH 1'), ['False', 'PRIV A', 'RAS A'], 'CH 1 is the existing record and keeps its keys.');
  assert.deepEqual(security(twinDoc, 'CH 2'), templateSecurity, 'CH 2 is new and inherits nothing.');
  assert.deepEqual(labels(twinDoc, 'CH 2'), ['None', 'None']);
  assert.equal(shape(channelIn(twinDoc, 'CH 2')), blueprintShape('DGTLCONV6PT25'));
  const twinBack = rowsFor(workbookFrom(await exportWorkbookFrom(twinDoc.documentElement.toString())), 'Channel Pool');
  const reexported = alias => SECURITY.map(header => twinBack.find(row => row[1] === alias)[twinBack[0].indexOf(header)]);
  assert.deepEqual(reexported('CH 1'), ['False', 'PRIV A', 'RAS A'], 'Re-exported, CH 1 shows the keys it carries.');
  assert.deepEqual(reexported('CH 2'), ['False', 'NONE', 'NONE']);

  // Clearing the keys on the existing channel writes NONE with CPS's label;
  // choosing keys and privacy on a new channel writes them.
  const cleared = book();
  edit(cleared, 'Channel Pool', { 'POOL ALL CALL': { 'Privacy Key Alias': '', 'RAS Key Alias': 'None' } });
  appendRow(cleared, 'Channel Pool', { 'Name / Alias': 'CH 3', 'Channel Type': 'Digital', 'Receive Frequency': '463.9000', 'Transmit Frequency': '468.9000', 'Color Code': '1', 'Privacy Enabled': 'true', 'Privacy Key Alias': 'PRIV A', 'RAS Key Alias': 'RAS A' });
  const clearedDoc = await imported(cleared);
  assert.deepEqual(security(clearedDoc, 'POOL ALL CALL'), ['False', 'NONE', 'NONE']);
  assert.deepEqual(labels(clearedDoc, 'POOL ALL CALL'), ['None', 'None']);
  assert.deepEqual(security(clearedDoc, 'CH 3'), ['True', 'PRIV A', 'RAS A']);
  assert.deepEqual(labels(clearedDoc, 'CH 3'), ['PRIV A', 'RAS A']);
  for (const alias of ['POOL ALL CALL', 'CH 3']) assert.equal(shape(channelIn(clearedDoc, alias)), blueprintShape('DGTLCONV6PT25'));
  // A blank Privacy Enabled keeps the stored value, as on Zone Channels.
  const privacyOn = await imported(edit(book(), 'Channel Pool', { 'POOL ALL CALL': { 'Privacy Enabled': 'True' } }));
  assert.equal(value(channelIn(privacyOn, 'POOL ALL CALL'), 'CP_VOICEPRIVACYEN'), 'True');
  const keptBlank = await imported(edit(workbookFrom(await exportWorkbookFrom(privacyOn.documentElement.toString())), 'Channel Pool', { 'POOL ALL CALL': { 'Privacy Enabled': '' } }));
  assert.equal(value(channelIn(keptBlank, 'POOL ALL CALL'), 'CP_VOICEPRIVACYEN'), 'True');

  // Validation matches Zone Channels.
  await rejected(edit(book(), 'Channel Pool', { 'POOL ALL CALL': { 'Privacy Key Alias': 'PRIV Z' } }), /POOL ALL CALL: Privacy Key Alias "PRIV Z" is not listed on the Security sheet/);
  await rejected(edit(book(), 'Channel Pool', { 'POOL ALL CALL': { 'Privacy Enabled': 'Yes' } }), /POOL ALL CALL: Privacy Enabled must be True or False, not "Yes"/);

  // Capacity Plus Voice: no value, even False; None for a key adds nothing.
  for (const [header, entry, pattern] of [
    ['Privacy Enabled', 'False', /VOICE 1: Capacity Plus Voice channels have no Privacy Enabled setting in CPS\. Leave Privacy Enabled blank\./],
    ['Privacy Enabled', 'True', /VOICE 1: Capacity Plus Voice channels have no Privacy Enabled setting in CPS/],
    ['Privacy Key Alias', 'PRIV A', /VOICE 1: Capacity Plus Voice channels have no Privacy Key Alias in CPS\. Leave Privacy Key Alias blank or None\./],
    ['RAS Key Alias', 'RAS A', /VOICE 1: Capacity Plus Voice channels have no RAS Key Alias in CPS/]
  ]) {
    await rejected(edit(book(), 'Channel Pool', { 'VOICE 1': { [header]: entry } }), pattern);
    const added = book();
    appendRow(added, 'Channel Pool', { 'Name / Alias': 'VOICE 1 NEW', 'Channel Type': 'Capacity Plus Voice', 'Receive Frequency': '463.1500', 'Transmit Frequency': '468.1500', 'Color Code': '3', [header]: entry });
    await rejected(added, new RegExp(pattern.source.replace('VOICE 1', 'VOICE 1 NEW')));
  }
  const voiceNone = await imported(edit(book(), 'Channel Pool', { 'VOICE 1': { 'Privacy Key Alias': 'None', 'RAS Key Alias': 'NONE' } }));
  assert.deepEqual(canonical(channelIn(voiceNone, 'VOICE 1')), canonical(channelIn(poolSource, 'VOICE 1')));
  // Changing the keyed Digital pool channel to Capacity Plus Voice needs its
  // security cells cleared, like any setting the new type lacks.
  const toVoice = { 'Channel Type': 'Capacity Plus Voice', 'Contact Name': '', 'Receive Only': '', 'Scan List': '' };
  await rejected(edit(book(), 'Channel Pool', { 'POOL ALL CALL': { ...toVoice, 'Privacy Enabled': '' } }), /POOL ALL CALL: Capacity Plus Voice channels have no Privacy Key Alias in CPS/);
  const voiced = await imported(edit(book(), 'Channel Pool', { 'POOL ALL CALL': { ...toVoice, 'Privacy Enabled': '', 'Privacy Key Alias': '', 'RAS Key Alias': '' } }));
  assert.equal(shape(channelIn(voiced, 'POOL ALL CALL')), blueprintShape('CAPPLUSVOICECONV'));

  // A workbook exported before these columns existed keeps the stored
  // security settings, and its new Digital rows still get the CPS defaults.
  const old = book();
  const oldRows = rowsFor(old, 'Channel Pool');
  const drop = SECURITY.map(header => oldRows[0].indexOf(header));
  replaceRows(old, 'Channel Pool', oldRows.map(row => row.filter((_, index) => !drop.includes(index))));
  const meta = rowsFor(old, '_CPS_META');
  const poolMeta = meta.find(row => row[0] === 'Channel Pool');
  const columnMap = JSON.parse(poolMeta[5]);
  SECURITY.forEach(header => delete columnMap[header]);
  poolMeta[5] = JSON.stringify(columnMap);
  replaceRows(old, '_CPS_META', meta);
  assert.ok(!rowsFor(old, 'Channel Pool')[0].includes('Privacy Key Alias'));
  appendRow(old, 'Channel Pool', { 'Name / Alias': 'OLD BOOK NEW', 'Channel Type': 'Digital', 'Receive Frequency': '463.9500', 'Transmit Frequency': '468.9500', 'Color Code': '1' });
  const oldDoc = await imported(old);
  assert.deepEqual(canonical(channelIn(oldDoc, 'POOL ALL CALL')), canonical(channelIn(poolSource, 'POOL ALL CALL')), 'An older workbook leaves the pool keys exactly as stored.');
  assert.deepEqual(security(oldDoc, 'OLD BOOK NEW'), templateSecurity);

  // None in any letter case is written as CPS's NONE on Zone Channels too.
  const zoneKeyed = await exportWorkbookFrom(syntheticCodeplug({ keys: true }));
  const zoneNone = await imported(edit(workbookFrom(zoneKeyed), 'Zone Channels', { 'DIG TALK': { 'Privacy Key Alias': 'none', 'RAS Key Alias': 'None' } }));
  assert.deepEqual(['CP_ENHCPRVCYIT', 'CP_RASDATAITEM'].map(name => [value(channelIn(zoneNone, 'DIG TALK'), name), fieldOf(channelIn(zoneNone, 'DIG TALK'), name).getAttribute('Name')]), [['NONE', 'None'], ['NONE', 'None']]);
}

// Channel record identity. A row with a hidden record ID is that channel; a
// row without one is a new channel from its CPS template, whatever its name.
// "Create all channels as new" builds every channel from its template.
{
  // Old channels carry a hidden marker the workbook never shows, so any
  // inheritance is visible: CP_TOT on most types, CP_SARQ_AT on Capacity
  // Plus Voice. DIG TALK also holds a hidden RX group list, and SCAN A lists
  // DIG TALK by name.
  const MARK = '987';
  const markerField = node => fieldOf(node, 'CP_TOT') ? 'CP_TOT' : 'CP_SARQ_AT';
  const markedDoc = xmlDocument(syntheticCodeplug({ keys: true }));
  for (const channel of allChannels(markedDoc)) fieldOf(channel, markerField(channel)).textContent = MARK;
  const tgList = fieldOf(allChannels(markedDoc).find(item => item.getAttribute('alias') === 'DIG TALK'), 'CP_TGLISTIT');
  tgList.textContent = 'GROUP A'; tgList.setAttribute('Name', 'GROUP A');
  const scanMembers = directElements(directElements(category(markedDoc, 'ScanItems'), 'set')[0], 'multifield')[0];
  const member = markedDoc.createElement('value'); member.textContent = 'DIG TALK'; scanMembers.append(member);
  const markedXml = '<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' + markedDoc.documentElement.toString();
  const marked = xmlDocument(markedXml);
  const markedBytes = await exportWorkbookFrom(markedXml);
  const book = () => workbookFrom(markedBytes);
  const markerOf = (doc, alias) => { const channel = channelIn(doc, alias); return value(channel, markerField(channel)); };
  const templateMarker = type => value(xmlDocument(`<r>${CPS_BLUEPRINTS.channels[type]}</r>`).documentElement.firstElementChild, type === 'CAPPLUSVOICECONV' ? 'CP_SARQ_AT' : 'CP_TOT');
  const recordId = (workbook, sheet, alias) => { const rows = rowsFor(workbook, sheet); return rows.find(row => row[1] === alias)[rows[0].indexOf('_record_id')]; };
  const setRows = (workbook, sheet, change) => { const rows = rowsFor(workbook, sheet); change(rows, Object.fromEntries(rows[0].map((header, index) => [header, index]))); replaceRows(workbook, sheet, rows); return workbook; };
  const run = async (workbook, options) => { const result = await importWorkbook(workbook, options); assert.equal(result.ok, true, result.message); return { ...result, doc: xmlDocument(result.xml) }; };
  const refuse = async (workbook, pattern, options) => { const result = await importWorkbook(workbook, options); assert.equal(result.ok, false, `Expected ${pattern}`); assert.equal(result.xml, ''); assert.match(result.message, pattern); assert.deepEqual(result.summary, []); return result; };
  // Everything outside the zones keeps its exact records.
  const othersUnchanged = (doc, label) => {
    for (const name of ['PCRContacts', 'CapacityPlusVoiceList', 'CapacityPlusSiteList', 'ScanItems', 'RoamList']) assert.deepEqual(directElements(category(doc, name)).map(canonical), directElements(category(marked, name)).map(canonical), `${label}: ${name} untouched`);
  };

  // An unchanged import keeps every record and says so.
  const same = await run(book());
  for (const original of allChannels(marked)) assert.deepEqual(canonical(channelIn(same.doc, original.getAttribute('alias'))), canonical(original));
  assert.ok(same.summary.includes('8 existing channels kept, with their other settings.') && same.summary.includes('3 existing channels kept, with their other settings.'), same.summary.join(' | '));
  assert.deepEqual(same.notes, []);

  // Same-name replacement: the old row deleted and a new one typed with the
  // same name, below or above the others. The new channel is built from the
  // template; nothing hidden carries over, the old record is removed.
  const newDigTalk = { 'Assigned Zone': 'ZONE A', 'Name / Alias': 'DIG TALK', 'Contact Name': 'DIG 2', 'Channel Type': 'Digital', 'Receive Frequency': '461.0250', 'Transmit Frequency': '466.0250', 'Color Code': '2' };
  for (const position of ['below', 'above']) {
    const workbook = setRows(book(), 'Zone Channels', (rows, col) => {
      rows.splice(rows.findIndex(row => row[1] === 'DIG TALK'), 1);
      const fresh = rows[0].map(header => newDigTalk[header] ?? '');
      if (position === 'above') rows.splice(1, 0, fresh); else rows.push(fresh);
    });
    const { doc, notes } = await run(workbook);
    assert.equal(markerOf(doc, 'DIG TALK'), templateMarker('DGTLCONV6PT25'), `${position}: the new DIG TALK does not inherit the hidden setting`);
    assert.deepEqual(['CP_ENHCPRVCYIT', 'CP_RASDATAITEM', 'CP_TGLISTIT', 'CP_VOICEPRIVACYEN'].map(name => value(channelIn(doc, 'DIG TALK'), name)), ['NONE', 'NONE', 'NONE', 'False']);
    assert.equal(shape(channelIn(doc, 'DIG TALK')), blueprintShape('DGTLCONV6PT25'));
    assert.equal(allChannels(doc).filter(item => item.getAttribute('alias') === 'DIG TALK').length, 1);
    assert.ok(notes.some(note => /replace a removed channel of the same name, without its other settings: DIG TALK/.test(note)), notes.join(' | '));
    othersUnchanged(doc, `same-name replacement ${position}`);
  }

  // A renamed channel keeps its record even when a new row reusing its old
  // name sits above it: row order no longer hands the record to the new row.
  const reordered = setRows(book(), 'Zone Channels', (rows, col) => {
    rows.find(row => row[1] === 'DIG TALK')[col['Name / Alias']] = 'DIG TALK 2';
    rows.splice(1, 0, rows[0].map(header => newDigTalk[header] ?? ''));
  });
  const reorderedDoc = (await run(reordered)).doc;
  assert.equal(markerOf(reorderedDoc, 'DIG TALK 2'), MARK, 'The renamed channel keeps its record.');
  assert.equal(value(channelIn(reorderedDoc, 'DIG TALK 2'), 'CP_TGLISTIT'), 'GROUP A');
  assert.equal(markerOf(reorderedDoc, 'DIG TALK'), templateMarker('DGTLCONV6PT25'), 'The new same-name row is a new channel.');

  // A deleted hidden ID makes the row a new channel, never a name match.
  const idCleared = setRows(book(), 'Zone Channels', (rows, col) => { rows.find(row => row[1] === 'CP TALK')[col._record_id] = ''; });
  const idClearedRun = await run(idCleared);
  assert.equal(markerOf(idClearedRun.doc, 'CP TALK'), templateMarker('CAPPLUSSYSCONV'));
  assert.deepEqual(['CP_ENHCPRVCYIT', 'CP_RASDATAITEM'].map(name => value(channelIn(idClearedRun.doc, 'CP TALK'), name)), ['PRIV A', 'RAS A'], 'Its visible keys come from the row.');
  assert.ok(idClearedRun.notes.some(note => /same name, without its other settings: CP TALK/.test(note)));

  // Retyped in another zone without an ID: new there, removed from ZONE A.
  const crossZone = book();
  appendRow(crossZone, 'Zones', { 'Name / Alias': 'ZONE B' });
  setRows(crossZone, 'Zone Channels', (rows, col) => { const row = rows.find(item => item[1] === 'DIG TALK'); row[col['Assigned Zone']] = 'ZONE B'; row[col._record_id] = ''; });
  const crossDoc = (await run(crossZone)).doc;
  assert.equal(channelIn(crossDoc, 'DIG TALK').parentElement.parentElement.getAttribute('alias'), 'ZONE B');
  assert.equal(markerOf(crossDoc, 'DIG TALK'), templateMarker('DGTLCONV6PT25'));

  // Legitimate edits keep the record: rename, zone move, and two channels
  // exchanging names. A type change rebuilds from the new type's template.
  const edits = book();
  appendRow(edits, 'Zones', { 'Name / Alias': 'ZONE B' });
  setRows(edits, 'Zone Channels', (rows, col) => {
    const row = alias => rows.find(item => item[1] === alias);
    row('ANA TALK')[col['Name / Alias']] = 'ANA RENAMED';
    row('DIG RXONLY')[col['Assigned Zone']] = 'ZONE B';
    const [rxOnly, talk] = [row('CP RXONLY'), row('CP TALK')];
    rxOnly[col['Name / Alias']] = 'CP TALK'; talk[col['Name / Alias']] = 'CP RXONLY';
    Object.assign(row('LCP TALK'), { [col['Channel Type']]: 'Digital', [col['Contact Name']]: 'DIG 1', [col['Site List']]: '', [col['Auto Roam']]: '', [col['Receive Frequency']]: '461.8000', [col['Transmit Frequency']]: '461.8000', [col['Color Code']]: '1' });
  });
  const editRun = await run(edits);
  for (const alias of ['ANA RENAMED', 'DIG RXONLY', 'CP TALK', 'CP RXONLY']) assert.equal(markerOf(editRun.doc, alias), MARK, `${alias} keeps its record`);
  assert.equal(value(channelIn(editRun.doc, 'CP TALK'), 'CP_RXONLYEN'), 'True', 'Exchanged names: CP TALK is the old CP RXONLY record.');
  assert.equal(channelIn(editRun.doc, 'DIG RXONLY').parentElement.parentElement.getAttribute('alias'), 'ZONE B');
  assert.equal(markerOf(editRun.doc, 'LCP TALK'), templateMarker('DGTLCONV6PT25'), 'A type change is rebuilt from the new template.');
  for (const expected of [/^Renamed: .*ANA TALK → ANA RENAMED/, /^Moved to another zone: DIG RXONLY \(ZONE A → ZONE B\)/, /^Changed type, rebuilt from the new type's CPS template: LCP TALK \(Linked Capacity Plus → Digital\)/]) assert.ok(editRun.summary.some(line => expected.test(line)), `${expected} in ${editRun.summary.join(' | ')}`);
  assert.ok(editRun.summary.includes('7 existing channels kept, with their other settings.'), `The retyped channel is not counted as kept as-is: ${editRun.summary.join(' | ')}`);
  assert.deepEqual(editRun.notes, [], 'Ordinary edits raise no notes.');

  // Swapped or misaligned IDs, as when the hidden column is left out of a
  // sort: provable, so refused. Rotating three rows' IDs is refused too.
  await refuse(setRows(book(), 'Zone Channels', (rows, col) => {
    const a = rows.find(row => row[1] === 'DIG TALK'), b = rows.find(row => row[1] === 'ANA TALK');
    [a[col._record_id], b[col._record_id]] = [b[col._record_id], a[col._record_id]];
  }), /Zone Channels: the row "DIG TALK" \(ZONE A\) exactly matches the exported row of "DIG TALK" \(ZONE A\), but its hidden record ID belongs to "ANA TALK" \(ZONE A\)\. The hidden ID column was probably left out of a sort or paste\./);
  await refuse(setRows(book(), 'Zone Channels', (rows, col) => {
    const ids = rows.slice(1, 4).map(row => row[col._record_id]);
    rows.slice(1, 4).forEach((row, index) => { row[col._record_id] = ids[(index + 1) % 3]; });
  }), /exactly matches the exported row of .* but its hidden record ID belongs to .*\(\+?|exactly matches the exported row/);
  await refuse(setRows(book(), 'Channel Pool', (rows, col) => {
    const a = rows.find(row => row[1] === 'VOICE 1'), b = rows.find(row => row[1] === 'VOICE 2');
    [a[col._record_id], b[col._record_id]] = [b[col._record_id], a[col._record_id]];
  }), /Channel Pool: the row "VOICE 1" \(Channel Pool\) exactly matches the exported row of "VOICE 1" \(Channel Pool\)/);

  // A typed-over row keeps its record unless rebuild mode is on; the summary
  // points that out.
  const typedOver = () => setRows(book(), 'Zone Channels', (rows, col) => {
    const row = rows.find(item => item[1] === 'DIG TALK');
    rows[0].forEach((header, index) => { if (header !== '_record_id') row[index] = ''; });
    Object.entries({ 'Assigned Zone': 'ZONE A', 'Name / Alias': 'CH 1', 'Contact Name': 'DIG 1', 'Channel Type': 'Digital', 'Receive Frequency': '461.9000', 'Transmit Frequency': '461.9000', 'Color Code': '1' }).forEach(([header, cell]) => { row[col[header]] = cell; });
  });
  const keptOver = await run(typedOver());
  assert.equal(markerOf(keptOver.doc, 'CH 1'), MARK);
  assert.ok(keptOver.notes.some(note => /Renamed and changed in several other ways, but these rows kept their existing channel records .*: CH 1\. If they are meant to be new channels, tick "Create all channels as new"/.test(note)), keptOver.notes.join(' | '));
  const newOver = await run(typedOver(), { createAllNew: true });
  assert.equal(markerOf(newOver.doc, 'CH 1'), templateMarker('DGTLCONV6PT25'));
  // The note counts old channels whose settings outside the workbook were
  // reset, without naming internal fields.
  assert.ok(newOver.notes.includes('8 channels were rebuilt using CPS template defaults. Some settings not included in the workbook were reset. Review these settings in CPS if needed.'), newOver.notes.join(' | '));
  assert.ok(newOver.notes.includes('3 channels were rebuilt using CPS template defaults. Some settings not included in the workbook were reset. Review these settings in CPS if needed.'), newOver.notes.join(' | '));

  // Duplicate IDs (a copied row): refused normally, ignored in rebuild mode.
  const copied = () => setRows(book(), 'Zone Channels', (rows, col) => { const copy = [...rows.find(row => row[1] === 'DIG TALK')]; copy[col['Name / Alias']] = 'DIG TALK COPY'; rows.push(copy); });
  await refuse(copied(), /Zone Channels: row "DIG TALK COPY" \(ZONE A\) has the same hidden record ID as another row, so it was probably copied from it\. Add new channels in blank rows, clear the copied row's hidden _record_id cell, or tick "Create all channels as new"\./);
  const copiedNew = await run(copied(), { createAllNew: true });
  for (const alias of ['DIG TALK', 'DIG TALK COPY']) {
    assert.equal(markerOf(copiedNew.doc, alias), templateMarker('DGTLCONV6PT25'));
    assert.deepEqual(['CP_ENHCPRVCYIT', 'CP_RASDATAITEM'].map(name => value(channelIn(copiedNew.doc, alias), name)), ['PRIV A', 'RAS A'], `${alias} takes its visible keys from the row`);
  }

  // Invalid IDs: unknown, or from the other channel tab. Rebuild mode ignores them.
  await refuse(setRows(book(), 'Zone Channels', (rows, col) => { rows.find(row => row[1] === 'ANA TALK')[col._record_id] = 'n999999'; }), /Zone Channels: row "ANA TALK" \(ZONE A\) has a hidden record ID \(n999999\) that does not belong to a channel on this tab\. Clear that row's hidden _record_id cell to add it as a new channel, or export the workbook again\./);
  const moveToPool = () => {
    const workbook = book();
    const zone = rowsFor(workbook, 'Zone Channels'), zoneHeaders = zone[0];
    const moved = zone.splice(zone.findIndex(row => row[1] === 'DIG TALK'), 1)[0];
    replaceRows(workbook, 'Zone Channels', zone);
    return setRows(workbook, 'Channel Pool', rows => rows.push(rows[0].map(header => header === 'Assigned Zone' ? 'Channel Pool' : zoneHeaders.includes(header) ? moved[zoneHeaders.indexOf(header)] : '')));
  };
  await refuse(moveToPool(), /Channel Pool: row "DIG TALK" \(Channel Pool\) carries the hidden record ID of a Zone Channels channel\. A channel cannot move between Zone Channels and Channel Pool; delete the row and add the channel again in a blank row on this tab, or tick "Create all channels as new"\./);
  const movedNew = await run(moveToPool(), { createAllNew: true });
  assert.ok(poolChannels(movedNew.doc).some(item => item.getAttribute('alias') === 'DIG TALK'));

  // A channel tab without hidden IDs cannot keep any channel's settings, so it
  // is refused unless every channel is created as new.
  const noIds = () => {
    const workbook = book();
    for (const sheet of ['Zone Channels', 'Channel Pool']) {
      const rows = rowsFor(workbook, sheet), drop = rows[0].indexOf('_record_id');
      replaceRows(workbook, sheet, rows.map(row => row.filter((_, index) => index !== drop)));
    }
    return setRows(workbook, '_CPS_META', rows => { for (const row of rows.filter(item => ['Zone Channels', 'Channel Pool'].includes(item[0]))) { const map = JSON.parse(row[5]); delete map._record_id; row[5] = JSON.stringify(map); } });
  };
  await refuse(noIds(), /Zone Channels: this workbook has no hidden record IDs, so unArchive cannot tell which channels already exist\. Export the workbook again, or tick "Create all channels as new"/);
  assert.equal(allChannels((await run(noIds(), { createAllNew: true })).doc).length, allChannels(marked).length);

  // A list outside the zones that named a removed channel is reported.
  const removedScanned = await run(setRows(book(), 'Zone Channels', rows => rows.splice(rows.findIndex(row => row[1] === 'DIG TALK'), 1)));
  assert.ok(removedScanned.notes.includes('The list "SCAN A" still lists "DIG TALK", which is no longer a channel. Update that list in CPS.'), removedScanned.notes.join(' | '));
  assert.deepEqual(directElements(category(removedScanned.doc, 'ScanItems')).map(canonical), directElements(category(marked, 'ScanItems')).map(canonical), 'The scan list itself is left unchanged.');
}

// The primary workflow: clear an existing configuration and build a new one,
// reusing common channel names, by typing over the old rows, by deleting them
// and adding new rows, or with "Create all channels as new". No new channel
// may carry anything hidden from an old one, and nothing outside the
// channels is rebuilt or dropped.
{
  const MARK = '987';
  const doc = xmlDocument(syntheticCodeplug({ keys: true, poolKeys: true }));
  for (const channel of allChannels(doc)) fieldOf(channel, fieldOf(channel, 'CP_TOT') ? 'CP_TOT' : 'CP_SARQ_AT').textContent = MARK;
  const sourceXml = '<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' + doc.documentElement.toString();
  const original = xmlDocument(sourceXml);
  const workbookBytes = await exportWorkbookFrom(sourceXml);
  // The new configuration reuses names such as DIG TALK and VOICE 1.
  const zonePlan = [
    { 'Name / Alias': 'CH 1', 'Contact Name': 'DIG 1', 'Channel Type': 'Digital', 'Receive Frequency': '451.1000', 'Transmit Frequency': '456.1000', 'Color Code': '5' },
    { 'Name / Alias': 'CH 2', 'Contact Name': 'DIG 2', 'Channel Type': 'Digital', 'Receive Frequency': '451.2000', 'Transmit Frequency': '456.2000', 'Color Code': '5', 'Privacy Enabled': 'True', 'Privacy Key Alias': 'PRIV A' },
    { 'Name / Alias': 'DIG TALK', 'Contact Name': 'DIG 1', 'Channel Type': 'Digital', 'Receive Frequency': '451.3000', 'Transmit Frequency': '456.3000', 'Color Code': '5' },
    { 'Name / Alias': 'CP 1', 'Contact Name': 'CP 1', 'Channel Type': 'Capacity Plus', 'Voice List': 'VOICE A' },
    { 'Name / Alias': 'ANA TALK', 'Channel Type': 'Analog', 'Receive Frequency': '452.1000', 'Transmit Frequency': '452.1000' }
  ].map(row => ({ 'Assigned Zone': 'ZONE A', ...row }));
  const poolPlan = [
    { 'Name / Alias': 'VOICE 1', 'Channel Type': 'Capacity Plus Voice', 'Receive Frequency': '453.1000', 'Transmit Frequency': '458.1000', 'Color Code': '6' },
    { 'Name / Alias': 'VOICE 2', 'Channel Type': 'Capacity Plus Voice', 'Receive Frequency': '453.2000', 'Transmit Frequency': '458.2000', 'Color Code': '6' },
    { 'Name / Alias': 'CH 1', 'Contact Name': 'DIG ALL CALL', 'Channel Type': 'Digital', 'Receive Frequency': '453.3000', 'Transmit Frequency': '458.3000', 'Color Code': '6' }
  ].map(row => ({ 'Assigned Zone': 'Channel Pool', ...row }));
  const typeFor = label => Object.entries({ Digital: 'DGTLCONV6PT25', Analog: 'ANLGCONV', 'Capacity Plus': 'CAPPLUSSYSCONV', 'Capacity Plus Voice': 'CAPPLUSVOICECONV' }).find(([name]) => name === label)[1];
  // Every new channel matches its CPS template except the values the plan sets.
  const assertBuiltFromTemplates = (built, label) => {
    for (const [sheet, plan, parentOf] of [['Zone Channels', zonePlan, () => 'ZONE A'], ['Channel Pool', poolPlan, () => 'CHANNELPOOL']]) {
      const channels = allChannels(built).filter(item => (sheet === 'Channel Pool') === (item.parentElement.parentElement.getAttribute('key') === 'CHANNELPOOL'));
      assert.deepEqual(channels.map(item => item.getAttribute('alias')), plan.map(row => row['Name / Alias']), `${label}: ${sheet} holds exactly the new plan`);
      channels.forEach((channel, index) => {
        const template = xmlDocument(`<r>${CPS_BLUEPRINTS.channels[typeFor(plan[index]['Channel Type'])]}</r>`).documentElement.firstElementChild;
        assert.equal(shape(channel), shape(template), `${label}: ${channel.getAttribute('alias')} has its template's fields`);
        const marker = fieldOf(channel, 'CP_TOT') ? 'CP_TOT' : 'CP_SARQ_AT';
        assert.equal(value(channel, marker), value(template, marker), `${label}: ${channel.getAttribute('alias')} inherits nothing hidden`);
        const expectKeys = plan[index]['Privacy Key Alias'] || (fieldOf(template, 'CP_ENHCPRVCYIT') ? 'NONE' : undefined);
        assert.equal(value(channel, 'CP_ENHCPRVCYIT'), expectKeys, `${label}: ${channel.getAttribute('alias')} privacy key`);
        if (fieldOf(template, 'CP_RASDATAITEM')) assert.equal(value(channel, 'CP_RASDATAITEM'), 'NONE', `${label}: ${channel.getAttribute('alias')} RAS key`);
      });
    }
    for (const name of ['PCRContacts', 'CapacityPlusSiteList', 'ScanItems', 'RoamList']) assert.deepEqual(directElements(category(built, name)).map(canonical), directElements(category(original, name)).map(canonical), `${label}: ${name} kept`);
    assert.deepEqual(directElements(directElements(category(built, 'CapacityPlusVoiceList'), 'set')[0], 'multifield')[0].children.length, 2, `${label}: voice list members still resolve`);
  };
  const fill = (rows, plan) => [rows[0], ...plan.map(entry => rows[0].map(header => entry[header] ?? ''))];
  // a. Typed over the old rows, which keep their hidden IDs, with rebuild mode.
  const overtyped = workbookFrom(workbookBytes);
  for (const [sheet, plan] of [['Zone Channels', zonePlan], ['Channel Pool', poolPlan]]) {
    const rows = rowsFor(overtyped, sheet), idColumn = rows[0].indexOf('_record_id');
    const next = fill(rows, plan);
    next.slice(1).forEach((row, index) => { row[idColumn] = rows[index + 1]?.[idColumn] || ''; });
    replaceRows(overtyped, sheet, next);
  }
  const rebuilt = await importWorkbook(overtyped, { createAllNew: true });
  assert.equal(rebuilt.ok, true, rebuilt.message);
  assertBuiltFromTemplates(xmlDocument(rebuilt.xml), 'typed over, rebuild mode');
  assert.ok(rebuilt.summary.includes('5 channels created as new from their CPS templates, replacing 8 old channels.'), rebuilt.summary.join(' | '));
  assert.ok(rebuilt.summary.includes('3 channels created as new from their CPS templates, replacing 3 old channels.'), rebuilt.summary.join(' | '));
  assert.ok(!rebuilt.notes.some(note => /still lists/.test(note)));
  // b. Old rows deleted and the plan entered in blank rows, without rebuild mode.
  const retyped = workbookFrom(workbookBytes);
  for (const [sheet, plan] of [['Zone Channels', zonePlan], ['Channel Pool', poolPlan]]) replaceRows(retyped, sheet, fill(rowsFor(retyped, sheet), plan));
  const fresh = await importWorkbook(retyped);
  assert.equal(fresh.ok, true, fresh.message);
  assertBuiltFromTemplates(xmlDocument(fresh.xml), 'deleted and retyped');
  assert.ok(fresh.notes.some(note => /same name, without its other settings: DIG TALK, ANA TALK/.test(note)), fresh.notes.join(' | '));
  // c. Rebuild mode on an export whose rows were deleted and retyped gives the same result.
  assertBuiltFromTemplates(xmlDocument((await importWorkbook(retyped, { createAllNew: true })).xml), 'deleted and retyped, rebuild mode');
}

// The reported case: old channels with internal settings that differ from
// the template (for example CP_DEVMEDFORINDDATACALL and OffSet) get one
// plain, counted note when rebuilt, and none when nothing differed.
{
  const doc = xmlDocument(syntheticCodeplug());
  const digital = allChannels(doc).filter(channel => channelType(channel) === 'DGTLCONV6PT25');
  for (const channel of digital.slice(0, 2)) {
    fieldOf(channel, 'CP_DEVMEDFORINDDATACALL').textContent = 'True';
    fieldOf(channel, 'OffSet').textContent = '3.000000';
  }
  const changedBytes = await exportWorkbookFrom('<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' + doc.documentElement.toString());
  const rebuilt = await importWorkbook(workbookFrom(changedBytes), { createAllNew: true });
  assert.equal(rebuilt.ok, true, rebuilt.message);
  assert.deepEqual(rebuilt.notes, ['2 channels were rebuilt using CPS template defaults. Some settings not included in the workbook were reset. Review these settings in CPS if needed.']);
  assert.doesNotMatch(rebuilt.summary.join(' '), /DEVMEDFORINDDATACALL|OffSet|Off Set|3\.000000/);
  const untouched = await importWorkbook(workbookFrom(bytes), { createAllNew: true });
  assert.deepEqual(untouched.notes, [], 'No note when no old channel differed from its template.');
  assert.ok(untouched.summary.every(line => /channels created as new from their CPS templates/.test(line)));
  // The checker itself catches the text that used to appear.
  assert.throws(() => assertPlainSummary({ summary: ['Old channels had settings: Channel — DEVMEDFORINDDATACALL on 25, Off Set on 25.'] }));
  assert.throws(() => assertPlainSummary({ summary: ['Reset CP_TOT on 3 and STR_12PT5KHZ.'] }));
  assert.throws(() => assertPlainSummary({ summary: ['Scan list in ScanItems.'] }));
}

console.log('Fast tests passed.');
