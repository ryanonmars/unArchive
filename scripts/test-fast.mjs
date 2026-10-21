// Fast regression suite. It runs on a synthetic codeplug built from the
// committed blueprints, so it needs no customer data and is tracked in Git.
// The full suite against real CPS fixtures stays local (see README).
import assert from 'node:assert/strict';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { CPS_BLUEPRINTS } from '../blueprints.generated.js';
import {
  exportWorkbookFrom, workbookFrom, rowsFor, replaceRows, appendRow, importWorkbook,
  xmlDocument, directElements, category, allChannels, fieldOf, channelType
} from './test-harness.mjs';
import { syntheticCodeplug } from './test-fixture.mjs';

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

console.log('Fast tests passed.');
