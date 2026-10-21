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
    await rejected(edit(fresh(), 'Channel Pool', { 'POOL ALL CALL': { 'Channel Type': 'Capacity Plus Voice', 'Contact Name': '', 'Scan List': '', 'Receive Only': entry } }), /POOL ALL CALL: Capacity Plus Voice channels have no Receive Only setting in CPS/);
  }
  const converted = await imported(edit(fresh(), 'Channel Pool', { 'POOL ALL CALL': { 'Channel Type': 'Capacity Plus Voice', 'Contact Name': '', 'Scan List': '', 'Receive Only': '' } }));
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
  assert.ok(!rowsFor(fresh(), 'Channel Pool')[0].includes('Privacy Enabled') && !rowsFor(fresh(), 'Channel Pool')[0].includes('Auto Roam'), 'Only Zone Channels shows these columns.');

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

console.log('Fast tests passed.');
