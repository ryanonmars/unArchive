import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { CPS_BLUEPRINTS } from './blueprints.generated.js';
import { readText as readClipboardText, writeText as writeClipboardText } from '@tauri-apps/plugin-clipboard-manager';
import { ask, save } from '@tauri-apps/plugin-dialog';
import { writeFile } from '@tauri-apps/plugin-fs';
import { getVersion } from '@tauri-apps/api/app';
import { check as checkForUpdate } from '@tauri-apps/plugin-updater';
(() => {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const input = $('xml-input'), output = $('xml-output');
  const xmlStatus = $('xml-status'), outputStatus = $('output-status'), summary = $('sheet-summary');
  let currentModel = null;
  const RESERVED = new Set(['_CPS_TREE', '_CPS_META', '_CPS_CHOICES']);
  const CATEGORY_LABELS = {
    GeneralSettings: 'General Settings', PCRContacts: 'Contacts', Zone: 'Zones & Channel Pool',
    ControlButtons: 'Programmable Buttons', DigitalRXGroupList: 'Digital RX Groups',
    CapacityPlusVoiceList: 'Capacity Plus Voice Lists', CapacityPlusDataList: 'Capacity Plus Data Lists',
    CapacityPlusSiteList: 'Capacity Plus Site Lists', TextMessages: 'Text Messages',
    SignalingSystems: 'Signaling Systems', MDCSystems: 'MDC Systems', QuikCallIISystems: 'Quik Call II Systems'
  };
  const COLLECTION_LABELS = {
    ZoneItems: 'Channels', DigitalCalls: 'Digital Calls', MDCCalls: 'MDC Calls',
    QuikCallIICalls: 'Quik Call II Calls', CapacityPlusCalls: 'Capacity Plus Calls',
    CapacityPlusSiteListItems: 'Sites', RoamListItems: 'Roam List Members',
    ScanListItems: 'Scan List Members', OneTouchAccessItems: 'One Touch Assignments'
  };
  const FIELD_LABELS = {
    ContactName: 'Contact Name', Comments: 'Comments', CP_CNVPERSALIAS: 'Channel Name',
    CP_RXFREQ: 'Receive Frequency', CP_TXFREQ: 'Transmit Frequency', CP_COLORCODE: 'Color Code',
    CP_PERSTYPE: 'Channel Type', CP_SLTASSGMNT: 'Time Slot', CP_TXPWR: 'Transmit Power',
    CP_RXONLYEN: 'Receive Only', CP_VOXSTATE: 'VOX Enabled', CP_VOICEPRIVACYEN: 'Privacy Enabled', CP_TOT: 'Time-Out Timer',
    CP_TXINHXPLEN: 'Admit Criteria', CP_MYCALLADCRTR: 'In Call Criteria',
    CP_TRUNKCHANLISTIT: 'Voice List', CP_LCPSTLSTIT: 'Site List', CP_LCPALATRM: 'Auto Roam',
    CP_ENHCPRVCYIT: 'Privacy Key Alias', CP_RASDATAITEM: 'RAS Key Alias', RC_RAS_PROPRIETARY_KEY_ID: 'RAS Proprietary Key ID',
    CP_TGLISTIT: 'RX Group List', CP_SCANLISTIT: 'Scan List', CP_EMSYSIT: 'Emergency System', CP_UKPPERS: 'Contact Name',
    DU_CALLALIAS: 'Contact Name', DU_CALLLSTID: 'Call ID', DU_CALLTYPE: 'Call Type',
    CAPPLUSUCL_CALLALIAS: 'Contact Name', CAPPLUSUCL_CALLLSTID: 'Call ID',
    CAPPLUSUCL_CALLTYPE: 'Call Type', TP_TRUNKPERALIAS: 'Voice List Name',
    RC_PWRUPBPFLNM: 'Power-Up Bitmap', GPS_GPSEN: 'GPS Enabled', RC_CPPASSWORD: 'Codeplug Password'
  };
  const FOCUSED_COLUMNS = {
    Zones: ['@alias'],
    'Zone Channels': ['@alias', 'field:CP_UKPPERS', 'field:CP_PERSTYPE', 'field:CP_VOICEPRIVACYEN', 'field:CP_ENHCPRVCYIT', 'field:CP_RASDATAITEM', 'field:CP_TRUNKCHANLISTIT', 'field:CP_LCPSTLSTIT', 'field:CP_LCPALATRM', 'field:CP_RXFREQ', 'field:CP_TXFREQ', 'field:CP_COLORCODE', 'field:CP_TXINHXPLEN', 'field:CP_MYCALLADCRTR', 'field:CP_RXONLYEN', 'field:CP_SCANLISTIT'],
    'Channel Pool': ['@alias', 'field:CP_UKPPERS', 'field:CP_PERSTYPE', 'field:CP_RXFREQ', 'field:CP_TXFREQ', 'field:CP_COLORCODE', 'field:CP_RXONLYEN', 'field:CP_SCANLISTIT']
  };
  const CHANNEL_TYPE_LABELS = {
    CAPPLUSSYSCONV: 'Capacity Plus',
    CAPPLUSVOICECONV: 'Capacity Plus Voice',
    LINKEDCAPPLUSSYSCONV: 'Linked Capacity Plus',
    DGTLCONV6PT25: 'Digital',
    ANLGCONV: 'Analog'
  };
  const CRITERIA_OPTIONS = {
    'field:CP_TXINHXPLEN': [
      { label: 'Always', value: 'ALWAYS' },
      { label: 'Channel Free', value: 'ONCHNNLFREE' },
      { label: 'Color Code Free', value: 'MTCHCLRCD' }
    ],
    'field:CP_MYCALLADCRTR': [
      { label: 'Follow Admit Criteria', value: 'FOLLOW_ADMIT_CRITERIA' },
      { label: 'Always', value: 'ALWAYS' },
      { label: 'TX Interrupt', value: 'TRANSMITTER_INTERRUPT' }
    ]
  };
  const CONTACT_TYPE_LABELS = { DigitalCalls: 'Digital', CapacityPlusCalls: 'Capacity Plus', MDCCalls: 'MDC', QuikCallIICalls: 'Quik Call II', PhoneCalls: 'Phone' };
  const CONTACT_TYPE_ORDER = ['CapacityPlusCalls', 'DigitalCalls', 'MDCCalls', 'QuikCallIICalls', 'PhoneCalls'];
  const CONTACT_COLLECTIONS = ['MDCCalls', 'QuikCallIICalls', 'DigitalCalls', 'CapacityPlusCalls', 'PhoneCalls'];
  const CONTACT_CALL_TYPE_OPTIONS = ['MDC', 'Quik Call II', 'Digital', 'Capacity Plus'];
  const ALL_CALL_IDS = { Digital: '16777215', 'Capacity Plus': '255' };
  const ZONE_CHANNEL_TYPE_OPTIONS = ['Analog', 'Digital', 'Capacity Plus', 'Linked Capacity Plus'];
  // Channel Pool rows always import into the pool, whatever Assigned Zone
  // says. Capacity Plus Voice is the repeater channel that voice lists name.
  // A Digital pool channel keeps a contact, such as a Digital All Call, in the
  // radio for scanning without putting it in a zone everyone can select.
  const CHANNEL_POOL_ZONE_OPTIONS = ['Channel Pool'];
  const CHANNEL_POOL_TYPE_OPTIONS = ['Capacity Plus Voice', 'Digital'];
  const TRUE_FALSE_OPTIONS = ['True', 'False'];
  const SECURITY_PRIVACY_TYPE_OPTIONS = ['None', 'Basic', 'Enhanced'];
  // CPS silently cuts zone, channel, and contact names to 16 characters on
  // paste, and a cut channel name no longer matches its slot, so CPS moves
  // that channel to the end of its zone. Longer names are rejected instead.
  const CPS_NAME_LENGTH = 16;
  const NAME_LENGTH_COLUMNS = new Set(['@alias', '__contact_name']);

  function message(el, text, type = '') { el.textContent = text; el.className = `status ${type}`; }
  function text(value) { return value == null ? '' : String(value); }
  function safeName(name, used) {
    let base = name.replace(/[\\/?*\[\]:]/g, '_').slice(0, 31) || 'Sheet', candidate = base, n = 2;
    while (used.has(candidate)) candidate = `${base.slice(0, 28)}_${n++}`;
    used.add(candidate); return candidate;
  }
  function words(value) {
    const abbreviations = { RX: 'Receive', TX: 'Transmit', FREQ: 'Frequency', EN: 'Enabled', PWR: 'Power', PERS: 'Personality', PERSTYPE: 'Personality Type', ALIAS: 'Name', LIST: 'List', ID: 'ID', NUM: 'Number', TMR: 'Timer', DUR: 'Duration', VOL: 'Volume', MSG: 'Message', SYS: 'System', GRP: 'Group', CALL: 'Call', CH: 'Channel', CHNL: 'Channel', DGT: 'Digital', DGTL: 'Digital', ANLG: 'Analog', RVR: 'Revert', RVRTPER: 'Revert Channel', RSSI: 'RSSI', GPS: 'GPS', VOX: 'VOX' };
    return value.split('_').filter(Boolean).map(part => abbreviations[part] || part.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/\b\w/g, c => c.toUpperCase())).join(' ');
  }
  function friendlyField(name) {
    if (FIELD_LABELS[name]) return FIELD_LABELS[name];
    const prefixes = { CP: 'Channel', RC: 'Radio', SS: 'Scan', DU: 'Digital Call', AU: 'MDC Call', QU: 'Quik Call II', TP: 'Capacity Plus', LCPSTLST: 'Site List', DATAREVERT: 'Data Revert', SURC: 'Radio', PSWD: 'Password' };
    const parts = name.split('_'); const prefix = prefixes[parts[0]];
    return `${prefix ? `${prefix} — ` : ''}${words(prefix ? parts.slice(1).join('_') : name)}`;
  }
  function friendlyAttribute(name) { return ({ alias: 'Name / Alias', key: 'Record Type', index: 'Position', name: 'Record Name' })[name] || words(name); }
  function criteriaOption(technical, value) {
    const normalized = text(value).trim().toLowerCase();
    return (CRITERIA_OPTIONS[technical] || []).find(option => option.value.toLowerCase() === normalized || option.label.toLowerCase() === normalized);
  }
  function criteriaLabel(technical, value) { return criteriaOption(technical, value)?.label || ''; }
  function displayValue(technical, value) {
    if (technical === 'field:CP_PERSTYPE') return CHANNEL_TYPE_LABELS[text(value)] || value;
    return CRITERIA_OPTIONS[technical] ? (criteriaLabel(technical, value) || value) : value;
  }
  function cpsValue(technical, value) {
    if (technical === 'field:CP_PERSTYPE') {
      const found = Object.entries(CHANNEL_TYPE_LABELS).find(([, label]) => label.toLowerCase() === text(value).trim().toLowerCase());
      return found ? found[0] : value;
    }
    return criteriaOption(technical, value)?.value || value;
  }
  function categoryLabel(name) { return CATEGORY_LABELS[name] || words(name); }
  function collectionLabel(name) { return COLLECTION_LABELS[name] || words(name); }
  function fieldKey(name) { return `field:${name}`; }
  function labelKey(name) { return `label:${name}`; }
  function parseXml(source) {
    const doc = new DOMParser().parseFromString(source, 'application/xml');
    const error = doc.querySelector('parsererror');
    if (error) throw new Error(error.textContent.replace(/\s+/g, ' ').slice(0, 250));
    if (!doc.documentElement || doc.documentElement.tagName !== 'config') throw new Error('Expected a CPS document with a <config> root element.');
    return doc;
  }
  const blueprintCache = new Map();
  function blueprintNode(doc, xml) {
    if (!xml) return null;
    if (!blueprintCache.has(xml)) {
      const parsed = new DOMParser().parseFromString(`<config>${xml}</config>`, 'application/xml');
      const error = parsed.querySelector('parsererror');
      if (error || !parsed.documentElement?.firstElementChild) throw new Error('An internal CPS blueprint is invalid.');
      blueprintCache.set(xml, parsed.documentElement.firstElementChild);
    }
    const source = blueprintCache.get(xml);
    return typeof doc.importNode === 'function' ? doc.importNode(source, true) : source.cloneNode(true);
  }
  function assignedTree(doc) {
    const records = [], nodes = new Map(); let count = 0;
    const visit = (node, parent, order) => {
      if (node.nodeType !== Node.ELEMENT_NODE) return;
      const id = `n${++count}`; nodes.set(node, id);
      records.push({ id, parentId: parent || '', order, tag: node.tagName, attrs: [...node.attributes].reduce((o, a) => (o[a.name] = a.value, o), {}), text: [...node.childNodes].filter(n => n.nodeType === Node.TEXT_NODE).map(n => n.nodeValue).join('').trim() });
      [...node.children].forEach((child, i) => visit(child, id, i));
    };
    visit(doc.documentElement, '', 0); return { records, nodes };
  }
  function directChildren(node, tag) { return [...node.children].filter(n => n.tagName === tag); }
  function findOrCreateCollection(parent, name) {
    let collection = directChildren(parent, 'collection').find(n => n.getAttribute('name') === name);
    if (!collection) { collection = parent.ownerDocument.createElement('collection'); collection.setAttribute('name', name); parent.append(collection); }
    return collection;
  }
  function sheetRows(nodes, sheet) {
    const rows = [];
    for (const item of sheet.items) {
      const row = {};
      for (const attr of item.node.attributes) row[`@${attr.name}`] = attr.value;
      for (const field of directChildren(item.node, 'field')) { row[fieldKey(field.getAttribute('name'))] = field.textContent; if (field.hasAttribute('Name')) row[labelKey(field.getAttribute('name'))] = field.getAttribute('Name'); }
      for (const multi of directChildren(item.node, 'multifield')) row[`multifield:${multi.getAttribute('name')}`] = JSON.stringify(directChildren(multi, 'value').map(v => v.textContent));
      row._record_id = nodes.get(item.node) || '';
      rows.push(row);
    }
    return rows;
  }
  function buildModel(doc) {
    const tree = assignedTree(doc), sheets = [], used = new Set();
    const addSheet = (label, category, collectionName, items, layout = label === 'Contacts' ? 'contacts' : 'grid') => {
      const sheet = { id: `s${sheets.length + 1}`, label, category, collectionName, layout, sheetName: safeName(label, used), items };
      sheets.push(sheet);
    };
    const category = (name) => directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === name);
    const records = (node) => node ? directChildren(node, 'set').map((child, order) => ({ node: child, parent: node, order })) : [];
    const collectionRecords = (parents, collectionName) => parents.flatMap(parent => {
      const collection = directChildren(parent, 'collection').find(node => node.getAttribute('name') === collectionName);
      return collection ? directChildren(collection, 'set').map((node, order) => ({ node, parent, order })) : [];
    });
    const contactsCategory = category('PCRContacts');
    const zoneCategory = category('Zone');
    const voiceListCategory = category('CapacityPlusVoiceList');
    const siteListCategory = category('CapacityPlusSiteList');
    const contacts = records(contactsCategory);
    const zoneRecords = records(zoneCategory);
    const channelPool = zoneRecords.filter(item => item.node.getAttribute('key') === 'CHANNELPOOL' || item.node.getAttribute('alias') === 'Channel Pool');
    const zones = zoneRecords.filter(item => !channelPool.includes(item));
    const securityChannels = collectionRecords(zoneRecords.map(item => item.node), 'ZoneItems');
    const isConfiguredSecurityAlias = (value) => Boolean(text(value).trim()) && text(value).trim().toUpperCase() !== 'NONE';
    const uniqueSecurityAliases = (field) => [...new Set(securityChannels.map(item => directFieldValue(item.node, field)).filter(isConfiguredSecurityAlias))];
    const rasAliases = uniqueSecurityAliases('CP_RASDATAITEM');
    const privacyAliases = uniqueSecurityAliases('CP_ENHCPRVCYIT');
    const securityItems = Array.from({ length: Math.max(1, rasAliases.length, privacyAliases.length) }, (_, index) => ({
      rasAlias: rasAliases[index] || '', privacyAlias: privacyAliases[index] || ''
    }));

    addSheet('Contacts', 'PCRContacts', '', contacts);
    addSheet('Zones', 'Zone', '', zones);
    addSheet('Zone Channels', 'Zone', 'ZoneItems', collectionRecords(zones.map(item => item.node), 'ZoneItems'));
    addSheet('Channel Pool', 'Zone', 'ZoneItems', collectionRecords(channelPool.map(item => item.node), 'ZoneItems'));
    addSheet('Voice Lists', 'CapacityPlusVoiceList', '', records(voiceListCategory), 'voiceLists');
    addSheet('Site Lists', 'CapacityPlusSiteList', '', records(siteListCategory), 'siteLists');
    addSheet('Security', 'Security', '', securityItems, 'security');
    return { doc, tree, sheets };
  }
  function columns(rows) {
    const base = []; const extra = new Set();
    rows.forEach(r => Object.keys(r).forEach(k => { if (!base.includes(k)) extra.add(k); }));
    return base.concat([...extra].sort((a, b) => a.localeCompare(b)));
  }
  function uniqueHeader(candidate, technical, used) {
    let value = candidate || technical, number = 2;
    while (used.has(value)) value = `${candidate} (${number++})`;
    used.add(value); return value;
  }
  function gridHeader(technical) {
    if (technical.startsWith('@')) return friendlyAttribute(technical.slice(1));
    if (technical.startsWith('field:')) return friendlyField(technical.slice(6));
    if (technical.startsWith('label:')) return `CPS display label — ${friendlyField(technical.slice(6))}`;
    if (technical.startsWith('multifield:')) return `${friendlyField(technical.slice(11))} (multiple values)`;
    return technical;
  }
  function settingsPresentation(sheet, nodes) {
    const headers = ['Setting', 'Value to send to CPS', 'Current CPS display label', '_record_id', '_field', '_kind', 'Delete setting?'];
    const rows = [];
    for (const item of sheet.items) {
      const id = nodes.get(item.node);
      for (const attribute of item.node.attributes) rows.push([friendlyAttribute(attribute.name), attribute.value, '', id, attribute.name, 'attribute', '']);
      for (const field of directChildren(item.node, 'field')) rows.push([friendlyField(field.getAttribute('name')), field.textContent, field.getAttribute('Name') || '', id, field.getAttribute('name'), 'field', '']);
      for (const multi of directChildren(item.node, 'multifield')) rows.push([`${friendlyField(multi.getAttribute('name'))} (multiple values)`, JSON.stringify(directChildren(multi, 'value').map(v => v.textContent)), '', id, multi.getAttribute('name'), 'multifield', '']);
    }
    return { headers, rows, columnMap: {}, hidden: new Set(['_record_id', '_field', '_kind']), guide: rows.map(r => [sheet.label, r[0], r[4], r[5] === 'multifield' ? 'Use a JSON array for multiple values.' : '']) };
  }
  function contactCalls(contact) { return directChildren(contact, 'collection').flatMap(collection => directChildren(collection, 'set').map(node => ({ node, collection: collection.getAttribute('name') }))); }
  function activeContactTypes(contact) { return [...new Set(contactCalls(contact).map(call => CONTACT_TYPE_LABELS[call.collection]).filter(Boolean))]; }
  function contactIds(contact) { return [...new Set(contactCalls(contact).flatMap(call => directChildren(call.node, 'field').filter(field => field.getAttribute('name') === 'PeudoCallId').map(field => field.textContent)))]; }
  function primaryContactType(contact) {
    const collections = new Set(contactCalls(contact).map(call => call.collection));
    const primary = CONTACT_TYPE_ORDER.find(collection => collections.has(collection));
    return primary ? CONTACT_TYPE_LABELS[primary] : '';
  }
  function requestedContactType(value) {
    const normalized = text(value).trim().toLowerCase().replace(/\s+/g, ' ');
    return Object.values(CONTACT_TYPE_LABELS).find(label => label.toLowerCase() === normalized) || '';
  }
  function isAllCallContact(node) {
    return contactCalls(node).some(call => {
      const callType = directFieldValue(call.node, 'CallType').toUpperCase();
      const key = text(call.node.getAttribute('key')).toUpperCase();
      return callType.endsWith('-ALL CALL') || ['DGTLALLSYSCALL', 'CAPPLUSDGTLALLSYSCALL'].includes(key);
    });
  }
  function isCpsAllCallContact(node) {
    return contactCalls(node).some(call => {
      const key = text(call.node.getAttribute('key')).toUpperCase();
      const id = directFieldValue(call.node, 'PeudoCallId');
      return (call.collection === 'DigitalCalls' && key === 'DGTLALLSYSCALL' && id === '16777215') ||
        (call.collection === 'CapacityPlusCalls' && key === 'CAPPLUSDGTLALLSYSCALL' && id === '255');
    });
  }
  function contactTemplateForType(candidates, type) {
    const collectionName = Object.entries(CONTACT_TYPE_LABELS).find(([, label]) => label === type)?.[0];
    if (!collectionName) return null;
    const matches = candidates.filter(candidate => directChildren(candidate, 'collection').some(collection => collection.getAttribute('name') === collectionName && directChildren(collection, 'set').length));
    return matches.find(candidate => !isAllCallContact(candidate)) || matches[0] || null;
  }
  function isCapacityPlusPersonality(value) { return text(value).toUpperCase().includes('CAPPLUS'); }
  function contactPresentation(sheet, nodes) {
    const headers = ['Contact Name', 'Call ID', 'Call Type', '_record_id'];
    const columnMap = { 'Contact Name': '__contact_name', 'Call ID': '__contact_id', 'Call Type': '__contact_type', '_record_id': '_record_id' };
    const rows = sheet.items.map(item => {
      const calls = contactCalls(item.node);
      const values = (name) => calls.flatMap(call => directChildren(call.node, 'field').filter(field => field.getAttribute('name') === name).map(field => field.textContent));
      const ids = values('PeudoCallId');
      return [directChildren(item.node, 'field').find(field => field.getAttribute('name') === 'ContactName')?.textContent || item.node.getAttribute('alias') || '', [...new Set(ids)].join(' / '), primaryContactType(item.node), nodes.get(item.node) || ''];
    });
    const hidden = new Set(['_record_id']);
    return { headers, rows, columnMap, hidden, guide: headers.slice(0, 3).map(header => [sheet.label, header, columnMap[header], '']) };
  }
  function gridPresentation(sheet, nodes) {
    const rawRows = sheetRows(nodes, sheet), allColumns = columns(rawRows);
    const requested = FOCUSED_COLUMNS[sheet.label] || allColumns;
    // Planning sheets keep their full editing schema even when the selected
    // source currently contains no records of a requested type.
    const technical = (FOCUSED_COLUMNS[sheet.label] ? requested : requested.filter(key => allColumns.includes(key))).filter(key => !key.startsWith('_'));
    const used = new Set(), columnMap = {}, headers = technical.map(key => {
      const label = uniqueHeader(gridHeader(key), key, used); columnMap[label] = key; return label;
    });
    headers.push('_record_id'); columnMap._record_id = '_record_id';
    const parentLabel = sheet.collectionName ? (sheet.collectionName === 'ZoneItems' ? 'Assigned Zone' : 'Parent Record') : '';
    if (parentLabel) { headers.splice(0, 0, parentLabel); columnMap[parentLabel] = '__parent_label'; }
    const hidden = new Set(headers.filter(h => columnMap[h].startsWith('label:') || h === '_record_id'));
    const guide = headers.filter(h => !hidden.has(h)).map(h => [sheet.label, h, columnMap[h], columnMap[h] === '_delete' ? 'Enter TRUE only when you want CPS to remove this record.' : columnMap[h] === '__parent_label' ? 'For reference. Copy a row to keep its zone assignment.' : '']);
    return { headers, rows: rawRows.map((row, index) => {
      const values = technical.map(key => displayValue(key, row[key] ?? ''));
      values.push(row._record_id || '');
      if (parentLabel) { const parent = sheet.items[index].parent; values.splice(0, 0, parent.getAttribute('alias') || parent.getAttribute('name') || ''); }
      return values;
    }), columnMap, hidden, guide };
  }
  function voiceListPresentation(sheet) {
    const headers = ['Voice List Name', 'Member Channel'];
    const columnMap = { 'Voice List Name': '__voice_list_name', 'Member Channel': '__voice_list_member' };
    const rows = sheet.items.flatMap(item => {
      const name = item.node.getAttribute('alias') || directFieldValue(item.node, 'TP_TRUNKPERALIAS');
      const memberField = directChildren(item.node, 'multifield').find(field => field.getAttribute('name') === 'CapacityPlusVoiceListItems');
      const members = memberField ? directChildren(memberField, 'value').map(value => value.textContent) : [];
      return members.length ? members.map(member => [name, member]) : [[name, '']];
    });
    return { headers, rows, columnMap, hidden: new Set(), guide: [
      ['Voice Lists', 'Voice List Name', '__voice_list_name', 'Repeat the same name on each row to add members to that voice list.'],
      ['Voice Lists', 'Member Channel', '__voice_list_member', 'Enter one Capacity Plus channel name per row. Leave this blank to keep an empty voice list.']
    ] };
  }
  function siteListPresentation(sheet) {
    const headers = ['Site Name', 'Site Alias', 'Site ID', 'Voice List'];
    const columnMap = {
      'Site Name': '__site_list_name', 'Site Alias': '__site_alias', 'Site ID': '__site_id',
      'Voice List': '__site_voice_list'
    };
    const rows = sheet.items.flatMap(item => {
      const name = item.node.getAttribute('alias') || directFieldValue(item.node, 'LCPSTLST_ALIAS');
      const collection = directChildren(item.node, 'collection').find(node => node.getAttribute('name') === 'CapacityPlusSiteListItems');
      const sites = collection ? directChildren(collection, 'set') : [];
      const siteRow = (site) => [name, directFieldValue(site, 'LCPSTLST_SITEALIAS'), directFieldValue(site, 'LCPSTLST_SITEID'), directFieldValue(site, 'LCPSTLST_VCLSTIT')];
      return sites.length ? sites.map(siteRow) : [[name, '', '', '']];
    });
    return { headers, rows, columnMap, hidden: new Set(), guide: [
      ['Site Lists', 'Site Name', '__site_list_name', 'Repeat this name for every site in the same list.'],
      ['Site Lists', 'Site Alias', '__site_alias', 'One CPS site per row.'],
      ['Site Lists', 'Site ID', '__site_id', 'The numeric site ID used by the linked system.'],
      ['Site Lists', 'Voice List', '__site_voice_list', 'The Capacity Plus Voice List used at this site.']
    ] };
  }
  function securityPresentation(sheet) {
    const headers = ['RAS Key Alias', 'Privacy Key Alias', 'Privacy Key ID', 'Privacy Type'];
    const columnMap = {
      'RAS Key Alias': '__ras_alias', 'Privacy Key Alias': '__privacy_alias',
      'Privacy Key ID': '__privacy_key_id', 'Privacy Type': '__privacy_type'
    };
    const rows = sheet.items.map(item => [item.rasAlias, item.privacyAlias, '', '']);
    return { headers, rows, columnMap, hidden: new Set(), guide: [
      ['Security', 'RAS Key Alias', '__ras_alias', 'Enter aliases already configured in the target codeplug. The RAS key value itself stays target-managed.'],
      ['Security', 'Privacy Key Alias', '__privacy_alias', 'Enter aliases already configured in the target codeplug.'],
      ['Security', 'Privacy Key ID', '__privacy_key_id', 'For your target-codeplug reference and validation. CPS does not include this definition in clipboard XML.'],
      ['Security', 'Privacy Type', '__privacy_type', 'For your target-codeplug reference and validation. CPS does not include this definition in clipboard XML.']
    ] };
  }
  function presentationFor(sheet, nodes) { if (sheet.label === 'Contacts') return contactPresentation(sheet, nodes); if (sheet.layout === 'voiceLists') return voiceListPresentation(sheet); if (sheet.layout === 'siteLists') return siteListPresentation(sheet); if (sheet.layout === 'security') return securityPresentation(sheet); return sheet.layout === 'settings' ? settingsPresentation(sheet, nodes) : gridPresentation(sheet, nodes); }
  function worksheetFromPresentation(presentation, namedChoices = {}, fixedChoices = {}) {
    const ws = XLSX.utils.aoa_to_sheet([presentation.headers, ...presentation.rows]);
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(0, presentation.rows.length), c: Math.max(0, presentation.headers.length - 1) } }) };
    ws['!cols'] = presentation.headers.map(header => ({ wch: Math.min(38, Math.max(14, header.length + 2)), hidden: presentation.hidden.has(header) }));
    ws.__unArchiveDataValidations = presentation.headers.flatMap((header, index) => {
      const technical = presentation.columnMap[header], options = CRITERIA_OPTIONS[technical], fixedOptions = fixedChoices[technical];
      const sqref = `${XLSX.utils.encode_col(index)}2:${XLSX.utils.encode_col(index)}5000`;
      if (NAME_LENGTH_COLUMNS.has(technical)) return [{ sqref, maxLength: CPS_NAME_LENGTH }];
      if (options) return [{ sqref, values: options.map(option => option.label) }];
      if (fixedOptions) return [{ sqref, values: fixedOptions }];
      const namedChoice = namedChoices[technical];
      return namedChoice ? [{ sqref, formula: typeof namedChoice === 'string' ? namedChoice : namedChoice.formula, strict: typeof namedChoice === 'string' ? false : namedChoice.strict }] : [];
    });
    return ws;
  }
  function showModel(model) {
    summary.innerHTML = '';
    model.sheets.forEach(sheet => { const tag = document.createElement('span'); tag.className = 'tag'; tag.textContent = `${sheet.label} (${sheet.items.length})`; summary.append(tag); });
    $('export-workbook').disabled = false;
  }
  function prepareModelFromInput() {
    try {
      currentModel = buildModel(parseXml(input.value)); showModel(currentModel); $('download-source').disabled = false;
      message(xmlStatus, `Ready: ${currentModel.sheets.length} friendly workbook tabs will be created.`, 'ok');
      return true;
    } catch (err) { currentModel = null; $('export-workbook').disabled = true; message(xmlStatus, err.message, 'error'); return false; }
  }
  // A new zone layout reuses the pasted codeplug for every hidden CPS setting
  // and replaces only its zones. Generated zones and channels are written to
  // the visible sheets without hidden record IDs, so import treats them as new
  // records: they are built from the CPS blueprints and must pass the same
  // required-field checks as any channel the user adds by hand.
  const TEMPLATE_MAX_ZONES = 250, TEMPLATE_MAX_CHANNELS = 160, TEMPLATE_DEFAULT_GROUP = { type: 'Digital', channels: 16 };
  const TEMPLATE_REQUIREMENTS = {
    Analog: 'Receive Frequency and Transmit Frequency',
    Digital: 'Receive Frequency, Transmit Frequency, Color Code, and a Digital Contact Name',
    'Capacity Plus': 'a Capacity Plus Contact Name and Voice List',
    'Linked Capacity Plus': 'a Capacity Plus Contact Name and Site List'
  };
  // Backing store for the per-zone rows. It is longer than the visible zone
  // count while the user is typing, so shrinking and regrowing keeps choices.
  // Each zone holds one or more channel groups, generated in order.
  const templateZones = [];
  function templateZoneName(index) { return `Zone ${index + 1}`; }
  function templateChannelName(index, count) { return `CH ${String(index + 1).padStart(Math.max(2, String(count).length), '0')}`; }
  function blueprintHasField(type, name) { return text(CPS_BLUEPRINTS.channels[cpsValue('field:CP_PERSTYPE', type)]).includes(`<field name="${name}"`); }
  // Defaults apply only where the channel type's CPS record has the field.
  // Analog has no privacy, RAS, or In Call Criteria, and CPS does not offer
  // Color Code Free as an analog admit criterion.
  function templateChannelValues(type, defaults) {
    const values = { 'field:CP_PERSTYPE': type };
    const set = (name, value) => { if (blueprintHasField(type, name)) values[fieldKey(name)] = value; };
    set('CP_VOICEPRIVACYEN', defaults.privacyEnabled);
    set('CP_ENHCPRVCYIT', defaults.privacyAlias);
    set('CP_RASDATAITEM', defaults.rasAlias);
    set('CP_TXINHXPLEN', type === 'Analog' && defaults.admit === 'Color Code Free' ? 'Always' : defaults.admit);
    set('CP_MYCALLADCRTR', defaults.inCall);
    set('CP_RXONLYEN', defaults.receiveOnly);
    return values;
  }
  function templateZoneCount() { return Number($('template-zones').value); }
  function copyTemplateZone(zone) { return { name: '', groups: (zone?.groups || [TEMPLATE_DEFAULT_GROUP]).map(group => ({ ...group })) }; }
  function templateTypeSelect(group, label) {
    const select = document.createElement('select');
    select.setAttribute('aria-label', label);
    for (const option of ZONE_CHANNEL_TYPE_OPTIONS) {
      const item = document.createElement('option');
      item.setAttribute('value', option); item.textContent = option;
      if (option === group.type) item.setAttribute('selected', '');
      select.append(item);
    }
    select.addEventListener('change', () => { group.type = select.value; showTemplatePreview(); });
    return select;
  }
  function templateCountInput(group, label) {
    const input = document.createElement('input');
    input.setAttribute('type', 'number'); input.setAttribute('min', '1'); input.setAttribute('max', String(TEMPLATE_MAX_CHANNELS));
    input.setAttribute('aria-label', label); input.setAttribute('value', String(group.channels));
    input.addEventListener('input', () => { group.channels = input.value; showTemplatePreview(); });
    return input;
  }
  function templateButton(label, ariaLabel, className, onClick) {
    const button = document.createElement('button');
    button.setAttribute('type', 'button'); button.setAttribute('class', className); button.setAttribute('aria-label', ariaLabel);
    button.textContent = label;
    button.addEventListener('click', onClick);
    return button;
  }
  function renderTemplateZones(focus) {
    const count = Math.min(TEMPLATE_MAX_ZONES, Math.max(0, Math.floor(templateZoneCount()) || 0));
    while (templateZones.length < count) templateZones.push(copyTemplateZone(templateZones.at(-1)));
    const cell = (...children) => { const td = document.createElement('td'); td.append(...children); return td; };
    const rows = [];
    let focusTarget = null;
    templateZones.slice(0, count).forEach((zone, index) => {
      const name = templateZoneName(index);
      const zoneName = document.createElement('input');
      zoneName.setAttribute('type', 'text'); zoneName.setAttribute('maxlength', String(CPS_NAME_LENGTH));
      zoneName.setAttribute('placeholder', name); zoneName.setAttribute('aria-label', `${name} name`); zoneName.setAttribute('value', text(zone.name));
      zoneName.addEventListener('input', () => { zone.name = zoneName.value; showTemplatePreview(); });
      zone.groups.forEach((group, groupIndex) => {
        const label = groupIndex ? `${name} channel type ${groupIndex + 1}` : `${name} channel type`;
        const type = templateTypeSelect(group, label);
        const channels = templateCountInput(group, groupIndex ? `${name} channel count ${groupIndex + 1}` : `${name} channel count`);
        if (focus && focus.zone === zone && focus.group === group) focusTarget = type;
        const row = document.createElement('tr');
        if (groupIndex) {
          row.setAttribute('class', 'template-subrow');
          const remove = templateButton('×', `Remove ${label}`, 'template-remove', () => {
            zone.groups.splice(zone.groups.indexOf(group), 1);
            renderTemplateZones(); showTemplatePreview();
          });
          const and = cell(document.createTextNode('and'));
          and.setAttribute('class', 'template-and');
          row.append(cell(), and, cell(type), cell(channels), cell(remove));
        } else {
          // A new group defaults to a type the zone does not use yet.
          const add = templateButton('+ Add type', `Add another channel type to ${name}`, 'template-add', () => {
            const used = new Set(zone.groups.map(item => item.type));
            const added = { type: ZONE_CHANNEL_TYPE_OPTIONS.find(option => !used.has(option)) || ZONE_CHANNEL_TYPE_OPTIONS[0], channels: 1 };
            zone.groups.push(added);
            renderTemplateZones({ zone, group: added }); showTemplatePreview();
          });
          row.append(cell(document.createTextNode(String(index + 1))), cell(zoneName), cell(type), cell(channels), cell(add));
        }
        rows.push(row);
      });
    });
    $('template-zone-rows').replaceChildren(...rows);
    focusTarget?.focus?.();
  }
  // Contact rows of the new layout. Linked Capacity Plus channels use
  // Capacity Plus contacts, so both labels create Capacity Plus contacts and
  // share its Call ID range; only the default name prefix differs.
  const TEMPLATE_CONTACT_TYPES = {
    'Capacity Plus': { callType: 'Capacity Plus', prefix: 'CP' },
    'Linked Capacity Plus': { callType: 'Capacity Plus', prefix: 'LCP' },
    Digital: { callType: 'Digital', prefix: 'DIG' }
  };
  const TEMPLATE_MAX_CONTACTS = 1000, TEMPLATE_MAX_CALL_ID = 16776415;
  const TEMPLATE_ALL_CALL_NAMES = { 'Capacity Plus': 'ALL CALL', Digital: 'DIG ALL CALL' };
  const TEMPLATE_CALL_OPTIONS = ['Group Call', 'All Call'];
  // The first row has no type, so the pasted contacts are kept until one is
  // chosen. A blank start continues the IDs of the call type's previous block.
  const templateContacts = [{ type: '', call: 'Group Call', count: 10, start: '' }];
  function renderTemplateContacts(focus) {
    const cell = (...children) => { const td = document.createElement('td'); td.append(...children); return td; };
    let focusTarget = null;
    $('template-contact-rows').replaceChildren(...templateContacts.map((row, index) => {
      const label = index ? `Contact type ${index + 1}` : 'Contact type';
      const type = document.createElement('select');
      type.setAttribute('aria-label', label);
      for (const option of ['', ...Object.keys(TEMPLATE_CONTACT_TYPES)]) {
        const item = document.createElement('option');
        item.setAttribute('value', option); item.textContent = option || 'None';
        if (option === row.type) item.setAttribute('selected', '');
        type.append(item);
      }
      type.addEventListener('change', () => { row.type = type.value; renderTemplateContacts(); showTemplatePreview(); });
      const call = document.createElement('select');
      call.setAttribute('aria-label', `${label} call`);
      for (const option of TEMPLATE_CALL_OPTIONS) {
        const item = document.createElement('option');
        item.setAttribute('value', option); item.textContent = option;
        if (option === row.call) item.setAttribute('selected', '');
        call.append(item);
      }
      call.addEventListener('change', () => { row.call = call.value; renderTemplateContacts(); showTemplatePreview(); });
      const count = document.createElement('input');
      count.setAttribute('type', 'number'); count.setAttribute('min', '1'); count.setAttribute('max', String(TEMPLATE_MAX_CONTACTS));
      count.setAttribute('aria-label', `${label} count`);
      // All Call is a single contact per system with a fixed ID.
      count.setAttribute('value', row.call === 'All Call' ? '1' : String(row.count));
      count.addEventListener('input', () => { row.count = count.value; showTemplatePreview(); });
      const start = document.createElement('input');
      start.setAttribute('type', 'number'); start.setAttribute('min', '1'); start.setAttribute('max', String(TEMPLATE_MAX_CALL_ID));
      start.setAttribute('aria-label', `${label} starting Call ID`);
      start.setAttribute('placeholder', row.call === 'All Call' && row.type ? ALL_CALL_IDS[TEMPLATE_CONTACT_TYPES[row.type].callType] : 'Auto');
      start.setAttribute('value', row.call === 'All Call' ? '' : text(row.start));
      start.addEventListener('input', () => { row.start = start.value; showTemplatePreview(); });
      if (!row.type) [call, count, start].forEach(input => input.setAttribute('disabled', ''));
      else if (row.call === 'All Call') [count, start].forEach(input => input.setAttribute('disabled', ''));
      if (focus === row) focusTarget = type;
      const action = index
        ? templateButton('×', `Remove ${label}`, 'template-remove', () => { templateContacts.splice(templateContacts.indexOf(row), 1); renderTemplateContacts(); showTemplatePreview(); })
        : templateButton('+ Add contact type', 'Add another contact type', 'template-add', () => {
          const used = new Set(templateContacts.map(item => item.type));
          const added = { type: Object.keys(TEMPLATE_CONTACT_TYPES).find(option => !used.has(option)) || 'Capacity Plus', call: 'Group Call', count: 10, start: '' };
          templateContacts.push(added);
          renderTemplateContacts(added); showTemplatePreview();
        });
      const tr = document.createElement('tr');
      if (index) tr.setAttribute('class', 'template-subrow');
      tr.append(cell(type), cell(call), cell(count), cell(start), cell(action));
      return tr;
    }));
    focusTarget?.focus?.();
  }
  function templateContactOptions() {
    const rows = templateContacts.filter(row => row.type);
    if (!rows.length) return null;
    const groups = rows.map(row => {
      const { callType, prefix } = TEMPLATE_CONTACT_TYPES[row.type] || {};
      if (!callType) throw new Error('New layout: select a supported contact type.');
      if (row.call === 'All Call') return { label: row.type, callType, allCall: true };
      const count = Number(row.count);
      if (!Number.isInteger(count) || count < 1 || count > TEMPLATE_MAX_CONTACTS) throw new Error(`New layout: ${row.type} contacts must be a whole number from 1 through ${TEMPLATE_MAX_CONTACTS}.`);
      const startText = text(row.start).trim(), start = startText ? Number(startText) : null;
      if (start !== null && (!Number.isInteger(start) || start < 1 || start > TEMPLATE_MAX_CALL_ID)) throw new Error(`New layout: the ${row.type} starting Call ID must be a whole number from 1 through ${TEMPLATE_MAX_CALL_ID}.`);
      return { label: row.type, callType, prefix, count, start };
    });
    const allCalls = new Set();
    for (const group of groups.filter(item => item.allCall)) {
      if (allCalls.has(group.callType)) throw new Error(`New layout: a codeplug has one ${group.callType} All Call; Capacity Plus and Linked Capacity Plus share it.`);
      allCalls.add(group.callType);
    }
    const nameTotals = {};
    for (const group of groups.filter(item => !item.allCall)) nameTotals[group.prefix] = (nameTotals[group.prefix] || 0) + group.count;
    // Names number per prefix through all rows. Call IDs run from each block's
    // starting ID, or continue the call type's previous block (from 1). Linked
    // Capacity Plus shares the Capacity Plus IDs, so blocks of either label
    // must not overlap, and a Capacity Plus block must skip the All Call ID.
    const nextId = {}, nextName = {}, usedIds = {}, contacts = [];
    for (const group of groups) {
      if (group.allCall) { contacts.push({ name: TEMPLATE_ALL_CALL_NAMES[group.callType], id: ALL_CALL_IDS[group.callType], callType: group.callType, label: 'All Call' }); continue; }
      const first = group.start ?? (nextId[group.callType] || 1), last = first + group.count - 1;
      if (last > TEMPLATE_MAX_CALL_ID) throw new Error(`New layout: the ${group.label} block would end at Call ID ${last}; the highest allowed is ${TEMPLATE_MAX_CALL_ID}.`);
      if (group.callType === 'Capacity Plus' && first <= 255 && last >= 255) throw new Error(`New layout: the ${group.label} block (Call IDs ${first}–${last}) includes 255, which is reserved for the Capacity Plus All Call. Choose a starting Call ID that avoids it.`);
      const used = usedIds[group.callType] = usedIds[group.callType] || new Map();
      for (let id = first; id <= last; id++) {
        if (used.has(id)) throw new Error(`New layout: Call ID ${id} is used by both the ${used.get(id)} and ${group.label} blocks. Change a starting Call ID so the blocks do not overlap.`);
        used.set(id, group.label);
        const number = nextName[group.prefix] = (nextName[group.prefix] || 0) + 1;
        contacts.push({ name: `${group.prefix} ${String(number).padStart(Math.max(2, String(nameTotals[group.prefix]).length), '0')}`, id: String(id), callType: group.callType, label: group.label });
      }
      nextId[group.callType] = last + 1;
    }
    return contacts;
  }
  function templateContactSummary(contacts) {
    const labels = [...new Set(contacts.map(contact => contact.label))];
    return `${contacts.length} contact${contacts.length === 1 ? '' : 's'} (${labels.map(label => `${contacts.filter(contact => contact.label === label).length} ${label}`).join(', ')})`;
  }
  // Template contacts replace the pasted ones, except CPS's own scaffold
  // contacts that combine several call types (such as Call1), which import
  // preserves and the simplified sheet cannot recreate.
  function removeContacts(doc) {
    const category = directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === 'PCRContacts');
    recordsForCategory(category).filter(node => activeContactTypes(node).length < 2).forEach(node => node.remove());
  }
  // Voice list members name Channel Pool channels, so they are emptied with
  // the pool; the voice lists themselves are kept.
  function clearChannelPool(doc) {
    const category = (name) => directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === name);
    for (const pool of recordsForCategory(category('Zone')).filter(isChannelPoolRecord)) {
      for (const collection of directChildren(pool, 'collection').filter(node => node.getAttribute('name') === 'ZoneItems')) collection.replaceChildren();
    }
    for (const list of recordsForCategory(category('CapacityPlusVoiceList'))) {
      directChildren(list, 'multifield').filter(node => node.getAttribute('name') === 'CapacityPlusVoiceListItems').forEach(node => node.replaceChildren());
    }
  }
  function templateOptions() {
    if (!$('template-enabled').checked) return null;
    const count = templateZoneCount();
    if (!Number.isInteger(count) || count < 1 || count > TEMPLATE_MAX_ZONES) throw new Error(`New zone layout: Number of zones must be a whole number from 1 through ${TEMPLATE_MAX_ZONES}.`);
    const usedNames = new Map();
    const zones = templateZones.slice(0, count).map((zone, index) => {
      // A blank name keeps the numbered default.
      const name = text(zone.name).trim() || templateZoneName(index);
      if (name.length > CPS_NAME_LENGTH) throw new Error(`New zone layout: zone name "${name}" is longer than ${CPS_NAME_LENGTH} characters.`);
      if (matchName(name) === matchName('Channel Pool')) throw new Error('New zone layout: "Channel Pool" is reserved by CPS. Choose another zone name.');
      if (usedNames.has(matchName(name))) throw new Error(`New zone layout: zone ${index + 1} and zone ${usedNames.get(matchName(name))} are both named "${name}".`);
      usedNames.set(matchName(name), index + 1);
      const groups = zone.groups.map(group => {
        const channels = Number(group.channels);
        if (!Number.isInteger(channels) || channels < 1 || channels > TEMPLATE_MAX_CHANNELS) throw new Error(`New zone layout: ${name} needs a whole number of channels from 1 through ${TEMPLATE_MAX_CHANNELS}.`);
        if (!ZONE_CHANNEL_TYPE_OPTIONS.includes(group.type)) throw new Error(`New zone layout: select a supported channel type for ${name}.`);
        return { type: group.type, channels };
      });
      const total = groups.reduce((sum, group) => sum + group.channels, 0);
      if (total > TEMPLATE_MAX_CHANNELS) throw new Error(`New zone layout: ${name} has ${total} channels; a zone can have at most ${TEMPLATE_MAX_CHANNELS}.`);
      // Channel numbers run through the whole zone, across its channel types.
      const types = groups.flatMap(group => Array(group.channels).fill(group.type));
      return { name, channels: types.map((type, i) => ({ name: templateChannelName(i, total), type })) };
    });
    const defaults = {
      privacyEnabled: text($('template-privacy-enabled').value), privacyAlias: text($('template-privacy-alias').value).trim(),
      rasAlias: text($('template-ras-alias').value).trim(), admit: text($('template-admit').value),
      inCall: text($('template-in-call').value), receiveOnly: text($('template-receive-only').value)
    };
    const channels = zones.flatMap(zone => zone.channels);
    const types = ZONE_CHANNEL_TYPE_OPTIONS.filter(type => channels.some(channel => channel.type === type));
    return { zones, defaults, types, channelCount: channels.length, contacts: templateContactOptions(), clearPool: $('template-clear-pool').checked };
  }
  function templateSummary(template) {
    const byType = template.types.map(type => `${template.zones.flatMap(zone => zone.channels).filter(channel => channel.type === type).length} ${type}`).join(', ');
    return `${template.zones.length} zone${template.zones.length === 1 ? '' : 's'} with ${template.channelCount} channel${template.channelCount === 1 ? '' : 's'} (${byType})`;
  }
  function templateRequirements(template) { return template.types.map(type => `${type}: ${TEMPLATE_REQUIREMENTS[type]}`).join('; '); }
  function removeZones(doc) {
    const zoneCategory = directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === 'Zone');
    recordsForCategory(zoneCategory).filter(node => !isChannelPoolRecord(node)).forEach(node => node.remove());
  }
  // Import rejects a channel key alias that is not on the Security sheet, so
  // list each template alias there when a generated channel uses it.
  function addSecurityAlias(presentation, column, alias) {
    if (!presentation || column < 0 || !alias || presentation.rows.some(row => matchName(row[column]) === matchName(alias))) return;
    const free = presentation.rows.find(row => !text(row[column]).trim());
    if (free) free[column] = alias;
    else { const row = presentation.headers.map(() => ''); row[column] = alias; presentation.rows.push(row); }
  }
  function addTemplateRows(presentations, template) {
    const rowFor = (presentation, values) => presentation.headers.map(header => values[presentation.columnMap[header]] ?? '');
    const sheet = (label) => presentations.find(item => item.sheet.label === label)?.presentation;
    const zones = sheet('Zones'), channels = sheet('Zone Channels'), security = sheet('Security'), contacts = sheet('Contacts');
    if (!zones || !channels) throw new Error('New zone layout: the Zones and Zone Channels sheets are unavailable.');
    for (const contact of template.contacts || []) contacts?.rows.push(rowFor(contacts, { '__contact_name': contact.name, '__contact_id': contact.id, '__contact_type': contact.callType }));
    for (const zone of template.zones) {
      zones.rows.push(rowFor(zones, { '@alias': zone.name }));
      for (const channel of zone.channels) channels.rows.push(rowFor(channels, { ...templateChannelValues(channel.type, template.defaults), '__parent_label': zone.name, '@alias': channel.name }));
    }
    if (template.types.some(type => blueprintHasField(type, 'CP_ENHCPRVCYIT'))) {
      addSecurityAlias(security, security?.headers.indexOf('RAS Key Alias'), template.defaults.rasAlias);
      addSecurityAlias(security, security?.headers.indexOf('Privacy Key Alias'), template.defaults.privacyAlias);
    }
  }
  function showTemplatePreview() {
    const enabled = $('template-enabled').checked;
    $('template-fields').disabled = !enabled;
    const preview = $('template-preview');
    if (!enabled) { message(preview, ''); return; }
    try {
      const template = templateOptions();
      message(preview, [
        `Export will replace the pasted zones with ${templateSummary(template)}.`,
        template.contacts ? `Contacts will be replaced with ${templateContactSummary(template.contacts)}.` : 'The pasted contacts will be kept.',
        template.clearPool ? 'The Channel Pool will be cleared.' : ''
      ].filter(Boolean).join(' '), 'ok');
    }
    catch (err) { message(preview, err.message, 'error'); }
  }
  function requireXlsx() { if (!XLSX?.utils) throw new Error('The bundled Excel library has not loaded. Restart the app and try again.'); }
  function requireZip() { if (!JSZip) throw new Error('The bundled workbook helper has not loaded. Restart the app and try again.'); }
  function freezeTopRowInSheetXml(xml) {
    return xml.replace(/<sheetView\b([^>]*?)(?:\/>|>([\s\S]*?)<\/sheetView>)/, (_, rawAttributes, contents = '') => {
      const attributes = rawAttributes.replace(/\/\s*$/, '');
      const remaining = contents.replace(/<pane\b[^>]*\/>/g, '').replace(/<selection\b[^>]*\/>/g, '');
      return `<sheetView${attributes}><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/>${remaining}</sheetView>`;
    });
  }
  function addDataValidationsToSheetXml(xml, validations) {
    if (!validations?.length) return xml;
    const entries = validations.map(validation => validation.maxLength ? `<dataValidation type="textLength" operator="lessThanOrEqual" allowBlank="1" showErrorMessage="1" errorTitle="Name too long" error="CPS allows at most ${validation.maxLength} characters and cuts longer names when pasting." sqref="${validation.sqref}"><formula1>${validation.maxLength}</formula1></dataValidation>` : `<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="${validation.strict === false ? '0' : '1'}" sqref="${validation.sqref}"><formula1>${validation.formula || `&quot;${validation.values.join(',')}&quot;`}</formula1></dataValidation>`).join('');
    const block = `<dataValidations count="${validations.length}">${entries}</dataValidations>`;
    // SheetJS writes <ignoredErrors> after <autoFilter>.  Excel requires
    // <dataValidations> before that element, so appending it at the end makes
    // the worksheet invalid even though LibreOffice is lenient about the order.
    if (xml.includes('<ignoredErrors')) return xml.replace('<ignoredErrors', `${block}<ignoredErrors`);
    return xml.includes('<pageMargins') ? xml.replace('<pageMargins', `${block}<pageMargins`) : xml.replace('</worksheet>', `${block}</worksheet>`);
  }
  async function workbookBlobWithFrozenHeaders(wb) {
    const zip = await JSZip.loadAsync(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }));
    const worksheetPaths = Object.keys(zip.files).filter(path => /^xl\/worksheets\/sheet\d+\.xml$/.test(path));
    await Promise.all(worksheetPaths.map(async path => {
      const file = zip.file(path);
      const index = Number(path.match(/sheet(\d+)\.xml$/)?.[1]) - 1;
      const sheet = wb.Sheets[wb.SheetNames[index]];
      zip.file(path, addDataValidationsToSheetXml(freezeTopRowInSheetXml(await file.async('string')), sheet?.__unArchiveDataValidations));
    }));
    return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  async function saveFile(contents, filename, filter) {
    const path = await save({ title: `Save ${filename}`, defaultPath: filename, filters: [filter] });
    if (!path) return false;
    const bytes = contents instanceof Blob ? new Uint8Array(await contents.arrayBuffer()) : new TextEncoder().encode(contents);
    await writeFile(path, bytes);
    return true;
  }
  async function exportWorkbook() {
    try {
      requireXlsx(); requireZip();
      const template = templateOptions();
      let model;
      if (template) {
        const doc = parseXml(input.value); removeZones(doc);
        if (template.contacts) removeContacts(doc);
        if (template.clearPool) clearChannelPool(doc);
        model = buildModel(doc);
      } else { model = currentModel || buildModel(parseXml(input.value)); currentModel = model; }
      showModel(model);
      const wb = XLSX.utils.book_new();
      const instructions = XLSX.utils.aoa_to_sheet([
        ['READ THIS BEFORE EDITING'],
        ['Topic', 'What you need to know'],
        ['Start with the correct radio model', 'Copy from a codeplug made for the same radio model and feature set as the target radios. unArchive does not convert one radio model into another. The source may otherwise be empty or contain records you plan to replace.'],
        ['Editing an existing record', "Changing names, frequencies, assignments, and other visible values keeps the record's original hidden CPS structure."],
        ['Adding a record or changing its type', 'New records use complete CPS-created templates built into unArchive. If you change a Digital contact or channel to Capacity Plus, Analog, or another type, unArchive replaces the old record with a complete record of the selected type. The source does not need to already contain that type.'],
        ['Add records on blank rows', 'Dragging values down a visible column is safe. Do not copy an entire existing row to create a new record; that also copies a hidden identity and import will stop with an error.'],
        ['Analog channel defaults', 'New channels changed to Analog start as carrier-squelch, 12.5 kHz, high-power channels with no PL/DPL tone. Configure any required PL/DPL, bandwidth, power, or signaling changes in CPS after pasting.'],
        ['RAS and privacy', 'The target codeplug must already contain every referenced RAS and privacy key. CPS clipboard XML carries the key names but not the secret key values.'],
        ['Capacity Plus Voice Lists', 'Check Voice List membership in CPS after pasting. Some CPS clipboard operations may not transfer those members.'],
        ['Import safety checks', 'unArchive verifies the visible workbook values and important links before enabling Copy XML. If it finds a problem, it clears the output and explains what must be corrected.'],
        ['Before fleet deployment', 'Run CPS validation, save and reopen the codeplug, and program and function-test representative radios before programming the fleet. Keep an untouched backup codeplug.'],
        ...(template ? [['New zone layout', `This workbook replaces the source zones with ${templateSummary(template)}. Rename the zones and channels as needed, then complete every generated channel — ${templateRequirements(template)}. You may also change a row's Channel Type. Import stops and names the first incomplete channel until every one is filled in. Scan lists, roam lists, and buttons that pointed at the old zone channels must be rechecked in CPS.`]] : []),
        ...(template?.contacts ? [['New contacts', `The source contacts were replaced with ${templateContactSummary(template.contacts)}, numbered with sequential Call IDs. Rename them and change their Call IDs as needed, then pick them in the Zone Channels Contact Name column. Capacity Plus and Linked Capacity Plus channels both use Capacity Plus contacts. All Call contacts keep their reserved IDs (Capacity Plus 255, Digital 16777215). CPS scaffold contacts that combine several call types are kept.`]] : []),
        ...(template?.clearPool ? [['Channel Pool cleared', 'The source Channel Pool channels were removed, and voice list members that pointed at them were emptied. Add the Capacity Plus repeater channels to Channel Pool and their members to Voice Lists before importing.']] : [])
      ]);
      instructions['!cols'] = [{ wch: 27 }, { wch: 118 }];
      XLSX.utils.book_append_sheet(wb, instructions, 'READ FIRST');
      const meta = [['sheet', 'id', 'category', 'collectionName', 'layout', 'column_map_json']];
      const presentations = model.sheets.map(sheet => ({ sheet, presentation: presentationFor(sheet, model.tree.nodes) }));
      if (template) addTemplateRows(presentations, template);
      const contacts = presentations.find(item => item.sheet.label === 'Contacts');
      const zones = presentations.find(item => item.sheet.label === 'Zones');
      const voiceLists = presentations.find(item => item.sheet.label === 'Voice Lists');
      const channelPool = presentations.find(item => item.sheet.label === 'Channel Pool');
      const siteLists = presentations.find(item => item.sheet.label === 'Site Lists');
      const security = presentations.find(item => item.sheet.label === 'Security');
      // Voice/Site list names occur on one row per member.  Their helper ranges
      // use one inexpensive first-occurrence formula per row.  This preserves a
      // live, de-duplicated list without the large array calculations that can
      // make a workbook slow to open or use.
      // Keep the Excel drop-downs live when the user adds Contacts or Zones after export.
      // Column A is the name column on both sheets; COUNTA excludes the header before sizing the range.
      const range = (item, column = 'A') => {
        const sheetName = item.sheet.sheetName.replace(/'/g, "''");
        return `OFFSET('${sheetName}'!$${column}$2,0,0,MAX(1,COUNTA('${sheetName}'!$${column}:$${column})-1),1)`;
      };
      const namedChoices = {
        'field:CP_UKPPERS': 'unArchiveContactNames',
        '__parent_label': 'unArchiveZoneNames',
        'field:CP_ENHCPRVCYIT': { formula: 'unArchivePrivacyKeyAliases', strict: true },
        'field:CP_RASDATAITEM': { formula: 'unArchiveRASKeyAliases', strict: true },
        'field:CP_TRUNKCHANLISTIT': { formula: 'unArchiveVoiceListNames', strict: false },
        'field:CP_LCPSTLSTIT': { formula: 'unArchiveSiteListNames', strict: false }
      };
      const choiceRange = (column) => `'_CPS_CHOICES'!$${column}$2:$${column}$501`;
      wb.Workbook = {
        Names: [
          ...(contacts ? [{ Name: 'unArchiveContactNames', Ref: range(contacts) }] : []),
          ...(zones ? [{ Name: 'unArchiveZoneNames', Ref: range(zones) }] : []),
          ...(voiceLists ? [{ Name: 'unArchiveVoiceListNames', Ref: choiceRange('A') }] : []),
          ...(channelPool ? [{ Name: 'unArchiveChannelPoolNames', Ref: range(channelPool, 'B') }] : []),
          ...(siteLists ? [{ Name: 'unArchiveSiteListNames', Ref: choiceRange('B') }] : []),
          ...(security ? [{ Name: 'unArchiveRASKeyAliases', Ref: choiceRange('C') }, { Name: 'unArchivePrivacyKeyAliases', Ref: choiceRange('D') }] : [])
        ],
        CalcPr: { calcMode: 'auto', fullCalcOnLoad: '1', forceFullCalc: '1' }
      };
      for (const { sheet, presentation } of presentations) {
        const choices = sheet.label === 'Zone Channels' ? namedChoices : sheet.label === 'Channel Pool' ? { 'field:CP_UKPPERS': namedChoices['field:CP_UKPPERS'] } : sheet.label === 'Site Lists' ? { '__site_voice_list': { formula: 'unArchiveVoiceListNames', strict: false } } : sheet.label === 'Voice Lists' ? { '__voice_list_member': { formula: 'unArchiveChannelPoolNames', strict: true } } : {};
        const fixedChoices = sheet.label === 'Contacts' ? { '__contact_type': CONTACT_CALL_TYPE_OPTIONS } : sheet.label === 'Zone Channels' ? { 'field:CP_PERSTYPE': ZONE_CHANNEL_TYPE_OPTIONS, 'field:CP_VOICEPRIVACYEN': TRUE_FALSE_OPTIONS, 'field:CP_LCPALATRM': TRUE_FALSE_OPTIONS, 'field:CP_RXONLYEN': TRUE_FALSE_OPTIONS } : sheet.label === 'Channel Pool' ? { '__parent_label': CHANNEL_POOL_ZONE_OPTIONS, 'field:CP_PERSTYPE': CHANNEL_POOL_TYPE_OPTIONS, 'field:CP_RXONLYEN': TRUE_FALSE_OPTIONS } : sheet.label === 'Security' ? { '__privacy_type': SECURITY_PRIVACY_TYPE_OPTIONS } : {};
        XLSX.utils.book_append_sheet(wb, worksheetFromPresentation(presentation, choices, fixedChoices), sheet.sheetName);
        meta.push([sheet.sheetName, sheet.id, sheet.category, sheet.collectionName, sheet.layout, JSON.stringify(presentation.columnMap)]);
      }
      const treeRows = [['id', 'parent_id', 'order', 'tag', 'attributes_json', 'text'], ...model.tree.records.map(r => [r.id, r.parentId, r.order, r.tag, JSON.stringify(r.attrs), r.text])];
      const choiceRows = Array.from({ length: 500 }, () => ['', '', '', '']);
      const choices = XLSX.utils.aoa_to_sheet([['Voice List Names', 'Site List Names', 'RAS Key Aliases', 'Privacy Key Aliases'], ...choiceRows]);
      const securitySheet = security ? `'${security.sheet.sheetName.replace(/'/g, "''")}'` : '';
      const liveUniqueChoiceFormula = (item, column, row) => {
        if (!item) return '';
        const sheetName = item.sheet.sheetName.replace(/'/g, "''");
        const value = `'${sheetName}'!A${row}`;
        const seen = `'${sheetName}'!$A$2:A${row}`;
        return `IF(AND(${value}<>\"\",COUNTIF(${seen},${value})=1),${value},\"\")`;
      };
      for (let row = 2; row <= 501; row++) {
        if (voiceLists) choices[`A${row}`] = { t: 'n', f: liveUniqueChoiceFormula(voiceLists, 'A', row) };
        if (siteLists) choices[`B${row}`] = { t: 'n', f: liveUniqueChoiceFormula(siteLists, 'B', row) };
        // Key-alias choices start with NONE so a channel's key can be removed.
        if (row === 2) { choices.C2 = { t: 's', v: 'NONE' }; choices.D2 = { t: 's', v: 'NONE' }; }
        else if (security) {
          choices[`C${row}`] = { t: 'n', f: `IF(${securitySheet}!A${row - 1}<>"",${securitySheet}!A${row - 1},"")` };
          choices[`D${row}`] = { t: 'n', f: `IF(${securitySheet}!B${row - 1}<>"",${securitySheet}!B${row - 1},"")` };
        }
      }
      XLSX.utils.book_append_sheet(wb, choices, '_CPS_CHOICES');
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(treeRows), '_CPS_TREE');
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(meta), '_CPS_META');
      wb.Workbook = wb.Workbook || {}; wb.Workbook.Sheets = wb.SheetNames.map(n => ({ Hidden: RESERVED.has(n) ? 1 : 0 }));
      const saved = await saveFile(await workbookBlobWithFrozenHeaders(wb), 'codeplug-editable.xlsx', { name: 'Excel Workbook', extensions: ['xlsx'] });
      const savedText = template ? `Workbook saved with a new layout of ${templateSummary(template)}${template.contacts ? ` and ${templateContactSummary(template.contacts)}` : ''}${template.clearPool ? ', with the Channel Pool cleared' : ''}. Before importing, complete each generated channel — ${templateRequirements(template)}.` : 'Workbook saved.';
      message(xmlStatus, saved ? savedText : 'Workbook save cancelled.', saved ? 'ok' : '');
    } catch (err) { message(xmlStatus, err.message, 'error'); }
  }
  // A row counts only when a visible column has a value. Clearing a row's
  // visible cells leaves its hidden _record_id behind; that row was deleted.
  function rowsFromSheet(wb, name) { const values = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: '' }); const headers = values.shift() || []; return values.filter(r => r.some((v, i) => text(v) !== '' && !text(headers[i]).startsWith('_'))).map(r => Object.fromEntries(headers.map((h, i) => [text(h), r[i] ?? '']))); }
  function treeFromRows(rows) {
    const doc = document.implementation.createDocument('', '', null), byId = new Map();
    rows.forEach(row => {
      const node = doc.createElement(row.tag); try { Object.entries(JSON.parse(row.attributes_json || '{}')).forEach(([k,v]) => node.setAttribute(k, text(v))); } catch { throw new Error(`Invalid attributes_json in tree row ${row.id}.`); }
      if (row.text) node.textContent = row.text; byId.set(row.id, node);
    });
    const groups = new Map(); rows.forEach(row => { const parent = text(row.parent_id); if (!groups.has(parent)) groups.set(parent, []); groups.get(parent).push(row); });
    for (const [parentId, children] of groups) {
      const parent = parentId ? byId.get(parentId) : doc;
      if (!parent) throw new Error('The hidden _CPS_TREE sheet is incomplete or has broken parent IDs.');
      children.sort((a, b) => (+a.order || 0) - (+b.order || 0)).forEach(row => parent.append(byId.get(row.id)));
    }
    if (!doc.documentElement) throw new Error('The hidden _CPS_TREE sheet is incomplete.');
    return { doc, byId };
  }
  function truthy(value) { return ['true', '1', 'yes'].includes(text(value).trim().toLowerCase()); }
  function updateRecord(doc, node, row) {
    Object.entries(row).forEach(([key, value]) => {
      if (key.startsWith('@')) node.setAttribute(key.slice(1), text(value));
      if (key.startsWith(fieldKey(''))) { const name = key.slice(6); if (!name) return; let field = directChildren(node, 'field').find(f => f.getAttribute('name') === name); if (!field && text(value) === '') return; if (!field) { field = doc.createElement('field'); field.setAttribute('name', name); node.append(field); } field.textContent = text(value); const label = row[labelKey(name)] || (CRITERIA_OPTIONS[fieldKey(name)] ? criteriaLabel(fieldKey(name), value) : ''); if (text(label) !== '') field.setAttribute('Name', text(label)); }
      if (key.startsWith('multifield:')) { const name = key.slice(11); let multi = directChildren(node, 'multifield').find(m => m.getAttribute('name') === name); if (!multi && text(value) === '') return; if (!multi) { multi = doc.createElement('multifield'); multi.setAttribute('name', name); node.append(multi); } let values; try { values = JSON.parse(text(value) || '[]'); } catch { values = text(value).split('|').map(v => v.trim()).filter(Boolean); } multi.replaceChildren(...values.map(v => { const el = doc.createElement('value'); el.textContent = text(v); return el; })); }
    });
    if (row['field:CP_PERSTYPE'] !== undefined && text(row['field:CP_PERSTYPE']) !== '') node.setAttribute('key', text(row['field:CP_PERSTYPE']));
  }
  function updateSetting(doc, byId, row) {
    const node = byId.get(text(row._record_id)); if (!node) return;
    const name = text(row._field), kind = text(row._kind);
    if (!name) return;
    if (truthy(row['Delete setting?'])) {
      if (kind === 'attribute') node.removeAttribute(name);
      else if (kind === 'field') directChildren(node, 'field').find(f => f.getAttribute('name') === name)?.remove();
      else if (kind === 'multifield') directChildren(node, 'multifield').find(f => f.getAttribute('name') === name)?.remove();
      return;
    }
    if (kind === 'attribute') { node.setAttribute(name, text(row['Value to send to CPS'])); return; }
    if (kind === 'field') updateRecord(doc, node, { [fieldKey(name)]: row['Value to send to CPS'], [labelKey(name)]: row['Current CPS display label'] });
    if (kind === 'multifield') updateRecord(doc, node, { [`multifield:${name}`]: row['Value to send to CPS'] });
  }
  function syncPlanningNames(info, node, row) {
    const setField = (target, name, value) => updateRecord(node.ownerDocument, target, { [fieldKey(name)]: value });
    if (info.category === 'PCRContacts' && row['field:ContactName'] !== undefined) {
      const name = text(row['field:ContactName']); node.setAttribute('alias', name); setField(node, 'ContactName', name);
      for (const collection of directChildren(node, 'collection')) for (const call of directChildren(collection, 'set')) {
        ['DU_CALLALIAS', 'CAPPLUSUCL_CALLALIAS', 'AU_CALLALIAS', 'QU_CALLENTRYALIAS'].forEach(field => {
          if (directChildren(call, 'field').some(item => item.getAttribute('name') === field)) setField(call, field, name);
        });
      }
    }
    if (info.sheet === 'Zones' && row['@alias'] !== undefined) setField(node, 'ZP_ZONEALIAS', text(row['@alias']));
    if (info.collectionName === 'ZoneItems' && row['@alias'] !== undefined) setField(node, 'CP_CNVPERSALIAS', text(row['@alias']));
    if (info.collectionName === 'ZoneItems') ['CP_UKPPERS', 'CP_ENHCPRVCYIT', 'CP_RASDATAITEM', 'CP_TRUNKCHANLISTIT', 'CP_LCPSTLSTIT'].forEach(name => {
      if (row[fieldKey(name)] === undefined) return;
      const field = directChildren(node, 'field').find(item => item.getAttribute('name') === name);
      if (!field) return;
      const value = text(row[fieldKey(name)]);
      if (value) field.setAttribute('Name', value.trim().toUpperCase() === 'NONE' ? 'None' : value); else field.removeAttribute('Name');
    });
    if (info.category === 'CapacityPlusVoiceList' && row['@alias'] !== undefined) setField(node, 'TP_TRUNKPERALIAS', text(row['@alias']));
  }
  function buildContactForType(doc, candidates, type) {
    const collectionName = Object.entries(CONTACT_TYPE_LABELS).find(([, label]) => label === type)?.[0];
    if (!collectionName) throw new Error(`Unsupported Call Type "${type}".`);
    const callOwner = blueprintNode(doc, CPS_BLUEPRINTS.contacts[type]) || contactTemplateForType(candidates, type);
    const targetSource = callOwner && directChildren(callOwner, 'collection').find(collection => collection.getAttribute('name') === collectionName && directChildren(collection, 'set').length);
    if (!targetSource || isAllCallContact(callOwner)) throw new Error(`No regular ${type} contact blueprint is available in this source. The application will not invent a CPS record from another type.`);

    // Clone the complete same-type CPS record, then remove only the other known
    // call definitions. Unknown wrapper data from newer CPS versions is retained.
    const node = callOwner.cloneNode(true);
    for (const collection of directChildren(node, 'collection')) {
      if (CONTACT_COLLECTIONS.includes(collection.getAttribute('name')) && collection.getAttribute('name') !== collectionName) collection.replaceChildren();
    }
    const target = directChildren(node, 'collection').find(collection => collection.getAttribute('name') === collectionName);
    if (!target) throw new Error(`The CPS contact wrapper does not support ${type} contacts.`);
    return node;
  }
  function validateRegularContactId(name, id, type) {
    if (type === 'Quik Call II') return;
    if (!/^\d+$/.test(id)) throw new Error(`${name}: ${type} Call ID must be a whole number.`);
    const value = Number(id);
    if (!Number.isSafeInteger(value)) throw new Error(`${name}: Call ID is too large.`);
    if (type === 'Capacity Plus' && (value < 1 || value > 16776415 || value === 255)) throw new Error(`${name}: Capacity Plus Call ID must be from 1 through 16776415, other than 255, which is reserved for All Call.`);
    if (type === 'Digital' && (value < 1 || value > 16776415)) throw new Error(`${name}: Digital Call ID must be from 1 through 16776415; 16777215 is reserved for All Call.`);
    if (type === 'MDC' && (value < 0 || value > 65535)) throw new Error(`${name}: MDC Call ID must be from 0 through 65535.`);
  }
  function prepareContactNode(doc, node, row, candidates) {
    const type = requestedContactType(row.__contact_type);
    const name = text(row.__contact_name).trim(), id = text(row.__contact_id).trim();
    if (!name) throw new Error('Contacts: every row needs a Contact Name.');
    if (!type) throw new Error(`${name}: select a supported Call Type.`);
    if (!id) throw new Error(`${name}: enter a Call ID.`);

    if (node) {
      const types = activeContactTypes(node), originalType = primaryContactType(node);
      const originalName = directFieldValue(node, 'ContactName') || node.getAttribute('alias');
      const originalIds = contactIds(node).join(' / ');
      // CPS sometimes keeps one required scaffold contact with more than one
      // populated subtype. The simplified sheet cannot safely reinterpret its
      // combined ID/type, so preserve it unless only its name was changed.
      if (types.length > 1) {
        if (type !== originalType || id !== originalIds) throw new Error(`${originalName}: this CPS scaffold contains ${types.join(' and ')} definitions. Its Call Type and combined Call ID must remain unchanged.`);
        return node;
      }
      if (isCpsAllCallContact(node)) {
        if (type !== originalType || id !== originalIds || name !== originalName) throw new Error(`${originalName}: the CPS All Call record is reserved and cannot be edited.`);
        return node;
      }
      validateRegularContactId(name, id, type);
      if (type === originalType) return node;
      const replacement = buildContactForType(doc, candidates, type);
      node.replaceWith(replacement);
      return replacement;
    }
    // A new row with the system's reserved All Call ID becomes a complete CPS
    // All Call record. A second All Call for the same system is reported as a
    // duplicate Call ID by validateReferences.
    if (ALL_CALL_IDS[type] === id) {
      const allCall = blueprintNode(doc, CPS_BLUEPRINTS.allCallContacts?.[type]);
      if (!allCall) throw new Error(`The internal ${type} All Call blueprint is unavailable.`);
      return allCall;
    }
    validateRegularContactId(name, id, type);
    return buildContactForType(doc, candidates, type);
  }
  function updateContact(node, row) {
    const name = text(row.__contact_name).trim(), id = text(row.__contact_id).trim();
    syncPlanningNames({ category: 'PCRContacts' }, node, { 'field:ContactName': name });
    const types = activeContactTypes(node);
    if (types.length > 1) return;
    for (const call of contactCalls(node)) {
      if (call.collection === 'QuikCallIICalls') {
        const current = directFieldValue(call.node, 'PeudoCallId');
        if (id !== current) throw new Error(`${name}: Quik Call II codes cannot be safely derived from a single Call ID. Keep ${current} or configure this contact in CPS.`);
        continue;
      }
      for (const field of directChildren(call.node, 'field')) {
        if (['PeudoCallId', 'DU_CALLLSTID', 'CAPPLUSUCL_CALLLSTID', 'AU_CALLLSTID'].includes(field.getAttribute('name'))) field.textContent = id;
      }
    }
  }
  function isChannelPoolRecord(node) { return node?.getAttribute('key') === 'CHANNELPOOL' || node?.getAttribute('alias') === 'Channel Pool'; }
  function ensureChannelPool(doc) {
    let zoneCategory = directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === 'Zone');
    if (!zoneCategory) {
      zoneCategory = doc.createElement('category');
      zoneCategory.setAttribute('name', 'Zone');
      doc.documentElement.append(zoneCategory);
    }
    let pool = recordsForCategory(zoneCategory).find(isChannelPoolRecord);
    if (!pool) {
      pool = blueprintNode(doc, CPS_BLUEPRINTS.channelPool);
      if (!pool) throw new Error('The internal Channel Pool blueprint is unavailable.');
      zoneCategory.append(pool);
    }
    return pool;
  }
  function matchName(value) { return text(value).trim().toUpperCase().replace(/\s+/g, ''); }
  function parentFromVisibleLabel(doc, info, label) {
    if (!info.collectionName || !text(label).trim()) return null;
    const category = directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === info.category);
    if (!category) return null;
    return directChildren(category, 'set').find(node => {
      if (matchName(node.getAttribute('alias')) !== matchName(label)) return false;
      if (info.sheet === 'Channel Pool') return isChannelPoolRecord(node);
      if (info.sheet === 'Zone Channels') return !isChannelPoolRecord(node);
      return true;
    }) || null;
  }
  function recordMatchesSheet(node, info) {
    if (info.collectionName) {
      const collection = node.parentElement, parent = collection?.parentElement;
      if (!collection || collection.tagName !== 'collection' || collection.getAttribute('name') !== info.collectionName) return false;
      if (info.sheet === 'Channel Pool') return isChannelPoolRecord(parent);
      if (info.sheet === 'Zone Channels') return !isChannelPoolRecord(parent);
      return true;
    }
    const category = node.parentElement;
    if (!category || category.tagName !== 'category' || category.getAttribute('name') !== info.category) return false;
    return info.sheet !== 'Zones' || !isChannelPoolRecord(node);
  }
  function originalRecordsForSheet(info, byId) {
    return [...byId.entries()].filter(([id, node]) => id.startsWith('n') && recordMatchesSheet(node, info)).map(([, node]) => node);
  }
  function directFieldValue(node, name) { return directChildren(node, 'field').find(field => field.getAttribute('name') === name)?.textContent || ''; }
  function callId(node) { return contactCalls(node).flatMap(call => directChildren(call.node, 'field').filter(field => field.getAttribute('name') === 'PeudoCallId').map(field => field.textContent)).find(Boolean) || ''; }
  function findExistingRecord(doc, info, row, candidates, used, byId) {
    const available = candidates.filter(node => !used.has(node));
    const identified = byId.get(text(row._record_id));
    if (identified && available.includes(identified) && recordMatchesSheet(identified, info)) return identified;
    if (info.layout === 'contacts') {
      const name = text(row.__contact_name), id = text(row.__contact_id), type = requestedContactType(row.__contact_type);
      const usableId = id && !id.includes(' / ');
      const sameName = node => directFieldValue(node, 'ContactName') === name;
      const sameId = node => usableId && callId(node) === id;
      const sameType = node => !type || primaryContactType(node) === type;
      return available.find(node => sameName(node) && sameId(node) && sameType(node)) ||
        available.find(node => sameName(node) && sameType(node)) ||
        available.find(node => sameId(node) && sameType(node)) ||
        available.find(sameName) || available.find(sameId) || null;
    }
    if (info.collectionName === 'ZoneItems') {
      const alias = text(row['@alias']); const parent = parentFromVisibleLabel(doc, info, row.__parent_label);
      return available.find(node => node.getAttribute('alias') === alias && (!parent || node.parentElement?.parentElement === parent)) || available.find(node => node.getAttribute('alias') === alias) || null;
    }
    return available.find(node => node.getAttribute('alias') === text(row['@alias'])) || null;
  }
  function reorderSheetRecords(info, orderedNodes) {
    if (!orderedNodes.length) return;
    if (info.collectionName) {
      const byCollection = new Map();
      for (const node of orderedNodes) {
        const collection = node.parentElement;
        if (!byCollection.has(collection)) byCollection.set(collection, []);
        byCollection.get(collection).push(node);
      }
      for (const [collection, nodes] of byCollection) nodes.forEach(node => collection.append(node));
      return;
    }
    const category = orderedNodes[0].parentElement;
    if (category) {
      orderedNodes.forEach(node => category.append(node));
      // Channel Pool is intentionally not shown on the Zones sheet. Keep the
      // CPS-owned pool after the user-visible zones instead of accidentally
      // moving it to the front when zones are reordered.
      if (info.sheet === 'Zones') recordsForCategory(category).filter(isChannelPoolRecord).forEach(node => category.append(node));
    }
  }
  function templateForRow(doc, info, row, candidates) {
    if (info.collectionName === 'ZoneItems') {
      const type = text(row['field:CP_PERSTYPE']);
      if (!type) throw new Error(`New rows in ${info.sheet} need a Channel Type.`);
      const template = blueprintNode(doc, CPS_BLUEPRINTS.channels[type]) || candidates.find(node => directFieldValue(node, 'CP_PERSTYPE') === type);
      if (!template) throw new Error(`No exact ${displayValue('field:CP_PERSTYPE', type)} channel blueprint is available.`);
      return template;
    }
    if (info.sheet === 'Zones') return blueprintNode(doc, CPS_BLUEPRINTS.zone) || candidates[0];
    return candidates[0] || null;
  }
  function prepareChannelNode(doc, node, info, row, candidates, wasExisting) {
    const alias = text(row['@alias']).trim();
    const type = text(row['field:CP_PERSTYPE']).trim();
    if (!alias) throw new Error(`${info.sheet}: every channel row needs a Name / Alias.`);
    if (!CHANNEL_TYPE_LABELS[type]) throw new Error(`${alias}: select a supported Channel Type.`);
    const originalType = node ? directFieldValue(node, 'CP_PERSTYPE') : '';
    const originalValues = new Map(['CP_RXFREQ', 'CP_TXFREQ', 'CP_COLORCODE', 'CP_UKPPERS', 'CP_TRUNKCHANLISTIT', 'CP_LCPSTLSTIT'].map(name => [name, node ? directFieldValue(node, name).trim() : '']));
    const typeChanged = Boolean(node && originalType !== type);
    if (typeChanged) {
      const template = templateForRow(doc, info, row, candidates);
      const replacement = template.cloneNode(true);
      node.replaceWith(replacement);
      node = replacement;
    }
    const emptyMeansNone = new Set(['CP_ENHCPRVCYIT', 'CP_RASDATAITEM', 'CP_SCANLISTIT', 'CP_TGLISTIT', 'CP_EMSYSIT']);
    if (wasExisting && !typeChanged) {
      // CPS has no blank form for these references; clearing one means "none".
      for (const name of emptyMeansNone) {
        const key = fieldKey(name);
        if (key in row && !text(row[key]).trim() && directFieldValue(node, name).trim()) row[key] = 'NONE';
      }
    } else {
      const userRequired = new Set(['CP_RXFREQ', 'CP_TXFREQ', 'CP_COLORCODE', 'CP_UKPPERS', 'CP_TRUNKCHANLISTIT', 'CP_LCPSTLSTIT']);
      for (const key of Object.keys(row).filter(item => item.startsWith('field:'))) {
        const name = key.slice(6);
        const blueprintValue = directFieldValue(node, name);
        const blueprintHasField = directChildren(node, 'field').some(field => field.getAttribute('name') === name);
        if (!blueprintHasField) row[key] = '';
        else if (!text(row[key]).trim() && emptyMeansNone.has(name)) row[key] = 'NONE';
        else if (!text(row[key]).trim() && !userRequired.has(name)) row[key] = blueprintValue;
      }
    }
    const requiredMissing = (fieldName, value) => !value && (!wasExisting || typeChanged || Boolean(originalValues.get(fieldName)));
    const frequencyBased = ['ANLGCONV', 'DGTLCONV6PT25', 'CAPPLUSVOICECONV'].includes(type);
    for (const fieldName of ['CP_RXFREQ', 'CP_TXFREQ']) {
      const value = text(row[`field:${fieldName}`]).trim();
      if (frequencyBased && (requiredMissing(fieldName, value) || (value && (!Number.isFinite(Number(value)) || Number(value) <= 0)))) throw new Error(`${alias}: ${friendlyField(fieldName)} is required and must be a positive number.`);
      if (!frequencyBased && value && (!Number.isFinite(Number(value)) || Number(value) <= 0)) throw new Error(`${alias}: ${friendlyField(fieldName)} must be a positive number when populated.`);
    }
    const colorCode = text(row['field:CP_COLORCODE']).trim();
    if (['DGTLCONV6PT25', 'CAPPLUSVOICECONV'].includes(type) && (requiredMissing('CP_COLORCODE', colorCode) || (colorCode && (!/^\d+$/.test(colorCode) || Number(colorCode) > 15)))) throw new Error(`${alias}: Color Code is required and must be a whole number from 0 through 15.`);
    if (colorCode && (!/^\d+$/.test(colorCode) || Number(colorCode) > 15)) throw new Error(`${alias}: Color Code must be a whole number from 0 through 15.`);
    const contact = text(row['field:CP_UKPPERS']).trim();
    if (info.sheet === 'Channel Pool' && 'field:CP_UKPPERS' in row) {
      const hasContact = directChildren(node, 'field').some(field => field.getAttribute('name') === 'CP_UKPPERS');
      if (!hasContact && contact && contact.toUpperCase() !== 'NONE') throw new Error(`${alias}: ${displayValue('field:CP_PERSTYPE', type)} channels have no Contact Name in CPS. Leave it blank.`);
      // A pool channel's contact is optional; blank means none.
      if (hasContact && !contact) row['field:CP_UKPPERS'] = 'NONE';
      if (!hasContact) row['field:CP_UKPPERS'] = '';
    }
    if (info.sheet === 'Zone Channels' && ['DGTLCONV6PT25', 'CAPPLUSSYSCONV', 'LINKEDCAPPLUSSYSCONV'].includes(type) && (!contact || contact.toUpperCase() === 'NONE') && (!wasExisting || typeChanged || (originalValues.get('CP_UKPPERS') && originalValues.get('CP_UKPPERS').toUpperCase() !== 'NONE'))) throw new Error(`${alias}: ${displayValue('field:CP_PERSTYPE', type)} needs a Contact Name.`);
    const voiceList = text(row['field:CP_TRUNKCHANLISTIT']).trim();
    if (info.sheet === 'Zone Channels' && type === 'CAPPLUSSYSCONV' && (!voiceList || voiceList.toUpperCase() === 'NONE') && (!wasExisting || typeChanged || (originalValues.get('CP_TRUNKCHANLISTIT') && originalValues.get('CP_TRUNKCHANLISTIT').toUpperCase() !== 'NONE'))) throw new Error(`${alias}: Capacity Plus needs a Voice List.`);
    const siteList = text(row['field:CP_LCPSTLSTIT']).trim();
    if (info.sheet === 'Zone Channels' && type === 'LINKEDCAPPLUSSYSCONV' && (!siteList || siteList.toUpperCase() === 'NONE') && (!wasExisting || typeChanged || (originalValues.get('CP_LCPSTLSTIT') && originalValues.get('CP_LCPSTLSTIT').toUpperCase() !== 'NONE'))) throw new Error(`${alias}: Linked Capacity Plus needs a Site List.`);

    return node;
  }
  function assertVisibleRowApplied(info, node, row) {
    const fail = (field, expected, actual) => { throw new Error(`${info.sheet}: ${node.getAttribute('alias') || 'record'} did not round-trip ${field}. Excel has "${expected}" but generated XML has "${actual}".`); };
    if (info.layout === 'contacts') {
      const actual = {
        'Contact Name': directFieldValue(node, 'ContactName') || node.getAttribute('alias'),
        'Call ID': contactIds(node).join(' / '),
        'Call Type': primaryContactType(node)
      };
      const expected = { 'Contact Name': text(row.__contact_name).trim(), 'Call ID': text(row.__contact_id).trim(), 'Call Type': requestedContactType(row.__contact_type) };
      for (const name of Object.keys(expected)) if (actual[name] !== expected[name]) fail(name, expected[name], actual[name]);
      return;
    }
    for (const [key, expectedValue] of Object.entries(row)) {
      if (key.startsWith('label:') || key.startsWith('_')) continue;
      const expected = text(expectedValue);
      if (key.startsWith('@')) {
        const actual = node.getAttribute(key.slice(1));
        if (actual !== expected) fail(key.slice(1), expected, actual);
      } else if (key.startsWith('field:')) {
        const actual = directFieldValue(node, key.slice(6));
        if (actual !== expected) fail(friendlyField(key.slice(6)), expected, actual);
      }
    }
    if (info.collectionName === 'ZoneItems') {
      const type = directFieldValue(node, 'CP_PERSTYPE');
      if (node.getAttribute('key') !== type) fail('Record Type', type, node.getAttribute('key'));
      const expectedParent = text(row.__parent_label).trim();
      const actualParent = node.parentElement?.parentElement?.getAttribute('alias') || '';
      if (expectedParent && matchName(actualParent) !== matchName(expectedParent)) fail('Assigned Zone', expectedParent, actualParent);
    }
  }
  function importVoiceLists(doc, sourceRows) {
    const lists = new Map();
    for (const source of sourceRows) {
      const name = text(source['Voice List Name']).trim();
      const member = text(source['Member Channel']).trim();
      if (!name) {
        if (member) throw new Error('Voice Lists: every Member Channel needs a Voice List Name.');
        continue;
      }
      const key = matchName(name);
      if (!lists.has(key)) lists.set(key, { name, members: [] });
      const list = lists.get(key);
      if (list.name !== name) throw new Error(`Voice Lists: "${name}" conflicts with "${list.name}" after CPS name matching.`);
      if (member && list.members.some(value => matchName(value) === matchName(member))) throw new Error(`Voice List "${name}": Member Channel "${member}" appears more than once.`);
      if (member) list.members.push(member);
    }
    let category = directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === 'CapacityPlusVoiceList');
    if (!category && !lists.size) return;
    if (!category) {
      category = doc.createElement('category');
      category.setAttribute('name', 'CapacityPlusVoiceList');
      const zone = directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === 'Zone');
      doc.documentElement.insertBefore(category, zone || null);
    }
    const originals = recordsForCategory(category), matched = new Set();
    for (const { name, members } of lists.values()) {
      let node = originals.find(candidate => !matched.has(candidate) && matchName(candidate.getAttribute('alias')) === matchName(name));
      if (!node) {
        const template = originals[0];
        node = blueprintNode(doc, CPS_BLUEPRINTS.voiceList) || (template ? template.cloneNode(true) : null);
        if (!node) throw new Error('The internal Capacity Plus Voice List blueprint is unavailable.');
        category.append(node);
      } else matched.add(node);
      updateRecord(doc, node, {
        '@alias': name,
        'field:TP_TRUNKPERALIAS': name,
        'multifield:CapacityPlusVoiceListItems': JSON.stringify(members)
      });
    }
    originals.filter(node => !matched.has(node)).forEach(node => node.remove());
  }
  function importSiteLists(doc, sourceRows) {
    const lists = new Map();
    for (const source of sourceRows) {
      const name = text(source['Site Name']).trim();
      const siteAlias = text(source['Site Alias']).trim();
      const siteId = text(source['Site ID']).trim();
      const voiceList = text(source['Voice List']).trim();
      if (!name) {
        if (siteAlias || siteId || voiceList) throw new Error('Site Lists: every site row needs a Site Name.');
        continue;
      }
      const key = matchName(name);
      if (!lists.has(key)) lists.set(key, { name, sites: [] });
      const list = lists.get(key);
      if (list.name !== name) throw new Error(`Site Lists: "${name}" conflicts with "${list.name}" after CPS name matching.`);
      if (siteAlias || siteId || voiceList) {
        if (!siteAlias || !siteId || !voiceList) throw new Error(`Site List "${name}": Site Alias, Site ID, and Voice List are required together.`);
        list.sites.push({ siteAlias, siteId, voiceList });
      }
    }
    let category = directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === 'CapacityPlusSiteList');
    if (!category && !lists.size) return;
    if (!category) {
      category = doc.createElement('category');
      category.setAttribute('name', 'CapacityPlusSiteList');
      const zone = directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === 'Zone');
      doc.documentElement.insertBefore(category, zone || null);
    }
    const originals = recordsForCategory(category), matched = new Set();
    for (const { name, sites } of lists.values()) {
      let node = originals.find(candidate => !matched.has(candidate) && matchName(candidate.getAttribute('alias')) === matchName(name));
      let existingSites = [];
      if (!node) {
        const template = originals[0];
        node = blueprintNode(doc, CPS_BLUEPRINTS.siteList) || (template ? template.cloneNode(true) : null);
        if (!node) throw new Error('The internal Capacity Plus Site List blueprint is unavailable.');
        category.append(node);
      } else {
        matched.add(node);
        const collection = directChildren(node, 'collection').find(child => child.getAttribute('name') === 'CapacityPlusSiteListItems');
        existingSites = collection ? directChildren(collection, 'set') : [];
      }
      updateRecord(doc, node, { '@alias': name, 'field:LCPSTLST_ALIAS': name });
      const collection = findOrCreateCollection(node, 'CapacityPlusSiteListItems');
      const template = blueprintNode(doc, CPS_BLUEPRINTS.siteItem) || existingSites[0] || originals.flatMap(list => {
        const items = directChildren(list, 'collection').find(child => child.getAttribute('name') === 'CapacityPlusSiteListItems');
        return items ? directChildren(items, 'set') : [];
      })[0];
      const used = new Set(), rebuilt = [];
      for (const site of sites) {
        let siteNode = existingSites.find(candidate => !used.has(candidate) && (directFieldValue(candidate, 'LCPSTLST_SITEALIAS') === site.siteAlias || directFieldValue(candidate, 'LCPSTLST_SITEID') === site.siteId));
        if (!siteNode) {
          siteNode = template ? template.cloneNode(true) : null;
          if (!siteNode) throw new Error('The internal Capacity Plus Site blueprint is unavailable.');
        } else used.add(siteNode);
        updateRecord(doc, siteNode, {
          'field:LCPSTLST_SITEALIAS': site.siteAlias,
          'field:LCPSTLST_SITEID': site.siteId,
          'field:LCPSTLST_VCLSTIT': site.voiceList
        });
        const voiceField = directChildren(siteNode, 'field').find(field => field.getAttribute('name') === 'LCPSTLST_VCLSTIT');
        if (voiceField) voiceField.setAttribute('Name', site.voiceList);
        rebuilt.push(siteNode);
      }
      collection.replaceChildren(...rebuilt);
    }
    originals.filter(node => !matched.has(node)).forEach(node => node.remove());
  }
  function securityAliasesFromDoc(doc) {
    const rasAliases = new Map(), privacyAliases = new Map();
    const zoneCategory = directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === 'Zone');
    for (const zone of recordsForCategory(zoneCategory)) for (const collection of directChildren(zone, 'collection').filter(item => item.getAttribute('name') === 'ZoneItems')) for (const channel of directChildren(collection, 'set')) {
      const rasAlias = directFieldValue(channel, 'CP_RASDATAITEM');
      const privacyAlias = directFieldValue(channel, 'CP_ENHCPRVCYIT');
      if (rasAlias && rasAlias.toUpperCase() !== 'NONE') rasAliases.set(matchName(rasAlias), rasAlias);
      if (privacyAlias && privacyAlias.toUpperCase() !== 'NONE') privacyAliases.set(matchName(privacyAlias), privacyAlias);
    }
    return { rasAliases, privacyAliases };
  }
  function securityMapFromRows(sourceRows, originalAliases = { rasAliases: new Map(), privacyAliases: new Map() }) {
    const rasAliases = new Map(), privacyAliases = new Map();
    for (const row of sourceRows) {
      const rasAlias = text(row['RAS Key Alias']).trim();
      const privacyAlias = text(row['Privacy Key Alias']).trim();
      if (rasAlias && rasAlias.toUpperCase() !== 'NONE') {
        const key = matchName(rasAlias);
        if (rasAliases.has(key)) throw new Error(`Security: RAS Key Alias "${rasAlias}" appears more than once.`);
        rasAliases.set(key, rasAlias);
      }
      if (privacyAlias && privacyAlias.toUpperCase() !== 'NONE') {
        const key = matchName(privacyAlias);
        if (privacyAliases.has(key)) throw new Error(`Security: Privacy Key Alias "${privacyAlias}" appears more than once.`);
        privacyAliases.set(key, privacyAlias);
      }
    }
    for (const [key, alias] of originalAliases.rasAliases) if (!rasAliases.has(key)) rasAliases.set(key, alias);
    for (const [key, alias] of originalAliases.privacyAliases) if (!privacyAliases.has(key)) privacyAliases.set(key, alias);
    return { rasAliases, privacyAliases };
  }
  function validateSecurityAssignments(doc, securityMap) {
    if (!securityMap) return [];
    const zoneCategory = directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === 'Zone');
    const issues = [];
    for (const zone of recordsForCategory(zoneCategory)) for (const collection of directChildren(zone, 'collection').filter(item => item.getAttribute('name') === 'ZoneItems')) for (const channel of directChildren(collection, 'set')) {
      const channelName = channel.getAttribute('alias') || 'Unnamed channel';
      const privacyAlias = directFieldValue(channel, 'CP_ENHCPRVCYIT');
      const rasAlias = directFieldValue(channel, 'CP_RASDATAITEM');
      if (privacyAlias && privacyAlias.toUpperCase() !== 'NONE' && !securityMap.privacyAliases.has(matchName(privacyAlias))) issues.push(`${channelName}: Privacy Key Alias "${privacyAlias}" is not listed on the Security sheet.`);
      if (rasAlias && rasAlias.toUpperCase() !== 'NONE' && !securityMap.rasAliases.has(matchName(rasAlias))) issues.push(`${channelName}: RAS Key Alias "${rasAlias}" is not listed on the Security sheet.`);
    }
    return issues;
  }
  function validateReferences(doc) {
    const category = (name) => directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === name);
    const issues = [], contactTypes = new Map(), normalizedContacts = new Map(), contactIdsByType = new Map();
    for (const node of recordsForCategory(category('PCRContacts'))) {
      const name = directFieldValue(node, 'ContactName');
      if (!name) { issues.push('A contact has no Contact Name.'); continue; }
      const normalized = matchName(name);
      if (normalizedContacts.has(normalized)) issues.push(`Contact "${name}" appears more than once.`); else normalizedContacts.set(normalized, name);
      if (!contactTypes.has(name)) contactTypes.set(name, new Set());
      for (const type of activeContactTypes(node)) contactTypes.get(name).add(type);
      for (const call of contactCalls(node)) {
        const id = directFieldValue(call.node, 'PeudoCallId');
        if (!id) { issues.push(`${name}: ${CONTACT_TYPE_LABELS[call.collection] || call.collection} definition has no Call ID.`); continue; }
        // CPS permits the same numeric ID in different call namespaces (for
        // example a Digital Group Call and a Digital Private Call). Only flag
        // duplicates that also share their concrete CPS call subtype.
        const subtype = text(call.node.getAttribute('key')) || directFieldValue(call.node, 'CallType') || call.collection;
        const idKey = `${call.collection}|${subtype.toUpperCase()}|${id.toUpperCase()}`;
        if (contactIdsByType.has(idKey)) issues.push(`${name}: Call ID "${id}" duplicates ${contactIdsByType.get(idKey)} in ${CONTACT_TYPE_LABELS[call.collection] || call.collection}.`);
        else contactIdsByType.set(idKey, name);
      }
    }
    const contacts = new Set(contactTypes.keys());
    const voiceListRecords = recordsForCategory(category('CapacityPlusVoiceList'));
    const siteListRecords = recordsForCategory(category('CapacityPlusSiteList'));
    const voiceLists = new Set(voiceListRecords.map(node => node.getAttribute('alias')).filter(Boolean));
    const siteLists = new Set(siteListRecords.map(node => node.getAttribute('alias')).filter(Boolean));
    const zones = category('Zone'), zoneNames = new Set(), poolChannelNames = new Set();
    for (const zone of recordsForCategory(zones)) {
      const zoneName = zone.getAttribute('alias') || 'Unnamed zone';
      const zoneKey = matchName(zoneName);
      if (zoneNames.has(zoneKey)) issues.push(`Zone "${zoneName}" appears more than once.`); else zoneNames.add(zoneKey);
      for (const collection of directChildren(zone, 'collection').filter(node => node.getAttribute('name') === 'ZoneItems')) {
        const channelNames = new Set();
        for (const channel of directChildren(collection, 'set')) {
          const itemName = channel.getAttribute('alias') || 'Unnamed channel';
          const channelKey = matchName(itemName);
          if (channelNames.has(channelKey)) issues.push(`${zoneName}: Channel "${itemName}" appears more than once.`); else channelNames.add(channelKey);
          if (isChannelPoolRecord(zone) && channel.getAttribute('alias')) {
            if (poolChannelNames.has(channel.getAttribute('alias'))) issues.push(`Channel Pool: "${channel.getAttribute('alias')}" appears more than once.`);
            poolChannelNames.add(channel.getAttribute('alias'));
          }
          const value = (name) => directFieldValue(channel, name);
          const channelName = channel.getAttribute('alias') || 'Unnamed channel';
          const type = value('CP_PERSTYPE'), contact = value('CP_UKPPERS'), voiceList = value('CP_TRUNKCHANLISTIT'), siteList = value('CP_LCPSTLSTIT');
          if (!type || channel.getAttribute('key') !== type) issues.push(`${channelName}: Record Type does not match its Channel Type.`);
          const duplicateFields = directChildren(channel, 'field').map(field => field.getAttribute('name')).filter((name, index, names) => names.indexOf(name) !== index);
          if (duplicateFields.length) issues.push(`${channelName}: duplicate CPS fields are present (${[...new Set(duplicateFields)].join(', ')}).`);
          if (siteList && siteList !== 'NONE' && !siteLists.has(siteList)) issues.push(`${channelName}: Site List "${siteList}" is not present on the Site Lists sheet.`);
          if (contact && contact !== 'NONE' && !contacts.has(contact)) issues.push(`${channelName}: Contact Name "${contact}" is not present on the Contacts sheet.`);
          if (contact && contact !== 'NONE' && contacts.has(contact)) {
            const requiredType = isCapacityPlusPersonality(type) ? 'Capacity Plus' : type === 'DGTLCONV6PT25' ? 'Digital' : '';
            if (requiredType && !contactTypes.get(contact).has(requiredType)) issues.push(`${channelName}: ${displayValue('field:CP_PERSTYPE', type)} needs a ${requiredType} contact; "${contact}" is ${[...contactTypes.get(contact)].join(' / ') || 'not callable'}.`);
          }
          if (voiceList && voiceList !== 'NONE' && !voiceLists.has(voiceList)) issues.push(`${channelName}: Voice List "${voiceList}" is not present in this codeplug.`);
          for (const fieldName of ['CP_RXFREQ', 'CP_TXFREQ']) {
            const frequency = value(fieldName);
            if (frequency && (!Number.isFinite(Number(frequency)) || Number(frequency) <= 0)) issues.push(`${channelName}: ${friendlyField(fieldName)} is invalid.`);
          }
          const colorCode = value('CP_COLORCODE');
          if (colorCode && (!/^\d+$/.test(colorCode) || Number(colorCode) > 15)) issues.push(`${channelName}: Color Code must be 0 through 15.`);
        }
      }
    }
    const voiceListNames = new Set();
    for (const list of voiceListRecords) {
      const name = list.getAttribute('alias') || 'Unnamed voice list';
      const nameKey = matchName(name);
      if (voiceListNames.has(nameKey)) issues.push(`Voice List "${name}" appears more than once.`); else voiceListNames.add(nameKey);
      const members = directChildren(list, 'multifield').find(field => field.getAttribute('name') === 'CapacityPlusVoiceListItems');
      for (const member of members ? directChildren(members, 'value').map(value => value.textContent) : []) {
        if (member && !poolChannelNames.has(member)) issues.push(`Voice List "${name}": Member Channel "${member}" is not present in the Channel Pool.`);
      }
    }
    const siteListNames = new Set();
    for (const list of siteListRecords) {
      const name = list.getAttribute('alias') || 'Unnamed site list';
      const nameKey = matchName(name);
      if (siteListNames.has(nameKey)) issues.push(`Site List "${name}" appears more than once.`); else siteListNames.add(nameKey);
      const siteNames = new Set(), siteIds = new Set();
      const collection = directChildren(list, 'collection').find(node => node.getAttribute('name') === 'CapacityPlusSiteListItems');
      for (const site of collection ? directChildren(collection, 'set') : []) {
        const siteName = directFieldValue(site, 'LCPSTLST_SITEALIAS');
        const siteId = directFieldValue(site, 'LCPSTLST_SITEID');
        const voiceList = directFieldValue(site, 'LCPSTLST_VCLSTIT');
        const siteNameKey = matchName(siteName);
        if (siteName && siteNames.has(siteNameKey)) issues.push(`Site List "${name}": Site Name "${siteName}" appears more than once.`); else if (siteName) siteNames.add(siteNameKey);
        if (siteId && siteIds.has(siteId)) issues.push(`Site List "${name}": Site ID "${siteId}" appears more than once.`); else if (siteId) siteIds.add(siteId);
        if (voiceList && voiceList !== 'NONE' && !voiceLists.has(voiceList)) issues.push(`Site List "${name}": Voice List "${voiceList}" is not present on the Voice Lists sheet.`);
      }
    }
    return issues;
  }
  function nameLengthIssues(doc) {
    const category = (name) => directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === name);
    const issues = [];
    const check = (label, name) => {
      if (name.length > CPS_NAME_LENGTH) issues.push(`${label} "${name}" is ${name.length} characters; CPS allows ${CPS_NAME_LENGTH} and would cut it to "${name.slice(0, CPS_NAME_LENGTH)}".`);
    };
    for (const contact of recordsForCategory(category('PCRContacts'))) check('Contact', directFieldValue(contact, 'ContactName') || text(contact.getAttribute('alias')));
    for (const zone of recordsForCategory(category('Zone'))) {
      const zoneName = text(zone.getAttribute('alias'));
      if (!isChannelPoolRecord(zone)) check('Zone', zoneName);
      for (const collection of directChildren(zone, 'collection').filter(node => node.getAttribute('name') === 'ZoneItems')) {
        for (const channel of directChildren(collection, 'set')) check(`${zoneName}: Channel`, text(channel.getAttribute('alias')));
      }
    }
    return issues;
  }
  function channelSchemaSignature(node) {
    return [...node.children].map(child => {
      const name = child.getAttribute('name');
      return `${child.tagName}:${name}:${child.hasAttribute('Name') ? 'display' : ''}`;
    }).join('|');
  }
  function channelSchemasFromDoc(doc) {
    const schemas = new Map();
    const zone = directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === 'Zone');
    for (const parent of recordsForCategory(zone)) for (const collection of directChildren(parent, 'collection').filter(node => node.getAttribute('name') === 'ZoneItems')) for (const channel of directChildren(collection, 'set')) {
      const type = directFieldValue(channel, 'CP_PERSTYPE');
      if (!type) continue;
      if (!schemas.has(type)) schemas.set(type, new Set());
      schemas.get(type).add(channelSchemaSignature(channel));
    }
    return schemas;
  }
  function addInternalChannelSchemas(doc, schemas) {
    for (const [type, xml] of Object.entries(CPS_BLUEPRINTS.channels)) {
      const channel = blueprintNode(doc, xml);
      if (!channel) continue;
      if (!schemas.has(type)) schemas.set(type, new Set());
      schemas.get(type).add(channelSchemaSignature(channel));
    }
    return schemas;
  }
  function validateChannelSchemas(doc, sourceSchemas) {
    const issues = [], zone = directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === 'Zone');
    for (const parent of recordsForCategory(zone)) for (const collection of directChildren(parent, 'collection').filter(node => node.getAttribute('name') === 'ZoneItems')) for (const channel of directChildren(collection, 'set')) {
      const name = channel.getAttribute('alias') || 'Unnamed channel', type = directFieldValue(channel, 'CP_PERSTYPE');
      const allowed = sourceSchemas.get(type);
      if (!allowed) issues.push(`${name}: no trusted source schema exists for ${displayValue('field:CP_PERSTYPE', type) || type || 'this channel type'}.`);
      else if (!allowed.has(channelSchemaSignature(channel))) issues.push(`${name}: its CPS field structure does not match a trusted ${displayValue('field:CP_PERSTYPE', type)} source record.`);
    }
    return issues;
  }
  function recordsForCategory(category) { return category ? directChildren(category, 'set') : []; }
  function placeDependentListsAfterZones(doc) {
    const config = doc.documentElement;
    const zone = directChildren(config, 'category').find(node => node.getAttribute('name') === 'Zone');
    const voiceLists = directChildren(config, 'category').find(node => node.getAttribute('name') === 'CapacityPlusVoiceList');
    const siteLists = directChildren(config, 'category').find(node => node.getAttribute('name') === 'CapacityPlusSiteList');
    const isAfter = (node, anchor) => directChildren(config, 'category').indexOf(node) > directChildren(config, 'category').indexOf(anchor);
    if (zone && voiceLists && !isAfter(voiceLists, zone)) config.insertBefore(voiceLists, zone.nextSibling);
    const siteAnchor = zone && voiceLists && isAfter(voiceLists, zone) ? voiceLists : zone;
    if (siteAnchor && siteLists && !isAfter(siteLists, siteAnchor)) config.insertBefore(siteLists, siteAnchor.nextSibling);
  }
  function planningCounts(doc) {
    const category = (name) => directChildren(doc.documentElement, 'category').find(node => node.getAttribute('name') === name);
    const contacts = recordsForCategory(category('PCRContacts')).length;
    const zones = recordsForCategory(category('Zone')).filter(node => !isChannelPoolRecord(node));
    const channels = zones.flatMap(zone => directChildren(zone, 'collection').filter(collection => collection.getAttribute('name') === 'ZoneItems').flatMap(collection => directChildren(collection, 'set'))).length;
    return { contacts, zones: zones.length, channels };
  }
  function importWorkbook(file) {
    $('copy-output').disabled = true; $('download-xml').disabled = true; output.value = '';
    const reader = new FileReader(); reader.onload = () => {
      try {
        requireXlsx(); const wb = XLSX.read(reader.result, { type: 'array', raw: false });
        if (!wb.SheetNames.includes('_CPS_TREE') || !wb.SheetNames.includes('_CPS_META')) throw new Error('This is not an unArchive export (hidden sheets are missing).');
        const { doc, byId } = treeFromRows(rowsFromSheet(wb, '_CPS_TREE'));
        const before = planningCounts(doc);
        // CPS-generated templates can contain intentional records that look
        // duplicate or incomplete to a generic validator. A no-change
        // round-trip must preserve and accept those. Report only issues newly
        // introduced by workbook edits.
        const originalReferenceIssues = new Set(validateReferences(doc));
        const sourceChannelSchemas = addInternalChannelSchemas(doc, channelSchemasFromDoc(doc));
        const originalSecurityAliases = securityAliasesFromDoc(doc);
        const meta = rowsFromSheet(wb, '_CPS_META');
        let verifiedRows = 0; const seenVisibleRecordIds = new Set();
        const securityInfo = meta.find(info => info.layout === 'security');
        const securityMap = securityInfo && wb.Sheets[securityInfo.sheet] ? securityMapFromRows(rowsFromSheet(wb, securityInfo.sheet), originalSecurityAliases) : null;
        for (const info of meta) {
          if (!wb.Sheets[info.sheet]) continue;
          const sourceRows = rowsFromSheet(wb, info.sheet);
          if (info.layout === 'settings') { sourceRows.forEach(row => updateSetting(doc, byId, row)); continue; }
          if (info.layout === 'voiceLists') { importVoiceLists(doc, sourceRows); verifiedRows += sourceRows.length; continue; }
          if (info.layout === 'siteLists') { importSiteLists(doc, sourceRows); verifiedRows += sourceRows.length; continue; }
          if (info.layout === 'security') { verifiedRows += sourceRows.length; continue; }
          let columnMap = {}; try { columnMap = JSON.parse(info.column_map_json || '{}'); } catch { throw new Error(`The workbook field map for ${info.sheet} is invalid.`); }
          const originals = originalRecordsForSheet(info, byId), matched = new Set(), orderedNodes = [];
          for (const source of sourceRows) {
            const row = Object.fromEntries(Object.entries(source).map(([header, value]) => { const technical = columnMap[header] || header; return [technical, cpsValue(technical, value)]; }));
            const workbookId = text(row._record_id).trim();
            if (workbookId) {
              if (seenVisibleRecordIds.has(workbookId)) throw new Error(`${info.sheet}: hidden record ID ${workbookId} appears more than once. Add new records using a blank row instead of copying an entire existing row.`);
              seenVisibleRecordIds.add(workbookId);
              const identified = byId.get(workbookId);
              if (!identified || !originals.includes(identified) || !recordMatchesSheet(identified, info)) throw new Error(`${info.sheet}: hidden record ID ${workbookId} does not identify a valid record on this sheet.`);
            }
            if (info.sheet === 'Zones' && !text(row['@alias']).trim()) throw new Error('Zones: every row needs a Name / Alias.');
            const originalNode = findExistingRecord(doc, info, row, originals, matched, byId);
            let node = originalNode;
            if (info.layout === 'contacts' && node) node = prepareContactNode(doc, node, row, originals);
            let visibleParent = parentFromVisibleLabel(doc, info, row.__parent_label);
            if (info.sheet === 'Channel Pool') {
              visibleParent = visibleParent || ensureChannelPool(doc);
              row.__parent_label = visibleParent.getAttribute('alias') || 'Channel Pool';
            }
            if (info.collectionName && text(row.__parent_label).trim() && !visibleParent) throw new Error(`${info.sheet}: Assigned Zone "${row.__parent_label}" does not match a Zone name.`);
            if (!node) {
              let parent = visibleParent;
              if (!parent && !info.collectionName) parent = directChildren(doc.documentElement, 'category').find(c => c.getAttribute('name') === info.category);
              if (!parent) throw new Error(`New row in ${info.sheet} needs an Assigned Zone that exactly matches a zone name.`);
              if (info.collectionName) parent = findOrCreateCollection(parent, info.collectionName);
              if (info.layout === 'contacts') node = prepareContactNode(doc, null, row, originals);
              else {
                const template = templateForRow(doc, info, row, originals);
                node = template ? template.cloneNode(true) : doc.createElement('set');
              }
              if (info.sheet === 'Zones') {
                const items = directChildren(node, 'collection').find(collection => collection.getAttribute('name') === 'ZoneItems');
                if (items) items.replaceChildren();
              }
              parent.append(node);
            }
            if (visibleParent && info.collectionName) {
              const target = findOrCreateCollection(visibleParent, info.collectionName);
              if (node.parentElement !== target) target.append(node);
            }
            if (info.collectionName === 'ZoneItems') node = prepareChannelNode(doc, node, info, row, originals, Boolean(originalNode));
            if (info.layout === 'contacts') updateContact(node, row);
            else { updateRecord(doc, node, row); syncPlanningNames(info, node, row); }
            assertVisibleRowApplied(info, node, row);
            verifiedRows++;
            if (originalNode && originals.includes(originalNode)) matched.add(originalNode);
            orderedNodes.push(node);
          }
          originals.filter(node => !matched.has(node)).forEach(node => node.remove());
          reorderSheetRecords(info, orderedNodes);
        }
        placeDependentListsAfterZones(doc);
        const serialized = '<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' + new XMLSerializer().serializeToString(doc.documentElement);
        output.value = serialized;
        const newReferenceIssues = validateReferences(doc).filter(issue => !originalReferenceIssues.has(issue));
        const issues = [...nameLengthIssues(doc), ...newReferenceIssues, ...validateChannelSchemas(doc, sourceChannelSchemas), ...validateSecurityAssignments(doc, securityMap)];
        if (issues.length) {
          output.value = '';
          $('copy-output').disabled = true; $('download-xml').disabled = true;
          message(outputStatus, `Fix these before copying: ${issues.slice(0, 4).join('  ')}${issues.length > 4 ? ` (+${issues.length - 4} more)` : ''}`, 'error');
        } else {
          $('copy-output').disabled = false; $('download-xml').disabled = false;
          const after = planningCounts(doc);
          message(outputStatus, `Ready to paste: verified ${verifiedRows} visible workbook rows. ${after.contacts} contacts, ${after.zones} zones, ${after.channels} zone channels (was ${before.contacts}, ${before.zones}, ${before.channels}).`, 'ok');
        }
      } catch (err) { message(outputStatus, err.message || 'Unable to import workbook.', 'error'); }
    }; reader.onerror = () => message(outputStatus, 'Could not read that workbook.', 'error'); reader.readAsArrayBuffer(file);
  }
  input.addEventListener('input', () => { currentModel = null; $('export-workbook').disabled = !input.value.trim(); $('download-source').disabled = !input.value.trim(); });
  input.addEventListener('change', () => { if (input.value.trim()) prepareModelFromInput(); });
  // Reset each picker after taking its file so choosing the same file again
  // (for example after fixing an import error) still fires a change event.
  $('xml-file').addEventListener('change', e => { const file = e.target.files[0]; e.target.value = ''; if (!file) return; const r = new FileReader(); r.onload = () => { input.value = r.result; prepareModelFromInput(); }; r.readAsText(file); });
  $('export-workbook').addEventListener('click', exportWorkbook);
  $('template-enabled').addEventListener('change', showTemplatePreview);
  ['input', 'change'].forEach(type => $('template-zones').addEventListener(type, () => { renderTemplateZones(); showTemplatePreview(); }));
  $('template-clear-pool').addEventListener('change', showTemplatePreview);
  renderTemplateZones();
  renderTemplateContacts();
  $('paste-clipboard').addEventListener('click', async () => {
    try {
      input.value = await readClipboardText();
      if (!input.value.trim()) throw new Error('Your clipboard is empty.');
      $('download-source').disabled = false; prepareModelFromInput();
    } catch (err) {
      input.focus(); message(xmlStatus, `${err.message || 'Clipboard access was blocked.'} Click the box and press Ctrl+V instead.`, 'error');
    }
  });
  $('download-source').addEventListener('click', async () => { try { const saved = await saveFile(input.value, 'cps-clipboard-source.xml', { name: 'CPS XML', extensions: ['xml'] }); message(xmlStatus, saved ? 'Source XML saved.' : 'Source XML save cancelled.', saved ? 'ok' : ''); } catch (err) { message(xmlStatus, err.message, 'error'); } });
  $('workbook-file').addEventListener('change', e => { const file = e.target.files[0]; e.target.value = ''; if (file) importWorkbook(file); });
  $('copy-output').addEventListener('click', async () => { try { await writeClipboardText(output.value); message(outputStatus, 'XML copied to your clipboard.', 'ok'); } catch { output.select(); document.execCommand('copy'); message(outputStatus, 'XML selected—use Ctrl+C to copy it.', 'ok'); } });
  $('download-xml').addEventListener('click', async () => { try { const saved = await saveFile(output.value, 'codeplug-rebuilt.xml', { name: 'CPS XML', extensions: ['xml'] }); message(outputStatus, saved ? 'Rebuilt XML saved.' : 'Rebuilt XML save cancelled.', saved ? 'ok' : ''); } catch (err) { message(outputStatus, err.message, 'error'); } });

  // Updates come from the signed latest.json published with each GitHub
  // release; the updater refuses any installer not signed with our key. The
  // startup check is quiet, so only a manual check reports "up to date" or a
  // network problem.
  const updateStatus = $('update-status'), updateCheck = $('update-check'), updateInstall = $('update-install');
  let pendingUpdate = null;
  async function checkUpdates(manual) {
    updateCheck.disabled = true;
    if (manual) message(updateStatus, 'Checking for updates…');
    try {
      pendingUpdate = await checkForUpdate();
      updateInstall.hidden = !pendingUpdate;
      if (pendingUpdate) message(updateStatus, `Version ${pendingUpdate.version} is available.`, 'ok');
      else if (manual) message(updateStatus, 'You have the latest version.', 'ok');
    } catch {
      if (manual) message(updateStatus, 'Could not check for updates. Check your internet connection and try again.', 'error');
    } finally { updateCheck.disabled = false; }
  }
  async function installUpdate() {
    if (!pendingUpdate) return;
    const proceed = await ask(`unArchive ${pendingUpdate.version} will download and install now. unArchive closes during the install, so save any workbook or XML you still need first.`, { title: 'Install update', kind: 'info', okLabel: 'Install', cancelLabel: 'Not now' });
    if (!proceed) return;
    updateInstall.disabled = true; updateCheck.disabled = true;
    try {
      let total = 0, received = 0;
      await pendingUpdate.downloadAndInstall(event => {
        if (event.event === 'Started') total = event.data.contentLength || 0;
        if (event.event === 'Progress') { received += event.data.chunkLength; message(updateStatus, total ? `Downloading update… ${Math.round(received / total * 100)}%` : 'Downloading update…'); }
        if (event.event === 'Finished') message(updateStatus, 'Installing update…');
      });
    } catch (err) {
      message(updateStatus, `The update could not be installed: ${err?.message || err}`, 'error');
      updateInstall.disabled = false; updateCheck.disabled = false;
    }
  }
  (async () => {
    let version = '';
    try { version = await getVersion(); } catch { /* Not running inside the desktop app. */ }
    if (!version) return;
    $('app-version').textContent = `unArchive v${version}`;
    updateCheck.hidden = false;
    updateCheck.addEventListener('click', () => checkUpdates(true));
    updateInstall.addEventListener('click', installUpdate);
    checkUpdates(false);
  })();
})();
