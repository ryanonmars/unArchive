import fs from 'node:fs';
import path from 'node:path';
import { DOMParser } from 'linkedom';

const root = path.resolve(import.meta.dirname, '..');
const sourceFiles = [
  'cps-clipboard-source.xml',
  'cps-clipboard-source-linked-plus-analog.xml'
];
const documents = sourceFiles.map(file => new DOMParser().parseFromString(fs.readFileSync(path.join(root, file), 'utf8'), 'application/xml'));
const direct = (node, tag) => [...(node?.children || [])].filter(child => !tag || child.tagName === tag);
const category = (doc, name) => direct(doc.documentElement, 'category').find(node => node.getAttribute('name') === name);
const records = (doc, name) => direct(category(doc, name), 'set');
const field = (node, name) => direct(node, 'field').find(item => item.getAttribute('name') === name)?.textContent || '';
const allChannels = (doc) => records(doc, 'Zone').flatMap(zone => direct(zone, 'collection')
  .filter(collection => collection.getAttribute('name') === 'ZoneItems')
  .flatMap(collection => direct(collection, 'set')));
const activeContactCollection = (contact, collectionName) => direct(contact, 'collection')
  .find(collection => collection.getAttribute('name') === collectionName && direct(collection, 'set').length);
const isAllCall = (contact) => direct(contact, 'collection').flatMap(collection => direct(collection, 'set')).some(call =>
  ['DGTLALLSYSCALL', 'CAPPLUSDGTLALLSYSCALL'].includes((call.getAttribute('key') || '').toUpperCase()) || field(call, 'CallType').toUpperCase().endsWith('-ALL CALL'));
const serialize = node => node?.toString() || '';

const contactTypes = {
  Digital: 'DigitalCalls',
  'Capacity Plus': 'CapacityPlusCalls',
  MDC: 'MDCCalls',
  'Quik Call II': 'QuikCallIICalls'
};
const channelTypes = ['ANLGCONV', 'DGTLCONV6PT25', 'CAPPLUSSYSCONV', 'LINKEDCAPPLUSSYSCONV', 'CAPPLUSVOICECONV'];
const contacts = Object.fromEntries(Object.entries(contactTypes).map(([label, collectionName]) => {
  const template = documents.flatMap(doc => records(doc, 'PCRContacts')).find(contact => activeContactCollection(contact, collectionName) && !isAllCall(contact));
  if (!template) throw new Error(`Missing ${label} contact blueprint.`);
  return [label, serialize(template)];
}));
// Fixture channels reference contacts, keys, and lists that belong to the
// fixture codeplug. A new channel must not silently inherit them, so reset
// every per-codeplug reference to CPS's own "nothing selected" form and turn
// privacy off. The field set and display-label attributes are left intact so
// the record keeps the exact CPS schema.
const channelReferenceFields = [
  'CP_UKPPERS', 'CP_ENHCPRVCYIT', 'CP_RASDATAITEM', 'CP_SCNROAMLISTIT', 'CP_TRUNKCHANLISTIT',
  'CP_LCPSTLSTIT', 'CP_TGLISTIT', 'CP_EMSYSIT', 'CP_DATAREVCHANLISTIT', 'CP_PHONESYSIT',
  'CP_CVFNLITEM', 'CP_MANDOWNPROFILEITEM', 'CP_TXSIGSYSIT', 'CP_RXSIGSYSIT', 'CP_SCANLISTIT'
];
// Enumerated settings that use the same "...IT" naming but are not references.
const channelEnumeratedFields = new Set(['CP_GPSRVRTPERSIT']);
const neutralChannel = (channel, type) => {
  const clone = channel.cloneNode(true);
  for (const item of direct(clone, 'field')) {
    const name = item.getAttribute('name');
    if (channelReferenceFields.includes(name)) {
      item.textContent = 'NONE';
      if (item.hasAttribute('Name')) item.setAttribute('Name', 'None');
    } else if (name === 'CP_VOICEPRIVACYEN') item.textContent = 'False';
    else if (item.hasAttribute('Name') && /(IT|ITEM)$/.test(name) && !channelEnumeratedFields.has(name) && item.textContent.toUpperCase() !== 'NONE') {
      throw new Error(`${type} blueprint field ${name} references "${item.textContent}". Classify it in generate-blueprints.mjs.`);
    }
  }
  return clone;
};
const channels = Object.fromEntries(channelTypes.map(type => {
  const template = documents.flatMap(allChannels).find(channel => field(channel, 'CP_PERSTYPE') === type);
  if (!template) throw new Error(`Missing ${type} channel blueprint.`);
  return [type, serialize(neutralChannel(template, type))];
}));
const regularZones = documents.flatMap(doc => records(doc, 'Zone')).filter(zone => zone.getAttribute('key') !== 'CHANNELPOOL' && zone.getAttribute('alias') !== 'Channel Pool');
const pool = documents.flatMap(doc => records(doc, 'Zone')).find(zone => zone.getAttribute('key') === 'CHANNELPOOL' || zone.getAttribute('alias') === 'Channel Pool');
const voiceList = documents.flatMap(doc => records(doc, 'CapacityPlusVoiceList'))[0];
const siteList = documents.flatMap(doc => records(doc, 'CapacityPlusSiteList'))[0];
const siteItem = direct(siteList, 'collection').find(collection => collection.getAttribute('name') === 'CapacityPlusSiteListItems');
const siteItemSet = direct(siteItem, 'set')[0];
if (!regularZones[0] || !pool || !voiceList || !siteList || !siteItemSet) throw new Error('One or more structural CPS blueprints are missing.');

const emptyCollection = (node, name) => {
  const clone = node.cloneNode(true);
  const collection = direct(clone, 'collection').find(item => item.getAttribute('name') === name);
  if (collection) collection.replaceChildren();
  return clone;
};
const voiceListShell = voiceList.cloneNode(true);
const voiceMembers = direct(voiceListShell, 'multifield').find(item => item.getAttribute('name') === 'CapacityPlusVoiceListItems');
if (voiceMembers) voiceMembers.replaceChildren();

// The generated file is committed, so replace every value that came from the
// fixture codeplugs (names, call IDs, frequencies) with a generic placeholder.
// Import always overwrites these from the workbook, so the CPS schema and
// defaults are all that the blueprints need to carry.
const NAME_FIELDS = new Set([
  'ContactName', 'DU_CALLALIAS', 'CAPPLUSUCL_CALLALIAS', 'AU_CALLALIAS', 'QU_CALLENTRYALIAS',
  'CP_CNVPERSALIAS', 'ZP_ZONEALIAS', 'TP_TRUNKPERALIAS', 'LCPSTLST_ALIAS', 'LCPSTLST_SITEALIAS'
]);
const ID_FIELDS = new Set(['DU_CALLLSTID', 'CAPPLUSUCL_CALLLSTID', 'LCPSTLST_SITEID']);
const FREQUENCY_FIELDS = new Set(['CP_RXFREQ', 'CP_TXFREQ']);
const scrub = (xml, placeholder) => {
  const node = new DOMParser().parseFromString(`<config>${xml}</config>`, 'application/xml').documentElement.firstElementChild;
  const visit = (element, inRegularCall) => {
    if (element.hasAttribute('alias') && element.getAttribute('alias') !== 'Channel Pool') element.setAttribute('alias', placeholder);
    const name = element.getAttribute('name');
    if (element.tagName === 'field' && element.textContent !== 'Channel Pool') {
      if (NAME_FIELDS.has(name)) element.textContent = placeholder;
      else if (ID_FIELDS.has(name) || (inRegularCall && name === 'PeudoCallId')) element.textContent = '1';
      else if (FREQUENCY_FIELDS.has(name)) element.textContent = '450.0000';
    }
    const regularCall = inRegularCall || (element.tagName === 'collection' && ['DigitalCalls', 'CapacityPlusCalls'].includes(name));
    for (const child of direct(element)) visit(child, regularCall);
  };
  visit(node, false);
  return serialize(node);
};
const scrubAll = (values, placeholder) => Object.fromEntries(Object.entries(values).map(([key, xml]) => [key, scrub(xml, placeholder)]));

const blueprints = {
  contacts: scrubAll(contacts, 'CONTACT'),
  channels: scrubAll(channels, 'CHANNEL'),
  zone: scrub(serialize(emptyCollection(regularZones[0], 'ZoneItems')), 'ZONE'),
  channelPool: scrub(serialize(emptyCollection(pool, 'ZoneItems')), 'ZONE'),
  voiceList: scrub(serialize(voiceListShell), 'VOICE LIST'),
  siteList: scrub(serialize(emptyCollection(siteList, 'CapacityPlusSiteListItems')), 'SITE LIST'),
  siteItem: scrub(serialize(siteItemSet), 'SITE')
};
const output = `// Generated by scripts/generate-blueprints.mjs from CPS-created fixtures, with fixture names, IDs, and frequencies replaced by placeholders.\nexport const CPS_BLUEPRINTS = ${JSON.stringify(blueprints, null, 2)};\n`;
fs.writeFileSync(path.join(root, 'blueprints.generated.js'), output, 'utf8');
console.log('Generated internal CPS blueprint library.');
