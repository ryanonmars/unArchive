// A small synthetic CPS codeplug built only from the committed blueprints, so
// the tracked tests need no customer codeplug. Names, call IDs, and
// frequencies are invented; every record keeps its exact CPS blueprint schema.
import { DOMParser } from 'linkedom';
import { CPS_BLUEPRINTS } from '../blueprints.generated.js';

const NAME_FIELDS = ['ContactName', 'DU_CALLALIAS', 'CAPPLUSUCL_CALLALIAS'];
const ID_FIELDS = ['DU_CALLLSTID', 'CAPPLUSUCL_CALLLSTID', 'PeudoCallId'];

function node(doc, xml) {
  const parsed = new DOMParser().parseFromString(`<config>${xml}</config>`, 'application/xml');
  return doc.importNode ? doc.importNode(parsed.documentElement.firstElementChild, true) : parsed.documentElement.firstElementChild;
}
function descendants(element, tag) { return [...element.getElementsByTagName(tag)]; }
function children(element, tag) { return [...element.children].filter(child => child.tagName === tag); }
function setFields(element, names, value) {
  for (const field of descendants(element, 'field')) if (names.includes(field.getAttribute('name'))) field.textContent = value;
}
// A reference field keeps CPS's display label in its Name attribute.
function setField(element, name, value) {
  const field = children(element, 'field').find(item => item.getAttribute('name') === name);
  if (!field) throw new Error(`Synthetic fixture: ${name} is not a field of this record.`);
  field.textContent = value;
  if (field.hasAttribute('Name')) field.setAttribute('Name', value === 'NONE' ? 'None' : value);
}

function contact(doc, xml, name, id) {
  const record = node(doc, xml);
  record.setAttribute('alias', name);
  setFields(record, NAME_FIELDS, name);
  if (id) setFields(record, ID_FIELDS, id);
  return record;
}

const CHANNEL_FIELDS = {
  rx: 'CP_RXFREQ', tx: 'CP_TXFREQ', colorCode: 'CP_COLORCODE', contact: 'CP_UKPPERS', voiceList: 'CP_TRUNKCHANLISTIT',
  siteList: 'CP_LCPSTLSTIT', scanList: 'CP_SCNROAMLISTIT', receiveOnly: 'CP_RXONLYEN',
  privacy: 'CP_VOICEPRIVACYEN', privacyKey: 'CP_ENHCPRVCYIT', rasKey: 'CP_RASDATAITEM'
};
function channel(doc, type, name, values = {}) {
  const record = node(doc, CPS_BLUEPRINTS.channels[type]);
  record.setAttribute('alias', name);
  setField(record, 'CP_CNVPERSALIAS', name);
  for (const [key, value] of Object.entries(values)) setField(record, CHANNEL_FIELDS[key], value);
  return record;
}

function zone(doc, xml, name, channels) {
  const record = node(doc, xml);
  if (name) { record.setAttribute('alias', name); setField(record, 'ZP_ZONEALIAS', name); }
  const items = children(record, 'collection').find(item => item.getAttribute('name') === 'ZoneItems');
  channels.forEach(item => items.append(item));
  return record;
}

function categoryOf(doc, name, records) {
  const element = doc.createElement('category');
  element.setAttribute('name', name);
  records.forEach(record => element.append(record));
  return element;
}

// Each channel type and its Receive Only and Scan List values, chosen so a
// test can tell a kept value from a default.
// With keys, the TALK channels of each type that has privacy carry invented
// privacy and RAS key aliases. With poolKeys, the Digital pool channel carries
// keys while Privacy Enabled stays False, as a channel set up in CPS can.
export function syntheticCodeplug({ pool = true, roamListChannel = false, keys = false, poolKeys = false } = {}) {
  const doc = new DOMParser().parseFromString('<config></config>', 'application/xml');
  const keyed = keys ? { privacy: 'True', privacyKey: 'PRIV A', rasKey: 'RAS A' } : {};
  const contacts = [
    contact(doc, CPS_BLUEPRINTS.contacts.Digital, 'DIG 1', '9001'),
    contact(doc, CPS_BLUEPRINTS.contacts.Digital, 'DIG 2', '9002'),
    contact(doc, CPS_BLUEPRINTS.contacts['Capacity Plus'], 'CP 1', '9101'),
    contact(doc, CPS_BLUEPRINTS.allCallContacts.Digital, 'DIG ALL CALL')
  ];
  const zoneChannels = [
    channel(doc, 'DGTLCONV6PT25', 'DIG RXONLY', { rx: '461.0125', tx: '461.0125', colorCode: '1', contact: 'DIG 1', scanList: 'ScanItems/SCAN A', receiveOnly: 'True' }),
    channel(doc, 'DGTLCONV6PT25', 'DIG TALK', { rx: '461.0250', tx: '466.0250', colorCode: '2', contact: 'DIG 2', scanList: roamListChannel ? 'RoamList/ROAM A' : 'NONE', receiveOnly: 'False', ...keyed }),
    channel(doc, 'ANLGCONV', 'ANA RXONLY', { rx: '462.5500', tx: '462.5500', scanList: 'ScanItems/SCAN B', receiveOnly: 'True' }),
    channel(doc, 'ANLGCONV', 'ANA TALK', { rx: '462.5750', tx: '467.5750', receiveOnly: 'False' }),
    channel(doc, 'CAPPLUSSYSCONV', 'CP RXONLY', { contact: 'CP 1', voiceList: 'VOICE A', receiveOnly: 'True' }),
    channel(doc, 'CAPPLUSSYSCONV', 'CP TALK', { contact: 'CP 1', voiceList: 'VOICE A', receiveOnly: 'False', ...keyed }),
    channel(doc, 'LINKEDCAPPLUSSYSCONV', 'LCP RXONLY', { contact: 'CP 1', siteList: 'SITES A', receiveOnly: 'True' }),
    channel(doc, 'LINKEDCAPPLUSSYSCONV', 'LCP TALK', { contact: 'CP 1', siteList: 'SITES A', receiveOnly: 'False', ...keyed })
  ];
  const poolChannels = pool ? [
    channel(doc, 'CAPPLUSVOICECONV', 'VOICE 1', { rx: '463.1000', tx: '468.1000', colorCode: '3' }),
    channel(doc, 'CAPPLUSVOICECONV', 'VOICE 2', { rx: '463.2000', tx: '468.2000', colorCode: '3' }),
    channel(doc, 'DGTLCONV6PT25', 'POOL ALL CALL', { rx: '463.3000', tx: '468.3000', colorCode: '4', contact: 'DIG ALL CALL', scanList: 'ScanItems/SCAN A', receiveOnly: 'True', ...(poolKeys ? { privacyKey: 'PRIV A', rasKey: 'RAS A' } : {}) })
  ] : [];

  const voiceList = node(doc, CPS_BLUEPRINTS.voiceList);
  voiceList.setAttribute('alias', 'VOICE A');
  setField(voiceList, 'TP_TRUNKPERALIAS', 'VOICE A');
  const members = children(voiceList, 'multifield')[0];
  for (const name of pool ? ['VOICE 1', 'VOICE 2'] : []) {
    const value = doc.createElement('value'); value.textContent = name; members.append(value);
  }
  const siteList = node(doc, CPS_BLUEPRINTS.siteList);
  siteList.setAttribute('alias', 'SITES A');
  setField(siteList, 'LCPSTLST_ALIAS', 'SITES A');
  const site = node(doc, CPS_BLUEPRINTS.siteItem);
  setField(site, 'LCPSTLST_SITEALIAS', 'SITE 1');
  setField(site, 'LCPSTLST_VCLSTIT', 'VOICE A');
  children(siteList, 'collection').find(item => item.getAttribute('name') === 'CapacityPlusSiteListItems').append(site);

  // Scan lists are kept exactly as stored; the workbook only names them.
  const scanList = (name) => node(doc, `<set name="ScanItems" alias="${name}" key="CONVSCAN"><multifield name="ScanListItems"><value>Selected</value></multifield><field name="SP_SCANLISTALIAS">${name}</field><field name="SP_SCLSTTYPE" Name="Scan List">CONVSCAN</field></set>`);
  const roamList = node(doc, '<set name="RoamList" alias="ROAM A"><field name="SS_ROAMLISTALIAS">ROAM A</field></set>');

  const config = doc.documentElement;
  config.append(
    categoryOf(doc, 'PCRContacts', contacts),
    categoryOf(doc, 'Zone', [zone(doc, CPS_BLUEPRINTS.zone, 'ZONE A', zoneChannels), zone(doc, CPS_BLUEPRINTS.channelPool, '', poolChannels)]),
    categoryOf(doc, 'CapacityPlusVoiceList', [voiceList]),
    categoryOf(doc, 'CapacityPlusSiteList', [siteList]),
    categoryOf(doc, 'ScanItems', [scanList('SCAN A'), scanList('SCAN B')]),
    categoryOf(doc, 'RoamList', [roamList])
  );
  return '<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' + config.toString();
}
