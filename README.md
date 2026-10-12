# unArchive

A local Windows desktop app that converts CPS clipboard XML into an editable Excel workbook, then rebuilds CPS XML from that workbook. It exposes Contacts, Zones, Zone Channels, and Channel Pool while preserving all other CPS configuration in hidden reconstruction data.

## Use the desktop app

1. In CPS, copy the full configuration.
2. In unArchive, select **Paste from Clipboard** and then **Export editable workbook**.
3. Read the **READ FIRST** tab, then edit the visible planning worksheets. Keep the hidden `_CPS_TREE` and `_CPS_META` worksheets.
4. Select **Import workbook**, then **Copy XML to clipboard** and paste it into CPS.

### Start a new layout

To plan a fresh layout instead of editing existing zones, paste any codeplug for the target radio model (a factory-default codeplug works), open **Start a new layout**, and choose the number of zones, then an optional name, channel type, and channel count for each zone. Select **+ Add type** to mix channel types in one zone (for example 12 Digital and 2 Analog); each added type gets its own count and can be removed with **×**, and a zone can hold up to 160 channels in total. Unnamed zones become `Zone 1`, `Zone 2`…; zone names must be unique, 16 characters or fewer, and cannot be `Channel Pool`. Channels are numbered `CH 01`, `CH 02`… through the whole zone, ready to rename in the workbook. Channel defaults (Privacy Enabled, Privacy Key Alias, RAS Key Alias, Admit Criteria, In Call Criteria, Receive Only) are applied to every generated channel whose type supports them; Analog channels skip the privacy, RAS, and In Call settings and use Always in place of Color Code Free. Key aliases entered here are added to the Security sheet and must already exist in the target codeplug. The exported workbook keeps the pasted codeplug's radio settings, contacts, Channel Pool, voice lists, and site lists, and replaces its zones with the generated rows. Each generated channel is a new record, so import stops until it has the fields its type requires (frequencies and color code for conventional channels, a contact, and a voice or site list for Capacity Plus). Recheck scan lists, roam lists, and buttons that referred to the old zone channels in CPS.

The **Contacts** section is optional. Leave its type on **None** to keep the pasted contacts, or choose Capacity Plus, Linked Capacity Plus, or Digital with a count (select **+ Add contact type** for more rows) to replace them with numbered contacts: `CP 01…`, `LCP 01…`, and `DIG 01…` with sequential Call IDs, ready to rename. Each row has an optional **Starting Call ID** (for example 1001 for Capacity Plus and 1 for Digital); left on Auto, a block continues from the previous block of the same call type, starting at 1. Linked Capacity Plus channels use Capacity Plus contacts, so both share the Capacity Plus Call IDs: blocks may not overlap, and a Capacity Plus block may not include 255, the All Call ID. Choosing **All Call** instead of Group Call adds the system's single All Call contact (`ALL CALL`, Capacity Plus ID 255, or `DIG ALL CALL`, Digital ID 16777215). CPS scaffold contacts that combine several call types, such as `Call1`, are always kept. Analog channels have no talkgroup contacts, so analog is not offered.

Tick **Clear the Channel Pool** to drop the pasted Channel Pool channels and their frequencies; voice list members that pointed at them are emptied too, while the voice lists themselves are kept.

Your codeplug is processed locally. The packaged app bundles its Excel dependencies, so it works without internet access after installation.

### Updates

The top of the app shows the installed version. When it starts, unArchive quietly checks the latest GitHub release; if a newer version exists, **Install update** downloads it, verifies its signature, closes the app, and runs the installer. **Check for updates** checks on demand. Without internet access the check simply does nothing. Versions before 0.2.0 have no updater, so install 0.2.0 once by hand.

## Import safety rules

The editable workbook is authoritative for the fields it shows. During import, unArchive verifies that every visible contact, zone, and channel value made it into the rebuilt XML. A failed check clears the generated XML and leaves the copy/download buttons disabled, so an older successful result cannot be pasted by mistake.

- Adding a contact creates a complete CPS contact record from an exact, same-type CPS-generated blueprint. A new row with a system's reserved All Call ID (Capacity Plus 255, Digital 16777215) creates that system's All Call contact; a second All Call for the same system is rejected. Changing an existing contact's call type also rebuilds the complete record; it does not retain the old type's hidden fields.
- Built-in CPS-generated blueprints support Analog, Digital, Capacity Plus, Linked Capacity Plus, and Capacity Plus Voice channels even when the source contains no examples or the user deletes all planning records. Existing records retain their original hidden structure, while new records and type conversions use the controlled internal blueprints.
- Capacity Plus channels must reference Capacity Plus contacts, and conventional digital channels must reference Digital contacts. Missing and duplicate names, IDs, zones, lists, members, sites, and other cross-references are rejected.
- Zone, channel, and contact names are limited to 16 characters. CPS cuts longer names when pasting and moves a cut channel to the end of its zone, so import lists every longer name and refuses to build XML until it is shortened. The workbook also flags longer names as they are typed (pasted cells bypass Excel's check, so import remains the safeguard).
- CPS-required contacts that contain more than one call definition are preserved. Their name may be changed, but their combined type and ID cannot be reinterpreted through the simplified worksheet.
- Add new records using blank rows. Copying a whole existing row also copies its hidden record identity; duplicate or invalid identities are rejected rather than silently overwriting another record.

RAS/privacy key material is not present in CPS clipboard XML. The workbook shows referenced aliases for planning and validation, but the corresponding keys must already exist in the target codeplug before pasting. Capacity Plus voice-list membership should also be checked in CPS before deployment because CPS clipboard behavior can vary by version.

## Create a shareable Windows installer

The application is packaged with Tauri as a small Windows NSIS setup executable. Building it requires Node.js LTS, Rust, and Microsoft C++ Build Tools with the **Desktop development with C++** workload.

```powershell
npm install
npm run tauri -- dev
npm run package:windows
```

The shareable installer will be created under `src-tauri\target\release\bundle\nsis\`.

## Publish a release

The **Windows installer** GitHub Actions workflow builds and signs the installer on a Windows runner, then publishes a release with the setup executable, its signature, and the `latest.json` update manifest that installed apps check. The version comes from the tag, so there is no need to edit version numbers before releasing, but each tag must be higher than the last release or installed apps will not offer it. Run the import safety tests locally first:

```powershell
npm test
git tag v0.2.1
git push origin v0.2.1
```

The latest installer is always available from the repository's **Releases** page. Running the workflow manually from the **Actions** tab builds a signed installer artifact for testing without publishing a release. To set the version in the source files locally, run `npm run version:set -- 0.2.1`.

### Update signing key

Installed apps only accept updates signed with the private key that matches the public key in `src-tauri/tauri.conf.json`. The workflow reads it from the repository secret `TAURI_SIGNING_PRIVATE_KEY` (the full contents of the key file; the key has no password, so `TAURI_SIGNING_PRIVATE_KEY_PASSWORD` stays unset). Keep a backup of the key file somewhere safe: if it is lost, installed apps can no longer be updated in place and everyone must reinstall by hand once with a new key. Never commit it. Local builds with `npm run package:windows` do not need the key; they skip the updater signature.

## CPS blueprints and local tests

New contacts, channels, zones, and lists are built from the CPS-generated records in `blueprints.generated.js`, with every name, call ID, and frequency replaced by a placeholder. Codeplugs and workbooks are never committed. The blueprint generator (`npm run generate:blueprints`) and the import safety tests (`npm test`) read two local CPS fixtures, `cps-clipboard-source.xml` and `cps-clipboard-source-linked-plus-analog.xml`, which stay outside the repository along with `scripts/test-import-safety.mjs`.

Windows 10 (version 1803+) and Windows 11 normally include the WebView2 runtime Tauri uses. The installer is configured to download it if it is missing, keeping the installer lightweight. For air-gapped computers, switch the `webviewInstallMode` in `src-tauri/tauri.conf.json` to Tauri’s offline installer mode; this substantially increases the installer size.

## Important caution

This is an XML structure editor, not a radio-model validator. Start from a CPS template compatible with the target radio. Keep a known-good codeplug and complete a CPS validation plus a representative radio read/write test before releasing a configuration for fleet programming.
