// Loads the app once, as it starts on the desktop, and prints what the user
// sees after the quiet startup update check. The fast suite runs it in a
// fresh process for each startup scenario (see UNARCHIVE_TEST_* in
// scripts/test-harness.mjs).
import { document, tauri, waitFor } from './test-harness.mjs';

await waitFor(() => tauri.calls.some(call => call.command === 'plugin:updater|check'), 'the startup update check');
await new Promise(resolve => setTimeout(resolve, 50));
const element = id => document.getElementById(id);
process.stdout.write(JSON.stringify({
  checks: tauri.calls.filter(call => call.command === 'plugin:updater|check').length,
  menus: tauri.calls.filter(call => call.command === 'plugin:menu|new').length,
  indicatorHidden: element('update-available').hidden,
  indicatorTitle: element('update-available').title,
  status: element('update-status').textContent,
  fallbackButtonHidden: element('update-check').hidden,
  dialogsOpen: ['update-dialog', 'about-dialog'].filter(id => element(id).hasAttribute('open')),
  prompts: tauri.calls.filter(call => call.command === 'plugin:dialog|message').length
}));
