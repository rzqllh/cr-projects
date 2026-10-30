import {copyFileSync} from 'node:fs';
// Build a safe site for CI/deployment with NO internal business records.
copyFileSync(new URL('../dist/data.template.js', import.meta.url), new URL('../dist/data.js', import.meta.url));
console.log('Safe empty snapshot generated; SharePoint data will require authorized Microsoft Graph sign-in.');
