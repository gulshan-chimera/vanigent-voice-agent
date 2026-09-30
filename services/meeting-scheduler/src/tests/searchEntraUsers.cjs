// Read-only local verification. Never registered as an Azure Function.
const fs = require('node:fs');
const path = require('node:path');

async function main() {
  const name = process.argv.slice(2).join(' ').trim();
  if (!name) throw new Error('Usage: npm run search:entra -- "Sneha"');
  const settingsPath = path.join(__dirname, '..', '..', 'local.settings.json');
  if (fs.existsSync(settingsPath)) {
    let values;
    try { values = JSON.parse(fs.readFileSync(settingsPath, 'utf8')).Values; }
    catch { throw new Error('Could not read local.settings.json. Check its JSON format.'); }
    for (const key of ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET']) {
      if (!process.env[key] && typeof values?.[key] === 'string') process.env[key] = values[key];
    }
  }
  if (['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET'].some(key => !process.env[key])) {
    throw new Error('Set TENANT_ID, CLIENT_ID and CLIENT_SECRET in the environment or local.settings.json Values.');
  }
  const { searchEntraUsers } = require('../../dist/meeting-scheduler/src/lib/entraUserSearch.js');
  console.log(JSON.stringify(await searchEntraUsers(name), null, 2));
}

main().catch(error => {
  // Search errors contain only sanitized messages/status, never raw Graph bodies.
  console.error(error.message);
  if (error.status) console.error(`HTTP status: ${error.status}`);
  if (error.retryAfterSeconds !== undefined) console.error(`Retry after ${error.retryAfterSeconds} seconds.`);
  process.exitCode = 1;
});
