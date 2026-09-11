// Checks site/js/config.js against the Entra app registration, without contacting Microsoft and without
// touching a workbook. Run it after any change to the configuration or to js/auth.js:
//
//   npm run verify:config
//
// What it cannot do: prove that the registration in Entra still matches these values, that consent is still
// granted, or that an assigned account can sign in. Those are Checkpoint 2 checks against the real service.
import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const configUrl = new URL('../site/js/config.js', import.meta.url).href;

// The registration as Cody completed it on 2026-09-11. Change these only alongside a real registration change.
const EXPECTED = {
  clientId: '6c5bf2f8-a50c-4908-ba93-535816364785',
  tenantId: 'ef5ad4e3-2b18-4542-8fd2-ed7212cbcc93',
  site: 'https://tpahavp.github.io/tpaha-books/',
  scopes: ['Files.ReadWrite', 'User.Read'],
};
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;

const results = [];
const check = (name, ok, detail = '') => results.push({ name, ok, detail });
/** Loads the config as a page served from `hostname` would see it. */
const asServedFrom = async hostname => {
  globalThis.location = { hostname, href: `https://${hostname}/tpaha-books/ledger.html` };
  const { CONFIG } = await import(`${configUrl}?h=${encodeURIComponent(hostname)}`);
  return CONFIG;
};

const live = await asServedFrom(new URL(EXPECTED.site).hostname);
check('Client id is a well-formed GUID and matches the registration', GUID.test(live.clientId) && live.clientId === EXPECTED.clientId, live.clientId);
check('Tenant id is a well-formed GUID and matches the directory', GUID.test(live.tenantId) && live.tenantId === EXPECTED.tenantId, live.tenantId);
check('The two ids are different values', live.clientId !== live.tenantId);

// The published site is configured; anything served from this machine is not, so tests and local work stay offline.
for (const host of ['localhost', '127.0.0.1']) {
  const c = await asServedFrom(host);
  check(`Served from ${host}: test mode, no Microsoft sign-in`, c.clientId === '' && c.tenantId === '', `clientId="${c.clientId}"`);
}
for (const host of [new URL(EXPECTED.site).hostname, 'books.tpaha.ca']) {
  const c = await asServedFrom(host);
  check(`Served from ${host}: configured for Microsoft sign-in`, c.clientId === EXPECTED.clientId);
}

// The redirect URI the app actually sends is derived in js/auth.js as new URL('./', location.href).
for (const page of ['index.html', 'ledger.html', 'diagnostics.html', 'storage.html']) {
  const derived = new URL('./', `${EXPECTED.site}${page}`).href;
  check(`Redirect URI derived from ${page} equals the registered SPA redirect`, derived === EXPECTED.site, derived);
}

const auth = readFileSync(join(root, 'site/js/auth.js'), 'utf8');
const scopes = /const SCOPES = \[([^\]]*)\]/.exec(auth)[1].replace(/['\s]/g, '').split(',').sort();
check('Requests exactly the consented delegated scopes', JSON.stringify(scopes) === JSON.stringify([...EXPECTED.scopes].sort()), scopes.join(', '));
check('Authority is the single tenant, not /common or /organizations', auth.includes('login.microsoftonline.com/${config.tenantId'));
check('Sign-in is redirect-based, so no implicit-grant or popup token flow is needed', auth.includes('loginRedirect') && !auth.includes('loginPopup'));

// Nothing secret may ever appear in the published site.
const secretish = /client[_-]?secret|certificate_thumbprint|BEGIN (RSA |EC )?PRIVATE KEY/i;
const leaks = [];
const walk = dir => {
  for (const f of readdirSync(join(root, dir), { withFileTypes: true })) {
    const rel = `${dir}/${f.name}`;
    if (f.isDirectory()) { if (f.name !== 'vendor') walk(rel); continue; }
    if (!/\.(js|html|json|css|webmanifest)$/.test(f.name)) continue;
    if (secretish.test(readFileSync(join(root, rel), 'utf8'))) leaks.push(rel);
  }
};
walk('site');
check('No client secret, certificate or private key anywhere in the published site', leaks.length === 0, leaks.join(', '));

check('No workbook is pinned: the member must identify and confirm one', live.workbooks.ledger.driveId === '' && live.workbooks.ledger.itemId === '');
check('Pilot writer model is single-writer', live.pilot.writerModel === 'single');

const failed = results.filter(r => !r.ok);
for (const r of results) console.log(`${r.ok ? 'ok  ' : 'FAIL'}  ${r.name}${r.detail ? `  [${r.detail}]` : ''}`);
console.log(`\n${results.length - failed.length} passed, ${failed.length} failed`);
process.exit(failed.length ? 1 : 0);
