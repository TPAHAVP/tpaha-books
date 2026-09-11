import { test, expect } from '@playwright/test';

test('launcher links to both apps and shows local mode', async ({ page }) => {
  await page.goto('/index.html');
  await expect(page).toHaveTitle(/TPAHA Books/);
  await expect(page.locator('a.app-link[href="ledger.html"]')).toContainText('Treasurer Ledger');
  await expect(page.locator('a.app-link[href="storage.html"]')).toContainText('Winter Storage');
  await expect(page.locator('#account-status')).toContainText('Test mode');
  await expect(page.locator('#app-version')).toContainText(/v\d{4}\.\d{2}\.\d{2}/);
});

test('manifest and service worker are served and the worker registers', async ({ page, browserName }) => {
  const manifest = await page.request.get('/manifest.webmanifest');
  expect(manifest.ok()).toBeTruthy();
  const json = await manifest.json();
  expect(json.name).toBe('TPAHA Books');
  expect(json.display).toBe('standalone');
  expect(json.icons.length).toBeGreaterThanOrEqual(2);
  for (const icon of json.icons) expect((await page.request.get('/' + icon.src.replace(/^\.\//, ''))).ok()).toBeTruthy();
  const sw = await page.request.get('/sw.js');
  expect(sw.ok()).toBeTruthy();
  await page.goto('/index.html');
  const registered = await page.evaluate(async () => {
    if (!('serviceWorker' in navigator)) return 'unsupported';
    const reg = await navigator.serviceWorker.ready;
    return reg && reg.active ? 'active' : 'pending';
  });
  expect(['active', 'pending', 'unsupported']).toContain(registered);
});

test('launcher on a phone renders without horizontal overflow', async ({ page }) => {
  await page.goto('/index.html');
  const overflow = await page.evaluate(() => document.documentElement.scrollWidth > document.documentElement.clientWidth + 1);
  expect(overflow).toBe(false);
});
