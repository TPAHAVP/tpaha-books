import { defineConfig, devices } from '@playwright/test';

export default defineConfig({
  testDir: 'tests/e2e',
  timeout: 30_000,
  retries: 0,
  reporter: [['list']],
  use: { baseURL: 'http://localhost:8787/', trace: 'retain-on-failure' },
  webServer: { command: 'node tools/serve.mjs 8787', url: 'http://localhost:8787/vendor/xlsx.full.min.js', reuseExistingServer: true, timeout: 20_000 },
  projects: [
    { name: 'desktop', use: { ...devices['Desktop Chrome'] } },
    { name: 'iphone', use: { ...devices['iPhone 13'], browserName: 'chromium' } },
    { name: 'iphone-webkit', use: { ...devices['iPhone 13'] } },   // Safari engine (WebKit), phone size
  ],
});
