const { chromium } = require('@playwright/test');

// Setup Playwright test environment
async function setupPlaywright() {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext();
  const page = await context.newPage();
  
  return { browser, context, page };
}

// Teardown Playwright test environment
async function teardownPlaywright(browser) {
  await browser.close();
}

module.exports = {
  setupPlaywright,
  teardownPlaywright
};
