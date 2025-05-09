module.exports = {
  testDir: './tests',
  timeout: 60000,
  expect: {
    timeout: 5000
  },
  reporter: [
    ['html'],
    ['json', { outputFile: 'test-results/test-results.json' }]
  ],
  use: {
    headless: false,
    viewport: { width: 1280, height: 720 },
    ignoreHTTPSErrors: true,
    video: 'on-first-retry',
    screenshot: 'only-on-failure',
    trace: 'on-first-retry',
  },
  projects: [
    {
      name: 'Chrome',
      use: { browserName: 'chromium' },
    },
    {
      name: 'Firefox',
      use: { browserName: 'firefox' },
    },
  ],
};
