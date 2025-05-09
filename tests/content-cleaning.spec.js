const { test, expect } = require('@playwright/test');
const { EmailAnalyzerPage } = require('./pages/email-analyzer-page');

test.describe('Email Content Cleaning', () => {
  let page;
  let emailAnalyzerPage;

  test.beforeEach(async ({ browser }) => {
    page = await browser.newPage();
    emailAnalyzerPage = new EmailAnalyzerPage(page);
    await emailAnalyzerPage.goto();
  });

  test('should properly remove email metadata while preserving property details', async () => {
    // Set filter to find emails with property details
    await emailAnalyzerPage.setFilter({
      subjectFilter: 'property'
    });
    
    await emailAnalyzerPage.fetchThreads();
    await emailAnalyzerPage.selectThreadByIndex(0);
    await emailAnalyzerPage.processThread();
    await emailAnalyzerPage.goToResultsTab();
    
    // Get the first email's content
    await emailAnalyzerPage.expandEmailContent();
    const cleanedContent = await emailAnalyzerPage.getEmailContent();
    
    // Check metadata is removed
    const metadataPatterns = [
      /From:.+?</i,
      /To:.+?</i,
      /Sent:.+?</i,
      /Subject:.+?</i,
      /CAUTION:/i,
      /DISCLAIMER:/i,
      /Confidential/i,
      /Original Message/i,
      /Forwarded Message/i
    ];
    
    for (const pattern of metadataPatterns) {
      expect(cleanedContent).not.toMatch(pattern);
    }
    
    // Check property details are preserved
    const propertyPatterns = [
      /\d+\s+.+?(Road|Street|Avenue|Boulevard|Drive|Lane)/i,
      /(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}/i,
      /\d{1,2}:\d{2}\s*(AM|PM)/i
    ];
    
    // At least one property pattern should be found if this is truly a property email
    const hasPropertyDetails = propertyPatterns.some(pattern => pattern.test(cleanedContent));
    expect(hasPropertyDetails).toBeTruthy();
  });

  test('should correctly handle emails with HTML content', async () => {
    // Set filter to find emails likely to have HTML content
    await emailAnalyzerPage.setFilter({
      fromFilter: '',
      subjectFilter: '' // Empty to get a variety of emails
    });
    
    await emailAnalyzerPage.fetchThreads();
    await emailAnalyzerPage.selectThreadByIndex(0);
    await emailAnalyzerPage.processThread();
    await emailAnalyzerPage.goToResultsTab();
    
    // Get the content
    await emailAnalyzerPage.expandEmailContent();
    const content = await emailAnalyzerPage.getEmailContent();
    
    // Verify HTML tags are stripped
    expect(content).not.toMatch(/<html|<body|<div|<span|<table|<tr|<td/i);
    
    // Ensure actual text content is preserved
    expect(content.length).toBeGreaterThan(10);
    expect(content.trim()).toBeTruthy();
  });
});
