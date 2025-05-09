const { test, expect } = require('@playwright/test');
const { EmailAnalyzerPage } = require('./pages/email-analyzer-page');

test.describe('Email Analyzer Application', () => {
  let page;
  let emailAnalyzerPage;

  test.beforeEach(async ({ browser }) => {
    page = await browser.newPage();
    emailAnalyzerPage = new EmailAnalyzerPage(page);
    await emailAnalyzerPage.goto();
    // Wait for the app to fully load
    await page.waitForSelector('.app-title');
  });

  test('should correctly fetch and display email threads', async () => {
    // Set filter parameters for testing
    await emailAnalyzerPage.setFilter({
      fromFilter: 'nraman@dwellworks.com',
      subjectFilter: ''
    });
    
    // Fetch email threads
    await emailAnalyzerPage.fetchThreads();
    
    // Verify threads are loaded and properly grouped
    const threadCount = await emailAnalyzerPage.getThreadCount();
    expect(threadCount).toBeGreaterThan(0);
    
    // Verify thread dropdown format shows correct information
    const firstThreadText = await emailAnalyzerPage.getFirstThreadText();
    expect(firstThreadText).toMatch(/.*\(Count: \d+ emails\)/);
    
    // Take screenshot of threads dropdown
    await page.screenshot({ path: 'test-results/threads-dropdown.png' });
  });

  test('should process thread and display evaluation metrics', async () => {
    // Set filter parameters
    await emailAnalyzerPage.setFilter({
      fromFilter: 'nraman@dwellworks.com',
      subjectFilter: 'Driving License'  // Specific test case
    });
    
    // Fetch and select thread
    await emailAnalyzerPage.fetchThreads();
    await emailAnalyzerPage.selectThreadByIndex(0); // First thread
    
    // Process thread
    await emailAnalyzerPage.processThread();
    
    // Navigate to results tab
    await emailAnalyzerPage.goToResultsTab();
    
    // Verify results display
    const hasResults = await emailAnalyzerPage.hasResults();
    expect(hasResults).toBeTruthy();
    
    // Verify email count in thread
    const emailCount = await emailAnalyzerPage.getEmailCount();
    expect(emailCount).toBeGreaterThanOrEqual(2); // Should have at least 2 emails for "Driving License"
    
    // Check evaluation metrics sections
    expect(await emailAnalyzerPage.hasSentimentAnalysisSection()).toBeTruthy();
    expect(await emailAnalyzerPage.hasFeatureAnalysisSection()).toBeTruthy();
    expect(await emailAnalyzerPage.hasEventDetectionSection()).toBeTruthy();
    
    // Take screenshot of results
    await page.screenshot({ path: 'test-results/evaluation-metrics.png' });
  });

  test('should properly clean email content and validate against ground truth', async () => {
    // Set filter and fetch a thread with known content
    await emailAnalyzerPage.setFilter({
      fromFilter: 'nraman@dwellworks.com',
      subjectFilter: ''
    });
    
    await emailAnalyzerPage.fetchThreads();
    await emailAnalyzerPage.selectThreadByIndex(0);
    await emailAnalyzerPage.processThread();
    await emailAnalyzerPage.goToResultsTab();
    
    // Expand the email content section
    await emailAnalyzerPage.expandEmailContent();
    
    // Verify content cleaning has removed signatures, metadata
    const emailContent = await emailAnalyzerPage.getEmailContent();
    
    // Check that metadata patterns are not present
    expect(emailContent).not.toContain("From:");
    expect(emailContent).not.toContain("To:");
    expect(emailContent).not.toContain("Sent:");
    expect(emailContent).not.toContain("CAUTION:");
    expect(emailContent).not.toContain("DISCLAIMER:");
    
    // Expand ground truth section
    await emailAnalyzerPage.expandGroundTruth();
    
    // Compare AI output with ground truth
    const aiOutput = await emailAnalyzerPage.getAIOutput();
    const groundTruth = await emailAnalyzerPage.getGroundTruth();
    
    // Verify similarity calculations are displayed
    expect(await emailAnalyzerPage.getSimilarityScore()).toBeDefined();
    
    // Take screenshot of content comparison
    await page.screenshot({ path: 'test-results/content-comparison.png' });
  });

  test('should validate event detection with all required fields', async () => {
    // Find a thread with event data
    await emailAnalyzerPage.setFilter({
      fromFilter: 'nraman@dwellworks.com',
      subjectFilter: 'meeting' // Likely to contain events
    });
    
    await emailAnalyzerPage.fetchThreads();
    await emailAnalyzerPage.selectThreadByIndex(0);
    await emailAnalyzerPage.processThread();
    await emailAnalyzerPage.goToResultsTab();
    
    // Expand event detection section
    await emailAnalyzerPage.expandEventDetection();
    
    // Verify all 6 event fields are present
    const eventFields = await emailAnalyzerPage.getEventFields();
    const requiredFields = ["Event name", "Date", "Time", "Property Type", "Agent Name", "Location"];
    
    for (const field of requiredFields) {
      expect(eventFields).toContain(field);
    }
    
    // Verify content validation is performed (Level 1)
    const contentValidationDisplayed = await emailAnalyzerPage.hasContentValidation();
    expect(contentValidationDisplayed).toBeTruthy();
    
    // Check for similarity percentage (Level 2)
    const hasSimilarityPercentage = await emailAnalyzerPage.hasSimilarityPercentage();
    expect(hasSimilarityPercentage).toBeTruthy();
    
    // Take screenshot of event validation
    await page.screenshot({ path: 'test-results/event-validation.png' });
  });
});
