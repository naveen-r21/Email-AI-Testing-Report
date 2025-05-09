const { test, expect } = require('@playwright/test');
const { EmailAnalyzerPage } = require('./pages/email-analyzer-page');

test.describe('Email Thread Grouping', () => {
  let page;
  let emailAnalyzerPage;

  test.beforeEach(async ({ browser }) => {
    page = await browser.newPage();
    emailAnalyzerPage = new EmailAnalyzerPage(page);
    await emailAnalyzerPage.goto();
  });

  test('should group emails by conversationId not subject', async () => {
    // Set filter
    await emailAnalyzerPage.setFilter({
      fromFilter: '',
      subjectFilter: 'RE:' // Find threads with reply prefixes
    });
    
    await emailAnalyzerPage.fetchThreads();
    
    // Select first thread
    await emailAnalyzerPage.selectThreadByIndex(0);
    
    // Get thread information
    const threadInfo = await page.$$eval('text="Thread Information"', elements => {
      const infoSection = elements[0].closest('div');
      return {
        subject: infoSection.querySelector('text="Original Subject:"').nextSibling.textContent,
        cleanSubject: infoSection.querySelector('text="Clean Subject:"').nextSibling.textContent,
        threadId: infoSection.querySelector('text="Thread ID:"').nextSibling.textContent
      };
    });
    
    // Verify clean subject has RE: removed
    expect(threadInfo.subject).toMatch(/RE:|re:/i);
    expect(threadInfo.cleanSubject).not.toMatch(/RE:|re:/i);
    
    // Expand emails in thread
    await page.click('text="🔍 View Emails in Thread"');
    
    // Get the emails in thread
    const emailSubjects = await page.$$eval('tr td:nth-child(2)', cells => 
      cells.map(cell => cell.textContent)
    );
    
    // Verify all emails in thread have the same clean subject (ignoring RE:, FW:, etc.)
    const expectedCleanSubject = threadInfo.cleanSubject.trim();
    
    for (const subject of emailSubjects) {
      const cleanedSubject = subject.replace(/^(RE:|FW:|FWD:)\s*/i, '').trim();
      expect(cleanedSubject.toLowerCase()).toEqual(expectedCleanSubject.toLowerCase());
    }
    
    // Process thread
    await emailAnalyzerPage.processThread();
    await emailAnalyzerPage.goToResultsTab();
    
    // Check thread overview
    const threadOverview = await page.textContent('h2:has-text("Thread Overview") + p');
    expect(threadOverview).toContain(expectedCleanSubject);
  });

  test('should display all emails for "Driving License" thread', async () => {
    // Set filter to find the "Driving License" thread
    await emailAnalyzerPage.setFilter({
      fromFilter: '',
      subjectFilter: 'Driving License'
    });
    
    await emailAnalyzerPage.fetchThreads();
    await emailAnalyzerPage.selectThreadByIndex(0);
    
    // Expand emails in thread view
    await page.click('text="🔍 View Emails in Thread"');
    
    // Get number of emails in the thread from the information section
    const emailCount = await page.textContent('text="Number of Emails:"');
    const count = parseInt(emailCount.replace('Number of Emails:', '').trim());
    
    // Verify number of emails in thread
    expect(count).toBeGreaterThanOrEqual(2); // Should have at least 2 emails
    
    // Count emails in the expanded view
    const emailRows = await page.$$('tr');
    // Subtract 1 for header row
    expect(emailRows.length - 1).toEqual(count);
    
    // Process thread
    await emailAnalyzerPage.processThread();
    await emailAnalyzerPage.goToResultsTab();
    
    // Verify all tabs for emails are present
    const emailTabs = await page.$$(emailAnalyzerPage.emailTabsSelector);
    expect(emailTabs.length).toEqual(count);
  });
});
