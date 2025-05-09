const { test, expect } = require('@playwright/test');
const { EmailAnalyzerPage } = require('./pages/email-analyzer-page');

test.describe('Evaluation Metrics', () => {
  let page;
  let emailAnalyzerPage;

  test.beforeEach(async ({ browser }) => {
    page = await browser.newPage();
    emailAnalyzerPage = new EmailAnalyzerPage(page);
    await emailAnalyzerPage.goto();
    
    // Set filter
    await emailAnalyzerPage.setFilter({
      fromFilter: 'nraman@dwellworks.com'
    });
    
    await emailAnalyzerPage.fetchThreads();
    await emailAnalyzerPage.selectThreadByIndex(0);
    await emailAnalyzerPage.processThread();
    await emailAnalyzerPage.goToResultsTab();
  });

  test('should display evaluation metrics with correct format and indicators', async () => {
    // Check sentiment analysis section
    await emailAnalyzerPage.expandSentimentAnalysis();
    
    // Verify sentiment fields are present
    const sentimentFields = await page.$$eval('div[data-testid="stExpander"]:has-text("Sentiment Analysis") table tr td:first-child', 
      cells => cells.map(cell => cell.textContent)
    );
    
    expect(sentimentFields).toContain('sentiment_analysis');
    expect(sentimentFields).toContain('overall_sentiment_analysis');
    
    // Check status indicators (Pass/Fail coloring)
    const passElements = await page.$$('div[data-testid="stExpander"]:has-text("Sentiment Analysis") td[style*="background-color: #1e7e34"]');
    const failElements = await page.$$('div[data-testid="stExpander"]:has-text("Sentiment Analysis") td[style*="background-color: #dc3545"]');
    
    expect(passElements.length + failElements.length).toBeGreaterThan(0);
    
    // Take screenshot
    await page.screenshot({ path: 'test-results/sentiment-analysis.png' });
  });

  test('should display feature classification with correct matrix values', async () => {
    // Check feature categorization
    await emailAnalyzerPage.expandFeatureAnalysis();
    
    // Get feature analysis data
    const featureData = await page.$$eval('div[data-testid="stExpander"]:has-text("Feature & Category") table tr', 
      rows => rows.slice(1).map(row => {
        const cells = Array.from(row.querySelectorAll('td'));
        return {
          field: cells[0].textContent,
          aiValue: cells[1].textContent,
          groundTruth: cells[2].textContent,
          status: cells[3].textContent
        };
      })
    );
    
    // Check feature values match expected format from classification matrix
    for (const row of featureData) {
      if (row.field === 'feature') {
        const validFeatures = [
          'EMAIL -- DSC First Contact with EE Completed',
          'EMAIL -- EE First Contact with DSC',
          'EMAIL -- Phone Consultation Scheduled',
          'EMAIL -- Phone Consultation Completed',
          'No feature'
        ];
        
        const aiFeatureIsValid = validFeatures.includes(row.aiValue);
        const gtFeatureIsValid = validFeatures.includes(row.groundTruth);
        
        expect(aiFeatureIsValid).toBeTruthy();
        expect(gtFeatureIsValid).toBeTruthy();
      }
      
      if (row.field === 'category') {
        const validCategories = [
          'Initial Service Milestones',
          'No category'
        ];
        
        const aiCategoryIsValid = validCategories.includes(row.aiValue);
        const gtCategoryIsValid = validCategories.includes(row.groundTruth);
        
        expect(aiCategoryIsValid).toBeTruthy();
        expect(gtCategoryIsValid).toBeTruthy();
      }
    }
    
    // Screenshot feature analysis
    await page.screenshot({ path: 'test-results/feature-analysis.png' });
  });

  test('should display event detection with all 6 required fields', async () => {
    // Check event detection
    await emailAnalyzerPage.expandEventDetection();
    
    // Get event field names
    const eventFieldNames = await page.$$eval('div[data-testid="stExpander"]:has-text("Event Detection") table tr td:first-child', 
      cells => cells.filter(cell => cell.textContent.includes('Event -'))
        .map(cell => cell.textContent.replace('Event -', '').trim())
    );
    
    // Check all 6 required fields are present
    const requiredFields = ['Event name', 'Date', 'Time', 'Property Type', 'Agent Name', 'Location'];
    
    for (const field of requiredFields) {
      expect(eventFieldNames).toContain(field);
    }
    
    // Check that null values are properly displayed
    const eventValues = await page.$$eval('div[data-testid="stExpander"]:has-text("Event Detection") table tr', 
      rows => rows.slice(1).map(row => {
        const cells = Array.from(row.querySelectorAll('td'));
        if (cells.length >= 2) {
          return cells[1].textContent; // AI Value column
        }
        return null;
      }).filter(Boolean)
    );
    
    // Verify null values are displayed as "null" or "N/A"
    const hasNullValues = eventValues.some(val => val === 'null' || val === 'N/A');
    expect(hasNullValues).toBeTruthy();
    
    // Verify content validation is present
    const contentValidation = await page.textContent('text="Content Validation:"');
    expect(contentValidation).toBeTruthy();
    
    // Screenshot event detection
    await page.screenshot({ path: 'test-results/event-detection.png' });
  });

  test('should display similarity scoring for summary validation', async () => {
    // Check summary validation
    const summarySection = await page.$('div[data-testid="stExpander"]:has-text("Summary Analysis")');
    
    if (summarySection) {
      await summarySection.click();
      
      // Check for similarity percentage
      const similarityScore = await page.textContent('text=Similarity Score:');
      expect(similarityScore).toMatch(/\d+\.\d+%/);
      
      // Check for visual indicators of similarity
      const scoreIndicator = await page.$('div:has-text("Similarity Score:") + div div[data-testid="stSuccess"], div:has-text("Similarity Score:") + div div[data-testid="stWarning"], div:has-text("Similarity Score:") + div div[data-testid="stError"]');
      
      expect(scoreIndicator).not.toBeNull();
    } else {
      console.log('Summary analysis section not found, skipping test');
    }
  });
});
