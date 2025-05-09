// tests/pages/email-analyzer-page.js
class EmailAnalyzerPage {
    constructor(page) {
      this.page = page;
      this.url = 'http://localhost:8501'; // Update with actual URL
      
      // Selectors
      this.fromFilterInput = 'input[aria-label="From (Sender Email)"]';
      this.toFilterInput = 'input[aria-label="To (Recipient Email)"]';
      this.subjectFilterInput = 'input[aria-label="Subject Contains"]';
      this.fetchThreadsButton = 'button:has-text("Fetch Email Threads")';
      this.threadSelectDropdown = '[aria-label="Email Threads"]';
      this.processThreadButton = 'button:has-text("Process Thread")';
      this.resultTabSelector = '[data-baseweb="tab"]:has-text("Results & Reports")';
      this.emailContentExpanderSelector = ':has-text("📧 Email Content")';
      this.groundTruthExpanderSelector = ':has-text("✓ Groundtruth")';
      this.aiOutputExpanderSelector = ':has-text("🤖 AI Output")';
      this.eventDetectionExpanderSelector = ':has-text("📅 Event Detection")';
      this.sentimentAnalysisExpanderSelector = ':has-text("📊 Sentiment Analysis")';
      this.featureAnalysisExpanderSelector = ':has-text("🔍 Feature & Category Analysis")';
      this.emailTabsSelector = '[role="tablist"] [role="tab"]';
      this.emailContentSelector = '.email-content';
      this.similarityScoreSelector = ':has-text("Similarity Score:")';
      this.contentValidationSelector = ':has-text("Content Validation:")';
    }
  
    async goto() {
      await this.page.goto(this.url);
    }
  
    async setFilter({ fromFilter = '', toFilter = '', subjectFilter = '' }) {
      if (fromFilter) {
        await this.page.fill(this.fromFilterInput, fromFilter);
      }
      
      if (toFilter) {
        await this.page.fill(this.toFilterInput, toFilter);
      }
      
      if (subjectFilter) {
        await this.page.fill(this.subjectFilterInput, subjectFilter);
      }
    }
  
    async fetchThreads() {
      await this.page.click(this.fetchThreadsButton);
      // Wait for success message or threads to load
      await this.page.waitForSelector('text=Found', { timeout: 30000 });
    }
  
    async getThreadCount() {
      // Count the options in the dropdown
      return this.page.$$eval(`${this.threadSelectDropdown} option`, options => options.length);
    }
  
    async getFirstThreadText() {
      return this.page.$eval(`${this.threadSelectDropdown}`, dropdown => dropdown.textContent);
    }
  
    async selectThreadByIndex(index) {
      // Select thread from dropdown by index
      await this.page.selectOption(this.threadSelectDropdown, { index });
      // Wait for thread details to load
      await this.page.waitForSelector('text=Thread Information');
    }
  
    async processThread() {
      await this.page.click(this.processThreadButton);
      // Wait for processing to complete
      await this.page.waitForSelector('text=Successfully analyzed', { timeout: 60000 });
    }
  
    async goToResultsTab() {
      await this.page.click(this.resultTabSelector);
      // Wait for results tab to load
      await this.page.waitForSelector('text=Overall Analysis Metrics');
    }
  
    async hasResults() {
      const resultsSection = await this.page.$('text=Overall Analysis Metrics');
      return !!resultsSection;
    }
  
    async getEmailCount() {
      // Count the email tabs
      return this.page.$$eval(this.emailTabsSelector, tabs => tabs.length);
    }
  
    async expandEmailContent() {
      await this.page.click(this.emailContentExpanderSelector);
    }
  
    async expandGroundTruth() {
      await this.page.click(this.groundTruthExpanderSelector);
    }
  
    async expandAIOutput() {
      await this.page.click(this.aiOutputExpanderSelector);
    }
  
    async expandEventDetection() {
      await this.page.click(this.eventDetectionExpanderSelector);
    }
  
    async getEmailContent() {
      return this.page.$eval(this.emailContentSelector, content => content.innerText);
    }
  
    async getAIOutput() {
      // Extract AI output JSON
      await this.expandAIOutput();
      const aiOutputText = await this.page.$eval('.stJson', element => element.innerText);
      return JSON.parse(aiOutputText);
    }
  
    async getGroundTruth() {
      // Extract Ground Truth JSON
      await this.expandGroundTruth();
      const groundTruthText = await this.page.$eval('.stJson', element => element.innerText);
      return JSON.parse(groundTruthText);
    }
  
    async getSimilarityScore() {
      const scoreElement = await this.page.$(this.similarityScoreSelector);
      return scoreElement ? scoreElement.innerText : null;
    }
  
    async hasSentimentAnalysisSection() {
      const section = await this.page.$(this.sentimentAnalysisExpanderSelector);
      return !!section;
    }
  
    async hasFeatureAnalysisSection() {
      const section = await this.page.$(this.featureAnalysisExpanderSelector);
      return !!section;
    }
  
    async hasEventDetectionSection() {
      const section = await this.page.$(this.eventDetectionExpanderSelector);
      return !!section;
    }
  
    async getEventFields() {
      await this.expandEventDetection();
      // Get all field names from the event section
      return this.page.$$eval('text="Event - "', fields => 
        fields.map(field => field.textContent.replace('Event - ', '').trim())
      );
    }
  
    async hasContentValidation() {
      const validation = await this.page.$(this.contentValidationSelector);
      return !!validation;
    }
  
    async hasSimilarityPercentage() {
      const similarity = await this.page.$(this.similarityScoreSelector);
      return !!similarity;
    }
  }
  
  module.exports = { EmailAnalyzerPage };