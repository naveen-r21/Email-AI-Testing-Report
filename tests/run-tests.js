const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

// Create test results directory if it doesn't exist
const resultsDir = path.join(__dirname, 'test-results');
if (!fs.existsSync(resultsDir)) {
  fs.mkdirSync(resultsDir, { recursive: true });
}

// Run Playwright tests
console.log('Running Playwright tests...');
const playwright = spawn('npx', ['playwright', 'test'], { stdio: 'inherit' });

playwright.on('close', (code) => {
  console.log(`Playwright tests completed with exit code ${code}`);
  
  if (code === 0) {
    console.log('All tests passed!');
  } else {
    console.error('Some tests failed. Check test-results directory for details.');
  }
  
  // Generate HTML report
  console.log('Generating HTML report...');
  const report = spawn('npx', ['playwright', 'show-report'], { stdio: 'inherit' });
  
  report.on('close', () => {
    console.log('Report generation completed.');
  });
});
