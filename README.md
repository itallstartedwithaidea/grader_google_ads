# Google Ads Account Grader

![Version](https://img.shields.io/badge/version-2.0-blue)
![Platform](https://img.shields.io/badge/platform-Google%20Ads-red)
![License](https://img.shields.io/badge/license-MIT-green)

## Overview

The Google Ads Account Grader is a comprehensive script that evaluates your Google Ads account performance across 10 key categories of PPC best practices. This tool provides actionable insights and recommendations to improve your account's effectiveness and efficiency.

## Features

- **Comprehensive Analysis**: Evaluates your account across 10 critical areas of PPC management
- **Scoring System**: Each category is scored on a 0-100% scale with letter grades (A-F)
- **Prioritized Recommendations**: Provides actionable suggestions ranked by potential impact
- **Email Reports**: Automatically sends detailed reports to designated recipients
- **Customizable Parameters**: Adapt thresholds and benchmarks to your specific industry

## Evaluation Categories

1. **Campaign Organization**: Assesses your account structure, naming conventions, and segmentation
2. **Conversion Tracking**: Evaluates conversion setup, implementation quality, and tracking depth
3. **Keyword Strategy**: Analyzes keyword selection, match types, and organization
4. **Negative Keywords**: Reviews negative keyword implementation and effectiveness
5. **Bidding Strategy**: Evaluates bidding approaches and alignment with business goals
6. **Ad Creative & Extensions**: Assesses ad quality, testing methodology, and extension usage
7. **Quality Score**: Analyzes quality score metrics and component performance
8. **Audience Strategy**: Reviews audience targeting, remarketing, and audience segmentation
9. **Landing Page Optimization**: Evaluates landing page experience, performance, and testing
10. **Competitive Analysis**: Assesses competitive positioning and strategy

## Installation

1. Log in to your Google Ads account
2. Navigate to Tools > Scripts
3. Click the + (plus) button to create a new script
4. Copy the entire script code and paste it into the editor
5. Edit the CONFIG section to customize settings for your account
6. Authorize the required permissions when prompted
7. Save and run the script

## Configuration

The script includes a comprehensive configuration object at the beginning of the code:

```javascript
const CONFIG = {
  // Date range for analysis (in days)
  dateRange: {
    lookbackDays: 30
  },
  
  // Email settings
  email: {
    sendEmail: true,
    sendReport: true,
    sendErrorNotifications: true,
    emailAddress: 'your-email@example.com',
    errorRecipients: ['your-email@example.com'],
    includeSpreadsheetLink: true
  },
  
  // Spreadsheet settings
  spreadsheet: {
    createNew: true,
    existingSpreadsheetUrl: '', // Only used if createNew is false
    includeRawData: false // Whether to include raw data sheets
  },
  
  // Thresholds for letter grades
  gradeThresholds: {
    A: 90, // 90-100%
    B: 80, // 80-89%
    C: 70, // 70-79%
    D: 60, // 60-69%
    F: 0   // 0-59%
  },
  
  // Industry benchmarks (customize for your industry)
  industryBenchmarks: {
    ctr: 3.17,
    conversionRate: 3.75,
    cpc: 2.69,
    qualityScore: 6
  },
  
  // Best practice thresholds
  bestPractices: {
    keywordsPerAdGroup: 20,
    adsPerAdGroup: 3,
    minExtensionTypes: 4,
    minQualityScore: 7,
    maxCampaignsPerNegativeList: 20
  },
  
  // Category weights (must sum to 100)
  categoryWeights: {
    campaignOrganization: 10,
    conversionTracking: 15,
    keywordStrategy: 12,
    negativeKeywords: 8,
    biddingStrategy: 12,
    adCreative: 10,
    qualityScore: 10,
    audienceStrategy: 8,
    landingPage: 8,
    competitiveAnalysis: 7
  }
}
```

## Usage

### Running the Script

1. After installation and configuration, click "Run" to execute the script
2. For regular analysis, set up a schedule (e.g., weekly or monthly)
3. Review the generated report in the linked Google Sheets document

### Understanding Results

The script generates a comprehensive report with:

- An overall account grade (A-F)
- Scores for each of the 10 evaluation categories
- Specific metrics and findings within each category
- Prioritized recommendations for improvement
- Detailed data sheets for deeper analysis

## Outputs

### Email Report

If enabled, the script sends an email with:
- Overall account grade
- Summary of category grades
- Top recommendations
- Link to the detailed report spreadsheet

### Spreadsheet Report

The detailed spreadsheet includes:
- Summary dashboard with overall performance
- Individual sheets for each evaluation category
- Prioritized recommendations
- Raw data (optional)

## Best Practices

- Run the account grader monthly to track improvements
- Focus on addressing high-impact recommendations first
- Use the category scores to identify your weakest areas
- Compare results over time to measure progress
- Customize industry benchmarks to your specific vertical

## Requirements

- Google Ads account with administrative access
- Google Ads Scripts permissions
- Google Sheets access for reporting

## Technical Notes

- The script requires approximately 10-15 minutes to run on large accounts
- Memory usage is optimized for accounts with up to 100,000 keywords
- API quotas are respected with built-in throttling

## Support

For issues, questions, or feature requests, please contact:


## License

This project is licensed under the MIT License - see the LICENSE file for details.

---

*Copyright © 2025 Your Company Name. All rights reserved.*
