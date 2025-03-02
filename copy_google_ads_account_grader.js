/**
 * Google Ads Account Grader
 * 
 * This script performs a comprehensive analysis of your Google Ads account,
 * evaluating performance across 10 key categories of PPC best practices:
 * - Campaign Organization
 * - Conversion Tracking
 * - Keyword Strategy
 * - Negative Keywords
 * - Bidding Strategy
 * - Ad Creative & Extensions
 * - Quality Score
 * - Audience Strategy
 * - Landing Page Optimization
 * - Competitive Analysis
 * 
 * Each category is scored on a 0-100% scale with detailed metrics and formulas,
 * and assigned a letter grade (A-F). The script provides actionable recommendations
 * prioritized by potential impact.
 * 
 * @version 2.0
 */

// Configuration
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
      emailAddress: 'googlescriptupdates-aaaappoowiqsm2y5ofawatmrri@outerbox.slack.com',
      errorRecipients: ['googlescriptupdates-aaaappoowiqsm2y5ofawatmrri@outerbox.slack.com'],
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
  };
  
  // Define evaluation categories
  const EVALUATION_CATEGORIES = [
    {
      name: "Campaign Organization",
      weight: CONFIG.categoryWeights.campaignOrganization,
      criteria: [
        { name: "Logical Campaign & Ad Group Structure", weight: 40 },
        { name: "Clear Naming Conventions & Segmentation", weight: 30 },
        { name: "No Internal Competition", weight: 30 }
      ]
    },
    {
      name: "Conversion Tracking",
      weight: CONFIG.categoryWeights.conversionTracking,
      criteria: [
        { name: "Comprehensive Conversion Coverage", weight: 40 },
        { name: "Accurate and Verified Tracking Implementation", weight: 35 },
        { name: "Enhanced & Offline Conversion Tracking", weight: 25 }
      ]
    },
    {
      name: "Keyword Strategy",
      weight: CONFIG.categoryWeights.keywordStrategy,
      criteria: [
        { name: "Extensive Keyword Research & Relevance", weight: 30 },
        { name: "Strategic Match Type Use", weight: 25 },
        { name: "Brand vs Non-Brand Segmentation", weight: 25 },
        { name: "Continuous Keyword Optimization", weight: 20 }
      ]
    },
    {
      name: "Negative Keywords",
      weight: CONFIG.categoryWeights.negativeKeywords,
      criteria: [
        { name: "Routine Search Query Mining", weight: 40 },
        { name: "Negative Keyword Lists and Hierarchy", weight: 35 },
        { name: "Balanced Exclusion (Avoid False Negatives)", weight: 25 }
      ]
    },
    {
      name: "Bidding Strategy",
      weight: CONFIG.categoryWeights.biddingStrategy,
      criteria: [
        { name: "Goal-Aligned Bidding Approach", weight: 35 },
        { name: "Optimize Automated Bidding with Data", weight: 25 },
        { name: "Device, Location, and Time Bid Adjustments", weight: 20 },
        { name: "Budget Management & Bid Strategy Alignment", weight: 20 }
      ]
    },
    {
      name: "Ad Creative & Extensions",
      weight: CONFIG.categoryWeights.adCreative,
      criteria: [
        { name: "Compelling Ad Copy with Relevance", weight: 30 },
        { name: "Ad Variety and Continuous Testing", weight: 25 },
        { name: "Leverage Ad Extensions", weight: 30 },
        { name: "Ad Quality and Compliance", weight: 15 }
      ]
    },
    {
      name: "Quality Score",
      weight: CONFIG.categoryWeights.qualityScore,
      criteria: [
        { name: "Monitor Quality Score & Components", weight: 25 },
        { name: "Improve Ad Relevance", weight: 25 },
        { name: "Improve Expected CTR", weight: 25 },
        { name: "Improve Landing Page Experience", weight: 25 }
      ]
    },
    {
      name: "Audience Strategy",
      weight: CONFIG.categoryWeights.audienceStrategy,
      criteria: [
        { name: "Remarketing & Retargeting", weight: 35 },
        { name: "Customer Match & Similar Audiences", weight: 25 },
        { name: "In-Market, Affinity, and Demographic Targeting", weight: 25 },
        { name: "Personalized Ad Experiences by Audience", weight: 15 }
      ]
    },
    {
      name: "Landing Page Optimization",
      weight: CONFIG.categoryWeights.landingPage,
      criteria: [
        { name: "Relevance and Message Match", weight: 30 },
        { name: "Conversion-Focused Design", weight: 30 },
        { name: "Page Speed and Mobile Optimization", weight: 25 },
        { name: "A/B Testing & Iteration", weight: 15 }
      ]
    },
    {
      name: "Competitive Analysis",
      weight: CONFIG.categoryWeights.competitiveAnalysis,
      criteria: [
        { name: "Auction Insights Monitoring", weight: 35 },
        { name: "Competitor Keyword and Ad Analysis", weight: 25 },
        { name: "Benchmarking Performance Metrics", weight: 25 },
        { name: "Adaptive Strategy to Competitor Moves", weight: 15 }
      ]
    }
  ];
  
  /**
   * Main function that runs the account grader
   */
  function main() {
    Logger.log("Starting Google Ads Account Grader...");
    
    try {
      // Collect account data
      Logger.log("Collecting account data...");
      const accountData = collectAccountData();
      
      // Evaluate each category
      const evaluationResults = {
        campaignorganization: evaluateCampaignOrganization(accountData),
        conversiontracking: evaluateConversionTracking(accountData),
        keywordstrategy: evaluateKeywordStrategy(accountData),
        negativekeywords: evaluateNegativeKeywords(accountData),
        biddingstrategy: evaluateBiddingStrategy(accountData),
        adcreativeextensions: evaluateAdCreative(accountData),
        qualityscore: evaluateQualityScore(accountData),
        audiencestrategy: evaluateAudienceStrategy(accountData),
        landingpageoptimization: evaluateLandingPage(accountData),
        competitiveanalysis: evaluateCompetitiveAnalysis(accountData)
      };
      
      // Calculate overall grade
      const overallGrade = calculateOverallGrade(evaluationResults);
      
      // Generate prioritized recommendations
      const prioritizedRecommendations = generatePrioritizedRecommendations(evaluationResults);
      
      // Create report
      const reportSpreadsheet = createReport(evaluationResults, overallGrade, prioritizedRecommendations, accountData);
      
      // Send email notification if configured
      if (CONFIG.email.sendReport) {
        sendEmailReport(reportSpreadsheet.getUrl(), evaluationResults, overallGrade, accountData);
      }
      
      Logger.log("Account grading completed successfully!");
      Logger.log("Overall grade: " + overallGrade.letter + " (" + overallGrade.score.toFixed(1) + ")");
      Logger.log("Report spreadsheet: " + reportSpreadsheet.getUrl());
      
    } catch (error) {
      Logger.log("Error running account grader: " + error.message);
      Logger.log(error.stack);
      
      // Send error notification if configured
      if (CONFIG.email.sendErrorNotifications) {
        sendErrorNotification(error);
      }
    }
  }
  
  /**
   * Gets the date range for analysis
   * @return {Object} Date range object with start and end dates
   */
  function getDateRange() {
    Logger.log("Getting date range...");
    
    // Get date range from CONFIG or use default (last 30 days)
    const today = new Date();
    const lookbackDays = CONFIG.dateRange.lookbackDays || 30;
    
    // Calculate start date by subtracting lookback days
    const startDate = new Date(today);
    startDate.setDate(today.getDate() - lookbackDays);
    
    // Format dates in YYYYMMDD format required by Google Ads
    const formattedStartDate = Utilities.formatDate(startDate, AdsApp.currentAccount().getTimeZone(), 'yyyyMMdd');
    const formattedEndDate = Utilities.formatDate(today, AdsApp.currentAccount().getTimeZone(), 'yyyyMMdd');
    
    Logger.log("Date range: " + formattedStartDate + " to " + formattedEndDate);
    
    return {
      start: formattedStartDate,
      end: formattedEndDate
    };
  }
  
  /**
   * Calculates the overall account grade based on category scores
   * Generates prioritized recommendations based on evaluation results
   * @param {Object} evaluationResults The results of all category evaluations
   * @return {Array} Prioritized list of recommendations
   */
  function generatePrioritizedRecommendations(evaluationResults) {
    Logger.log("Generating prioritized recommendations...");
    
    // Collect all recommendations from all categories
    const allRecommendations = [];
    
    for (const category in evaluationResults) {
      const categoryName = EVALUATION_CATEGORIES.find(c => 
        c.name.toLowerCase().replace(/\s+/g, '') === category).name;
      
      evaluationResults[category].recommendations.forEach(rec => {
        allRecommendations.push({
          category: categoryName,
          text: rec.text,
          impact: rec.impact
        });
      });
    }
    
    // Sort recommendations by impact (highest first)
    allRecommendations.sort((a, b) => b.impact - a.impact);
    
    return allRecommendations;
  }
  
  /**
   * Creates a report spreadsheet with the evaluation results
   * @param {Object} evaluationResults The results of all category evaluations
   * @param {Object} overallGrade The overall account grade
   * @param {Array} prioritizedRecommendations The prioritized recommendations
   * @param {Object} accountData The collected account data
   * @return {Spreadsheet} The created spreadsheet
   */
  function createReport(evaluationResults, overallGrade, prioritizedRecommendations, accountData) {
    Logger.log("Creating report spreadsheet...");
    
    const accountName = AdsApp.currentAccount().getName();
    const accountId = AdsApp.currentAccount().getCustomerId();
    const date = Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), "yyyy-MM-dd");
    
    // Create a new spreadsheet
    const spreadsheet = SpreadsheetApp.create("Google Ads Account Grader - " + accountName + " - " + date);
    const summarySheet = spreadsheet.getActiveSheet().setName("Summary");
    
    // Add account info
    let row = 1;
    summarySheet.getRange(row, 1).setValue("Google Ads Account Grader Report");
    summarySheet.getRange(row, 1).setFontWeight("bold").setFontSize(16);
    row += 2;
    
    summarySheet.getRange(row, 1).setValue("Account:");
    summarySheet.getRange(row, 2).setValue(accountName + " (" + accountId + ")");
    row++;
    
    summarySheet.getRange(row, 1).setValue("Date:");
    summarySheet.getRange(row, 2).setValue(date);
    row += 2;
    
    // Add overall grade
    summarySheet.getRange(row, 1).setValue("Overall Account Grade:");
    summarySheet.getRange(row, 2).setValue(overallGrade.letter + " (" + overallGrade.score.toFixed(1) + ")");
    
    // Format the grade cell based on the grade
    const gradeCell = summarySheet.getRange(row, 2);
    if (overallGrade.letter === 'A') {
      gradeCell.setBackground("#b7e1cd"); // Green
    } else if (overallGrade.letter === 'B') {
      gradeCell.setBackground("#c9daf8"); // Light blue
    } else if (overallGrade.letter === 'C') {
      gradeCell.setBackground("#fce8b2"); // Yellow
    } else if (overallGrade.letter === 'D') {
      gradeCell.setBackground("#f7c8a2"); // Orange
    } else {
      gradeCell.setBackground("#f4c7c3"); // Red
    }
    
    row += 2;
    
    // Add category summary
    summarySheet.getRange(row, 1).setValue("Category Grades");
    summarySheet.getRange(row, 1).setFontWeight("bold").setFontSize(14);
    row += 1;
    
    // Create header row
    summarySheet.getRange(row, 1).setValue("Category");
    summarySheet.getRange(row, 2).setValue("Grade");
    summarySheet.getRange(row, 3).setValue("Score");
    summarySheet.getRange(row, 1, 1, 3).setFontWeight("bold").setBackground("#efefef");
    row++;
    
    // Add each category grade
    for (const category in evaluationResults) {
      const categoryName = EVALUATION_CATEGORIES.find(c => 
        c.name.toLowerCase().replace(/\s+/g, '') === category).name;
      const result = evaluationResults[category];
      
      summarySheet.getRange(row, 1).setValue(categoryName);
      summarySheet.getRange(row, 2).setValue(result.letter);
      summarySheet.getRange(row, 3).setValue(result.score.toFixed(1));
      
      // Format the grade cell based on the grade
      const gradeCellCategory = summarySheet.getRange(row, 2);
      if (result.letter === 'A') {
        gradeCellCategory.setBackground("#b7e1cd"); // Green
      } else if (result.letter === 'B') {
        gradeCellCategory.setBackground("#c9daf8"); // Light blue
      } else if (result.letter === 'C') {
        gradeCellCategory.setBackground("#fce8b2"); // Yellow
      } else if (result.letter === 'D') {
        gradeCellCategory.setBackground("#f7c8a2"); // Orange
      } else {
        gradeCellCategory.setBackground("#f4c7c3"); // Red
      }
      
      row++;
    }
    
    row += 2;
    
    // Add top recommendations
    summarySheet.getRange(row, 1).setValue("Top Recommendations");
    summarySheet.getRange(row, 1).setFontWeight("bold").setFontSize(14);
    row += 1;
    
    // Create header row
    summarySheet.getRange(row, 1).setValue("Category");
    summarySheet.getRange(row, 2).setValue("Recommendation");
    summarySheet.getRange(row, 3).setValue("Impact");
    summarySheet.getRange(row, 1, 1, 3).setFontWeight("bold").setBackground("#efefef");
    row++;
    
    // Add top 10 recommendations
    const topRecommendations = prioritizedRecommendations.slice(0, 10);
    topRecommendations.forEach(rec => {
      summarySheet.getRange(row, 1).setValue(rec.category);
      summarySheet.getRange(row, 2).setValue(rec.text);
      
      // Convert impact score to text
      let impactText = "";
      if (rec.impact >= 0.9) {
        impactText = "Critical";
      } else if (rec.impact >= 0.7) {
        impactText = "High";
      } else if (rec.impact >= 0.5) {
        impactText = "Medium";
      } else {
        impactText = "Low";
      }
      
      summarySheet.getRange(row, 3).setValue(impactText);
      
      // Format impact cell
      const impactCell = summarySheet.getRange(row, 3);
      if (rec.impact >= 0.9) {
        impactCell.setBackground("#f4c7c3"); // Red for critical
      } else if (rec.impact >= 0.7) {
        impactCell.setBackground("#f7c8a2"); // Orange for high
      } else if (rec.impact >= 0.5) {
        impactCell.setBackground("#fce8b2"); // Yellow for medium
      } else {
        impactCell.setBackground("#b7e1cd"); // Green for low
      }
      
      row++;
    });
    
    // Auto-resize columns
    summarySheet.autoResizeColumns(1, 3);
    
    // Create detailed sheets for each category
    for (const category in evaluationResults) {
      const categoryName = EVALUATION_CATEGORIES.find(c => 
        c.name.toLowerCase().replace(/\s+/g, '') === category).name;
      const result = evaluationResults[category];
      
      // Create a new sheet for this category
      const categorySheet = spreadsheet.insertSheet(categoryName);
      
      // Add category info
      row = 1;
      categorySheet.getRange(row, 1).setValue(categoryName + " Evaluation");
      categorySheet.getRange(row, 1).setFontWeight("bold").setFontSize(16);
      row += 2;
      
      categorySheet.getRange(row, 1).setValue("Grade:");
      categorySheet.getRange(row, 2).setValue(result.letter + " (" + result.score.toFixed(1) + ")");
      
      // Format the grade cell
      const gradeCellDetail = categorySheet.getRange(row, 2);
      if (result.letter === 'A') {
        gradeCellDetail.setBackground("#b7e1cd"); // Green
      } else if (result.letter === 'B') {
        gradeCellDetail.setBackground("#c9daf8"); // Light blue
      } else if (result.letter === 'C') {
        gradeCellDetail.setBackground("#fce8b2"); // Yellow
      } else if (result.letter === 'D') {
        gradeCellDetail.setBackground("#f7c8a2"); // Orange
      } else {
        gradeCellDetail.setBackground("#f4c7c3"); // Red
      }
      
      row += 2;
      
      // Add criteria details
      categorySheet.getRange(row, 1).setValue("Evaluation Criteria");
      categorySheet.getRange(row, 1).setFontWeight("bold").setFontSize(14);
      row += 1;
      
      // Create header row
      categorySheet.getRange(row, 1).setValue("Criterion");
      categorySheet.getRange(row, 2).setValue("Score");
      categorySheet.getRange(row, 1, 1, 2).setFontWeight("bold").setBackground("#efefef");
      row++;
      
      // Add each criterion
      for (const criterionKey in result.criteria) {
        const criterion = result.criteria[criterionKey];
        
        // Convert camelCase to readable text
        const criterionName = criterionKey
          .replace(/([A-Z])/g, ' $1')
          .replace(/^./, function(str) { return str.toUpperCase(); });
        
        categorySheet.getRange(row, 1).setValue(criterionName);
        categorySheet.getRange(row, 2).setValue(criterion.score.toFixed(1));
        
        row++;
      }
      
      row += 2;
      
      // Add recommendations
      categorySheet.getRange(row, 1).setValue("Recommendations");
      categorySheet.getRange(row, 1).setFontWeight("bold").setFontSize(14);
      row += 1;
      
      if (result.recommendations.length > 0) {
        // Create header row
        categorySheet.getRange(row, 1).setValue("Recommendation");
        categorySheet.getRange(row, 2).setValue("Impact");
        categorySheet.getRange(row, 1, 1, 2).setFontWeight("bold").setBackground("#efefef");
        row++;
        
        // Add each recommendation
        result.recommendations.forEach(rec => {
          categorySheet.getRange(row, 1).setValue(rec.text);
          
          // Convert impact score to text
          let impactText = "";
          if (rec.impact >= 0.9) {
            impactText = "Critical";
          } else if (rec.impact >= 0.7) {
            impactText = "High";
          } else if (rec.impact >= 0.5) {
            impactText = "Medium";
          } else {
            impactText = "Low";
          }
          
          categorySheet.getRange(row, 2).setValue(impactText);
          
          // Format impact cell
          const impactCellDetail = categorySheet.getRange(row, 2);
          if (rec.impact >= 0.9) {
            impactCellDetail.setBackground("#f4c7c3"); // Red for critical
          } else if (rec.impact >= 0.7) {
            impactCellDetail.setBackground("#f7c8a2"); // Orange for high
          } else if (rec.impact >= 0.5) {
            impactCellDetail.setBackground("#fce8b2"); // Yellow for medium
          } else {
            impactCellDetail.setBackground("#b7e1cd"); // Green for low
          }
          
          row++;
        });
      } else {
        categorySheet.getRange(row, 1).setValue("No recommendations for this category.");
        row++;
      }
      
      row += 2;
      
      // Add details section if available
      if (Object.keys(result.criteria).length > 0) {
        categorySheet.getRange(row, 1).setValue("Detailed Metrics");
        categorySheet.getRange(row, 1).setFontWeight("bold").setFontSize(14);
        row += 1;
        
        // For each criterion, add its details
        for (const criterionKey in result.criteria) {
          const criterion = result.criteria[criterionKey];
          
          if (criterion.details && Object.keys(criterion.details).length > 0) {
            // Convert camelCase to readable text
            const criterionName = criterionKey
              .replace(/([A-Z])/g, ' $1')
              .replace(/^./, function(str) { return str.toUpperCase(); });
            
            categorySheet.getRange(row, 1).setValue(criterionName + " Details");
            categorySheet.getRange(row, 1).setFontWeight("bold");
            row++;
            
            // Add each detail
            for (const detailKey in criterion.details) {
              // Convert camelCase to readable text
              const detailName = detailKey
                .replace(/([A-Z])/g, ' $1')
                .replace(/^./, function(str) { return str.toUpperCase(); });
              
              let detailValue = criterion.details[detailKey];
              
              // Format percentages
              if (detailName.toLowerCase().includes('percentage') || 
                  detailKey.toLowerCase().includes('rate') ||
                  detailKey.toLowerCase().includes('share')) {
                if (typeof detailValue === 'number' && detailValue <= 1) {
                  detailValue = (detailValue * 100).toFixed(2) + '%';
                } else if (typeof detailValue === 'number') {
                  detailValue = detailValue.toFixed(2) + '%';
                }
              }
              
              categorySheet.getRange(row, 1).setValue(detailName);
              categorySheet.getRange(row, 2).setValue(detailValue);
              row++;
            }
            
            row++; // Add a blank row between criteria
          }
        }
      }
      
      // Auto-resize columns
      categorySheet.autoResizeColumns(1, 2);
    }
    
    // Create a data sheet with raw account data
    const dataSheet = spreadsheet.insertSheet("Account Data");
    
    // Add account data
    row = 1;
    dataSheet.getRange(row, 1).setValue("Account Data");
    dataSheet.getRange(row, 1).setFontWeight("bold").setFontSize(16);
    row += 2;
    
    // Helper function to add a data section
    function addDataSection(dataObj, prefix = '') {
      for (const key in dataObj) {
        if (typeof dataObj[key] === 'object' && dataObj[key] !== null && !Array.isArray(dataObj[key])) {
          // This is a nested object, add a header for it
          const sectionName = key
            .replace(/([A-Z])/g, ' $1')
            .replace(/^./, function(str) { return str.toUpperCase(); });
          
          dataSheet.getRange(row, 1).setValue(prefix + sectionName);
          dataSheet.getRange(row, 1).setFontWeight("bold");
          row++;
          
          // Recursively add its contents
          addDataSection(dataObj[key], '  '); // Add indentation for nested items
        } else {
          // This is a value, add it directly
          const valueName = key
            .replace(/([A-Z])/g, ' $1')
            .replace(/^./, function(str) { return str.toUpperCase(); });
          
          let value = dataObj[key];
          
          // Format percentages
          if (valueName.toLowerCase().includes('percentage') || 
              key.toLowerCase().includes('rate') ||
              key.toLowerCase().includes('share')) {
            if (typeof value === 'number' && value <= 1) {
              value = (value * 100).toFixed(2) + '%';
            } else if (typeof value === 'number') {
              value = value.toFixed(2) + '%';
            }
          }
          
          dataSheet.getRange(row, 1).setValue(prefix + valueName);
          dataSheet.getRange(row, 2).setValue(value);
          row++;
        }
      }
    }
    
    // Add each main section of account data
    for (const section in accountData) {
      if (section === 'campaigns') continue; // Skip the campaigns array as it's too large
      
      const sectionName = section
        .replace(/([A-Z])/g, ' $1')
        .replace(/^./, function(str) { return str.toUpperCase(); });
      
      dataSheet.getRange(row, 1).setValue(sectionName);
      dataSheet.getRange(row, 1).setFontWeight("bold").setFontSize(14);
      row++;
      
      addDataSection(accountData[section]);
      
      row++; // Add a blank row between sections
    }
    
    // Auto-resize columns
    dataSheet.autoResizeColumns(1, 2);
    
    // Set the Summary sheet as active
    spreadsheet.setActiveSheet(summarySheet);
    
    return spreadsheet;
  }
  
  /**
   * Sends an email report with the evaluation results
   * @param {string} spreadsheetUrl URL of the report spreadsheet
   * @param {Object} evaluationResults The results of all category evaluations
   * @param {Object} overallGrade The overall account grade
   * @param {Object} accountData The collected account data
   */
  function sendEmailReport(spreadsheetUrl, evaluationResults, overallGrade, accountData) {
    Logger.log("Sending email report...");
    
    const accountName = AdsApp.currentAccount().getName();
    const accountId = AdsApp.currentAccount().getCustomerId();
    
    // Create email subject
    const subject = "Google Ads Account Grader Report - " + accountName + " (" + accountId + ")";
    
    // Create email body
    let body = "<h2>Google Ads Account Grader Report</h2>";
    body += "<p><strong>Account:</strong> " + accountName + " (" + accountId + ")</p>";
    body += "<p><strong>Date:</strong> " + Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), "yyyy-MM-dd") + "</p>";
    
    // Add overall grade
    body += "<h3>Overall Account Grade: " + overallGrade.letter + " (" + overallGrade.score.toFixed(1) + ")</h3>";
    
    // Add category summary
    body += "<h3>Category Grades:</h3>";
    body += "<table border='1' cellpadding='5' style='border-collapse: collapse;'>";
    body += "<tr><th>Category</th><th>Grade</th></tr>";
    
    for (const category in evaluationResults) {
      const categoryName = EVALUATION_CATEGORIES.find(c => c.name.toLowerCase().replace(/\s+/g, '') === category).name;
      const result = evaluationResults[category];
      
      body += "<tr>";
      body += "<td>" + categoryName + "</td>";
      body += "<td>" + result.letter + " (" + result.score.toFixed(1) + ")</td>";
      body += "</tr>";
    }
    
    body += "</table>";
    
    // Add top recommendations
    body += "<h3>Top Recommendations:</h3>";
    body += "<ul>";
    
    // Get top 5 recommendations across all categories
    const allRecommendations = [];
    for (const category in evaluationResults) {
      evaluationResults[category].recommendations.forEach(rec => {
        allRecommendations.push({
          category: EVALUATION_CATEGORIES.find(c => c.name.toLowerCase().replace(/\s+/g, '') === category).name,
          text: rec.text,
          impact: rec.impact
        });
      });
    }
    
    // Sort by impact and take top 5
    allRecommendations.sort((a, b) => b.impact - a.impact);
    const topRecommendations = allRecommendations.slice(0, 5);
    
    topRecommendations.forEach(rec => {
      body += "<li><strong>" + rec.category + ":</strong> " + rec.text + "</li>";
    });
    
    body += "</ul>";
    
    // Add link to full report
    body += "<p><a href='" + spreadsheetUrl + "'>View Full Report</a></p>";
    
    // Send the email
    MailApp.sendEmail({
      to: CONFIG.email.emailAddress,
      subject: subject,
      htmlBody: body
    });
  }
  
  /**
   * Collects all account data needed for evaluation
   * @return {Object} Collected account data
   */
  function collectAccountData() {
    // Initialize account data object
    const accountData = {
      account: {
        name: AdsApp.currentAccount().getName(),
        id: AdsApp.currentAccount().getCustomerId(),
        currencyCode: AdsApp.currentAccount().getCurrencyCode(),
        timeZone: AdsApp.currentAccount().getTimeZone()
      }
    };
    
    // Get date range
    const dateRange = getDateRange();
    accountData.dateRange = dateRange;  // Store the date range in accountData
    
    // Collect performance data with the date range
    collectPerformanceData(accountData, dateRange);
    
    // Collect account structure data
    collectAccountStructure(accountData);
    
    // Collect keyword data
    collectKeywordData(accountData, dateRange);
    
    // Collect negative keyword data
    collectNegativeKeywordData(accountData);
    
    // Collect bidding data
    collectBiddingData(accountData, dateRange);
    
    // Collect ad data
    collectAdData(accountData, dateRange);
    
    // Collect extension data
    collectExtensionData(accountData, dateRange);
    
    // Collect quality score data
    collectQualityScoreData(accountData);
    
    // Collect conversion data
    collectConversionData(accountData, dateRange);
    
    // Collect audience data
    collectAudienceData(accountData);
    
    // Collect landing page data
    collectLandingPageData(accountData);
    
    // Collect competitive data
    collectCompetitiveData(accountData, dateRange);
    
    return accountData;
  }
  
  /**
   * Collects overall account performance data
   * @param {Object} accountData The account data object to populate
   */
  function collectPerformanceData(accountData, dateRange) {
    Logger.log("Collecting performance data...");
    
    const query = "SELECT Impressions, Clicks, Cost, Conversions " +
                 "FROM ACCOUNT_PERFORMANCE_REPORT " +
                 "WHERE Impressions > 0 " +
                 "DURING " + dateRange.start + "_" + dateRange.end;
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    let metrics = {
      impressions: 0,
      clicks: 0,
      cost: 0,
      conversions: 0
    };
    
    while (rows.hasNext()) {
      const row = rows.next();
      metrics.impressions += parseInt(row['Impressions']);
      metrics.clicks += parseInt(row['Clicks']);
      metrics.cost += parseFloat(row['Cost']);
      metrics.conversions += parseInt(row['Conversions']);
    }
    
    return metrics;
  }
  
  /**
   * Collects campaign data
   * @param {Object} accountData The account data object to populate
   */
  function collectCampaignData(accountData) {
    Logger.log("Collecting campaign data...");
    
    const dateRange = accountData.dateRange;
    
    // Create a query to get campaign data
    const query = `
      SELECT
        campaign.id,
        campaign.name,
        campaign.status,
        campaign.advertising_channel_type,
        campaign.bidding_strategy_type,
        campaign_budget.amount_micros,
        metrics.impressions,
        metrics.clicks,
        metrics.cost_micros,
        metrics.conversions,
        metrics.conversions_value,
        metrics.search_impression_share,
        metrics.search_rank_lost_impression_share,
        metrics.search_budget_lost_impression_share
      FROM campaign
      WHERE segments.date BETWEEN '${dateRange.start}' AND '${dateRange.end}'
    `;
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    const campaigns = [];
    let campaignCount = 0;
    let searchImpressionShare = 0;
    let searchRankLost = 0;
    let searchBudgetLost = 0;
    let campaignsWithImpressionShare = 0;
    
    // Also collect bidding strategy data
    const biddingStrategies = {
      manual: 0,
      enhanced: 0,
      targetCpa: 0,
      targetRoas: 0,
      maximizeConversions: 0,
      maximizeValue: 0,
      targetImpressionShare: 0
    };
    
    while (rows.hasNext()) {
      const row = rows.next();
      campaignCount++;
      
      const campaignId = row['campaign.id'];
      const campaignName = row['campaign.name'];
      const status = row['campaign.status'];
      const channelType = row['campaign.advertising_channel_type'];
      const biddingStrategyType = row['campaign.bidding_strategy_type'];
      const budgetMicros = parseInt(row['campaign_budget.amount_micros'], 10) || 0;
      const impressions = parseInt(row['metrics.impressions'], 10) || 0;
      const clicks = parseInt(row['metrics.clicks'], 10) || 0;
      const costMicros = parseInt(row['metrics.cost_micros'], 10) || 0;
      const conversions = parseFloat(row['metrics.conversions']) || 0;
      const conversionValue = parseFloat(row['metrics.conversions_value']) || 0;
      const impressionShare = parseFloat(row['metrics.search_impression_share']) || 0;
      const rankLost = parseFloat(row['metrics.search_rank_lost_impression_share']) || 0;
      const budgetLost = parseFloat(row['metrics.search_budget_lost_impression_share']) || 0;
      
      // Convert cost and budget from micros to actual currency
      const cost = costMicros / 1000000;
      const budget = budgetMicros / 1000000;
      
      // Calculate derived metrics
      const ctr = clicks > 0 ? (clicks / impressions) * 100 : 0;
      const averageCpc = clicks > 0 ? cost / clicks : 0;
      const conversionRate = clicks > 0 ? (conversions / clicks) * 100 : 0;
      const costPerConversion = conversions > 0 ? cost / conversions : 0;
      const roas = cost > 0 ? conversionValue / cost : 0;
      
      // Add to campaigns array
      campaigns.push({
        id: campaignId,
        name: campaignName,
        status: status,
        channelType: channelType,
        biddingStrategyType: biddingStrategyType,
        budget: budget,
        impressions: impressions,
        clicks: clicks,
        cost: cost,
        conversions: conversions,
        conversionValue: conversionValue,
        ctr: ctr,
        averageCpc: averageCpc,
        conversionRate: conversionRate,
        costPerConversion: costPerConversion,
        roas: roas,
        impressionShare: impressionShare,
        rankLost: rankLost,
        budgetLost: budgetLost
      });
      
      // Update bidding strategy counts
      if (biddingStrategyType) {
        if (biddingStrategyType.includes('MANUAL_CPC')) {
          biddingStrategies.manual++;
        } else if (biddingStrategyType.includes('ENHANCED_CPC')) {
          biddingStrategies.enhanced++;
        } else if (biddingStrategyType.includes('TARGET_CPA')) {
          biddingStrategies.targetCpa++;
        } else if (biddingStrategyType.includes('TARGET_ROAS')) {
          biddingStrategies.targetRoas++;
        } else if (biddingStrategyType.includes('MAXIMIZE_CONVERSIONS')) {
          biddingStrategies.maximizeConversions++;
        } else if (biddingStrategyType.includes('MAXIMIZE_CONVERSION_VALUE')) {
          biddingStrategies.maximizeValue++;
        } else if (biddingStrategyType.includes('TARGET_IMPRESSION_SHARE')) {
          biddingStrategies.targetImpressionShare++;
        }
      }
      
      // Update impression share metrics (only for search campaigns with data)
      if (channelType === 'SEARCH' && !isNaN(impressionShare) && impressionShare > 0) {
        searchImpressionShare += impressionShare;
        searchRankLost += rankLost;
        searchBudgetLost += budgetLost;
        campaignsWithImpressionShare++;
      }
    }
    
    // Calculate average impression share metrics
    if (campaignsWithImpressionShare > 0) {
      searchImpressionShare /= campaignsWithImpressionShare;
      searchRankLost /= campaignsWithImpressionShare;
      searchBudgetLost /= campaignsWithImpressionShare;
    }
    
    // Update account data
    accountData.structure.campaignCount = campaignCount;
    accountData.campaigns = campaigns;
    accountData.biddingStrategies = biddingStrategies;
    accountData.competitive.searchImpressionShare = searchImpressionShare;
    accountData.competitive.searchRankLost = searchRankLost;
    accountData.competitive.searchBudgetLost = searchBudgetLost;
    
    // Check for bid adjustments
    checkBidAdjustments(accountData);
  }
  
  /**
   * Checks for device, location, and schedule bid adjustments
   * @param {Object} accountData The account data object to populate
   */
  function checkBidAdjustments(accountData) {
    Logger.log("Checking bid adjustments...");
    
    let hasDeviceAdjustments = false;
    let hasLocationAdjustments = false;
    let hasScheduleAdjustments = false;
    
    // Check a sample of campaigns for bid adjustments
    const campaignIterator = AdsApp.campaigns()
      .withCondition("Status = ENABLED")
      .withLimit(50)
      .get();
    
    while (campaignIterator.hasNext()) {
      const campaign = campaignIterator.next();
      
      // Check device bid adjustments
      if (!hasDeviceAdjustments) {
        const deviceIterator = campaign.targeting().platforms().get();
        while (deviceIterator.hasNext()) {
          const device = deviceIterator.next();
          if (device.getBidModifier() !== 1.0) {
            hasDeviceAdjustments = true;
            break;
          }
        }
      }
      
      // Check location bid adjustments
      if (!hasLocationAdjustments) {
        const locationIterator = campaign.targeting().targetedLocations().get();
        while (locationIterator.hasNext()) {
          const location = locationIterator.next();
          if (location.getBidModifier() !== 1.0) {
            hasLocationAdjustments = true;
            break;
          }
        }
      }
      
      // Check ad schedule bid adjustments
      if (!hasScheduleAdjustments) {
        const scheduleIterator = campaign.targeting().adSchedules().get();
        while (scheduleIterator.hasNext()) {
          const schedule = scheduleIterator.next();
          if (schedule.getBidModifier() !== 1.0) {
            hasScheduleAdjustments = true;
            break;
          }
        }
      }
      
      // If all adjustments found, no need to check more campaigns
      if (hasDeviceAdjustments && hasLocationAdjustments && hasScheduleAdjustments) {
        break;
      }
    }
    
    // Update account data
    accountData.bidAdjustments = {
      device: hasDeviceAdjustments,
      location: hasLocationAdjustments,
      schedule: hasScheduleAdjustments
    };
  }
  
  /**
   * Collects ad group data
   * @param {Object} accountData The account data object to populate
   */
  function collectAdGroupData(accountData) {
    Logger.log("Collecting ad group data...");
    
    const dateRange = accountData.dateRange;
    
    // Create a query to get ad group data
    const query = `
      SELECT
        ad_group.id,
        ad_group.name,
        ad_group.status,
        campaign.id,
        campaign.name,
        metrics.impressions,
        metrics.clicks,
        metrics.cost_micros,
        metrics.conversions
      FROM ad_group
      WHERE segments.date BETWEEN '${dateRange.start}' AND '${dateRange.end}'
    `;
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    const adGroups = [];
    let adGroupCount = 0;
    
    while (rows.hasNext()) {
      const row = rows.next();
      adGroupCount++;
      
      const adGroupId = row['ad_group.id'];
      const adGroupName = row['ad_group.name'];
      const status = row['ad_group.status'];
      const campaignId = row['campaign.id'];
      const campaignName = row['campaign.name'];
      const impressions = parseInt(row['metrics.impressions'], 10) || 0;
      const clicks = parseInt(row['metrics.clicks'], 10) || 0;
      const costMicros = parseInt(row['metrics.cost_micros'], 10) || 0;
      const conversions = parseFloat(row['metrics.conversions']) || 0;
      
      // Convert cost from micros to actual currency
      const cost = costMicros / 1000000;
      
      // Calculate derived metrics
      const ctr = clicks > 0 ? (clicks / impressions) * 100 : 0;
      const conversionRate = clicks > 0 ? (conversions / clicks) * 100 : 0;
      
      // Add to ad groups array
      adGroups.push({
        id: adGroupId,
        name: adGroupName,
        status: status,
        campaignId: campaignId,
        campaignName: campaignName,
        impressions: impressions,
        clicks: clicks,
        cost: cost,
        conversions: conversions,
        ctr: ctr,
        conversionRate: conversionRate,
        keywordCount: 0,
        adCount: 0
      });
    }
    
    // Update account data
    accountData.structure.adGroupCount = adGroupCount;
    accountData.adGroups = adGroups;
  }
  
  /**
   * Collects keyword data
   * @param {Object} accountData The account data object to populate
   * @param {Object} dateRange Date range for data collection
   */
  function collectKeywordData(accountData, dateRange) {
    Logger.log("Collecting keyword data...");
    
    const query = `SELECT Id, Criteria, KeywordMatchType, QualityScore, SearchImpressionShare, ` +
                  `Impressions, Clicks, Cost, Conversions, AverageCpc, FirstPageCpc, TopOfPageCpc ` +
                  `FROM KEYWORDS_PERFORMANCE_REPORT ` +
                  `WHERE Status = "ENABLED" AND Date BETWEEN '${dateRange.start}' AND '${dateRange.end}'`;
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    // Initialize counters and distributions
    let matchTypeDistribution = {
      'EXACT': 0,
      'PHRASE': 0,
      'BROAD': 0
    };
    
    let keywordLengthDistribution = {
      'short': 0,  // 1 word
      'medium': 0, // 2-3 words
      'long': 0    // 4+ words
    };
    
    let lowQualityKeywords = 0;
    let nonConvertingKeywords = 0;
    let brandKeywords = 0;
    let totalKeywords = 0;
    
    // Get brand terms from account name (simplified approach)
    const accountName = accountData.account.name.toLowerCase();
    const brandTerms = accountName.split(/\s+/).filter(word => word.length > 3);
    
    // Process keyword data
    while (rows.hasNext()) {
      const row = rows.next();
      totalKeywords++;
      
      // Match type distribution
      const matchType = row['KeywordMatchType'];
      if (matchTypeDistribution[matchType] !== undefined) {
        matchTypeDistribution[matchType]++;
      }
      
      // Quality score analysis
      const qualityScore = parseInt(row['QualityScore'], 10) || 0;
      if (qualityScore < CONFIG.bestPractices.minQualityScore && qualityScore > 0) {
        lowQualityKeywords++;
      }
      
      // Conversion analysis
      const conversions = parseFloat(row['Conversions']) || 0;
      const clicks = parseInt(row['Clicks'], 10) || 0;
      if (clicks > 10 && conversions === 0) {
        nonConvertingKeywords++;
      }
      
      // Keyword length analysis
      const keyword = row['Criteria'].toLowerCase();
      const wordCount = keyword.split(/\s+/).length;
      
      if (wordCount === 1) {
        keywordLengthDistribution.short++;
      } else if (wordCount <= 3) {
        keywordLengthDistribution.medium++;
      } else {
        keywordLengthDistribution.long++;
      }
      
      // Brand keyword analysis
      if (brandTerms.some(term => keyword.includes(term))) {
        brandKeywords++;
      }
    }
    
    // Calculate percentages
    const lowQualityKeywordPercentage = totalKeywords > 0 ? 
      lowQualityKeywords / totalKeywords : 0;
    const nonConvertingKeywordPercentage = totalKeywords > 0 ? 
      nonConvertingKeywords / totalKeywords : 0;
    const brandKeywordPercentage = totalKeywords > 0 ? 
      brandKeywords / totalKeywords : 0;
    
    // Populate keyword data
    accountData.keywords.matchTypeDistribution = matchTypeDistribution;
    accountData.keywords.lowQualityKeywordPercentage = lowQualityKeywordPercentage;
    accountData.keywords.nonConvertingKeywordPercentage = nonConvertingKeywordPercentage;
    accountData.keywords.brandKeywordPercentage = brandKeywordPercentage;
    accountData.keywords.keywordLengthDistribution = keywordLengthDistribution;
  }
  
  /**
   * Collects ad data
   * @param {Object} accountData The account data object to populate
   */
  function collectAdData(accountData) {
    Logger.log("Collecting ad data...");
    
    const dateRange = accountData.dateRange;
    
    // Create a query to get ad data
    const query = `
      SELECT
        ad_group_ad.ad.id,
        ad_group_ad.ad.type,
        ad_group_ad.ad.final_urls,
        ad_group_ad.ad_strength,
        ad_group_ad.policy_summary.approval_status,
        ad_group.id,
        ad_group.name,
        campaign.id,
        campaign.name,
        metrics.impressions,
        metrics.clicks,
        metrics.cost_micros,
        metrics.conversions
      FROM ad_group_ad
      WHERE segments.date BETWEEN '${dateRange.start}' AND '${dateRange.end}'
    `;
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    const ads = [];
    let adCount = 0;
    let disapprovedAdCount = 0;
    let rsaCount = 0;
    let excellentOrGoodAdStrengthCount = 0;
    
    // Track ad counts by ad group
    const adGroupAdCounts = new Map();
    
    while (rows.hasNext()) {
      const row = rows.next();
      adCount++;
      
      const adId = row['ad_group_ad.ad.id'];
      const adType = row['ad_group_ad.ad.type'];
      const finalUrls = row['ad_group_ad.ad.final_urls'];
      const adStrength = row['ad_group_ad.ad_strength'];
      const approvalStatus = row['ad_group_ad.policy_summary.approval_status'];
      const adGroupId = row['ad_group.id'];
      const adGroupName = row['ad_group.name'];
      const campaignId = row['campaign.id'];
      const campaignName = row['campaign.name'];
      const impressions = parseInt(row['metrics.impressions'], 10) || 0;
      const clicks = parseInt(row['metrics.clicks'], 10) || 0;
      const costMicros = parseInt(row['metrics.cost_micros'], 10) || 0;
      const conversions = parseFloat(row['metrics.conversions']) || 0;
      
      // Convert cost from micros to actual currency
      const cost = costMicros / 1000000;
      
      // Calculate derived metrics
      const ctr = clicks > 0 ? (clicks / impressions) * 100 : 0;
      const conversionRate = clicks > 0 ? (conversions / clicks) * 100 : 0;
      
      // Track disapproved ads
      if (approvalStatus === 'DISAPPROVED') {
        disapprovedAdCount++;
      }
      
      // Track RSA count and ad strength
      if (adType === 'RESPONSIVE_SEARCH_AD') {
        rsaCount++;
        if (adStrength === 'EXCELLENT' || adStrength === 'GOOD') {
          excellentOrGoodAdStrengthCount++;
        }
      }
      
      // Track ad counts by ad group
      if (adGroupAdCounts.has(adGroupId)) {
        adGroupAdCounts.set(adGroupId, adGroupAdCounts.get(adGroupId) + 1);
      } else {
        adGroupAdCounts.set(adGroupId, 1);
      }
      
      // Add to ads array
      ads.push({
        id: adId,
        type: adType,
        finalUrls: finalUrls,
        adStrength: adStrength,
        approvalStatus: approvalStatus,
        adGroupId: adGroupId,
        adGroupName: adGroupName,
        campaignId: campaignId,
        campaignName: campaignName,
        impressions: impressions,
        clicks: clicks,
        cost: cost,
        conversions: conversions,
        ctr: ctr,
        conversionRate: conversionRate
      });
      
      // Update ad group ad count
      const adGroup = accountData.adGroups.find(ag => ag.id === adGroupId);
      if (adGroup) {
        adGroup.adCount++;
      }
    }
    
    // Calculate ad disapproval rate
    const adDisapprovalRate = adCount > 0 ? disapprovedAdCount / adCount : 0;
    
    // Calculate RSA adoption rate
    const rsaAdoptionRate = adCount > 0 ? rsaCount / adCount : 0;
    
    // Calculate RSA quality rate
    const rsaQualityRate = rsaCount > 0 ? excellentOrGoodAdStrengthCount / rsaCount : 0;
    
    // Calculate ad groups with multiple ads
    let adGroupsWithMultipleAds = 0;
    adGroupAdCounts.forEach(count => {
      if (count >= CONFIG.bestPractices.adsPerAdGroup) {
        adGroupsWithMultipleAds++;
      }
    });
    const adGroupsWithMultipleAdsRate = accountData.structure.adGroupCount > 0 ? 
      adGroupsWithMultipleAds / accountData.structure.adGroupCount : 0;
    
    // Update account data
    accountData.structure.adCount = adCount;
    accountData.structure.adDisapprovalRate = adDisapprovalRate;
    accountData.structure.rsaAdoptionRate = rsaAdoptionRate;
    accountData.structure.rsaQualityRate = rsaQualityRate;
    accountData.structure.adGroupsWithMultipleAdsRate = adGroupsWithMultipleAdsRate;
    accountData.ads = ads;
  }
  
  /**
   * Collects negative keyword data
   * @param {Object} accountData The account data object to populate
   */
  function collectNegativeKeywordData(accountData) {
    Logger.log("Collecting negative keyword data...");
    
    // Get campaign negative keywords
    const campaignNegativeIterator = AdsApp.negativeKeywords()
      .withCondition("Status = ENABLED")
      .get();
    
    // Get ad group negative keywords
    const adGroupNegativeIterator = AdsApp.adGroupNegativeKeywords()
      .withCondition("Status = ENABLED")
      .get();
    
    // Get shared negative keyword sets
    const sharedSetIterator = AdsApp.negativeKeywordLists()
      .get();
    
    // Count negative keywords
    const campaignNegativeCount = campaignNegativeIterator.totalNumEntities();
    const adGroupNegativeCount = adGroupNegativeIterator.totalNumEntities();
    const sharedSetCount = sharedSetIterator.totalNumEntities();
    
    // Count campaigns using shared sets
    let campaignsUsingSharedSets = 0;
    if (sharedSetCount > 0) {
      const campaignSharedSetQuery = "SELECT CampaignName " +
        "FROM CAMPAIGN_NEGATIVE_KEYWORDS_LIST_REPORT";
      
      const report = AdsApp.report(campaignSharedSetQuery);
      const rows = report.rows();
      
      // Count unique campaigns
      const campaignsWithSharedSets = new Set();
      while (rows.hasNext()) {
        const row = rows.next();
        campaignsWithSharedSets.add(row['CampaignName']);
      }
      
      campaignsUsingSharedSets = campaignsWithSharedSets.size;
    }
    
    // Estimate irrelevant search percentage (simplified approach)
    // In a real implementation, this would require search query report analysis
    const irrelevantSearchPercentage = 0.1; // Placeholder value
    
    // Populate negative keyword data
    accountData.negativeKeywords.campaignNegativeCount = campaignNegativeCount;
    accountData.negativeKeywords.adGroupNegativeCount = adGroupNegativeCount;
    accountData.negativeKeywords.sharedSetCount = sharedSetCount;
    accountData.negativeKeywords.campaignsUsingSharedSets = campaignsUsingSharedSets;
    accountData.negativeKeywords.irrelevantSearchPercentage = irrelevantSearchPercentage;
  }
  
  /**
   * Collects extension data
   * @param {Object} accountData The account data object to populate
   */
  function collectExtensionData(accountData) {
    Logger.log("Collecting extension data...");
    
    const dateRange = accountData.dateRange;
    
    // Create a query to get extension data
    const query = `
      SELECT
        ad_group_ad.ad_group,
        campaign.id,
        campaign.name,
        metrics.impressions,
        segments.ad_network_type,
        segments.extension_type
      FROM ad_group_ad
      WHERE segments.date BETWEEN '${dateRange.start}' AND '${dateRange.end}'
        AND segments.extension_type != ""
    `;
    
    try {
      const report = AdsApp.report(query);
      const rows = report.rows();
      
      const extensionTypes = new Set();
      let impressionsWithExtensions = 0;
      
      while (rows.hasNext()) {
        const row = rows.next();
        
        const extensionType = row['segments.extension_type'];
        const impressions = parseInt(row['metrics.impressions'], 10) || 0;
        
        extensionTypes.add(extensionType);
        impressionsWithExtensions += impressions;
      }
      
      // Update account data
      accountData.extensions = {
        types: Array.from(extensionTypes),
        totalCount: extensionTypes.size,
        impressionWithExtensions: impressionsWithExtensions
      };
    } catch (e) {
      Logger.log("Error collecting extension data: " + e);
      accountData.extensions = {
        types: [],
        totalCount: 0,
        impressionWithExtensions: 0
      };
    }
  }
  
  /**
   * Collects audience data
   * @param {Object} accountData The account data object to populate
   */
  function collectAudienceData(accountData) {
    Logger.log("Collecting audience data...");
    
    // Check for remarketing lists
    const remarketingListIterator = AdsApp.audiences().get();
    let remarketingListCount = 0;
    
    while (remarketingListIterator.hasNext()) {
      remarketingListIterator.next();
      remarketingListCount++;
    }
    
    // Check for active remarketing campaigns
    const query = "SELECT CampaignId " +
      "FROM CAMPAIGN_AUDIENCE_VIEW " +
      "WHERE CampaignStatus = 'ENABLED'";
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    let activeRemarketingCampaigns = 0;
    const campaignIds = new Set();
    
    while (rows.hasNext()) {
      const row = rows.next();
      const campaignId = row['CampaignId'];
      
      if (!campaignIds.has(campaignId)) {
        campaignIds.add(campaignId);
        activeRemarketingCampaigns++;
      }
    }
    
    // Check for customer match lists
    const customerMatchQuery = "SELECT AudienceId " +
      "FROM AUDIENCE_REPORT " +
      "WHERE AudienceType = 'CUSTOMER_LIST'";
    
    const customerMatchReport = AdsApp.report(customerMatchQuery);
    const customerMatchRows = customerMatchReport.rows();
    
    let customerMatchListCount = 0;
    while (customerMatchRows.hasNext()) {
      customerMatchRows.next();
      customerMatchListCount++;
    }
    
    // Check for in-market and affinity audiences
    const audienceTypeQuery = "SELECT AudienceType, AudienceId " +
      "FROM AUDIENCE_REPORT " +
      "WHERE AudienceType IN ['IN_MARKET', 'AFFINITY']";
    
    const audienceTypeReport = AdsApp.report(audienceTypeQuery);
    const audienceTypeRows = audienceTypeReport.rows();
    
    let inMarketAudienceCount = 0;
    let affinityAudienceCount = 0;
    
    while (audienceTypeRows.hasNext()) {
      const row = audienceTypeRows.next();
      if (row['AudienceType'] === 'IN_MARKET') {
        inMarketAudienceCount++;
      } else if (row['AudienceType'] === 'AFFINITY') {
        affinityAudienceCount++;
      }
    }
    
    // Check for audience bid adjustments
    const bidAdjustmentQuery = "SELECT CampaignId, AdGroupId " +
      "FROM CAMPAIGN_AUDIENCE_VIEW " +
      "WHERE BidModifier != 1.0 AND BidModifier > 0";
    
    const bidAdjustmentReport = AdsApp.report(bidAdjustmentQuery);
    const bidAdjustmentRows = bidAdjustmentReport.rows();
    
    let audiencesWithBidAdjustments = 0;
    let totalAudiences = remarketingListCount + customerMatchListCount + inMarketAudienceCount + affinityAudienceCount;
    
    while (bidAdjustmentRows.hasNext()) {
      bidAdjustmentRows.next();
      audiencesWithBidAdjustments++;
    }
    
    const audienceBidAdjustmentPercentage = totalAudiences > 0 ? audiencesWithBidAdjustments / totalAudiences : 0;
    
    // Populate audience data
    accountData.audiences.remarketingListCount = remarketingListCount;
    accountData.audiences.activeRemarketingCampaigns = activeRemarketingCampaigns;
    accountData.audiences.hasCustomerMatch = customerMatchListCount > 0;
    accountData.audiences.customerMatchListCount = customerMatchListCount;
    accountData.audiences.hasInMarketAudiences = inMarketAudienceCount > 0;
    accountData.audiences.hasAffinityAudiences = affinityAudienceCount > 0;
    accountData.audiences.inMarketAudienceCount = inMarketAudienceCount;
    accountData.audiences.affinityAudienceCount = affinityAudienceCount;
    accountData.audiences.audienceBidAdjustmentPercentage = audienceBidAdjustmentPercentage;
  }
  
  /**
   * Collects landing page data
   * @param {Object} accountData The account data object to populate
   */
  function collectLandingPageData(accountData) {
    Logger.log("Collecting landing page data...");
    
    // Get landing page performance data
    const query = "SELECT FinalUrl, Conversions, Clicks, Device " +
      "FROM FINAL_URL_REPORT " +
      "WHERE Impressions > 100";
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    let totalConversions = 0;
    let totalClicks = 0;
    let mobileConversions = 0;
    let mobileClicks = 0;
    let desktopConversions = 0;
    let desktopClicks = 0;
    let uniqueUrls = new Set();
    
    while (rows.hasNext()) {
      const row = rows.next();
      const url = row['FinalUrl'];
      const conversions = parseFloat(row['Conversions']) || 0;
      const clicks = parseInt(row['Clicks'], 10) || 0;
      const device = row['Device'];
      
      uniqueUrls.add(url);
      totalConversions += conversions;
      totalClicks += clicks;
      
      if (device === 'MOBILE') {
        mobileConversions += conversions;
        mobileClicks += clicks;
      } else if (device === 'DESKTOP') {
        desktopConversions += conversions;
        desktopClicks += clicks;
      }
    }
    
    // Calculate conversion rates
    const conversionRate = totalClicks > 0 ? totalConversions / totalClicks : 0;
    const mobileConversionRate = mobileClicks > 0 ? mobileConversions / mobileClicks : 0;
    const desktopConversionRate = desktopClicks > 0 ? desktopConversions / desktopClicks : 0;
    
    // Check for mobile-friendliness (approximation based on mobile vs desktop conversion rate)
    const isMobileFriendly = mobileConversionRate >= (desktopConversionRate * 0.8);
    
    // Check for A/B testing (approximation based on URL patterns)
    let abTestCount = 0;
    const abTestPatterns = [
      /variant/i, /version/i, /test/i, /exp/i, /ab\d/i, /a-b/i, /split/i
    ];
    
    uniqueUrls.forEach(url => {
      for (const pattern of abTestPatterns) {
        if (pattern.test(url)) {
          abTestCount++;
          break;
        }
      }
    });
    
    const isABTestingImplemented = abTestCount > 0;
    
    // Populate landing page data
    accountData.landingPage.conversionRate = conversionRate;
    accountData.landingPage.isMobileFriendly = isMobileFriendly;
    accountData.landingPage.mobileConversionRate = mobileConversionRate;
    accountData.landingPage.desktopConversionRate = desktopConversionRate;
    accountData.landingPage.isABTestingImplemented = isABTestingImplemented;
    accountData.landingPage.abTestCount = abTestCount;
  }
  
  /**
   * Collects competitive data
   * @param {Object} accountData The account data object to populate
   * @param {Object} dateRange Date range for data collection
   */
  function collectCompetitiveData(accountData, dateRange) {
    Logger.log("Collecting competitive data...");
    
    const query = `SELECT SearchImpressionShare, SearchRankLostImpressionShare, ` +
                  `SearchBudgetLostImpressionShare, SearchExactMatchImpressionShare ` +
                  `FROM CAMPAIGN_PERFORMANCE_REPORT ` +
                  `WHERE Date BETWEEN '${dateRange.start}' AND '${dateRange.end}'`;
    
    let hasAuctionInsightsData = false;
    let impressionShare = 0;
    let topImpressionShare = 0;
    let absoluteTopImpressionShare = 0;
    let recordCount = 0;
    
    try {
      const report = AdsApp.report(query);
      const rows = report.rows();
      
      while (rows.hasNext()) {
        const row = rows.next();
        hasAuctionInsightsData = true;
        recordCount++;
        
        impressionShare += parseFloat(row['SearchImpressionShare']) || 0;
        topImpressionShare += parseFloat(row['SearchTopImpressionShare']) || 0;
        absoluteTopImpressionShare += parseFloat(row['SearchAbsoluteTopImpressionShare']) || 0;
      }
      
      // Calculate averages
      if (recordCount > 0) {
        impressionShare /= recordCount;
        topImpressionShare /= recordCount;
        absoluteTopImpressionShare /= recordCount;
      }
    } catch (e) {
      Logger.log("Error getting auction insights data: " + e.message);
    }
    
    // Check for competitor campaigns
    const competitorKeywordQuery = "SELECT CampaignId, Criteria " +
      "FROM KEYWORDS_PERFORMANCE_REPORT " +
      "WHERE Status IN ['ENABLED', 'PAUSED'] " +
      "AND Criteria CONTAINS_ANY ['competitor', 'vs', 'versus', 'alternative']";
    
    let hasCompetitorCampaigns = false;
    let competitorKeywordCount = 0;
    
    try {
      const competitorReport = AdsApp.report(competitorKeywordQuery);
      const competitorRows = competitorReport.rows();
      
      while (competitorRows.hasNext()) {
        competitorRows.next();
        hasCompetitorCampaigns = true;
        competitorKeywordCount++;
      }
    } catch (e) {
      Logger.log("Error getting competitor keyword data: " + e.message);
    }
    
    // Check for competitive messaging in ads
    const adCopyQuery = "SELECT HeadlinePart1, HeadlinePart2, HeadlinePart3, Description1, Description2 " +
      "FROM AD_PERFORMANCE_REPORT " +
      "WHERE Status IN ['ENABLED', 'PAUSED'] " +
      "AND AdType = 'EXPANDED_TEXT_AD'";
    
    let hasCompetitiveAdCopyAnalysis = false;
    let competitiveAdCount = 0;
    let totalAdCount = 0;
    
    try {
      const adCopyReport = AdsApp.report(adCopyQuery);
      const adCopyRows = adCopyReport.rows();
      
      const competitiveTerms = [
        /better than/i, /vs/i, /versus/i, /compared to/i, /alternative to/i,
        /switch from/i, /unlike/i, /outperform/i, /superior/i, /best in class/i
      ];
      
      while (adCopyRows.hasNext()) {
        const row = adCopyRows.next();
        totalAdCount++;
        
        const adText = [
          row['HeadlinePart1'] || '',
          row['HeadlinePart2'] || '',
          row['HeadlinePart3'] || '',
          row['Description1'] || '',
          row['Description2'] || ''
        ].join(' ');
        
        for (const term of competitiveTerms) {
          if (term.test(adText)) {
            competitiveAdCount++;
            hasCompetitiveAdCopyAnalysis = true;
            break;
          }
        }
      }
    } catch (e) {
      Logger.log("Error getting ad copy data: " + e.message);
    }
    
    const competitiveMessagingScore = totalAdCount > 0 ? (competitiveAdCount / totalAdCount) * 100 : 0;
    
    // Populate competitive data
    accountData.competitive.hasAuctionInsightsData = hasAuctionInsightsData;
    accountData.competitive.impressionShare = impressionShare;
    accountData.competitive.topImpressionShare = topImpressionShare;
    accountData.competitive.absoluteTopImpressionShare = absoluteTopImpressionShare;
    accountData.competitive.hasCompetitorCampaigns = hasCompetitorCampaigns;
    accountData.competitive.competitorKeywordCount = competitorKeywordCount;
    accountData.competitive.hasCompetitiveAdCopyAnalysis = hasCompetitiveAdCopyAnalysis;
    accountData.competitive.competitiveMessagingScore = competitiveMessagingScore;
  }
  
  /**
   * Collects conversion data
   * @param {Object} accountData The account data object to populate
   * @param {Object} dateRange Date range for data collection
   */
  function collectConversionData(accountData, dateRange) {
    Logger.log("Collecting conversion data...");
    
    const query = `SELECT ConversionRate, CostPerConversion, ConversionValue, ` +
                  `ValuePerConversion, AllConversionRate, AllConversionValue ` +
                  `FROM ACCOUNT_PERFORMANCE_REPORT ` +
                  `WHERE Date BETWEEN '${dateRange.start}' AND '${dateRange.end}'`;
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    let conversions = 0;
    let conversionValue = 0;
    let cost = 0;
    let clicks = 0;
    let impressions = 0;
    
    while (rows.hasNext()) {
      const row = rows.next();
      conversions += parseFloat(row['Conversions']) || 0;
      conversionValue += parseFloat(row['ConversionValue']) || 0;
      cost += parseFloat(row['Cost']) || 0;
      clicks += parseInt(row['Clicks'], 10) || 0;
      impressions += parseInt(row['Impressions'], 10) || 0;
    }
    
    // Calculate derived metrics
    const ctr = impressions > 0 ? clicks / impressions : 0;
    const conversionRate = clicks > 0 ? conversions / clicks : 0;
    const roas = cost > 0 ? conversionValue / cost : 0;
    const averageCpc = clicks > 0 ? cost / clicks : 0;
    
    // Populate performance data
    accountData.performance.conversions = conversions;
    accountData.performance.conversionValue = conversionValue;
    accountData.performance.cost = cost;
    accountData.performance.clicks = clicks;
    accountData.performance.impressions = impressions;
    accountData.performance.ctr = ctr;
    accountData.performance.conversionRate = conversionRate;
    accountData.performance.roas = roas;
    accountData.performance.averageCpc = averageCpc;
  }
  
  /**
   * Collects conversion tracking data
   * @param {Object} accountData The account data object to populate
   */
  function collectConversionTrackingData(accountData) {
    Logger.log("Collecting conversion tracking data...");
    
    // Query for conversion actions
    const query = "SELECT ConversionActionName, ConversionActionCategory, " +
      "IncludeInConversionsMetric, ValueSettingStatus, AttributionModelType " +
      "FROM CONVERSION_ACTION_REPORT";
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    // Initialize counters
    let conversionCount = 0;
    let valueTrackingCount = 0;
    let hasPhoneCallTracking = false;
    let hasImportedConversions = false;
    
    // Process conversion actions
    while (rows.hasNext()) {
      const row = rows.next();
      
      // Count active conversion actions
      if (row['IncludeInConversionsMetric'] === 'true') {
        conversionCount++;
        
        // Check for value tracking
        if (row['ValueSettingStatus'] === 'ACTIVE') {
          valueTrackingCount++;
        }
        
        // Check for phone call tracking
        const category = row['ConversionActionCategory'];
        if (category === 'PHONE_CALL_LEAD' || category === 'PHONE_CALL_CONVERSION') {
          hasPhoneCallTracking = true;
        }
        
        // Check for imported conversions
        if (category === 'UPLOAD' || category === 'IMPORT') {
          hasImportedConversions = true;
        }
      }
    }
    
    // Populate conversion tracking data
    accountData.conversionTracking.count = conversionCount;
    accountData.conversionTracking.valueTrackingCount = valueTrackingCount;
    accountData.conversionTracking.hasPhoneCallTracking = hasPhoneCallTracking;
    accountData.conversionTracking.hasImportedConversions = hasImportedConversions;
  }
  
  /**
   * Collects bidding strategy data
   * @param {Object} accountData The account data object to populate
   */
  function collectBiddingStrategyData(accountData) {
    Logger.log("Collecting bidding strategy data...");
    
    const dateRange = accountData.dateRange;
    
    // Create a query to get campaign bidding strategies
    const query = `
      SELECT
        campaign.id,
        campaign.name,
        campaign.bidding_strategy_type,
        campaign.target_cpa.target_cpa_micros,
        campaign.target_roas.target_roas,
        campaign.maximize_conversion_value.target_roas,
        campaign.maximize_conversions.target_cpa_micros,
        metrics.impressions,
        metrics.clicks,
        metrics.cost_micros,
        metrics.conversions,
        metrics.conversions_value,
        metrics.search_impression_share,
        metrics.search_budget_lost_impression_share,
        metrics.search_rank_lost_impression_share
      FROM campaign
      WHERE segments.date BETWEEN '${dateRange.start}' AND '${dateRange.end}'
    `;
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    // Initialize bidding strategy counts
    const biddingStrategyCounts = {
      MANUAL_CPC: 0,
      MAXIMIZE_CONVERSIONS: 0,
      MAXIMIZE_CONVERSION_VALUE: 0,
      TARGET_CPA: 0,
      TARGET_ROAS: 0,
      TARGET_IMPRESSION_SHARE: 0,
      ENHANCED_CPC: 0,
      OTHER: 0
    };
    
    // Initialize impression share metrics
    let totalImpressions = 0;
    let totalSearchImpressionShare = 0;
    let totalSearchBudgetLost = 0;
    let totalSearchRankLost = 0;
    let campaignsWithImpressionShareData = 0;
    
    while (rows.hasNext()) {
      const row = rows.next();
      
      const biddingStrategyType = row['campaign.bidding_strategy_type'];
      const impressions = parseInt(row['metrics.impressions'], 10) || 0;
      const searchImpressionShare = parseFloat(row['metrics.search_impression_share']) || 0;
      const searchBudgetLost = parseFloat(row['metrics.search_budget_lost_impression_share']) || 0;
      const searchRankLost = parseFloat(row['metrics.search_rank_lost_impression_share']) || 0;
      
      // Count bidding strategies
      if (biddingStrategyType in biddingStrategyCounts) {
        biddingStrategyCounts[biddingStrategyType]++;
      } else {
        biddingStrategyCounts.OTHER++;
      }
      
      // Aggregate impression share data (weighted by impressions)
      if (impressions > 0 && !isNaN(searchImpressionShare)) {
        totalImpressions += impressions;
        totalSearchImpressionShare += searchImpressionShare * impressions;
        totalSearchBudgetLost += searchBudgetLost * impressions;
        totalSearchRankLost += searchRankLost * impressions;
        campaignsWithImpressionShareData++;
      }
    }
    
    // Calculate weighted averages
    const avgSearchImpressionShare = totalImpressions > 0 ? totalSearchImpressionShare / totalImpressions : 0;
    const avgSearchBudgetLost = totalImpressions > 0 ? totalSearchBudgetLost / totalImpressions : 0;
    const avgSearchRankLost = totalImpressions > 0 ? totalSearchRankLost / totalImpressions : 0;
    
    // Check for bid adjustments
    let hasDeviceBidAdjustments = false;
    let hasLocationBidAdjustments = false;
    let hasScheduleBidAdjustments = false;
    
    try {
      // Check device bid adjustments
      const deviceQuery = `
        SELECT campaign.id
        FROM campaign_criterion
        WHERE campaign_criterion.type = "DEVICE"
          AND campaign_criterion.bid_modifier != 1.0
        LIMIT 1
      `;
      
      const deviceReport = AdsApp.report(deviceQuery);
      hasDeviceBidAdjustments = deviceReport.rows().hasNext();
      
      // Check location bid adjustments
      const locationQuery = `
        SELECT campaign.id
        FROM campaign_criterion
        WHERE campaign_criterion.type = "LOCATION"
          AND campaign_criterion.bid_modifier != 1.0
        LIMIT 1
      `;
      
      const locationReport = AdsApp.report(locationQuery);
      hasLocationBidAdjustments = locationReport.rows().hasNext();
      
      // Check ad schedule bid adjustments
      const scheduleQuery = `
        SELECT campaign.id
        FROM campaign_criterion
        WHERE campaign_criterion.type = "AD_SCHEDULE"
          AND campaign_criterion.bid_modifier != 1.0
        LIMIT 1
      `;
      
      const scheduleReport = AdsApp.report(scheduleQuery);
      hasScheduleBidAdjustments = scheduleReport.rows().hasNext();
    } catch (e) {
      Logger.log("Error checking bid adjustments: " + e);
    }
    
    // Update account data
    accountData.bidding = {
      strategies: biddingStrategyCounts,
      smartBiddingPercentage: (biddingStrategyCounts.MAXIMIZE_CONVERSIONS + 
                              biddingStrategyCounts.MAXIMIZE_CONVERSION_VALUE + 
                              biddingStrategyCounts.TARGET_CPA + 
                              biddingStrategyCounts.TARGET_ROAS) / 
                              Object.values(biddingStrategyCounts).reduce((a, b) => a + b, 0),
      impressionShare: avgSearchImpressionShare,
      budgetLost: avgSearchBudgetLost,
      rankLost: avgSearchRankLost,
      hasDeviceBidAdjustments: hasDeviceBidAdjustments,
      hasLocationBidAdjustments: hasLocationBidAdjustments,
      hasScheduleBidAdjustments: hasScheduleBidAdjustments
    };
  }
  
  /**
   * Evaluates campaign organization
   * @param {Object} accountData The collected account data
   * @return {Object} Evaluation results for campaign organization
   */
  function evaluateCampaignOrganization(accountData) {
    Logger.log("Evaluating campaign organization...");
    
    // Initialize results object
    const results = {
      score: 0,
      letter: '',
      criteria: {},
      recommendations: []
    };
    
    // 1. Evaluate logical campaign & ad group structure
    const structureResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get structure data
    const campaignCount = accountData.structure.campaignCount;
    const adGroupCount = accountData.structure.adGroupCount;
    const keywordCount = accountData.structure.keywordCount;
    const averageAdGroupsPerCampaign = accountData.structure.averageAdGroupsPerCampaign;
    const averageKeywordsPerAdGroup = accountData.structure.averageKeywordsPerAdGroup;
    
    // Store metrics in details
    structureResult.details.campaignCount = campaignCount;
    structureResult.details.adGroupCount = adGroupCount;
    structureResult.details.keywordCount = keywordCount;
    structureResult.details.averageAdGroupsPerCampaign = averageAdGroupsPerCampaign;
    structureResult.details.averageKeywordsPerAdGroup = averageKeywordsPerAdGroup;
    
    // Evaluate average keywords per ad group
    let keywordsPerAdGroupScore = 100;
    if (averageKeywordsPerAdGroup > CONFIG.bestPractices.keywordsPerAdGroup * 2) {
      keywordsPerAdGroupScore = 50;
      structureResult.recommendations.push({
        text: "Your ad groups contain too many keywords on average (" + averageKeywordsPerAdGroup.toFixed(1) + "). Consider restructuring to have fewer, more tightly themed keywords per ad group.",
        impact: 0.8
      });
    } else if (averageKeywordsPerAdGroup > CONFIG.bestPractices.keywordsPerAdGroup) {
      keywordsPerAdGroupScore = 75;
      structureResult.recommendations.push({
        text: "Your ad groups contain more keywords than recommended (" + averageKeywordsPerAdGroup.toFixed(1) + " vs. ideal " + CONFIG.bestPractices.keywordsPerAdGroup + "). Consider tightening your ad group themes.",
        impact: 0.6
      });
    }
    
    // Evaluate ad groups per campaign
    let adGroupsPerCampaignScore = 100;
    if (averageAdGroupsPerCampaign < 2) {
      adGroupsPerCampaignScore = 70;
      structureResult.recommendations.push({
        text: "Your campaigns have very few ad groups on average (" + averageAdGroupsPerCampaign.toFixed(1) + "). Consider creating more specific ad groups to better organize your keywords.",
        impact: 0.5
      });
    } else if (averageAdGroupsPerCampaign > 20) {
      adGroupsPerCampaignScore = 80;
      structureResult.recommendations.push({
        text: "Your campaigns have a high number of ad groups on average (" + averageAdGroupsPerCampaign.toFixed(1) + "). Consider splitting large campaigns into more focused ones.",
        impact: 0.4
      });
    }
    
    // Calculate structure score
    structureResult.score = (keywordsPerAdGroupScore + adGroupsPerCampaignScore) / 2;
    results.criteria.logicalStructure = structureResult;
    
    // 2. Evaluate naming conventions & segmentation
    const namingResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Simplified naming convention check (would be more sophisticated in a real implementation)
    // Check if campaigns follow a consistent naming pattern
    const campaigns = accountData.campaigns || [];
    let consistentNamingCount = 0;
    let segmentationScore = 0;
    
    // Check for common naming patterns
    const namingPatterns = {
      locationPattern: /\b(north|south|east|west|regional|local|national|global)\b/i,
      productPattern: /\b(product|service|category|brand)\b/i,
      purposePattern: /\b(brand|non-brand|generic|competitor|display|search|shopping)\b/i
    };
    
    const campaignsByPattern = {
      locationPattern: 0,
      productPattern: 0,
      purposePattern: 0
    };
    
    campaigns.forEach(campaign => {
      for (const pattern in namingPatterns) {
        if (namingPatterns[pattern].test(campaign.name)) {
          campaignsByPattern[pattern]++;
          break;
        }
      }
    });
    
    // Calculate percentage of campaigns with consistent naming
    const maxPatternCount = Math.max(...Object.values(campaignsByPattern));
    const consistentNamingPercentage = campaigns.length > 0 ? maxPatternCount / campaigns.length : 0;
    
    namingResult.details.consistentNamingPercentage = consistentNamingPercentage * 100;
    
    // Check for campaign segmentation by type
    const campaignTypes = {};
    campaigns.forEach(campaign => {
      if (!campaignTypes[campaign.type]) {
        campaignTypes[campaign.type] = 0;
      }
      campaignTypes[campaign.type]++;
    });
    
    const campaignTypeCount = Object.keys(campaignTypes).length;
    namingResult.details.campaignTypeCount = campaignTypeCount;
    
    // Score naming conventions
    if (consistentNamingPercentage >= 0.8) {
      namingResult.score = 90;
    } else if (consistentNamingPercentage >= 0.6) {
      namingResult.score = 75;
      namingResult.recommendations.push({
        text: "Only " + Math.round(consistentNamingPercentage * 100) + "% of your campaigns follow a consistent naming convention. Standardize naming for better organization.",
        impact: 0.6
      });
    } else {
      namingResult.score = 50;
      namingResult.recommendations.push({
        text: "Your campaign naming lacks consistency. Implement a standardized naming convention that includes purpose, product/service, and targeting criteria.",
        impact: 0.7
      });
    }
    
    // Add segmentation recommendation if needed
    if (campaignTypeCount < 2 && campaigns.length > 5) {
      namingResult.score -= 10;
      namingResult.recommendations.push({
        text: "Consider segmenting campaigns by different campaign types (Search, Display, Shopping) for better performance management.",
        impact: 0.5
      });
    }
    
    results.criteria.namingConventions = namingResult;
    
    // 3. Evaluate internal competition
    const competitionResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Check for duplicate keywords across ad groups
    const duplicateKeywordPercentage = accountData.structure.duplicateKeywords / accountData.structure.keywordCount || 0;
    competitionResult.details.duplicateKeywordPercentage = duplicateKeywordPercentage * 100;
    
    // Score based on duplicate keywords
    if (duplicateKeywordPercentage <= 0.05) {
      competitionResult.score = 90;
    } else if (duplicateKeywordPercentage <= 0.1) {
      competitionResult.score = 75;
      competitionResult.recommendations.push({
        text: "You have " + Math.round(duplicateKeywordPercentage * 100) + "% duplicate keywords across ad groups. Review and remove duplicates to prevent internal competition.",
        impact: 0.6
      });
    } else {
      competitionResult.score = 50;
      competitionResult.recommendations.push({
        text: "High level of keyword duplication (" + Math.round(duplicateKeywordPercentage * 100) + "%) across ad groups. This causes internal competition and wasted spend.",
        impact: 0.8
      });
    }
    
    results.criteria.internalCompetition = competitionResult;
    
    // Calculate overall category score (weighted average of criteria scores)
    const categoryInfo = EVALUATION_CATEGORIES.find(c => c.name === "Campaign Organization");
    let weightedScoreSum = 0;
    let weightSum = 0;
    
    categoryInfo.criteria.forEach(criterion => {
      const criterionKey = criterion.name.toLowerCase().replace(/\s+|&/g, '').replace(/[^a-z0-9]/g, '');
      const criterionResult = results.criteria[criterionKey];
      
      if (criterionResult) {
        weightedScoreSum += criterionResult.score * criterion.weight;
        weightSum += criterion.weight;
        
        // Add recommendations to the category level
        criterionResult.recommendations.forEach(rec => {
          results.recommendations.push(rec);
        });
      }
    });
    
    results.score = weightSum > 0 ? weightedScoreSum / weightSum : 0;
    
    // Assign letter grade
    if (results.score >= CONFIG.gradeThresholds.A) {
      results.letter = 'A';
    } else if (results.score >= CONFIG.gradeThresholds.B) {
      results.letter = 'B';
    } else if (results.score >= CONFIG.gradeThresholds.C) {
      results.letter = 'C';
    } else if (results.score >= CONFIG.gradeThresholds.D) {
      results.letter = 'D';
    } else {
      results.letter = 'F';
    }
    
    return results;
  }
  
  /**
   * Evaluates conversion tracking
   * @param {Object} accountData The collected account data
   * @return {Object} Evaluation results for conversion tracking
   */
  function evaluateConversionTracking(accountData) {
    Logger.log("Evaluating conversion tracking...");
    
    // Initialize results object
    const results = {
      score: 0,
      letter: '',
      criteria: {},
      recommendations: []
    };
    
    // 1. Evaluate comprehensive conversion coverage
    const coverageResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get conversion data
    const conversionCount = accountData.conversionTracking.count || 0;
    const hasPhoneCallTracking = accountData.conversionTracking.hasPhoneCallTracking || false;
    const hasImportedConversions = accountData.conversionTracking.hasImportedConversions || false;
    
    coverageResult.details.conversionCount = conversionCount;
    coverageResult.details.hasPhoneCallTracking = hasPhoneCallTracking;
    coverageResult.details.hasImportedConversions = hasImportedConversions;
    
    // Score based on conversion count and types
    if (conversionCount >= 3 && hasPhoneCallTracking && hasImportedConversions) {
      coverageResult.score = 95;
    } else if (conversionCount >= 2) {
      coverageResult.score = 80;
      
      if (!hasPhoneCallTracking) {
        coverageResult.recommendations.push({
          text: "Add phone call conversion tracking to capture valuable phone leads.",
          impact: 0.7
        });
      }
      
      if (!hasImportedConversions && accountData.account.isEcommerce) {
        coverageResult.recommendations.push({
          text: "Set up offline conversion imports to track sales that happen outside of your website.",
          impact: 0.6
        });
      }
    } else if (conversionCount >= 1) {
      coverageResult.score = 60;
      coverageResult.recommendations.push({
        text: "You have only " + conversionCount + " conversion action. Set up additional conversion actions to track different user goals.",
        impact: 0.8
      });
    } else {
      coverageResult.score = 20;
      coverageResult.recommendations.push({
        text: "No conversion tracking detected. Set up conversion tracking immediately to measure campaign effectiveness.",
        impact: 1.0
      });
    }
    
    results.criteria.comprehensiveConversionCoverage = coverageResult;
    
    // 2. Evaluate accurate and verified tracking implementation
    const implementationResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Check for conversion value tracking
    const valueTrackingCount = accountData.conversionTracking.valueTrackingCount || 0;
    const valueTrackingPercentage = conversionCount > 0 ? valueTrackingCount / conversionCount : 0;
    
    implementationResult.details.valueTrackingPercentage = valueTrackingPercentage * 100;
    
    // Score based on value tracking
    if (valueTrackingPercentage >= 0.8) {
      implementationResult.score = 90;
    } else if (valueTrackingPercentage >= 0.5) {
      implementationResult.score = 75;
      implementationResult.recommendations.push({
        text: "Only " + Math.round(valueTrackingPercentage * 100) + "% of your conversion actions have value tracking. Add values to all meaningful conversions.",
        impact: 0.6
      });
    } else if (conversionCount > 0) {
      implementationResult.score = 50;
      implementationResult.recommendations.push({
        text: "Most of your conversion actions don't have value tracking. Add conversion values to optimize for revenue/value instead of just conversion count.",
        impact: 0.8
      });
    } else {
      implementationResult.score = 0;
    }
    
    results.criteria.accurateAndVerifiedTrackingImplementation = implementationResult;
    
    // 3. Evaluate enhanced & offline conversion tracking
    const enhancedResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Check for enhanced conversions
    const hasEnhancedConversions = accountData.conversionTracking.hasEnhancedConversions || false;
    const hasDataDrivenAttribution = accountData.conversionTracking.hasDataDrivenAttribution || false;
    
    enhancedResult.details.hasEnhancedConversions = hasEnhancedConversions;
    enhancedResult.details.hasDataDrivenAttribution = hasDataDrivenAttribution;
    
    // Score based on enhanced features
    if (hasEnhancedConversions && hasDataDrivenAttribution) {
      enhancedResult.score = 95;
    } else if (hasEnhancedConversions || hasDataDrivenAttribution) {
      enhancedResult.score = 75;
      
      if (!hasEnhancedConversions) {
        enhancedResult.recommendations.push({
          text: "Implement enhanced conversions to improve measurement accuracy, especially in browsers with tracking limitations.",
          impact: 0.7
        });
      }
      
      if (!hasDataDrivenAttribution) {
        enhancedResult.recommendations.push({
          text: "Switch to data-driven attribution to more accurately credit conversions across the customer journey.",
          impact: 0.6
        });
      }
    } else if (conversionCount > 0) {
      enhancedResult.score = 50;
      enhancedResult.recommendations.push({
        text: "Your account isn't using enhanced conversion features. Implement enhanced conversions and data-driven attribution for better measurement.",
        impact: 0.7
      });
    } else {
      enhancedResult.score = 0;
    }
    
    results.criteria.enhancedOfflineConversionTracking = enhancedResult;
    
    // Calculate overall category score (weighted average of criteria scores)
    const categoryInfo = EVALUATION_CATEGORIES.find(c => c.name === "Conversion Tracking");
    let weightedScoreSum = 0;
    let weightSum = 0;
    
    categoryInfo.criteria.forEach(criterion => {
      const criterionKey = criterion.name.toLowerCase().replace(/\s+|&/g, '').replace(/[^a-z0-9]/g, '');
      const criterionResult = results.criteria[criterionKey];
      
      if (criterionResult) {
        weightedScoreSum += criterionResult.score * criterion.weight;
        weightSum += criterion.weight;
        
        // Add recommendations to the category level
        criterionResult.recommendations.forEach(rec => {
          results.recommendations.push(rec);
        });
      }
    });
    
    results.score = weightSum > 0 ? weightedScoreSum / weightSum : 0;
    
    // Assign letter grade
    if (results.score >= CONFIG.gradeThresholds.A) {
      results.letter = 'A';
    } else if (results.score >= CONFIG.gradeThresholds.B) {
      results.letter = 'B';
    } else if (results.score >= CONFIG.gradeThresholds.C) {
      results.letter = 'C';
    } else if (results.score >= CONFIG.gradeThresholds.D) {
      results.letter = 'D';
    } else {
      results.letter = 'F';
    }
    
    return results;
  }
  
  /**
   * Evaluates keyword strategy
   * @param {Object} accountData The collected account data
   * @return {Object} Evaluation results for keyword strategy
   */
  function evaluateKeywordStrategy(accountData) {
    Logger.log("Evaluating keyword strategy...");
    
    // Initialize results object
    const results = {
      score: 0,
      letter: '',
      criteria: {},
      recommendations: []
    };
    
    // 1. Evaluate extensive keyword research & relevance
    const researchResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get keyword data
    const keywordCount = accountData.structure.keywordCount || 0;
    const keywordLengthDistribution = accountData.keywords.lengthDistribution || {
      short: 0,
      medium: 0,
      long: 0
    };
    
    // Calculate percentage of long-tail keywords
    const longTailPercentage = keywordCount > 0 ? 
      (keywordLengthDistribution.medium + keywordLengthDistribution.long) / keywordCount : 0;
    
    researchResult.details.keywordCount = keywordCount;
    researchResult.details.longTailPercentage = longTailPercentage * 100;
    
    // Score based on keyword count and long-tail percentage
    if (keywordCount >= 500 && longTailPercentage >= 0.7) {
      researchResult.score = 90;
    } else if (keywordCount >= 200 && longTailPercentage >= 0.5) {
      researchResult.score = 75;
      
      if (longTailPercentage < 0.7) {
        researchResult.recommendations.push({
          text: "Increase your long-tail keyword coverage. Only " + Math.round(longTailPercentage * 100) + "% of your keywords are medium or long-tail phrases.",
          impact: 0.6
        });
      }
    } else if (keywordCount >= 100) {
      researchResult.score = 60;
      researchResult.recommendations.push({
        text: "Expand your keyword list with more specific, long-tail keywords that match user search intent.",
        impact: 0.7
      });
    } else {
      researchResult.score = 40;
      researchResult.recommendations.push({
        text: "Your keyword list is limited (" + keywordCount + " keywords). Conduct comprehensive keyword research to expand your reach.",
        impact: 0.8
      });
    }
    
    results.criteria.extensiveKeywordResearchRelevance = researchResult;
    
    // 2. Evaluate strategic match type use
    const matchTypeResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get match type distribution
    const matchTypeDistribution = accountData.keywords.matchTypeDistribution || {
      EXACT: 0,
      PHRASE: 0,
      BROAD: 0
    };
    
    // Calculate percentages
    const totalMatchTypes = matchTypeDistribution.EXACT + matchTypeDistribution.PHRASE + matchTypeDistribution.BROAD;
    const exactPercentage = totalMatchTypes > 0 ? matchTypeDistribution.EXACT / totalMatchTypes : 0;
    const phrasePercentage = totalMatchTypes > 0 ? matchTypeDistribution.PHRASE / totalMatchTypes : 0;
    const broadPercentage = totalMatchTypes > 0 ? matchTypeDistribution.BROAD / totalMatchTypes : 0;
    
    matchTypeResult.details.exactPercentage = exactPercentage * 100;
    matchTypeResult.details.phrasePercentage = phrasePercentage * 100;
    matchTypeResult.details.broadPercentage = broadPercentage * 100;
    
    // Score based on match type balance
    if (exactPercentage >= 0.3 && phrasePercentage >= 0.2 && broadPercentage >= 0.2) {
      matchTypeResult.score = 90;
    } else if (exactPercentage >= 0.2 && (phrasePercentage + broadPercentage) >= 0.3) {
      matchTypeResult.score = 75;
      
      if (exactPercentage < 0.3) {
        matchTypeResult.recommendations.push({
          text: "Increase your exact match keywords from " + Math.round(exactPercentage * 100) + "% to at least 30% for better control over traffic quality.",
          impact: 0.6
        });
      }
    } else if (totalMatchTypes > 0) {
      matchTypeResult.score = 60;
      
      if (exactPercentage < 0.2) {
        matchTypeResult.recommendations.push({
          text: "Your exact match keyword percentage is too low (" + Math.round(exactPercentage * 100) + "%). Add more exact match keywords for your core terms.",
          impact: 0.7
        });
      }
      
      if (broadPercentage > 0.7) {
        matchTypeResult.recommendations.push({
          text: "Your account relies too heavily on broad match (" + Math.round(broadPercentage * 100) + "%). Balance with more exact and phrase match keywords.",
          impact: 0.7
        });
      }
    } else {
      matchTypeResult.score = 0;
    }
    
    results.criteria.strategicMatchTypeUse = matchTypeResult;
    
    // 3. Evaluate brand vs non-brand segmentation
    const brandResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get brand keyword data
    const brandKeywordPercentage = accountData.keywords.brandKeywordPercentage || 0;
    const hasBrandCampaigns = accountData.keywords.hasBrandCampaigns || false;
    
    brandResult.details.brandKeywordPercentage = brandKeywordPercentage * 100;
    brandResult.details.hasBrandCampaigns = hasBrandCampaigns;
    
    // Score based on brand segmentation
    if (hasBrandCampaigns && brandKeywordPercentage <= 0.3) {
      brandResult.score = 90;
    } else if (hasBrandCampaigns) {
      brandResult.score = 75;
      
      if (brandKeywordPercentage > 0.3) {
        brandResult.recommendations.push({
          text: "Your account has a high percentage of brand keywords (" + Math.round(brandKeywordPercentage * 100) + "%). Focus more on non-brand terms to expand reach.",
          impact: 0.5
        });
      }
    } else if (brandKeywordPercentage > 0) {
      brandResult.score = 60;
      brandResult.recommendations.push({
        text: "Create dedicated brand campaigns to separate brand and non-brand performance metrics and bidding strategies.",
        impact: 0.7
      });
    } else {
      brandResult.score = 50;
      brandResult.recommendations.push({
        text: "No brand keywords detected. If applicable, add brand terms in dedicated campaigns for high-quality traffic.",
        impact: 0.6
      });
    }
    
    results.criteria.brandVsNonbrandSegmentation = brandResult;
    
    // 4. Evaluate continuous keyword optimization
    const optimizationResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get optimization data
    const lowQualityKeywordPercentage = accountData.keywords.lowQualityKeywordPercentage || 0;
    const nonConvertingKeywordPercentage = accountData.keywords.nonConvertingKeywordPercentage || 0;
    
    optimizationResult.details.lowQualityKeywordPercentage = lowQualityKeywordPercentage * 100;
    optimizationResult.details.nonConvertingKeywordPercentage = nonConvertingKeywordPercentage * 100;
    
    // Score based on keyword optimization
    if (lowQualityKeywordPercentage <= 0.1 && nonConvertingKeywordPercentage <= 0.2) {
      optimizationResult.score = 90;
    } else if (lowQualityKeywordPercentage <= 0.2 && nonConvertingKeywordPercentage <= 0.3) {
      optimizationResult.score = 75;
      
      if (lowQualityKeywordPercentage > 0.1) {
        optimizationResult.recommendations.push({
          text: "Address the " + Math.round(lowQualityKeywordPercentage * 100) + "% of keywords with low quality scores by improving ad relevance and landing pages.",
          impact: 0.6
        });
      }
    } else {
      optimizationResult.score = 50;
      
      if (nonConvertingKeywordPercentage > 0.3) {
        optimizationResult.recommendations.push({
          text: "Review and optimize the " + Math.round(nonConvertingKeywordPercentage * 100) + "% of keywords with clicks but no conversions.",
          impact: 0.8
        });
      }
      
      if (lowQualityKeywordPercentage > 0.2) {
        optimizationResult.recommendations.push({
          text: "Your account has a high percentage (" + Math.round(lowQualityKeywordPercentage * 100) + "%) of low quality score keywords. Pause or improve these keywords.",
          impact: 0.7
        });
      }
    }
    
    results.criteria.continuousKeywordOptimization = optimizationResult;
    
    // Calculate overall category score (weighted average of criteria scores)
    const categoryInfo = EVALUATION_CATEGORIES.find(c => c.name === "Keyword Strategy");
    let weightedScoreSum = 0;
    let weightSum = 0;
    
    categoryInfo.criteria.forEach(criterion => {
      const criterionKey = criterion.name.toLowerCase().replace(/\s+|&/g, '').replace(/[^a-z0-9]/g, '');
      const criterionResult = results.criteria[criterionKey];
      
      if (criterionResult) {
        weightedScoreSum += criterionResult.score * criterion.weight;
        weightSum += criterion.weight;
        
        // Add recommendations to the category level
        criterionResult.recommendations.forEach(rec => {
          results.recommendations.push(rec);
        });
      }
    });
    
    results.score = weightSum > 0 ? weightedScoreSum / weightSum : 0;
    
    // Assign letter grade
    if (results.score >= CONFIG.gradeThresholds.A) {
      results.letter = 'A';
    } else if (results.score >= CONFIG.gradeThresholds.B) {
      results.letter = 'B';
    } else if (results.score >= CONFIG.gradeThresholds.C) {
      results.letter = 'C';
    } else if (results.score >= CONFIG.gradeThresholds.D) {
      results.letter = 'D';
    } else {
      results.letter = 'F';
    }
    
    return results;
  }
  
  /**
   * Evaluates negative keywords
   * @param {Object} accountData The collected account data
   * @return {Object} Evaluation results for negative keywords
   */
  function evaluateNegativeKeywords(accountData) {
    Logger.log("Evaluating negative keywords...");
    
    // Initialize results object
    const results = {
      score: 0,
      letter: '',
      criteria: {},
      recommendations: []
    };
    
    // 1. Evaluate routine search query mining
    const miningResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get negative keyword data
    const negativeKeywordCount = accountData.negativeKeywords.count || 0;
    const campaignCount = accountData.structure.campaignCount || 0;
    const adGroupCount = accountData.structure.adGroupCount || 0;
    
    // Calculate negatives per campaign/ad group
    const negativesPerCampaign = campaignCount > 0 ? negativeKeywordCount / campaignCount : 0;
    const negativesPerAdGroup = adGroupCount > 0 ? negativeKeywordCount / adGroupCount : 0;
    
    miningResult.details.negativeKeywordCount = negativeKeywordCount;
    miningResult.details.negativesPerCampaign = negativesPerCampaign;
    miningResult.details.negativesPerAdGroup = negativesPerAdGroup;
    
    // Score based on negative keyword volume
    if (negativesPerCampaign >= 30) {
      miningResult.score = 90;
    } else if (negativesPerCampaign >= 15) {
      miningResult.score = 75;
      miningResult.recommendations.push({
        text: "Increase your negative keyword count through regular search query mining. You have " + 
              negativeKeywordCount + " negatives (" + negativesPerCampaign.toFixed(1) + " per campaign).",
        impact: 0.6
      });
    } else if (negativesPerCampaign > 0) {
      miningResult.score = 50;
      miningResult.recommendations.push({
        text: "Your negative keyword coverage is limited (" + negativesPerCampaign.toFixed(1) + 
              " per campaign). Implement weekly search query mining to identify irrelevant terms.",
        impact: 0.8
      });
    } else {
      miningResult.score = 20;
      miningResult.recommendations.push({
        text: "No negative keywords found. Add negative keywords immediately to prevent wasted spend on irrelevant searches.",
        impact: 0.9
      });
    }
    
    results.criteria.routineSearchQueryMining = miningResult;
    
    // 2. Evaluate negative keyword organization
    const organizationResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get negative keyword organization data
    const campaignLevelCount = accountData.negativeKeywords.campaignLevel || 0;
    const adGroupLevelCount = accountData.negativeKeywords.adGroupLevel || 0;
    const sharedSetCount = accountData.negativeKeywords.sharedSetCount || 0;
    
    organizationResult.details.campaignLevelCount = campaignLevelCount;
    organizationResult.details.adGroupLevelCount = adGroupLevelCount;
    organizationResult.details.sharedSetCount = sharedSetCount;
    
    // Calculate percentage at each level
    const totalNegatives = campaignLevelCount + adGroupLevelCount;
    const campaignLevelPercentage = totalNegatives > 0 ? campaignLevelCount / totalNegatives : 0;
    const adGroupLevelPercentage = totalNegatives > 0 ? adGroupLevelCount / totalNegatives : 0;
    
    organizationResult.details.campaignLevelPercentage = campaignLevelPercentage * 100;
    organizationResult.details.adGroupLevelPercentage = adGroupLevelPercentage * 100;
    
    // Score based on negative keyword organization
    if (sharedSetCount >= 3 && campaignLevelCount > 0 && adGroupLevelCount > 0) {
      organizationResult.score = 90;
    } else if (sharedSetCount >= 1 && (campaignLevelCount > 0 || adGroupLevelCount > 0)) {
      organizationResult.score = 75;
      
      if (campaignLevelCount === 0 || adGroupLevelCount === 0) {
        organizationResult.recommendations.push({
          text: "Use both campaign-level and ad group-level negative keywords for more granular control.",
          impact: 0.5
        });
      }
    } else if (totalNegatives > 0) {
      organizationResult.score = 60;
      
      if (sharedSetCount === 0) {
        organizationResult.recommendations.push({
          text: "Create negative keyword lists to efficiently manage negatives across multiple campaigns.",
          impact: 0.7
        });
      }
    } else {
      organizationResult.score = 0;
    }
    
    results.criteria.negativeKeywordOrganization = organizationResult;
    
    // 3. Evaluate negative match types
    const matchTypeResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get negative match type data (simplified for this example)
    const hasExactNegatives = accountData.negativeKeywords.hasExactNegatives || false;
    const hasPhraseNegatives = accountData.negativeKeywords.hasPhraseNegatives || false;
    const exactNegativePercentage = accountData.negativeKeywords.exactNegativePercentage || 0;
    const phraseNegativePercentage = accountData.negativeKeywords.phraseNegativePercentage || 0;
    
    matchTypeResult.details.hasExactNegatives = hasExactNegatives;
    matchTypeResult.details.hasPhraseNegatives = hasPhraseNegatives;
    matchTypeResult.details.exactNegativePercentage = exactNegativePercentage * 100;
    matchTypeResult.details.phraseNegativePercentage = phraseNegativePercentage * 100;
    
    // Score based on negative match type usage
    if (hasExactNegatives && hasPhraseNegatives) {
      matchTypeResult.score = 90;
      
      if (exactNegativePercentage < 0.2) {
        matchTypeResult.recommendations.push({
          text: "Consider using more exact match negatives for precise exclusion of specific terms.",
          impact: 0.5
        });
      }
    } else if (hasPhraseNegatives) {
      matchTypeResult.score = 70;
      matchTypeResult.recommendations.push({
        text: "Add exact match negative keywords for more precise control over excluded terms.",
        impact: 0.6
      });
    } else if (hasExactNegatives) {
      matchTypeResult.score = 60;
      matchTypeResult.recommendations.push({
        text: "Add phrase match negative keywords to exclude broader variations of irrelevant terms.",
        impact: 0.7
      });
    } else if (negativeKeywordCount > 0) {
      matchTypeResult.score = 50;
      matchTypeResult.recommendations.push({
        text: "Use both phrase and exact match negative keywords for comprehensive exclusion control.",
        impact: 0.7
      });
    } else {
      matchTypeResult.score = 0;
    }
    
    results.criteria.strategicNegativeMatchTypes = matchTypeResult;
    
    // Calculate overall category score (weighted average of criteria scores)
    const categoryInfo = EVALUATION_CATEGORIES.find(c => c.name === "Negative Keywords");
    let weightedScoreSum = 0;
    let weightSum = 0;
    
    categoryInfo.criteria.forEach(criterion => {
      const criterionKey = criterion.name.toLowerCase().replace(/\s+|&/g, '').replace(/[^a-z0-9]/g, '');
      const criterionResult = results.criteria[criterionKey];
      
      if (criterionResult) {
        weightedScoreSum += criterionResult.score * criterion.weight;
        weightSum += criterion.weight;
        
        // Add recommendations to the category level
        criterionResult.recommendations.forEach(rec => {
          results.recommendations.push(rec);
        });
      }
    });
    
    results.score = weightSum > 0 ? weightedScoreSum / weightSum : 0;
    
    // Assign letter grade
    if (results.score >= CONFIG.gradeThresholds.A) {
      results.letter = 'A';
    } else if (results.score >= CONFIG.gradeThresholds.B) {
      results.letter = 'B';
    } else if (results.score >= CONFIG.gradeThresholds.C) {
      results.letter = 'C';
    } else if (results.score >= CONFIG.gradeThresholds.D) {
      results.letter = 'D';
    } else {
      results.letter = 'F';
    }
    
    return results;
  }
  
  /**
   * Evaluates bidding strategy
   * @param {Object} accountData The collected account data
   * @return {Object} Evaluation results for bidding strategy
   */
  function evaluateBiddingStrategy(accountData) {
    Logger.log("Evaluating bidding strategy...");
    
    // Initialize results object
    const results = {
      score: 0,
      letter: '',
      criteria: {},
      recommendations: []
    };
    
    // 1. Evaluate smart bidding adoption
    const smartBiddingResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get bidding strategy data
    const biddingStrategies = accountData.bidding.strategies || {};
    const campaignCount = accountData.structure.campaignCount || 0;
    
    // Calculate smart bidding percentage
    const smartBiddingCount = (biddingStrategies.targetCpa || 0) + 
                             (biddingStrategies.targetRoas || 0) + 
                             (biddingStrategies.maximizeConversions || 0) + 
                             (biddingStrategies.maximizeConversionValue || 0);
    
    const smartBiddingPercentage = campaignCount > 0 ? smartBiddingCount / campaignCount : 0;
    
    smartBiddingResult.details.smartBiddingPercentage = smartBiddingPercentage * 100;
    smartBiddingResult.details.manualBiddingPercentage = (1 - smartBiddingPercentage) * 100;
    
    // Score based on smart bidding adoption
    if (smartBiddingPercentage >= 0.8) {
      smartBiddingResult.score = 90;
    } else if (smartBiddingPercentage >= 0.5) {
      smartBiddingResult.score = 75;
      smartBiddingResult.recommendations.push({
        text: "Increase smart bidding adoption from " + Math.round(smartBiddingPercentage * 100) + 
              "% to at least 80% of campaigns to leverage Google's machine learning.",
        impact: 0.7
      });
    } else if (smartBiddingPercentage > 0) {
      smartBiddingResult.score = 50;
      smartBiddingResult.recommendations.push({
        text: "Significantly increase smart bidding usage. Only " + Math.round(smartBiddingPercentage * 100) + 
              "% of your campaigns use smart bidding strategies.",
        impact: 0.8
      });
    } else {
      smartBiddingResult.score = 20;
      smartBiddingResult.recommendations.push({
        text: "Implement smart bidding strategies (Target CPA, Target ROAS) to optimize for conversions instead of manual bidding.",
        impact: 0.9
      });
    }
    
    results.criteria.smartBiddingAdoption = smartBiddingResult;
    
    // 2. Evaluate bid strategy alignment with goals
    const alignmentResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get goal alignment data
    const hasConversionTracking = accountData.conversionTracking.count > 0;
    const hasValueTracking = accountData.conversionTracking.valueTrackingCount > 0;
    const hasTargetRoas = (biddingStrategies.targetRoas || 0) > 0;
    const hasTargetCpa = (biddingStrategies.targetCpa || 0) > 0;
    
    alignmentResult.details.hasConversionTracking = hasConversionTracking;
    alignmentResult.details.hasValueTracking = hasValueTracking;
    alignmentResult.details.hasTargetRoas = hasTargetRoas;
    alignmentResult.details.hasTargetCpa = hasTargetCpa;
    
    // Score based on bid strategy alignment
    if (hasValueTracking && hasTargetRoas) {
      alignmentResult.score = 90;
    } else if (hasConversionTracking && hasTargetCpa) {
      alignmentResult.score = 80;
      
      if (hasValueTracking && !hasTargetRoas) {
        alignmentResult.recommendations.push({
          text: "You're tracking conversion values but not using Target ROAS bidding. Switch appropriate campaigns to Target ROAS to optimize for value.",
          impact: 0.7
        });
      }
    } else if (hasConversionTracking) {
      alignmentResult.score = 60;
      alignmentResult.recommendations.push({
        text: "Align your bidding strategy with your conversion goals by implementing Target CPA bidding for conversion-focused campaigns.",
        impact: 0.8
      });
    } else {
      alignmentResult.score = 30;
      alignmentResult.recommendations.push({
        text: "Set up conversion tracking before implementing smart bidding strategies.",
        impact: 1.0
      });
    }
    
    results.criteria.bidStrategyAlignmentWithGoals = alignmentResult;
    
    // 3. Evaluate bid adjustments
    const adjustmentsResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get bid adjustment data
    const hasMobileBidAdjustments = accountData.bidding.hasMobileBidAdjustments || false;
    const hasLocationBidAdjustments = accountData.bidding.hasLocationBidAdjustments || false;
    const hasAudienceBidAdjustments = accountData.bidding.hasAudienceBidAdjustments || false;
    const hasScheduleBidAdjustments = accountData.bidding.hasScheduleBidAdjustments || false;
    
    adjustmentsResult.details.hasMobileBidAdjustments = hasMobileBidAdjustments;
    adjustmentsResult.details.hasLocationBidAdjustments = hasLocationBidAdjustments;
    adjustmentsResult.details.hasAudienceBidAdjustments = hasAudienceBidAdjustments;
    adjustmentsResult.details.hasScheduleBidAdjustments = hasScheduleBidAdjustments;
    
    // Count number of adjustment types used
    const adjustmentCount = [
      hasMobileBidAdjustments,
      hasLocationBidAdjustments,
      hasAudienceBidAdjustments,
      hasScheduleBidAdjustments
    ].filter(Boolean).length;
    
    // Score based on bid adjustments
    if (adjustmentCount >= 3) {
      adjustmentsResult.score = 90;
    } else if (adjustmentCount >= 2) {
      adjustmentsResult.score = 75;
      
      if (!hasMobileBidAdjustments) {
        adjustmentsResult.recommendations.push({
          text: "Add mobile bid adjustments to optimize performance across devices.",
          impact: 0.6
        });
      }
      
      if (!hasLocationBidAdjustments) {
        adjustmentsResult.recommendations.push({
          text: "Implement location bid adjustments to optimize for geographic performance differences.",
          impact: 0.6
        });
      }
    } else if (adjustmentCount >= 1) {
      adjustmentsResult.score = 60;
      adjustmentsResult.recommendations.push({
        text: "Expand your use of bid adjustments. You're only using " + adjustmentCount + " out of 4 possible adjustment types.",
        impact: 0.7
      });
    } else {
      adjustmentsResult.score = 40;
      adjustmentsResult.recommendations.push({
        text: "Implement bid adjustments for device, location, audience, and ad schedule to optimize performance.",
        impact: 0.8
      });
    }
    
    results.criteria.comprehensiveBidAdjustments = adjustmentsResult;
    
    // Calculate overall category score (weighted average of criteria scores)
    const categoryInfo = EVALUATION_CATEGORIES.find(c => c.name === "Bidding Strategy");
    let weightedScoreSum = 0;
    let weightSum = 0;
    
    categoryInfo.criteria.forEach(criterion => {
      const criterionKey = criterion.name.toLowerCase().replace(/\s+|&/g, '').replace(/[^a-z0-9]/g, '');
      const criterionResult = results.criteria[criterionKey];
      
      if (criterionResult) {
        weightedScoreSum += criterionResult.score * criterion.weight;
        weightSum += criterion.weight;
        
        // Add recommendations to the category level
        criterionResult.recommendations.forEach(rec => {
          results.recommendations.push(rec);
        });
      }
    });
    
    results.score = weightSum > 0 ? weightedScoreSum / weightSum : 0;
    
    // Assign letter grade
    if (results.score >= CONFIG.gradeThresholds.A) {
      results.letter = 'A';
    } else if (results.score >= CONFIG.gradeThresholds.B) {
      results.letter = 'B';
    } else if (results.score >= CONFIG.gradeThresholds.C) {
      results.letter = 'C';
    } else if (results.score >= CONFIG.gradeThresholds.D) {
      results.letter = 'D';
    } else {
      results.letter = 'F';
    }
    
    return results;
  }
  
  /**
   * Evaluates ad creative and extensions
   * @param {Object} accountData The collected account data
   * @return {Object} Evaluation results for ad creative and extensions
   */
  function evaluateAdCreative(accountData) {
    Logger.log("Evaluating ad creative and extensions...");
    
    // Initialize results object
    const results = {
      score: 0,
      letter: '',
      criteria: {},
      recommendations: []
    };
    
    // 1. Evaluate responsive search ad adoption
    const rsaResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get RSA data
    const rsaPercentage = accountData.ads.rsaPercentage || 0;
    const averageHeadlinesPerRsa = accountData.ads.averageHeadlinesPerRsa || 0;
    const averageDescriptionsPerRsa = accountData.ads.averageDescriptionsPerRsa || 0;
    
    rsaResult.details.rsaPercentage = rsaPercentage * 100;
    rsaResult.details.averageHeadlinesPerRsa = averageHeadlinesPerRsa;
    rsaResult.details.averageDescriptionsPerRsa = averageDescriptionsPerRsa;
    
    // Score based on RSA adoption and asset count
    if (rsaPercentage >= 0.9 && averageHeadlinesPerRsa >= 12 && averageDescriptionsPerRsa >= 4) {
      rsaResult.score = 95;
    } else if (rsaPercentage >= 0.8 && averageHeadlinesPerRsa >= 10 && averageDescriptionsPerRsa >= 3) {
      rsaResult.score = 80;
      
      if (averageHeadlinesPerRsa < 12) {
        rsaResult.recommendations.push({
          text: "Add more headlines to your RSAs. You're using " + averageHeadlinesPerRsa.toFixed(1) + " on average, but should aim for all 15 possible headlines.",
          impact: 0.6
        });
      }
      
      if (averageDescriptionsPerRsa < 4) {
        rsaResult.recommendations.push({
          text: "Add more descriptions to your RSAs. You're using " + averageDescriptionsPerRsa.toFixed(1) + " on average, but should aim for all 4 possible descriptions.",
          impact: 0.5
        });
      }
    } else if (rsaPercentage >= 0.6) {
      rsaResult.score = 60;
      rsaResult.recommendations.push({
        text: "Increase your RSA adoption from " + Math.round(rsaPercentage * 100) + "% to at least 90% of all ad groups.",
        impact: 0.8
      });
    } else {
      rsaResult.score = 40;
      rsaResult.recommendations.push({
        text: "Upgrade to Responsive Search Ads in all ad groups. Only " + Math.round(rsaPercentage * 100) + "% of your ad groups use RSAs.",
        impact: 0.9
      });
    }
    
    results.criteria.responsiveSearchAdAdoption = rsaResult;
    
    // 2. Evaluate ad rotation and testing
    const rotationResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get ad rotation data
    const averageAdsPerAdGroup = accountData.ads.averageAdsPerAdGroup || 0;
    const singleAdAdGroupPercentage = accountData.ads.singleAdAdGroupPercentage || 0;
    
    rotationResult.details.averageAdsPerAdGroup = averageAdsPerAdGroup;
    rotationResult.details.singleAdAdGroupPercentage = singleAdAdGroupPercentage * 100;
    
    // Score based on ad count per ad group
    if (averageAdsPerAdGroup >= CONFIG.bestPractices.adsPerAdGroup && singleAdAdGroupPercentage <= 0.05) {
      rotationResult.score = 90;
    } else if (averageAdsPerAdGroup >= 2 && singleAdAdGroupPercentage <= 0.2) {
      rotationResult.score = 75;
      
      if (averageAdsPerAdGroup < CONFIG.bestPractices.adsPerAdGroup) {
        rotationResult.recommendations.push({
          text: "Increase your average ads per ad group from " + averageAdsPerAdGroup.toFixed(1) + " to at least " + CONFIG.bestPractices.adsPerAdGroup + " for better testing and performance.",
          impact: 0.7
        });
      }
    } else {
      rotationResult.score = 50;
      rotationResult.recommendations.push({
        text: Math.round(singleAdAdGroupPercentage * 100) + "% of your ad groups have only one ad. Add at least one more ad to each ad group for testing.",
        impact: 0.8
      });
    }
    
    results.criteria.adRotationAndTesting = rotationResult;
    
    // 3. Evaluate ad extensions
    const extensionsResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get extension data
    const extensionTypeCount = accountData.extensions.totalCount || 0;
    const impressionWithExtensionsPercentage = accountData.extensions.impressionWithExtensions / accountData.performance.impressions || 0;
    
    extensionsResult.details.extensionTypeCount = extensionTypeCount;
    extensionsResult.details.impressionWithExtensionsPercentage = impressionWithExtensionsPercentage * 100;
    
    // Score based on extension usage
    if (extensionTypeCount >= CONFIG.bestPractices.minExtensionTypes && impressionWithExtensionsPercentage >= 0.7) {
      extensionsResult.score = 90;
    } else if (extensionTypeCount >= 3 && impressionWithExtensionsPercentage >= 0.5) {
      extensionsResult.score = 75;
      
      if (extensionTypeCount < CONFIG.bestPractices.minExtensionTypes) {
        extensionsResult.recommendations.push({
          text: "Add more extension types. You're using " + extensionTypeCount + " types, but should aim for at least " + CONFIG.bestPractices.minExtensionTypes + ".",
          impact: 0.7
        });
      }
    } else if (extensionTypeCount >= 1) {
      extensionsResult.score = 50;
      extensionsResult.recommendations.push({
        text: "Expand your ad extension usage. Only " + Math.round(impressionWithExtensionsPercentage * 100) + "% of impressions include extensions.",
        impact: 0.8
      });
    } else {
      extensionsResult.score = 20;
      extensionsResult.recommendations.push({
        text: "Implement ad extensions immediately. Extensions increase CTR and provide additional information to users.",
        impact: 0.9
      });
    }
    
    results.criteria.adExtensions = extensionsResult;
    
    // Calculate overall category score (weighted average of criteria scores)
    const categoryInfo = EVALUATION_CATEGORIES.find(c => c.name === "Ad Creative & Extensions");
    let weightedScoreSum = 0;
    let weightSum = 0;
    
    categoryInfo.criteria.forEach(criterion => {
      const criterionKey = criterion.name.toLowerCase().replace(/\s+|&/g, '').replace(/[^a-z0-9]/g, '');
      const criterionResult = results.criteria[criterionKey];
      
      if (criterionResult) {
        weightedScoreSum += criterionResult.score * criterion.weight;
        weightSum += criterion.weight;
        
        // Add recommendations to the category level
        criterionResult.recommendations.forEach(rec => {
          results.recommendations.push(rec);
        });
      }
    });
    
    results.score = weightSum > 0 ? weightedScoreSum / weightSum : 0;
    
    // Assign letter grade
    if (results.score >= CONFIG.gradeThresholds.A) {
      results.letter = 'A';
    } else if (results.score >= CONFIG.gradeThresholds.B) {
      results.letter = 'B';
    } else if (results.score >= CONFIG.gradeThresholds.C) {
      results.letter = 'C';
    } else if (results.score >= CONFIG.gradeThresholds.D) {
      results.letter = 'D';
    } else {
      results.letter = 'F';
    }
    
    return results;
  }
  
  /**
   * Evaluates quality score
   * @param {Object} accountData The collected account data
   * @return {Object} Evaluation results for quality score
   */
  function evaluateQualityScore(accountData) {
    Logger.log("Evaluating quality score...");
    
    // Initialize results object
    const results = {
      score: 0,
      letter: '',
      criteria: {},
      recommendations: []
    };
    
    // 1. Evaluate average quality score
    const averageQsResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get quality score data
    const averageQualityScore = accountData.structure.averageQualityScore || 0;
    const benchmarkQualityScore = CONFIG.industryBenchmarks.qualityScore || 0;
    
    averageQsResult.details.averageQualityScore = averageQualityScore;
    averageQsResult.details.benchmarkQualityScore = benchmarkQualityScore;
    
    // Calculate QS comparison to benchmark
    const qsComparison = benchmarkQualityScore > 0 ? averageQualityScore / benchmarkQualityScore : 0;
    averageQsResult.details.qualityScoreComparison = qsComparison;
    
    // Score based on average quality score
    if (averageQualityScore >= CONFIG.bestPractices.minQualityScore + 1) {
      averageQsResult.score = 90;
    } else if (averageQualityScore >= CONFIG.bestPractices.minQualityScore) {
      averageQsResult.score = 80;
      averageQsResult.recommendations.push({
        text: "Your average quality score of " + averageQualityScore.toFixed(1) + " is good, but aim to improve it further through better ad relevance and landing page experience.",
        impact: 0.6
      });
    } else if (averageQualityScore >= 5) {
      averageQsResult.score = 60;
      averageQsResult.recommendations.push({
        text: "Improve your average quality score from " + averageQualityScore.toFixed(1) + " to at least " + CONFIG.bestPractices.minQualityScore + " to reduce CPCs and improve ad position.",
        impact: 0.8
      });
    } else {
      averageQsResult.score = 40;
      averageQsResult.recommendations.push({
        text: "Your average quality score of " + averageQualityScore.toFixed(1) + " is significantly below the recommended minimum of " + CONFIG.bestPractices.minQualityScore + ". Focus on improving ad relevance and landing page experience.",
        impact: 0.9
      });
    }
    
    results.criteria.averageQualityScore = averageQsResult;
    
    // 2. Evaluate quality score distribution
    const distributionResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get QS distribution data
    const keywordsByQualityScore = accountData.structure.keywordsByQualityScore || {};
    const totalKeywords = accountData.structure.keywordCount || 0;
    
    // Calculate percentages for each QS bucket
    const lowQsCount = (keywordsByQualityScore[1] || 0) + 
                    (keywordsByQualityScore[2] || 0) + 
                    (keywordsByQualityScore[3] || 0) + 
                    (keywordsByQualityScore[4] || 0);
    
    const mediumQsCount = (keywordsByQualityScore[5] || 0) + 
                       (keywordsByQualityScore[6] || 0);
    
    const highQsCount = (keywordsByQualityScore[7] || 0) + 
                     (keywordsByQualityScore[8] || 0) + 
                     (keywordsByQualityScore[9] || 0) + 
                     (keywordsByQualityScore[10] || 0);
    
    const lowQsPercentage = totalKeywords > 0 ? lowQsCount / totalKeywords : 0;
    const mediumQsPercentage = totalKeywords > 0 ? mediumQsCount / totalKeywords : 0;
    const highQsPercentage = totalKeywords > 0 ? highQsCount / totalKeywords : 0;
    
    distributionResult.details.lowQsPercentage = lowQsPercentage * 100;
    distributionResult.details.mediumQsPercentage = mediumQsPercentage * 100;
    distributionResult.details.highQsPercentage = highQsPercentage * 100;
    
    // Score based on QS distribution
    if (highQsPercentage >= 0.7 && lowQsPercentage <= 0.1) {
      distributionResult.score = 90;
    } else if (highQsPercentage >= 0.5 && lowQsPercentage <= 0.2) {
      distributionResult.score = 75;
      distributionResult.recommendations.push({
        text: "Improve the " + Math.round(lowQsPercentage * 100) + "% of keywords with low quality scores (1-4) by making ads and landing pages more relevant.",
        impact: 0.7
      });
    } else if (highQsPercentage >= 0.3) {
      distributionResult.score = 60;
      distributionResult.recommendations.push({
        text: "Only " + Math.round(highQsPercentage * 100) + "% of your keywords have high quality scores (7-10). Focus on improving your low-performing keywords.",
        impact: 0.8
      });
    } else {
      distributionResult.score = 40;
      distributionResult.recommendations.push({
        text: "Your quality score distribution is poor. " + Math.round(lowQsPercentage * 100) + "% of keywords have low quality scores (1-4). Consider pausing the worst performers and restructuring ad groups.",
        impact: 0.9
      });
    }
    
    results.criteria.qualityScoreDistribution = distributionResult;
    
    // 3. Evaluate ad relevance component
    const adRelevanceResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get ad relevance data
    const goodAdRelevancePercentage = accountData.qualityScore.goodAdRelevancePercentage || 0;
    const poorAdRelevancePercentage = accountData.qualityScore.poorAdRelevancePercentage || 0;
    
    adRelevanceResult.details.goodAdRelevancePercentage = goodAdRelevancePercentage * 100;
    adRelevanceResult.details.poorAdRelevancePercentage = poorAdRelevancePercentage * 100;
    
    // Score based on ad relevance
    if (goodAdRelevancePercentage >= 0.7 && poorAdRelevancePercentage <= 0.1) {
      adRelevanceResult.score = 90;
    } else if (goodAdRelevancePercentage >= 0.5 && poorAdRelevancePercentage <= 0.2) {
      adRelevanceResult.score = 75;
      adRelevanceResult.recommendations.push({
        text: "Improve ad relevance for the " + Math.round(poorAdRelevancePercentage * 100) + "% of keywords with below average ad relevance.",
        impact: 0.7
      });
    } else {
      adRelevanceResult.score = 50;
      adRelevanceResult.recommendations.push({
        text: "Only " + Math.round(goodAdRelevancePercentage * 100) + "% of your keywords have above average ad relevance. Create more specific ads that include your keywords.",
        impact: 0.8
      });
    }
    
    results.criteria.adRelevance = adRelevanceResult;
    
    // 4. Evaluate landing page experience
    const landingPageResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get landing page data
    const goodLandingPagePercentage = accountData.qualityScore.goodLandingPagePercentage || 0;
    const poorLandingPagePercentage = accountData.qualityScore.poorLandingPagePercentage || 0;
    const landingPageSpeed = accountData.qualityScore.landingPageSpeed || 0;
    
    landingPageResult.details.goodLandingPagePercentage = goodLandingPagePercentage * 100;
    landingPageResult.details.poorLandingPagePercentage = poorLandingPagePercentage * 100;
    landingPageResult.details.landingPageSpeed = landingPageSpeed;
    
    // Score based on landing page experience
    if (goodLandingPagePercentage >= 0.7 && poorLandingPagePercentage <= 0.1 && landingPageSpeed >= 80) {
      landingPageResult.score = 90;
    } else if (goodLandingPagePercentage >= 0.5 && landingPageSpeed >= 70) {
      landingPageResult.score = 75;
      
      if (poorLandingPagePercentage > 0.1) {
        landingPageResult.recommendations.push({
          text: "Improve landing pages for the " + Math.round(poorLandingPagePercentage * 100) + "% of keywords with below average landing page experience.",
          impact: 0.7
        });
      }
    } else {
      landingPageResult.score = 50;
      
      if (goodLandingPagePercentage < 0.3) {
        landingPageResult.recommendations.push({
          text: "Improve landing page experience by ensuring your landing pages are relevant to your keywords and ads. Only " + Math.round(goodLandingPagePercentage * 100) + "% of your keywords have above average landing page experience.",
          impact: 0.8
        });
      }
      
      if (landingPageSpeed < 70 && landingPageSpeed > 0) {
        landingPageResult.recommendations.push({
          text: `Improve landing page load speed (currently ${landingPageSpeed}/100). Faster pages provide better user experience and can significantly improve conversion rates.`,
          impact: 0.7
        });
      }
    }
    
    results.criteria.landingPageExperience = landingPageResult;
    
    // Calculate overall category score (weighted average of criteria scores)
    const categoryInfo = EVALUATION_CATEGORIES.find(c => c.name === "Quality Score");
    let weightedScoreSum = 0;
    let weightSum = 0;
    
    categoryInfo.criteria.forEach(criterion => {
      const criterionKey = criterion.name.toLowerCase().replace(/\s+|&/g, '').replace(/[^a-z0-9]/g, '');
      const criterionResult = results.criteria[criterionKey];
      
      if (criterionResult) {
        weightedScoreSum += criterionResult.score * criterion.weight;
        weightSum += criterion.weight;
        
        // Add recommendations to the category level
        criterionResult.recommendations.forEach(rec => {
          results.recommendations.push(rec);
        });
      }
    });
    
    results.score = weightSum > 0 ? weightedScoreSum / weightSum : 0;
    
    // Assign letter grade
    if (results.score >= CONFIG.gradeThresholds.A) {
      results.letter = 'A';
    } else if (results.score >= CONFIG.gradeThresholds.B) {
      results.letter = 'B';
    } else if (results.score >= CONFIG.gradeThresholds.C) {
      results.letter = 'C';
    } else if (results.score >= CONFIG.gradeThresholds.D) {
      results.letter = 'D';
    } else {
      results.letter = 'F';
    }
    
    return results;
  }
  
  /**
   * Evaluates audience strategy
   * @param {Object} accountData The collected account data
   * @return {Object} Evaluation results for audience strategy
   */
  function evaluateAudienceStrategy(accountData) {
    Logger.log("Evaluating audience strategy...");
    
    // Initialize results object
    const results = {
      score: 0,
      letter: '',
      criteria: {},
      recommendations: []
    };
    
    // 1. Evaluate remarketing implementation
    const remarketingResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get remarketing data
    const remarketingListCount = accountData.audiences.remarketingListCount || 0;
    const activeRemarketingCampaigns = accountData.audiences.activeRemarketingCampaigns || 0;
    const campaignCount = accountData.structure.campaignCount || 0;
    
    // Calculate remarketing coverage
    const remarketingCampaignPercentage = campaignCount > 0 ? 
      activeRemarketingCampaigns / campaignCount : 0;
    
    remarketingResult.details.remarketingListCount = remarketingListCount;
    remarketingResult.details.activeRemarketingCampaigns = activeRemarketingCampaigns;
    remarketingResult.details.remarketingCampaignPercentage = remarketingCampaignPercentage * 100;
    
    // Score based on remarketing implementation
    if (remarketingListCount >= 3 && remarketingCampaignPercentage >= 0.5) {
      remarketingResult.score = 90;
    } else if (remarketingListCount >= 1 && activeRemarketingCampaigns >= 1) {
      remarketingResult.score = 70;
      
      if (remarketingListCount < 3) {
        remarketingResult.recommendations.push({
          text: "Create more remarketing lists to target different user behaviors and funnel stages. You currently have " + remarketingListCount + " list(s).",
          impact: 0.7
        });
      }
      
      if (remarketingCampaignPercentage < 0.5) {
        remarketingResult.recommendations.push({
          text: "Expand remarketing to more campaigns. Only " + Math.round(remarketingCampaignPercentage * 100) + "% of your campaigns use remarketing audiences.",
          impact: 0.6
        });
      }
    } else if (remarketingListCount >= 1) {
      remarketingResult.score = 50;
      remarketingResult.recommendations.push({
        text: "Activate your remarketing lists in campaigns. You have lists created but they're not being used effectively.",
        impact: 0.8
      });
    } else {
      remarketingResult.score = 20;
      remarketingResult.recommendations.push({
        text: "Set up remarketing to re-engage past website visitors. This is a fundamental audience strategy missing from your account.",
        impact: 0.9
      });
    }
    
    results.criteria.remarketingImplementation = remarketingResult;
    
    // 2. Evaluate customer list targeting
    const customerListResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get customer list data
    const hasCustomerMatch = accountData.audiences.hasCustomerMatch || false;
    const customerMatchListCount = accountData.audiences.customerMatchListCount || 0;
    
    customerListResult.details.hasCustomerMatch = hasCustomerMatch;
    customerListResult.details.customerMatchListCount = customerMatchListCount;
    
    // Score based on customer list usage
    if (hasCustomerMatch && customerMatchListCount >= 2) {
      customerListResult.score = 90;
    } else if (hasCustomerMatch) {
      customerListResult.score = 70;
      customerListResult.recommendations.push({
        text: "Create more customer match lists to target different customer segments. You currently have only " + customerMatchListCount + " list(s).",
        impact: 0.6
      });
    } else {
      customerListResult.score = 30;
      customerListResult.recommendations.push({
        text: "Implement customer match to target your existing customers and similar audiences. This powerful feature is missing from your account.",
        impact: 0.8
      });
    }
    
    results.criteria.customerListTargeting = customerListResult;
    
    // 3. Evaluate in-market and affinity audiences
    const inMarketResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get in-market and affinity data
    const hasInMarketAudiences = accountData.audiences.hasInMarketAudiences || false;
    const hasAffinityAudiences = accountData.audiences.hasAffinityAudiences || false;
    const inMarketAudienceCount = accountData.audiences.inMarketAudienceCount || 0;
    const affinityAudienceCount = accountData.audiences.affinityAudienceCount || 0;
    
    inMarketResult.details.hasInMarketAudiences = hasInMarketAudiences;
    inMarketResult.details.hasAffinityAudiences = hasAffinityAudiences;
    inMarketResult.details.inMarketAudienceCount = inMarketAudienceCount;
    inMarketResult.details.affinityAudienceCount = affinityAudienceCount;
    
    // Score based on in-market and affinity usage
    if (hasInMarketAudiences && hasAffinityAudiences) {
      inMarketResult.score = 90;
    } else if (hasInMarketAudiences || hasAffinityAudiences) {
      inMarketResult.score = 70;
      
      if (!hasInMarketAudiences) {
        inMarketResult.recommendations.push({
          text: "Add in-market audiences to target users actively researching products or services like yours.",
          impact: 0.7
        });
      }
      
      if (!hasAffinityAudiences) {
        inMarketResult.recommendations.push({
          text: "Implement affinity audiences to reach users based on their long-term interests and habits.",
          impact: 0.6
        });
      }
    } else {
      inMarketResult.score = 40;
      inMarketResult.recommendations.push({
        text: "Start using Google's in-market and affinity audiences to expand your targeting to relevant prospects.",
        impact: 0.8
      });
    }
    
    results.criteria.inMarketAffinityAudiences = inMarketResult;
    
    // 4. Evaluate audience bid adjustments
    const bidAdjustmentResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get audience bid adjustment data
    const hasAudienceBidAdjustments = accountData.bidding.hasAudienceBidAdjustments || false;
    const audienceBidAdjustmentPercentage = accountData.audiences.audienceBidAdjustmentPercentage || 0;
    
    bidAdjustmentResult.details.hasAudienceBidAdjustments = hasAudienceBidAdjustments;
    bidAdjustmentResult.details.audienceBidAdjustmentPercentage = audienceBidAdjustmentPercentage * 100;
    
    // Score based on audience bid adjustments
    if (hasAudienceBidAdjustments && audienceBidAdjustmentPercentage >= 0.7) {
      bidAdjustmentResult.score = 90;
    } else if (hasAudienceBidAdjustments) {
      bidAdjustmentResult.score = 70;
      bidAdjustmentResult.recommendations.push({
        text: "Expand audience bid adjustments to more of your audiences. Currently only " + 
              Math.round(audienceBidAdjustmentPercentage * 100) + "% of your audiences have bid adjustments.",
        impact: 0.6
      });
    } else {
      bidAdjustmentResult.score = 40;
      bidAdjustmentResult.recommendations.push({
        text: "Implement bid adjustments for your audiences to optimize performance based on user behavior and interests.",
        impact: 0.7
      });
    }
    
    results.criteria.audienceBidAdjustments = bidAdjustmentResult;
    
    // Calculate overall category score (weighted average of criteria scores)
    const categoryInfo = EVALUATION_CATEGORIES.find(c => c.name === "Audience Strategy");
    let weightedScoreSum = 0;
    let weightSum = 0;
    
    categoryInfo.criteria.forEach(criterion => {
      const criterionKey = criterion.name.toLowerCase().replace(/\s+|&/g, '').replace(/[^a-z0-9]/g, '');
      const criterionResult = results.criteria[criterionKey];
      
      if (criterionResult) {
        weightedScoreSum += criterionResult.score * criterion.weight;
        weightSum += criterion.weight;
        
        // Add recommendations to the category level
        criterionResult.recommendations.forEach(rec => {
          results.recommendations.push(rec);
        });
      }
    });
    
    results.score = weightSum > 0 ? weightedScoreSum / weightSum : 0;
    
    // Assign letter grade
    if (results.score >= CONFIG.gradeThresholds.A) {
      results.letter = 'A';
    } else if (results.score >= CONFIG.gradeThresholds.B) {
      results.letter = 'B';
    } else if (results.score >= CONFIG.gradeThresholds.C) {
      results.letter = 'C';
    } else if (results.score >= CONFIG.gradeThresholds.D) {
      results.letter = 'D';
    } else {
      results.letter = 'F';
    }
    
    return results;
  }
  
  /**
   * Evaluates landing page optimization
   * @param {Object} accountData The collected account data
   * @return {Object} Evaluation results for landing page optimization
   */
  function evaluateLandingPage(accountData) {
    Logger.log("Evaluating landing page optimization...");
    
    // Initialize results object
    const results = {
      score: 0,
      letter: '',
      criteria: {},
      recommendations: []
    };
    
    // 1. Evaluate landing page relevance
    const relevanceResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get landing page relevance data
    const goodLandingPagePercentage = accountData.qualityScore.goodLandingPagePercentage || 0;
    const poorLandingPagePercentage = accountData.qualityScore.poorLandingPagePercentage || 0;
    
    relevanceResult.details.goodLandingPagePercentage = goodLandingPagePercentage * 100;
    relevanceResult.details.poorLandingPagePercentage = poorLandingPagePercentage * 100;
    
    // Score based on landing page relevance
    if (goodLandingPagePercentage >= 0.7 && poorLandingPagePercentage <= 0.1) {
      relevanceResult.score = 90;
    } else if (goodLandingPagePercentage >= 0.5) {
      relevanceResult.score = 70;
      relevanceResult.recommendations.push({
        text: "Improve landing page relevance for the " + Math.round(poorLandingPagePercentage * 100) + 
              "% of keywords with below average landing page experience.",
        impact: 0.7
      });
    } else {
      relevanceResult.score = 50;
      relevanceResult.recommendations.push({
        text: "Only " + Math.round(goodLandingPagePercentage * 100) + 
              "% of your keywords have above average landing page experience. Create more relevant landing pages that match search intent.",
        impact: 0.8
      });
    }
    
    results.criteria.landingPageRelevance = relevanceResult;
    
    // 2. Evaluate landing page performance
    const performanceResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get landing page performance data
    const landingPageSpeed = accountData.qualityScore.landingPageSpeed || 0;
    const landingPageConversionRate = accountData.landingPage.conversionRate || 0;
    const industryAvgConversionRate = CONFIG.industryBenchmarks.conversionRate || 3.75;
    
    performanceResult.details.landingPageSpeed = landingPageSpeed;
    performanceResult.details.landingPageConversionRate = landingPageConversionRate;
    performanceResult.details.industryAvgConversionRate = industryAvgConversionRate;
    
    // Score based on landing page performance
    if (landingPageSpeed >= 80 && landingPageConversionRate >= industryAvgConversionRate * 1.2) {
      performanceResult.score = 90;
    } else if (landingPageSpeed >= 70 && landingPageConversionRate >= industryAvgConversionRate * 0.8) {
      performanceResult.score = 70;
      
      if (landingPageSpeed < 80) {
        performanceResult.recommendations.push({
          text: "Improve landing page speed (currently " + landingPageSpeed + "/100). Faster pages lead to better user experience and higher conversion rates.",
          impact: 0.7
        });
      }
      
      if (landingPageConversionRate < industryAvgConversionRate) {
        performanceResult.recommendations.push({
          text: "Your landing page conversion rate (" + landingPageConversionRate.toFixed(2) + 
                "%) is below industry average (" + industryAvgConversionRate.toFixed(2) + "%). Test different layouts and calls-to-action.",
          impact: 0.8
        });
      }
    } else {
      performanceResult.score = 50;
      performanceResult.recommendations.push({
        text: "Your landing pages need significant improvement in both speed and conversion rate. Consider a redesign focused on user experience and conversion optimization.",
        impact: 0.9
      });
    }
    
    results.criteria.landingPagePerformance = performanceResult;
    
    // 3. Evaluate mobile optimization
    const mobileResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get mobile optimization data
    const isMobileFriendly = accountData.landingPage.isMobileFriendly || false;
    const mobileConversionRate = accountData.landingPage.mobileConversionRate || 0;
    const desktopConversionRate = accountData.landingPage.desktopConversionRate || 0;
    
    mobileResult.details.isMobileFriendly = isMobileFriendly;
    mobileResult.details.mobileConversionRate = mobileConversionRate;
    mobileResult.details.desktopConversionRate = desktopConversionRate;
    
    // Calculate mobile-to-desktop conversion ratio
    const mobileDesktopRatio = desktopConversionRate > 0 ? mobileConversionRate / desktopConversionRate : 0;
    mobileResult.details.mobileDesktopRatio = mobileDesktopRatio;
    
    // Score based on mobile optimization
    if (isMobileFriendly && mobileDesktopRatio >= 0.9) {
      mobileResult.score = 90;
    } else if (isMobileFriendly && mobileDesktopRatio >= 0.7) {
      mobileResult.score = 70;
      mobileResult.recommendations.push({
        text: "Your mobile conversion rate is " + Math.round(mobileDesktopRatio * 100) + 
              "% of your desktop rate. Improve mobile UX to close this gap.",
        impact: 0.7
      });
    } else if (isMobileFriendly) {
      mobileResult.score = 50;
      mobileResult.recommendations.push({
        text: "Your mobile conversion rate is significantly lower than desktop. Conduct mobile-specific usability testing to identify and fix issues.",
        impact: 0.8
      });
    } else {
      mobileResult.score = 30;
      mobileResult.recommendations.push({
        text: "Your landing pages are not mobile-friendly. Implement responsive design immediately as mobile traffic continues to grow.",
        impact: 0.9
      });
    }
    
    results.criteria.mobileOptimization = mobileResult;
    
    // 4. Evaluate A/B testing
    const testingResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get A/B testing data
    const isABTestingImplemented = accountData.landingPage.isABTestingImplemented || false;
    const abTestCount = accountData.landingPage.abTestCount || 0;
    
    testingResult.details.isABTestingImplemented = isABTestingImplemented;
    testingResult.details.abTestCount = abTestCount;
    
    // Score based on A/B testing
    if (isABTestingImplemented && abTestCount >= 3) {
      testingResult.score = 90;
    } else if (isABTestingImplemented) {
      testingResult.score = 70;
      testingResult.recommendations.push({
        text: "Expand your A/B testing program. You've run only " + abTestCount + 
              " tests. Aim for continuous testing of different page elements.",
        impact: 0.6
      });
    } else {
      testingResult.score = 40;
      testingResult.recommendations.push({
        text: "Implement A/B testing for your landing pages to systematically improve conversion rates over time.",
        impact: 0.8
      });
    }
    
    results.criteria.abTesting = testingResult;
    
    // Calculate overall category score (weighted average of criteria scores)
    const categoryInfo = EVALUATION_CATEGORIES.find(c => c.name === "Landing Page Optimization");
    let weightedScoreSum = 0;
    let weightSum = 0;
    
    categoryInfo.criteria.forEach(criterion => {
      const criterionKey = criterion.name.toLowerCase().replace(/\s+|&/g, '').replace(/[^a-z0-9]/g, '');
      const criterionResult = results.criteria[criterionKey];
      
      if (criterionResult) {
        weightedScoreSum += criterionResult.score * criterion.weight;
        weightSum += criterion.weight;
        
        // Add recommendations to the category level
        criterionResult.recommendations.forEach(rec => {
          results.recommendations.push(rec);
        });
      }
    });
    
    results.score = weightSum > 0 ? weightedScoreSum / weightSum : 0;
    
    // Assign letter grade
    if (results.score >= CONFIG.gradeThresholds.A) {
      results.letter = 'A';
    } else if (results.score >= CONFIG.gradeThresholds.B) {
      results.letter = 'B';
    } else if (results.score >= CONFIG.gradeThresholds.C) {
      results.letter = 'C';
    } else if (results.score >= CONFIG.gradeThresholds.D) {
      results.letter = 'D';
    } else {
      results.letter = 'F';
    }
    
    return results;
  }
  
  /**
   * Evaluates competitive analysis
   * @param {Object} accountData The collected account data
   * @return {Object} Evaluation results for competitive analysis
   */
  function evaluateCompetitiveAnalysis(accountData) {
    Logger.log("Evaluating competitive analysis...");
    
    // Initialize results object
    const results = {
      score: 0,
      letter: '',
      criteria: {},
      recommendations: []
    };
    
    // 1. Evaluate auction insights monitoring
    const auctionInsightsResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get auction insights data
    const hasAuctionInsightsData = accountData.competitive.hasAuctionInsightsData || false;
    const impressionShare = accountData.competitive.impressionShare || 0;
    const topImpressionShare = accountData.competitive.topImpressionShare || 0;
    
    auctionInsightsResult.details.hasAuctionInsightsData = hasAuctionInsightsData;
    auctionInsightsResult.details.impressionShare = impressionShare * 100;
    auctionInsightsResult.details.topImpressionShare = topImpressionShare * 100;
    
    // Score based on auction insights monitoring
    if (hasAuctionInsightsData && impressionShare >= 0.7) {
      auctionInsightsResult.score = 90;
    } else if (hasAuctionInsightsData && impressionShare >= 0.5) {
      auctionInsightsResult.score = 75;
      auctionInsightsResult.recommendations.push({
        text: "Work on improving your impression share from " + Math.round(impressionShare * 100) + 
              "% to at least 70% in your core markets.",
        impact: 0.7
      });
    } else if (hasAuctionInsightsData) {
      auctionInsightsResult.score = 60;
      auctionInsightsResult.recommendations.push({
        text: "Your impression share is low at " + Math.round(impressionShare * 100) + 
              "%. Increase budgets or improve quality scores to compete more effectively.",
        impact: 0.8
      });
    } else {
      auctionInsightsResult.score = 30;
      auctionInsightsResult.recommendations.push({
        text: "Start monitoring auction insights reports to understand your competitive position in the ad auction.",
        impact: 0.9
      });
    }
    
    results.criteria.auctionInsightsMonitoring = auctionInsightsResult;
    
    // 2. Evaluate competitor keyword targeting
    const competitorKeywordResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get competitor keyword data
    const hasCompetitorCampaigns = accountData.competitive.hasCompetitorCampaigns || false;
    const competitorKeywordCount = accountData.competitive.competitorKeywordCount || 0;
    
    competitorKeywordResult.details.hasCompetitorCampaigns = hasCompetitorCampaigns;
    competitorKeywordResult.details.competitorKeywordCount = competitorKeywordCount;
    
    // Score based on competitor keyword targeting
    if (hasCompetitorCampaigns && competitorKeywordCount >= 50) {
      competitorKeywordResult.score = 90;
    } else if (hasCompetitorCampaigns) {
      competitorKeywordResult.score = 70;
      competitorKeywordResult.recommendations.push({
        text: "Expand your competitor keyword targeting. You're only targeting " + 
              competitorKeywordCount + " competitor keywords.",
        impact: 0.7
      });
    } else {
      competitorKeywordResult.score = 40;
      competitorKeywordResult.recommendations.push({
        text: "Create campaigns targeting competitor brand terms to capture users comparing solutions.",
        impact: 0.8
      });
    }
    
    results.criteria.competitorKeywordTargeting = competitorKeywordResult;
    
    // 3. Evaluate competitive ad copy analysis
    const adCopyResult = {
      score: 0,
      details: {},
      recommendations: []
    };
    
    // Get competitive ad copy data
    const hasCompetitiveAdCopyAnalysis = accountData.competitive.hasCompetitiveAdCopyAnalysis || false;
    const competitiveMessagingScore = accountData.competitive.competitiveMessagingScore || 0;
    
    adCopyResult.details.hasCompetitiveAdCopyAnalysis = hasCompetitiveAdCopyAnalysis;
    adCopyResult.details.competitiveMessagingScore = competitiveMessagingScore;
    
    // Score based on competitive ad copy analysis
    if (hasCompetitiveAdCopyAnalysis && competitiveMessagingScore >= 80) {
      adCopyResult.score = 90;
    } else if (hasCompetitiveAdCopyAnalysis) {
      adCopyResult.score = 70;
      adCopyResult.recommendations.push({
        text: "Improve your competitive messaging in ad copy to better differentiate from competitors.",
        impact: 0.6
      });
    } else {
      adCopyResult.score = 40;
      adCopyResult.recommendations.push({
        text: "Analyze competitor ads and develop messaging that highlights your unique value propositions.",
        impact: 0.7
      });
    }
    
    results.criteria.competitiveAdCopyAnalysis = adCopyResult;
    
    // Calculate overall category score (weighted average of criteria scores)
    const categoryInfo = EVALUATION_CATEGORIES.find(c => c.name === "Competitive Analysis");
    let weightedScoreSum = 0;
    let weightSum = 0;
    
    categoryInfo.criteria.forEach(criterion => {
      const criterionKey = criterion.name.toLowerCase().replace(/\s+|&/g, '').replace(/[^a-z0-9]/g, '');
      const criterionResult = results.criteria[criterionKey];
      
      if (criterionResult) {
        weightedScoreSum += criterionResult.score * criterion.weight;
        weightSum += criterion.weight;
        
        // Add recommendations to the category level
        criterionResult.recommendations.forEach(rec => {
          results.recommendations.push(rec);
        });
      }
    });
    
    results.score = weightSum > 0 ? weightedScoreSum / weightSum : 0;
    
    // Assign letter grade
    if (results.score >= CONFIG.gradeThresholds.A) {
      results.letter = 'A';
    } else if (results.score >= CONFIG.gradeThresholds.B) {
      results.letter = 'B';
    } else if (results.score >= CONFIG.gradeThresholds.C) {
      results.letter = 'C';
    } else if (results.score >= CONFIG.gradeThresholds.D) {
      results.letter = 'D';
    } else {
      results.letter = 'F';
    }
    
    return results;
  }
  
  /**
   * Collects account structure data
   * @param {Object} accountData The account data object to populate
   */
  function collectAccountStructure(accountData) {
    Logger.log("Collecting account structure data...");
    
    // Get campaign count
    const campaignIterator = AdsApp.campaigns()
      .withCondition("Status IN ['ENABLED', 'PAUSED']")
      .get();
    accountData.structure.campaignCount = campaignIterator.totalNumEntities();
    
    // Get ad group count
    const adGroupIterator = AdsApp.adGroups()
      .withCondition("Status IN ['ENABLED', 'PAUSED']")
      .get();
    accountData.structure.adGroupCount = adGroupIterator.totalNumEntities();
    
    // Get keyword count and quality score data
    const keywordIterator = AdsApp.keywords()
      .withCondition("Status IN ['ENABLED', 'PAUSED']")
      .get();
    
    let keywordCount = 0;
    let qualityScoreSum = 0;
    let keywordsByQualityScore = {};
    let adRelevanceDistribution = {
      'ABOVE_AVERAGE': 0,
      'AVERAGE': 0,
      'BELOW_AVERAGE': 0
    };
    let expectedCtrDistribution = {
      'ABOVE_AVERAGE': 0,
      'AVERAGE': 0,
      'BELOW_AVERAGE': 0
    };
    let landingPageDistribution = {
      'ABOVE_AVERAGE': 0,
      'AVERAGE': 0,
      'BELOW_AVERAGE': 0
    };
    
    while (keywordIterator.hasNext()) {
      const keyword = keywordIterator.next();
      keywordCount++;
      
      // Quality Score
      const qualityScore = keyword.getQualityScore();
      if (qualityScore > 0) { // Only count valid quality scores
        qualityScoreSum += qualityScore;
        
        // Count keywords by quality score
        if (!keywordsByQualityScore[qualityScore]) {
          keywordsByQualityScore[qualityScore] = 0;
        }
        keywordsByQualityScore[qualityScore]++;
        
        // Count QS components
        const adRelevance = keyword.getAdRelevanceStatus();
        if (adRelevance) {
          adRelevanceDistribution[adRelevance]++;
        }
        
        const expectedCtr = keyword.getExpectedCtrStatus();
        if (expectedCtr) {
          expectedCtrDistribution[expectedCtr]++;
        }
        
        const landingPageExp = keyword.getLandingPageExperienceStatus();
        if (landingPageExp) {
          landingPageDistribution[landingPageExp]++;
        }
      }
    }
    
    // Calculate averages
    accountData.structure.keywordCount = keywordCount;
    accountData.structure.averageKeywordsPerAdGroup = accountData.structure.adGroupCount > 0 ? 
      keywordCount / accountData.structure.adGroupCount : 0;
    accountData.structure.averageAdGroupsPerCampaign = accountData.structure.campaignCount > 0 ? 
      accountData.structure.adGroupCount / accountData.structure.campaignCount : 0;
    accountData.structure.averageQualityScore = keywordCount > 0 ? 
      qualityScoreSum / keywordCount : 0;
    
    // Store distributions
    accountData.structure.keywordsByQualityScore = keywordsByQualityScore;
    accountData.structure.adRelevanceDistribution = adRelevanceDistribution;
    accountData.structure.expectedCtrDistribution = expectedCtrDistribution;
    accountData.structure.landingPageDistribution = landingPageDistribution;
  }
  
  /**
   * Collects performance data
   * @param {Object} accountData The account data object to populate
   * @param {Object} dateRange Date range for data collection
   */
  function collectPerformanceData(accountData, dateRange) {
    Logger.log("Collecting performance data...");
    
    const query = `SELECT Impressions, Clicks, Cost, Conversions ` +
                  `FROM ACCOUNT_PERFORMANCE_REPORT ` +
                  `WHERE Impressions > 0 AND Date DURING '${dateRange.start}', '${dateRange.end}'`;
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    let metrics = {
      impressions: 0,
      clicks: 0,
      cost: 0,
      conversions: 0
    };
    
    while (rows.hasNext()) {
      const row = rows.next();
      metrics.impressions += parseInt(row['Impressions']);
      metrics.clicks += parseInt(row['Clicks']);
      metrics.cost += parseFloat(row['Cost']);
      metrics.conversions += parseInt(row['Conversions']);
    }
    
    return metrics;
  }
  
  /**
   * Collects campaign data
   * @param {Object} accountData The account data object to populate
   * @param {Object} dateRange Date range for data collection
   */
  function collectCampaignData(accountData, dateRange) {
    Logger.log("Collecting campaign data...");
    
    // Query for campaign performance metrics
    const query = `SELECT CampaignName, CampaignStatus, Impressions, Clicks, Cost, ` +
                  `Conversions, ConversionValue ` +
                  `FROM CAMPAIGN_PERFORMANCE_REPORT ` +
                  `WHERE Date BETWEEN '${dateRange.start}' AND '${dateRange.end}'`;
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    // Process campaign data
    let campaigns = [];
    let totalBudgetLost = 0;
    let totalRankLost = 0;
    let campaignsWithImpressionShareData = 0;
    
    while (rows.hasNext()) {
      const row = rows.next();
      
      // Extract campaign data
      const campaign = {
        id: row['CampaignId'],
        name: row['CampaignName'],
        status: row['CampaignStatus'],
        type: row['AdvertisingChannelType'],
        impressions: parseInt(row['Impressions'], 10) || 0,
        clicks: parseInt(row['Clicks'], 10) || 0,
        cost: parseFloat(row['Cost'].replace(/,/g, '')) || 0,
        conversions: parseFloat(row['Conversions']) || 0,
        impressionShare: parseFloat(row['SearchImpressionShare']) || 0,
        topImpressionShare: parseFloat(row['SearchTopImpressionShare']) || 0,
        absoluteTopImpressionShare: parseFloat(row['SearchAbsoluteTopImpressionShare']) || 0,
        budgetLostImpressionShare: parseFloat(row['SearchBudgetLostImpressionShare']) || 0,
        rankLostImpressionShare: parseFloat(row['SearchRankLostImpressionShare']) || 0
      };
      
      // Add to campaigns array
      campaigns.push(campaign);
      
      // Track impression share metrics
      if (campaign.impressionShare > 0) {
        campaignsWithImpressionShareData++;
        totalBudgetLost += campaign.budgetLostImpressionShare;
        totalRankLost += campaign.rankLostImpressionShare;
      }
    }
    
    // Calculate average impression share metrics
    const avgBudgetLost = campaignsWithImpressionShareData > 0 ? 
      totalBudgetLost / campaignsWithImpressionShareData : 0;
    const avgRankLost = campaignsWithImpressionShareData > 0 ? 
      totalRankLost / campaignsWithImpressionShareData : 0;
    
    // Store campaign data
    accountData.campaigns = campaigns;
    accountData.bidding.budgetLost = avgBudgetLost;
    accountData.bidding.rankLost = avgRankLost;
  }
  
  /**
   * Collects bidding data
   * @param {Object} accountData The account data object to populate
   */
  function collectBiddingData(accountData) {
    Logger.log("Collecting bidding data...");
    
    // Query for campaign bidding strategies
    const query = "SELECT CampaignId, CampaignName, BiddingStrategyType, " +
      "EnhancedCpcEnabled, TargetCpa, TargetRoas " +
      "FROM CAMPAIGN_PERFORMANCE_REPORT";
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    // Initialize counters
    let strategies = {};
    let smartBiddingCount = 0;
    let totalCampaigns = 0;
    
    // Process bidding data
    while (rows.hasNext()) {
      const row = rows.next();
      totalCampaigns++;
      
      // Extract bidding strategy
      const biddingStrategyType = row['BiddingStrategyType'];
      const enhancedCpcEnabled = row['EnhancedCpcEnabled'] === 'true';
      
      // Count strategies
      if (!strategies[biddingStrategyType]) {
        strategies[biddingStrategyType] = 0;
      }
      strategies[biddingStrategyType]++;
      
      // Count smart bidding campaigns
      if (biddingStrategyType === 'TARGET_CPA' || 
          biddingStrategyType === 'TARGET_ROAS' || 
          biddingStrategyType === 'MAXIMIZE_CONVERSIONS' || 
          biddingStrategyType === 'MAXIMIZE_CONVERSION_VALUE' ||
          enhancedCpcEnabled) {
        smartBiddingCount++;
      }
    }
    
    // Check for bid adjustments
    let hasDeviceBidAdjustments = false;
    let hasLocationBidAdjustments = false;
    let hasScheduleBidAdjustments = false;
    
    // Check device bid adjustments
    const deviceBidAdjustmentQuery = "SELECT CampaignId, Device, DeviceBidModifier " +
      "FROM CAMPAIGN_CRITERIA_REPORT " +
      "WHERE CriterionType = 'DEVICE' AND DeviceBidModifier != 0.0";
    
    const deviceReport = AdsApp.report(deviceBidAdjustmentQuery);
    const deviceRows = deviceReport.rows();
    hasDeviceBidAdjustments = deviceRows.hasNext();
    
    // Check location bid adjustments
    const locationBidAdjustmentQuery = "SELECT CampaignId, Id, BidModifier " +
      "FROM CAMPAIGN_LOCATION_TARGET_REPORT " +
      "WHERE BidModifier != 1.0";
    
    const locationReport = AdsApp.report(locationBidAdjustmentQuery);
    const locationRows = locationReport.rows();
    hasLocationBidAdjustments = locationRows.hasNext();
    
    // Check ad schedule bid adjustments
    const scheduleBidAdjustmentQuery = "SELECT CampaignId, AdSchedule, BidModifier " +
      "FROM CAMPAIGN_AD_SCHEDULE_TARGET_REPORT " +
      "WHERE BidModifier != 1.0";
    
    const scheduleReport = AdsApp.report(scheduleBidAdjustmentQuery);
    const scheduleRows = scheduleReport.rows();
    hasScheduleBidAdjustments = scheduleRows.hasNext();
    
    // Calculate smart bidding percentage
    const smartBiddingPercentage = totalCampaigns > 0 ? 
      smartBiddingCount / totalCampaigns : 0;
    
    // Populate bidding data
    accountData.bidding.strategies = strategies;
    accountData.bidding.smartBiddingPercentage = smartBiddingPercentage;
    accountData.bidding.hasDeviceBidAdjustments = hasDeviceBidAdjustments;
    accountData.bidding.hasLocationBidAdjustments = hasLocationBidAdjustments;
    accountData.bidding.hasScheduleBidAdjustments = hasScheduleBidAdjustments;
  }
  
  /**
   * Collects ad data
   * @param {Object} accountData The account data object to populate
   * @param {Object} dateRange Date range for data collection
   */
  function collectAdData(accountData, dateRange) {
    Logger.log("Collecting ad data...");
    
    // Query for ad performance metrics
    const query = `SELECT Id, HeadlinePart1, HeadlinePart2, HeadlinePart3, ` +
                  `Description1, Description2, Impressions, Clicks, Cost, Conversions ` +
                  `FROM AD_PERFORMANCE_REPORT ` +
                  `WHERE Date BETWEEN '${dateRange.start}' AND '${dateRange.end}'`;
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    // Initialize counters
    let totalAds = 0;
    let rsaCount = 0;
    let disapprovedCount = 0;
    let limitedByPolicyCount = 0;
    let totalHeadlines = 0;
    let totalDescriptions = 0;
    let rsaWithHeadlines = 0;
    
    // Track ads per ad group
    const adGroupAdsMap = new Map();
    
    // Process ad data
    while (rows.hasNext()) {
      const row = rows.next();
      totalAds++;
      
      // Extract ad data
      const adType = row['AdType'];
      const adGroupId = row['AdGroupId'];
      const status = row['Status'];
      const approvalStatus = row['PolicySummaryApprovalStatus'];
      
      // Count ads by ad group
      if (!adGroupAdsMap.has(adGroupId)) {
        adGroupAdsMap.set(adGroupId, 0);
      }
      adGroupAdsMap.set(adGroupId, adGroupAdsMap.get(adGroupId) + 1);
      
      // Count RSAs
      if (adType === 'RESPONSIVE_SEARCH_AD') {
        rsaCount++;
        
        // Count headlines and descriptions for RSAs
        let headlineCount = 0;
        if (row['HeadlinePart1'] && row['HeadlinePart1'].trim() !== '') headlineCount++;
        if (row['HeadlinePart2'] && row['HeadlinePart2'].trim() !== '') headlineCount++;
        if (row['HeadlinePart3'] && row['HeadlinePart3'].trim() !== '') headlineCount++;
        
        let descriptionCount = 0;
        if (row['Description1'] && row['Description1'].trim() !== '') descriptionCount++;
        if (row['Description2'] && row['Description2'].trim() !== '') descriptionCount++;
        
        totalHeadlines += headlineCount;
        totalDescriptions += descriptionCount;
        rsaWithHeadlines++;
      }
      
      // Count disapproved and limited ads
      if (approvalStatus === 'DISAPPROVED') {
        disapprovedCount++;
      } else if (approvalStatus === 'LIMITED') {
        limitedByPolicyCount++;
      }
    }
    
    // Calculate metrics
    const rsaPercentage = totalAds > 0 ? rsaCount / totalAds : 0;
    const disapprovedPercentage = totalAds > 0 ? disapprovedCount / totalAds : 0;
    const limitedByPolicyPercentage = totalAds > 0 ? limitedByPolicyCount / totalAds : 0;
    
    // Calculate average headlines and descriptions per RSA
    const averageHeadlinesPerRsa = rsaWithHeadlines > 0 ? totalHeadlines / rsaWithHeadlines : 0;
    const averageDescriptionsPerRsa = rsaWithHeadlines > 0 ? totalDescriptions / rsaWithHeadlines : 0;
    
    // Calculate average ads per ad group
    let adGroupsWithAds = 0;
    let singleAdAdGroups = 0;
    
    adGroupAdsMap.forEach((adCount, adGroupId) => {
      adGroupsWithAds++;
      if (adCount === 1) {
        singleAdAdGroups++;
      }
    });
    
    const averageAdsPerAdGroup = adGroupsWithAds > 0 ? totalAds / adGroupsWithAds : 0;
    const singleAdAdGroupPercentage = adGroupsWithAds > 0 ? singleAdAdGroups / adGroupsWithAds : 0;
    
    // Populate ad data
    accountData.ads.rsaPercentage = rsaPercentage;
    accountData.ads.averageAdsPerAdGroup = averageAdsPerAdGroup;
    accountData.ads.singleAdAdGroupPercentage = singleAdAdGroupPercentage;
    accountData.ads.averageHeadlinesPerRsa = averageHeadlinesPerRsa;
    accountData.ads.averageDescriptionsPerRsa = averageDescriptionsPerRsa;
    accountData.ads.disapprovedPercentage = disapprovedPercentage;
    accountData.ads.limitedByPolicyPercentage = limitedByPolicyPercentage;
  }
  
  /**
   * Collects extension data
   * @param {Object} accountData The account data object to populate
   * @param {Object} dateRange Date range for data collection
   */
  function collectExtensionData(accountData, dateRange) {
    Logger.log("Collecting extension data...");
    
    // Count extension types
    let extensionTypes = 0;
    
    // Check for sitelink extensions
    const sitelinkIterator = AdsApp.extensions().sitelinks().get();
    if (sitelinkIterator.totalNumEntities() > 0) {
      extensionTypes++;
    }
    
    // Check for callout extensions
    const calloutIterator = AdsApp.extensions().callouts().get();
    if (calloutIterator.totalNumEntities() > 0) {
      extensionTypes++;
    }
    
    // Check for structured snippet extensions
    const snippetIterator = AdsApp.extensions().snippets().get();
    if (snippetIterator.totalNumEntities() > 0) {
      extensionTypes++;
    }
    
    // Check for call extensions
    const callIterator = AdsApp.extensions().phoneNumbers().get();
    if (callIterator.totalNumEntities() > 0) {
      extensionTypes++;
    }
    
    // Check for price extensions
    const priceIterator = AdsApp.extensions().prices().get();
    if (priceIterator.totalNumEntities() > 0) {
      extensionTypes++;
    }
    
    // Check for app extensions
    const appIterator = AdsApp.extensions().mobileApps().get();
    if (appIterator.totalNumEntities() > 0) {
      extensionTypes++;
    }
    
    // Check for lead form extensions
    const leadFormIterator = AdsApp.extensions().leadForms().get();
    if (leadFormIterator.totalNumEntities() > 0) {
      extensionTypes++;
    }
    
    // Get extension impressions data
    const query = "SELECT Impressions, ClickType " +
      "FROM CLICK_PERFORMANCE_REPORT " +
      `WHERE Impressions > 0 AND Date BETWEEN ${dateRange.start} AND ${dateRange.end}`;
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    let totalImpressions = accountData.performance.impressions;
    let impressionsWithExtensions = 0;
    
    while (rows.hasNext()) {
      const row = rows.next();
      const clickType = row['ClickType'];
      const impressions = parseInt(row['Impressions'], 10);
      
      // Count impressions with extension clicks
      if (clickType !== 'headline' && !isNaN(impressions)) {
        impressionsWithExtensions += impressions;
      }
    }
    
    // Populate extension data
    accountData.extensions.totalCount = extensionTypes;
    accountData.extensions.impressionWithExtensions = impressionsWithExtensions;
  }
  
  /**
   * Collects quality score data
   * @param {Object} accountData The account data object to populate
   */
  function collectQualityScoreData(accountData) {
    Logger.log("Collecting quality score data...");
    
    // Quality score data is already collected in collectAccountStructure
    // This function extracts and processes that data further
    
    // Get quality score distribution from the report
    const query = "SELECT QualityScore, Impressions " +
      "FROM KEYWORDS_PERFORMANCE_REPORT " +
      "WHERE Status IN ['ENABLED', 'PAUSED'] " +
      "AND QualityScore > 0";
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    let totalKeywords = 0;
    let weightedQualityScoreSum = 0;
    let qualityScoreDistribution = {};
    
    while (rows.hasNext()) {
      const row = rows.next();
      const qualityScore = parseInt(row['QualityScore'], 10);
      const impressions = parseInt(row['Impressions'], 10) || 1; // Use 1 if no impressions
      
      if (!isNaN(qualityScore) && qualityScore > 0) {
        totalKeywords++;
        weightedQualityScoreSum += qualityScore * impressions;
        
        // Count by quality score
        if (!qualityScoreDistribution[qualityScore]) {
          qualityScoreDistribution[qualityScore] = 0;
        }
        qualityScoreDistribution[qualityScore]++;
      }
    }
    
    // Calculate impression-weighted average quality score
    const averageQualityScore = totalKeywords > 0 ? weightedQualityScoreSum / totalKeywords : 0;
    
    // Calculate percentage of keywords with low quality score (below 5)
    let lowQualityScoreCount = 0;
    for (let i = 1; i <= 4; i++) {
      lowQualityScoreCount += qualityScoreDistribution[i] || 0;
    }
    const lowQualityScorePercentage = totalKeywords > 0 ? lowQualityScoreCount / totalKeywords : 0;
    
    // Calculate good ad relevance percentage
    const totalWithAdRelevance = 
      (accountData.qualityScore.adRelevanceDistribution['ABOVE_AVERAGE'] || 0) +
      (accountData.qualityScore.adRelevanceDistribution['AVERAGE'] || 0) +
      (accountData.qualityScore.adRelevanceDistribution['BELOW_AVERAGE'] || 0);
    
    const goodAdRelevancePercentage = totalWithAdRelevance > 0 ? 
      accountData.qualityScore.adRelevanceDistribution['ABOVE_AVERAGE'] / totalWithAdRelevance : 0;
    
    // Calculate good expected CTR percentage
    const totalWithExpectedCtr = 
      (accountData.qualityScore.expectedCtrDistribution['ABOVE_AVERAGE'] || 0) +
      (accountData.qualityScore.expectedCtrDistribution['AVERAGE'] || 0) +
      (accountData.qualityScore.expectedCtrDistribution['BELOW_AVERAGE'] || 0);
    
    const goodExpectedCtrPercentage = totalWithExpectedCtr > 0 ? 
      accountData.qualityScore.expectedCtrDistribution['ABOVE_AVERAGE'] / totalWithExpectedCtr : 0;
    
    // Calculate good landing page percentage
    const totalWithLandingPage = 
      (accountData.qualityScore.landingPageDistribution['ABOVE_AVERAGE'] || 0) +
      (accountData.qualityScore.landingPageDistribution['AVERAGE'] || 0) +
      (accountData.qualityScore.landingPageDistribution['BELOW_AVERAGE'] || 0);
    
    const goodLandingPagePercentage = totalWithLandingPage > 0 ? 
      accountData.qualityScore.landingPageDistribution['ABOVE_AVERAGE'] / totalWithLandingPage : 0;
    
    const poorLandingPagePercentage = totalWithLandingPage > 0 ? 
      accountData.qualityScore.landingPageDistribution['BELOW_AVERAGE'] / totalWithLandingPage : 0;
    
    // Populate quality score data
    accountData.qualityScore.average = averageQualityScore;
    accountData.qualityScore.distribution = qualityScoreDistribution;
    accountData.qualityScore.lowQualityScorePercentage = lowQualityScorePercentage;
    accountData.qualityScore.goodAdRelevancePercentage = goodAdRelevancePercentage;
    accountData.qualityScore.goodExpectedCtrPercentage = goodExpectedCtrPercentage;
    accountData.qualityScore.goodLandingPagePercentage = goodLandingPagePercentage;
    accountData.qualityScore.poorLandingPagePercentage = poorLandingPagePercentage;
  }
  
  /**
   * Collects conversion tracking data
   * @param {Object} accountData The account data object to populate
   */
  function collectConversionTrackingData(accountData) {
    Logger.log("Collecting conversion tracking data...");
    
    // Query for conversion actions
    const query = "SELECT ConversionActionName, ConversionActionCategory, " +
      "IncludeInConversionsMetric, ValueSettingStatus, AttributionModelType " +
      "FROM CONVERSION_ACTION_REPORT";
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    // Initialize counters
    let conversionCount = 0;
    let valueTrackingCount = 0;
    let hasPhoneCallTracking = false;
    let hasImportedConversions = false;
    
    // Process conversion actions
    while (rows.hasNext()) {
      const row = rows.next();
      
      // Count active conversion actions
      if (row['IncludeInConversionsMetric'] === 'true') {
        conversionCount++;
        
        // Check for value tracking
        if (row['ValueSettingStatus'] === 'ACTIVE') {
          valueTrackingCount++;
        }
        
        // Check for phone call tracking
        const category = row['ConversionActionCategory'];
        if (category === 'PHONE_CALL_LEAD' || category === 'PHONE_CALL_CONVERSION') {
          hasPhoneCallTracking = true;
        }
        
        // Check for imported conversions
        if (category === 'UPLOAD' || category === 'IMPORT') {
          hasImportedConversions = true;
        }
      }
    }
    
    // Populate conversion tracking data
    accountData.conversionTracking.count = conversionCount;
    accountData.conversionTracking.valueTrackingCount = valueTrackingCount;
    accountData.conversionTracking.hasPhoneCallTracking = hasPhoneCallTracking;
    accountData.conversionTracking.hasImportedConversions = hasImportedConversions;
  }
  
  /**
   * Sends an error notification email
   * @param {Error} error The error that occurred
   */
  function sendErrorNotification(error) {
    Logger.log("Sending error notification...");
    
    const accountName = AdsApp.currentAccount().getName();
    const accountId = AdsApp.currentAccount().getCustomerId();
    
    // Create email subject
    const subject = "ERROR: Google Ads Account Grader - " + accountName + " (" + accountId + ")";
    
    // Create email body
    let body = "<h2>Error Running Google Ads Account Grader</h2>";
    body += "<p><strong>Account:</strong> " + accountName + " (" + accountId + ")</p>";
    body += "<p><strong>Date:</strong> " + Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), "yyyy-MM-dd") + "</p>";
    body += "<p><strong>Error Message:</strong> " + error.message + "</p>";
    
    if (error.stack) {
      body += "<p><strong>Stack Trace:</strong></p>";
      body += "<pre>" + error.stack + "</pre>";
    }
    
    // Send the email
    MailApp.sendEmail({
      to: CONFIG.email.errorRecipients ? CONFIG.email.errorRecipients.join(",") : CONFIG.email.emailAddress,
      subject: subject,
      htmlBody: body
    });
  }