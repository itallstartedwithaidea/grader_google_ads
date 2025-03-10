Below is an example **README.md** file (GitHub-friendly) that you can include with your script, giving step-by-step instructions—**with approximate line references**—on how to copy, install, and configure the script in Google Ads. Adjust any line numbers as needed if your script differs.

---

# Google Ads Account Grader - Setup Guide

This document explains how to install and configure the Google Ads Account Grader script in your Google Ads account. Follow these steps carefully to ensure proper setup.

---

## 1. Copy the Script into Google Ads

1. **Log into Google Ads** using the account (or Manager account) where you want to run this script.
2. **Navigate to “Tools & Settings”** in the top menu bar, then under **“Bulk Actions,”** click **“Scripts.”**
3. **Click the blue plus (+) button** to create a new script.
4. **Paste** the entire script (the one in `GoogleAdsAccountGrader.js`, for example) into the editor window.
5. **Press “Save.”**

> **Tip**: If you use a **Manager account**, you might find “Scripts” under **“Setup”** → **“Scripts”** in the left menu instead.

---

## 2. Update the `CONFIG` Settings

Near the **top** of the script (around **lines 35–140**), you’ll see:

```js
// Configuration
const CONFIG = {
  // Date range for data collection
  dateRange: {
    // Set to true to use custom date range, false to use lookback period
    useCustomDateRange: true,
    
    // Custom date range (only used if useCustomDateRange is true)
    // Format: YYYYMMDD (e.g., 20220701 for July 1, 2022)
    customStartDate: "20220715", // Example
    customEndDate: "20220915",   // Example
    
    // Lookback period in days (only used if useCustomDateRange is false)
    lookbackDays: 30
  },
  
  // ...
};
```

You **must** edit this object to match your needs:

### 2.1 Date Range

- **Line ~42:**  
  `useCustomDateRange: true,`  
  - Set this to `true` if you want **fixed start/end dates**.
  - Set this to `false` if you want a **rolling lookback** (like last 30 days).

- **Lines ~46–47:**  
  ```js
  customStartDate: "20220715",
  customEndDate: "20220915",
  ```
  - If `useCustomDateRange` is `true`, these two values are used.
  - Must be in **YYYYMMDD** format (e.g. `"20230301"` for March 1, 2023).

- **Line ~50:**  
  `lookbackDays: 30`  
  - If `useCustomDateRange` is `false`, the script looks back this many days from today.

### 2.2 Email Settings

- **Lines ~57–61:**
  ```js
  email: {
    sendEmail: true,
    sendReport: true,
    sendErrorNotifications: true,
    emailAddress: 'your.email@domain.com',
    errorRecipients: ['your.email@domain.com'],
    // ...
  },
  ```
  - **`emailAddress`**: Put your own email address here.
  - **`sendReport`**: If `true`, the script will send a full report email with a link to the spreadsheet.
  - **`sendEmail`**: If `true`, the script can send emails (both reports and any debug messages).
  - **If you only want debugging in the logs,** set `sendEmail: false` (the script won’t email anything).

> **Important**: Make sure your Google Ads script is authorized to send email. When you run it for the first time, you’ll need to grant permission.

---

## 3. Authorize the Script

When you run a Google Ads script for the first time, you’ll see an **“Authorization”** prompt:

1. Click **“Authorize”** in the script editor.
2. Confirm any permissions requested (access to Google Ads data, sending mail, drive/spreadsheets, etc.).
3. Once authorized, your script is ready to run or schedule.

---

## 4. Test & Run the Script

1. **Click “Preview”** to do a test run. You’ll see logs in the console if it runs successfully (no permanent changes).
2. **Check the logs** to confirm it’s fetching data as expected. Watch for any warnings or errors.
3. **Click “Run”** (or “Preview → Run” if you want to skip the preview). This does a full run:
   - It generates the grading report.
   - If `CONFIG.email.sendReport` is `true`, it emails you the results.
   - If `CONFIG.spreadsheet.createNew` is `true`, you’ll get a new Google Sheet link in the logs and/or your email.

---

## 5. (Optional) Schedule the Script

If you’d like the script to run **automatically**:

1. In your Google Ads account, go to **“Scripts”** (Bulk Actions).
2. Find your new script in the list and click **“Schedule.”**
3. Choose the frequency (daily, weekly, monthly, etc.), start time, and date range.
4. Save it. The script will now run on that schedule, emailing fresh reports each time.

---

## 6. Common Updates & Troubleshooting

- **Email Not Sent**: Check that
  - `CONFIG.email.sendEmail` and `CONFIG.email.sendReport` are both `true`.
  - You’ve authorized the script to send email.
  - Your email address (`CONFIG.email.emailAddress`) is correct.
- **Spreadsheet Not Created**: Make sure `CONFIG.spreadsheet.createNew` is `true`. If you want to reuse a single spreadsheet, set `createNew: false` and provide the link in `existingSpreadsheetUrl`.
- **Date Format Errors**: The script expects `YYYYMMDD` (e.g. `"20230228"`). Make sure you aren’t using a dash or slash format.
- **Script Logs**: Use the **Logs** button in the Google Ads script editor to see detailed run-time logs or error messages.

---

Make your changes in place to suit your own account setup.

---

### That’s It!

You now have a working “Google Ads Account Grader” script. Once installed and configured, it will:

- Evaluate your account across multiple categories (bidding, keywords, negatives, ads, etc.).
- Produce an overall grade plus separate category grades.
- Generate an easy-to-read spreadsheet report (and an email if you enabled it).

For further customization, feel free to explore or modify the individual evaluation functions in the script. Happy grading!
