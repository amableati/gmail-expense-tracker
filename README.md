# 📧 Gmail Expense Tracker

An automated personal finance management tool that extracts transaction details from Gmail and syncs them into Google Sheets using a single, modular script file.

## 🚀 Key Features

* **Auto-Fetch:** Parses bank, credit card, and UPI transaction alerts directly from your inbox.
* **Dynamic Categorization:** Uses a dedicated `Mappings` sheet to learn your preferences. The system checks this sheet first, falling back to a hardcoded logic engine if no match is found.
* **Smart Sync:** Automatically cleans duplicate entries and maintains chronological order.
* **Correction Tool:** Select any row, update the category, and click "Save Correction" to train the system for future transactions.

## 🛠 Setup Instructions

### 1. Spreadsheet Setup

1. Create a Google Sheet named **Expenses**.
2. Create two tabs:
* **`Expenses`**: Required columns are: `Date`, `Amount`, `Merchant`, `Payment Mode`, `Category`, `Column Type`, `Source Bank`, and `Raw Text`.
* **`Mappings`**: Requires two columns: `Merchant Keyword` (A) and `Category` (B).



### 2. Apps Script Configuration

1. In your Google Sheet, go to **Extensions > Apps Script**.
2. Delete any existing code and paste your complete script.
3. Save the project.
4. Refresh your spreadsheet—the **"📊 Expense Tracker"** menu will appear in your toolbar.

### 3. Automation (The "Set and Forget" Step)

1. In the Apps Script editor, click the **Triggers** tab (the alarm clock icon on the left sidebar).
2. Click **+ Add Trigger**.
3. Select `fetchAllEmailsRaw` as the function to run.
4. Change the "Select event source" to **Time-driven**.
5. Set it to run **Daily** or **Hourly** based on your preference.

## 💡 How to "Train" the Tracker

To teach the system how to categorize your transactions:

1. Identify a transaction in your `Expenses` sheet that needs a better category.
2. Manually change the **Category** in that row to your preferred value.
3. Select the row(s) you just updated.
4. Click **📊 Expense Tracker > 3. Save Correction**.
5. The system will now remember this mapping for all future transactions matching that merchant!

## ⚙️ How it Works

The script uses a cached mapping approach to ensure high performance:

* **`autoCategorize`**: First checks your `Mappings` sheet. If it finds the merchant, it uses your custom category; otherwise, it uses the default hardcoded rules.
* **`saveCategoryCorrection`**: Updates your `Mappings` sheet and flushes the cache, ensuring the system "learns" your changes instantly without needing a code update.
* **`clearSheetDuplicates`**: A built-in wizard to keep your data clean by removing duplicate entries.

---

### Pro-Tip for Single-File Management

Because everything is in one file, **the order of functions does not matter** to the script execution. However, for your own sanity, keep the "Main" functions (like `onOpen`) at the top and the "Helper" functions (like `cleanText` or `parseMyDate`) at the bottom. This makes the code much easier to navigate when you need to make changes!
