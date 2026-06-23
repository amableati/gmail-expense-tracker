# gmail-expense-tracker

Automated tool to extract transaction details from Gmail and sync them to Google Sheets.

## Features
- **Auto-Fetch:** Parses bank, credit card, and UPI alerts.
- **Categorization:** Intelligent auto-categorization engine.
- **Smart Sync:** Merges duplicate entries and maintains chronological order.
- **Modular Design:** Clean code separated into Core, Parsing, and Utilities.

## Setup
1. Create a Google Sheet named "Expenses".
2. Open **Extensions > Apps Script**.
3. Create three files: `CoreLogic.gs`, `ParsingEngine.gs`, and `Utilities.gs`.
4. Copy the respective functions into each file.
5. Set up a **Time-driven Trigger** in the Triggers tab to automate execution.

## License
MIT
