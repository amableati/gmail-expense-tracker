const ALLOWED_CATEGORIES = "Travel, Shopping, Entertainment, Utility, Groceries, Investment, Food, Self Transfer, Health, Misc";

//  SCHEDULAR
function runAutomatedExpenseWorkflow() {
  // 1. Fetch and process emails
  // Note: We modified fetchAllEmailsRaw to be UI-safe
  fetchAllEmailsRaw(); 
  
  // 2. Clean duplicates automatically
  // Note: You must ensure clearSheetDuplicates is also UI-safe!
  //clearSheetDuplicates();
  
  // 3. Ensure everything is sorted
  sortExpensesByDate();
  
  console.log("Automated workflow completed successfully.");
}

function getUiSafe() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    return null; // Return null if running in background
  }
}


function getLastExtractedDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Expenses");
  
  // Safety check: Does the sheet exist?
  if (!sheet) {
    throw new Error("Sheet named 'Expenses' not found. Please check your sheet name.");
  }
  
  const data = sheet.getDataRange().getValues();
  
  // If sheet is empty (only header or no rows), default to a start date
  if (data.length <= 1) return "2026/04/01"; 

  let maxDate = new Date(0); 
  for (let r = 1; r < data.length; r++) {
    // Ensure Column A contains a valid date
    let cellDate = new Date(data[r][0]);
    if (!isNaN(cellDate.getTime()) && cellDate > maxDate) {
      maxDate = cellDate;
    }
  }
  
  return Utilities.formatDate(maxDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
}

function fetchAllEmailsRaw() {
  const ui = getUiSafe();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expenses");
  if (!sheet) { ui && ui.alert("Error: 'Expenses' sheet not found."); return; }

  const startDateStr = getLastExtractedDate(); 
  let tomorrow = new Date(); tomorrow.setDate(tomorrow.getDate() + 1);
  const endDateStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), "yyyy/MM/dd");
  
  const gmailSearchQuery = `("debited" OR "credited" OR "Transaction alert" OR "SIP Auto Payment" OR "Buy order placed" OR "SmartPay" OR "Amazon Pay" OR "ATM withdrawal") -("reminder" OR "due") after:${startDateStr.replace(/\//g, '-')} before:${endDateStr.replace(/\//g, '-')}`;
  const threads = GmailApp.search(gmailSearchQuery, 0, 500); 

  if (threads.length === 0) { ui && ui.alert("No new transactions found."); return; }

  let rowsAddedCount = 0;
  for (const thread of threads) {
    for (const msg of thread.getMessages()) {
      let data = parseEmailMessage(msg);
      // Only process if we found a valid amount
      if (data && data.amount !== "0.00") {
        if (processAndLogExpense(sheet, msg, data)) rowsAddedCount++;
      }
    }
  }
  sortExpensesByDate();
  const msgText = "Done! Fetched " + rowsAddedCount + " records.";
  ui ? ui.alert(msgText) : console.log(msgText);
}

function parseEmailMessage(msg) {
  let fullBody = (msg.getBody() || "").replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  let lowerBody = fullBody.toLowerCase();
  let lowerSub = (msg.getSubject() || "").toLowerCase();
  
  // Initialize result object
  let result = { 
    amount: "0.00", 
    merchant: "Unknown", 
    txnDate: "", 
    paymentMode: "Bank Account", 
    columnType: "Expense", 
    sourceBank: "Unknown Bank" 
  };

  let parsedSuccessfully = false;

  // --- TEMPLATES ---
// HDFC Debit Card - ATM Withdrawal
if (lowerBody.includes("atm withdrawal")) {
  let amtM = fullBody.match(
    /ATM\s+withdrawal\s+for\s+Rs\.?\s*([\d,]+(?:\.\d{1,2})?)/i
  );

  let merM = fullBody.match(
    /\bin\s+(.+?)\s+at\s+(.+?)\s+on\s+/i
  );

  let datM = fullBody.match(
    /\bon\s+(\d{2}-\d{2}-\d{4})\s+(\d{2}:\d{2}:\d{2})/i
  );

  result.amount = amtM ? amtM[1] : "0.00";

  if (merM) {
    result.merchant = merM[2].trim();
  } else {
    result.merchant = "ATM Withdrawal";
  }

  result.txnDate = datM
    ? datM[1] + " " + datM[2]
    : "";

  result.paymentMode = "ATM";
  result.columnType = "Expense";
  result.sourceBank = "HDFC Bank";

  parsedSuccessfully = true;
}
else if (fullBody.includes("towards VPA")) {
    let amtM = fullBody.match(/Rs\.\s*([\d,]+\.\d{2})/i);
    let datM = fullBody.match(/on\s+(\d{2}-\d{2}-\d{2,4})/);
    let bracketM = fullBody.match(/towards VPA\s+[^\s]+\s+\((.*?)\)/i);
    let rawVpaM = fullBody.match(/towards VPA\s+([^\s]+)/i);
    result.amount = amtM ? amtM[1] : "0.00";
    result.merchant = bracketM ? bracketM[1].trim() : (rawVpaM ? rawVpaM[1].trim() : "UPI Payment");
    result.txnDate = datM ? datM[1] : "";
    result.paymentMode = "UPI";
    parsedSuccessfully = true;
  }
  else if (fullBody.includes("ICICI Bank Account") && fullBody.includes("credited")) {
    let amtM = fullBody.match(/credited with INR\s*([\d,]+)/i);
    let datM = fullBody.match(/on\s+(\d{2}-[A-Za-z]{3}-\d{2,4})/i);
    result.amount = amtM ? amtM[1] : "0.00";
    result.txnDate = datM ? datM[1] : "";
    result.merchant = "ICICI Interest Payment";
    result.paymentMode = "Bank Credit";
    result.columnType = "Income";
    parsedSuccessfully = true;
  }
  else if (fullBody.includes("successfully debited") && fullBody.includes("towards")) {
    let amtM = fullBody.match(/Rs\.\s*([\d,]+)/i);
    let merM = fullBody.match(/towards\s+(.*?)\./i);
    let datM = fullBody.match(/on\s+(\d{4}\/\d{2}\/\d{2})/);
    result.amount = amtM ? amtM[1] : "0.00";
    result.merchant = merM ? merM[1].replace(/[*]/g, '').trim() : "Kotak Debit";
    result.txnDate = datM ? datM[1] : "";
    result.paymentMode = "Bank Account";
    parsedSuccessfully = true;
  }
  else if (fullBody.includes("credited to your Kotak Bank")) {
    let amtM = fullBody.match(/Rs\.\s*([\d,]+)/i);
    let datM = fullBody.match(/on\s+(\d{2}-[A-Za-z]{3}-\d{2,4})/i);
    let merM = fullBody.match(/transaction from\s+(.*?)\./i);
    result.amount = amtM ? amtM[1] : "0.00";
    result.merchant = merM ? merM[1].trim() : "NEFT Credit";
    result.txnDate = datM ? datM[1] : "";
    result.paymentMode = "NEFT Inward";
    result.columnType = "Income";
    parsedSuccessfully = true;
  }
  else if (fullBody.includes("State Bank of India") && fullBody.includes("Amount:")) {
    let amtM = fullBody.match(/Amount:\s*INR\s*([\d,]+\.\d{2})/i);
    let datM = fullBody.match(/Date:\s*([\d\/]+)/i);
    let merM = fullBody.match(/Sent by:\s*(.*?)\s*Sender/i);
    result.amount = amtM ? amtM[1] : "0.00";
    result.merchant = merM ? merM[1].trim() : "SBI Credit";
    result.txnDate = datM ? datM[1] : "";
    result.paymentMode = "NEFT Inward";
    result.columnType = "Income";
    parsedSuccessfully = true;
  }
  else if (lowerSub.includes("sip auto payment") || lowerSub.includes("buy order placed")) {
    let amtM = msg.getSubject().match(/(?:₹|Rs\.)\s*([\d,]+\.\d{2}|[\d,]+)/i) || fullBody.match(/(?:₹|Rs\.)\s*([\d,]+\.\d{2}|[\d,]+)/i);
    let merM = fullBody.match(/in\s+(.*?)\s+has been/i) || fullBody.match(/order of.*?\s+in\s+(.*?)\s+has/i);
    result.amount = amtM ? amtM[1] : "0.00";
    result.merchant = merM ? merM[1].trim() : "Mutual Fund SIP";
    result.paymentMode = "Auto-Debit";
    parsedSuccessfully = true;
  }
  else if (lowerBody.includes("smartpay")) {
    let amtM = fullBody.match(/Rs\.\s*([\d,]+\.\d{2})/i);
    let merM = fullBody.match(/Biller Name:\s*(.*?)(?:Unique|$)/i);
    result.amount = amtM ? amtM[1] : "0.00";
    result.merchant = merM ? merM[1].trim() : "SmartPay Bill";
    result.paymentMode = "Credit Card";
    result.sourceBank = "HDFC Bank (Card 8554)";
    parsedSuccessfully = true;
  }
  else if (fullBody.includes("HDFC Bank Credit Card") && fullBody.includes("debited")) {
    let amtM = fullBody.match(/Rs\.\s*([\d,]+\.\d{2})/i);
    let merM = fullBody.match(/towards\s+(.*?)\s+on/i);
    let datM = fullBody.match(/on\s+(\d{1,2}\s+[A-Za-z]+\s*,\s*\d{4})/i);
    result.amount = amtM ? amtM[1] : "0.00";
    result.merchant = merM ? merM[1].trim() : "HDFC Debit";
    result.txnDate = datM ? datM[1] : "";
    result.paymentMode = "Credit Card";
    parsedSuccessfully = true;
  }
  else if (fullBody.includes("Dividend") || lowerSub.includes("dividend")) {
    let amtM = fullBody.match(/(?:Rs\.|₹)\s*([\d,]+\.?\d*)/i);
    result.amount = amtM ? amtM[1] : "0.00";
    let merM = fullBody.match(/(.*?)(?:dividend)/i);
    result.merchant = merM ? merM[1].trim() : "Stock Dividend";
    result.paymentMode = "Bank Credit";
    result.columnType = "Income";
    result.sourceBank = "Investment Account";
    parsedSuccessfully = true;
  }
  else if (fullBody.includes("PRAN") || fullBody.includes("NPS")) {
    let amtM = fullBody.match(/Rs\.\s*([\d,]+\.\d{2})/i);
    result.amount = amtM ? amtM[1] : "0.00";
    result.merchant = "NPS Contribution";
    result.txnDate = fullBody.match(/(\d{2}\/\d{2}\/\d{4})/)?.[1] || "";
    result.paymentMode = "Auto-Debit";
    result.columnType = "Expense";
    result.sourceBank = "NPS";
    parsedSuccessfully = true;
  }

  // Fallback Engine
  if (!parsedSuccessfully) {
    result.merchant = extractMerchant(fullBody, msg.getSubject());
    let amtM = fullBody.match(/(?:Rs\.|INR|₹)\s*([\d,]+\.\d{2})/i) || fullBody.match(/(?:Rs\.|INR|₹)\s*([\d,]+)/i) || msg.getSubject().match(/(?:₹|Rs\.|INR)\s*([\d,]+)/i);
    if (amtM) result.amount = amtM[1];
    
    let dateM = fullBody.match(/on\s+\*?(\d{1,2}\s+[A-Za-z]{3},\s*\d{4})/i) || fullBody.match(/on\s+(\d{2}-\d{2}-\d{2,4})/);
    if (dateM) result.txnDate = dateM[1];
    
    if (lowerBody.includes("credited") || lowerBody.includes("received")) {
      result.columnType = "Income"; 
      result.paymentMode = "Inward Transfer";
    } else if (lowerBody.includes("debited") || lowerBody.includes("spent") || lowerSub.includes("payment")) {
      result.columnType = "Expense"; 
      result.paymentMode = (lowerBody.includes("credit card") || lowerSub.includes("credit card")) ? "Credit Card" : (lowerBody.includes("vpa") ? "UPI" : "Bank Account");
    }
  }

  return result;
}

function processAndLogExpense(sheet, msg, data) {
  let { amount, merchant, txnDate, paymentMode, columnType, sourceBank } = data;
  let fullBody = (msg.getBody() || "").replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  let lowerBody = fullBody.toLowerCase();
  
  // 1. Centralized Bank Logic
  if (["Unknown Bank", "Unknown", "General Bank"].includes(sourceBank)) {
    if (lowerBody.includes("amazon pay")) sourceBank = "Amazon Pay eGift Card";
    else if (lowerBody.includes("hdfc")) sourceBank = "HDFC Bank";
    else if (lowerBody.includes("icici")) sourceBank = "ICICI Bank";
    else if (lowerBody.includes("kotak")) sourceBank = "Kotak Bank";
    else if (lowerBody.includes("sbi")) sourceBank = "SBI Bank";
    else if (lowerBody.includes("axis")) sourceBank = "Axis Bank";
    else sourceBank = "General Bank";

    let details = extractBankDetails(fullBody);
    if (details && !sourceBank.includes("Amazon Pay")) {
      let typeLabel = (lowerBody.includes("credit card")) ? "Card " : "a/c ";
      sourceBank += " (" + typeLabel + details + ")";
    }
  }

  // 2. Formatting Date
  let dateObj = txnDate ? parseMyDate(txnDate) : msg.getDate();

  if (txnDate) {
    const messageDate = msg.getDate();

    // parseMyDate() currently returns only the date at 00:00:00.
    // Use the Gmail message time as the time component.
    dateObj.setHours(
      messageDate.getHours(),
      messageDate.getMinutes(),
      messageDate.getSeconds(),
      0
    );
  }

  // 3. Categorization
  let category = autoCategorize(merchant, msg.getSubject(), columnType, sourceBank);
  if (category === "IGNORE") return false;

  // 4. Traceable Logging
  let combinedLog = msg.getSubject() + " | " + getReadableText(fullBody).substring(0, 100);
  
  // 5. Append
  sheet.appendRow([dateObj, amount.replace(/,/g, ''), cleanText(merchant), paymentMode, category, columnType, sourceBank, combinedLog]);
  return true;
}

function extractMerchant(body, subject) {
  let text = body.replace(/<[^>]*>/g, ' '); // Strip HTML tags

  // 1. Specific SIP Reminder rule
  if (text.includes("Upcoming SIP installment")) return "SIP Reminder";

  // 2. Credit Card Payment confirmations
  if (text.includes("payment confirmation") && text.includes("credit card payment was successful")) return "Credit Card Payment";

  // 3. ICICI / UPI Card Patterns: Look for "at [Merchant] on"
  let atMatch = text.match(/at\s+([A-Z0-9\s]+?)\s+(?:on|the|is)/i);
  if (atMatch) return atMatch[1].trim();

  // 4. Indian Clearing Corp (SIPs)
  if (text.includes("Indian Clearing")) return "Indian Clearing Corp";

  // 5. VPA/UPI Names
  let vpaMatch = text.match(/VPA\s+[^\s]+\s+([A-Z\s\.]+)(?:\s+on|\s+\d{2})/i);
  if (vpaMatch) return vpaMatch[1].trim();

  // 6. Bank Credit/Cashback rules
  if (text.includes("credited to your HDFC Bank")) return "HDFC Bank Credit";
  if (text.includes("cashback")) return "Cashback Received";

  return "Unknown Merchant";
}

function parseMyDate(dateStr) {
  if (!dateStr) return new Date();

  dateStr = dateStr.trim();

  // DD-MM-YYYY HH:mm:ss
  let dateTimeMatch = dateStr.match(
    /^(\d{2})-(\d{2})-(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/
  );

  if (dateTimeMatch) {
    return new Date(
      Number(dateTimeMatch[3]),
      Number(dateTimeMatch[2]) - 1,
      Number(dateTimeMatch[1]),
      Number(dateTimeMatch[4]),
      Number(dateTimeMatch[5]),
      Number(dateTimeMatch[6])
    );
  }

  // Existing date-only logic
  let cleanStr = dateStr
    .replace(/[\/\-\.\,]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  let parts = cleanStr.split(' ');

  let day, month, year;

  let firstPart = parseInt(parts[0]);
  let lastPart = parseInt(parts[parts.length - 1]);

  if (firstPart > 2000) {
    // YYYY MM DD
    year = firstPart;
    month = parseInt(parts[1]) - 1;
    day = parseInt(parts[2]);
  } else {
    // DD MM YYYY / DD Month YYYY
    day = firstPart;
    month = parts[1];
    year = lastPart;

    if (year < 100) year += 2000;

    if (isNaN(month)) {
      const months = [
        "jan", "feb", "mar", "apr", "may", "jun",
        "jul", "aug", "sep", "oct", "nov", "dec"
      ];

      month = months.indexOf(
        month.toLowerCase().substring(0, 3)
      );
    } else {
      month = parseInt(month) - 1;
    }
  }

  return new Date(year, month, day);
}

function getReadableText(text) {
  return text
    .replace(/<[^>]*>/g, ' ')           // Remove any leftover HTML tags
    .replace(/&nbsp;/g, ' ')            // Remove HTML space entities
    .replace(/&amp;/g, '&')             // Fix ampersands
    .replace(/\s+/g, ' ')               // Collapse multiple spaces into one
    .replace(/\s*\.\s*/g, '. ')         // Ensure proper spacing after dots
    .trim();                            // Trim start/end whitespace
}

function standardizeDate(dateStr) {
  if (!dateStr) return null;
  // This is a basic parser; you might need to expand it based on the exact formats in your emails
  let d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr; // Return original if parsing fails
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function extractBankDetails(text) {
  let match = text.match(/(?:account|a\/c|card|ending)\s*\*?\s*(?:ending)?\s*\*?\s*([X\d]{3,5})/i);
  return match ? match[1] : "";
}

function cleanText(text) {
  if (!text) return "";
  return text.replace(/[*]/g, '').replace(/\s+/g, ' ').trim();
}

// BUTTON 2: INTERACTIVE DUPLICATE & MERGE WIZARD
function clearSheetDuplicates() {
  const ui = getUiSafe();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expenses");
  let data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return;

  let trackedKeys = {}; 
  let rowsToDelete = []; 

  // First Pass: Scan all rows to handle Merges and Deletions
  for (let r = 1; r < data.length; r++) {
    let rowDate = data[r][0];

  let formattedDate;

  if (rowDate instanceof Date && !isNaN(rowDate.getTime())) {
    formattedDate = Utilities.formatDate(
      rowDate,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm:ss"
    );
  } else {
    formattedDate = rowDate.toString().trim();
  }

  let rowAmount = Number(data[r][1]).toFixed(2);

  let currentMerchant = data[r][2].toString().trim();
  let currentCategory = data[r][4].toString().trim();

  // Duplicate key = Date + Time + Amount
  let coreKey = formattedDate + "_" + rowAmount;

    if (trackedKeys[coreKey] !== undefined) {
      let baselineRowIndex = trackedKeys[coreKey];
      
      // Fetch Baseline values directly from the sheet arrays
      let baselineMerchant = sheet.getRange(baselineRowIndex, 3).getValue().toString().trim();
      let baselineCategory = sheet.getRange(baselineRowIndex, 5).getValue().toString().trim();

      // Silent deletion shortcut if rows are 100% identical mirror copies
      if (baselineMerchant.toLowerCase() === currentMerchant.toLowerCase() && baselineCategory.toLowerCase() === currentCategory.toLowerCase()) {
        rowsToDelete.push(r + 1);
        continue;
      }

      // --- STEP 1: WIZARD FOR MERGING ---
      if (baselineMerchant !== currentMerchant || baselineCategory !== currentCategory) {
        let mergeChoice = ui.alert(
          "Merge Matching Rows?",
          "Identical Date & Amount (Rs. " + rowAmount + " on " + formattedDate + ") found.\n\n" +
          "Row A: " + baselineMerchant + " [" + baselineCategory + "]\n" +
          "Row B: " + currentMerchant + " [" + currentCategory + "]\n\n" +
          "Would you like to MERGE these into a single row?",
          ui.ButtonSet.YES_NO
        );

        if (mergeChoice === ui.Button.YES) {
          let catChoice = ui.alert(
            "Select Category for Merged Row",
            "Which category should be preserved?\n\n" +
            "• Click YES for Row A: \"" + baselineCategory + "\"\n" +
            "• Click NO for Row B: \"" + currentCategory + "\"",
            ui.ButtonSet.YES_NO
          );
          
          let finalCategory = (catChoice === ui.Button.YES) ? baselineCategory : currentCategory;
          let combinedMerchant = baselineMerchant + " / " + currentMerchant;
          
          sheet.getRange(baselineRowIndex, 3).setValue(combinedMerchant);
          sheet.getRange(baselineRowIndex, 5).setValue(finalCategory);
          
          rowsToDelete.push(r + 1); // Delete Row B after blending its data up
          continue; 
        }
      }

      // --- STEP 2: WIZARD FOR ROW VERSION DELETION (If they decline merging) ---
      let deleteChoice = ui.alert(
        "Choose Row Version to REMOVE",
        "Amount: Rs. " + rowAmount + " | Date: " + formattedDate + "\n\n" +
        "Which row version would you like to delete from the sheet?\n\n" +
        "• Click YES to remove: \"" + baselineMerchant + "\"\n" +
        "• Click NO to remove: \"" + currentMerchant + "\"\n" +
        "• Click CANCEL to keep both separate rows.",
        ui.ButtonSet.YES_NO_CANCEL
      );

      if (deleteChoice === ui.Button.YES) {
        rowsToDelete.push(baselineRowIndex);
        trackedKeys[coreKey] = r + 1; // Hand pointer tracking authority over to Row B
      } else if (deleteChoice === ui.Button.NO) {
        rowsToDelete.push(r + 1);
      }
      
    } else {
      trackedKeys[coreKey] = r + 1;
    }
  }

  // Second Pass: Safe reverse-ordered row extraction exuecution
  if (rowsToDelete.length > 0) {
    rowsToDelete = rowsToDelete.filter((item, index) => rowsToDelete.indexOf(item) === index);
    rowsToDelete.sort(function(a, b) { return b - a; });
    
    for (let d = 0; d < rowsToDelete.length; d++) {
      sheet.deleteRow(rowsToDelete[d]);
    }
    ui.alert("Wizard finished! Modified, merged, or removed " + rowsToDelete.length + " conflict rows.");
  } else {
    ui.alert("No duplicate conflicts found!");
  }
}

function sortExpensesByDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Expenses");
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; 

  // Force the date column to be strictly YYYY-MM-DD
  const dateRange = sheet.getRange(2, 1, lastRow - 1, 1);
  dateRange.setNumberFormat("yyyy-mm-dd HH:mm:ss");

  // Sort the data
  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  range.sort({column: 1, ascending: true});
}

// 3. CATEGORIZATION ENGINE
function autoCategorize(merchant, subject, columnType, sourceBank) {
  const mappings = getMappings(); 
  const m = merchant.toLowerCase();
  const bank = (sourceBank || "").toLowerCase();

  // TIER 1: Check Mappings Sheet (Highest Priority - keeps your manual overrides)
  for (let key in mappings) {
    if (m.includes(key.toLowerCase())) return mappings[key];
  }

  // TIER 2: Bank Account-Based Defaults (The new logic)
  if (bank.includes("8910")) return "Misc";0
  
  if (bank.includes("3141")) return "Groceries"; // Defaulting to Groceries for now
  
  if (bank.includes("xx2722")) return "Shopping"; // Defaulting to Shopping for now
  
  if (bank.includes("xx094")) return "Food"; // Defaulting to Food for now
  
  if (bank.includes("xx540")) return "Travel";

  let s = subject.toLowerCase();

  // 1. HIGH PRIORITY OVERRIDES (Check these before anything else)
  if (m.includes("nps") || m.includes("pran")) return "Investment";
  if (m.includes("atindra") || s.includes("self transfer")) return "Self Transfer";

  // 2. INCOME LOGIC (Consolidated)
  if (columnType === "Income") {
    if (m.includes("dividend") || m.includes("interest") || m.includes("int.pd")) {
      return "Investment";
    }
    return "Inward Income";
  }

  // 2. Functional Categories
  if (m.includes("pharmeasy") || m.includes("apollo") || m.includes("1mg") || 
      m.includes("medical") || m.includes("pharmacy") || m.includes("hospital") || m.includes("bima")) {
    return "Health";
  }

  if (m.includes("finzoom") || m.includes("investment") || s.includes("sip") || 
      m.includes("fund") || m.includes("nippon") || m.includes("zerodha") || 
      m.includes("groww") || m.includes("payu") || m.includes("mutual fund")) {
    return "Investment";
  }

  if (m.includes("jio") || m.includes("airtel") || m.includes("vi ") || 
      m.includes("electricity") || m.includes("cesc") || m.includes("wbseb") || 
      m.includes("broadband") || m.includes("recharge") || m.includes("southern power") || 
      m.includes("telangana") || m.includes("smartpay") || 
      m.includes("athkur srilatha") || m.includes("mukteshwari b")) {
    return "Utility";
  }

  if (m.includes("cinema") || m.includes("pvr") || m.includes("netflix") || 
      m.includes("bookmyshow")) {
    return "Entertainment";
  }

  if (m.includes("zomato") || m.includes("swiggy") || m.includes("restaurant") || 
      m.includes("dine") || m.includes("pizza")) {
    return "Food";
  }

  if (m.includes("blinkit") || m.includes("instamart") || m.includes("dmart") || 
      m.includes("zepto") || m.includes("bigbasket")) {
    return "Groceries";
  }

  if (m.includes("uber") || m.includes("ola") || m.includes("fuel") || 
      m.includes("petrol") || m.includes("irctc") || m.includes("flight") || m.includes("gasoline")) {
    return "Travel";
  }

  if (m.includes("amazon") || m.includes("flipkart") || m.includes("myntra") || 
      m.includes("ajio") || m.includes("meesho")) {
    return "Shopping";
  }

  // 1. FILTER OUT NOISE
  if (s.includes("no-reply") || s.includes("add to contact") || s.includes("marketing") || s.includes("failed") 
      || s.includes("sorry") || s.includes("upcoming sip") || s.includes("surprise") || s.includes("wallet points") 
      || s.includes("cashback")
      || s.includes("reminder") 
      || s.includes("due tomorrow") 
      || s.includes("bill payment is due")) { // Added specific rules for reminders
    return "IGNORE";
  }

  if (m.includes("cashback")) return "Inward Income";
  if (merchant === "Reminder Auto Debit SIP") return "Investment";

  // Default
  return "Misc";
}

function getMappings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Mappings");
  if (!sheet) return {}; // Return empty object if sheet doesn't exist
  
  const data = sheet.getDataRange().getValues();
  let mappings = {};
  
  // Start from 1 to skip header row
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { // Column A: Merchant
      mappings[data[i][0].toString().toLowerCase()] = data[i][1]; // Column B: Category
    }
  }
  return mappings;
}

function syncMapping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expSheet = ss.getSheetByName("Expenses");
  const mapSheet = ss.getSheetByName("Mappings");
  
  if (!mapSheet || !expSheet) {
    SpreadsheetApp.getUi().alert("Error: 'Expenses' or 'Mappings' sheet not found.");
    return;
  }
  
  // 1. Get all data from Expenses
  const expData = expSheet.getDataRange().getValues();
  let uniqueMappings = {};
  
  // 2. Loop through all rows (skipping header)
  // Assuming Col C is Merchant (index 2) and Col E is Category (index 4)
  for (let i = 1; i < expData.length; i++) {
    let merchant = expData[i][2];
    let category = expData[i][4];
    
    // Only map if merchant and category exist and aren't 'Misc'
    // This ensures only "trained" data is added
    if (merchant && category && category.toString().toLowerCase() !== "misc") {
      uniqueMappings[merchant.toString().toLowerCase()] = category;
    }
  }
  
  // 3. Clear existing Mappings (keeping only the header)
  // We clear from row 2 downwards
  mapSheet.getRange(2, 1, Math.max(mapSheet.getLastRow() - 1, 1), 2).clearContent();
  
  // 4. Convert object to array for writing back to sheet
  let output = [];
  for (let merchant in uniqueMappings) {
    output.push([merchant, uniqueMappings[merchant]]);
  }
  
  // 5. Write to Mappings sheet
  if (output.length > 0) {
    mapSheet.getRange(2, 1, output.length, 2).setValues(output);
    SpreadsheetApp.getUi().alert("Success! " + output.length + " unique mappings synced.");
  } else {
    SpreadsheetApp.getUi().alert("No valid mappings found to sync.");
  }
}

// 4. MENU
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Expense Tracker')
    .addItem('1. Fetch New Emails', 'fetchAllEmailsRaw')
    .addItem('2. Clean Duplicate Rows', 'clearSheetDuplicates')
    .addItem('3. Sync Mappings', 'syncMapping')
    .addToUi();
}
