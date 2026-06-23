const ALLOWED_CATEGORIES = "Travel, Shopping, Entertainment, Utility, Groceries, Investment, Food, Self Transfer, Health, Misc";

function getLastExtractedDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Expenses");
  
  // Safety check: Does the sheet exist?
  if (!sheet) {
    throw new Error("Sheet named 'Expenses' not found. Please check your sheet name.");
  }
  
  const data = sheet.getDataRange().getValues();
  
  // If sheet is empty (only header or no rows), default to a start date
  if (data.length <= 1) return "2025/05/01"; 

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
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expenses");
  if (!sheet) { ui.alert("Error: 'Expenses' sheet not found."); return; }

  // 1. Get the last date from the sheet automatically [cite: 57, 63]
  const startDateStr = getLastExtractedDate(); 
  
  // 2. Set end date to tomorrow so that the search captures everything up to today
  let tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const endDateStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), "yyyy/MM/dd");
  
  const afterDate = startDateStr.replace(/\//g, '-');
  const beforeDate = endDateStr.replace(/\//g, '-');
  
  // 3. Updated Query: Now includes SmartPay and correct date range 
  const gmailSearchQuery = `("debited" OR "credited" OR "Transaction alert" OR "SIP Auto Payment" OR "Buy order placed" OR "SmartPay") after:${afterDate} before:${beforeDate}`;
  const threads = GmailApp.search(gmailSearchQuery, 0, 100); 

  if (threads.length === 0) {
    ui.alert("No new transaction emails found since " + startDateStr);
    return;
  }

  let rowsAddedCount = 0;

  for (let i = 0; i < threads.length; i++) {
    const messages = threads[i].getMessages();
    for (let j = 0; j < messages.length; j++) {
      const msg = messages[j];
      let fullBody = (msg.getPlainBody() || "").replace(/\s+/g, ' ').trim(); 
      let amount = "0.00", merchant = "Unknown", txnDate = "", paymentMode = "Bank Account", columnType = "Expense", sourceBank = "Unknown Bank"; 
      let lowerBody = fullBody.toLowerCase(), lowerSub = (msg.getSubject() || "").toLowerCase();

      let bankBase = "Unknown Bank";
      if (lowerBody.includes("hdfc bank") || lowerSub.includes("hdfc")) bankBase = "HDFC Bank";
      else if (lowerBody.includes("icici bank") || lowerSub.includes("icici")) bankBase = "ICICI Bank";
      else if (lowerBody.includes("kotak bank") || lowerSub.includes("kotak")) bankBase = "Kotak Bank";
      else if (lowerBody.includes("state bank of india") || lowerBody.includes("sbi")) bankBase = "SBI";
      else if (lowerBody.includes("axis bank") || lowerBody.includes("axis")) bankBase = "Axis Bank";

      let parsedSuccessfully = false;

            // Template 1: Axis / Generic UPI towards VPA (Fixed to extract bracket name over raw VPA link)
      if (fullBody.includes("towards VPA")) {
        let amtM = fullBody.match(/Rs\.\s*([\d,]+\.\d{2})/i);
        let datM = fullBody.match(/on\s+(\d{2}-\d{2}-\d{2,4})/);
        
        // Match name inside brackets first. If missing, fall back to raw VPA address.
        let bracketM = fullBody.match(/towards VPA\s+[^\s]+\s+\((.*?)\)/i);
        let rawVpaM = fullBody.match(/towards VPA\s+([^\s]+)/i);
        
        amount = amtM ? amtM[1] : "0.00";
        merchant = bracketM ? bracketM[1].trim() : (rawVpaM ? rawVpaM[1].trim() : "UPI Payment");
        txnDate = datM ? datM[1] : "";
        paymentMode = "UPI";
        
        let numM = fullBody.match(/account ending\s*(\d+)/i);
        sourceBank = bankBase + (numM ? " (a/c " + numM[1] + ")" : "");
        parsedSuccessfully = true;
      }
      else if (fullBody.includes("ICICI Bank Account") && fullBody.includes("credited")) {
        let amtM = fullBody.match(/credited with INR\s*([\d,]+)/i);
        let datM = fullBody.match(/on\s+(\d{2}-[A-Za-z]{3}-\d{2,4})/i);
        amount = amtM ? amtM[1] : "0.00";
        txnDate = datM ? datM[1] : "";
        merchant = "ICICI Interest Payment";
        paymentMode = "Bank Credit";
        columnType = "Income";
        let numM = fullBody.match(/Account\s+([X\d]+)/i);
        sourceBank = "ICICI Bank" + (numM ? " (" + numM[1] + ")" : "");
        parsedSuccessfully = true;
      }
      else if (fullBody.includes("successfully debited") && fullBody.includes("towards")) {
        let amtM = fullBody.match(/Rs\.\s*([\d,]+)/i);
        let merM = fullBody.match(/towards\s+(.*?)\./i);
        let datM = fullBody.match(/on\s+(\d{4}\/\d{2}\/\d{2})/);
        amount = amtM ? amtM[1] : "0.00";
        merchant = merM ? merM[1].replace(/[*]/g, '').trim() : "Kotak Debit";
        txnDate = datM ? datM[1] : "";
        paymentMode = "Bank Account";
        sourceBank = "Kotak Bank";
        parsedSuccessfully = true;
      }
      else if (fullBody.includes("credited to your Kotak Bank")) {
        let amtM = fullBody.match(/Rs\.\s*([\d,]+)/i);
        let datM = fullBody.match(/on\s+(\d{2}-[A-Za-z]{3}-\d{2,4})/i);
        let merM = fullBody.match(/transaction from\s+(.*?)\./i);
        amount = amtM ? amtM[1] : "0.00";
        merchant = merM ? merM[1].trim() : "NEFT Credit";
        txnDate = datM ? datM[1] : "";
        paymentMode = "NEFT Inward";
        columnType = "Income";
        let numM = fullBody.match(/a\/c\s+([X\d]+)/i);
        sourceBank = "Kotak Bank" + (numM ? " (" + numM[1] + ")" : "");
        parsedSuccessfully = true;
      }
      else if (fullBody.includes("State Bank of India") && fullBody.includes("Amount:")) {
        let amtM = fullBody.match(/Amount:\s*INR\s*([\d,]+\.\d{2})/i);
        let datM = fullBody.match(/Date:\s*([\d\/]+)/i);
        let merM = fullBody.match(/Sent by:\s*(.*?)\s*Sender/i);
        amount = amtM ? amtM[1] : "0.00";
        merchant = merM ? merM[1].trim() : "SBI Credit";
        txnDate = datM ? datM[1] : "";
        paymentMode = "NEFT Inward";
        columnType = "Income";
        let numM = fullBody.match(/Your A\/c:\s*([X\d]+)/i);
        sourceBank = "SBI" + (numM ? " (" + numM[1] + ")" : "");
        parsedSuccessfully = true;
      }
      else if (lowerSub.includes("sip auto payment") || lowerSub.includes("buy order placed")) {
        let amtM = msg.getSubject().match(/(?:₹|Rs\.)\s*([\d,]+\.\d{2}|[\d,]+)/i) || fullBody.match(/(?:₹|Rs\.)\s*([\d,]+\.\d{2}|[\d,]+)/i);
        let merM = fullBody.match(/in\s+(.*?)\s+has been/i) || fullBody.match(/order of.*?\s+in\s+(.*?)\s+has/i);
        amount = amtM ? amtM[1] : "0.00";
        merchant = merM ? merM[1].trim() : "Mutual Fund SIP";
        paymentMode = "Auto-Debit";
        let details = extractBankDetails(fullBody);
        sourceBank = bankBase + (details ? " (a/c " + details + ")" : "");
        parsedSuccessfully = true;
      }
      
      else if (lowerBody.includes("smartpay")) {
        // Extract Amount
        let amtM = fullBody.match(/Rs\.\s*([\d,]+\.\d{2})/i);
        amount = amtM ? amtM[1] : "0.00";
        
        // Extract Merchant/Biller Name
        let merM = fullBody.match(/Biller Name:\s*(.*?)(?:Unique|$)/i);
        merchant = merM ? merM[1].trim() : "SmartPay Bill";
        
        // Set other defaults
        txnDate = ""; // You may need to add a regex for date if present
        paymentMode = "SmartPay";
        columnType = "Expense";
        sourceBank = "SmartPay";
        
        parsedSuccessfully = true;
      }
      else if (fullBody.includes("HDFC Bank Credit Card") && fullBody.includes("debited")) {
          let amtM = fullBody.match(/Rs\.\s*([\d,]+\.\d{2})/i);
          let merM = fullBody.match(/towards\s+(.*?)\s+on/i); // Improved capture
          let datM = fullBody.match(/on\s+(\d{1,2}\s+[A-Za-z]+\s*,\s*\d{4})/i);
          
          amount = amtM ? amtM[1] : "0.00";
          merchant = merM ? merM[1].trim() : "HDFC Debit";
          txnDate = datM ? datM[1] : "";
          paymentMode = "Credit Card";
          sourceBank = "HDFC Bank";
          parsedSuccessfully = true;
      }
      else if (fullBody.includes("Dividend") || lowerSub.includes("dividend")) {
      // 1. Extract Amount: Matches "Rs. 4" or "₹4"
      let amtM = fullBody.match(/(?:Rs\.|₹)\s*([\d,]+\.?\d*)/i);
      amount = amtM ? amtM[1] : "0.00";
      
      // 2. Extract Merchant: Looks for text appearing BEFORE the word "dividend"
      // This assumes the format is "[Stock Name] dividend..."
      let merM = fullBody.match(/(.*?)(?:dividend)/i);
      merchant = merM ? merM[1].trim() : "Stock Dividend";
      
      // 3. Clean up the Merchant: If the extracted name is too long, truncate it
      if (merchant.length > 30) {
          // Often the name is at the end of the previous sentence
          let words = merchant.split(' ');
          merchant = words.slice(-2).join(' '); // Take only the last 2 words (e.g., "TATA STEEL")
      }
      
      txnDate = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      paymentMode = "Bank Credit";
      columnType = "Income";
      sourceBank = "Investment Account";
      parsedSuccessfully = true;
    }
    else if (fullBody.includes("PRAN") || fullBody.includes("NPS")) {
      // Extract Amount
      let amtM = fullBody.match(/Rs\.\s*([\d,]+\.\d{2})/i);
      amount = amtM ? amtM[1] : "0.00";
      
      // Merchant is NPS
      merchant = "NPS Contribution";
      
      // Extract Date (Look for DD/MM/YYYY or similar)
      let datM = fullBody.match(/(\d{2}\/\d{2}\/\d{4})/);
      txnDate = datM ? datM[1] : Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      
      paymentMode = "Auto-Debit";
      columnType = "Expense"; // It's an expense/investment outflow
      sourceBank = "NPS";
      parsedSuccessfully = true;
    }

  // Fallback Engine
if (!parsedSuccessfully) {
  let details = extractBankDetails(fullBody);
  let amtM = fullBody.match(/(?:Rs\.|INR|₹)\s*([\d,]+\.\d{2})/i) || fullBody.match(/(?:Rs\.|INR|₹)\s*([\d,]+)/i) || msg.getSubject().match(/(?:₹|Rs\.|INR)\s*([\d,]+)/i);
  if (amtM) amount = amtM[1];
  
  let dateM = fullBody.match(/on\s+\*?(\d{1,2}\s+[A-Za-z]{3},\s*\d{4})/i) || fullBody.match(/on\s+(\d{2}-\d{2}-\d{2,4})/);
  if (dateM) txnDate = dateM[1];
  
  let isCreditCard = lowerBody.includes("credit card") || lowerSub.includes("credit card");
  
  if (lowerBody.includes("credited") || lowerBody.includes("received")) {
    columnType = "Income"; 
    paymentMode = "Inward Transfer";
  } else if (lowerBody.includes("debited") || lowerBody.includes("spent") || lowerSub.includes("payment")) {
    columnType = "Expense"; 
    paymentMode = isCreditCard ? "Credit Card" : (lowerBody.includes("vpa") ? "UPI" : "Bank Account");
  }

  let merM = fullBody.match(/towards\s+([^on\.]+)\s+on/i) || fullBody.match(/from\s+([^on\.]+)\s+via/i) || fullBody.match(/Info:\s*([^\s]+)/i);
  if (merM) { 
    merchant = cleanText(merM[1]); // Ensure you use the cleanText helper
    if (merchant.length > 50) merchant = merchant.substring(0, 47) + "..."; 
  }

  // 1. Identify the base bank name
  let bankNameMatch = fullBody.match(/([A-Za-z0-9]+)\s+Bank/i);
  let baseBankName = bankNameMatch ? bankNameMatch[1] + " Bank" : (lowerBody.includes("sbi") ? "SBI" : "General Bank");

  // 2. Identify the identifier (Account/Card digits)
  let extractedDigits = fullBody.match(/(?:account|a\/c|card|ending)\s*\*?\s*(?:ending)?\s*\*?\s*([X\d]{3,5})/i);
  let identifier = details || (extractedDigits ? extractedDigits[1] : "");

  // 3. Construct the sourceBank string
  if (identifier) {
    let typeLabel = isCreditCard ? "Card " : "a/c ";
    sourceBank = baseBankName + " (" + typeLabel + identifier + ")";
  } else {
    sourceBank = baseBankName;
  }
}

      // --- UPDATED SAFETY LOGGING ZONE ---
    // --- UPDATED LOGGING ZONE ---
    if (amount !== "0.00") {
      amount = amount.replace(/,/g, '');
      
      // Clean the merchant name immediately
      merchant = cleanText(merchant); 

      // 1. Determine the source date
      let dateObj = txnDate ? parseMyDate(txnDate) : msg.getDate();
      if (isNaN(dateObj.getTime())) dateObj = msg.getDate();
      let emailDateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
      
      let category = autoCategorize(merchant, msg.getSubject(), columnType);
      
      // Generate a much cleaner log snippet
      let cleanLogText = fullBody.replace(/https?:\/\/[^\s]+/g, '') // Remove URLs
                                .replace(/[*]/g, '')              // Remove asterisks
                                .replace(/\s+/g, ' ')             // Normalize spaces
                                .trim();
      let textToLog = cleanLogText.substring(0, 150);
      
      // --- CENTRALIZED SOURCE BANK LOGIC ---
      // If it's still "Unknown Bank", determine it based on the full body text
      if (sourceBank === "Unknown Bank" || sourceBank === "Unknown") {
        let bankBase = "General Bank";
        if (lowerBody.includes("hdfc")) bankBase = "HDFC Bank";
        else if (lowerBody.includes("icici")) bankBase = "ICICI Bank";
        else if (lowerBody.includes("kotak")) bankBase = "Kotak Bank";
        else if (lowerBody.includes("sbi")) bankBase = "SBI";
        else if (lowerBody.includes("axis")) bankBase = "Axis Bank";

        let details = extractBankDetails(fullBody);
        if (details) {
          let typeLabel = (lowerBody.includes("credit card") || lowerSub.includes("credit card")) ? "Card " : "a/c ";
          sourceBank = bankBase + " (" + typeLabel + details + ")";
        } else {
          sourceBank = bankBase;
        }
      }
      // --- END OF CENTRALIZED LOGIC ---
      sheet.appendRow([emailDateStr, amount, merchant, paymentMode, category, columnType, sourceBank, textToLog]);
      rowsAddedCount++;
    }
    }
  }

  sortExpensesByDate();
  ui.alert("Done! Fetched " + rowsAddedCount + " records raw. Now use the clear duplicates tool.");
}

function parseMyDate(dateStr) {
  if (!dateStr) return new Date();
  
  // Clean: Replace separators with spaces, remove commas, normalize whitespace
  let cleanStr = dateStr.replace(/[\/\-\.\,]/g, ' ').replace(/\s+/g, ' ').trim();
  let parts = cleanStr.split(' ');
  
  let day, month, year;

  // Logic: Detect if Year is in the first position (e.g., 2026 06 15) or third (e.g., 15 06 2026)
  let firstPart = parseInt(parts[0]);
  let lastPart = parseInt(parts[parts.length - 1]);

  if (firstPart > 2000) {
    // Format: YYYY MM DD
    year = firstPart;
    month = parseInt(parts[1]) - 1;
    day = parseInt(parts[2]);
  } else {
    // Format: DD MM YYYY (or DD Month YYYY)
    day = firstPart;
    month = parts[1];
    year = lastPart;
    
    // Ensure 2-digit years are treated as 2000s
    if (year < 100) year += 2000;
    
    // Handle Month names
    if (isNaN(month)) {
      const months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
      month = months.indexOf(month.toLowerCase().substring(0, 3));
    } else {
      month = parseInt(month) - 1;
    }
  }
  
  return new Date(year, month, day);
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
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expenses");
  let data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return;

  let trackedKeys = {}; 
  let rowsToDelete = []; 

  // First Pass: Scan all rows to handle Merges and Deletions
  for (let r = 1; r < data.length; r++) {
    let rowDate = data[r][0];
    let formattedDate = rowDate instanceof Date ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : rowDate.toString().trim();
    let rowAmount = Number(data[r][1]).toFixed(2);
    
    let currentMerchant = data[r][2].toString().trim();
    let currentCategory = data[r][4].toString().trim();

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
  
  // Get data range excluding the header row
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // Nothing to sort
  
  // Define range (starting from row 2, column 1, down to last row)
  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  
  // Sort by Column 1 (Date) in Ascending order (oldest at top, newest at bottom)
  range.sort({column: 1, ascending: true});
}

// 3. CATEGORIZATION ENGINE
function autoCategorize(merchant, subject, columnType) {
  let m = merchant.toLowerCase();
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

  // Default
  return "Misc";
}

// 4. MENU
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Expense Tracker')
    .addItem('1. Fetch New Emails', 'fetchAllEmailsRaw')
    .addItem('2. Clean Duplicate Rows', 'clearSheetDuplicates')
    .addToUi();
}