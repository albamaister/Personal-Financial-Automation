// CONFIGURATION
const API_KEY =
  PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
const SHEET_ID =
  PropertiesService.getScriptProperties().getProperty("SHEET_ID");
const SHEET_NAME =
  PropertiesService.getScriptProperties().getProperty("SHEET_NAME");
const PROCESSED_LABEL =
  PropertiesService.getScriptProperties().getProperty("PROCESSED_LABEL");

function runExpenseAgent() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  let processedCount = 0;

  // 1. Setup label if it doesn't exist
  if (!GmailApp.getUserLabelByName(PROCESSED_LABEL)) {
    GmailApp.createLabel(PROCESSED_LABEL);
  }
  const label = GmailApp.getUserLabelByName(PROCESSED_LABEL);

  // --- LEER IDs EXISTENTES ---
  const lastRowCheck = sheet.getLastRow();
  let existingIds = [];
  if (lastRowCheck > 1) {
    const data = sheet.getRange(2, 6, lastRowCheck - 1).getValues();
    existingIds = data.flat();
  }
  // ---------------------------------------------------------

  // 2. Search for UNPROCESSED threads
  const query = `from:capitalone (subject:"transaction" OR subject:"withdrawal notice" OR subject:"You sent money with Zelle") -"CAPITAL ONE has initiated" -label:${PROCESSED_LABEL}`;

  const threads = GmailApp.search(query, 0, 15);

  if (threads.length === 0) {
    console.log("No new expense emails found.");
    return;
  }

  for (const thread of threads) {
    const messages = thread.getMessages();
    let threadSuccess = true;

    for (const msg of messages) {
      if (existingIds.includes(msg.getId())) {
        console.log(`Skipping duplicate message: ${msg.getId()}`);
        continue;
      }

      const body = msg.getPlainBody();
      const subject = msg.getSubject();
      const date = msg.getDate();

      // 3. Call the Agent (Gemini)
      try {
        const expenseData = extractDataWithGemini(body, subject, date);
        if (expenseData) {
          // 4. Save to Sheets
          sheet.appendRow([
            expenseData.date,
            expenseData.merchant,
            expenseData.category,
            expenseData.amount,
            expenseData.description,
            msg.getId(),
          ]);

          processedCount++;
          console.log(
            `Processed: ${expenseData.merchant} - ${expenseData.amount}`
          );

          existingIds.push(msg.getId());
        }

        // Security pause (10 seconds to avoid error 429)
        Utilities.sleep(10000);
      } catch (e) {
        console.error(`Error processing message ${msg.getId()}: ${e.message}`);
        threadSuccess = false;
        break;
      }
    }

    // 5. Mark thread as processed (ONLY IF NO ERRORS)
    if (threadSuccess) {
      thread.addLabel(label);
    } else {
      console.warn("Thread skipped due to errors. Will retry next run.");
    }
  }

  // 6. Sort Data
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 6).sort({ column: 1, ascending: true });
  }

  createDashboard();

  // 7. SEND EMAIL NOTIFICATION
  if (processedCount > 0) {
    const myEmail = Session.getActiveUser().getEmail();
    const subject = `Expense Update: ${processedCount} new transactions processed`;
    const body = `
      Hola,
      
      Your financial agent has completed its scheduled execution.
      
      Summary:
      - Successfully processed: ${processedCount} transactions.
      - Your "Expenses" spreadsheet has been updated.

      Link to your sheet: https://docs.google.com/spreadsheets/d/${SHEET_ID}
      
      
      Regards,
      Your Apps Script Agent
    `;

    GmailApp.sendEmail(myEmail, subject, body);
    console.log(`Notification email sent to ${myEmail}`);
  }
}

function extractDataWithGemini(emailBody, subject, emailDate) {
  const model = "gemini-2.5-flash-lite";

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${API_KEY}`;

  const prompt = `
    Act as a financial data extractor. Analyze the following bank transaction email.
    
    ### PRIORITY RULES (Check these first):

    1. IF Subject contains "withdrawal notice":
       - Merchant: Extract the specific entity name initiating the withdrawal (e.g., from "MIDAMERICAN has initiated...", extract "MIDAMERICAN").
       - Category: "Withdrawal".
       - Description: Extract the full sentence describing the action (e.g., "MIDAMERICAN has initiated the following withdrawal from your BryanAlba…3698 account").

    2. IF Subject contains "You sent money with Zelle":
       - Merchant: "Zelle".
       - Category: "Zelle".
       - Description: Construct a string with this exact format: "From: [Account Info] To: [Recipient Name/Number] Memo: [Memo text]".

    ### STANDARD CATEGORIES (Only use if above rules don't apply):
    1. "Dining": 
       - Specific brands: EUREST, EATFUTI, PANDA, STARBUCKS, UBER EATS, MCDONALDS, CHICK-FIL-A, APPLEBEES.
       - KEYWORDS: "GRILL", "CAFE", "KITCHEN", "BURGER", "PIZZA", "TACO", "BAR", "RESTAURANT", "DINER".
    
    2. "General Shopping": 
       - TARGET, COSTCO, DOLLAR TREE, ROSS, MARSHALLS, AMAZON, kohl's.
    
    3. "Gas": 
       - SHELL, KUM&GO, MAVERIK, BP, EXXON, CASEYS.
    
    4. "Tech/Services": 
       - OPENAI, VERCEL, GOOGLE, NETFLIX, APPLE, SPOTIFY, UDEMY.
    
    5. "Home/Projects": 
       - HOME DEPOT, LOWES, MENARDS, HARDWARE STORES.
    
    6. "Pets": 
       - PETCO, PETSMART, VET.
    
    7. "Groceries": 
       - WALMART, SAMSCLUB, CARMQUINT FOOD & SERVI, ALDI, WHOLE FOODS, PUBLIX, TRADER JOES, HY-VEE.

    GENERAL RULE: Use the merchant name to infer the category if it's not explicitly listed. 
    
    Expected Output: ONLY a valid JSON object (no markdown, no code blocks) with these keys:
    {
      "date": "YYYY-MM-DD" (use email date if not found in text: ${emailDate}),
      "merchant": "Clean merchant name",
      "amount": "Number only (float)",
      "category": "One of the categories above",
      "description": "Brief description based on rules above"
    }

    Email Context:
    Subject: ${subject}
    Body: ${emailBody.substring(0, 2000)} 
  `;

  const payload = {
    contents: [
      {
        parts: [{ text: prompt }],
      },
    ],
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() !== 200) {
    throw new Error(
      `API Error (${response.getResponseCode()}): ${response.getContentText()}`
    );
  }

  const jsonResponse = JSON.parse(response.getContentText());

  if (!jsonResponse.candidates || jsonResponse.candidates.length === 0) {
    throw new Error(
      "Gemini did not return any results (Candidates array empty)."
    );
  }

  let rawText = jsonResponse.candidates[0].content.parts[0].text;
  rawText = rawText
    .replace(/```json/g, "")
    .replace(/```/g, "")
    .trim();

  return JSON.parse(rawText);
}

function createDashboard() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const dataSheet = ss.getSheetByName("Expenses");

  // 1. LIMPIEZA Y SETUP
  let dashSheet = ss.getSheetByName("Dashboard");
  if (dashSheet) {
    dashSheet.clear();
  } else {
    dashSheet = ss.insertSheet("Dashboard");
  }

  dashSheet.setHiddenGridlines(true);

  // 2. TITLES AND UI
  dashSheet
    .getRange("B2")
    .setValue("Expense Explorer")
    .setFontSize(18)
    .setFontWeight("bold");
  dashSheet
    .getRange("B4")
    .setValue("Filter by Month:")
    .setFontWeight("bold")
    .setFontColor("#666");
  dashSheet
    .getRange("D4")
    .setValue("Filter by Category:")
    .setFontWeight("bold")
    .setFontColor("#666");
  dashSheet
    .getRange("F4")
    .setValue("TOTAL SPENT:")
    .setFontWeight("bold")
    .setFontColor("#666");

  // 3. PREPARE DATA FOR THE SELECTORS
  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return;
  const dataValues = dataSheet.getRange(2, 1, lastRow - 1, 3).getValues();

  // Unique Categories (+ 'ALL')
  const categories = [...new Set(dataValues.map((r) => r[2]))]
    .filter(String)
    .sort();
  if (!categories.includes("ALL")) categories.unshift("ALL");

  // Uniques Months (+ 'ALL')
  const months = [
    ...new Set(
      dataValues.map((r) => {
        const d = new Date(r[0]);
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        return isNaN(d) ? "" : `${yyyy}-${mm}`;
      })
    ),
  ]
    .filter(String)
    .sort()
    .reverse();
  if (!months.includes("ALL")) months.unshift("ALL");

  // 4. DROPDOWNS
  const cellMonth = dashSheet.getRange("B5");
  cellMonth.clearDataValidations();
  cellMonth.setNumberFormat("@").setBackground("#FFF2CC").setValue("ALL");
  const ruleMonths = SpreadsheetApp.newDataValidation()
    .requireValueInList(months, true)
    .setAllowInvalid(false)
    .build();
  cellMonth.setDataValidation(ruleMonths);

  const cellCat = dashSheet.getRange("D5");
  cellCat.clearDataValidations();
  cellCat.setNumberFormat("@").setBackground("#E2F0CB").setValue("ALL");
  const ruleCats = SpreadsheetApp.newDataValidation()
    .requireValueInList(categories, true)
    .setAllowInvalid(false)
    .build();
  cellCat.setDataValidation(ruleCats);

  // 5. SUM OF TOTALS (Dynamic)
  dashSheet.getRange("F5").setFormula("=SUM(D9:D)");
  dashSheet
    .getRange("F5")
    .setNumberFormat("$#,##0.00")
    .setFontSize(14)
    .setFontWeight("bold");

  // FILTER LOGIC (WHERE CLAUSE)
  const whereClause = `
    WHERE D IS NOT NULL 
    " & IF(OR(B5="", B5="ALL"), "", " AND year(A) = " & LEFT(B5,4) & " AND month(A) = " & (RIGHT(B5,2)-1) ) & "
    " & IF(OR(D5="", D5="ALL"), "", " AND C = '" & D5 & "' ") & "
  `;

  // 6. MAIN TABLE
  const mainQuery = `=QUERY(Expenses!A:E, "SELECT A, B, C, D, E ${whereClause} ORDER BY A DESC", 1)`;
  dashSheet.getRange("A8").setFormula(mainQuery);

  dashSheet
    .getRange("A8:E8")
    .setBackground("#4285F4")
    .setFontColor("white")
    .setFontWeight("bold");
  dashSheet.getRange("A:E").setHorizontalAlignment("left");
  dashSheet.getRange("D:D").setNumberFormat("$#,##0.00");
  dashSheet.getRange("A:A").setNumberFormat("yyyy-mm-dd");
  dashSheet.setColumnWidth(1, 90);
  dashSheet.setColumnWidth(2, 160);
  dashSheet.setColumnWidth(5, 200);

  // 7. AUXILIARY CHART TABLE
  const chartQuery = `=QUERY(Expenses!A:E, "SELECT C, SUM(D) ${whereClause} GROUP BY C LABEL C 'Category', SUM(D) 'Amount'", 1)`;
  dashSheet.getRange("H8").setFormula(chartQuery);

  dashSheet
    .getRange("H8:I8")
    .setBackground("#34A853")
    .setFontColor("white")
    .setFontWeight("bold");
  dashSheet.getRange("I:I").setNumberFormat("$#,##0.00");

  // 8. BAR CHART
  const chartRange = dashSheet.getRange("H8:I20");

  const barChart = dashSheet
    .newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(chartRange)
    .setPosition(2, 8, 0, 0)
    .setOption("title", "Spending Analysis")
    .setOption("legend", { position: "none" })
    .setOption("hAxis", { format: "$#,##0" })
    .setOption("width", 600)
    .setOption("height", 350)
    .setOption("colors", ["#4285F4"])
    .build();

  dashSheet.insertChart(barChart);

  console.log("Dashboard UI Updated: Added 'ALL' options.");
}
