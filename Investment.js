
/** * GLOBAL CONFIGURATION */
const INVESTMENT_CONFIG = {  
  "SBI Automotive Opportunities Fund Growth": 3000,  
  "SBI Mid Cap Fund Growth": 3000,  
  "Nippon India Small Cap Fund Direct Growth": 1000,  
  "Public Provident Fund": 12500
};  

/** * Adds a custom menu to the toolbar. */
function onOpen() { 
  const ui = SpreadsheetApp.getUi();  
  ui.createMenu('Portfolio Actions') 
    .addItem('Recalculate All Totals', 'recalculateSheetTotals') 
    .addItem('Sync NAV & Units to Transactions', 'syncNavToTransactions') // New Menu Item
    .addToUi();
}  

/** * Appends new monthly entries for all investments in the specified order. */
function appendMonthlyEntries() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const sheet = ss.getSheetByName("Transactions"); 
  const liveNavSheet = ss.getSheetByName("Live NAV");

  if (!sheet) { 
    ss.toast("Sheet 'Transactions' not found!", "Error"); 
    return; 
  }  
  
  // 1. Fetch current NAVs from 'Live NAV' sheet
  let navMap = {};
  if (liveNavSheet) {
    const navData = liveNavSheet.getDataRange().getValues();
    for (let i = 1; i < navData.length; i++) {
      let name = navData[i][0]; // Column A: Investment Name
      let nav = parseFloat(navData[i][2]); // Column C: Value (NAV/Rate)
      if (name && !isNaN(nav)) {
        navMap[name] = nav;
      }
    }
  }

  // Helper function to match names and return NAV
  function getNav(invName) {
    for (let key in navMap) {
      if (invName.includes(key) && key !== "PPF Interest Rate") {
        return navMap[key];
      }
    }
    return "";
  }

  const today = new Date();  
  
  // 2. Get the NAV for each Mutual Fund
  let sbiAutoNav = getNav("SBI Automotive Opportunities Fund Growth");
  let sbiMidNav = getNav("SBI Mid Cap Fund Growth");
  let nipponNav = getNav("Nippon India Small Cap Fund Direct Growth");

  // 3. Prepare the data rows (8 columns wide to include Transaction Type)
  // We leave "Current Units" blank here because recalculateSheetTotals will calculate it cumulatively!
  const newEntries = [ 
    [today, "PPF", "Public Provident Fund", INVESTMENT_CONFIG["Public Provident Fund"], "", "", "", "Deposit"], 
    [today, "SIP", "SBI Automotive Opportunities Fund Growth", INVESTMENT_CONFIG["SBI Automotive Opportunities Fund Growth"], "", sbiAutoNav, "", "Deposit"], 
    [today, "SIP", "SBI Mid Cap Fund Growth", INVESTMENT_CONFIG["SBI Mid Cap Fund Growth"], "", sbiMidNav, "", "Deposit"], 
    [today, "SIP", "Nippon India Small Cap Fund Direct Growth", INVESTMENT_CONFIG["Nippon India Small Cap Fund Direct Growth"], "", nipponNav, "", "Deposit"] 
  ];  
  
  // Append the rows to the bottom of the table
  sheet.getRange(sheet.getLastRow() + 1, 1, newEntries.length, 8).setValues(newEntries);  
  
  // Recalculate totals (This will now handle Cumulative Units and Invested Amount)
  recalculateSheetTotals();  
  ss.toast("Monthly entries appended successfully.", "Success");
}  

/** * NEW: Syncs missing NAV, Units, and Transaction Types for existing rows */
function syncNavToTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName("Transactions");
  const liveNavSheet = ss.getSheetByName("Live NAV");

  if (!transSheet || !liveNavSheet) return;

  // Build NAV Map from Live NAV sheet
  let navMap = {};
  const navData = liveNavSheet.getDataRange().getValues();
  for (let i = 1; i < navData.length; i++) {
    let name = navData[i][0];
    let nav = parseFloat(navData[i][2]);
    if (name && !isNaN(nav)) navMap[name] = nav;
  }

  const lastRow = transSheet.getLastRow();
  if (lastRow > 1) {
    const range = transSheet.getRange(2, 1, lastRow - 1, 8);
    const data = range.getValues();
    let isUpdated = false;

    for (let i = 0; i < data.length; i++) {
      const invName = data[i][2]; // Column C
      const currentNav = data[i][5]; // Column F
      const transType = data[i][7]; // Column H

      // 1. Set Transaction Type to Deposit if empty
      if (transType === "") {
        data[i][7] = "Deposit";
        isUpdated = true;
      }

      // 2. Update NAV if NAV is empty
      if (invName && currentNav === "") {
        for (let key in navMap) {
          if (invName.includes(key) && key !== "PPF Interest Rate") {
            data[i][5] = navMap[key]; // Set NAV
            isUpdated = true;
            break;
          }
        }
      }
    }

    // Write NAVs and Deposit tags back to sheet
    if (isUpdated) {
      range.setValues(data);
    }
  }
  
  // Always recalculate at the end so units are updated cumulatively
  recalculateSheetTotals();
}

/** * CORE LOGIC: Recalculates 'Invested Amount' AND 'Current Units' cumulatively. */
function recalculateSheetTotals() {  
  const ss = SpreadsheetApp.getActiveSpreadsheet();  
  const sheet = ss.getSheetByName("Transactions");  
  if (!sheet) return;  
  
  const range = sheet.getDataRange();  
  const data = range.getValues();  
  const totals = {};  
  const unitTotals = {}; // NEW: Track cumulative units
  
  for (let i = 1; i < data.length; i++) {  
    const name = data[i][2]; // Column C 
    const sip = Number(data[i][3]) || 0; // Column D  
    const nav = Number(data[i][5]) || 0; // Column F
    
    if (name) {  
      // 1. Cumulative Invested Amount
      totals[name] = (totals[name] || 0) + sip;  
      data[i][6] = totals[name]; // Update Column G  
      
      // 2. Cumulative Units (Added from last purchased unit)
      if (nav > 0) {
        const transUnits = sip / nav; // Units for this specific transaction
        unitTotals[name] = (unitTotals[name] || 0) + transUnits; // Add to cumulative total
        data[i][4] = Number(unitTotals[name].toFixed(3)); // <--- CHANGED: Rounds to 3 decimal places
      }
    }  
  }  
  
  range.setValues(data);  
  ss.toast("Totals and Cumulative Units updated successfully.", "Portfolio Sync", 1);
}  

/** * MANUAL TRIGGER: Handles cell edits (Value changes). */
function onEdit(e) {  
  const sheet = e.source.getActiveSheet();  
  const range = e.range;  
  if (sheet.getName() === "Transactions" && range.getRow() >= 2) {  
    const col = range.getColumn();  
    if (col === 3 || col === 4 || col === 6) {  
      // Auto-fill SIP if name is entered 
      if (col === 3 && range.getValue() && INVESTMENT_CONFIG[range.getValue()]) {  
        sheet.getRange(range.getRow(), 4).setValue(INVESTMENT_CONFIG[range.getValue()]);  
      }  
      recalculateSheetTotals();  
    }  
  }
}  

/** * INSTALLABLE TRIGGER: Handles structural changes like deleting rows. */
function handleChange(e) {  
  if (e.changeType === "REMOVE_ROW") {  
    recalculateSheetTotals();  
  }
} 
