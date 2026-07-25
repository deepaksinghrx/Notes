
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
    .addItem('Recalculate PPF Interset', 'updatePPFInterestTable')
    .addToUi();
}

/** * Appends new monthly entries for all investments in the specified order. */
function appendMonthlyEntries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Transactions");
  if (!sheet) {
    ss.toast("Sheet 'Transactions' not found!", "Error");
    return;
  }

  const today = new Date();
  
  // Prepare the data rows in the requested order:
  // Order: PPF, SBI Automotive, SBI Mid, Nippon
  // Format: [Date, Investment Type, Investment Name, Monthly SIP / Put, Current Units, Avg NAV/Price, Invested Amount]
  const newEntries = [
    [today, "PPF", "Public Provident Fund", INVESTMENT_CONFIG["Public Provident Fund"], "", "", ""],
    [today, "SIP", "SBI Automotive Opportunities Fund Growth", INVESTMENT_CONFIG["SBI Automotive Opportunities Fund Growth"], "", "", ""],
    [today, "SIP", "SBI Mid Cap Fund Growth", INVESTMENT_CONFIG["SBI Mid Cap Fund Growth"], "", "", ""],
    [today, "SIP", "Nippon India Small Cap Fund Direct Growth", INVESTMENT_CONFIG["Nippon India Small Cap Fund Direct Growth"], "", "", ""]
  ];

  // Append the rows to the bottom of the table
  sheet.getRange(sheet.getLastRow() + 1, 1, newEntries.length, 7).setValues(newEntries);

  // Recalculate totals to update Column G for the new rows
  recalculateSheetTotals();
  
  ss.toast("Monthly entries appended successfully.", "Success");
}

/** * CORE LOGIC: Recalculates every 'Invested Amount' in the sheet. */
function recalculateSheetTotals() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const sheet = ss.getSheetByName("Transactions"); 
  if (!sheet) return;

  const range = sheet.getDataRange(); 
  const data = range.getValues(); 
  const totals = {};

  for (let i = 1; i < data.length; i++) { 
    const name = data[i][2]; // Column C
    const sip = Number(data[i][3] || 0); // Column D 
    if (name) { 
      totals[name] = (totals[name] || 0) + sip; 
      data[i][6] = totals[name]; // Update Column G 
    } 
  }

  range.setValues(data); 
  ss.toast("Totals updated successfully.", "Portfolio Sync", 1);
}

/** * MANUAL TRIGGER: Handles cell edits (Value changes). */
function onEdit(e) { 
  const sheet = e.source.getActiveSheet(); 
  const range = e.range; 
  if (sheet.getName() === "Transactions" && range.getRow() >= 2) { 
    const col = range.getColumn(); 
    if (col === 3 || col === 4) { 
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

