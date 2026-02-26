// ╔════════════════════════════════════════════════════════════════════════════╗
// ║                    DC CONSULTING ACCOUNTING SYSTEM v3.0                     ║
// ║                              Part 9 of 9 (FINAL)                            ║
// ║                    Utilities + User Guide + System Setup                    ║
// ╚════════════════════════════════════════════════════════════════════════════╝

// ==================== 1. SYSTEM SETUP (SECURE) ====================
function setupSystemSecure() {
  const ui = SpreadsheetApp.getUi();
  
  // First time setup doesn't require password
  const props = PropertiesService.getScriptProperties();
  const isFirstTime = !props.getProperty('SYSTEM_INITIALIZED');
  
  if (!isFirstTime) {
    if (!verifyPassword('setup system')) return;
  }
  
  const confirm = ui.alert(
    '🔐 DC Consulting System Setup\n\n' +
    'This will create all required sheets:\n\n' +
    '• Settings & Holidays\n' +
    '• Categories & Movement Types\n' +
    '• Items Database & Activities\n' +
    '• Activity Profiles (per-activity branding)\n' +
    '• Clients, Client Activities, Vendors, Employees\n' +
    '• Cash Boxes & Bank Accounts\n' +
    '• Transactions\n' +
    '• Invoice Template & Log\n' +
    '• Email Log & Alerts Log\n' +
    '• Dashboard\n\n' +
    '⚠️ Existing sheets with same names will be recreated!\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    ui.alert('🔄 Setting up system...\n\nThis may take a minute. Click OK to continue.');
    
    // Part 2: Database sheets
    createSettingsSheet(ss);
    createHolidaysSheet(ss);
    createCategoriesSheet(ss);
    createMovementTypesSheet(ss);
    createItemsDatabase(ss);
    createActivitiesSheet(ss);
    createActivityProfilesSheet(ss);

    // Part 3: Party sheets
    createClientsSheet(ss);
    createClientActivitiesSheet(ss);
    createVendorsSheet(ss);
    createEmployeesSheet(ss);
    
    // Part 4: Cash & Bank
    createCashBoxesDatabase(ss);
    createBankAccountsDatabase(ss);
    
    // Part 5: Transactions
    createTransactionsSheet(ss);
    
    // Part 6: Invoices
    createInvoiceLogSheet(ss);
    createInvoiceTemplateSheet(ss);
    
    // Part 7: Email & Alerts
    createEmailLogSheet(ss);
    createAlertsLogSheet(ss);
    
    // Part 8: Dashboard
    createDashboardSheet(ss);
    
    // Setup dropdowns
    setupTransactionDropdowns();
    
    // Mark as initialized
    props.setProperty('SYSTEM_INITIALIZED', 'true');
    props.setProperty('SETUP_DATE', new Date().toISOString());
    
    // Navigate to Transactions
    const transSheet = ss.getSheetByName('Transactions');
    if (transSheet) ss.setActiveSheet(transSheet);
    
    ui.alert(
      '✅ System Setup Complete!\n\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      'Version: ' + SYSTEM_VERSION + '\n' +
      'Default Password: DC2025\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
      'Next Steps:\n' +
      '1. Add bank accounts in "Bank Accounts"\n' +
      '2. Add cash boxes in "Cash Boxes"\n' +
      '3. Click "Create Cash/Bank Sheets"\n' +
      '4. Add clients in "Clients"\n' +
      '5. Start recording transactions!\n\n' +
      '📖 See "User Guide" for more help.'
    );
    
  } catch (error) {
    ui.alert('❌ Setup Error:\n\n' + error.message);
  }
}

// ==================== 2. USER GUIDE ====================
function showUserGuide() {
  const ui = SpreadsheetApp.getUi();
  
  const guide = `
📖 DC CONSULTING SYSTEM - USER GUIDE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🏢 SYSTEM OVERVIEW
This system manages:
• Clients, Vendors & Employees (3 languages)
• Cash Boxes & Bank Accounts (multi-currency)
• Transactions with smart dropdowns
• Invoices (3 methods + PDF + Email)
• Reports & Dashboard

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

⚙️ INITIAL SETUP
1. Run "Setup System" from menu
2. Add bank accounts → Bank Accounts sheet
3. Add cash boxes → Cash Boxes sheet
4. Click "Create Cash/Bank Sheets"
5. Add clients → Clients sheet
6. Set Folder ID for each client (for invoices)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💰 TRANSACTIONS
• Select Movement Type first
• Client Code auto-fills Client Name (and vice versa)
• Party Type changes Party Name dropdown dynamically
• Payment Method colors the row automatically
• Amount TRY calculated from Amount × Exchange Rate

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🧾 INVOICES
3 ways to create:
1. From Transaction - select row, generate
2. Custom Invoice - enter details manually
3. All Monthly - batch from Client Activities sheet

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📧 EMAIL
• Set "Send Email = Yes" in Invoice Log
• Run "Send Pending Invoices"
• Or setup triggers for automatic sending

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🔐 SECURITY
• Default password: DC2025
• Change via Settings → Change Password
• Sensitive sheets can be hidden

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Version: ${SYSTEM_VERSION}
© 2025 Dewan Consulting
`;
  
  ui.alert(guide);
}

// ==================== 3. QUICK REFERENCE ====================
function showQuickReference() {
  const ui = SpreadsheetApp.getUi();
  
  const ref = `
📋 QUICK REFERENCE

━━━ MOVEMENT TYPES ━━━
REV-DUE = Revenue Accrual (استحقاق إيراد)
REV-COL = Revenue Collection (تحصيل)
EXP-DUE = Expense Accrual (استحقاق مصروف)
EXP-PAY = Expense Payment (دفع)
TRF-CC = Cash to Cash
TRF-BB = Bank to Bank
TRF-CB = Cash to Bank (إيداع)
TRF-BC = Bank to Cash (سحب)

━━━ PAYMENT COLORS ━━━
🟡 Yellow = Accrual (لم يُدفع)
🟢 Green = Cash (نقدي)
🔵 Blue = Bank Transfer
🟣 Purple = Credit Card

━━━ SHORTCUTS ━━━
• Client Code → auto-fills Name
• Party Type → changes Party dropdown
• Amount × Rate = Amount TRY (auto)

━━━ INVOICE SCHEDULE ━━━
• Generation: Day 25 (or next working day)
• Sending: 2 working days after generation
• Skips weekends & Turkish holidays
`;
  
  ui.alert(ref);
}

// ==================== 4. VALIDATE SYSTEM ====================
function validateSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const requiredSheets = [
    'Settings', 'Holidays', 'Categories', 'Movement Types', 'Items Database', 'Activities',
    'Clients', 'Vendors', 'Employees',
    'Cash Boxes', 'Bank Accounts',
    'Transactions', 'Invoice Log', 'Invoice Template',
    'Dashboard'
  ];
  
  const missing = [];
  const found = [];
  
  requiredSheets.forEach(name => {
    if (ss.getSheetByName(name)) {
      found.push('✅ ' + name);
    } else {
      missing.push('❌ ' + name);
    }
  });
  
  // Check for cash/bank sheets
  const cashBoxes = getCashBoxesList();
  const bankAccounts = getBankAccountsList();
  
  let cashBankStatus = '\n\n━━━ Cash & Bank Sheets ━━━\n';
  cashBoxes.forEach(c => {
    cashBankStatus += (ss.getSheetByName(c.sheetName) ? '✅ ' : '❌ ') + c.sheetName + '\n';
  });
  bankAccounts.forEach(b => {
    cashBankStatus += (ss.getSheetByName(b.sheetName) ? '✅ ' : '❌ ') + b.sheetName + '\n';
  });
  
  const result = 
    '🔍 SYSTEM VALIDATION\n\n' +
    '━━━ Required Sheets ━━━\n' +
    found.join('\n') + '\n' +
    (missing.length > 0 ? '\n' + missing.join('\n') : '') +
    cashBankStatus +
    '\n\n' +
    (missing.length === 0 ? '✅ System is complete!' : '⚠️ Some sheets are missing. Run Setup again.');
  
  ui.alert(result);
}

// ==================== 5. BACKUP DATA ====================
function backupData() {
  const ui = SpreadsheetApp.getUi();
  
  const confirm = ui.alert(
    '💾 Backup Data\n\n' +
    'This will create a copy of the entire spreadsheet.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const backupName = ss.getName() + ' - Backup ' + formatDate(new Date(), 'yyyy-MM-dd HH-mm');
    
    const backup = ss.copy(backupName);
    
    ui.alert(
      '✅ Backup Created!\n\n' +
      'Name: ' + backupName + '\n' +
      'Location: Same folder as original\n\n' +
      'URL: ' + backup.getUrl()
    );
    
  } catch (error) {
    ui.alert('❌ Backup Error:\n\n' + error.message);
  }
}

// ==================== 6. EXPORT TRANSACTIONS TO CSV ====================
function exportTransactionsToCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const transSheet = ss.getSheetByName('Transactions');
  if (!transSheet || transSheet.getLastRow() < 2) {
    ui.alert('⚠️ No transactions to export!');
    return;
  }
  
  const data = transSheet.getDataRange().getValues();
  
  // Convert to CSV
  const csv = data.map(row => 
    row.map(cell => {
      if (cell === null || cell === undefined) return '';
      if (cell instanceof Date) return formatDate(cell, 'yyyy-MM-dd');
      const str = String(cell);
      if (str.includes(',') || str.includes('"') || str.includes('\n')) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    }).join(',')
  ).join('\n');
  
  // Create file
  const fileName = 'DC_Transactions_' + formatDate(new Date(), 'yyyyMMdd') + '.csv';
  const file = DriveApp.createFile(fileName, csv, MimeType.CSV);
  
  ui.alert(
    '✅ CSV Exported!\n\n' +
    'File: ' + fileName + '\n' +
    'Rows: ' + data.length + '\n\n' +
    'Download: ' + file.getUrl()
  );
}

// ==================== 7. MAINTENANCE FUNCTIONS ====================
function clearOldAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const sheet = ss.getSheetByName('Alerts Log');
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('ℹ️ No alerts to clear.');
    return;
  }
  
  const confirm = ui.alert(
    '🗑️ Clear Alerts Log\n\n' +
    'Delete all alerts older than 30 days?\n\n' +
    'Current alerts: ' + (sheet.getLastRow() - 1),
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  const data = sheet.getDataRange().getValues();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 30);
  
  let deleted = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    const alertDate = new Date(data[i][0]);
    if (alertDate < cutoff) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  
  ui.alert('✅ Cleared ' + deleted + ' old alerts.');
}

function resetInvoiceNumber() {
  const ui = SpreadsheetApp.getUi();
  
  if (!verifyPassword('reset invoice number')) return;
  
  const response = ui.prompt(
    '🔄 Reset Invoice Number\n\n' +
    'Current: ' + (getSettingValue('Next Invoice Number') || 1),
    'Enter new starting number:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const newNum = parseInt(response.getResponseText());
  if (isNaN(newNum) || newNum < 1) {
    ui.alert('⚠️ Invalid number!');
    return;
  }
  
  setSettingValue('Next Invoice Number', newNum);
  ui.alert('✅ Invoice number reset to: ' + newNum);
}

function fixTransactionFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const transSheet = ss.getSheetByName('Transactions');
  if (!transSheet) {
    ui.alert('❌ Transactions sheet not found!');
    return;
  }
  
  const lastRow = transSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('ℹ️ No data to fix.');
    return;
  }
  
  let fixed = 0;
  
  for (let row = 2; row <= lastRow; row++) {
    // Fix Amount TRY (column O)
    const amount = transSheet.getRange(row, 12).getValue();
    const rate = transSheet.getRange(row, 14).getValue() || 1;
    transSheet.getRange(row, 15).setValue(amount * rate);

    // Fix Remaining (column W)
    const paid = transSheet.getRange(row, 22).getValue() || 0;
    transSheet.getRange(row, 23).setValue(amount - paid);
    
    fixed++;
  }
  
  ui.alert('✅ Fixed formulas in ' + fixed + ' rows.');
}

function recalculateBalances() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    '🔄 Recalculate Balances\n\n' +
    'Cash/Bank balances are calculated automatically using SUMIF formulas.\n\n' +
    'If you see incorrect balances:\n' +
    '1. Check that Direction column (G) is "IN" or "OUT"\n' +
    '2. Check that Amount column (F) has numbers\n' +
    '3. The formula in B2 should be:\n' +
    '   =SUMIF(G4:G1000,"IN",F4:F1000)-SUMIF(G4:G1000,"OUT",F4:F1000)'
  );
}

// ==================== 8. ADD HOLIDAYS FOR NEW YEAR ====================
function addNewYearHolidays() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    '📅 Add Holidays for New Year',
    'Enter year (e.g., 2026):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const year = parseInt(response.getResponseText());
  if (isNaN(year) || year < 2025 || year > 2030) {
    ui.alert('⚠️ Invalid year!');
    return;
  }
  
  ui.alert(
    '📅 Add Holidays for ' + year + '\n\n' +
    'Please add holidays manually to the "Holidays" sheet.\n\n' +
    'Turkish holidays to add:\n' +
    '• Jan 1 - New Year\n' +
    '• Apr 23 - Children\'s Day\n' +
    '• May 1 - Labour Day\n' +
    '• May 19 - Youth Day\n' +
    '• Jul 15 - Democracy Day\n' +
    '• Aug 30 - Victory Day\n' +
    '• Oct 29 - Republic Day\n' +
    '• Eid al-Fitr (3 days) - check calendar\n' +
    '• Eid al-Adha (4 days) - check calendar'
  );
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Holidays');
  if (sheet) SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
}

// ==================== 9. DIAGNOSTIC INFO ====================
function showDiagnosticInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const props = PropertiesService.getScriptProperties();
  
  const info = `
🔧 DIAGNOSTIC INFORMATION

━━━ System ━━━
Version: ${SYSTEM_VERSION}
Spreadsheet ID: ${ss.getId()}
Timezone: ${Session.getScriptTimeZone()}

━━━ Properties ━━━
Initialized: ${props.getProperty('SYSTEM_INITIALIZED') || 'No'}
Setup Date: ${props.getProperty('SETUP_DATE') || 'Never'}
Password Set: ${props.getProperty('ADMIN_PASSWORD') ? 'Yes' : 'No (using default)'}

━━━ Sheets ━━━
Total Sheets: ${ss.getSheets().length}

━━━ Triggers ━━━
Active Triggers: ${ScriptApp.getProjectTriggers().length}

━━━ Quotas ━━━
Email remaining today: ${MailApp.getRemainingDailyQuota()}
`;
  
  ui.alert(info);
}

// ==================== 10. TEST EMAIL ====================
function sendTestEmail() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    '📧 Send Test Email',
    'Enter email address:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const email = response.getResponseText().trim();
  if (!email || !email.includes('@')) {
    ui.alert('⚠️ Invalid email!');
    return;
  }
  
  try {
    const companyName = getSettingValue('Company Name (EN)') || 'DC Consulting';
    
    GmailApp.sendEmail(
      email,
      'Test Email from ' + companyName,
      '',
      {
        name: companyName,
        htmlBody: '<h2>✅ Test Email Successful!</h2><p>Your DC Consulting system is configured correctly.</p>'
      }
    );
    
    ui.alert('✅ Test email sent to: ' + email);
    
  } catch (error) {
    ui.alert('❌ Error sending email:\n\n' + error.message);
  }
}

// ==================== 11. SYNC TRANSACTIONS TO CASH/BANK ====================
function syncTransactionsToCashBank() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const transSheet = ss.getSheetByName('Transactions');
  if (!transSheet || transSheet.getLastRow() < 2) {
    ui.alert('⚠️ No transactions to sync!');
    return;
  }
  
  const confirm = ui.alert(
    '🔄 Sync Transactions to Cash/Bank\n\n' +
    'This will add missing entries from Transactions to Cash/Bank sheets.\n\n' +
    '⚠️ Only transactions with:\n' +
    '• Payment Method ≠ Accrual\n' +
    '• Cash/Bank specified\n' +
    '• Status = Paid\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  const data = transSheet.getDataRange().getValues();
  let synced = 0, skipped = 0;
  
  for (let i = 1; i < data.length; i++) {
    const paymentMethod = data[i][15]; // Column P
    const cashBank = data[i][16]; // Column Q
    const status = data[i][19]; // Column T
    
    // Skip accruals
    if (!paymentMethod || paymentMethod.includes('Accrual')) {
      skipped++;
      continue;
    }
    
    // Skip if no cash/bank
    if (!cashBank) {
      skipped++;
      continue;
    }
    
    // Skip if not paid
    if (!status || !status.includes('Paid')) {
      skipped++;
      continue;
    }
    
    // Extract sheet name from dropdown value (e.g., "💰 Cash TRY - Main (TRY)" → "Cash TRY - Main")
    let sheetName = cashBank.replace(/^[💰🏦]\s*/, '').replace(/\s*\([^)]+\)$/, '');
    
    const targetSheet = ss.getSheetByName(sheetName);
    if (!targetSheet) {
      skipped++;
      continue;
    }
    
    // Check if already synced (by transaction code)
    const transCode = 'TRX-' + data[i][0];
    const targetData = targetSheet.getDataRange().getValues();
    let exists = false;
    
    for (let j = 3; j < targetData.length; j++) {
      if (targetData[j][4] === transCode) {
        exists = true;
        break;
      }
    }
    
    if (exists) {
      skipped++;
      continue;
    }
    
    // Add entry
    const movementType = data[i][3];
    const direction = (movementType && movementType.includes('Revenue')) ? 'IN' : 'OUT';

    addCashBankEntry(
      sheetName,
      data[i][1], // Date
      data[i][8] || data[i][7], // Description or Item
      data[i][17], // Reference
      data[i][9] || data[i][6], // Party Name or Client Name
      transCode,
      data[i][11], // Amount
      direction
    );
    
    synced++;
  }
  
  ui.alert('✅ Sync Complete!\n\nSynced: ' + synced + '\nSkipped: ' + skipped);
}
// ==================== DIAGNOSTIC TOOLS ====================

/**
 * تشخيص مشاكل الـ Dropdowns والشيتات
 */
function runSystemDiagnostic() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  let report = '🔍 DIAGNOSTIC REPORT\n\n';
  
  // 1. Check required sheets
  const requiredSheets = ['Transactions', 'Clients', 'Vendors', 'Employees', 'Cash Boxes', 'Bank Accounts', 'Items Database'];
  report += '━━━ Required Sheets ━━━\n';
  requiredSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    const exists = sheet ? '✅' : '❌';
    const rows = sheet ? sheet.getLastRow() : 0;
    report += exists + ' ' + name + ' (' + rows + ' rows)\n';
  });
  
  // 2. Check Clients data
  report += '\n━━━ Clients Data ━━━\n';
  const clientsSheet = ss.getSheetByName('Clients');
  if (clientsSheet && clientsSheet.getLastRow() > 1) {
    const data = clientsSheet.getRange(2, 1, Math.min(5, clientsSheet.getLastRow() - 1), 16).getValues();
    let active = 0;
    data.forEach((row, i) => {
      const code = row[0];
      const name = row[1];
      const status = row[15];
      report += (i+1) + '. ' + code + ' | ' + name + ' | Status: ' + (status || 'EMPTY') + '\n';
      if (status === 'Active') active++;
    });
    report += 'Active clients: ' + active + '\n';
  } else {
    report += '⚠️ No client data!\n';
  }
  
  // 3. Check global constants
  report += '\n━━━ Global Constants ━━━\n';
  try {
    report += 'COLORS: ' + (typeof COLORS !== 'undefined' ? '✅ Defined' : '❌ Not defined') + '\n';
  } catch(e) { report += 'COLORS: ❌ Not defined\n'; }
  
  try {
    report += 'CURRENCIES: ' + (typeof CURRENCIES !== 'undefined' ? '✅ ' + CURRENCIES.join(', ') : '❌ Not defined') + '\n';
  } catch(e) { report += 'CURRENCIES: ❌ Not defined\n'; }
  
  ui.alert(report);
}
// ==================== END OF PART 9 ====================
// ════════════════════════════════════════════════════════════════
// ║          DC CONSULTING ACCOUNTING SYSTEM v3.0 COMPLETE!      ║
// ║                        ~150 Functions                        ║
// ║                        9 Parts Total                         ║
// ════════════════════════════════════════════════════════════════

