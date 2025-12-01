// ╔════════════════════════════════════════════════════════════════════════════╗
// ║                    DC CONSULTING ACCOUNTING SYSTEM v3.0                     ║
// ║                              Part 3 of 9                                    ║
// ║                    Clients, Vendors, Employees Databases                    ║
// ╚════════════════════════════════════════════════════════════════════════════╝

// ==================== 1. CLIENTS SHEET ====================
function createClientsSheet(ss) {
  let sheet = ss.getSheetByName('Clients');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Clients');
  sheet.setTabColor('#4caf50');
  
  const headers = [
    'Client Code',           // A
    'Company Name (EN)',     // B
    'Company Name (AR)',     // C
    'Company Name (TR)',     // D
    'Tax Number',            // E
    'Tax Office',            // F
    'Address',               // G
    'Phone',                 // H
    'Email',                 // I
    'Contact Person',        // J
    'Monthly Fee',           // K
    'Fee Currency',          // L
    'Language',              // M
    'Folder ID',             // N
    'Contract Start',        // O
    'Status',                // P
    'Notes',                 // Q
    'Created Date'           // R
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(COLORS.header)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const widths = [100, 180, 150, 180, 120, 120, 250, 120, 200, 150, 100, 80, 70, 280, 100, 80, 200, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  const lastRow = 500;
  
  // Data validations
  const currencyValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CURRENCIES, true)
    .build();
  sheet.getRange(2, 12, lastRow, 1).setDataValidation(currencyValidation);
  
  const languageValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['EN', 'AR', 'TR'], true)
    .build();
  sheet.getRange(2, 13, lastRow, 1).setDataValidation(languageValidation);
  
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'Suspended'], true)
    .build();
  sheet.getRange(2, 16, lastRow, 1).setDataValidation(statusValidation);
  
  // Number formats
  sheet.getRange(2, 11, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 15, lastRow, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 18, lastRow, 1).setNumberFormat('yyyy-mm-dd');
  
  // Conditional formatting for Status
  const statusRange = sheet.getRange(2, 16, lastRow, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Active').setBackground(COLORS.success).setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Inactive').setBackground(COLORS.warning).setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Suspended').setBackground(COLORS.danger).setRanges([statusRange]).build()
  ]);
  
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  
  // Add notes
  sheet.getRange('A1').setNote('Client Code: Auto-generated (CLT-001, CLT-002, ...)');
  sheet.getRange('N1').setNote('Folder ID: Google Drive folder for invoices');
  
  return sheet;
}

function addNewClient() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Clients');
  
  if (!sheet) {
    ui.alert('⚠️ Clients sheet not found!\n\nRun "Setup System" first.');
    return;
  }
  
  const lastRow = sheet.getLastRow() + 1;
  const newCode = generateNextCode('CLT', sheet, 1);
  
  // Set defaults
  sheet.getRange(lastRow, 1).setValue(newCode);
  sheet.getRange(lastRow, 12).setValue('TRY');
  sheet.getRange(lastRow, 13).setValue('AR');
  sheet.getRange(lastRow, 16).setValue('Active');
  sheet.getRange(lastRow, 18).setValue(new Date());
  
  sheet.setActiveRange(sheet.getRange(lastRow, 2));
  ss.setActiveSheet(sheet);
  
  ui.alert(
    '👤 Add New Client (إضافة عميل جديد)\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Client Code: ' + newCode + '\n' +
    'Row: ' + lastRow + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'Required fields (الحقول المطلوبة):\n' +
    '• Company Name (EN/AR/TR)\n' +
    '• Tax Number\n' +
    '• Email\n' +
    '• Monthly Fee\n' +
    '• Folder ID (for invoices)'
  );
}

function getClientData(clientCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Clients');
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const cols = {};
  headers.forEach((h, i) => cols[h] = i);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][cols['Client Code']] === clientCode) {
      return {
        row: i + 1,
        code: data[i][cols['Client Code']],
        nameEN: data[i][cols['Company Name (EN)']] || '',
        nameAR: data[i][cols['Company Name (AR)']] || '',
        nameTR: data[i][cols['Company Name (TR)']] || '',
        taxNumber: data[i][cols['Tax Number']] || '',
        taxOffice: data[i][cols['Tax Office']] || '',
        address: data[i][cols['Address']] || '',
        phone: data[i][cols['Phone']] || '',
        email: data[i][cols['Email']] || '',
        contactPerson: data[i][cols['Contact Person']] || '',
        monthlyFee: data[i][cols['Monthly Fee']] || 0,
        feeCurrency: data[i][cols['Fee Currency']] || 'TRY',
        language: data[i][cols['Language']] || 'AR',
        folderId: data[i][cols['Folder ID']] || '',
        contractStart: data[i][cols['Contract Start']] || '',
        status: data[i][cols['Status']] || 'Active',
        notes: data[i][cols['Notes']] || ''
      };
    }
  }
  return null;
}

function getActiveClients() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Clients');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const cols = {};
  headers.forEach((h, i) => cols[h] = i);
  
  const clients = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][cols['Status']] === 'Active' && data[i][cols['Company Name (EN)']]) {
      clients.push({
        row: i + 1,
        code: data[i][cols['Client Code']],
        nameEN: data[i][cols['Company Name (EN)']],
        nameAR: data[i][cols['Company Name (AR)']],
        nameTR: data[i][cols['Company Name (TR)']],
        monthlyFee: data[i][cols['Monthly Fee']] || 0,
        feeCurrency: data[i][cols['Fee Currency']] || 'TRY',
        email: data[i][cols['Email']] || '',
        folderId: data[i][cols['Folder ID']] || '',
        language: data[i][cols['Language']] || 'AR',
        display: data[i][cols['Company Name (EN)']] + ' (' + data[i][cols['Company Name (AR)']] + ')'
      });
    }
  }
  return clients;
}

function getClientByName(clientName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Clients');
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    // Check if name matches EN, AR, or TR
    if (data[i][1] === clientName || data[i][2] === clientName || data[i][3] === clientName) {
      return {
        code: data[i][0],
        nameEN: data[i][1],
        nameAR: data[i][2],
        nameTR: data[i][3]
      };
    }
  }
  return null;
}

// ==================== 2. VENDORS SHEET ====================
function createVendorsSheet(ss) {
  let sheet = ss.getSheetByName('Vendors');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Vendors');
  sheet.setTabColor('#ff9800');
  
  const headers = [
    'Vendor Code',           // A
    'Vendor Name (EN)',      // B
    'Vendor Name (AR)',      // C
    'Vendor Name (TR)',      // D
    'Tax Number',            // E
    'Tax Office',            // F
    'Address',               // G
    'Phone',                 // H
    'Email',                 // I
    'Contact Person',        // J
    'Category',              // K
    'Payment Terms',         // L
    'Currency',              // M
    'Bank Name',             // N
    'IBAN',                  // O
    'Status',                // P
    'Notes',                 // Q
    'Created Date'           // R
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#e65100')
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const widths = [100, 180, 150, 180, 120, 120, 250, 120, 200, 150, 120, 100, 80, 150, 250, 80, 200, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  const lastRow = 500;
  
  // Category validation
  const categoryValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Rent', 'Utilities', 'Services', 'Supplies', 'Government', 'Insurance', 'Other'], true)
    .build();
  sheet.getRange(2, 11, lastRow, 1).setDataValidation(categoryValidation);
  
  // Currency validation
  const currencyValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CURRENCIES, true)
    .build();
  sheet.getRange(2, 13, lastRow, 1).setDataValidation(currencyValidation);
  
  // Status validation
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive'], true)
    .build();
  sheet.getRange(2, 16, lastRow, 1).setDataValidation(statusValidation);
  
  sheet.getRange(2, 18, lastRow, 1).setNumberFormat('yyyy-mm-dd');
  sheet.setFrozenRows(1);
  
  return sheet;
}

function addNewVendor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Vendors');
  
  if (!sheet) {
    ui.alert('⚠️ Vendors sheet not found!');
    return;
  }
  
  const lastRow = sheet.getLastRow() + 1;
  const newCode = generateNextCode('VND', sheet, 1);
  
  sheet.getRange(lastRow, 1).setValue(newCode);
  sheet.getRange(lastRow, 13).setValue('TRY');
  sheet.getRange(lastRow, 16).setValue('Active');
  sheet.getRange(lastRow, 18).setValue(new Date());
  
  sheet.setActiveRange(sheet.getRange(lastRow, 2));
  ss.setActiveSheet(sheet);
  
  ui.alert(
    '🏪 Add New Vendor (إضافة مورد جديد)\n\n' +
    'Vendor Code: ' + newCode + '\n' +
    'Row: ' + lastRow
  );
}

function getActiveVendors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Vendors');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const vendors = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][15] === 'Active' && data[i][1]) {
      vendors.push({
        code: data[i][0],
        nameEN: data[i][1],
        nameAR: data[i][2],
        nameTR: data[i][3],
        display: data[i][1] + ' (' + data[i][2] + ')'
      });
    }
  }
  return vendors;
}

// ==================== 3. EMPLOYEES SHEET ====================
function createEmployeesSheet(ss) {
  let sheet = ss.getSheetByName('Employees');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Employees');
  sheet.setTabColor('#9c27b0');
  
  const headers = [
    'Employee Code',         // A
    'Full Name (EN)',        // B
    'Full Name (AR)',        // C
    'Full Name (TR)',        // D
    'National ID',           // E
    'Phone',                 // F
    'Email',                 // G
    'Position',              // H
    'Department',            // I
    'Start Date',            // J
    'Salary',                // K
    'Currency',              // L
    'Bank Name',             // M
    'IBAN',                  // N
    'Status',                // O
    'Notes',                 // P
    'Created Date'           // Q
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#6a1b9a')
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const widths = [100, 160, 140, 160, 120, 120, 200, 150, 120, 100, 100, 80, 150, 250, 80, 200, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  const lastRow = 200;
  
  // Currency validation
  const currencyValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CURRENCIES, true)
    .build();
  sheet.getRange(2, 12, lastRow, 1).setDataValidation(currencyValidation);
  
  // Status validation
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'On Leave'], true)
    .build();
  sheet.getRange(2, 15, lastRow, 1).setDataValidation(statusValidation);
  
  // Number formats
  sheet.getRange(2, 10, lastRow, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 11, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 17, lastRow, 1).setNumberFormat('yyyy-mm-dd');
  
  sheet.setFrozenRows(1);
  
  return sheet;
}

function addNewEmployee() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Employees');
  
  if (!sheet) {
    ui.alert('⚠️ Employees sheet not found!');
    return;
  }
  
  const lastRow = sheet.getLastRow() + 1;
  const newCode = generateNextCode('EMP', sheet, 1);
  
  sheet.getRange(lastRow, 1).setValue(newCode);
  sheet.getRange(lastRow, 12).setValue('TRY');
  sheet.getRange(lastRow, 15).setValue('Active');
  sheet.getRange(lastRow, 17).setValue(new Date());
  
  sheet.setActiveRange(sheet.getRange(lastRow, 2));
  ss.setActiveSheet(sheet);
  
  ui.alert(
    '👨‍💼 Add New Employee (إضافة موظف جديد)\n\n' +
    'Employee Code: ' + newCode + '\n' +
    'Row: ' + lastRow
  );
}

function getActiveEmployees() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Employees');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const employees = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][14] === 'Active' && data[i][1]) {
      employees.push({
        code: data[i][0],
        nameEN: data[i][1],
        nameAR: data[i][2],
        nameTR: data[i][3],
        display: data[i][1] + ' (' + data[i][2] + ')'
      });
    }
  }
  return employees;
}

// ==================== 4. GET PARTY BY TYPE ====================
/**
 * Get list of parties based on type - for dynamic Party Name dropdown
 * @param {string} partyType - 'Client', 'Vendor', 'Employee', 'Internal'
 * @returns {Array} - List of party names for dropdown
 */
function getPartyListByType(partyType) {
  switch (partyType) {
    case 'Client (عميل)':
    case 'Client':
      return getActiveClients().map(c => c.display);
    
    case 'Vendor (مورد)':
    case 'Vendor':
      return getActiveVendors().map(v => v.display);
    
    case 'Employee (موظف)':
    case 'Employee':
      return getActiveEmployees().map(e => e.display);
    
    case 'Internal (داخلي)':
    case 'Internal':
      // Return cash boxes and bank accounts for internal transfers
      const cashBanks = [];
      const cashBoxes = getCashBoxesList();
      const bankAccounts = getBankAccountsList();
      
      cashBoxes.forEach(c => cashBanks.push('💰 ' + c.name));
      bankAccounts.forEach(b => cashBanks.push('🏦 ' + b.name));
      
      return cashBanks;
    
    default:
      return [];
  }
}

// ==================== CLIENT UTILITIES ====================

/**
 * توليد أكواد تلقائية للعملاء الذين ليس لديهم كود
 */
function generateMissingClientCodes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Clients');
  const ui = SpreadsheetApp.getUi();
  
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('❌ No clients found!');
    return;
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  let fixed = 0;
  
  for (let i = 0; i < data.length; i++) {
    const code = data[i][0];
    const name = data[i][1];
    
    if (name && !code) {
      const newCode = 'CLI-' + String(i + 1).padStart(3, '0');
      sheet.getRange(i + 2, 1).setValue(newCode);
      fixed++;
    }
  }
  
  // Refresh dropdowns
  try {
    refreshClientDropdowns();
  } catch(e) {}
  
  ui.alert('✅ Generated ' + fixed + ' client codes!\n\nDropdowns updated.');
}

// ==================== END OF PART 3 ====================
