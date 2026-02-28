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
    'Company Type',          // E
    'Tax Number',            // F
    'Tax Office',            // G
    'Address',               // H
    'Phone',                 // I
    'Email',                 // J
    'Contact Person',        // K
    'Language',              // L
    'Folder ID',             // M
    'Contract Start',        // N
    'Status',                // O
    'Notes',                 // P
    'Created Date'           // Q
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(COLORS.header)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const widths = [100, 180, 150, 180, 120, 120, 120, 250, 120, 200, 150, 70, 280, 100, 80, 200, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  const lastRow = 500;

  // Data validations
  // Company Type validation (column E)
  const companyTypeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Limited', 'Şahıs', 'Anonim', 'Mükellef'], true)
    .build();
  sheet.getRange(2, 5, lastRow, 1).setDataValidation(companyTypeValidation);

  // Language validation (column L = 12)
  const languageValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['EN', 'AR', 'TR'], true)
    .build();
  sheet.getRange(2, 12, lastRow, 1).setDataValidation(languageValidation);

  // Status validation (column O = 15)
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'Suspended'], true)
    .build();
  sheet.getRange(2, 15, lastRow, 1).setDataValidation(statusValidation);

  // Number formats
  sheet.getRange(2, 14, lastRow, 1).setNumberFormat('dd.mm.yyyy'); // Contract Start
  sheet.getRange(2, 17, lastRow, 1).setNumberFormat('dd.mm.yyyy'); // Created Date

  // Conditional formatting for Status (column O = 15)
  const statusRange = sheet.getRange(2, 15, lastRow, 1);
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
  sheet.getRange('M1').setNote('Folder ID: Google Drive folder for invoices');
  
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
  sheet.getRange(lastRow, 5).setValue('Limited'); // Company Type
  sheet.getRange(lastRow, 12).setValue('AR'); // Language
  sheet.getRange(lastRow, 15).setValue('Active'); // Status
  sheet.getRange(lastRow, 17).setValue(new Date()); // Created Date

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
    '• Folder ID (for invoices)\n\n' +
    'Then add activities in "Client Activities" sheet'
  );
}

/**
 * توليد الأكواد المفقودة للعملاء
 * Generate missing codes for clients that have data but no code
 */
function generateMissingClientCodes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Clients');

  if (!sheet) {
    ui.alert('⚠️ Clients sheet not found!');
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    ui.alert('ℹ️ No client data found.');
    return;
  }

  let generatedCount = 0;
  const codeCol = 1; // Column A
  const nameCol = 2; // Column B (Company Name EN)

  for (let i = 1; i < data.length; i++) {
    const code = data[i][codeCol - 1];
    const name = data[i][nameCol - 1];

    // If no code but has name, generate code
    if ((!code || code.toString().trim() === '') && name && name.toString().trim() !== '') {
      const newCode = generateNextCode('CLT', sheet, codeCol);
      sheet.getRange(i + 1, codeCol).setValue(newCode);
      generatedCount++;

      // Update data array to avoid duplicate codes
      data[i][codeCol - 1] = newCode;
    }
  }

  if (generatedCount > 0) {
    ui.alert(
      '✅ تم توليد الأكواد بنجاح!\n\n' +
      'Generated ' + generatedCount + ' client code(s).\n\n' +
      'تم توليد ' + generatedCount + ' كود للعملاء.'
    );
  } else {
    ui.alert('ℹ️ All clients already have codes.\n\nجميع العملاء لديهم أكواد بالفعل.');
  }
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
        companyType: data[i][cols['Company Type']] || '',
        taxNumber: data[i][cols['Tax Number']] || '',
        taxOffice: data[i][cols['Tax Office']] || '',
        address: data[i][cols['Address']] || '',
        phone: data[i][cols['Phone']] || '',
        email: data[i][cols['Email']] || '',
        contactPerson: data[i][cols['Contact Person']] || '',
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
    'Sector',                // L - Which business sector this vendor serves
    'Payment Terms',         // M
    'Currency',              // N
    'Bank Name',             // O
    'IBAN',                  // P
    'Status',                // Q
    'Notes',                 // R
    'Created Date'           // S
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#e65100')
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const widths = [100, 180, 150, 180, 120, 120, 250, 120, 200, 150, 120, 120, 100, 80, 150, 250, 80, 200, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  const lastRow = 500;

  // Category validation (column K = 11)
  const categoryValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Rent', 'Utilities', 'Services', 'Supplies', 'Government', 'Insurance', 'Other'], true)
    .build();
  sheet.getRange(2, 11, lastRow, 1).setDataValidation(categoryValidation);

  // Sector validation (column L = 12)
  const sectorValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Accounting', 'Consulting', 'Logistics', 'Trading', 'Inspection', 'Tourism', 'All', 'Other'], true)
    .build();
  sheet.getRange(2, 12, lastRow, 1).setDataValidation(sectorValidation);

  // Currency validation (column N = 14)
  const currencyValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CURRENCIES, true)
    .build();
  sheet.getRange(2, 14, lastRow, 1).setDataValidation(currencyValidation);

  // Status validation (column Q = 17)
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive'], true)
    .build();
  sheet.getRange(2, 17, lastRow, 1).setDataValidation(statusValidation);

  sheet.getRange(2, 19, lastRow, 1).setNumberFormat('dd.mm.yyyy');
  sheet.setFrozenRows(1);

  sheet.getRange('L1').setNote('Sector: القطاع الذي يخدمه المورد\nAll = يخدم كل القطاعات');
  
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
  sheet.getRange(lastRow, 14).setValue('TRY');      // Currency (col N)
  sheet.getRange(lastRow, 17).setValue('Active');    // Status (col Q)
  sheet.getRange(lastRow, 19).setValue(new Date());  // Created Date (col S)
  
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
    if (data[i][16] === 'Active' && data[i][1]) {  // Status is col Q (index 16)
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
    'Sector',                // J - Which business sector this employee works in
    'Start Date',            // K
    'Salary',                // L
    'Currency',              // M
    'Bank Name',             // N
    'IBAN',                  // O
    'Status',                // P
    'Notes',                 // Q
    'Created Date'           // R
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#6a1b9a')
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const widths = [100, 160, 140, 160, 120, 120, 200, 150, 120, 120, 100, 100, 80, 150, 250, 80, 200, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  const lastRow = 200;

  // Sector validation (column J = 10)
  const sectorValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Accounting', 'Consulting', 'Logistics', 'Trading', 'Inspection', 'Tourism', 'All', 'Other'], true)
    .build();
  sheet.getRange(2, 10, lastRow, 1).setDataValidation(sectorValidation);

  // Currency validation (column M = 13)
  const currencyValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CURRENCIES, true)
    .build();
  sheet.getRange(2, 13, lastRow, 1).setDataValidation(currencyValidation);

  // Status validation (column P = 16)
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'On Leave'], true)
    .build();
  sheet.getRange(2, 16, lastRow, 1).setDataValidation(statusValidation);

  // Number formats
  sheet.getRange(2, 11, lastRow, 1).setNumberFormat('dd.mm.yyyy');  // Start Date (col K)
  sheet.getRange(2, 12, lastRow, 1).setNumberFormat('#,##0.00');    // Salary (col L)
  sheet.getRange(2, 18, lastRow, 1).setNumberFormat('dd.mm.yyyy');  // Created Date (col R)

  sheet.setFrozenRows(1);

  sheet.getRange('J1').setNote('Sector: القطاع الذي يعمل فيه الموظف\nAll = يعمل في كل القطاعات');

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
  sheet.getRange(lastRow, 13).setValue('TRY');      // Currency (col M)
  sheet.getRange(lastRow, 16).setValue('Active');    // Status (col P)
  sheet.getRange(lastRow, 18).setValue(new Date());  // Created Date (col R)
  
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
    if (data[i][15] === 'Active' && data[i][1]) {  // Status is col P (index 15)
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
 * إضافة عامود Company Type للشيت الموجود بدون حذف البيانات
 */
function addCompanyTypeColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Clients');

  if (!sheet) {
    ui.alert('❌ Clients sheet not found!');
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Check if Company Type column already exists
  if (headers.includes('Company Type')) {
    // Update the validation with new options
    const companyTypeCol = headers.indexOf('Company Type') + 1;
    const lastRow = Math.max(sheet.getLastRow(), 500);
    const companyTypeValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Limited', 'Şahıs', 'Anonim', 'Mükellef'], true)
      .build();
    sheet.getRange(2, companyTypeCol, lastRow, 1).setDataValidation(companyTypeValidation);

    ui.alert('✅ Company Type validation updated!\n\nOptions: Limited, Şahıs, Anonim, Mükellef');
    return;
  }

  // Find where to insert (after Company Name (TR) - column D)
  const insertAfterCol = 4; // Column D

  // Insert new column at position 5 (E)
  sheet.insertColumnAfter(insertAfterCol);

  // Set header
  sheet.getRange(1, 5).setValue('Company Type')
    .setBackground(COLORS.header)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Set column width
  sheet.setColumnWidth(5, 150);

  // Add validation
  const lastRow = Math.max(sheet.getLastRow(), 500);
  const companyTypeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Limited', 'Şahıs', 'Anonim', 'Mükellef'], true)
    .build();
  sheet.getRange(2, 5, lastRow, 1).setDataValidation(companyTypeValidation);

  // Set default value for existing clients
  const existingRows = sheet.getLastRow() - 1;
  if (existingRows > 0) {
    for (let i = 2; i <= sheet.getLastRow(); i++) {
      if (sheet.getRange(i, 2).getValue()) { // If has company name
        sheet.getRange(i, 5).setValue('Limited');
      }
    }
  }

  ui.alert(
    '✅ Company Type column added!\n\n' +
    'Şirket Türü sütunu eklendi\n\n' +
    'Options: Limited, Şahıs, Anonim, Mükellef\n' +
    'Default: Limited'
  );
}

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
