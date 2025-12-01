// ╔════════════════════════════════════════════════════════════════════════════╗
// ║                    DC CONSULTING ACCOUNTING SYSTEM v3.0                     ║
// ║                              Part 4 of 9                                    ║
// ║                       Cash & Bank Management                                ║
// ╚════════════════════════════════════════════════════════════════════════════╝

// ==================== 1. CASH BOXES DATABASE ====================
function createCashBoxesDatabase(ss) {
  let sheet = ss.getSheetByName('Cash Boxes');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Cash Boxes');
  sheet.setTabColor('#ff5722');
  
  const headers = [
    'Cash Code',
    'Cash Name',
    'Currency',
    'Responsible Person',
    'Location',
    'Opening Balance',
    'Opening Date',
    'Status',
    'Notes',
    'Sheet Created'
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#d84315')
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold');
  
  // Sample data
  const sampleData = [
    ['CSH-001', 'Cash TRY - Main', 'TRY', 'Accountant', 'Office', 0, new Date(), 'Active', '', 'No'],
    ['CSH-002', 'Cash USD - Main', 'USD', 'Accountant', 'Office', 0, new Date(), 'Active', '', 'No'],
    ['CSH-003', 'Cash EUR - Main', 'EUR', 'Accountant', 'Office', 0, new Date(), 'Active', '', 'No']
  ];
  
  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  
  const widths = [80, 160, 80, 150, 100, 120, 100, 80, 200, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  // Validations
  const currencyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CURRENCIES, true).build();
  sheet.getRange(2, 3, 100, 1).setDataValidation(currencyRule);
  
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive'], true).build();
  sheet.getRange(2, 8, 100, 1).setDataValidation(statusRule);
  
  const sheetCreatedRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true).build();
  sheet.getRange(2, 10, 100, 1).setDataValidation(sheetCreatedRule);
  
  sheet.getRange(2, 6, 100, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 7, 100, 1).setNumberFormat('yyyy-mm-dd');
  sheet.setFrozenRows(1);
  
  return sheet;
}

function addNewCashBox() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Cash Boxes');
  
  if (!sheet) {
    ui.alert('⚠️ Cash Boxes sheet not found!');
    return;
  }
  
  // Prompt for name
  const nameResponse = ui.prompt(
    '💰 Add New Cash Box (إضافة خزينة)\n\nStep 1/3',
    'Enter Cash Box Name:\nأدخل اسم الخزينة:\n\nExample: Cash TRY - Shehata',
    ui.ButtonSet.OK_CANCEL
  );
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
  const cashName = nameResponse.getResponseText().trim();
  if (!cashName) { ui.alert('⚠️ Name cannot be empty!'); return; }
  
  // Prompt for currency
  const currencyResponse = ui.prompt(
    '💰 Add New Cash Box\n\nStep 2/3',
    'Enter Currency (العملة):\n\nOptions: TRY, USD, EUR, SAR, EGP, AED, GBP',
    ui.ButtonSet.OK_CANCEL
  );
  if (currencyResponse.getSelectedButton() !== ui.Button.OK) return;
  const currency = currencyResponse.getResponseText().trim().toUpperCase();
  if (!CURRENCIES.includes(currency)) { ui.alert('⚠️ Invalid currency!'); return; }
  
  // Prompt for opening balance
  const balanceResponse = ui.prompt(
    '💰 Add New Cash Box\n\nStep 3/3',
    'Enter Opening Balance (الرصيد الافتتاحي):\n\n(Enter 0 if empty)',
    ui.ButtonSet.OK_CANCEL
  );
  if (balanceResponse.getSelectedButton() !== ui.Button.OK) return;
  const openingBalance = parseFloat(balanceResponse.getResponseText()) || 0;
  
  // Add to database
  const lastRow = sheet.getLastRow() + 1;
  const newCode = generateNextCode('CSH', sheet, 1);
  
  sheet.getRange(lastRow, 1).setValue(newCode);
  sheet.getRange(lastRow, 2).setValue(cashName);
  sheet.getRange(lastRow, 3).setValue(currency);
  sheet.getRange(lastRow, 6).setValue(openingBalance);
  sheet.getRange(lastRow, 7).setValue(new Date());
  sheet.getRange(lastRow, 8).setValue('Active');
  sheet.getRange(lastRow, 10).setValue('No');
  
  // Ask to create sheet
  const createResponse = ui.alert(
    '✅ Cash Box Added!\n\n' +
    'Code: ' + newCode + '\n' +
    'Name: ' + cashName + '\n' +
    'Currency: ' + currency + '\n' +
    'Balance: ' + formatCurrency(openingBalance, currency) + '\n\n' +
    'Create cash sheet now?',
    ui.ButtonSet.YES_NO
  );
  
  if (createResponse === ui.Button.YES) {
    createSingleCashSheet(ss, cashName, currency, openingBalance);
    sheet.getRange(lastRow, 10).setValue('Yes');
    ui.alert('✅ Cash sheet "' + cashName + '" created!');
  }
}

function showCashBoxes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Cash Boxes');
  if (sheet) ss.setActiveSheet(sheet);
}

// ==================== 2. BANK ACCOUNTS DATABASE ====================
function createBankAccountsDatabase(ss) {
  let sheet = ss.getSheetByName('Bank Accounts');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Bank Accounts');
  sheet.setTabColor('#1565c0');
  
  const headers = [
    'Account Code',
    'Account Name',
    'Bank Name',
    'Currency',
    'IBAN',
    'SWIFT/BIC',
    'Account Holder',
    'Branch',
    'Opening Balance',
    'Opening Date',
    'Status',
    'Notes',
    'Sheet Created'
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#0d47a1')
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold');
  
  // Sample data
  const sampleData = [
    ['BNK-001', 'Kuveyt Türk - TRY', 'Kuveyt Türk', 'TRY', 'TR250020500009448735700002', 'KTEFTRIS', 'Dewan Consulting', 'Esenyurt', 0, new Date(), 'Active', '', 'No'],
    ['BNK-002', 'Kuveyt Türk - USD', 'Kuveyt Türk', 'USD', 'TR680020500009448735700101', 'KTEFTRIS', 'Dewan Consulting', 'Esenyurt', 0, new Date(), 'Active', '', 'No']
  ];
  
  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  
  const widths = [90, 160, 120, 70, 250, 100, 150, 100, 120, 100, 80, 200, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  // Validations
  const currencyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CURRENCIES, true).build();
  sheet.getRange(2, 4, 100, 1).setDataValidation(currencyRule);
  
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive'], true).build();
  sheet.getRange(2, 11, 100, 1).setDataValidation(statusRule);
  
  sheet.getRange(2, 9, 100, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 10, 100, 1).setNumberFormat('yyyy-mm-dd');
  sheet.setFrozenRows(1);
  
  return sheet;
}

function addNewBankAccount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Bank Accounts');
  
  if (!sheet) {
    ui.alert('⚠️ Bank Accounts sheet not found!');
    return;
  }
  
  const nameResponse = ui.prompt(
    '🏦 Add Bank Account (إضافة حساب بنكي)\n\nStep 1/4',
    'Enter Account Name:\n\nExample: Kuveyt Türk - TRY',
    ui.ButtonSet.OK_CANCEL
  );
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
  const accountName = nameResponse.getResponseText().trim();
  if (!accountName) { ui.alert('⚠️ Name cannot be empty!'); return; }
  
  const bankResponse = ui.prompt(
    '🏦 Add Bank Account\n\nStep 2/4',
    'Enter Bank Name:\n\nExample: Kuveyt Türk',
    ui.ButtonSet.OK_CANCEL
  );
  if (bankResponse.getSelectedButton() !== ui.Button.OK) return;
  const bankName = bankResponse.getResponseText().trim();
  
  const currencyResponse = ui.prompt(
    '🏦 Add Bank Account\n\nStep 3/4',
    'Enter Currency:\n\nOptions: TRY, USD, EUR, SAR, EGP, AED, GBP',
    ui.ButtonSet.OK_CANCEL
  );
  if (currencyResponse.getSelectedButton() !== ui.Button.OK) return;
  const currency = currencyResponse.getResponseText().trim().toUpperCase();
  if (!CURRENCIES.includes(currency)) { ui.alert('⚠️ Invalid currency!'); return; }
  
  const ibanResponse = ui.prompt(
    '🏦 Add Bank Account\n\nStep 4/4',
    'Enter IBAN:\n\nExample: TR250020500009448735700002',
    ui.ButtonSet.OK_CANCEL
  );
  if (ibanResponse.getSelectedButton() !== ui.Button.OK) return;
  const iban = ibanResponse.getResponseText().trim().replace(/\s/g, '');
  
  // Add to database
  const lastRow = sheet.getLastRow() + 1;
  const newCode = generateNextCode('BNK', sheet, 1);
  
  sheet.getRange(lastRow, 1).setValue(newCode);
  sheet.getRange(lastRow, 2).setValue(accountName);
  sheet.getRange(lastRow, 3).setValue(bankName);
  sheet.getRange(lastRow, 4).setValue(currency);
  sheet.getRange(lastRow, 5).setValue(iban);
  sheet.getRange(lastRow, 9).setValue(0);
  sheet.getRange(lastRow, 10).setValue(new Date());
  sheet.getRange(lastRow, 11).setValue('Active');
  sheet.getRange(lastRow, 13).setValue('No');
  
  const createResponse = ui.alert(
    '✅ Bank Account Added!\n\n' +
    'Code: ' + newCode + '\n' +
    'Name: ' + accountName + '\n' +
    'Bank: ' + bankName + '\n' +
    'Currency: ' + currency + '\n\n' +
    'Create bank sheet now?',
    ui.ButtonSet.YES_NO
  );
  
  if (createResponse === ui.Button.YES) {
    createSingleBankSheet(ss, accountName, currency, 0);
    sheet.getRange(lastRow, 13).setValue('Yes');
    ui.alert('✅ Bank sheet "' + accountName + '" created!');
  }
}

function showBankAccounts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Bank Accounts');
  if (sheet) ss.setActiveSheet(sheet);
}

// ==================== 3. CREATE INDIVIDUAL SHEETS ====================
function createSingleCashSheet(ss, cashName, currency, openingBalance) {
  let sheet = ss.getSheetByName(cashName);
  if (sheet) return sheet;
  
  sheet = ss.insertSheet(cashName);
  sheet.setTabColor('#ff5722');
  
  // Title
  sheet.getRange('A1:H1').merge()
    .setValue('💰 ' + cashName + ' (' + currency + ')')
    .setBackground('#ff5722')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center');
  
  // Current Balance
  sheet.getRange('A2').setValue('Current Balance (الرصيد الحالي)').setFontWeight('bold');
  sheet.getRange('B2')
    .setFormula('=SUMIF(G4:G1000,"IN",F4:F1000)-SUMIF(G4:G1000,"OUT",F4:F1000)')
    .setNumberFormat('#,##0.00')
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#fff3e0');
  sheet.getRange('C2').setValue(currency).setFontWeight('bold');
  
  // Headers
  const headers = ['Date', 'Description', 'Reference', 'Party', 'Trans. Code', 'Amount', 'Direction', 'Balance'];
  sheet.getRange(3, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#e64a19')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  // Opening Balance row
  sheet.getRange('A4').setValue(new Date()).setNumberFormat('yyyy-mm-dd');
  sheet.getRange('B4').setValue('Opening Balance (رصيد افتتاحي)');
  sheet.getRange('F4').setValue(openingBalance).setNumberFormat('#,##0.00');
  sheet.getRange('G4').setValue('IN');
  sheet.getRange('H4').setFormula('=F4').setNumberFormat('#,##0.00');
  
  // Column widths
  const widths = [100, 200, 120, 150, 120, 120, 80, 120];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  // Direction validation
  const dirRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['IN', 'OUT'], true).build();
  sheet.getRange('G4:G1000').setDataValidation(dirRule);
  
  // Conditional formatting
  const dirRange = sheet.getRange('G4:G1000');
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('IN').setBackground(COLORS.success).setRanges([dirRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('OUT').setBackground(COLORS.danger).setRanges([dirRange]).build()
  ]);
  
  sheet.getRange('A4:A1000').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('F4:F1000').setNumberFormat('#,##0.00');
  sheet.getRange('H4:H1000').setNumberFormat('#,##0.00');
  sheet.setFrozenRows(3);
  
  return sheet;
}

function createSingleBankSheet(ss, accountName, currency, openingBalance) {
  let sheet = ss.getSheetByName(accountName);
  if (sheet) return sheet;
  
  sheet = ss.insertSheet(accountName);
  sheet.setTabColor('#1565c0');
  
  // Title
  sheet.getRange('A1:H1').merge()
    .setValue('🏦 ' + accountName + ' (' + currency + ')')
    .setBackground('#1565c0')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center');
  
  // Current Balance
  sheet.getRange('A2').setValue('Current Balance (الرصيد الحالي)').setFontWeight('bold');
  sheet.getRange('B2')
    .setFormula('=SUMIF(G4:G1000,"IN",F4:F1000)-SUMIF(G4:G1000,"OUT",F4:F1000)')
    .setNumberFormat('#,##0.00')
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#e3f2fd');
  sheet.getRange('C2').setValue(currency).setFontWeight('bold');
  
  // Headers
  const headers = ['Date', 'Description', 'Reference', 'Party', 'Trans. Code', 'Amount', 'Direction', 'Balance'];
  sheet.getRange(3, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#0d47a1')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  // Opening Balance row
  sheet.getRange('A4').setValue(new Date()).setNumberFormat('yyyy-mm-dd');
  sheet.getRange('B4').setValue('Opening Balance (رصيد افتتاحي)');
  sheet.getRange('F4').setValue(openingBalance).setNumberFormat('#,##0.00');
  sheet.getRange('G4').setValue('IN');
  sheet.getRange('H4').setFormula('=F4').setNumberFormat('#,##0.00');
  
  const widths = [100, 200, 120, 150, 120, 120, 80, 120];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  const dirRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['IN', 'OUT'], true).build();
  sheet.getRange('G4:G1000').setDataValidation(dirRule);
  
  const dirRange = sheet.getRange('G4:G1000');
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('IN').setBackground(COLORS.success).setRanges([dirRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('OUT').setBackground(COLORS.danger).setRanges([dirRange]).build()
  ]);
  
  sheet.getRange('A4:A1000').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('F4:F1000').setNumberFormat('#,##0.00');
  sheet.getRange('H4:H1000').setNumberFormat('#,##0.00');
  sheet.setFrozenRows(3);
  
  return sheet;
}

// ==================== 4. CREATE ALL CASH/BANK SHEETS ====================
function createCashBankSheetsFromDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '🔄 Create Cash & Bank Sheets',
    'This will create individual sheets for all cash boxes and bank accounts.\n\n' +
    'Existing sheets will NOT be deleted.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  let cashCreated = 0, bankCreated = 0, skipped = 0;
  
  // Process Cash Boxes
  const cashSheet = ss.getSheetByName('Cash Boxes');
  if (cashSheet) {
    const cashData = cashSheet.getDataRange().getValues();
    for (let i = 1; i < cashData.length; i++) {
      const name = cashData[i][1];
      const currency = cashData[i][2];
      const opening = cashData[i][5] || 0;
      const status = cashData[i][7];
      const created = cashData[i][9];
      
      if (name && status === 'Active' && created !== 'Yes') {
        if (!ss.getSheetByName(name)) {
          createSingleCashSheet(ss, name, currency, opening);
          cashSheet.getRange(i + 1, 10).setValue('Yes');
          cashCreated++;
        } else {
          cashSheet.getRange(i + 1, 10).setValue('Yes');
          skipped++;
        }
      }
    }
  }
  
  // Process Bank Accounts
  const bankSheet = ss.getSheetByName('Bank Accounts');
  if (bankSheet) {
    const bankData = bankSheet.getDataRange().getValues();
    for (let i = 1; i < bankData.length; i++) {
      const name = bankData[i][1];
      const currency = bankData[i][3];
      const opening = bankData[i][8] || 0;
      const status = bankData[i][10];
      const created = bankData[i][12];
      
      if (name && status === 'Active' && created !== 'Yes') {
        if (!ss.getSheetByName(name)) {
          createSingleBankSheet(ss, name, currency, opening);
          bankSheet.getRange(i + 1, 13).setValue('Yes');
          bankCreated++;
        } else {
          bankSheet.getRange(i + 1, 13).setValue('Yes');
          skipped++;
        }
      }
    }
  }
  
  ui.alert(
    '✅ Cash & Bank Sheets Created!\n\n' +
    'Cash Boxes: ' + cashCreated + '\n' +
    'Bank Accounts: ' + bankCreated + '\n' +
    'Skipped (exists): ' + skipped
  );
}

// ==================== 5. GET LISTS ====================
function getCashBoxesList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Cash Boxes');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const cashBoxes = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][7] === 'Active' && data[i][1]) {
      cashBoxes.push({
        code: data[i][0],
        name: data[i][1],
        currency: data[i][2],
        sheetName: data[i][1]
      });
    }
  }
  return cashBoxes;
}

function getBankAccountsList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Bank Accounts');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const accounts = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][10] === 'Active' && data[i][1]) {
      accounts.push({
        code: data[i][0],
        name: data[i][1],
        bankName: data[i][2],
        currency: data[i][3],
        iban: data[i][4],
        sheetName: data[i][1]
      });
    }
  }
  return accounts;
}

function getCashBankBalance(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;
  return sheet.getRange('B2').getValue() || 0;
}

// ==================== 6. ADD ENTRY TO CASH/BANK ====================
function addCashBankEntry(sheetName, date, description, reference, party, transCode, amount, direction) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    console.error('Sheet not found: ' + sheetName);
    return false;
  }
  
  const lastRow = sheet.getLastRow() + 1;
  
  sheet.getRange(lastRow, 1).setValue(date).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(lastRow, 2).setValue(description);
  sheet.getRange(lastRow, 3).setValue(reference);
  sheet.getRange(lastRow, 4).setValue(party);
  sheet.getRange(lastRow, 5).setValue(transCode);
  sheet.getRange(lastRow, 6).setValue(amount).setNumberFormat('#,##0.00');
  sheet.getRange(lastRow, 7).setValue(direction);
  
  // Running balance formula
  if (lastRow > 4) {
    sheet.getRange(lastRow, 8).setFormula(
      '=H' + (lastRow - 1) + '+IF(G' + lastRow + '="IN",F' + lastRow + ',-F' + lastRow + ')'
    ).setNumberFormat('#,##0.00');
  } else {
    sheet.getRange(lastRow, 8).setFormula('=F' + lastRow).setNumberFormat('#,##0.00');
  }
  
  return true;
}

// ==================== 7. TRANSFER OPERATIONS ====================
function transferBetweenCashes() {
  const ui = SpreadsheetApp.getUi();
  const cashBoxes = getCashBoxesList();
  
  if (cashBoxes.length < 2) {
    ui.alert('⚠️ You need at least 2 cash boxes.');
    return;
  }
  
  // Select source
  const sourceList = cashBoxes.map((c, i) => (i + 1) + '. ' + c.name + ' (' + c.currency + ')').join('\n');
  const sourceResponse = ui.prompt(
    '🔄 Cash Transfer (1/3) - Select Source',
    'Available cash boxes:\n\n' + sourceList + '\n\nEnter number:',
    ui.ButtonSet.OK_CANCEL
  );
  if (sourceResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const sourceIndex = parseInt(sourceResponse.getResponseText()) - 1;
  if (isNaN(sourceIndex) || sourceIndex < 0 || sourceIndex >= cashBoxes.length) {
    ui.alert('⚠️ Invalid selection!'); return;
  }
  
  const sourceCash = cashBoxes[sourceIndex];
  const sourceBalance = getCashBankBalance(sourceCash.sheetName);
  
  // Select destination (same currency)
  const destCashBoxes = cashBoxes.filter((c, i) => i !== sourceIndex && c.currency === sourceCash.currency);
  if (destCashBoxes.length === 0) {
    ui.alert('⚠️ No other cash boxes with ' + sourceCash.currency);
    return;
  }
  
  const destList = destCashBoxes.map((c, i) => (i + 1) + '. ' + c.name).join('\n');
  const destResponse = ui.prompt(
    '🔄 Cash Transfer (2/3) - Select Destination\n\n' +
    'Source: ' + sourceCash.name + '\n' +
    'Balance: ' + formatCurrency(sourceBalance, sourceCash.currency),
    'Select destination:\n\n' + destList + '\n\nEnter number:',
    ui.ButtonSet.OK_CANCEL
  );
  if (destResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const destIndex = parseInt(destResponse.getResponseText()) - 1;
  if (isNaN(destIndex) || destIndex < 0 || destIndex >= destCashBoxes.length) {
    ui.alert('⚠️ Invalid selection!'); return;
  }
  
  const destCash = destCashBoxes[destIndex];
  
  // Amount
  const amountResponse = ui.prompt(
    '🔄 Cash Transfer (3/3) - Enter Amount\n\n' +
    'From: ' + sourceCash.name + '\n' +
    'To: ' + destCash.name + '\n' +
    'Available: ' + formatCurrency(sourceBalance, sourceCash.currency),
    'Enter amount:',
    ui.ButtonSet.OK_CANCEL
  );
  if (amountResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const amount = parseFloat(amountResponse.getResponseText());
  if (isNaN(amount) || amount <= 0) { ui.alert('⚠️ Invalid amount!'); return; }
  if (amount > sourceBalance) { ui.alert('⚠️ Insufficient balance!'); return; }
  
  // Execute
  const today = new Date();
  const transCode = 'TRF-' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  
  addCashBankEntry(sourceCash.sheetName, today, 'Transfer to ' + destCash.name, transCode, destCash.name, transCode, amount, 'OUT');
  addCashBankEntry(destCash.sheetName, today, 'Transfer from ' + sourceCash.name, transCode, sourceCash.name, transCode, amount, 'IN');
  
  ui.alert(
    '✅ Transfer Complete!\n\n' +
    'From: ' + sourceCash.name + '\n' +
    'To: ' + destCash.name + '\n' +
    'Amount: ' + formatCurrency(amount, sourceCash.currency) + '\n' +
    'Reference: ' + transCode
  );
}

function transferBetweenBanks() {
  const ui = SpreadsheetApp.getUi();
  const bankAccounts = getBankAccountsList();
  
  if (bankAccounts.length < 2) {
    ui.alert('⚠️ You need at least 2 bank accounts.');
    return;
  }
  
  // Similar logic to transferBetweenCashes
  const sourceList = bankAccounts.map((b, i) => (i + 1) + '. ' + b.name + ' (' + b.currency + ')').join('\n');
  const sourceResponse = ui.prompt(
    '🔄 Bank Transfer (1/3) - Select Source',
    'Available banks:\n\n' + sourceList + '\n\nEnter number:',
    ui.ButtonSet.OK_CANCEL
  );
  if (sourceResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const sourceIndex = parseInt(sourceResponse.getResponseText()) - 1;
  if (isNaN(sourceIndex) || sourceIndex < 0 || sourceIndex >= bankAccounts.length) {
    ui.alert('⚠️ Invalid selection!'); return;
  }
  
  const sourceBank = bankAccounts[sourceIndex];
  const sourceBalance = getCashBankBalance(sourceBank.sheetName);
  
  const destBanks = bankAccounts.filter((b, i) => i !== sourceIndex && b.currency === sourceBank.currency);
  if (destBanks.length === 0) {
    ui.alert('⚠️ No other banks with ' + sourceBank.currency);
    return;
  }
  
  const destList = destBanks.map((b, i) => (i + 1) + '. ' + b.name).join('\n');
  const destResponse = ui.prompt(
    '🔄 Bank Transfer (2/3)\n\nSource: ' + sourceBank.name + '\nBalance: ' + formatCurrency(sourceBalance, sourceBank.currency),
    'Select destination:\n\n' + destList,
    ui.ButtonSet.OK_CANCEL
  );
  if (destResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const destIndex = parseInt(destResponse.getResponseText()) - 1;
  const destBank = destBanks[destIndex];
  
  const amountResponse = ui.prompt('🔄 Bank Transfer (3/3)', 'Enter amount:', ui.ButtonSet.OK_CANCEL);
  if (amountResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const amount = parseFloat(amountResponse.getResponseText());
  if (isNaN(amount) || amount <= 0 || amount > sourceBalance) {
    ui.alert('⚠️ Invalid amount!'); return;
  }
  
  const today = new Date();
  const transCode = 'TRF-' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  
  addCashBankEntry(sourceBank.sheetName, today, 'Transfer to ' + destBank.name, transCode, destBank.name, transCode, amount, 'OUT');
  addCashBankEntry(destBank.sheetName, today, 'Transfer from ' + sourceBank.name, transCode, sourceBank.name, transCode, amount, 'IN');
  
  ui.alert('✅ Transfer Complete!\n\nAmount: ' + formatCurrency(amount, sourceBank.currency));
}

function cashToBankDeposit() {
  const ui = SpreadsheetApp.getUi();
  const cashBoxes = getCashBoxesList();
  const bankAccounts = getBankAccountsList();
  
  if (cashBoxes.length === 0 || bankAccounts.length === 0) {
    ui.alert('⚠️ Need at least 1 cash box and 1 bank account.');
    return;
  }
  
  // Select cash
  const cashList = cashBoxes.map((c, i) => (i + 1) + '. ' + c.name + ' (' + c.currency + ')').join('\n');
  const cashResponse = ui.prompt('🏦 Deposit (1/3) - Select Cash', cashList, ui.ButtonSet.OK_CANCEL);
  if (cashResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const cashIndex = parseInt(cashResponse.getResponseText()) - 1;
  const cash = cashBoxes[cashIndex];
  const cashBalance = getCashBankBalance(cash.sheetName);
  
  // Select bank (same currency)
  const banks = bankAccounts.filter(b => b.currency === cash.currency);
  if (banks.length === 0) { ui.alert('⚠️ No bank with ' + cash.currency); return; }
  
  const bankList = banks.map((b, i) => (i + 1) + '. ' + b.name).join('\n');
  const bankResponse = ui.prompt('🏦 Deposit (2/3) - Select Bank\n\nCash Balance: ' + formatCurrency(cashBalance, cash.currency), bankList, ui.ButtonSet.OK_CANCEL);
  if (bankResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const bankIndex = parseInt(bankResponse.getResponseText()) - 1;
  const bank = banks[bankIndex];
  
  const amountResponse = ui.prompt('🏦 Deposit (3/3) - Enter Amount', 'Available: ' + formatCurrency(cashBalance, cash.currency), ui.ButtonSet.OK_CANCEL);
  if (amountResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const amount = parseFloat(amountResponse.getResponseText());
  if (isNaN(amount) || amount <= 0 || amount > cashBalance) { ui.alert('⚠️ Invalid amount!'); return; }
  
  const today = new Date();
  const transCode = 'DEP-' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  
  addCashBankEntry(cash.sheetName, today, 'Deposit to ' + bank.name, transCode, bank.name, transCode, amount, 'OUT');
  addCashBankEntry(bank.sheetName, today, 'Deposit from ' + cash.name, transCode, cash.name, transCode, amount, 'IN');
  
  ui.alert('✅ Deposit Complete!\n\nAmount: ' + formatCurrency(amount, cash.currency));
}

function bankToCashWithdrawal() {
  const ui = SpreadsheetApp.getUi();
  const cashBoxes = getCashBoxesList();
  const bankAccounts = getBankAccountsList();
  
  if (cashBoxes.length === 0 || bankAccounts.length === 0) {
    ui.alert('⚠️ Need at least 1 cash box and 1 bank account.');
    return;
  }
  
  // Select bank
  const bankList = bankAccounts.map((b, i) => (i + 1) + '. ' + b.name + ' (' + b.currency + ')').join('\n');
  const bankResponse = ui.prompt('💵 Withdrawal (1/3) - Select Bank', bankList, ui.ButtonSet.OK_CANCEL);
  if (bankResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const bankIndex = parseInt(bankResponse.getResponseText()) - 1;
  const bank = bankAccounts[bankIndex];
  const bankBalance = getCashBankBalance(bank.sheetName);
  
  // Select cash (same currency)
  const cashes = cashBoxes.filter(c => c.currency === bank.currency);
  if (cashes.length === 0) { ui.alert('⚠️ No cash box with ' + bank.currency); return; }
  
  const cashList = cashes.map((c, i) => (i + 1) + '. ' + c.name).join('\n');
  const cashResponse = ui.prompt('💵 Withdrawal (2/3)\n\nBank Balance: ' + formatCurrency(bankBalance, bank.currency), cashList, ui.ButtonSet.OK_CANCEL);
  if (cashResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const cashIndex = parseInt(cashResponse.getResponseText()) - 1;
  const cash = cashes[cashIndex];
  
  const amountResponse = ui.prompt('💵 Withdrawal (3/3)', 'Available: ' + formatCurrency(bankBalance, bank.currency), ui.ButtonSet.OK_CANCEL);
  if (amountResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const amount = parseFloat(amountResponse.getResponseText());
  if (isNaN(amount) || amount <= 0 || amount > bankBalance) { ui.alert('⚠️ Invalid amount!'); return; }
  
  const today = new Date();
  const transCode = 'WDR-' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  
  addCashBankEntry(bank.sheetName, today, 'Withdrawal to ' + cash.name, transCode, cash.name, transCode, amount, 'OUT');
  addCashBankEntry(cash.sheetName, today, 'Withdrawal from ' + bank.name, transCode, bank.name, transCode, amount, 'IN');
  
  ui.alert('✅ Withdrawal Complete!\n\nAmount: ' + formatCurrency(amount, bank.currency));
}

function currencyExchange() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('💱 Currency Exchange\n\nThis feature will be available in a future update.\n\nUse manual entries for now.');
}
// ==================== 8. SYNC FROM TRANSACTIONS ====================

/**
 * مزامنة الحركات من شيت Transactions إلى شيتات الخزائن والبنوك
 * تُسجل فقط الحركات النقدية الفعلية (تحصيل/دفعة) وليس الاستحقاقات
 */
function syncAllCashAndBankSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '🔄 Sync Cash & Bank Sheets',
    'سيتم مزامنة الحركات من Transactions إلى شيتات الخزائن والبنوك.\n\n' +
    '✅ تحصيل إيراد (Revenue Collection) → IN\n' +
    '✅ دفع مصروف (Expense Payment) → OUT\n' +
    '❌ استحقاق → لا يُسجل\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const transSheet = ss.getSheetByName('Transactions');
  if (!transSheet || transSheet.getLastRow() < 2) {
    ui.alert('⚠️ No transactions found!');
    return;
  }
  
  const data = transSheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find column indices
  const colIndex = {
    date: 1,           // B
    movementType: 2,   // C
    description: 7,    // H
    partyName: 8,      // I
    amount: 10,        // K
    paymentMethod: 14, // O
    cashBank: 15,      // P
    reference: 16,     // Q
    transNum: 0        // A
  };
  
  let synced = 0, skipped = 0, errors = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    const movementType = row[colIndex.movementType] || '';
    const paymentMethod = row[colIndex.paymentMethod] || '';
    const cashBankName = row[colIndex.cashBank] || '';
    
    // Skip if no Cash/Bank selected
    if (!cashBankName) {
      skipped++;
      continue;
    }
    
    // Skip Accruals - only process actual cash movements
    if (paymentMethod.includes('Accrual') || paymentMethod.includes('استحقاق')) {
      skipped++;
      continue;
    }
    
    // Skip if not Cash or Bank payment
    if (!paymentMethod.includes('Cash') && !paymentMethod.includes('Bank') && 
        !paymentMethod.includes('نقدي') && !paymentMethod.includes('تحويل بنكي')) {
      skipped++;
      continue;
    }
    
    // Determine direction based on movement type
    let direction = '';
    if (movementType.includes('Collection') || movementType.includes('تحصيل')) {
      direction = 'IN';
    } else if (movementType.includes('Payment') || movementType.includes('دفع')) {
      direction = 'OUT';
    } else {
      skipped++;
      continue;
    }
    
    // Extract sheet name from dropdown value (remove emoji and currency)
    // Format: "💰 Cash TRY - Main (TRY)" → "Cash TRY - Main"
    let sheetName = cashBankName
      .replace(/^💰\s*/, '')
      .replace(/^🏦\s*/, '')
      .replace(/\s*\([A-Z]{3}\)$/, '')
      .trim();
    
    const targetSheet = ss.getSheetByName(sheetName);
    if (!targetSheet) {
      console.log('Sheet not found: ' + sheetName);
      errors++;
      continue;
    }
    
    // Check if already synced (by Trans Code)
    const transCode = 'TRX-' + (row[colIndex.transNum] || (i + 1));
    const existingData = targetSheet.getRange('E4:E' + targetSheet.getLastRow()).getValues();
    let alreadyExists = false;
    
    for (let j = 0; j < existingData.length; j++) {
      if (existingData[j][0] === transCode) {
        alreadyExists = true;
        break;
      }
    }
    
    if (alreadyExists) {
      skipped++;
      continue;
    }
    
    // Add entry
    const success = addCashBankEntry(
      sheetName,
      row[colIndex.date] || new Date(),
      row[colIndex.description] || movementType,
      row[colIndex.reference] || '',
      row[colIndex.partyName] || '',
      transCode,
      row[colIndex.amount] || 0,
      direction
    );
    
    if (success) {
      synced++;
    } else {
      errors++;
    }
  }
  
  ui.alert(
    '✅ Sync Complete!\n\n' +
    '📥 Synced: ' + synced + '\n' +
    '⏭️ Skipped: ' + skipped + '\n' +
    '❌ Errors: ' + errors
  );
}
// ==================== END OF PART 4 ====================
