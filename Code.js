// ╔════════════════════════════════════════════════════════════════════════════╗
// ║                    DC CONSULTING ACCOUNTING SYSTEM v3.0                     ║
// ║                              Part 1 of 9                                    ║
// ║                    Core + Menu + Config + Security                          ║
// ╚════════════════════════════════════════════════════════════════════════════╝

// ==================== GLOBAL CONSTANTS ====================
const SYSTEM_VERSION = '3.0';
const DEFAULT_PASSWORD = 'DC2025';

const CURRENCIES = ['TRY', 'USD', 'EUR', 'SAR', 'EGP', 'AED', 'GBP'];

const COLORS = {
  header: '#1565c0',
  headerText: '#ffffff',
  success: '#c8e6c9',
  warning: '#fff9c4',
  danger: '#ffcdd2',
  info: '#bbdefb',
  purple: '#e1bee7',
  rowEven: '#e3f2fd',
  rowOdd: '#ffffff'
};

// ==================== 1. MENU SYSTEM ====================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🏢 DC Consulting')
    // Setup
    .addItem('🔐 Setup System (إعداد النظام)', 'setupSystemSecure')
    .addItem('🔄 Refresh All Data (تحديث البيانات)', 'refreshAllData')
    .addSeparator()
    
    // Financial Operations
    .addSubMenu(ui.createMenu('💰 Financial Operations (العمليات المالية)')
      .addItem('➕ Add Transaction (إضافة معاملة)', 'addTransaction')
      .addSeparator()
      .addItem('🔄 Transfer: Cash ↔ Cash', 'transferBetweenCashes')
      .addItem('🔄 Transfer: Bank ↔ Bank', 'transferBetweenBanks')
      .addItem('🏦 Deposit: Cash → Bank', 'cashToBankDeposit')
      .addItem('💵 Withdraw: Bank → Cash', 'bankToCashWithdrawal')
      .addItem('💱 Currency Exchange (صرف عملات)', 'currencyExchange'))
    
    // Invoices
    .addSubMenu(ui.createMenu('🧾 Invoices (الفواتير)')
      .addItem('📄 Generate from Transaction', 'generateInvoiceFromTransaction')
      .addItem('📝 Generate Custom Invoice', 'generateCustomInvoice')
      .addItem('📋 Generate All Monthly', 'generateAllMonthlyInvoices')
      .addSeparator()
      .addItem('📧 Send Pending Invoices', 'sendPendingInvoices')
      .addItem('👁️ Preview Invoice', 'previewInvoice')
      .addSeparator()
      .addItem('📊 Invoice Log', 'showInvoiceLog'))
    
    // Clients & Parties
    .addSubMenu(ui.createMenu('👥 Clients & Parties (العملاء والأطراف)')
      .addItem('➕ Add Client (إضافة عميل)', 'addNewClient')
      .addItem('➕ Add Vendor (إضافة مورد)', 'addNewVendor')
      .addItem('➕ Add Employee (إضافة موظف)', 'addNewEmployee')
      .addSeparator()
      .addItem('📄 Client Statement (كشف حساب)', 'showClientStatement')
      .addItem('💹 Client Profitability (ربحية العميل)', 'showClientProfitability')
      .addSeparator()
      .addItem('🏢 Add Company Type Column', 'addCompanyTypeColumn'))
    
    // Cash & Bank
    .addSubMenu(ui.createMenu('🏦 Cash & Bank (الخزائن والبنوك)')
      .addItem('➕ Add Cash Box (إضافة خزينة)', 'addNewCashBox')
      .addItem('➕ Add Bank Account (إضافة حساب بنكي)', 'addNewBankAccount')
      .addItem('🔄 Create Cash/Bank Sheets', 'createCashBankSheetsFromDatabase')
      .addSeparator()
      .addItem('🔄 Sync to Cash/Bank (مزامنة)', 'syncAllCashAndBankSheets')
      .addSeparator()
      .addItem('📊 View Cash Boxes', 'showCashBoxes')
      .addItem('📊 View Bank Accounts', 'showBankAccounts'))

    // Advances (العهد)
    .addSubMenu(ui.createMenu('💼 Advances (العهد)')
      .addItem('💵 Issue Advance (صرف عهدة)', 'issueAdvance')
      .addItem('📝 Add Expense (إضافة مصروف)', 'addAdvanceExpense')
      .addItem('✅ Settle Advance (تسوية عهدة)', 'settleAdvance')
      .addSeparator()
      .addItem('📊 Advance Statement (كشف عهدة)', 'showAdvanceStatement')
      .addSeparator()
      .addItem('📋 View Advances', 'showAdvances')
      .addItem('📋 View Advance Expenses', 'showAdvanceExpenses'))

    
    // Reports
    .addSubMenu(ui.createMenu('📊 Reports (التقارير)')
      .addItem('📈 Dashboard (لوحة التحكم)', 'showDashboard')
      .addItem('📋 Clients Report', 'generateClientsReport')
      .addItem('⚠️ Overdue Report', 'generateOverdueReport')
      .addSeparator()
      .addItem('🔔 Check Overdue Now', 'checkOverdueNow')
      .addItem('📧 Send Overdue Reminders', 'sendOverdueReminders'))
    
    // Show/Hide Sheets
    .addSubMenu(ui.createMenu('👁️ Show/Hide Sheets (إظهار/إخفاء)')
      .addItem('📊 Show Reports', 'showReports')
      .addItem('🙈 Hide Reports', 'hideReports')
      .addSeparator()
      .addItem('🏦 Show Banks', 'showBanks')
      .addItem('🙈 Hide Banks', 'hideBanks')
      .addSeparator()
      .addItem('💰 Show Cash', 'showCash')
      .addItem('🙈 Hide Cash', 'hideCash')
      .addSeparator()
      .addItem('🗄️ Show Databases', 'showDatabases')
      .addItem('🙈 Hide Databases', 'hideDatabases')
      .addSeparator()
      .addItem('🔐 Hide All Sensitive', 'hideAllSensitive')
      .addItem('🔓 Show All Sheets', 'showAllSheets')
      .addSeparator()
      .addItem('👨‍💼 Accountant View', 'accountantView')
      .addItem('👔 Manager View', 'managerView')
      .addItem('📝 Data Entry View', 'dataEntryView'))
    
    // Settings
    .addSubMenu(ui.createMenu('⚙️ Settings (الإعدادات)')
      .addItem('📅 Manage Holidays', 'showHolidays')
      .addItem('🔧 System Settings', 'showSettingsSheet')
      .addItem('⏰ Setup Triggers', 'setupTriggers')
      .addItem('❌ Remove Triggers', 'removeAllTriggers')
      .addSeparator()
      .addItem('🔑 Change Password', 'changeAdminPassword')
      .addItem('🔄 Reset Password', 'resetPassword')
      .addSeparator()
      .addItem('🔄 Refresh Dropdowns', 'refreshAllDropdowns'))
    // في onOpen() أضف:
    .addSubMenu(ui.createMenu('📧 Email (البريد)')
      .addItem('📤 Send Pending Invoices', 'sendPendingInvoices')
      .addItem('📤 Send Selected Invoice', 'sendSelectedInvoice')
      .addSeparator()
      .addItem('📊 Email Statistics', 'showEmailStatistics')
      .addItem('📋 View Email Log', 'showEmailLog')
      .addSeparator()
      .addItem('⏰ Setup Auto Triggers', 'setupAutoTriggers')
      .addItem('🔍 Show Triggers Status', 'showTriggersStatus')
      .addItem('🗑️ Remove All Triggers', 'removeAllTriggers')
      .addSeparator()
      .addItem('📅 Test Invoice Schedule', 'testInvoiceSchedule'))
    // Help
    .addSeparator()
    .addItem('📖 User Guide (دليل المستخدم)', 'showUserGuide')
    .addItem('ℹ️ About (حول النظام)', 'showAbout')
    
    .addToUi();
  
  // ══════════════════════════════════════════════════════════════════
  // تحديث الـ Dropdowns تلقائياً عند فتح الشيت
  // ══════════════════════════════════════════════════════════════════
  try {
    refreshClientDropdowns();
    refreshItemsDropdown();
    refreshCashBankDropdown();
  } catch (e) {
    // تجاهل الأخطاء - قد تحدث إذا الشيتات غير موجودة بعد
    console.log('Dropdowns refresh skipped: ' + e.message);
  }
}

// ==================== 2. PASSWORD SYSTEM ====================
function verifyPassword(action) {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    '🔐 Security Verification (التحقق الأمني)',
    'Enter admin password to ' + action + ':\nأدخل كلمة السر للـ ' + action + ':',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return false;
  }
  
  const props = PropertiesService.getScriptProperties();
  const savedPassword = props.getProperty('ADMIN_PASSWORD') || DEFAULT_PASSWORD;
  
  if (response.getResponseText() !== savedPassword) {
    ui.alert('❌ Incorrect password! (كلمة سر خاطئة)');
    return false;
  }
  
  return true;
}

function changeAdminPassword() {
  const ui = SpreadsheetApp.getUi();
  
  if (!verifyPassword('change password')) return;
  
  const newPass = ui.prompt(
    '🔑 New Password (كلمة سر جديدة)',
    'Enter new admin password (minimum 4 characters):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (newPass.getSelectedButton() !== ui.Button.OK) return;
  
  if (!newPass.getResponseText() || newPass.getResponseText().length < 4) {
    ui.alert('❌ Password must be at least 4 characters!');
    return;
  }
  
  const confirmPass = ui.prompt(
    '🔑 Confirm Password (تأكيد كلمة السر)',
    'Confirm new password:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (confirmPass.getSelectedButton() !== ui.Button.OK) return;
  
  if (newPass.getResponseText() !== confirmPass.getResponseText()) {
    ui.alert('❌ Passwords do not match!');
    return;
  }
  
  const props = PropertiesService.getScriptProperties();
  props.setProperty('ADMIN_PASSWORD', newPass.getResponseText());
  
  ui.alert('✅ Password changed successfully!');
}

function resetPassword() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    '⚠️ Reset Password',
    'Type "RESET" to confirm resetting to default (DC2025):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  if (response.getResponseText() !== 'RESET') {
    ui.alert('❌ Reset cancelled.');
    return;
  }
  
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('ADMIN_PASSWORD');
  
  ui.alert('✅ Password reset to default: DC2025');
}

// ==================== 3. SHEET GROUPS ====================
function getSheetGroups() {
  return {
    'reports': {
      name: '📊 Reports',
      patterns: ['Dashboard', 'Client Statement', 'Profitability', 'Alerts', 'Email Log', 'Invoice Log']
    },
    'banks': {
      name: '🏦 Bank Accounts',
      patterns: ['Kuveyt', 'Ziraat', 'Garanti', 'İş Bank', 'Yapı Kredi', 'Akbank', 'Halk', 'Vakıf', 'QNB']
    },
    'cash': {
      name: '💰 Cash Boxes',
      patterns: ['Cash TRY', 'Cash USD', 'Cash EUR', 'Cash EGP', 'Cash SAR', 'Cash AED', 'Cash GBP']
    },
    'databases': {
      name: '🗄️ Databases',
      patterns: ['Clients', 'Vendors', 'Employees', 'Items Database', 'Movement Types', 'Categories', 'Holidays', 'Cash Boxes', 'Bank Accounts']
    },
    'settings': {
      name: '⚙️ Settings',
      patterns: ['Settings', 'Invoice Template']
    }
  };
}

// ==================== 4. SHEET VISIBILITY ====================
function getSheetsInGroup(groupKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const groups = getSheetGroups();
  const group = groups[groupKey];
  
  if (!group) return [];
  
  const matchedSheets = [];
  
  allSheets.forEach(sheet => {
    const name = sheet.getName();
    for (const pattern of group.patterns) {
      if (name.includes(pattern) || name.startsWith(pattern)) {
        matchedSheets.push(sheet);
        break;
      }
    }
  });
  
  return matchedSheets;
}

function hideSheetGroup(groupKey, silent = false) {
  const ui = SpreadsheetApp.getUi();
  const sheets = getSheetsInGroup(groupKey);
  const groups = getSheetGroups();
  
  if (sheets.length === 0) {
    if (!silent) ui.alert('⚠️ No sheets found in this group!');
    return 0;
  }
  
  let hiddenCount = 0;
  sheets.forEach(sheet => {
    try {
      sheet.hideSheet();
      hiddenCount++;
    } catch (e) { /* Cannot hide last visible sheet */ }
  });
  
  if (!silent) {
    ui.alert(`✅ Hidden ${hiddenCount} sheets in "${groups[groupKey].name}"`);
  }
  
  return hiddenCount;
}

function showSheetGroup(groupKey, silent = false) {
  const ui = SpreadsheetApp.getUi();
  const sheets = getSheetsInGroup(groupKey);
  const groups = getSheetGroups();
  
  if (sheets.length === 0) {
    if (!silent) ui.alert('⚠️ No sheets found in this group!');
    return 0;
  }
  
  let shownCount = 0;
  sheets.forEach(sheet => {
    sheet.showSheet();
    shownCount++;
  });
  
  if (!silent) {
    ui.alert(`✅ Shown ${shownCount} sheets in "${groups[groupKey].name}"`);
  }
  
  return shownCount;
}

// Menu functions for show/hide
function hideReports() { hideSheetGroup('reports'); }
function showReports() { showSheetGroup('reports'); }
function hideBanks() { hideSheetGroup('banks'); }
function showBanks() { showSheetGroup('banks'); }
function hideCash() { hideSheetGroup('cash'); }
function showCash() { showSheetGroup('cash'); }
function hideDatabases() { hideSheetGroup('databases'); }
function showDatabases() { showSheetGroup('databases'); }

function hideAllSensitive() {
  if (!verifyPassword('hide all sensitive sheets')) return;
  
  const ui = SpreadsheetApp.getUi();
  let total = 0;
  total += hideSheetGroup('banks', true);
  total += hideSheetGroup('cash', true);
  total += hideSheetGroup('databases', true);
  total += hideSheetGroup('settings', true);
  
  ui.alert(`✅ Hidden ${total} sensitive sheets!`);
}

function showAllSheets() {
  if (!verifyPassword('show all sheets')) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  sheets.forEach(sheet => sheet.showSheet());
  
  SpreadsheetApp.getUi().alert(`✅ All ${sheets.length} sheets are now visible!`);
}

// ==================== 5. VIEW MODES ====================
function accountantView() {
  if (!verifyPassword('switch to Accountant View')) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(sheet => sheet.showSheet());
  
  SpreadsheetApp.getUi().alert('👨‍💼 Accountant View Active\n\nAll sheets visible.');
}

function managerView() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  sheets.forEach(sheet => {
    try { sheet.hideSheet(); } catch(e) {}
  });
  
  const trans = ss.getSheetByName('Transactions');
  if (trans) trans.showSheet();
  
  const dash = ss.getSheetByName('Dashboard');
  if (dash) dash.showSheet();
  
  showSheetGroup('reports', true);
  
  SpreadsheetApp.getUi().alert('👔 Manager View Active\n\nTransactions + Dashboard + Reports visible.');
}

function dataEntryView() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  sheets.forEach(sheet => {
    try { sheet.hideSheet(); } catch(e) {}
  });
  
  const trans = ss.getSheetByName('Transactions');
  if (trans) {
    trans.showSheet();
    ss.setActiveSheet(trans);
  }
  
  SpreadsheetApp.getUi().alert('📝 Data Entry View Active\n\nOnly Transactions visible.');
}

// ==================== 6. HELPER FUNCTIONS ====================
function formatCurrency(amount, currency) {
  if (amount === null || amount === undefined || isNaN(amount)) {
    amount = 0;
  }
  
  const formatted = Math.abs(amount).toLocaleString('tr-TR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
  
  const symbols = {
    'TRY': '₺', 'USD': '$', 'EUR': '€', 
    'EGP': 'E£', 'SAR': 'ر.س', 'GBP': '£', 'AED': 'د.إ'
  };
  
  const symbol = symbols[currency] || currency;
  const sign = amount < 0 ? '-' : '';
  
  return sign + formatted + ' ' + symbol;
}

function formatDate(date, format) {
  if (!date) return '';
  format = format || 'yyyy-MM-dd';
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), format);
}

function generateNextCode(prefix, sheet, codeColumn) {
  const data = sheet.getDataRange().getValues();
  let maxNum = 0;
  
  for (let i = 1; i < data.length; i++) {
    const code = data[i][codeColumn - 1];
    if (code && typeof code === 'string' && code.startsWith(prefix + '-')) {
      const num = parseInt(code.replace(prefix + '-', ''));
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }
  
  return prefix + '-' + String(maxNum + 1).padStart(3, '0');
}

function showAbout() {
  SpreadsheetApp.getUi().alert(
    'ℹ️ DC Consulting Accounting System\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Version: ' + SYSTEM_VERSION + '\n' +
    'Developer: Dewan Group\n' +
    'Contact: sales@aldewan.net\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    '© 2025 Dewan Consulting'
  );
}

// ==================== END OF PART 1 ====================
// ╔════════════════════════════════════════════════════════════════════════════╗
// ║                    DC CONSULTING ACCOUNTING SYSTEM v3.0                     ║
// ║                              Part 2 of 9                                    ║
// ║           Database Sheets: Settings, Holidays, Categories,                  ║
// ║                    Movement Types, Items Database                           ║
// ╚════════════════════════════════════════════════════════════════════════════╝

// ==================== 1. SETTINGS SHEET ====================
function createSettingsSheet(ss) {
  let sheet = ss.getSheetByName('Settings');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Settings');
  sheet.setTabColor('#607d8b');
  
  const headers = [['Setting', 'Value']];
  sheet.getRange('A1:B1').setValues(headers)
    .setBackground(COLORS.header)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold');
  
  const settings = [
    ['Company Name (EN)', 'Dewan Consulting'],
    ['Company Name (AR)', 'ديوان للاستشارات'],
    ['Company Name (TR)', 'DİVAN DANIŞMANLIK'],
    ['Company Address', 'Beycenter, Cumhuriyet, 1991. Sk., 34515 Esenyurt/İstanbul'],
    ['Company Phone', '+90 (552) 740 60 13'],
    ['Company Email', 'sales@aldewan.net'],
    ['Tax Office', 'Gunesli'],
    ['Tax Number', '0471079224'],
    ['', ''],
    ['── Bank Details ──', ''],
    ['Bank Name', 'Kuveyt Türk'],
    ['IBAN TRY', 'TR250020500009448735700002'],
    ['IBAN USD', 'TR680020500009448735700101'],
    ['SWIFT Code', 'KTEFTRIS'],
    ['', ''],
    ['── Invoice Settings ──', ''],
    ['Invoice Prefix', 'INV-'],
    ['Next Invoice Number', '1'],
    ['Invoice Due Days', '30'],
    ['', ''],
    ['── Reminder Settings ──', ''],
    ['First Reminder (Days)', '7'],
    ['Recurring Reminder (Days)', '90'],
    ['Admin Email', 'sales@aldewan.net'],
    ['', ''],
    ['── Schedule Settings ──', ''],
    ['Invoice Generation Day', '25'],
    ['Invoice Generation Hour', '9'],
    ['Invoice Send Day Offset', '2'],
    ['Invoice Send Hour', '18'],
    ['', ''],
    ['── System ──', ''],
    ['System Version', SYSTEM_VERSION],
    ['Last Setup Date', new Date().toISOString().split('T')[0]]
  ];
  
  sheet.getRange(2, 1, settings.length, 2).setValues(settings);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 350);
  sheet.setFrozenRows(1);
  
  return sheet;
}

function getSettingValue(settingName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === settingName) {
      return data[i][1];
    }
  }
  return null;
}

function setSettingValue(settingName, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) return false;
  
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === settingName) {
      sheet.getRange(i + 1, 2).setValue(value);
      return true;
    }
  }
  return false;
}

function showSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Settings');
  if (sheet) ss.setActiveSheet(sheet);
  else SpreadsheetApp.getUi().alert('⚠️ Settings sheet not found!');
}

// ==================== 2. HOLIDAYS SHEET ====================
function createHolidaysSheet(ss) {
  let sheet = ss.getSheetByName('Holidays');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Holidays');
  sheet.setTabColor('#e91e63');
  
  const headers = ['Date', 'Holiday Name (EN)', 'Holiday Name (AR)', 'Holiday Name (TR)', 'Type', 'Year'];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(COLORS.header)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold');
  
  const holidays2025 = [
    ['2025-01-01', "New Year's Day", 'رأس السنة', 'Yılbaşı', 'National', 2025],
    ['2025-03-30', 'Eid al-Fitr Day 1', 'عيد الفطر 1', 'Ramazan Bayramı 1', 'Religious', 2025],
    ['2025-03-31', 'Eid al-Fitr Day 2', 'عيد الفطر 2', 'Ramazan Bayramı 2', 'Religious', 2025],
    ['2025-04-01', 'Eid al-Fitr Day 3', 'عيد الفطر 3', 'Ramazan Bayramı 3', 'Religious', 2025],
    ['2025-04-23', "Children's Day", 'يوم الطفل', 'Çocuk Bayramı', 'National', 2025],
    ['2025-05-01', 'Labour Day', 'عيد العمال', 'İşçi Bayramı', 'National', 2025],
    ['2025-05-19', 'Youth Day', 'يوم الشباب', 'Gençlik Bayramı', 'National', 2025],
    ['2025-06-06', 'Eid al-Adha Day 1', 'عيد الأضحى 1', 'Kurban Bayramı 1', 'Religious', 2025],
    ['2025-06-07', 'Eid al-Adha Day 2', 'عيد الأضحى 2', 'Kurban Bayramı 2', 'Religious', 2025],
    ['2025-06-08', 'Eid al-Adha Day 3', 'عيد الأضحى 3', 'Kurban Bayramı 3', 'Religious', 2025],
    ['2025-06-09', 'Eid al-Adha Day 4', 'عيد الأضحى 4', 'Kurban Bayramı 4', 'Religious', 2025],
    ['2025-07-15', 'Democracy Day', 'يوم الديمقراطية', 'Demokrasi Günü', 'National', 2025],
    ['2025-08-30', 'Victory Day', 'يوم النصر', 'Zafer Bayramı', 'National', 2025],
    ['2025-10-29', 'Republic Day', 'يوم الجمهورية', 'Cumhuriyet Bayramı', 'National', 2025]
  ];
  
  sheet.getRange(2, 1, holidays2025.length, 6).setValues(holidays2025);
  
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 60);
  
  sheet.getRange(2, 1, holidays2025.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.setFrozenRows(1);
  
  return sheet;
}

function showHolidays() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Holidays');
  if (sheet) ss.setActiveSheet(sheet);
  else SpreadsheetApp.getUi().alert('⚠️ Holidays sheet not found!');
}

function isHolidayOrWeekend(date) {
  const dayOfWeek = date.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) return true;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Holidays');
  if (!sheet) return false;
  
  const holidays = sheet.getDataRange().getValues();
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  for (let i = 1; i < holidays.length; i++) {
    if (holidays[i][0]) {
      const holidayDate = Utilities.formatDate(new Date(holidays[i][0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (holidayDate === dateStr) return true;
    }
  }
  return false;
}

// ==================== 3. CATEGORIES SHEET (3 Languages) ====================
function createCategoriesSheet(ss) {
  let sheet = ss.getSheetByName('Categories');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Categories');
  sheet.setTabColor('#009688');
  
  const headers = [
    'Category Code',
    'Category Name (EN)',
    'Category Name (AR)',
    'Category Name (TR)',
    'Type',
    'Status'
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(COLORS.header)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold');
  
  const data = [
    ['SRV-REV', 'Service Revenue', 'إيرادات خدمات', 'Hizmet Geliri', 'REVENUE', 'Active'],
    ['DIR-EXP', 'Direct Expenses', 'مصاريف مباشرة', 'Doğrudan Giderler', 'EXPENSE', 'Active'],
    ['ADM-EXP', 'Administrative Expenses', 'مصاريف إدارية', 'İdari Giderler', 'EXPENSE', 'Active'],
    ['SAL-EXP', 'Salaries & Wages', 'رواتب وأجور', 'Maaş ve Ücretler', 'EXPENSE', 'Active'],
    ['TRF', 'Transfers', 'تحويلات', 'Transferler', 'TRANSFER', 'Active'],
    ['FX', 'Currency Exchange', 'صرف عملات', 'Döviz Bozdurma', 'TRANSFER', 'Active'],
    ['ADJ', 'Adjustments', 'تسويات', 'Düzeltmeler', 'ADJUSTMENT', 'Active'],
    ['OPN', 'Opening Balance', 'رصيد افتتاحي', 'Açılış Bakiyesi', 'ADJUSTMENT', 'Active']
  ];
  
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 80);
  
  // Data Validations
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['REVENUE', 'EXPENSE', 'TRANSFER', 'ADJUSTMENT'], true)
    .build();
  sheet.getRange(2, 5, 50, 1).setDataValidation(typeRule);
  
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive'], true)
    .build();
  sheet.getRange(2, 6, 50, 1).setDataValidation(statusRule);
  
  sheet.setFrozenRows(1);
  applyAlternatingColors(sheet, 2, data.length, headers.length);
  
  return sheet;
}

// ==================== 4. MOVEMENT TYPES SHEET (3 Languages) ====================
function createMovementTypesSheet(ss) {
  let sheet = ss.getSheetByName('Movement Types');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Movement Types');
  sheet.setTabColor('#795548');
  
  const headers = [
    'Type Code',
    'Type Name (EN)',
    'Type Name (AR)',
    'Type Name (TR)',
    'Category Code',
    'Direction',
    'Affects Cash/Bank',
    'Icon',
    'Status'
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(COLORS.header)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold');
  
  // بدون "مصروف مباشر" - كما طلبت
  const data = [
    ['REV-DUE', 'Revenue Accrual', 'استحقاق إيراد', 'Gelir Tahakkuku', 'SRV-REV', 'IN', 'No', '📈', 'Active'],
    ['REV-COL', 'Revenue Collection', 'تحصيل إيراد', 'Gelir Tahsilatı', 'SRV-REV', 'IN', 'Yes', '✅', 'Active'],
    ['EXP-DUE', 'Expense Accrual', 'استحقاق مصروف', 'Gider Tahakkuku', '', 'OUT', 'No', '📉', 'Active'],
    ['EXP-PAY', 'Expense Payment', 'دفع مصروف', 'Gider Ödemesi', '', 'OUT', 'Yes', '💸', 'Active'],
    ['TRF-CC', 'Cash to Cash', 'تحويل خزينة ↔ خزينة', 'Kasa Transferi', 'TRF', 'INTERNAL', 'Yes', '🔄', 'Active'],
    ['TRF-BB', 'Bank to Bank', 'تحويل بنك ↔ بنك', 'Banka Transferi', 'TRF', 'INTERNAL', 'Yes', '🔄', 'Active'],
    ['TRF-CB', 'Cash to Bank', 'إيداع خزينة → بنك', 'Kasadan Bankaya', 'TRF', 'INTERNAL', 'Yes', '🏦', 'Active'],
    ['TRF-BC', 'Bank to Cash', 'سحب بنك → خزينة', 'Bankadan Kasaya', 'TRF', 'INTERNAL', 'Yes', '💵', 'Active'],
    ['FX-EXC', 'Currency Exchange', 'صرف عملات', 'Döviz Bozdurma', 'FX', 'INTERNAL', 'Yes', '💱', 'Active'],
    ['ADJ-IN', 'Adjustment (Add)', 'تسوية إضافة', 'Düzeltme (+)', 'ADJ', 'IN', 'Yes', '➕', 'Active'],
    ['ADJ-OUT', 'Adjustment (Deduct)', 'تسوية خصم', 'Düzeltme (-)', 'ADJ', 'OUT', 'Yes', '➖', 'Active'],
    ['OPN-BAL', 'Opening Balance', 'رصيد افتتاحي', 'Açılış Bakiyesi', 'OPN', 'IN', 'Yes', '🔰', 'Active']
  ];
  
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  
  // Column widths
  const widths = [90, 160, 160, 160, 100, 90, 110, 50, 80];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  // Data Validations
  const dirRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['IN', 'OUT', 'INTERNAL'], true)
    .build();
  sheet.getRange(2, 6, 50, 1).setDataValidation(dirRule);
  
  const affectsRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .build();
  sheet.getRange(2, 7, 50, 1).setDataValidation(affectsRule);
  
  sheet.setFrozenRows(1);
  applyAlternatingColors(sheet, 2, data.length, headers.length);
  
  return sheet;
}

// ==================== 5. ITEMS DATABASE (3 Languages) ====================
function createItemsDatabase(ss) {
  let sheet = ss.getSheetByName('Items Database');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Items Database');
  sheet.setTabColor('#00bcd4');
  
  const headers = [
    'Item Code',
    'Item Name (EN)',
    'Item Name (AR)',
    'Item Name (TR)',
    'Type',
    'Default Price',
    'Currency',
    'Status'
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(COLORS.header)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold');
  
  const data = [
    // Services
    ['SRV-CONS', 'Monthly Consulting', 'استشارات شهرية', 'Aylık Danışmanlık', 'SERVICE', '', 'TRY', 'Active'],
    ['SRV-COMP', 'Company Formation', 'تأسيس شركة', 'Şirket Kuruluşu', 'SERVICE', '', 'TRY', 'Active'],
    ['SRV-TRANS', 'Translation', 'ترجمة', 'Tercüme', 'SERVICE', '', 'TRY', 'Active'],
    ['SRV-RESID', 'Residence Permit', 'إقامة', 'İkamet İzni', 'SERVICE', '', 'TRY', 'Active'],
    ['SRV-ADDR', 'Address Change', 'تغيير عنوان', 'Adres Değişikliği', 'SERVICE', '', 'TRY', 'Active'],
    // Admin Expenses
    ['EXP-RENT', 'Office Rent', 'إيجار مكتب', 'Ofis Kirası', 'EXPENSE', '', 'TRY', 'Active'],
    ['EXP-ELEC', 'Electricity', 'كهرباء', 'Elektrik', 'EXPENSE', '', 'TRY', 'Active'],
    ['EXP-INET', 'Internet', 'إنترنت', 'İnternet', 'EXPENSE', '', 'TRY', 'Active'],
    ['EXP-TEL', 'Telephone', 'هاتف', 'Telefon', 'EXPENSE', '', 'TRY', 'Active'],
    ['EXP-WATER', 'Water', 'مياه', 'Su', 'EXPENSE', '', 'TRY', 'Active'],
    ['EXP-GAS', 'Natural Gas', 'غاز', 'Doğalgaz', 'EXPENSE', '', 'TRY', 'Active'],
    // Salaries
    ['EXP-SAL', 'Salary', 'راتب', 'Maaş', 'SALARY', '', 'TRY', 'Active'],
    ['EXP-BONUS', 'Bonus', 'مكافأة', 'Prim', 'SALARY', '', 'TRY', 'Active'],
    // Government
    ['EXP-TAX', 'Tax Office Fees', 'رسوم ضرائب', 'Vergi Harçları', 'EXPENSE', '', 'TRY', 'Active'],
    ['EXP-CHMBR', 'Chamber of Commerce', 'غرفة تجارة', 'Ticaret Odası', 'EXPENSE', '', 'TRY', 'Active'],
    ['EXP-NOTR', 'Notary Fees', 'رسوم نوتر', 'Noter Harçları', 'EXPENSE', '', 'TRY', 'Active'],
    // Other
    ['EXP-OFFC', 'Office Supplies', 'مستلزمات مكتب', 'Ofis Malzemeleri', 'EXPENSE', '', 'TRY', 'Active'],
    ['EXP-TRVL', 'Transportation', 'مواصلات', 'Ulaşım', 'EXPENSE', '', 'TRY', 'Active'],
    ['EXP-BANK', 'Bank Charges', 'مصاريف بنكية', 'Banka Masrafları', 'EXPENSE', '', 'TRY', 'Active'],
    ['EXP-MISC', 'Miscellaneous', 'متنوعات', 'Çeşitli', 'EXPENSE', '', 'TRY', 'Active']
  ];
  
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  
  // Column widths
  const widths = [100, 160, 140, 160, 90, 100, 80, 80];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  // Data Validations
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['SERVICE', 'EXPENSE', 'SALARY'], true)
    .build();
  sheet.getRange(2, 5, 100, 1).setDataValidation(typeRule);
  
  const currencyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CURRENCIES, true)
    .build();
  sheet.getRange(2, 7, 100, 1).setDataValidation(currencyRule);
  
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive'], true)
    .build();
  sheet.getRange(2, 8, 100, 1).setDataValidation(statusRule);
  
  sheet.getRange(2, 6, 100, 1).setNumberFormat('#,##0.00');
  sheet.setFrozenRows(1);
  applyAlternatingColors(sheet, 2, data.length, headers.length);
  
  return sheet;
}

// ==================== 6. HELPER: ALTERNATING COLORS ====================
function applyAlternatingColors(sheet, startRow, numRows, numCols) {
  for (let i = 0; i < numRows; i++) {
    const rowRange = sheet.getRange(startRow + i, 1, 1, numCols);
    if (i % 2 === 0) {
      rowRange.setBackground(COLORS.rowEven);
    } else {
      rowRange.setBackground(COLORS.rowOdd);
    }
  }
}

// ==================== 7. GET FUNCTIONS ====================
function getCategoriesList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Categories');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const categories = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === 'Active') {
      categories.push({
        code: data[i][0],
        nameEN: data[i][1],
        nameAR: data[i][2],
        nameTR: data[i][3],
        type: data[i][4],
        display: data[i][1] + ' (' + data[i][2] + ')'
      });
    }
  }
  return categories;
}

function getMovementTypesList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Movement Types');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const types = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === 'Active') {
      types.push({
        code: data[i][0],
        nameEN: data[i][1],
        nameAR: data[i][2],
        nameTR: data[i][3],
        categoryCode: data[i][4],
        direction: data[i][5],
        affectsCashBank: data[i][6] === 'Yes',
        icon: data[i][7],
        display: data[i][1] + ' (' + data[i][2] + ')'
      });
    }
  }
  return types;
}

function getItemsList(type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Items Database');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const items = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][7] === 'Active' && (!type || data[i][4] === type)) {
      items.push({
        code: data[i][0],
        nameEN: data[i][1],
        nameAR: data[i][2],
        nameTR: data[i][3],
        type: data[i][4],
        defaultPrice: data[i][5] || 0,
        currency: data[i][6] || 'TRY',
        display: data[i][1] + ' (' + data[i][2] + ')'
      });
    }
  }
  return items;
}

// ==================== END OF PART 2 ====================
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
    'Company Type',          // E - NEW
    'Tax Number',            // F
    'Tax Office',            // G
    'Address',               // H
    'Phone',                 // I
    'Email',                 // J
    'Contact Person',        // K
    'Monthly Fee',           // L
    'Fee Currency',          // M
    'Language',              // N
    'Folder ID',             // O
    'Contract Start',        // P
    'Status',                // Q
    'Notes',                 // R
    'Created Date'           // S
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(COLORS.header)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const widths = [100, 180, 150, 180, 120, 120, 120, 250, 120, 200, 150, 100, 80, 70, 280, 100, 80, 200, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  const lastRow = 500;

  // Data validations
  // Company Type validation (column E)
  const companyTypeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Limited', 'Şahıs', 'Anonim', 'Mükellef'], true)
    .build();
  sheet.getRange(2, 5, lastRow, 1).setDataValidation(companyTypeValidation);

  const currencyValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CURRENCIES, true)
    .build();
  sheet.getRange(2, 13, lastRow, 1).setDataValidation(currencyValidation);

  const languageValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['EN', 'AR', 'TR'], true)
    .build();
  sheet.getRange(2, 14, lastRow, 1).setDataValidation(languageValidation);

  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'Suspended'], true)
    .build();
  sheet.getRange(2, 17, lastRow, 1).setDataValidation(statusValidation);

  // Number formats
  sheet.getRange(2, 12, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 16, lastRow, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 19, lastRow, 1).setNumberFormat('yyyy-mm-dd');
  
  // Conditional formatting for Status (column Q = 17)
  const statusRange = sheet.getRange(2, 17, lastRow, 1);
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
  sheet.getRange(lastRow, 5).setValue('Limited'); // Company Type
  sheet.getRange(lastRow, 13).setValue('TRY'); // Fee Currency
  sheet.getRange(lastRow, 14).setValue('AR'); // Language
  sheet.getRange(lastRow, 17).setValue('Active'); // Status
  sheet.getRange(lastRow, 19).setValue(new Date()); // Created Date
  
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
        companyType: data[i][cols['Company Type']] || '',
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
// ╔════════════════════════════════════════════════════════════════════════════╗
// ║                    DC CONSULTING ACCOUNTING SYSTEM v3.0                     ║
// ║                              Part 5 of 9                                    ║
// ║            Transactions Sheet + Smart Dropdowns + onEdit Handler            ║
// ║                         *** UPDATED VERSION ***                             ║
// ╚════════════════════════════════════════════════════════════════════════════╝

// ==================== BILINGUAL DROPDOWN VALUES ====================
const DROPDOWN_VALUES = {
  movementTypes: [
    'Revenue Accrual (استحقاق إيراد)',
    'Revenue Collection (تحصيل إيراد)',
    'Expense Accrual (استحقاق مصروف)',
    'Expense Payment (دفع مصروف)',
    'Cash Transfer (تحويل خزينة)',
    'Bank Transfer (تحويل بنكي)',
    'Cash to Bank (إيداع)',
    'Bank to Cash (سحب)',
    'Currency Exchange (صرف عملات)',
    'Adjustment Add (تسوية +)',
    'Adjustment Deduct (تسوية -)',
    'Opening Balance (رصيد افتتاحي)',
    'Advance Issue (صرف عهدة)',
    'Advance Return (رد عهدة)',
    'Expense (مصروف)'
  ],
  categories: [
    'Service Revenue (إيرادات خدمات)',
    'Direct Expenses (مصاريف مباشرة)',
    'Administrative Expenses (مصاريف إدارية)',
    'Salaries & Wages (رواتب وأجور)',
    'Transfers (تحويلات)',
    'Currency Exchange (صرف عملات)',
    'Adjustments (تسويات)',
    'Opening Balance (رصيد افتتاحي)',
    'Petty Cash Advance (عهدة مؤقتة)',
    'Other Expense (مصروف آخر)'
  ],
  partyTypes: [
    'Client (عميل)',
    'Vendor (مورد)',
    'Employee (موظف)',
    'Internal (داخلي)'
  ],
  paymentMethods: [
    'Cash (نقدي)',
    'Bank Transfer (تحويل بنكي)',
    'Accrual (استحقاق)',
    'Credit Card (بطاقة ائتمان)',
    'Advance (عهدة)'
  ],
  paymentStatus: [
    'Pending (معلق)',
    'Partial (جزئي)',
    'Paid (مدفوع)',
    'Cancelled (ملغي)'
  ],
  showInStatement: [
    'Yes (نعم)',
    'No (لا)'
  ]
};

// ==================== 1. CREATE TRANSACTIONS SHEET ====================
function createTransactionsSheet(ss) {
  let sheet = ss.getSheetByName('Transactions');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Transactions');
  sheet.setTabColor('#3f51b5');
  
  // 25 columns (A-Y)
  const headers = [
    '#',                      // A (1)
    'Date (التاريخ)',         // B (2)
    'Movement Type (نوع الحركة)', // C (3)
    'Category (التصنيف)',     // D (4)
    'Client Code (كود العميل)', // E (5)
    'Client Name (اسم العميل)', // F (6)
    'Item (البند)',           // G (7)
    'Description (الوصف)',    // H (8)
    'Party Name (اسم الطرف)', // I (9)
    'Party Type (نوع الطرف)', // J (10)
    'Amount (المبلغ)',        // K (11)
    'Currency (العملة)',      // L (12)
    'Exchange Rate (سعر الصرف)', // M (13)
    'Amount TRY (بالليرة)',   // N (14)
    'Payment Method (طريقة الدفع)', // O (15)
    'Cash/Bank (الخزينة/البنك)', // P (16)
    'Reference (المرجع)',     // Q (17)
    'Invoice No (رقم الفاتورة)', // R (18)
    'Status (الحالة)',        // S (19)
    'Due Date (تاريخ الاستحقاق)', // T (20)
    'Paid Amount (المدفوع)',  // U (21)
    'Remaining (المتبقي)',    // V (22)
    'Notes (ملاحظات)',        // W (23)
    'Attachment (مرفق)',      // X (24)
    'Show in Statement (إظهار في الكشف)' // Y (25)
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(COLORS.header)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true);
  
  // Column widths
  const widths = [40, 90, 180, 170, 100, 180, 160, 200, 180, 130, 100, 70, 80, 100, 150, 160, 100, 100, 120, 100, 100, 100, 200, 150, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  sheet.setRowHeight(1, 45);
  
  const lastRow = 1000;
  
  // ===== Static Data Validations =====
  
  // Movement Type (C)
  const movementRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DROPDOWN_VALUES.movementTypes, true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 3, lastRow, 1).setDataValidation(movementRule);
  
  // Category (D)
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DROPDOWN_VALUES.categories, true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 4, lastRow, 1).setDataValidation(categoryRule);
  
  // Party Type (J)
  const partyTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DROPDOWN_VALUES.partyTypes, true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 10, lastRow, 1).setDataValidation(partyTypeRule);
  
  // Currency (L)
  const currencyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CURRENCIES, true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 12, lastRow, 1).setDataValidation(currencyRule);
  
  // Payment Method (O)
  const paymentMethodRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DROPDOWN_VALUES.paymentMethods, true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 15, lastRow, 1).setDataValidation(paymentMethodRule);
  
  // Payment Status (S)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DROPDOWN_VALUES.paymentStatus, true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 19, lastRow, 1).setDataValidation(statusRule);
  
  // Show in Statement (Y)
  const showRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DROPDOWN_VALUES.showInStatement, true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 25, lastRow, 1).setDataValidation(showRule);
  
  // ===== Number Formats =====
  sheet.getRange(2, 2, lastRow, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 11, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 13, lastRow, 1).setNumberFormat('#,##0.0000');
  sheet.getRange(2, 14, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 20, lastRow, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 21, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 22, lastRow, 1).setNumberFormat('#,##0.00');
  
  // ===== Conditional Formatting - تلوين الصف بالكامل حسب الحالة =====
  // نطبق التنسيق على كل الأعمدة (A-Y = 25 عمود) بناءً على قيمة Status (العمود S)
  const fullRowRange = sheet.getRange(2, 1, lastRow, 25);

  sheet.setConditionalFormatRules([
    // ✅ Paid (مدفوع) - أخضر
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($S2,"Paid")')
      .setBackground('#c8e6c9')
      .setRanges([fullRowRange])
      .build(),
    // ⏳ Pending (معلق) - أصفر
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($S2,"Pending")')
      .setBackground('#fff9c4')
      .setRanges([fullRowRange])
      .build(),
    // 🔶 Partial (جزئي) - برتقالي
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($S2,"Partial")')
      .setBackground('#ffe0b2')
      .setRanges([fullRowRange])
      .build(),
    // ❌ Cancelled (ملغي) - أحمر
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($S2,"Cancelled")')
      .setBackground('#ffcdd2')
      .setRanges([fullRowRange])
      .build()
  ]);
  
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  
  // Add notes
  sheet.getRange('E1').setNote('Client Code: اختر الكود → الاسم يُملأ تلقائياً');
  sheet.getRange('F1').setNote('Client Name: اختر الاسم → الكود يُملأ تلقائياً');
  sheet.getRange('J1').setNote('Party Type: اختر النوع → يتغير dropdown في Party Name');
  sheet.getRange('I1').setNote('Party Name: Dropdown ديناميكي حسب Party Type');
  sheet.getRange('Y1').setNote('Show in Statement:\nYes = يظهر في كشف الحساب\nNo = مخفي (تكلفة داخلية)');
  
  return sheet;
}

// ==================== 2. REFRESH CLIENT DROPDOWNS ====================
/**
 * تحديث dropdown العملاء (الكود والاسم) ديناميكياً من شيت Clients
 */
function refreshClientDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName('Transactions');
  const clientsSheet = ss.getSheetByName('Clients');
  
  if (!transSheet || !clientsSheet) return;
  
  const lastClientRow = clientsSheet.getLastRow();
  if (lastClientRow < 2) return;
  
  const lastRow = 1000;
  
  // جمع بيانات العملاء النشطين
  const clientData = clientsSheet.getRange(2, 1, lastClientRow - 1, 16).getValues();
  
  const clientCodes = [];
  const clientNamesEN = [];
  
  clientData.forEach(row => {
    const code = row[0];      // A = Code
    const nameEN = row[1];    // B = Name EN
    const status = row[15];   // P = Status
    
    if (code && nameEN && status === 'Active') {
      clientCodes.push(code);
      clientNamesEN.push(nameEN);
    }
  });
  
  if (clientCodes.length === 0) return;
  
  // Client Code Dropdown (Column E)
  const codeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(clientCodes, true)
    .setAllowInvalid(true)
    .build();
  transSheet.getRange(2, 5, lastRow, 1).setDataValidation(codeRule);
  
  // Client Name Dropdown (Column F)
  const nameRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(clientNamesEN, true)
    .setAllowInvalid(true)
    .build();
  transSheet.getRange(2, 6, lastRow, 1).setDataValidation(nameRule);
}

// ==================== 3. REFRESH ITEMS DROPDOWN ====================
/**
 * تحديث dropdown البنود من شيت Items Database
 */
function refreshItemsDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName('Transactions');
  const itemsSheet = ss.getSheetByName('Items Database');
  
  if (!transSheet || !itemsSheet) return;
  
  const lastItemRow = itemsSheet.getLastRow();
  if (lastItemRow < 2) return;
  
  const lastRow = 1000;
  
  // جمع البنود بصيغة EN (AR)
  const itemData = itemsSheet.getRange(2, 2, lastItemRow - 1, 3).getValues();
  const items = [];
  
  itemData.forEach(row => {
    const nameEN = row[0]; // B = Name EN
    const nameAR = row[1]; // C = Name AR
    const status = row[2]; // يمكن إضافة عمود Status لاحقاً
    
    if (nameEN) {
      items.push(nameEN + ' (' + (nameAR || nameEN) + ')');
    }
  });
  
  if (items.length === 0) return;
  
  // Item Dropdown (Column G)
  const itemRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(items, true)
    .setAllowInvalid(true)
    .build();
  transSheet.getRange(2, 7, lastRow, 1).setDataValidation(itemRule);
}

// ==================== 4. REFRESH CASH/BANK DROPDOWN ====================
/**
 * تحديث dropdown الخزائن والبنوك
 */
function refreshCashBankDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName('Transactions');
  if (!transSheet) return;
  
  const cashBankList = [];
  
  // Cash Boxes
  const cashSheet = ss.getSheetByName('Cash Boxes');
  if (cashSheet && cashSheet.getLastRow() > 1) {
    const cashData = cashSheet.getRange(2, 2, cashSheet.getLastRow() - 1, 7).getValues();
    cashData.forEach(row => {
      const name = row[0];     // B = Name
      const currency = row[1]; // C = Currency
      const status = row[6];   // H = Status
      
      if (name && status === 'Active') {
        cashBankList.push('💰 ' + name + ' (' + currency + ')');
      }
    });
  }
  
  // Bank Accounts
  const bankSheet = ss.getSheetByName('Bank Accounts');
  if (bankSheet && bankSheet.getLastRow() > 1) {
    const bankData = bankSheet.getRange(2, 2, bankSheet.getLastRow() - 1, 10).getValues();
    bankData.forEach(row => {
      const name = row[0];     // B = Name
      const currency = row[2]; // D = Currency
      const status = row[9];   // K = Status
      
      if (name && status === 'Active') {
        cashBankList.push('🏦 ' + name + ' (' + currency + ')');
      }
    });
  }
  
  if (cashBankList.length > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(cashBankList, true)
      .setAllowInvalid(true)
      .build();
    transSheet.getRange(2, 16, 1000, 1).setDataValidation(rule);
  }
}

// ==================== 5. UPDATE PARTY NAME DROPDOWN ====================
/**
 * تحديث dropdown اسم الطرف ديناميكياً حسب نوع الطرف
 * يُستدعى من onEdit عند تغيير Party Type
 *
 * ✅ محدّث: منع تكرار الاسم + إصلاح قراءة البيانات
 */
function updatePartyNameDropdown(ss, sheet, row, partyType) {
  let partyList = [];

  /**
   * دالة مساعدة لتنسيق الاسم بدون تكرار
   * إذا الاسم العربي فارغ أو = الاسم الإنجليزي → نعرض الإنجليزي فقط
   */
  function formatPartyName(nameEN, nameAR) {
    if (nameAR && nameAR.trim() !== '' && nameAR.trim() !== nameEN.trim()) {
      return nameEN + ' (' + nameAR + ')';
    }
    return nameEN;
  }

  // ===== Client =====
  if (partyType.includes('Client') || partyType.includes('عميل')) {
    const clientsSheet = ss.getSheetByName('Clients');
    if (clientsSheet) {
      const lastRow = clientsSheet.getLastRow();
      if (lastRow >= 2) {
        const data = clientsSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const nameEN = data[i][1];  // B - Company Name (EN)
          const nameAR = data[i][2];  // C - Company Name (AR)
          const status = data[i][15]; // P - Status

          if (nameEN && status === 'Active') {
            partyList.push(formatPartyName(nameEN, nameAR));
          }
        }
      }
    }
  }

  // ===== Vendor =====
  else if (partyType.includes('Vendor') || partyType.includes('مورد')) {
    const vendorsSheet = ss.getSheetByName('Vendors');
    if (vendorsSheet) {
      const lastRow = vendorsSheet.getLastRow();
      if (lastRow >= 2) {
        const data = vendorsSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const nameEN = data[i][1];  // B - Vendor Name (EN)
          const nameAR = data[i][2];  // C - Vendor Name (AR)
          const status = data[i][15]; // P - Status

          if (nameEN && status === 'Active') {
            partyList.push(formatPartyName(nameEN, nameAR));
          }
        }
      }
    }
  }

  // ===== Employee =====
  else if (partyType.includes('Employee') || partyType.includes('موظف')) {
    const employeesSheet = ss.getSheetByName('Employees');
    if (employeesSheet) {
      const lastRow = employeesSheet.getLastRow();
      if (lastRow >= 2) {
        const data = employeesSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const nameEN = data[i][1];  // B - Full Name (EN)
          const nameAR = data[i][2];  // C - Full Name (AR)
          const status = data[i][14]; // O - Status

          if (nameEN && status === 'Active') {
            partyList.push(formatPartyName(nameEN, nameAR));
          }
        }
      }
    }
  }

  // ===== Internal (Cash/Bank) =====
  else if (partyType.includes('Internal') || partyType.includes('داخلي')) {
    // Cash Boxes
    const cashSheet = ss.getSheetByName('Cash Boxes');
    if (cashSheet) {
      const lastRow = cashSheet.getLastRow();
      if (lastRow >= 2) {
        const data = cashSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const name = data[i][1];     // B - Cash Name
          const currency = data[i][2]; // C - Currency
          const status = data[i][7];   // H - Status

          if (name && status === 'Active') {
            partyList.push('💰 ' + name + ' (' + currency + ')');
          }
        }
      }
    }

    // Bank Accounts
    const bankSheet = ss.getSheetByName('Bank Accounts');
    if (bankSheet) {
      const lastRow = bankSheet.getLastRow();
      if (lastRow >= 2) {
        const data = bankSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const name = data[i][1];     // B - Account Name
          const currency = data[i][3]; // D - Currency
          const status = data[i][10];  // K - Status

          if (name && status === 'Active') {
            partyList.push('🏦 ' + name + ' (' + currency + ')');
          }
        }
      }
    }
  }

  // تطبيق الـ dropdown على الخلية المحددة
  if (partyList.length > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(partyList, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(row, 9).setDataValidation(rule);
  } else {
    // إذا لم توجد بيانات، نمسح الـ validation ونضع رسالة
    sheet.getRange(row, 9).clearDataValidations();
  }
}

// ==================== 6. SETUP ALL TRANSACTION DROPDOWNS ====================
/**
 * إعداد جميع الـ dropdowns في شيت Transactions
 */
function setupTransactionDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const transSheet = ss.getSheetByName('Transactions');
  if (!transSheet) {
    ui.alert('❌ Transactions sheet not found!');
    return;
  }
  
  // 1. Client Dropdowns (Code & Name)
  refreshClientDropdowns();
  
  // 2. Items Dropdown
  refreshItemsDropdown();
  
  // 3. Cash/Bank Dropdown
  refreshCashBankDropdown();
  
  // 4. الـ dropdowns الثابتة موجودة في createTransactionsSheet
  
  ui.alert(
    '✅ Dropdowns Setup Complete!\n\n' +
    '• Client Code ✓ (ديناميكي)\n' +
    '• Client Name ✓ (ديناميكي)\n' +
    '• Items ✓ (ديناميكي)\n' +
    '• Cash/Bank ✓ (ديناميكي)\n' +
    '• Party Type → Party Name ✓ (ديناميكي)\n' +
    '• Movement Type ✓\n' +
    '• Category ✓\n' +
    '• Payment Method ✓\n' +
    '• Status ✓\n\n' +
    '💡 الـ dropdowns تتحدث تلقائياً!'
  );
}

function refreshAllDropdowns() {
  refreshClientDropdowns();
  refreshItemsDropdown();
  refreshCashBankDropdown();
  refreshTransactionsValidation();
  SpreadsheetApp.getUi().alert('✅ All dropdowns refreshed!');
}

// ==================== REFRESH TRANSACTIONS VALIDATION ====================
/**
 * تحديث جميع قواعد التحقق في شيت Transactions
 */
function refreshTransactionsValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transactions');

  if (!sheet) return;

  const lastRow = Math.max(sheet.getLastRow(), 1000);

  // Movement Type (C - column 3)
  const movementRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DROPDOWN_VALUES.movementTypes, true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 3, lastRow, 1).setDataValidation(movementRule);

  // Category (D - column 4)
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DROPDOWN_VALUES.categories, true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 4, lastRow, 1).setDataValidation(categoryRule);

  // Party Type (J - column 10)
  const partyTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DROPDOWN_VALUES.partyTypes, true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 10, lastRow, 1).setDataValidation(partyTypeRule);

  // Payment Method (O - column 15)
  const paymentMethodRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DROPDOWN_VALUES.paymentMethods, true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 15, lastRow, 1).setDataValidation(paymentMethodRule);

  // Payment Status (S - column 19)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DROPDOWN_VALUES.paymentStatus, true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 19, lastRow, 1).setDataValidation(statusRule);

  // Show in Statement (Y - column 25)
  const showRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DROPDOWN_VALUES.showInStatement, true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 25, lastRow, 1).setDataValidation(showRule);
}

// ==================== 7. ONEDIT HANDLER ====================
/**
 * Main onEdit trigger - يعالج التغييرات التلقائية
 */
function onEdit(e) {
  if (!e) return;
  
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  const value = e.value;
  const ss = e.source;
  
  // ═══════════════════════════════════════════════════════════
  // معالجة شيت Transactions
  // ═══════════════════════════════════════════════════════════
  if (sheetName === 'Transactions' && row >= 2) {
    
    // ───── Client Code (E, col 5) → Fill Client Name (F) ─────
    if (col === 5 && value) {
      const clientsSheet = ss.getSheetByName('Clients');
      if (clientsSheet && clientsSheet.getLastRow() > 1) {
        const clientData = clientsSheet.getDataRange().getValues();

        for (let i = 1; i < clientData.length; i++) {
          if (clientData[i][0] === value) { // Code match (Column A)
            const nameEN = clientData[i][1]; // Column B
            const nameAR = clientData[i][2]; // Column C

            // Fill Client Name
            sheet.getRange(row, 6).setValue(nameEN);

            // Fill Party Type
            sheet.getRange(row, 10).setValue('Client (عميل)');

            // Fill Party Name - ✅ محدّث: بدون تكرار الاسم
            const partyName = (nameAR && nameAR.trim() !== '' && nameAR.trim() !== nameEN.trim())
              ? nameEN + ' (' + nameAR + ')'
              : nameEN;
            sheet.getRange(row, 9).setValue(partyName);

            break;
          }
        }
      }
    }

    // ───── Client Name (F, col 6) → Fill Client Code (E) ─────
    if (col === 6 && value) {
      const clientsSheet = ss.getSheetByName('Clients');
      if (clientsSheet && clientsSheet.getLastRow() > 1) {
        const clientData = clientsSheet.getDataRange().getValues();

        for (let i = 1; i < clientData.length; i++) {
          const code = clientData[i][0];   // A
          const nameEN = clientData[i][1]; // B
          const nameAR = clientData[i][2]; // C
          const nameTR = clientData[i][3]; // D

          // Check if name matches EN, AR, or TR
          if (nameEN === value || nameAR === value || nameTR === value) {
            // Fill Client Code
            sheet.getRange(row, 5).setValue(code);

            // Fill Party Type
            sheet.getRange(row, 10).setValue('Client (عميل)');

            // Fill Party Name - ✅ محدّث: بدون تكرار الاسم
            const partyName = (nameAR && nameAR.trim() !== '' && nameAR.trim() !== nameEN.trim())
              ? nameEN + ' (' + nameAR + ')'
              : nameEN;
            sheet.getRange(row, 9).setValue(partyName);
            
            break;
          }
        }
      }
    }
    
    // ───── Party Type (J, col 10) → Update Party Name Dropdown (I) ─────
    if (col === 10 && value) {
      updatePartyNameDropdown(ss, sheet, row, value);

      // مسح القيمة القديمة في Party Name
      sheet.getRange(row, 9).setValue('');
    }

    // ───── Payment Method (O, col 15) ─────
    // ملاحظة: تم إزالة التلوين اليدوي لأنه يتعارض مع التنسيق الشرطي لـ Status
    // التلوين الآن يعتمد فقط على Status (العمود S) عبر التنسيق الشرطي

    // ───── Amount (K) / Currency (L) / Rate (M) → Amount TRY (N) ─────
    if (col === 11 || col === 12 || col === 13) {
      const amount = sheet.getRange(row, 11).getValue() || 0;
      const currency = sheet.getRange(row, 12).getValue() || 'TRY';
      const rate = sheet.getRange(row, 13).getValue() || 1;
      
      if (currency === 'TRY') {
        sheet.getRange(row, 14).setValue(amount);
      } else {
        sheet.getRange(row, 14).setValue(amount * rate);
      }
    }
    
    // ───── Amount (K) / Paid (U) → Remaining (V) ─────
    if (col === 11 || col === 21) {
      const amount = sheet.getRange(row, 11).getValue() || 0;
      const paid = sheet.getRange(row, 21).getValue() || 0;
      sheet.getRange(row, 22).setValue(amount - paid);
    }
  }
  
  // ═══════════════════════════════════════════════════════════
  // تحديث Dropdowns عند تعديل شيت Clients
  // ═══════════════════════════════════════════════════════════
  if (sheetName === 'Clients' && row >= 2) {
    // تحديث بعد تأخير قصير
    Utilities.sleep(300);
    refreshClientDropdowns();
  }
  
  // ═══════════════════════════════════════════════════════════
  // تحديث Dropdowns عند تعديل شيتات أخرى
  // ═══════════════════════════════════════════════════════════
  if (sheetName === 'Vendors' && row >= 2) {
    // لا نحتاج تحديث - سيتم عند اختيار Party Type
  }
  
  if (sheetName === 'Employees' && row >= 2) {
    // لا نحتاج تحديث - سيتم عند اختيار Party Type
  }
  
  if ((sheetName === 'Cash Boxes' || sheetName === 'Bank Accounts') && row >= 2) {
    Utilities.sleep(300);
    refreshCashBankDropdown();
  }
  
  if (sheetName === 'Items Database' && row >= 2) {
    Utilities.sleep(300);
    refreshItemsDropdown();
  }
}

// ==================== 8. PAYMENT METHOD COLORS ====================

/**
 * تلوين الصف حسب طريقة الدفع
 * دالة داخلية - لا تُشغّل مباشرة
 */
function applyPaymentMethodColor(sheet, row, paymentMethod) {
  // التحقق من المعاملات
  if (!sheet || !row) {
    console.log('applyPaymentMethodColor: Missing sheet or row');
    return;
  }
  
  const lastCol = 25;
  
  try {
    const rowRange = sheet.getRange(row, 1, 1, lastCol);
    
    // مسح اللون السابق
    rowRange.setBackground(null);
    
    if (!paymentMethod) return;
    
    let bgColor = null;
    
    if (paymentMethod.includes('Accrual') || paymentMethod.includes('استحقاق')) {
      bgColor = '#fff9c4'; // 🟡 أصفر - استحقاق
    } else if (paymentMethod.includes('Cash') || paymentMethod.includes('نقدي')) {
      bgColor = '#c8e6c9'; // 🟢 أخضر - نقدي
    } else if (paymentMethod.includes('Bank') || paymentMethod.includes('تحويل بنكي')) {
      bgColor = '#bbdefb'; // 🔵 أزرق - تحويل بنكي
    } else if (paymentMethod.includes('Credit') || paymentMethod.includes('بطاقة')) {
      bgColor = '#e1bee7'; // 🟣 بنفسجي - بطاقة ائتمان
    }
    
    if (bgColor) {
      rowRange.setBackground(bgColor);
    }
  } catch (e) {
    console.log('Error in applyPaymentMethodColor: ' + e.message);
  }
}

/**
 * تطبيق الألوان على كل الصفوف الموجودة
 * ✅ شغّل هذه الدالة
 */
function applyAllPaymentColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transactions');
  
  if (!sheet) {
    try {
      SpreadsheetApp.getUi().alert('❌ Transactions sheet not found!');
    } catch (e) {
      console.log('Transactions sheet not found!');
    }
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    try {
      SpreadsheetApp.getUi().alert('⚠️ No data in Transactions.');
    } catch (e) {
      console.log('No data in Transactions.');
    }
    return;
  }
  
  // قراءة عمود Payment Method (العمود O = 15)
  const paymentData = sheet.getRange(2, 15, lastRow - 1, 1).getValues();
  let colored = 0;
  
  for (let i = 0; i < paymentData.length; i++) {
    const paymentMethod = paymentData[i][0];
    if (paymentMethod) {
      applyPaymentMethodColor(sheet, i + 2, paymentMethod);
      colored++;
    }
  }
  
  console.log('Colored ' + colored + ' rows');
  
  try {
    SpreadsheetApp.getUi().alert(
      '✅ Colors Applied!\n\n' +
      colored + ' rows colored.\n\n' +
      '🟡 Yellow = Accrual (استحقاق)\n' +
      '🟢 Green = Cash (نقدي)\n' +
      '🔵 Blue = Bank Transfer (تحويل بنكي)\n' +
      '🟣 Purple = Credit Card (بطاقة ائتمان)'
    );
  } catch (e) {
    // Running from script editor
  }
}

// ==================== 9. ADD TRANSACTION ====================
/**
 * إضافة معاملة جديدة
 */
function addTransaction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Transactions');
  
  if (!sheet) {
    ui.alert('❌ Transactions sheet not found!');
    return;
  }
  
  ss.setActiveSheet(sheet);
  const lastRow = sheet.getLastRow() + 1;
  
  // Set auto number
  sheet.getRange(lastRow, 1).setValue(lastRow - 1);
  
  // Set default date
  sheet.getRange(lastRow, 2).setValue(new Date());
  
  // Set defaults
  sheet.getRange(lastRow, 12).setValue('TRY');
  sheet.getRange(lastRow, 13).setValue(1);
  sheet.getRange(lastRow, 19).setValue('Pending (معلق)');
  sheet.getRange(lastRow, 25).setValue('Yes (نعم)');
  
  // Select first input cell
  sheet.setActiveRange(sheet.getRange(lastRow, 3));
  
  ui.alert(
    '➕ Add Transaction (إضافة معاملة)\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Row #' + (lastRow - 1) + ' is ready.\n\n' +
    'Defaults:\n' +
    '• Date: Today\n' +
    '• Currency: TRY\n' +
    '• Exchange Rate: 1\n' +
    '• Status: Pending\n' +
    '• Show in Statement: Yes\n\n' +
    '💡 Tips:\n' +
    '• اختر Client Code → الاسم يُملأ تلقائياً\n' +
    '• اختر Party Type → يتغير dropdown الأسماء'
  );
}
function generateMissingTransactionNumbers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transactions');
  const ui = SpreadsheetApp.getUi();
  
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('❌ No transactions found!');
    return;
  }
  
  const lastRow = sheet.getLastRow();
  let fixed = 0;
  
  for (let i = 2; i <= lastRow; i++) {
    const currentNum = sheet.getRange(i, 1).getValue();
    if (!currentNum) {
      sheet.getRange(i, 1).setValue(i - 1);
      fixed++;
    }
  }
  
  ui.alert('✅ Generated ' + fixed + ' transaction numbers!');
}

// ==================== 10. UPDATE STATUS CONDITIONAL FORMATTING ====================
/**
 * تحديث التنسيق الشرطي لـ Status على الشيت الموجود
 * ✅ يمسح الألوان اليدوية القديمة ويطبق التنسيق الشرطي الجديد
 */
function updateStatusConditionalFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transactions');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('❌ Transactions sheet not found!');
    return;
  }

  const lastRow = Math.max(sheet.getLastRow(), 100);

  // 1. مسح كل الألوان اليدوية من الصفوف (ما عدا الهيدر)
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 25);
  dataRange.setBackground(null);

  // 2. مسح قواعد التنسيق الشرطي القديمة
  sheet.clearConditionalFormatRules();

  // 3. تطبيق قواعد التنسيق الشرطي الجديدة
  const fullRowRange = sheet.getRange(2, 1, lastRow, 25);

  sheet.setConditionalFormatRules([
    // ✅ Paid (مدفوع) - أخضر
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($S2,"Paid")')
      .setBackground('#c8e6c9')
      .setRanges([fullRowRange])
      .build(),
    // ⏳ Pending (معلق) - أصفر
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($S2,"Pending")')
      .setBackground('#fff9c4')
      .setRanges([fullRowRange])
      .build(),
    // 🔶 Partial (جزئي) - برتقالي
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($S2,"Partial")')
      .setBackground('#ffe0b2')
      .setRanges([fullRowRange])
      .build(),
    // ❌ Cancelled (ملغي) - أحمر
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($S2,"Cancelled")')
      .setBackground('#ffcdd2')
      .setRanges([fullRowRange])
      .build()
  ]);

  ui.alert(
    '✅ تم تحديث التنسيق الشرطي!\n\n' +
    '🟢 Paid = أخضر\n' +
    '🟡 Pending = أصفر\n' +
    '🟠 Partial = برتقالي\n' +
    '🔴 Cancelled = أحمر\n' +
    '⚪ فارغ = بدون لون\n\n' +
    '💡 الآن عند تغيير Status، يتغير لون الصف تلقائياً!'
  );
}
// ==================== END OF PART 5 ====================
// ╔════════════════════════════════════════════════════════════════════════════╗
// ║                    DC CONSULTING ACCOUNTING SYSTEM v3.0                     ║
// ║                              Part 6 of 9                                    ║
// ║                           Invoice System                                    ║
// ╚════════════════════════════════════════════════════════════════════════════╝

// ==================== 1. CREATE INVOICE LOG SHEET ====================
function createInvoiceLogSheet(ss) {
  let sheet = ss.getSheetByName('Invoice Log');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Invoice Log');
  sheet.setTabColor('#9c27b0');
  
  const headers = [
    'Invoice No',        // A
    'Invoice Date',      // B
    'Client Code',       // C
    'Client Name',       // D
    'Service',           // E
    'Period',            // F
    'Amount',            // G
    'Currency',          // H
    'Status',            // I
    'PDF Link',          // J
    'Send Email',        // K - Yes/No
    'Email Status',      // L - Pending/Sent/Failed
    'Email Sent Date',   // M
    'Trans. Code',       // N
    'Notes',             // O
    'Created Date'       // P
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#6a1b9a')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  const widths = [100, 100, 90, 180, 150, 100, 100, 70, 90, 250, 80, 100, 100, 120, 200, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  const lastRow = 500;
  
  // Status validation
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Draft', 'Issued', 'Sent', 'Paid', 'Cancelled'], true).build();
  sheet.getRange(2, 9, lastRow, 1).setDataValidation(statusRule);
  
  // Send Email validation
  const sendRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true).build();
  sheet.getRange(2, 11, lastRow, 1).setDataValidation(sendRule);
  
  // Email Status validation
  const emailStatusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Sent', 'Failed'], true).build();
  sheet.getRange(2, 12, lastRow, 1).setDataValidation(emailStatusRule);
  
  // Number formats
  sheet.getRange(2, 2, lastRow, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 7, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 13, lastRow, 1).setNumberFormat('yyyy-mm-dd HH:mm');
  sheet.getRange(2, 16, lastRow, 1).setNumberFormat('yyyy-mm-dd HH:mm');
  
  // Conditional formatting
  const statusRange = sheet.getRange(2, 9, lastRow, 1);
  const emailRange = sheet.getRange(2, 12, lastRow, 1);
  
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Paid').setBackground('#c8e6c9').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Sent').setBackground('#bbdefb').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Issued').setBackground('#fff9c4').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Cancelled').setBackground('#ffcdd2').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Sent').setBackground('#c8e6c9').setRanges([emailRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Failed').setBackground('#ffcdd2').setRanges([emailRange]).build()
  ]);
  
  sheet.setFrozenRows(1);
  
  return sheet;
}

// ==================== 2. CREATE INVOICE TEMPLATE SHEET ====================
function createInvoiceTemplateSheet(ss) {
  let sheet = ss.getSheetByName('Invoice Template');
  if (sheet) ss.deleteSheet(sheet);

  sheet = ss.insertSheet('Invoice Template');
  sheet.setTabColor('#673ab7');

  // Set column widths (6 columns now)
  sheet.setColumnWidth(1, 30);   // #
  sheet.setColumnWidth(2, 140);  // Item
  sheet.setColumnWidth(3, 180);  // Description
  sheet.setColumnWidth(4, 40);   // Qty
  sheet.setColumnWidth(5, 90);   // Unit Price
  sheet.setColumnWidth(6, 90);   // Total

  // Get logo URL from settings
  const companyLogo = getSettingValue('Company Logo URL') || '';
  let logoUrl = '';
  if (companyLogo && companyLogo.trim() !== '') {
    logoUrl = companyLogo.trim();
    // Handle Google Drive sharing links
    if (logoUrl.includes('drive.google.com/file/d/')) {
      const fileId = logoUrl.match(/\/d\/([^\/]+)/);
      if (fileId && fileId[1]) {
        logoUrl = 'https://drive.google.com/uc?export=view&id=' + fileId[1];
      }
    } else if (logoUrl.includes('drive.google.com/open?id=')) {
      const fileId = logoUrl.match(/id=([^&]+)/);
      if (fileId && fileId[1]) {
        logoUrl = 'https://drive.google.com/uc?export=view&id=' + fileId[1];
      }
    }
  }

  let currentRow = 1;

  // Row 1: Logo (centered) - if provided
  if (logoUrl) {
    sheet.getRange('A1:F1').merge();
    sheet.getRange('A1').setFormula('=IMAGE("' + logoUrl + '", 1)');
    sheet.setRowHeight(1, 60);
    sheet.getRange('A1').setHorizontalAlignment('center').setVerticalAlignment('middle');
    currentRow = 2;
  }

  // Company Header
  sheet.getRange('A' + currentRow + ':F' + currentRow).merge()
    .setValue(getSettingValue('Company Name (EN)') || 'Dewan Consulting')
    .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#1565c0').setFontColor('#ffffff');
  currentRow++;

  sheet.getRange('A' + currentRow + ':F' + currentRow).merge()
    .setValue(getSettingValue('Company Name (AR)') || 'ديوان للاستشارات')
    .setFontSize(14).setHorizontalAlignment('center')
    .setBackground('#1976d2').setFontColor('#ffffff');
  currentRow++;

  sheet.getRange('A' + currentRow + ':F' + currentRow).merge()
    .setValue(getSettingValue('Company Address') || '')
    .setFontSize(10).setHorizontalAlignment('center');
  currentRow++;

  // Empty row
  currentRow++;

  // Invoice Title
  sheet.getRange('A' + currentRow + ':F' + currentRow).merge()
    .setValue('INVOICE / فاتورة')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#e3f2fd');
  currentRow++;

  // Empty row
  currentRow++;

  // Invoice Details Section - store the starting row
  const detailsStartRow = currentRow;

  // Row 1: Invoice No & Date
  sheet.getRange('A' + detailsStartRow + ':B' + detailsStartRow).merge().setValue('Invoice No:').setFontWeight('bold');
  sheet.getRange('C' + detailsStartRow + ':D' + detailsStartRow).merge(); // Value placeholder
  sheet.getRange('E' + detailsStartRow).setValue('Date:').setFontWeight('bold');
  // F is for date value

  // Row 2: Client
  sheet.getRange('A' + (detailsStartRow + 1) + ':B' + (detailsStartRow + 1)).merge().setValue('Client:').setFontWeight('bold');
  sheet.getRange('C' + (detailsStartRow + 1) + ':F' + (detailsStartRow + 1)).merge(); // Value placeholder

  // Row 3: Company Type
  sheet.getRange('A' + (detailsStartRow + 2) + ':B' + (detailsStartRow + 2)).merge().setValue('Company Type:').setFontWeight('bold');
  sheet.getRange('C' + (detailsStartRow + 2) + ':F' + (detailsStartRow + 2)).merge(); // Value placeholder

  // Row 4: Tax Number
  sheet.getRange('A' + (detailsStartRow + 3) + ':B' + (detailsStartRow + 3)).merge().setValue('Tax Number:').setFontWeight('bold');
  sheet.getRange('C' + (detailsStartRow + 3) + ':F' + (detailsStartRow + 3)).merge(); // Value placeholder

  // Row 5: Address
  sheet.getRange('A' + (detailsStartRow + 4) + ':B' + (detailsStartRow + 4)).merge().setValue('Address:').setFontWeight('bold');
  sheet.getRange('C' + (detailsStartRow + 4) + ':F' + (detailsStartRow + 4)).merge(); // Value placeholder

  currentRow = detailsStartRow + 5;

  // Empty row
  currentRow++;

  // Items Table Header (6 columns)
  const tableHeaderRow = currentRow;
  sheet.getRange('A' + tableHeaderRow + ':F' + tableHeaderRow)
    .setValues([['#', 'Item', 'Description', 'Qty', 'Unit Price', 'Total']])
    .setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Totals Section (after items space)
  const totalsRow = tableHeaderRow + 12; // Leave space for items
  sheet.getRange('E' + totalsRow).setValue('Subtotal:').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('E' + (totalsRow + 1)).setValue('VAT (0%):').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('E' + (totalsRow + 2)).setValue('TOTAL:').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('right');
  sheet.getRange('F' + totalsRow + ':F' + (totalsRow + 2)).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
  sheet.getRange('F' + (totalsRow + 2)).setFontWeight('bold').setFontSize(12).setBackground('#e3f2fd');

  // Bank Details
  const bankRow = totalsRow + 4;
  sheet.getRange('A' + bankRow + ':F' + bankRow).merge()
    .setValue('Bank Details / التفاصيل البنكية')
    .setFontWeight('bold').setBackground('#f5f5f5');

  // Bank row 1
  sheet.getRange('A' + (bankRow + 1) + ':B' + (bankRow + 1)).merge().setValue('Bank:').setFontWeight('bold');
  sheet.getRange('C' + (bankRow + 1) + ':F' + (bankRow + 1)).merge().setValue(getSettingValue('Bank Name') || 'Kuveyt Türk');

  // Bank row 2
  sheet.getRange('A' + (bankRow + 2) + ':B' + (bankRow + 2)).merge().setValue('IBAN (TRY):').setFontWeight('bold');
  sheet.getRange('C' + (bankRow + 2) + ':F' + (bankRow + 2)).merge().setValue(getSettingValue('IBAN TRY') || '');

  // Bank row 3
  sheet.getRange('A' + (bankRow + 3) + ':B' + (bankRow + 3)).merge().setValue('IBAN (USD):').setFontWeight('bold');
  sheet.getRange('C' + (bankRow + 3) + ':F' + (bankRow + 3)).merge().setValue(getSettingValue('IBAN USD') || '');

  // Bank row 4
  sheet.getRange('A' + (bankRow + 4) + ':B' + (bankRow + 4)).merge().setValue('SWIFT:').setFontWeight('bold');
  sheet.getRange('C' + (bankRow + 4) + ':F' + (bankRow + 4)).merge().setValue(getSettingValue('SWIFT Code') || 'KTEFTRIS');

  // Footer
  const footerRow = bankRow + 6;
  sheet.getRange('A' + footerRow + ':F' + footerRow).merge()
    .setValue('Thank you for your business! / شكراً لتعاملكم معنا')
    .setHorizontalAlignment('center').setFontStyle('italic');

  sheet.setHiddenGridlines(true);

  return sheet;
}

// ==================== 3. GET NEXT INVOICE NUMBER ====================
function getNextInvoiceNumber() {
  const prefix = getSettingValue('Invoice Prefix') || 'INV-';
  const nextNum = parseInt(getSettingValue('Next Invoice Number')) || 1;
  const year = new Date().getFullYear();
  return prefix + year + '-' + String(nextNum).padStart(4, '0');
}

function incrementInvoiceNumber() {
  const currentNum = parseInt(getSettingValue('Next Invoice Number')) || 1;
  setSettingValue('Next Invoice Number', currentNum + 1);
}

// ==================== 4. GENERATE INVOICE FROM TRANSACTION (MULTI-ROW) ====================
function generateInvoiceFromTransaction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const transSheet = ss.getSheetByName('Transactions');
  
  if (!transSheet) {
    ui.alert('❌ Transactions sheet not found!');
    return;
  }
  
  const selection = transSheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  
  if (startRow < 2) {
    ui.alert('⚠️ Please select transaction row(s) first!\n\nاختر صف أو أكثر من الحركات');
    return;
  }
  
  const selectedData = [];
  let firstClientCode = null;
  let firstClientName = null;
  let totalAmount = 0;
  let currency = 'TRY';
  
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const rowData = transSheet.getRange(row, 1, 1, 25).getValues()[0];
    
    const transCode = rowData[0];
    const transDate = rowData[1];
    const clientCode = rowData[4];
    const clientName = rowData[5];
    const item = rowData[6];
    const description = rowData[7];
    const amount = rowData[10] || 0;
    const rowCurrency = rowData[11] || 'TRY';
    
    if (!amount || amount === 0) continue;
    
    if (firstClientCode === null) {
      firstClientCode = clientCode;
      firstClientName = clientName;
      currency = rowCurrency;
    } else if (clientCode !== firstClientCode && clientName !== firstClientName) {
      ui.alert('⚠️ All selected rows must be for the SAME client!\n\nكل الصفوف المختارة يجب أن تكون لنفس العميل');
      return;
    }
    
    if (rowCurrency !== currency) {
      ui.alert('⚠️ All selected rows must have the SAME currency!\n\nكل الصفوف يجب أن تكون بنفس العملة');
      return;
    }
    
    selectedData.push({
      row: row,
      transCode: transCode,
      transDate: transDate,
      item: item,
      description: description,
      amount: amount
    });
    
    totalAmount += amount;
  }
  
  if (selectedData.length === 0) {
    ui.alert('⚠️ No valid transactions selected!');
    return;
  }
  
  const clientData = firstClientCode ? getClientData(firstClientCode) : null;
  
  const itemsList = selectedData.map((d, i) => 
    (i + 1) + '. ' + (d.item || d.description || 'Item') + ': ' + formatCurrency(d.amount, currency)
  ).join('\n');
  
  const confirm = ui.alert(
    '📄 Generate Invoice (إنشاء فاتورة)\n\n' +
    'Client: ' + (firstClientName || firstClientCode) + '\n' +
    'Items: ' + selectedData.length + '\n\n' +
    itemsList + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━\n' +
    'TOTAL: ' + formatCurrency(totalAmount, currency) + '\n\n' +
    'Generate invoice?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  const invoiceNo = getNextInvoiceNumber();
  const invoiceDate = new Date();
  const period = Utilities.formatDate(selectedData[0].transDate || invoiceDate, Session.getScriptTimeZone(), 'MMMM yyyy');
  
  const items = selectedData.map(d => ({
    item: d.item || '',
    description: d.description || '',
    qty: 1,
    unitPrice: d.amount,
    total: d.amount
  }));
  
  fillInvoiceTemplate(ss, {
    invoiceNo: invoiceNo,
    invoiceDate: invoiceDate,
    clientName: firstClientName || (clientData ? clientData.nameEN : ''),
    clientNameAR: clientData ? clientData.nameAR : '',
    companyType: clientData ? clientData.companyType : '',
    taxNumber: clientData ? clientData.taxNumber : '',
    address: clientData ? clientData.address : '',
    period: period,
    items: items,
    currency: currency,
    subtotal: totalAmount,
    vat: 0,
    vatRate: 0,
    total: totalAmount
  });
  
  logInvoice({
    invoiceNo: invoiceNo,
    invoiceDate: invoiceDate,
    clientCode: firstClientCode,
    clientName: firstClientName,
    service: selectedData.length > 1 ? 'Multiple Items (' + selectedData.length + ')' : (selectedData[0].item || ''),
    period: period,
    amount: totalAmount,
    currency: currency,
    status: 'Issued',
    sendEmail: 'Yes',
    emailStatus: 'Pending',
    transCode: selectedData.map(d => 'TRX-' + d.transCode).join(', ')
  });
  
  incrementInvoiceNumber();
  
  selectedData.forEach(d => {
    transSheet.getRange(d.row, 18).setValue(invoiceNo);
  });
  
  // ===== Save PDF to client folder =====
  let pdfResult = null;
  if (clientData && clientData.folderId) {
    try {
      pdfResult = createInvoicePDF(invoiceNo, clientData.folderId);
      updateInvoicePDFLink(invoiceNo, pdfResult.url);
    } catch (e) {
      console.log('PDF creation error: ' + e.message);
    }
  }
  
  const templateSheet = ss.getSheetByName('Invoice Template');
  if (templateSheet) ss.setActiveSheet(templateSheet);
  
  ui.alert(
    '✅ Invoice Generated!\n\n' +
    'Invoice No: ' + invoiceNo + '\n' +
    'Items: ' + selectedData.length + '\n' +
    'Total: ' + formatCurrency(totalAmount, currency) + '\n\n' +
    (pdfResult ? '✅ PDF saved to client folder' : '⚠️ PDF not saved (no folder ID)') + '\n\n' +
    'All ' + selectedData.length + ' transactions updated with invoice number.'
  );
}

// ==================== 5. GENERATE CUSTOM INVOICE ====================
function generateCustomInvoice() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // ===== Step 1: Enter Client Code =====
  const codeResponse = ui.prompt(
    '📄 Custom Invoice (1/5) - Client Code\n\nإنشاء فاتورة مخصصة',
    'Enter Client Code (أدخل كود العميل):\n\nExample: CLI-001',
    ui.ButtonSet.OK_CANCEL
  );
  if (codeResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const clientCode = codeResponse.getResponseText().trim().toUpperCase();
  if (!clientCode) {
    ui.alert('⚠️ Client code cannot be empty!');
    return;
  }
  
  const clientData = getClientData(clientCode);
  if (!clientData) {
    ui.alert('❌ Client not found!\n\nالعميل غير موجود: ' + clientCode);
    return;
  }
  
  const clientConfirm = ui.alert(
    '✅ Client Found (تم العثور على العميل)\n\n' +
    'Code: ' + clientCode + '\n' +
    'Name (EN): ' + clientData.nameEN + '\n' +
    'Name (AR): ' + (clientData.nameAR || '-') + '\n' +
    'Tax Number: ' + (clientData.taxNumber || '-') + '\n\n' +
    'Continue with this client?',
    ui.ButtonSet.YES_NO
  );
  if (clientConfirm !== ui.Button.YES) return;
  
  // ===== Step 2: Enter Service Description =====
  const descResponse = ui.prompt(
    '📄 Custom Invoice (2/5) - Service\n\nClient: ' + clientData.nameEN,
    'Enter service description (وصف الخدمة):\n\nExample: Monthly Consulting - December 2025',
    ui.ButtonSet.OK_CANCEL
  );
  if (descResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const description = descResponse.getResponseText().trim();
  if (!description) {
    ui.alert('⚠️ Description cannot be empty!');
    return;
  }
  
  // ===== Step 3: Enter Amount (before VAT) =====
  const amountResponse = ui.prompt(
    '📄 Custom Invoice (3/5) - Amount\n\nClient: ' + clientData.nameEN + '\nService: ' + description,
    'Enter amount BEFORE VAT (المبلغ قبل الضريبة):\n\nThis is the net amount',
    ui.ButtonSet.OK_CANCEL
  );
  if (amountResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const netAmount = parseFloat(amountResponse.getResponseText().replace(/,/g, '')) || 0;
  if (netAmount <= 0) {
    ui.alert('⚠️ Invalid amount!');
    return;
  }
  
  // ===== Step 4: Select Currency =====
  const currencyResponse = ui.prompt(
    '📄 Custom Invoice (4/5) - Currency\n\nAmount: ' + netAmount.toLocaleString(),
    'Enter currency (العملة):\n\nOptions: TRY, USD, EUR, SAR, EGP, AED, GBP\n\nDefault: TRY',
    ui.ButtonSet.OK_CANCEL
  );
  if (currencyResponse.getSelectedButton() !== ui.Button.OK) return;
  
  let currency = currencyResponse.getResponseText().trim().toUpperCase() || 'TRY';
  if (!CURRENCIES.includes(currency)) {
    currency = 'TRY';
  }
  
  // ===== Step 5: VAT Selection =====
  const vatResponse = ui.alert(
    '📄 Custom Invoice (5/5) - KDV/VAT\n\n' +
    'Client: ' + clientData.nameEN + '\n' +
    'Service: ' + description + '\n' +
    'Net Amount: ' + formatCurrency(netAmount, currency) + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'Does this invoice include KDV (VAT 20%)?\n' +
    'هل الفاتورة تشمل ضريبة القيمة المضافة (20%)؟\n\n' +
    'YES = With KDV (مع ضريبة)\n' +
    'NO = Without KDV (بدون ضريبة)',
    ui.ButtonSet.YES_NO
  );
  
  const withVAT = (vatResponse === ui.Button.YES);
  const vatRate = withVAT ? 0.20 : 0;
  const vatAmount = netAmount * vatRate;
  const totalAmount = netAmount + vatAmount;
  
  // ===== Final Confirmation =====
  const finalConfirm = ui.alert(
    '📄 Confirm Invoice (تأكيد الفاتورة)\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Client: ' + clientData.nameEN + '\n' +
    'Service: ' + description + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Net Amount: ' + formatCurrency(netAmount, currency) + '\n' +
    'KDV (' + (vatRate * 100) + '%): ' + formatCurrency(vatAmount, currency) + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'TOTAL: ' + formatCurrency(totalAmount, currency) + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'Generate invoice and record transaction?',
    ui.ButtonSet.YES_NO
  );
  
  if (finalConfirm !== ui.Button.YES) return;
  
  // ===== Generate Invoice =====
  const invoiceNo = getNextInvoiceNumber();
  const invoiceDate = new Date();
  const period = Utilities.formatDate(invoiceDate, Session.getScriptTimeZone(), 'MMMM yyyy');
  
  fillInvoiceTemplate(ss, {
    invoiceNo: invoiceNo,
    invoiceDate: invoiceDate,
    clientName: clientData.nameEN,
    clientNameAR: clientData.nameAR || '',
    companyType: clientData.companyType || '',
    taxNumber: clientData.taxNumber || '',
    address: clientData.address || '',
    period: period,
    items: [{
      item: '',
      description: description,
      qty: 1,
      unitPrice: netAmount,
      total: netAmount
    }],
    currency: currency,
    subtotal: netAmount,
    vat: vatAmount,
    vatRate: vatRate * 100,
    total: totalAmount
  });
  
  logInvoice({
    invoiceNo: invoiceNo,
    invoiceDate: invoiceDate,
    clientCode: clientCode,
    clientName: clientData.nameEN,
    service: description,
    period: period,
    amount: totalAmount,
    currency: currency,
    status: 'Issued',
    sendEmail: 'Yes',
    emailStatus: 'Pending',
    transCode: '',
    notes: withVAT ? 'KDV 20% included' : 'No KDV'
  });
  
  // ===== Record Transaction =====
  const transSheet = ss.getSheetByName('Transactions');
  let transRow = null;
  if (transSheet) {
    const lastRow = transSheet.getLastRow() + 1;
    transRow = lastRow;
    
    transSheet.getRange(lastRow, 1).setValue(lastRow - 1);
    transSheet.getRange(lastRow, 2).setValue(invoiceDate);
    transSheet.getRange(lastRow, 3).setValue('Revenue Accrual (استحقاق إيراد)');
    transSheet.getRange(lastRow, 4).setValue('Service Revenue (إيرادات خدمات)');
    transSheet.getRange(lastRow, 5).setValue(clientCode);
    transSheet.getRange(lastRow, 6).setValue(clientData.nameEN);
    transSheet.getRange(lastRow, 8).setValue(description);
    transSheet.getRange(lastRow, 9).setValue(clientData.nameEN + ' (' + (clientData.nameAR || clientData.nameEN) + ')');
    transSheet.getRange(lastRow, 10).setValue('Client (عميل)');
    transSheet.getRange(lastRow, 11).setValue(totalAmount);
    transSheet.getRange(lastRow, 12).setValue(currency);
    transSheet.getRange(lastRow, 13).setValue(1);
    transSheet.getRange(lastRow, 14).setValue(totalAmount);
    transSheet.getRange(lastRow, 15).setValue('Accrual (استحقاق)');
    transSheet.getRange(lastRow, 18).setValue(invoiceNo);
    transSheet.getRange(lastRow, 19).setValue('Pending (معلق)');
    transSheet.getRange(lastRow, 25).setValue('Yes (نعم)');
    
    applyPaymentMethodColor(transSheet, lastRow, 'Accrual (استحقاق)');
  }
  
  incrementInvoiceNumber();
  
  // ===== Save PDF to client folder =====
  let pdfResult = null;
  let pdfSaved = false;
  
  if (clientData.folderId) {
    try {
      pdfResult = createInvoicePDF(invoiceNo, clientData.folderId);
      pdfSaved = true;
      updateInvoicePDFLink(invoiceNo, pdfResult.url);
      
      if (transSheet && transRow) {
        transSheet.getRange(transRow, 23).setValue('PDF: ' + pdfResult.url);
      }
    } catch (e) {
      console.log('PDF creation error: ' + e.message);
    }
  }
  
  const templateSheet = ss.getSheetByName('Invoice Template');
  if (templateSheet) ss.setActiveSheet(templateSheet);
  
  ui.alert(
    '✅ Invoice Generated & Transaction Recorded!\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Invoice No: ' + invoiceNo + '\n' +
    'Client: ' + clientData.nameEN + '\n' +
    'Total: ' + formatCurrency(totalAmount, currency) + '\n' +
    (withVAT ? '(Including KDV 20%)' : '(No KDV)') + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    '✅ Transaction added to Transactions sheet\n' +
    (pdfSaved ? '✅ PDF saved to client folder' : '⚠️ PDF not saved (no folder ID)') + '\n\n' +
    'Next: Email will be sent after 3 working days'
  );
}

// ==================== 6. GENERATE ALL MONTHLY INVOICES ====================
function generateAllMonthlyInvoices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const clients = getActiveClients().filter(c => c.monthlyFee > 0);
  
  if (clients.length === 0) {
    ui.alert('⚠️ No clients with monthly fees found!');
    return;
  }
  
  const clientsList = clients.map(c => '• ' + c.nameEN + ': ' + formatCurrency(c.monthlyFee, c.feeCurrency)).join('\n');
  
  const confirm = ui.alert(
    '📋 Generate All Monthly Invoices\n\n' +
    'This will create invoices for ' + clients.length + ' clients:\n\n' +
    clientsList + '\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  const invoiceDate = new Date();
  const period = Utilities.formatDate(invoiceDate, Session.getScriptTimeZone(), 'MMMM yyyy');
  let generated = 0;
  let pdfSaved = 0;
  
  clients.forEach(client => {
    const invoiceNo = getNextInvoiceNumber();
    const clientData = getClientData(client.code);
    
    // Fill template for each (to create PDF)
    fillInvoiceTemplate(ss, {
      invoiceNo: invoiceNo,
      invoiceDate: invoiceDate,
      clientName: client.nameEN,
      clientNameAR: clientData ? clientData.nameAR : '',
      companyType: clientData ? clientData.companyType : '',
      taxNumber: clientData ? clientData.taxNumber : '',
      address: clientData ? clientData.address : '',
      period: period,
      items: [{
        item: 'Monthly Consulting',
        description: 'استشارات شهرية',
        qty: 1,
        unitPrice: client.monthlyFee,
        total: client.monthlyFee
      }],
      currency: client.feeCurrency,
      subtotal: client.monthlyFee,
      vat: 0,
      vatRate: 0,
      total: client.monthlyFee
    });
    
    // Save PDF
    let pdfUrl = '';
    if (clientData && clientData.folderId) {
      try {
        const pdfResult = createInvoicePDF(invoiceNo, clientData.folderId);
        pdfUrl = pdfResult.url;
        pdfSaved++;
      } catch (e) {
        console.log('PDF error for ' + client.code + ': ' + e.message);
      }
    }
    
    // Log invoice
    logInvoice({
      invoiceNo: invoiceNo,
      invoiceDate: invoiceDate,
      clientCode: client.code,
      clientName: client.nameEN,
      service: 'Monthly Consulting (استشارات شهرية)',
      period: period,
      amount: client.monthlyFee,
      currency: client.feeCurrency,
      status: 'Issued',
      pdfLink: pdfUrl,
      sendEmail: 'Yes',
      emailStatus: 'Pending',
      transCode: ''
    });
    
    // Record transaction
    recordInvoiceTransaction(invoiceNo, client.code, client.nameEN, client.monthlyFee, client.feeCurrency, 'Monthly Consulting (استشارات شهرية)');
    
    incrementInvoiceNumber();
    generated++;
  });
  
  ui.alert(
    '✅ Monthly Invoices Generated!\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Total Invoices: ' + generated + '\n' +
    'PDFs Saved: ' + pdfSaved + '\n' +
    'Period: ' + period + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    '📧 Emails will be sent after 3 working days.\n' +
    'Use "Send Pending Invoices" to send manually.'
  );
}

// ==================== 7. FILL INVOICE TEMPLATE ====================
function fillInvoiceTemplate(ss, data) {
  let sheet = ss.getSheetByName('Invoice Template');
  if (!sheet) {
    sheet = createInvoiceTemplateSheet(ss);
  }

  // Check if logo exists to determine row offset
  const companyLogo = getSettingValue('Company Logo URL') || '';
  const hasLogo = companyLogo && companyLogo.trim() !== '';
  const rowOffset = hasLogo ? 1 : 0;

  // Details start row (7 without logo, 8 with logo)
  const detailsStartRow = 7 + rowOffset;
  // Items start row (14 without logo, 15 with logo) - 5 detail rows now (no Period)
  const itemsStartRow = 14 + rowOffset;
  // Totals row (26 without logo, 27 with logo)
  const totalsStartRow = 26 + rowOffset;

  // Clear previous data (5 detail rows now - no Period)
  sheet.getRange('C' + detailsStartRow + ':F' + (detailsStartRow + 4)).clearContent();
  sheet.getRange('A' + itemsStartRow + ':F' + (itemsStartRow + 9)).clearContent();
  sheet.getRange('F' + totalsStartRow + ':F' + (totalsStartRow + 2)).clearContent();

  // Invoice details (values in column C for merged cells, F for date)
  sheet.getRange('C' + detailsStartRow).setValue(data.invoiceNo);  // Invoice No value in C-D merged
  sheet.getRange('F' + detailsStartRow).setValue(formatDate(data.invoiceDate, 'yyyy-MM-dd'));  // Date value
  sheet.getRange('C' + (detailsStartRow + 1)).setValue(data.clientName + (data.clientNameAR ? ' / ' + data.clientNameAR : ''));  // Client in C-F merged
  sheet.getRange('C' + (detailsStartRow + 2)).setValue(data.companyType || '');  // Company Type in C-F merged
  sheet.getRange('C' + (detailsStartRow + 3)).setValue(data.taxNumber || '');  // Tax Number in C-F merged
  sheet.getRange('C' + (detailsStartRow + 4)).setValue(data.address || '');  // Address in C-F merged

  // Items - 6 columns: #, Item, Description, Qty, Unit Price, Total
  if (data.items && data.items.length > 0) {
    data.items.forEach((item, i) => {
      const row = itemsStartRow + i;
      if (row < itemsStartRow + 10) {
        sheet.getRange(row, 1).setValue(i + 1).setHorizontalAlignment('center');
        sheet.getRange(row, 2).setValue(item.item || '');  // Item column
        sheet.getRange(row, 3).setValue(item.description || '');  // Description column
        sheet.getRange(row, 4).setValue(item.qty || 1).setHorizontalAlignment('center');
        sheet.getRange(row, 5).setValue(item.unitPrice).setNumberFormat('#,##0.00');
        sheet.getRange(row, 6).setValue(item.total).setNumberFormat('#,##0.00');
      }
    });
  }

  // Totals (column F now)
  const currencySymbol = data.currency === 'TRY' ? '₺' : (data.currency === 'USD' ? '$' : (data.currency === 'EUR' ? '€' : data.currency));
  sheet.getRange('F' + totalsStartRow).setValue(data.subtotal).setNumberFormat('#,##0.00');
  sheet.getRange('E' + (totalsStartRow + 1)).setValue('VAT (' + (data.vatRate || 0) + '%):').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('F' + (totalsStartRow + 1)).setValue(data.vat || 0).setNumberFormat('#,##0.00');
  sheet.getRange('F' + (totalsStartRow + 2)).setValue(data.total).setNumberFormat('#,##0.00 "' + currencySymbol + '"');

  return sheet;
}

// ==================== 8. LOG INVOICE ====================
function logInvoice(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Invoice Log');
  if (!sheet) {
    sheet = createInvoiceLogSheet(ss);
  }
  
  const lastRow = sheet.getLastRow() + 1;
  
  sheet.getRange(lastRow, 1).setValue(data.invoiceNo);
  sheet.getRange(lastRow, 2).setValue(data.invoiceDate);
  sheet.getRange(lastRow, 3).setValue(data.clientCode);
  sheet.getRange(lastRow, 4).setValue(data.clientName);
  sheet.getRange(lastRow, 5).setValue(data.service);
  sheet.getRange(lastRow, 6).setValue(data.period);
  sheet.getRange(lastRow, 7).setValue(data.amount);
  sheet.getRange(lastRow, 8).setValue(data.currency);
  sheet.getRange(lastRow, 9).setValue(data.status || 'Issued');
  sheet.getRange(lastRow, 10).setValue(data.pdfLink || '');
  sheet.getRange(lastRow, 11).setValue(data.sendEmail || 'Yes');
  sheet.getRange(lastRow, 12).setValue(data.emailStatus || 'Pending');
  sheet.getRange(lastRow, 14).setValue(data.transCode || '');
  sheet.getRange(lastRow, 15).setValue(data.notes || '');
  sheet.getRange(lastRow, 16).setValue(new Date());
  
  return lastRow;
}

// ==================== 9. UPDATE INVOICE PDF LINK ====================
function updateInvoicePDFLink(invoiceNo, pdfUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Invoice Log');
  if (!logSheet) return;
  
  const data = logSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === invoiceNo) {
      logSheet.getRange(i + 1, 10).setValue(pdfUrl);
      logSheet.getRange(i + 1, 15).setValue('PDF saved to client folder');
      break;
    }
  }
}

// ==================== 10. CREATE INVOICE PDF ====================
function createInvoicePDF(invoiceNo, clientFolderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName('Invoice Template');
  
  if (!templateSheet) {
    throw new Error('Invoice Template not found!');
  }
  
  const url = ss.getUrl().replace(/edit$/, '') +
    'export?format=pdf' +
    '&gid=' + templateSheet.getSheetId() +
    '&size=A4' +
    '&portrait=true' +
    '&fitw=true' +
    '&gridlines=false' +
    '&printtitle=false' +
    '&sheetnames=false' +
    '&pagenum=false' +
    '&fzr=false';
  
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });
  
  const pdfBlob = response.getBlob().setName(invoiceNo + '.pdf');
  
  let file;
  
  if (clientFolderId) {
    // البحث عن أو إنشاء مجلد Invoices داخل مجلد العميل
    const invoicesFolder = getOrCreateInvoicesFolder(clientFolderId);
    
    if (invoicesFolder) {
      file = invoicesFolder.createFile(pdfBlob);
    } else {
      // إذا فشل، احفظ في مجلد العميل مباشرة
      try {
        const clientFolder = DriveApp.getFolderById(clientFolderId);
        file = clientFolder.createFile(pdfBlob);
      } catch (e) {
        file = DriveApp.createFile(pdfBlob);
      }
    }
  } else {
    file = DriveApp.createFile(pdfBlob);
  }
  
  return {
    fileId: file.getId(),
    url: file.getUrl(),
    name: file.getName()
  };
}


// ==================== 11. PREVIEW INVOICE (CREATE PDF) ====================
function previewInvoice() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const templateSheet = ss.getSheetByName('Invoice Template');
  if (!templateSheet) {
    ui.alert('❌ Invoice Template not found!');
    return;
  }
  
  const invoiceNo = templateSheet.getRange('B7').getValue();
  if (!invoiceNo) {
    ui.alert('⚠️ No invoice loaded in template!\n\nGenerate an invoice first.');
    return;
  }
  
  const logSheet = ss.getSheetByName('Invoice Log');
  let folderId = '';
  let clientCode = '';
  
  if (logSheet) {
    const logData = logSheet.getDataRange().getValues();
    for (let i = 1; i < logData.length; i++) {
      if (logData[i][0] === invoiceNo) {
        clientCode = logData[i][2];
        const clientData = getClientData(clientCode);
        if (clientData && clientData.folderId) {
          folderId = clientData.folderId;
        }
        break;
      }
    }
  }
  
  try {
    const pdf = createInvoicePDF(invoiceNo, folderId);
    updateInvoicePDFLink(invoiceNo, pdf.url);
    
    ui.alert(
      '✅ Invoice PDF Created!\n\n' +
      'Invoice: ' + invoiceNo + '\n' +
      'File: ' + pdf.name + '\n' +
      (folderId ? '📁 Saved to client folder' : '📁 Saved to root folder') + '\n\n' +
      'PDF Link:\n' + pdf.url
    );
    
  } catch (error) {
    ui.alert('❌ Error creating PDF:\n\n' + error.message);
  }
}

// ==================== 12. SHOW INVOICE LOG ====================
function showInvoiceLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Invoice Log');
  if (!sheet) {
    sheet = createInvoiceLogSheet(ss);
  }
  ss.setActiveSheet(sheet);
}

// ==================== 13. CLEAR INVOICE TEMPLATE ====================
function clearInvoiceTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Invoice Template');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Invoice Template not found!');
    return;
  }

  // Check if logo exists to determine row offset
  const companyLogo = getSettingValue('Company Logo URL') || '';
  const hasLogo = companyLogo && companyLogo.trim() !== '';
  const rowOffset = hasLogo ? 1 : 0;

  const detailsStartRow = 7 + rowOffset;
  const itemsStartRow = 14 + rowOffset;
  const totalsStartRow = 26 + rowOffset;

  sheet.getRange('C' + detailsStartRow + ':F' + (detailsStartRow + 4)).clearContent();
  sheet.getRange('A' + itemsStartRow + ':F' + (itemsStartRow + 9)).clearContent();
  sheet.getRange('F' + totalsStartRow + ':F' + (totalsStartRow + 2)).clearContent();

  SpreadsheetApp.getUi().alert('✅ Invoice Template cleared!');
}

// ==================== 14. RECORD INVOICE AS TRANSACTION ====================
function recordInvoiceTransaction(invoiceNo, clientCode, clientName, amount, currency, item) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName('Transactions');
  
  if (!transSheet) return null;
  
  const lastRow = transSheet.getLastRow() + 1;
  
  transSheet.getRange(lastRow, 1).setValue(lastRow - 1);
  transSheet.getRange(lastRow, 2).setValue(new Date());
  transSheet.getRange(lastRow, 3).setValue('Revenue Accrual (استحقاق إيراد)');
  transSheet.getRange(lastRow, 4).setValue('Service Revenue (إيرادات خدمات)');
  transSheet.getRange(lastRow, 5).setValue(clientCode);
  transSheet.getRange(lastRow, 6).setValue(clientName);
  transSheet.getRange(lastRow, 8).setValue(item);
  transSheet.getRange(lastRow, 10).setValue('Client (عميل)');
  transSheet.getRange(lastRow, 11).setValue(amount);
  transSheet.getRange(lastRow, 12).setValue(currency);
  transSheet.getRange(lastRow, 13).setValue(1);
  transSheet.getRange(lastRow, 15).setValue('Accrual (استحقاق)');
  transSheet.getRange(lastRow, 18).setValue(invoiceNo);
  transSheet.getRange(lastRow, 19).setValue('Pending (معلق)');
  transSheet.getRange(lastRow, 25).setValue('Yes (نعم)');
  
  applyPaymentMethodColor(transSheet, lastRow, 'Accrual (استحقاق)');
  
  return lastRow;
}
// ==================== 15. GET OR CREATE INVOICES FOLDER ====================
/**
 * البحث عن مجلد Invoices داخل مجلد الشركة أو إنشاؤه
 */
function getOrCreateInvoicesFolder(parentFolderId) {
  if (!parentFolderId) return null;
  
  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const folderName = 'Invoices';
    
    // البحث عن مجلد Invoices
    const folders = parentFolder.getFoldersByName(folderName);
    
    if (folders.hasNext()) {
      // المجلد موجود - إرجاعه
      return folders.next();
    } else {
      // المجلد غير موجود - إنشاؤه
      const newFolder = parentFolder.createFolder(folderName);
      console.log('Created Invoices folder in: ' + parentFolder.getName());
      return newFolder;
    }
    
  } catch (e) {
    console.log('Error getting/creating Invoices folder: ' + e.message);
    return null;
  }
}
// ==================== END OF PART 6 ====================
// ╔════════════════════════════════════════════════════════════════════════════╗
// ║                    DC CONSULTING ACCOUNTING SYSTEM v3.0                     ║
// ║                              Part 7 of 9                                    ║
// ║                    Email System + Triggers + Scheduling                     ║
// ╚════════════════════════════════════════════════════════════════════════════╝

// ==================== 1. CREATE EMAIL LOG SHEET ====================
function createEmailLogSheet(ss) {
  let sheet = ss.getSheetByName('Email Log');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Email Log');
  sheet.setTabColor('#f44336');
  
  const headers = [
    'Date/Time',       // A
    'Invoice No',      // B
    'Client Code',     // C
    'Client Name',     // D
    'Email',           // E
    'Language',        // F
    'Status',          // G - Sent/Failed
    'Error Message',   // H
    'Sent By',         // I - Auto/Manual
    'PDF Link'         // J
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#c62828')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  const widths = [150, 120, 90, 180, 200, 80, 80, 300, 80, 250];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  sheet.getRange(2, 1, 500, 1).setNumberFormat('yyyy-mm-dd HH:mm:ss');
  
  // Conditional formatting
  const statusRange = sheet.getRange(2, 7, 500, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Sent').setBackground('#c8e6c9').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Failed').setBackground('#ffcdd2').setRanges([statusRange]).build()
  ]);
  
  sheet.setFrozenRows(1);
  
  return sheet;
}

// ==================== 2. LOG EMAIL ====================
function logEmail(invoiceNo, clientCode, clientName, email, language, status, errorMsg, sentBy, pdfLink) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Email Log');
  if (!sheet) {
    sheet = createEmailLogSheet(ss);
  }
  
  const lastRow = sheet.getLastRow() + 1;
  
  sheet.getRange(lastRow, 1).setValue(new Date());
  sheet.getRange(lastRow, 2).setValue(invoiceNo);
  sheet.getRange(lastRow, 3).setValue(clientCode);
  sheet.getRange(lastRow, 4).setValue(clientName);
  sheet.getRange(lastRow, 5).setValue(email);
  sheet.getRange(lastRow, 6).setValue(language);
  sheet.getRange(lastRow, 7).setValue(status);
  sheet.getRange(lastRow, 8).setValue(errorMsg || '');
  sheet.getRange(lastRow, 9).setValue(sentBy || 'Manual');
  sheet.getRange(lastRow, 10).setValue(pdfLink || '');
}

// ==================== 3. WORKING DAYS CALCULATOR ====================

/**
 * التحقق هل اليوم عطلة رسمية
 */
function isHoliday(date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Holidays');
  if (!sheet) return false;
  
  const data = sheet.getDataRange().getValues();
  const checkDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      const holidayDate = Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (holidayDate === checkDate) {
        return true;
      }
    }
  }
  return false;
}

/**
 * التحقق هل اليوم عطلة أو نهاية أسبوع
 */
function isHolidayOrWeekend(date) {
  const day = date.getDay();
  // السبت = 6، الأحد = 0
  if (day === 0 || day === 6) return true;
  if (isHoliday(date)) return true;
  return false;
}

/**
 * التحقق هل اليوم يوم عمل
 */
function isWorkingDay(date) {
  return !isHolidayOrWeekend(date);
}

/**
 * إضافة أيام عمل لتاريخ معين
 */
function addWorkingDays(startDate, workingDays) {
  let currentDate = new Date(startDate);
  let addedDays = 0;
  
  while (addedDays < workingDays) {
    currentDate.setDate(currentDate.getDate() + 1);
    if (isWorkingDay(currentDate)) {
      addedDays++;
    }
  }
  
  return currentDate;
}

/**
 * حساب تاريخ الإرسال (3 أيام عمل بعد الإصدار)
 */
function calculateSendDate(issueDate) {
  return addWorkingDays(issueDate, 3);
}

/**
 * الحصول على تاريخ إصدار الفواتير (يوم 25 أو أقرب يوم عمل)
 */
function getInvoiceGenerationDate(year, month) {
  let genDate = new Date(year, month, 25);
  
  // إذا كان يوم 25 عطلة، انتقل لأقرب يوم عمل
  while (isHolidayOrWeekend(genDate)) {
    genDate.setDate(genDate.getDate() + 1);
  }
  
  genDate.setHours(9, 0, 0, 0);
  return genDate;
}

/**
 * الحصول على تاريخ إرسال الفواتير (3 أيام عمل بعد الإصدار)
 */
function getInvoiceSendDate(generationDate) {
  return addWorkingDays(generationDate, 3);
}

/**
 * اختبار جدول الفواتير
 */
function testInvoiceSchedule() {
  const ui = SpreadsheetApp.getUi();
  const now = new Date();
  
  const thisMonthGen = getInvoiceGenerationDate(now.getFullYear(), now.getMonth());
  const thisMonthSend = getInvoiceSendDate(thisMonthGen);
  
  const nextMonth = now.getMonth() === 11 ? 0 : now.getMonth() + 1;
  const nextYear = now.getMonth() === 11 ? now.getFullYear() + 1 : now.getFullYear();
  const nextMonthGen = getInvoiceGenerationDate(nextYear, nextMonth);
  const nextMonthSend = getInvoiceSendDate(nextMonthGen);
  
  ui.alert(
    '📅 Invoice Schedule Test\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'This Month:\n' +
    '• Generation: ' + Utilities.formatDate(thisMonthGen, Session.getScriptTimeZone(), 'EEEE, dd MMM yyyy') + '\n' +
    '• Send: ' + Utilities.formatDate(thisMonthSend, Session.getScriptTimeZone(), 'EEEE, dd MMM yyyy') + '\n\n' +
    'Next Month:\n' +
    '• Generation: ' + Utilities.formatDate(nextMonthGen, Session.getScriptTimeZone(), 'EEEE, dd MMM yyyy') + '\n' +
    '• Send: ' + Utilities.formatDate(nextMonthSend, Session.getScriptTimeZone(), 'EEEE, dd MMM yyyy')
  );
}

// ==================== 4. EMAIL TEMPLATES (3 LANGUAGES) ====================

/**
 * الحصول على قالب الإيميل حسب اللغة
 */
function getEmailTemplate(language, data) {
  const templates = {
    'EN': {
      subject: 'Invoice ' + data.invoiceNo + ' - ' + (getSettingValue('Company Name (EN)') || 'Dewan Consulting'),
      body: `
        <div style="font-family: Arial, sans-serif; padding: 20px; max-width: 600px;">
          <div style="background: #1565c0; color: white; padding: 20px; text-align: center;">
            <h1 style="margin: 0;">${getSettingValue('Company Name (EN)') || 'Dewan Consulting'}</h1>
            <p style="margin: 5px 0 0 0;">${getSettingValue('Company Name (AR)') || 'ديوان للاستشارات'}</p>
          </div>
          
          <div style="padding: 20px; background: #f5f5f5;">
            <p>Dear <strong>${data.contactPerson || data.clientName}</strong>,</p>
            
            <p>Please find attached your invoice for <strong>${data.period}</strong>.</p>
            
            <div style="background: white; padding: 15px; border-radius: 5px; margin: 20px 0;">
              <h3 style="margin-top: 0; color: #1565c0;">Invoice Details:</h3>
              <table style="width: 100%;">
                <tr><td><strong>Invoice No:</strong></td><td>${data.invoiceNo}</td></tr>
                <tr><td><strong>Date:</strong></td><td>${data.invoiceDate}</td></tr>
                <tr><td><strong>Amount:</strong></td><td style="font-size: 18px; color: #1565c0;"><strong>${data.amount}</strong></td></tr>
              </table>
            </div>
            
            <div style="background: white; padding: 15px; border-radius: 5px;">
              <h3 style="margin-top: 0; color: #1565c0;">Payment Details:</h3>
              <p><strong>Bank:</strong> ${getSettingValue('Bank Name') || 'Kuveyt Türk'}</p>
              <p><strong>IBAN (TRY):</strong> ${getSettingValue('IBAN TRY') || ''}</p>
              <p><strong>IBAN (USD):</strong> ${getSettingValue('IBAN USD') || ''}</p>
              <p><strong>SWIFT:</strong> ${getSettingValue('SWIFT Code') || 'KTEFTRIS'}</p>
            </div>
            
            <p style="margin-top: 20px;">If you have any questions, please don't hesitate to contact us.</p>
            
            <p>Best regards,<br>
            <strong>${getSettingValue('Company Name (EN)') || 'Dewan Consulting'}</strong><br>
            <a href="mailto:sales@aldewan.net">sales@aldewan.net</a></p>
          </div>
          
          <div style="background: #333; color: white; padding: 10px; text-align: center; font-size: 12px;">
            Thank you for your business!
          </div>
        </div>
      `
    },
    
    'TR': {
      subject: 'Fatura ' + data.invoiceNo + ' - ' + (getSettingValue('Company Name (EN)') || 'Dewan Consulting'),
      body: `
        <div style="font-family: Arial, sans-serif; padding: 20px; max-width: 600px;">
          <div style="background: #1565c0; color: white; padding: 20px; text-align: center;">
            <h1 style="margin: 0;">${getSettingValue('Company Name (EN)') || 'Dewan Consulting'}</h1>
            <p style="margin: 5px 0 0 0;">${getSettingValue('Company Name (AR)') || 'ديوان للاستشارات'}</p>
          </div>
          
          <div style="padding: 20px; background: #f5f5f5;">
            <p>Sayın <strong>${data.contactPerson || data.clientName}</strong>,</p>
            
            <p><strong>${data.period}</strong> dönemine ait faturanız ekte sunulmuştur.</p>
            
            <div style="background: white; padding: 15px; border-radius: 5px; margin: 20px 0;">
              <h3 style="margin-top: 0; color: #1565c0;">Fatura Detayları:</h3>
              <table style="width: 100%;">
                <tr><td><strong>Fatura No:</strong></td><td>${data.invoiceNo}</td></tr>
                <tr><td><strong>Tarih:</strong></td><td>${data.invoiceDate}</td></tr>
                <tr><td><strong>Tutar:</strong></td><td style="font-size: 18px; color: #1565c0;"><strong>${data.amount}</strong></td></tr>
              </table>
            </div>
            
            <div style="background: white; padding: 15px; border-radius: 5px;">
              <h3 style="margin-top: 0; color: #1565c0;">Ödeme Bilgileri:</h3>
              <p><strong>Banka:</strong> ${getSettingValue('Bank Name') || 'Kuveyt Türk'}</p>
              <p><strong>IBAN (TRY):</strong> ${getSettingValue('IBAN TRY') || ''}</p>
              <p><strong>IBAN (USD):</strong> ${getSettingValue('IBAN USD') || ''}</p>
              <p><strong>SWIFT:</strong> ${getSettingValue('SWIFT Code') || 'KTEFTRIS'}</p>
            </div>
            
            <p style="margin-top: 20px;">Herhangi bir sorunuz varsa lütfen bizimle iletişime geçin.</p>
            
            <p>Saygılarımızla,<br>
            <strong>${getSettingValue('Company Name (EN)') || 'Dewan Consulting'}</strong><br>
            <a href="mailto:sales@aldewan.net">sales@aldewan.net</a></p>
          </div>
          
          <div style="background: #333; color: white; padding: 10px; text-align: center; font-size: 12px;">
            İlginiz için teşekkür ederiz!
          </div>
        </div>
      `
    },
    
    'AR': {
      subject: 'فاتورة ' + data.invoiceNo + ' - ' + (getSettingValue('Company Name (AR)') || 'ديوان للاستشارات'),
      body: `
        <div dir="rtl" style="font-family: Arial, sans-serif; padding: 20px; max-width: 600px;">
          <div style="background: #1565c0; color: white; padding: 20px; text-align: center;">
            <h1 style="margin: 0;">${getSettingValue('Company Name (AR)') || 'ديوان للاستشارات'}</h1>
            <p style="margin: 5px 0 0 0;">${getSettingValue('Company Name (EN)') || 'Dewan Consulting'}</p>
          </div>
          
          <div style="padding: 20px; background: #f5f5f5;">
            <p>السيد/السيدة <strong>${data.contactPerson || data.clientName}</strong> المحترم/ة،</p>
            
            <p>مرفق لكم فاتورة <strong>${data.period}</strong>.</p>
            
            <div style="background: white; padding: 15px; border-radius: 5px; margin: 20px 0;">
              <h3 style="margin-top: 0; color: #1565c0;">تفاصيل الفاتورة:</h3>
              <table style="width: 100%;">
                <tr><td><strong>رقم الفاتورة:</strong></td><td>${data.invoiceNo}</td></tr>
                <tr><td><strong>التاريخ:</strong></td><td>${data.invoiceDate}</td></tr>
                <tr><td><strong>المبلغ:</strong></td><td style="font-size: 18px; color: #1565c0;"><strong>${data.amount}</strong></td></tr>
              </table>
            </div>
            
            <div style="background: white; padding: 15px; border-radius: 5px;">
              <h3 style="margin-top: 0; color: #1565c0;">معلومات الدفع:</h3>
              <p><strong>البنك:</strong> ${getSettingValue('Bank Name') || 'Kuveyt Türk'}</p>
              <p><strong>IBAN (TRY):</strong> ${getSettingValue('IBAN TRY') || ''}</p>
              <p><strong>IBAN (USD):</strong> ${getSettingValue('IBAN USD') || ''}</p>
              <p><strong>SWIFT:</strong> ${getSettingValue('SWIFT Code') || 'KTEFTRIS'}</p>
            </div>
            
            <p style="margin-top: 20px;">في حال وجود أي استفسار، لا تترددوا في التواصل معنا.</p>
            
            <p>مع أطيب التحيات،<br>
            <strong>${getSettingValue('Company Name (AR)') || 'ديوان للاستشارات'}</strong><br>
            <a href="mailto:sales@aldewan.net">sales@aldewan.net</a></p>
          </div>
          
          <div style="background: #333; color: white; padding: 10px; text-align: center; font-size: 12px;">
            شكراً لتعاملكم معنا!
          </div>
        </div>
      `
    }
  };
  
  return templates[language] || templates['EN'];
}

// ==================== 5. SEND SINGLE INVOICE EMAIL ====================

/**
 * إرسال فاتورة واحدة بالإيميل
 */
function sendInvoiceEmail(invoiceNo, sentBy) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get invoice data from Invoice Log
  const logSheet = ss.getSheetByName('Invoice Log');
  if (!logSheet) return { success: false, error: 'Invoice Log not found' };
  
  const logData = logSheet.getDataRange().getValues();
  let invoiceRow = -1;
  let invoiceData = null;
  
  for (let i = 1; i < logData.length; i++) {
    if (logData[i][0] === invoiceNo) {
      invoiceRow = i + 1;
      invoiceData = {
        invoiceNo: logData[i][0],
        invoiceDate: logData[i][1],
        clientCode: logData[i][2],
        clientName: logData[i][3],
        service: logData[i][4],
        period: logData[i][5],
        amount: logData[i][6],
        currency: logData[i][7],
        pdfLink: logData[i][9]
      };
      break;
    }
  }
  
  if (!invoiceData) {
    return { success: false, error: 'Invoice not found: ' + invoiceNo };
  }
  
  // Get client data
  const clientData = getClientData(invoiceData.clientCode);
  if (!clientData) {
    logEmail(invoiceNo, invoiceData.clientCode, invoiceData.clientName, '', '', 'Failed', 'Client not found', sentBy, '');
    return { success: false, error: 'Client not found: ' + invoiceData.clientCode };
  }
  
  const clientEmail = clientData.email;
  const clientLanguage = clientData.language || 'EN';
  const contactPerson = clientData.contactPerson || '';
  
  if (!clientEmail) {
    logEmail(invoiceNo, invoiceData.clientCode, invoiceData.clientName, '', clientLanguage, 'Failed', 'No email address', sentBy, '');
    return { success: false, error: 'No email for client: ' + invoiceData.clientCode };
  }
  
  // Get email template
  const template = getEmailTemplate(clientLanguage, {
    invoiceNo: invoiceData.invoiceNo,
    invoiceDate: Utilities.formatDate(new Date(invoiceData.invoiceDate), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    clientName: invoiceData.clientName,
    contactPerson: contactPerson,
    period: invoiceData.period,
    amount: formatCurrency(invoiceData.amount, invoiceData.currency)
  });
  
  try {
    // Get PDF file if exists
    let attachments = [];
    if (invoiceData.pdfLink) {
      try {
        const fileId = extractFileIdFromUrl(invoiceData.pdfLink);
        if (fileId) {
          const file = DriveApp.getFileById(fileId);
          attachments.push(file.getAs(MimeType.PDF));
        }
      } catch (e) {
        console.log('Could not attach PDF: ' + e.message);
      }
    }
    
    // Send email
    GmailApp.sendEmail(clientEmail, template.subject, '', {
      htmlBody: template.body,
      name: getSettingValue('Company Name (EN)') || 'Dewan Consulting',
      replyTo: 'sales@aldewan.net',
      attachments: attachments
    });
    
    // Update Invoice Log
    logSheet.getRange(invoiceRow, 9).setValue('Sent');
    logSheet.getRange(invoiceRow, 12).setValue('Sent');
    logSheet.getRange(invoiceRow, 13).setValue(new Date());
    
    // Log to Email Log
    logEmail(invoiceData.invoiceNo, invoiceData.clientCode, invoiceData.clientName, 
             clientEmail, clientLanguage, 'Sent', '', sentBy, invoiceData.pdfLink);
    
    return { success: true, email: clientEmail };
    
  } catch (e) {
    // Log failure
    logEmail(invoiceData.invoiceNo, invoiceData.clientCode, invoiceData.clientName,
             clientEmail, clientLanguage, 'Failed', e.message, sentBy, invoiceData.pdfLink);
    
    // Update Invoice Log
    logSheet.getRange(invoiceRow, 12).setValue('Failed');
    
    return { success: false, error: e.message };
  }
}

/**
 * استخراج File ID من رابط Drive
 */
function extractFileIdFromUrl(url) {
  if (!url) return null;
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

// ==================== 6. SEND PENDING INVOICES ====================

/**
 * إرسال كل الفواتير المعلقة (بعد 3 أيام عمل)
 */
function sendPendingInvoices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const logSheet = ss.getSheetByName('Invoice Log');
  
  if (!logSheet || logSheet.getLastRow() < 2) {
    ui.alert('⚠️ No invoices in Invoice Log!');
    return;
  }
  
  const data = logSheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  let readyToSend = [];
  let notReady = [];
  
  for (let i = 1; i < data.length; i++) {
    const invoiceNo = data[i][0];
    const invoiceDate = new Date(data[i][1]);
    const sendEmail = data[i][10];
    const emailStatus = data[i][11];
    
    if (sendEmail !== 'Yes' || emailStatus === 'Sent') continue;
    
    // Calculate send date (3 working days after invoice date)
    const sendDate = calculateSendDate(invoiceDate);
    sendDate.setHours(0, 0, 0, 0);
    
    if (today >= sendDate) {
      readyToSend.push({
        row: i + 1,
        invoiceNo: invoiceNo,
        clientName: data[i][3],
        sendDate: sendDate
      });
    } else {
      notReady.push({
        invoiceNo: invoiceNo,
        clientName: data[i][3],
        sendDate: sendDate
      });
    }
  }
  
  if (readyToSend.length === 0 && notReady.length === 0) {
    ui.alert('✅ No pending invoices to send!');
    return;
  }
  
  let message = '📧 Pending Invoices Report\n\n';
  
  if (readyToSend.length > 0) {
    message += '✅ Ready to Send (' + readyToSend.length + '):\n';
    readyToSend.forEach(inv => {
      message += '• ' + inv.invoiceNo + ' - ' + inv.clientName + '\n';
    });
    message += '\n';
  }
  
  if (notReady.length > 0) {
    message += '⏳ Not Ready Yet (' + notReady.length + '):\n';
    notReady.slice(0, 5).forEach(inv => {
      message += '• ' + inv.invoiceNo + ' - Send: ' + Utilities.formatDate(inv.sendDate, Session.getScriptTimeZone(), 'dd/MM/yyyy') + '\n';
    });
    if (notReady.length > 5) {
      message += '... and ' + (notReady.length - 5) + ' more\n';
    }
    message += '\n';
  }
  
  if (readyToSend.length === 0) {
    ui.alert(message + 'No invoices ready to send yet.');
    return;
  }
  
  const confirm = ui.alert(
    message + 'Send ' + readyToSend.length + ' invoices now?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  let sent = 0, failed = 0;
  const results = [];
  
  readyToSend.forEach(inv => {
    const result = sendInvoiceEmail(inv.invoiceNo, 'Manual');
    
    if (result.success) {
      sent++;
      results.push({ invoice: inv.invoiceNo, status: '✅ Sent', detail: result.email });
    } else {
      failed++;
      results.push({ invoice: inv.invoiceNo, status: '❌ Failed', detail: result.error });
    }
  });
  
  let report = '📧 Send Complete!\n\n';
  report += '✅ Sent: ' + sent + '\n';
  report += '❌ Failed: ' + failed + '\n\n';
  
  if (results.length <= 10) {
    report += 'Details:\n';
    results.forEach(r => {
      report += r.status + ' ' + r.invoice + ': ' + r.detail + '\n';
    });
  }
  
  ui.alert(report);
}

/**
 * إرسال فاتورة محددة يدوياً
 */
function sendSelectedInvoice() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();
  
  if (sheet.getName() !== 'Invoice Log') {
    ui.alert('⚠️ Please go to Invoice Log sheet and select an invoice row.');
    return;
  }
  
  const row = sheet.getActiveCell().getRow();
  if (row < 2) {
    ui.alert('⚠️ Please select an invoice row (not header).');
    return;
  }
  
  const invoiceNo = sheet.getRange(row, 1).getValue();
  const clientName = sheet.getRange(row, 4).getValue();
  const emailStatus = sheet.getRange(row, 12).getValue();
  
  if (emailStatus === 'Sent') {
    const resend = ui.alert(
      '⚠️ Invoice Already Sent!\n\n' +
      'Invoice: ' + invoiceNo + '\n' +
      'Client: ' + clientName + '\n\n' +
      'Send again?',
      ui.ButtonSet.YES_NO
    );
    if (resend !== ui.Button.YES) return;
  }
  
  const confirm = ui.alert(
    '📧 Send Invoice Email\n\n' +
    'Invoice: ' + invoiceNo + '\n' +
    'Client: ' + clientName + '\n\n' +
    'Send now?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  const result = sendInvoiceEmail(invoiceNo, 'Manual');
  
  if (result.success) {
    ui.alert('✅ Email sent successfully!\n\nTo: ' + result.email);
  } else {
    ui.alert('❌ Failed to send email:\n\n' + result.error);
  }
}

// ==================== 7. SHOW EMAIL LOG ====================
function showEmailLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Email Log');
  if (!sheet) {
    sheet = createEmailLogSheet(ss);
  }
  ss.setActiveSheet(sheet);
}

// ==================== 8. EMAIL STATISTICS ====================
function showEmailStatistics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Email Log');
  
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('📊 No email data yet.');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  let sent = 0, failed = 0;
  const byLanguage = { EN: 0, TR: 0, AR: 0 };
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] === 'Sent') {
      sent++;
      const lang = data[i][5] || 'EN';
      byLanguage[lang] = (byLanguage[lang] || 0) + 1;
    } else if (data[i][6] === 'Failed') {
      failed++;
    }
  }
  
  ui.alert(
    '📊 Email Statistics\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '✅ Sent: ' + sent + '\n' +
    '❌ Failed: ' + failed + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'By Language:\n' +
    '🇬🇧 English: ' + (byLanguage.EN || 0) + '\n' +
    '🇹🇷 Turkish: ' + (byLanguage.TR || 0) + '\n' +
    '🇸🇦 Arabic: ' + (byLanguage.AR || 0)
  );
}

// ==================== 9. TRIGGERS MANAGEMENT ====================

/**
 * إعداد الـ Triggers التلقائية
 */
function setupAutoTriggers() {
  const ui = SpreadsheetApp.getUi();
  
  const confirm = ui.alert(
    '⏰ Setup Automatic Triggers\n\n' +
    'This will create:\n\n' +
    '1. 📅 Monthly Invoice Generation\n' +
    '   Day 25 at 9:00 AM (or next working day)\n\n' +
    '2. 📧 Daily Email Check\n' +
    '   Every day at 10:00 AM\n' +
    '   (Sends invoices after 3 working days)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  // Remove existing triggers
  removeAllTriggers(true);
  
  try {
    // Monthly invoice generation - day 25 at 9:00 AM
    ScriptApp.newTrigger('autoGenerateMonthlyInvoices')
      .timeBased()
      .onMonthDay(25)
      .atHour(9)
      .create();
    
    // Daily email check - 10:00 AM
    ScriptApp.newTrigger('autoSendPendingInvoices')
      .timeBased()
      .everyDays(1)
      .atHour(10)
      .create();
    
    ui.alert(
      '✅ Triggers Created!\n\n' +
      '📅 Monthly Invoices: Day 25 at 9:00 AM\n' +
      '📧 Email Check: Daily at 10:00 AM\n\n' +
      'The system will:\n' +
      '• Generate invoices on day 25\n' +
      '• Wait 3 working days\n' +
      '• Send emails automatically\n\n' +
      'Holidays and weekends are respected.'
    );
    
  } catch (error) {
    ui.alert('❌ Error creating triggers:\n\n' + error.message);
  }
}

/**
 * إزالة كل الـ Triggers
 */
function removeAllTriggers(silent) {
  const triggers = ScriptApp.getProjectTriggers();
  
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  
  if (!silent) {
    SpreadsheetApp.getUi().alert('✅ All triggers removed!\n\nRemoved: ' + triggers.length + ' triggers');
  }
}

/**
 * عرض حالة الـ Triggers
 */
function showTriggersStatus() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  
  if (triggers.length === 0) {
    ui.alert('⚠️ No active triggers.\n\nUse "Setup Auto Triggers" to create them.');
    return;
  }
  
  let status = '⏰ Active Triggers:\n\n';
  
  triggers.forEach(trigger => {
    const funcName = trigger.getHandlerFunction();
    const type = trigger.getEventType();
    status += '• ' + funcName + ' (' + type + ')\n';
  });
  
  ui.alert(status);
}

// ==================== 10. AUTO TRIGGER FUNCTIONS ====================

/**
 * إصدار الفواتير الشهرية تلقائياً
 */
function autoGenerateMonthlyInvoices() {
  const today = new Date();
  
  // Check if today is a working day
  if (isHolidayOrWeekend(today)) {
    return;
  }
  
  // Check if this is the correct generation day
  const expectedDate = getInvoiceGenerationDate(today.getFullYear(), today.getMonth());
  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const expectedStr = Utilities.formatDate(expectedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  if (todayStr !== expectedStr) {
    return;
  }
  
  // Generate all monthly invoices
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const clients = getActiveClients().filter(c => c.monthlyFee > 0);
    
    if (clients.length === 0) return;
    
    const invoiceDate = new Date();
    const period = Utilities.formatDate(invoiceDate, Session.getScriptTimeZone(), 'MMMM yyyy');
    
    clients.forEach(client => {
      const invoiceNo = getNextInvoiceNumber();
      const clientData = getClientData(client.code);
      
      // Fill template
      fillInvoiceTemplate(ss, {
        invoiceNo: invoiceNo,
        invoiceDate: invoiceDate,
        clientName: client.nameEN,
        clientNameAR: clientData ? clientData.nameAR : '',
        taxNumber: clientData ? clientData.taxNumber : '',
        address: clientData ? clientData.address : '',
        period: period,
        items: [{
          description: 'Monthly Consulting (استشارات شهرية)',
          qty: 1,
          unitPrice: client.monthlyFee,
          total: client.monthlyFee
        }],
        currency: client.feeCurrency,
        subtotal: client.monthlyFee,
        vat: 0,
        vatRate: 0,
        total: client.monthlyFee
      });
      
      // Save PDF
      let pdfUrl = '';
      if (clientData && clientData.folderId) {
        try {
          const pdfResult = createInvoicePDF(invoiceNo, clientData.folderId);
          pdfUrl = pdfResult.url;
        } catch (e) {
          console.log('PDF error: ' + e.message);
        }
      }
      
      // Log invoice
      logInvoice({
        invoiceNo: invoiceNo,
        invoiceDate: invoiceDate,
        clientCode: client.code,
        clientName: client.nameEN,
        service: 'Monthly Consulting (استشارات شهرية)',
        period: period,
        amount: client.monthlyFee,
        currency: client.feeCurrency,
        status: 'Issued',
        pdfLink: pdfUrl,
        sendEmail: 'Yes',
        emailStatus: 'Pending',
        transCode: ''
      });
      
      // Record transaction
      recordInvoiceTransaction(invoiceNo, client.code, client.nameEN, client.monthlyFee, client.feeCurrency, 'Monthly Consulting (استشارات شهرية)');
      
      incrementInvoiceNumber();
    });
    
    // Log the action
    logAlert('Auto Invoice', clients.length + ' monthly invoices generated for ' + period, 'Info');
    
  } catch (error) {
    console.error('Auto generate error: ' + error);
    logAlert('Auto Invoice Error', error.message, 'Error');
  }
}

/**
 * إرسال الفواتير المعلقة تلقائياً
 */
function autoSendPendingInvoices() {
  const today = new Date();
  
  // Skip weekends
  if (today.getDay() === 0 || today.getDay() === 6) {
    return;
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('Invoice Log');
    
    if (!logSheet || logSheet.getLastRow() < 2) return;
    
    const data = logSheet.getDataRange().getValues();
    let sent = 0, failed = 0;
    
    for (let i = 1; i < data.length; i++) {
      const invoiceNo = data[i][0];
      const invoiceDate = new Date(data[i][1]);
      const sendEmail = data[i][10];
      const emailStatus = data[i][11];
      
      if (sendEmail !== 'Yes' || emailStatus === 'Sent') continue;
      
      // Check if 3 working days have passed
      const sendDate = calculateSendDate(invoiceDate);
      
      if (today >= sendDate) {
        const result = sendInvoiceEmail(invoiceNo, 'Auto');
        
        if (result.success) {
          sent++;
        } else {
          failed++;
        }
      }
    }
    
    if (sent > 0 || failed > 0) {
      logAlert('Auto Email', 'Sent: ' + sent + ', Failed: ' + failed, sent > 0 ? 'Info' : 'Warning');
    }
    
  } catch (error) {
    console.error('Auto send error: ' + error);
    logAlert('Auto Email Error', error.message, 'Error');
  }
}

// ==================== 11. ALERTS LOG ====================

function createAlertsLogSheet(ss) {
  let sheet = ss.getSheetByName('Alerts Log');
  if (sheet) return sheet;
  
  sheet = ss.insertSheet('Alerts Log');
  sheet.setTabColor('#ff9800');
  
  const headers = ['Date/Time', 'Type', 'Message', 'Severity', 'Acknowledged'];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#e65100')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  const widths = [150, 120, 400, 80, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  // Conditional formatting
  const severityRange = sheet.getRange(2, 4, 500, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Error').setBackground('#ffcdd2').setRanges([severityRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Warning').setBackground('#fff9c4').setRanges([severityRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Info').setBackground('#c8e6c9').setRanges([severityRange]).build()
  ]);
  
  sheet.setFrozenRows(1);
  
  return sheet;
}

function logAlert(type, message, severity) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Alerts Log');
  if (!sheet) {
    sheet = createAlertsLogSheet(ss);
  }
  
  const lastRow = sheet.getLastRow() + 1;
  
  sheet.getRange(lastRow, 1).setValue(new Date()).setNumberFormat('yyyy-mm-dd HH:mm');
  sheet.getRange(lastRow, 2).setValue(type);
  sheet.getRange(lastRow, 3).setValue(message);
  sheet.getRange(lastRow, 4).setValue(severity || 'Info');
  sheet.getRange(lastRow, 5).setValue('No');
}

function showAlertsLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Alerts Log');
  if (!sheet) {
    sheet = createAlertsLogSheet(ss);
  }
  ss.setActiveSheet(sheet);
}

// ==================== 12. OVERDUE REMINDERS ====================

/**
 * التحقق من المدفوعات المتأخرة
 */
function checkOverduePayments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const transSheet = ss.getSheetByName('Transactions');
  if (!transSheet || transSheet.getLastRow() < 2) {
    ui.alert('⚠️ No transactions found!');
    return;
  }
  
  const data = transSheet.getDataRange().getValues();
  const today = new Date();
  const overdueList = [];
  
  const reminderDays = parseInt(getSettingValue('First Reminder (Days)')) || 30;
  
  for (let i = 1; i < data.length; i++) {
    const status = data[i][18]; // Status
    const dueDate = data[i][19]; // Due Date
    const amount = data[i][10]; // Amount
    const clientName = data[i][5]; // Client Name
    const invoiceNo = data[i][17]; // Invoice No
    const currency = data[i][11]; // Currency
    
    if (status && status.includes('Pending') && dueDate) {
      const due = new Date(dueDate);
      const diffDays = Math.floor((today - due) / (1000 * 60 * 60 * 24));
      
      if (diffDays > 0) {
        overdueList.push({
          row: i + 1,
          clientName: clientName,
          invoiceNo: invoiceNo || 'N/A',
          amount: formatCurrency(amount, currency),
          dueDate: Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          daysOverdue: diffDays
        });
      }
    }
  }
  
  if (overdueList.length === 0) {
    ui.alert('✅ No overdue payments!\n\nAll pending payments are still within their due dates.');
    return;
  }
  
  // Sort by days overdue
  overdueList.sort((a, b) => b.daysOverdue - a.daysOverdue);
  
  let report = '⚠️ Overdue Payments Report\n\n';
  report += 'Found ' + overdueList.length + ' overdue payments:\n\n';
  
  overdueList.slice(0, 10).forEach(o => {
    report += '• ' + o.clientName + '\n';
    report += '  Invoice: ' + o.invoiceNo + ' | ' + o.amount + '\n';
    report += '  Overdue: ' + o.daysOverdue + ' days\n\n';
  });
  
  if (overdueList.length > 10) {
    report += '... and ' + (overdueList.length - 10) + ' more\n';
  }
  
  ui.alert(report);
}

// ==================== END OF PART 7 ====================
// ╔════════════════════════════════════════════════════════════════════════════╗
// ║                    DC CONSULTING ACCOUNTING SYSTEM v3.0                     ║
// ║                              Part 8 of 9                                    ║
// ║                        Reports + Dashboard                                  ║
// ╚════════════════════════════════════════════════════════════════════════════╝

// ==================== 1. CREATE DASHBOARD SHEET ====================
function createDashboardSheet(ss) {
  let sheet = ss.getSheetByName('Dashboard');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Dashboard');
  sheet.setTabColor('#2196f3');
  
  // Title
  sheet.getRange('A1:H1').merge()
    .setValue('📊 DC CONSULTING DASHBOARD')
    .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#1565c0').setFontColor('#ffffff');
  
  sheet.getRange('A2:H2').merge()
    .setValue('Last Updated: ' + formatDate(new Date(), 'yyyy-MM-dd HH:mm'))
    .setHorizontalAlignment('center').setFontStyle('italic');
  
  // Section: Cash & Bank Balances
  sheet.getRange('A4:D4').merge()
    .setValue('💰 Cash & Bank Balances')
    .setFontSize(14).setFontWeight('bold').setBackground('#e3f2fd');
  
  sheet.getRange('A5:D5')
    .setValues([['Account', 'Currency', 'Balance', 'Status']])
    .setFontWeight('bold').setBackground('#bbdefb');
  
  // Section: Monthly Summary
  sheet.getRange('F4:H4').merge()
    .setValue('📈 Monthly Summary')
    .setFontSize(14).setFontWeight('bold').setBackground('#e8f5e9');
  
  sheet.getRange('F5:H5')
    .setValues([['Metric', 'This Month', 'Last Month']])
    .setFontWeight('bold').setBackground('#c8e6c9');
  
  // Section: Client Statistics
  sheet.getRange('A20:D20').merge()
    .setValue('👥 Client Statistics')
    .setFontSize(14).setFontWeight('bold').setBackground('#fff3e0');
  
  // Section: Overdue Alerts
  sheet.getRange('F20:H20').merge()
    .setValue('⚠️ Overdue Alerts')
    .setFontSize(14).setFontWeight('bold').setBackground('#ffebee');
  
  // Column widths
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 30); // spacer
  sheet.setColumnWidth(6, 180);
  sheet.setColumnWidth(7, 120);
  sheet.setColumnWidth(8, 120);
  
  sheet.setFrozenRows(2);
  
  return sheet;
}

// ==================== 2. SHOW/REFRESH DASHBOARD ====================
function showDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Dashboard');
  
  if (!sheet) {
    sheet = createDashboardSheet(ss);
  }
  
  refreshDashboard();
  ss.setActiveSheet(sheet);
}

function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Dashboard');
  
  if (!sheet) return;
  
  // Update timestamp
  sheet.getRange('A2:H2').merge()
    .setValue('Last Updated: ' + formatDate(new Date(), 'yyyy-MM-dd HH:mm'));
  
  // Clear old data
  sheet.getRange('A6:D18').clearContent();
  sheet.getRange('F6:H18').clearContent();
  sheet.getRange('A21:D30').clearContent();
  sheet.getRange('F21:H30').clearContent();
  
  // ===== Cash & Bank Balances =====
  let row = 6;
  
  // Cash Boxes
  const cashBoxes = getCashBoxesList();
  cashBoxes.forEach(cash => {
    const balance = getCashBankBalance(cash.sheetName);
    sheet.getRange(row, 1).setValue('💰 ' + cash.name);
    sheet.getRange(row, 2).setValue(cash.currency);
    sheet.getRange(row, 3).setValue(balance).setNumberFormat('#,##0.00');
    sheet.getRange(row, 4).setValue(balance >= 0 ? '✅' : '⚠️');
    row++;
  });
  
  // Bank Accounts
  const bankAccounts = getBankAccountsList();
  bankAccounts.forEach(bank => {
    const balance = getCashBankBalance(bank.sheetName);
    sheet.getRange(row, 1).setValue('🏦 ' + bank.name);
    sheet.getRange(row, 2).setValue(bank.currency);
    sheet.getRange(row, 3).setValue(balance).setNumberFormat('#,##0.00');
    sheet.getRange(row, 4).setValue(balance >= 0 ? '✅' : '⚠️');
    row++;
  });
  
  // ===== Monthly Summary =====
  const transSheet = ss.getSheetByName('Transactions');
  if (transSheet && transSheet.getLastRow() > 1) {
    const transData = transSheet.getDataRange().getValues();
    const now = new Date();
    const thisMonth = now.getMonth();
    const lastMonth = thisMonth === 0 ? 11 : thisMonth - 1;
    const thisYear = now.getFullYear();
    const lastMonthYear = thisMonth === 0 ? thisYear - 1 : thisYear;
    
    let thisMonthRevenue = 0, lastMonthRevenue = 0;
    let thisMonthExpense = 0, lastMonthExpense = 0;
    let thisMonthInvoices = 0, lastMonthInvoices = 0;
    
    for (let i = 1; i < transData.length; i++) {
      const date = transData[i][1]; // Date
      const movementType = transData[i][2]; // Movement Type
      const amount = parseFloat(transData[i][13]) || 0; // Amount TRY
      
      if (!date) continue;
      
      const transDate = new Date(date);
      const transMonth = transDate.getMonth();
      const transYear = transDate.getFullYear();
      
      if (transYear === thisYear && transMonth === thisMonth) {
        if (movementType && movementType.includes('Revenue')) {
          thisMonthRevenue += amount;
        } else if (movementType && (movementType.includes('Expense') || movementType.includes('مصروف'))) {
          thisMonthExpense += amount;
        }
      }
      
      if (transYear === lastMonthYear && transMonth === lastMonth) {
        if (movementType && movementType.includes('Revenue')) {
          lastMonthRevenue += amount;
        } else if (movementType && (movementType.includes('Expense') || movementType.includes('مصروف'))) {
          lastMonthExpense += amount;
        }
      }
    }
    
    // Invoice counts
    const invoiceSheet = ss.getSheetByName('Invoice Log');
    if (invoiceSheet && invoiceSheet.getLastRow() > 1) {
      const invData = invoiceSheet.getDataRange().getValues();
      for (let i = 1; i < invData.length; i++) {
        const invDate = invData[i][1];
        if (!invDate) continue;
        const d = new Date(invDate);
        if (d.getFullYear() === thisYear && d.getMonth() === thisMonth) thisMonthInvoices++;
        if (d.getFullYear() === lastMonthYear && d.getMonth() === lastMonth) lastMonthInvoices++;
      }
    }
    
    // Fill summary
    const summaryData = [
      ['Total Revenue (TRY)', thisMonthRevenue, lastMonthRevenue],
      ['Total Expenses (TRY)', thisMonthExpense, lastMonthExpense],
      ['Net Income (TRY)', thisMonthRevenue - thisMonthExpense, lastMonthRevenue - lastMonthExpense],
      ['Invoices Issued', thisMonthInvoices, lastMonthInvoices]
    ];
    
    sheet.getRange(6, 6, summaryData.length, 3).setValues(summaryData);
    sheet.getRange(6, 7, 3, 2).setNumberFormat('#,##0.00');
    
    // Highlight net income
    const netRow = 8;
    const netThis = thisMonthRevenue - thisMonthExpense;
    sheet.getRange(netRow, 7).setBackground(netThis >= 0 ? '#c8e6c9' : '#ffcdd2');
  }
  
  // ===== Client Statistics =====
  const clients = getActiveClients();
  const totalClients = clients.length;
  const clientsWithFee = clients.filter(c => c.monthlyFee > 0).length;
  const totalMonthlyRevenue = clients.reduce((sum, c) => sum + (c.monthlyFee || 0), 0);
  
  const clientStats = [
    ['Total Active Clients', totalClients],
    ['Clients with Monthly Fee', clientsWithFee],
    ['Total Monthly Revenue (TRY)', totalMonthlyRevenue]
  ];
  
  sheet.getRange(21, 1, clientStats.length, 2).setValues(clientStats);
  sheet.getRange(23, 2).setNumberFormat('#,##0.00');
  
  // ===== Overdue Alerts =====
  if (transSheet && transSheet.getLastRow() > 1) {
    const transData = transSheet.getDataRange().getValues();
    const today = new Date();
    const firstReminderDays = parseInt(getSettingValue('First Reminder (Days)')) || 7;
    
    let overdueCount = 0;
    let overdueAmount = 0;
    
    for (let i = 1; i < transData.length; i++) {
      const status = transData[i][18];
      const dueDate = transData[i][19];
      const amount = parseFloat(transData[i][13]) || 0;
      
      if (status && status.includes('Pending') && dueDate) {
        const due = new Date(dueDate);
        const diffDays = Math.floor((today - due) / (1000 * 60 * 60 * 24));
        
        if (diffDays >= firstReminderDays) {
          overdueCount++;
          overdueAmount += amount;
        }
      }
    }
    
    const alertData = [
      ['Overdue Invoices', overdueCount],
      ['Overdue Amount (TRY)', overdueAmount],
      ['Status', overdueCount > 0 ? '⚠️ Action Required' : '✅ All Clear']
    ];
    
    sheet.getRange(21, 6, alertData.length, 2).setValues(alertData);
    sheet.getRange(22, 7).setNumberFormat('#,##0.00');
    
    if (overdueCount > 0) {
      sheet.getRange(21, 6, 3, 2).setBackground('#ffebee');
    } else {
      sheet.getRange(21, 6, 3, 2).setBackground('#e8f5e9');
    }
  }
}

// ==================== 3. CLIENT STATEMENT ====================
/**
 * ✅ محدّث: إدخال كود العميل مباشرة بدون قائمة
 */
function showClientStatement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // طلب كود العميل مباشرة
  const response = ui.prompt(
    '📄 Client Statement (كشف حساب عميل)',
    'أدخل كود العميل (Client Code):\n\nمثال: CLT-001',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const clientCode = response.getResponseText().trim();
  if (!clientCode) {
    ui.alert('⚠️ يرجى إدخال كود العميل!');
    return;
  }

  // البحث عن العميل بالكود
  const client = getClientData(clientCode);
  if (!client) {
    ui.alert('⚠️ لم يتم العثور على عميل بهذا الكود!\n\nClient Code: ' + clientCode);
    return;
  }

  generateClientStatement(client.code, client.nameEN);
}

/**
 * ✅ محدّث: كشف حساب بصيغة دائن/مدين/رصيد
 */
function generateClientStatement(clientCode, clientName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const transSheet = ss.getSheetByName('Transactions');
  if (!transSheet || transSheet.getLastRow() < 2) {
    ui.alert('⚠️ No transactions found!');
    return;
  }

  const transData = transSheet.getDataRange().getValues();

  // Filter transactions for this client where Show in Statement = Yes
  const clientTrans = [];

  for (let i = 1; i < transData.length; i++) {
    const code = transData[i][4]; // Client Code
    const name = transData[i][5]; // Client Name
    const showInStatement = transData[i][24]; // Column Y

    if ((code === clientCode || name === clientName) &&
        (!showInStatement || showInStatement.includes('Yes'))) {

      const movementType = transData[i][2] || '';
      const amount = parseFloat(transData[i][10]) || 0;
      const item = transData[i][6] || '';
      const description = transData[i][7] || '';

      // تحديد دائن/مدين بناءً على نوع الحركة
      let credit = 0; // له (دائن) - ما يستحق على العميل
      let debit = 0;  // عليه (مدين) - ما دفعه العميل

      // استحقاق إيراد = له (دائن) - فاتورة مستحقة على العميل
      if (movementType.includes('Revenue Accrual') || movementType.includes('استحقاق إيراد')) {
        credit = amount;
      }
      // تحصيل إيراد = عليه (مدين) - دفعة من العميل
      else if (movementType.includes('Revenue Collection') || movementType.includes('تحصيل إيراد')) {
        debit = amount;
      }

      // فقط نضيف الحركات ذات القيمة
      if (credit > 0 || debit > 0) {
        clientTrans.push({
          date: transData[i][1],
          description: description || item || movementType, // Description first
          credit: credit,
          debit: debit
        });
      }
    }
  }

  if (clientTrans.length === 0) {
    ui.alert('ℹ️ No statement items found for this client.');
    return;
  }

  // ترتيب حسب التاريخ
  clientTrans.sort((a, b) => new Date(a.date) - new Date(b.date));

  // حساب الإجماليات
  let totalCredit = 0, totalDebit = 0;
  clientTrans.forEach(t => {
    totalCredit += t.credit;
    totalDebit += t.debit;
  });
  const balance = totalCredit - totalDebit;

  // Show summary
  const summary =
    '📄 كشف حساب العميل: ' + clientName + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'عدد الحركات: ' + clientTrans.length + '\n' +
    'إجمالي له (دائن): ' + formatCurrency(totalCredit, 'TRY') + '\n' +
    'إجمالي عليه (مدين): ' + formatCurrency(totalDebit, 'TRY') + '\n' +
    '───────────────────────────\n' +
    'الرصيد المستحق: ' + formatCurrency(balance, 'TRY') + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'هل تريد تصدير الكشف إلى شيت؟';

  const exportConfirm = ui.alert(summary, ui.ButtonSet.YES_NO);

  if (exportConfirm === ui.Button.YES) {
    exportClientStatement(clientCode, clientName, clientTrans, {
      totalCredit: totalCredit,
      totalDebit: totalDebit,
      balance: balance
    });
  }
}

/**
 * ✅ محدّث: تصدير كشف الحساب بتصميم احترافي
 * - ترويسة الشركة مع اللوجو
 * - معلومات العميل
 * - جدول الحركات بصيغة دائن/مدين/رصيد
 */
function exportClientStatement(clientCode, clientName, transactions, totals) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Statement - ' + clientName.substring(0, 20);

  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);

  sheet = ss.insertSheet(sheetName);
  sheet.setTabColor('#1565c0');

  // ═══════════════════════════════════════════════════════════════════════════
  // الحصول على بيانات الشركة من الإعدادات
  // ═══════════════════════════════════════════════════════════════════════════
  const companyNameEN = getSettingValue('Company Name (EN)') || 'Dewan Consulting';
  const companyNameAR = getSettingValue('Company Name (AR)') || 'ديوان للاستشارات';
  const companyAddress = getSettingValue('Company Address') || '';
  const companyPhone = getSettingValue('Company Phone') || '';
  const companyEmail = getSettingValue('Company Email') || '';
  const companyLogo = getSettingValue('Company Logo URL') || '';

  let currentRow = 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // HEADER SECTION - ترويسة الشركة
  // ═══════════════════════════════════════════════════════════════════════════

  // Insert logo if URL is provided
  if (companyLogo && companyLogo.trim() !== '') {
    try {
      // Create a formula to display the image
      sheet.getRange('A1:A4').merge();
      sheet.getRange('A1').setFormula('=IMAGE("' + companyLogo + '", 2)');
      sheet.setColumnWidth(1, 80);

      // Company Name - shifted to B column
      sheet.getRange('B1:F1').merge()
        .setValue(companyNameEN)
        .setFontSize(22).setFontWeight('bold').setFontColor('#1565c0')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
      sheet.setRowHeight(1, 40);

      // Arabic Company Name
      sheet.getRange('B2:F2').merge()
        .setValue(companyNameAR)
        .setFontSize(16).setFontWeight('bold').setFontColor('#424242')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
      sheet.setRowHeight(2, 30);

      // Address
      sheet.getRange('B3:F3').merge()
        .setValue('📍 ' + companyAddress)
        .setFontSize(10).setFontColor('#616161')
        .setHorizontalAlignment('center');

      // Phone & Email
      sheet.getRange('B4:F4').merge()
        .setValue('📞 ' + companyPhone + '  |  ✉️ ' + companyEmail)
        .setFontSize(10).setFontColor('#616161')
        .setHorizontalAlignment('center');

    } catch (e) {
      // If logo fails, use text-only header
      insertTextOnlyHeader(sheet, companyNameEN, companyNameAR, companyAddress, companyPhone, companyEmail);
    }
  } else {
    // No logo - use text-only header
    insertTextOnlyHeader(sheet, companyNameEN, companyNameAR, companyAddress, companyPhone, companyEmail);
  }

  // Row 5: Decorative line
  sheet.getRange('A5:F5').merge()
    .setBackground('#1565c0');
  sheet.setRowHeight(5, 4);

  // Row 6: Empty spacer
  sheet.setRowHeight(6, 15);

  // ═══════════════════════════════════════════════════════════════════════════
  // STATEMENT TITLE - عنوان الكشف
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A7:F7').merge()
    .setValue('كشف حساب العميل  |  STATEMENT OF ACCOUNT')
    .setFontSize(14).setFontWeight('bold').setBackground('#e3f2fd').setFontColor('#1565c0')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(7, 35);

  // Row 8: Empty spacer
  sheet.setRowHeight(8, 10);

  // ═══════════════════════════════════════════════════════════════════════════
  // CLIENT INFO SECTION - معلومات العميل
  // ═══════════════════════════════════════════════════════════════════════════

  // Row 9-11: Client info - English labels only, in columns B-C
  sheet.getRange('B9').setValue('Client Name:').setFontWeight('bold').setFontColor('#424242').setHorizontalAlignment('right');
  sheet.getRange('C9:F9').merge().setValue(clientName).setFontColor('#1565c0').setFontWeight('bold');

  sheet.getRange('B10').setValue('Client Code:').setFontWeight('bold').setFontColor('#424242').setHorizontalAlignment('right');
  sheet.getRange('C10:F10').merge().setValue(clientCode).setFontColor('#1565c0');

  sheet.getRange('B11').setValue('Issue Date:').setFontWeight('bold').setFontColor('#424242').setHorizontalAlignment('right');
  sheet.getRange('C11:F11').merge().setValue(formatDate(new Date(), 'yyyy-MM-dd')).setFontColor('#1565c0');

  // Row 12: Decorative line
  sheet.getRange('A12:F12').merge().setBackground('#e0e0e0');
  sheet.setRowHeight(12, 2);

  // Row 13: Empty spacer
  sheet.setRowHeight(13, 10);

  // ═══════════════════════════════════════════════════════════════════════════
  // TABLE SECTION - جدول الحركات
  // ═══════════════════════════════════════════════════════════════════════════

  // Row 14: Table headers
  const headers = ['#', 'التاريخ\nDate', 'الوصف\nDescription', 'له (دائن)\nCredit', 'عليه (مدين)\nDebit', 'الرصيد\nBalance'];
  sheet.getRange(14, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1565c0')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(14, 40);

  // Data with running balance and row numbers
  let runningBalance = 0;
  const data = transactions.map((t, index) => {
    runningBalance += t.credit - t.debit;
    return [
      index + 1,
      formatDate(t.date, 'yyyy-MM-dd'),
      t.description,
      t.credit || '',
      t.debit || '',
      runningBalance
    ];
  });

  const dataStartRow = 15;
  if (data.length > 0) {
    sheet.getRange(dataStartRow, 1, data.length, headers.length).setValues(data);

    // Format numbers
    sheet.getRange(dataStartRow, 4, data.length, 3).setNumberFormat('#,##0.00');

    // Center align row numbers and dates
    sheet.getRange(dataStartRow, 1, data.length, 2).setHorizontalAlignment('center');

    // Alternate row colors
    for (let i = 0; i < data.length; i++) {
      const rowRange = sheet.getRange(dataStartRow + i, 1, 1, headers.length);
      if (i % 2 === 0) {
        rowRange.setBackground('#ffffff');
      } else {
        rowRange.setBackground('#f5f5f5');
      }
    }

    // Add thin borders to data
    sheet.getRange(dataStartRow, 1, data.length, headers.length)
      .setBorder(true, true, true, true, true, true, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // TOTALS SECTION - الإجماليات
  // ═══════════════════════════════════════════════════════════════════════════
  const totalRow = dataStartRow + data.length;

  // Empty row before totals
  sheet.setRowHeight(totalRow, 5);

  // Totals row
  const totalsRow = totalRow + 1;
  sheet.getRange(totalsRow, 1, 1, 6)
    .setValues([['', '', 'الإجمالي | Total', totals.totalCredit, totals.totalDebit, totals.balance]])
    .setFontWeight('bold')
    .setBackground('#e3f2fd')
    .setFontColor('#1565c0');
  sheet.getRange(totalsRow, 3).setHorizontalAlignment('right');
  sheet.getRange(totalsRow, 4, 1, 3).setNumberFormat('#,##0.00');
  sheet.getRange(totalsRow, 1, 1, 6)
    .setBorder(true, true, true, true, null, null, '#1565c0', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // ═══════════════════════════════════════════════════════════════════════════
  // BALANCE SUMMARY BOX - ملخص الرصيد
  // ═══════════════════════════════════════════════════════════════════════════
  const summaryRow = totalsRow + 2;

  // Balance color based on status
  let balanceColor;
  if (totals.balance > 0) {
    balanceColor = '#c62828'; // Red - Amount due from client
  } else if (totals.balance < 0) {
    balanceColor = '#2e7d32'; // Green - Credit balance for client
  } else {
    balanceColor = '#1565c0'; // Blue - Settled
  }

  // Balance label
  sheet.getRange(summaryRow, 1, 1, 3).merge()
    .setValue('Balance / الرصيد')
    .setFontWeight('bold').setFontSize(12)
    .setBackground('#fafafa')
    .setFontColor('#424242')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(summaryRow, 40);

  // Balance amount
  sheet.getRange(summaryRow, 4, 1, 3).merge()
    .setValue(totals.balance)
    .setFontWeight('bold').setFontSize(16)
    .setBackground(balanceColor)
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setNumberFormat('#,##0.00 "TRY"');

  // Border around summary box
  sheet.getRange(summaryRow, 1, 1, 6)
    .setBorder(true, true, true, true, null, null, balanceColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // ═══════════════════════════════════════════════════════════════════════════
  // FOOTER - تذييل
  // ═══════════════════════════════════════════════════════════════════════════
  const footerRow = summaryRow + 2;

  sheet.getRange(footerRow, 1, 1, 6).merge()
    .setValue('شكراً لتعاملكم معنا  |  Thank you for your business')
    .setFontSize(10).setFontStyle('italic').setFontColor('#757575')
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // COLUMN WIDTHS - عرض الأعمدة
  // ═══════════════════════════════════════════════════════════════════════════
  const widths = [40, 100, 220, 110, 110, 120];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Freeze header rows
  sheet.setFrozenRows(14);

  // Set print settings for A4
  sheet.getRange('A1:F' + footerRow).setFontFamily('Arial');

  ss.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert('✅ تم تصدير كشف الحساب بنجاح!\n\nStatement exported to sheet: ' + sheetName);
}

// ==================== 4. CLIENT PROFITABILITY ====================
/**
 * ✅ محدّث: إدخال كود العميل مباشرة بدون قائمة
 */
function showClientProfitability() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // طلب كود العميل مباشرة
  const response = ui.prompt(
    '💹 Client Profitability (ربحية العميل)',
    'أدخل كود العميل (Client Code):\n\nمثال: CLT-001',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const clientCode = response.getResponseText().trim();
  if (!clientCode) {
    ui.alert('⚠️ يرجى إدخال كود العميل!');
    return;
  }

  // البحث عن العميل بالكود
  const client = getClientData(clientCode);
  if (!client) {
    ui.alert('⚠️ لم يتم العثور على عميل بهذا الكود!\n\nClient Code: ' + clientCode);
    return;
  }

  generateClientProfitability(client.code, client.nameEN);
}

/**
 * ✅ محدّث: تقرير الربحية - يحسب فقط الاستحقاقات (بدون تكرار)
 *
 * المنطق الصحيح:
 * - الإيرادات = Revenue Accrual فقط (قيمة الخدمة/المنتج)
 * - المصروفات = Expense Accrual فقط (التكلفة الفعلية)
 * - التحصيلات والمدفوعات لا تدخل في حساب الربحية
 */
function generateClientProfitability(clientCode, clientName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const transSheet = ss.getSheetByName('Transactions');
  if (!transSheet || transSheet.getLastRow() < 2) {
    ui.alert('⚠️ No transactions found!');
    return;
  }

  const transData = transSheet.getDataRange().getValues();

  // جمع فقط الاستحقاقات (وليس التحصيلات/المدفوعات)
  const revenueItems = [];
  const expenseItems = [];
  let totalRevenue = 0;
  let totalDirectExpenses = 0;

  for (let i = 1; i < transData.length; i++) {
    const code = transData[i][4];
    const name = transData[i][5];
    const movementType = transData[i][2] || '';
    const item = transData[i][6] || '';
    const description = transData[i][7] || '';
    const amount = parseFloat(transData[i][13]) || 0; // Amount TRY
    const date = transData[i][1];

    if (code === clientCode || name === clientName) {
      // ✅ فقط استحقاق الإيراد (Revenue Accrual) - وليس التحصيل
      if (movementType.includes('Revenue Accrual') || movementType.includes('استحقاق إيراد')) {
        totalRevenue += amount;
        revenueItems.push({
          date: date,
          item: item || description || 'إيراد',
          description: description,
          amount: amount
        });
      }

      // ✅ فقط استحقاق المصروف (Expense Accrual) - وليس الدفع
      if (movementType.includes('Expense Accrual') || movementType.includes('استحقاق مصروف')) {
        totalDirectExpenses += amount;
        expenseItems.push({
          date: date,
          item: item || description || 'مصروف',
          description: description,
          amount: amount
        });
      }
    }
  }

  const grossProfit = totalRevenue - totalDirectExpenses;
  const profitMargin = totalRevenue > 0 ? (grossProfit / totalRevenue * 100).toFixed(1) : 0;
  const transCount = revenueItems.length + expenseItems.length;

  if (transCount === 0) {
    ui.alert('ℹ️ لا توجد استحقاقات لهذا العميل!\n\nClient Code: ' + clientCode + '\n\n💡 تأكد من وجود حركات من نوع:\n• Revenue Accrual (استحقاق إيراد)\n• Expense Accrual (استحقاق مصروف)');
    return;
  }

  // عرض ملخص وسؤال عن التصدير
  const summary =
    '💹 تقرير ربحية العميل: ' + clientName + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '📊 ملخص الربحية\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'إجمالي الإيرادات: ' + formatCurrency(totalRevenue, 'TRY') + '\n' +
    'إجمالي المصروفات: ' + formatCurrency(totalDirectExpenses, 'TRY') + '\n' +
    '───────────────────────────\n' +
    'صافي الربح: ' + formatCurrency(grossProfit, 'TRY') + '\n' +
    'هامش الربح: ' + profitMargin + '%\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'هل تريد تصدير التقرير إلى شيت؟';

  const exportConfirm = ui.alert(summary, ui.ButtonSet.YES_NO);

  if (exportConfirm === ui.Button.YES) {
    exportClientProfitability(clientCode, clientName, revenueItems, expenseItems, {
      totalRevenue: totalRevenue,
      totalExpenses: totalDirectExpenses,
      grossProfit: grossProfit,
      profitMargin: profitMargin
    });
  }
}

/**
 * ✅ محدّث: تصدير تقرير الربحية - بدون تكرار
 */
function exportClientProfitability(clientCode, clientName, revenueItems, expenseItems, totals) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Profit - ' + clientName.substring(0, 20);

  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);

  sheet = ss.insertSheet(sheetName);
  sheet.setTabColor('#9c27b0');

  let currentRow = 1;

  // ═══════════════════════════════════════════════════════════
  // العنوان الرئيسي
  // ═══════════════════════════════════════════════════════════
  sheet.getRange('A1:D1').merge()
    .setValue('💹 تقرير ربحية العميل: ' + clientName)
    .setFontSize(14).setFontWeight('bold').setBackground('#9c27b0').setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  sheet.getRange('A2').setValue('تاريخ الإصدار: ' + formatDate(new Date(), 'yyyy-MM-dd HH:mm'));
  sheet.getRange('A3').setValue('كود العميل: ' + clientCode);

  // ═══════════════════════════════════════════════════════════
  // ملخص الربحية
  // ═══════════════════════════════════════════════════════════
  currentRow = 5;
  sheet.getRange(currentRow, 1, 1, 3).merge()
    .setValue('📊 ملخص الربحية')
    .setFontWeight('bold').setBackground('#e1bee7').setFontSize(12);

  currentRow++;
  const summaryData = [
    ['إجمالي الإيرادات', totals.totalRevenue, 'TRY'],
    ['إجمالي المصروفات', totals.totalExpenses, 'TRY'],
    ['صافي الربح', totals.grossProfit, 'TRY'],
    ['هامش الربح', totals.profitMargin + '%', '']
  ];
  sheet.getRange(currentRow, 1, summaryData.length, 3).setValues(summaryData);
  sheet.getRange(currentRow, 2, 3, 1).setNumberFormat('#,##0.00');

  // تلوين صافي الربح
  const profitCell = sheet.getRange(currentRow + 2, 2);
  if (totals.grossProfit >= 0) {
    profitCell.setBackground('#c8e6c9').setFontWeight('bold');
  } else {
    profitCell.setBackground('#ffcdd2').setFontWeight('bold');
  }

  // ═══════════════════════════════════════════════════════════
  // تفاصيل الإيرادات
  // ═══════════════════════════════════════════════════════════
  currentRow += summaryData.length + 2;
  sheet.getRange(currentRow, 1, 1, 4).merge()
    .setValue('📈 تفاصيل الإيرادات (' + revenueItems.length + ' بند)')
    .setFontWeight('bold').setBackground('#c8e6c9').setFontSize(11);

  currentRow++;
  const revenueHeaders = ['التاريخ', 'البند', 'الوصف', 'المبلغ (TRY)'];
  sheet.getRange(currentRow, 1, 1, revenueHeaders.length)
    .setValues([revenueHeaders])
    .setFontWeight('bold').setBackground('#e8f5e9');

  currentRow++;
  if (revenueItems.length > 0) {
    const revenueData = revenueItems.map(r => [
      formatDate(r.date, 'yyyy-MM-dd'),
      r.item,
      r.description,
      r.amount
    ]);
    sheet.getRange(currentRow, 1, revenueData.length, 4).setValues(revenueData);
    sheet.getRange(currentRow, 4, revenueData.length, 1).setNumberFormat('#,##0.00');

    // إجمالي الإيرادات
    currentRow += revenueData.length;
    sheet.getRange(currentRow, 1, 1, 4)
      .setValues([['', '', 'الإجمالي', totals.totalRevenue]])
      .setFontWeight('bold').setBackground('#a5d6a7');
    sheet.getRange(currentRow, 4).setNumberFormat('#,##0.00');
    currentRow++;
  }

  // ═══════════════════════════════════════════════════════════
  // تفاصيل المصروفات
  // ═══════════════════════════════════════════════════════════
  currentRow += 1;
  sheet.getRange(currentRow, 1, 1, 4).merge()
    .setValue('📉 تفاصيل المصروفات (' + expenseItems.length + ' بند)')
    .setFontWeight('bold').setBackground('#ffcdd2').setFontSize(11);

  currentRow++;
  const expenseHeaders = ['التاريخ', 'البند', 'الوصف', 'المبلغ (TRY)'];
  sheet.getRange(currentRow, 1, 1, expenseHeaders.length)
    .setValues([expenseHeaders])
    .setFontWeight('bold').setBackground('#ffebee');

  currentRow++;
  if (expenseItems.length > 0) {
    const expenseData = expenseItems.map(e => [
      formatDate(e.date, 'yyyy-MM-dd'),
      e.item,
      e.description,
      e.amount
    ]);
    sheet.getRange(currentRow, 1, expenseData.length, 4).setValues(expenseData);
    sheet.getRange(currentRow, 4, expenseData.length, 1).setNumberFormat('#,##0.00');

    // إجمالي المصروفات
    currentRow += expenseData.length;
    sheet.getRange(currentRow, 1, 1, 4)
      .setValues([['', '', 'الإجمالي', totals.totalExpenses]])
      .setFontWeight('bold').setBackground('#ef9a9a');
    sheet.getRange(currentRow, 4).setNumberFormat('#,##0.00');
  }

  // ═══════════════════════════════════════════════════════════
  // تنسيق الأعمدة
  // ═══════════════════════════════════════════════════════════
  const widths = [100, 200, 200, 120];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  sheet.setFrozenRows(5);
  ss.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert('✅ تم تصدير التقرير إلى شيت: ' + sheetName);
}

// ==================== 5. CLIENTS REPORT ====================
function generateClientsReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const clients = getActiveClients();
  if (clients.length === 0) {
    ui.alert('⚠️ No active clients found!');
    return;
  }
  
  // Calculate summary
  const totalClients = clients.length;
  const totalMonthlyFees = clients.reduce((sum, c) => sum + (c.monthlyFee || 0), 0);
  const avgFee = totalMonthlyFees / totalClients;
  
  // Group by currency
  const byCurrency = {};
  clients.forEach(c => {
    const curr = c.feeCurrency || 'TRY';
    if (!byCurrency[curr]) byCurrency[curr] = { count: 0, total: 0 };
    byCurrency[curr].count++;
    byCurrency[curr].total += c.monthlyFee || 0;
  });
  
  let currencyBreakdown = '';
  Object.keys(byCurrency).forEach(curr => {
    currencyBreakdown += curr + ': ' + byCurrency[curr].count + ' clients, ' + 
                         formatCurrency(byCurrency[curr].total, curr) + '\n';
  });
  
  const report = 
    '📋 CLIENTS REPORT\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Total Active Clients: ' + totalClients + '\n' +
    'Total Monthly Revenue: ' + formatCurrency(totalMonthlyFees, 'TRY') + '\n' +
    'Average Fee/Client: ' + formatCurrency(avgFee, 'TRY') + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'By Currency:\n' + currencyBreakdown;
  
  ui.alert(report);
}

// ==================== 6. OVERDUE REPORT ====================
function generateOverdueReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const transSheet = ss.getSheetByName('Transactions');
  if (!transSheet || transSheet.getLastRow() < 2) {
    ui.alert('⚠️ No transactions found!');
    return;
  }
  
  const data = transSheet.getDataRange().getValues();
  const today = new Date();
  const overdueList = [];
  
  for (let i = 1; i < data.length; i++) {
    const status = data[i][18];
    const dueDate = data[i][19];
    const clientName = data[i][5];
    const amount = data[i][10];
    const currency = data[i][11];
    const invoiceNo = data[i][17];
    
    if (status && status.includes('Pending') && dueDate) {
      const due = new Date(dueDate);
      const diffDays = Math.floor((today - due) / (1000 * 60 * 60 * 24));
      
      if (diffDays > 0) {
        overdueList.push({
          client: clientName,
          invoice: invoiceNo || 'N/A',
          amount: formatCurrency(amount, currency || 'TRY'),
          days: diffDays
        });
      }
    }
  }
  
  if (overdueList.length === 0) {
    ui.alert('✅ No overdue payments!\n\nAll payments are on time.');
    return;
  }
  
  // Sort by days overdue
  overdueList.sort((a, b) => b.days - a.days);
  
  const list = overdueList.slice(0, 10).map(o => 
    '• ' + o.client + ' | ' + o.invoice + ' | ' + o.amount + ' | ' + o.days + ' days'
  ).join('\n');
  
  ui.alert(
    '⚠️ OVERDUE REPORT\n\n' +
    'Total Overdue: ' + overdueList.length + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    (overdueList.length > 10 ? 'Top 10:\n' : '') +
    list +
    (overdueList.length > 10 ? '\n\n... and ' + (overdueList.length - 10) + ' more' : '')
  );
}

// ==================== 7. REFRESH ALL DATA ====================
function refreshAllData() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert('🔄 Refreshing all data...\n\nPlease wait...');
  
  try {
    // Refresh dropdowns
    refreshCashBankDropdown();
    
    // Refresh dashboard
    refreshDashboard();
    
    ui.alert('✅ All data refreshed!\n\n• Cash/Bank dropdowns updated\n• Dashboard refreshed');
    
  } catch (error) {
    ui.alert('❌ Error refreshing data:\n\n' + error.message);
  }
}

// ==================== 8. HELPER FUNCTIONS ====================

/**
 * Helper function to insert text-only header (when no logo is provided)
 */
function insertTextOnlyHeader(sheet, companyNameEN, companyNameAR, companyAddress, companyPhone, companyEmail) {
  // Row 1: Company Name (Large, Bold)
  sheet.getRange('A1:F1').merge()
    .setValue(companyNameEN)
    .setFontSize(22).setFontWeight('bold').setFontColor('#1565c0')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);

  // Row 2: Arabic Company Name
  sheet.getRange('A2:F2').merge()
    .setValue(companyNameAR)
    .setFontSize(16).setFontWeight('bold').setFontColor('#424242')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(2, 30);

  // Row 3: Address
  sheet.getRange('A3:F3').merge()
    .setValue('📍 ' + companyAddress)
    .setFontSize(10).setFontColor('#616161')
    .setHorizontalAlignment('center');

  // Row 4: Phone & Email
  sheet.getRange('A4:F4').merge()
    .setValue('📞 ' + companyPhone + '  |  ✉️ ' + companyEmail)
    .setFontSize(10).setFontColor('#616161')
    .setHorizontalAlignment('center');
}

// ==================== END OF PART 8 ====================
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
    '• Items Database\n' +
    '• Clients, Vendors, Employees\n' +
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
    
    // Part 3: Party sheets
    createClientsSheet(ss);
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
3. All Monthly - batch for all clients with fees

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
    'Settings', 'Holidays', 'Categories', 'Movement Types', 'Items Database',
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
    // Fix Amount TRY (column N)
    const amount = transSheet.getRange(row, 11).getValue();
    const rate = transSheet.getRange(row, 13).getValue() || 1;
    transSheet.getRange(row, 14).setValue(amount * rate);
    
    // Fix Remaining (column V)
    const paid = transSheet.getRange(row, 21).getValue() || 0;
    transSheet.getRange(row, 22).setValue(amount - paid);
    
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
    const paymentMethod = data[i][14]; // Column O
    const cashBank = data[i][15]; // Column P
    const status = data[i][18]; // Column S
    
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
    const movementType = data[i][2];
    const direction = (movementType && movementType.includes('Revenue')) ? 'IN' : 'OUT';
    
    addCashBankEntry(
      sheetName,
      data[i][1], // Date
      data[i][7] || data[i][6], // Description or Item
      data[i][16], // Reference
      data[i][8] || data[i][5], // Party Name or Client Name
      transCode,
      data[i][10], // Amount
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
// ╔════════════════════════════════════════════════════════════════════════════╗
// ║                    DC CONSULTING ACCOUNTING SYSTEM v3.0                     ║
// ║                              Part 10 of 10                                   ║
// ║                    Petty Cash Advance System (نظام العهد)                    ║
// ╚════════════════════════════════════════════════════════════════════════════╝

// ==================== 1. CREATE ADVANCES SHEET ====================
function createAdvancesSheet(ss) {
  let sheet = ss.getSheetByName('Advances');
  if (sheet) ss.deleteSheet(sheet);

  sheet = ss.insertSheet('Advances');
  sheet.setTabColor('#ff5722');

  const headers = [
    'Advance Code',       // A - كود العهدة
    'Date',               // B - تاريخ الصرف
    'Employee Code',      // C - كود الموظف
    'Employee Name',      // D - اسم الموظف
    'Amount',             // E - المبلغ
    'Currency',           // F - العملة
    'Source Type',        // G - نوع المصدر (Cash/Bank)
    'Source Name',        // H - اسم المصدر
    'Purpose',            // I - الغرض
    'Status',             // J - الحالة
    'Total Expenses',     // K - إجمالي المصروفات
    'Remaining',          // L - المتبقي
    'Settlement Date',    // M - تاريخ التسوية
    'Notes',              // N - ملاحظات
    'Created Date'        // O - تاريخ الإنشاء
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#e65100')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const widths = [100, 100, 100, 150, 100, 70, 80, 150, 200, 90, 100, 100, 100, 200, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  const lastRow = 500;

  // Currency validation
  const currencyValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CURRENCIES, true)
    .build();
  sheet.getRange(2, 6, lastRow, 1).setDataValidation(currencyValidation);

  // Source Type validation
  const sourceTypeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Cash', 'Bank'], true)
    .build();
  sheet.getRange(2, 7, lastRow, 1).setDataValidation(sourceTypeValidation);

  // Status validation
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active (نشطة)', 'Settled (مسواة)', 'Cancelled (ملغاة)'], true)
    .build();
  sheet.getRange(2, 10, lastRow, 1).setDataValidation(statusValidation);

  // Number formats
  sheet.getRange(2, 2, lastRow, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 5, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 11, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 12, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 13, lastRow, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 15, lastRow, 1).setNumberFormat('yyyy-mm-dd HH:mm');

  // Conditional formatting for Status
  const statusRange = sheet.getRange(2, 10, lastRow, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Active').setBackground('#fff9c4').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Settled').setBackground('#c8e6c9').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Cancelled').setBackground('#ffcdd2').setRanges([statusRange]).build()
  ]);

  // Formulas for Total Expenses and Remaining
  sheet.getRange('K2').setNote('Auto-calculated from Advance Expenses');
  sheet.getRange('L2').setNote('= Amount - Total Expenses');

  sheet.setFrozenRows(1);

  return sheet;
}

// ==================== 2. CREATE ADVANCE EXPENSES SHEET ====================
function createAdvanceExpensesSheet(ss) {
  let sheet = ss.getSheetByName('Advance Expenses');
  if (sheet) ss.deleteSheet(sheet);

  sheet = ss.insertSheet('Advance Expenses');
  sheet.setTabColor('#ff7043');

  const headers = [
    'Expense Code',       // A - كود المصروف
    'Advance Code',       // B - كود العهدة
    'Date',               // C - التاريخ
    'Description',        // D - الوصف
    'Amount',             // E - المبلغ
    'Currency',           // F - العملة
    'Receipt No',         // G - رقم الإيصال
    'Vendor',             // H - المورد
    'Category',           // I - التصنيف
    'Notes',              // J - ملاحظات
    'Created Date'        // K - تاريخ الإنشاء
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#bf360c')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const widths = [100, 100, 100, 250, 100, 70, 100, 150, 120, 200, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  const lastRow = 1000;

  // Currency validation
  const currencyValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CURRENCIES, true)
    .build();
  sheet.getRange(2, 6, lastRow, 1).setDataValidation(currencyValidation);

  // Category validation
  const categoryValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      'Office Supplies (مستلزمات مكتبية)',
      'Transportation (مواصلات)',
      'Meals (وجبات)',
      'Utilities (مرافق)',
      'Printing (طباعة)',
      'Maintenance (صيانة)',
      'Communication (اتصالات)',
      'Other (أخرى)'
    ], true)
    .build();
  sheet.getRange(2, 9, lastRow, 1).setDataValidation(categoryValidation);

  // Number formats
  sheet.getRange(2, 3, lastRow, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 5, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 11, lastRow, 1).setNumberFormat('yyyy-mm-dd HH:mm');

  sheet.setFrozenRows(1);

  return sheet;
}

// ==================== 3. ISSUE ADVANCE (صرف عهدة) ====================
function issueAdvance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Check if Advances sheet exists
  let advSheet = ss.getSheetByName('Advances');
  if (!advSheet) {
    advSheet = createAdvancesSheet(ss);
  }

  // Step 1: Select Employee
  const employees = getActiveEmployees();
  if (employees.length === 0) {
    ui.alert('❌ No active employees found!\n\nلا يوجد موظفين نشطين');
    return;
  }

  const empList = employees.map((e, i) => (i + 1) + '. ' + e.nameEN + ' (' + e.code + ')').join('\n');
  const empResponse = ui.prompt(
    '💼 Issue Advance (1/5) - Employee\n\nصرف عهدة - اختر الموظف',
    'Select employee number:\n\n' + empList,
    ui.ButtonSet.OK_CANCEL
  );
  if (empResponse.getSelectedButton() !== ui.Button.OK) return;

  const empIndex = parseInt(empResponse.getResponseText()) - 1;
  if (isNaN(empIndex) || empIndex < 0 || empIndex >= employees.length) {
    ui.alert('⚠️ Invalid selection!');
    return;
  }
  const selectedEmployee = employees[empIndex];

  // Step 2: Enter Amount
  const amountResponse = ui.prompt(
    '💼 Issue Advance (2/5) - Amount\n\nEmployee: ' + selectedEmployee.nameEN,
    'Enter advance amount (أدخل مبلغ العهدة):',
    ui.ButtonSet.OK_CANCEL
  );
  if (amountResponse.getSelectedButton() !== ui.Button.OK) return;

  const amount = parseFloat(amountResponse.getResponseText().replace(/,/g, ''));
  if (isNaN(amount) || amount <= 0) {
    ui.alert('⚠️ Invalid amount!');
    return;
  }

  // Step 3: Select Currency
  const currencyResponse = ui.prompt(
    '💼 Issue Advance (3/5) - Currency\n\nAmount: ' + amount.toLocaleString(),
    'Enter currency (العملة):\n\nOptions: TRY, USD, EUR, SAR, EGP, AED, GBP\n\nDefault: TRY',
    ui.ButtonSet.OK_CANCEL
  );
  if (currencyResponse.getSelectedButton() !== ui.Button.OK) return;

  let currency = currencyResponse.getResponseText().trim().toUpperCase() || 'TRY';
  if (!CURRENCIES.includes(currency)) currency = 'TRY';

  // Step 4: Select Source (Cash or Bank)
  const sourceResponse = ui.alert(
    '💼 Issue Advance (4/5) - Source\n\nمصدر الصرف',
    'Select payment source:\n\nYES = Cash Box (خزينة)\nNO = Bank Account (حساب بنكي)',
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (sourceResponse === ui.Button.CANCEL) return;

  const sourceType = (sourceResponse === ui.Button.YES) ? 'Cash' : 'Bank';
  let sourceName = '';
  let sourceList = [];

  if (sourceType === 'Cash') {
    sourceList = getCashBoxesList();
    if (sourceList.length === 0) {
      ui.alert('❌ No cash boxes found!');
      return;
    }
  } else {
    sourceList = getBankAccountsList();
    if (sourceList.length === 0) {
      ui.alert('❌ No bank accounts found!');
      return;
    }
  }

  const sourceListStr = sourceList.map((s, i) => (i + 1) + '. ' + s.name + ' (' + s.currency + ')').join('\n');
  const sourceNameResponse = ui.prompt(
    '💼 Issue Advance (4b/5) - Select ' + sourceType,
    'Select ' + sourceType.toLowerCase() + ' number:\n\n' + sourceListStr,
    ui.ButtonSet.OK_CANCEL
  );
  if (sourceNameResponse.getSelectedButton() !== ui.Button.OK) return;

  const sourceIndex = parseInt(sourceNameResponse.getResponseText()) - 1;
  if (isNaN(sourceIndex) || sourceIndex < 0 || sourceIndex >= sourceList.length) {
    ui.alert('⚠️ Invalid selection!');
    return;
  }
  sourceName = sourceList[sourceIndex].name;

  // Step 5: Enter Purpose
  const purposeResponse = ui.prompt(
    '💼 Issue Advance (5/5) - Purpose\n\nEmployee: ' + selectedEmployee.nameEN + '\nAmount: ' + formatCurrency(amount, currency),
    'Enter purpose (الغرض من العهدة):',
    ui.ButtonSet.OK_CANCEL
  );
  if (purposeResponse.getSelectedButton() !== ui.Button.OK) return;

  const purpose = purposeResponse.getResponseText().trim() || 'Petty Cash Advance';

  // Final Confirmation
  const confirm = ui.alert(
    '💼 Confirm Advance (تأكيد صرف العهدة)\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Employee: ' + selectedEmployee.nameEN + '\n' +
    'Amount: ' + formatCurrency(amount, currency) + '\n' +
    'Source: ' + sourceType + ' - ' + sourceName + '\n' +
    'Purpose: ' + purpose + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'Proceed with advance?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // Generate Advance Code
  const advanceCode = generateNextCode('ADV', advSheet, 1);
  const advanceDate = new Date();

  // Record in Advances sheet
  const lastRow = advSheet.getLastRow() + 1;
  advSheet.getRange(lastRow, 1).setValue(advanceCode);
  advSheet.getRange(lastRow, 2).setValue(advanceDate);
  advSheet.getRange(lastRow, 3).setValue(selectedEmployee.code);
  advSheet.getRange(lastRow, 4).setValue(selectedEmployee.nameEN);
  advSheet.getRange(lastRow, 5).setValue(amount);
  advSheet.getRange(lastRow, 6).setValue(currency);
  advSheet.getRange(lastRow, 7).setValue(sourceType);
  advSheet.getRange(lastRow, 8).setValue(sourceName);
  advSheet.getRange(lastRow, 9).setValue(purpose);
  advSheet.getRange(lastRow, 10).setValue('Active (نشطة)');
  advSheet.getRange(lastRow, 11).setValue(0); // Total Expenses
  advSheet.getRange(lastRow, 12).setValue(amount); // Remaining
  advSheet.getRange(lastRow, 15).setValue(new Date());

  // Record withdrawal in Transactions
  const transSheet = ss.getSheetByName('Transactions');
  if (transSheet) {
    const transRow = transSheet.getLastRow() + 1;
    transSheet.getRange(transRow, 1).setValue(transRow - 1); // Transaction Code
    transSheet.getRange(transRow, 2).setValue(advanceDate); // Date
    transSheet.getRange(transRow, 3).setValue('Advance Issue (صرف عهدة)'); // Movement Type
    transSheet.getRange(transRow, 4).setValue('Petty Cash Advance (عهدة مؤقتة)'); // Sub Type
    transSheet.getRange(transRow, 5).setValue(selectedEmployee.code); // Party Code
    transSheet.getRange(transRow, 6).setValue(selectedEmployee.nameEN); // Party Name
    transSheet.getRange(transRow, 8).setValue(purpose); // Description
    transSheet.getRange(transRow, 10).setValue('Employee (موظف)'); // Party Type
    transSheet.getRange(transRow, 11).setValue(amount); // Amount
    transSheet.getRange(transRow, 12).setValue(currency); // Currency
    transSheet.getRange(transRow, 15).setValue(sourceType === 'Cash' ? 'Cash (نقدي)' : 'Bank Transfer (تحويل بنكي)'); // Payment Method
    transSheet.getRange(transRow, 16).setValue(sourceName); // Cash/Bank
    transSheet.getRange(transRow, 19).setValue('Paid (مدفوع)'); // Payment Status
    transSheet.getRange(transRow, 22).setValue(advanceCode); // Reference
    transSheet.getRange(transRow, 25).setValue('Yes (نعم)'); // Confirmed

    // Apply color
    if (typeof applyPaymentMethodColor === 'function') {
      applyPaymentMethodColor(transSheet, transRow, sourceType === 'Cash' ? 'Cash (نقدي)' : 'Bank Transfer (تحويل بنكي)');
    }
  }

  // Update Cash/Bank balance (withdrawal - OUT)
  addCashBankEntry(
    sourceName,                                    // Sheet name
    advanceDate,                                   // Date
    'Advance: ' + advanceCode + ' - ' + selectedEmployee.nameEN,  // Description
    advanceCode,                                   // Reference
    selectedEmployee.nameEN,                       // Party
    '',                                            // Trans Code
    amount,                                        // Amount
    'OUT'                                          // Direction (withdrawal)
  );

  ss.setActiveSheet(advSheet);

  ui.alert(
    '✅ Advance Issued Successfully!\n\n' +
    'تم صرف العهدة بنجاح\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Advance Code: ' + advanceCode + '\n' +
    'Employee: ' + selectedEmployee.nameEN + '\n' +
    'Amount: ' + formatCurrency(amount, currency) + '\n' +
    'Source: ' + sourceName + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'
  );
}

// ==================== 4. ADD ADVANCE EXPENSE (إضافة مصروف للعهدة) ====================
function addAdvanceExpense() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Check sheets
  const advSheet = ss.getSheetByName('Advances');
  if (!advSheet) {
    ui.alert('❌ Advances sheet not found!\n\nRun Setup System first.');
    return;
  }

  let expSheet = ss.getSheetByName('Advance Expenses');
  if (!expSheet) {
    expSheet = createAdvanceExpensesSheet(ss);
  }

  // Get active advances
  const advances = getActiveAdvances();
  if (advances.length === 0) {
    ui.alert('❌ No active advances found!\n\nلا توجد عهد نشطة');
    return;
  }

  // Step 1: Select Advance
  const advList = advances.map((a, i) =>
    (i + 1) + '. ' + a.code + ' - ' + a.employeeName + ' (' + formatCurrency(a.remaining, a.currency) + ' remaining)'
  ).join('\n');

  const advResponse = ui.prompt(
    '📝 Add Expense (1/4) - Select Advance\n\nاختر العهدة',
    'Select advance number:\n\n' + advList,
    ui.ButtonSet.OK_CANCEL
  );
  if (advResponse.getSelectedButton() !== ui.Button.OK) return;

  const advIndex = parseInt(advResponse.getResponseText()) - 1;
  if (isNaN(advIndex) || advIndex < 0 || advIndex >= advances.length) {
    ui.alert('⚠️ Invalid selection!');
    return;
  }
  const selectedAdvance = advances[advIndex];

  // Step 2: Enter Description
  const descResponse = ui.prompt(
    '📝 Add Expense (2/4) - Description\n\nAdvance: ' + selectedAdvance.code + '\nRemaining: ' + formatCurrency(selectedAdvance.remaining, selectedAdvance.currency),
    'Enter expense description (وصف المصروف):',
    ui.ButtonSet.OK_CANCEL
  );
  if (descResponse.getSelectedButton() !== ui.Button.OK) return;

  const description = descResponse.getResponseText().trim();
  if (!description) {
    ui.alert('⚠️ Description is required!');
    return;
  }

  // Step 3: Enter Amount
  const amountResponse = ui.prompt(
    '📝 Add Expense (3/4) - Amount\n\nRemaining: ' + formatCurrency(selectedAdvance.remaining, selectedAdvance.currency),
    'Enter expense amount (المبلغ):\n\nMax: ' + selectedAdvance.remaining,
    ui.ButtonSet.OK_CANCEL
  );
  if (amountResponse.getSelectedButton() !== ui.Button.OK) return;

  const expAmount = parseFloat(amountResponse.getResponseText().replace(/,/g, ''));
  if (isNaN(expAmount) || expAmount <= 0) {
    ui.alert('⚠️ Invalid amount!');
    return;
  }

  if (expAmount > selectedAdvance.remaining) {
    ui.alert('⚠️ Amount exceeds remaining balance!\n\nالمبلغ يتجاوز المتبقي من العهدة\n\nRemaining: ' + formatCurrency(selectedAdvance.remaining, selectedAdvance.currency));
    return;
  }

  // Step 4: Select Category
  const categories = [
    '1. Office Supplies (مستلزمات مكتبية)',
    '2. Transportation (مواصلات)',
    '3. Meals (وجبات)',
    '4. Utilities (مرافق)',
    '5. Printing (طباعة)',
    '6. Maintenance (صيانة)',
    '7. Communication (اتصالات)',
    '8. Other (أخرى)'
  ];

  const catResponse = ui.prompt(
    '📝 Add Expense (4/4) - Category\n\nالتصنيف',
    'Select category number:\n\n' + categories.join('\n'),
    ui.ButtonSet.OK_CANCEL
  );
  if (catResponse.getSelectedButton() !== ui.Button.OK) return;

  const catIndex = parseInt(catResponse.getResponseText()) - 1;
  const categoryValues = [
    'Office Supplies (مستلزمات مكتبية)',
    'Transportation (مواصلات)',
    'Meals (وجبات)',
    'Utilities (مرافق)',
    'Printing (طباعة)',
    'Maintenance (صيانة)',
    'Communication (اتصالات)',
    'Other (أخرى)'
  ];
  const category = (catIndex >= 0 && catIndex < categoryValues.length) ? categoryValues[catIndex] : 'Other (أخرى)';

  // Record Expense
  const expenseCode = generateNextCode('EXP', expSheet, 1);
  const expenseDate = new Date();

  const lastRow = expSheet.getLastRow() + 1;
  expSheet.getRange(lastRow, 1).setValue(expenseCode);
  expSheet.getRange(lastRow, 2).setValue(selectedAdvance.code);
  expSheet.getRange(lastRow, 3).setValue(expenseDate);
  expSheet.getRange(lastRow, 4).setValue(description);
  expSheet.getRange(lastRow, 5).setValue(expAmount);
  expSheet.getRange(lastRow, 6).setValue(selectedAdvance.currency);
  expSheet.getRange(lastRow, 9).setValue(category);
  expSheet.getRange(lastRow, 11).setValue(new Date());

  // Update Advance totals
  updateAdvanceTotals(selectedAdvance.code);

  ui.alert(
    '✅ Expense Added!\n\n' +
    'تم إضافة المصروف\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Expense Code: ' + expenseCode + '\n' +
    'Advance: ' + selectedAdvance.code + '\n' +
    'Amount: ' + formatCurrency(expAmount, selectedAdvance.currency) + '\n' +
    'Category: ' + category + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'New Remaining: ' + formatCurrency(selectedAdvance.remaining - expAmount, selectedAdvance.currency)
  );
}

// ==================== 5. SETTLE ADVANCE (تسوية العهدة) ====================
function settleAdvance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const advSheet = ss.getSheetByName('Advances');
  if (!advSheet) {
    ui.alert('❌ Advances sheet not found!');
    return;
  }

  // Get active advances
  const advances = getActiveAdvances();
  if (advances.length === 0) {
    ui.alert('❌ No active advances to settle!\n\nلا توجد عهد للتسوية');
    return;
  }

  // Select Advance
  const advList = advances.map((a, i) =>
    (i + 1) + '. ' + a.code + ' - ' + a.employeeName + '\n   Amount: ' + formatCurrency(a.amount, a.currency) +
    ' | Expenses: ' + formatCurrency(a.totalExpenses, a.currency) +
    ' | Remaining: ' + formatCurrency(a.remaining, a.currency)
  ).join('\n\n');

  const advResponse = ui.prompt(
    '✅ Settle Advance - Select\n\nتسوية العهدة',
    'Select advance number:\n\n' + advList,
    ui.ButtonSet.OK_CANCEL
  );
  if (advResponse.getSelectedButton() !== ui.Button.OK) return;

  const advIndex = parseInt(advResponse.getResponseText()) - 1;
  if (isNaN(advIndex) || advIndex < 0 || advIndex >= advances.length) {
    ui.alert('⚠️ Invalid selection!');
    return;
  }
  const selectedAdvance = advances[advIndex];

  // Update totals first
  updateAdvanceTotals(selectedAdvance.code);

  // Get updated data
  const updatedAdvance = getAdvanceData(selectedAdvance.code);
  if (!updatedAdvance) {
    ui.alert('❌ Error loading advance data!');
    return;
  }

  // Get expenses
  const expenses = getAdvanceExpenses(selectedAdvance.code);
  const expensesList = expenses.map(e => '• ' + e.description + ': ' + formatCurrency(e.amount, e.currency)).join('\n');

  // Confirm settlement
  const confirm = ui.alert(
    '✅ Settle Advance (تسوية العهدة)\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Advance: ' + updatedAdvance.code + '\n' +
    'Employee: ' + updatedAdvance.employeeName + '\n' +
    'Original Amount: ' + formatCurrency(updatedAdvance.amount, updatedAdvance.currency) + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Expenses (' + expenses.length + '):\n' + (expensesList || 'No expenses recorded') + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Total Expenses: ' + formatCurrency(updatedAdvance.totalExpenses, updatedAdvance.currency) + '\n' +
    'Remaining to Return: ' + formatCurrency(updatedAdvance.remaining, updatedAdvance.currency) + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'Proceed with settlement?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const settlementDate = new Date();

  // Transfer expenses to Transactions
  const transSheet = ss.getSheetByName('Transactions');
  if (transSheet && expenses.length > 0) {
    expenses.forEach(exp => {
      const transRow = transSheet.getLastRow() + 1;
      transSheet.getRange(transRow, 1).setValue(transRow - 1);
      transSheet.getRange(transRow, 2).setValue(exp.date);
      transSheet.getRange(transRow, 3).setValue('Expense (مصروف)');
      transSheet.getRange(transRow, 4).setValue(exp.category || 'Other Expense (مصروف آخر)');
      transSheet.getRange(transRow, 5).setValue(updatedAdvance.employeeCode);
      transSheet.getRange(transRow, 6).setValue(updatedAdvance.employeeName);
      transSheet.getRange(transRow, 8).setValue(exp.description + ' (Advance: ' + updatedAdvance.code + ')');
      transSheet.getRange(transRow, 10).setValue('Employee (موظف)');
      transSheet.getRange(transRow, 11).setValue(exp.amount);
      transSheet.getRange(transRow, 12).setValue(exp.currency);
      transSheet.getRange(transRow, 15).setValue('Advance (عهدة)');
      transSheet.getRange(transRow, 19).setValue('Paid (مدفوع)');
      transSheet.getRange(transRow, 22).setValue(updatedAdvance.code);
      transSheet.getRange(transRow, 25).setValue('Yes (نعم)');
    });
  }

  // Return remaining to Cash/Bank
  if (updatedAdvance.remaining > 0) {
    // Record return transaction
    if (transSheet) {
      const transRow = transSheet.getLastRow() + 1;
      transSheet.getRange(transRow, 1).setValue(transRow - 1);
      transSheet.getRange(transRow, 2).setValue(settlementDate);
      transSheet.getRange(transRow, 3).setValue('Advance Return (رد عهدة)');
      transSheet.getRange(transRow, 4).setValue('Advance Settlement (تسوية عهدة)');
      transSheet.getRange(transRow, 5).setValue(updatedAdvance.employeeCode);
      transSheet.getRange(transRow, 6).setValue(updatedAdvance.employeeName);
      transSheet.getRange(transRow, 8).setValue('Return of advance ' + updatedAdvance.code);
      transSheet.getRange(transRow, 10).setValue('Employee (موظف)');
      transSheet.getRange(transRow, 11).setValue(updatedAdvance.remaining);
      transSheet.getRange(transRow, 12).setValue(updatedAdvance.currency);
      transSheet.getRange(transRow, 15).setValue(updatedAdvance.sourceType === 'Cash' ? 'Cash (نقدي)' : 'Bank Transfer (تحويل بنكي)');
      transSheet.getRange(transRow, 16).setValue(updatedAdvance.sourceName);
      transSheet.getRange(transRow, 19).setValue('Received (مستلم)');
      transSheet.getRange(transRow, 22).setValue(updatedAdvance.code);
      transSheet.getRange(transRow, 25).setValue('Yes (نعم)');
    }

    // Update Cash/Bank (deposit - IN)
    addCashBankEntry(
      updatedAdvance.sourceName,                     // Sheet name
      settlementDate,                                // Date
      'Return Advance: ' + updatedAdvance.code,      // Description
      updatedAdvance.code,                           // Reference
      updatedAdvance.employeeName,                   // Party
      '',                                            // Trans Code
      updatedAdvance.remaining,                      // Amount
      'IN'                                           // Direction (deposit/return)
    );
  }

  // Update Advance status
  advSheet.getRange(updatedAdvance.row, 10).setValue('Settled (مسواة)');
  advSheet.getRange(updatedAdvance.row, 13).setValue(settlementDate);

  ui.alert(
    '✅ Advance Settled Successfully!\n\n' +
    'تمت تسوية العهدة بنجاح\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Advance: ' + updatedAdvance.code + '\n' +
    'Total Expenses: ' + formatCurrency(updatedAdvance.totalExpenses, updatedAdvance.currency) + '\n' +
    'Returned: ' + formatCurrency(updatedAdvance.remaining, updatedAdvance.currency) + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'Expenses transferred to Transactions sheet.'
  );
}

// ==================== 6. HELPER FUNCTIONS ====================

// Get Active Advances
function getActiveAdvances() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Advances');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const advances = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][9] && data[i][9].toString().includes('Active')) {
      advances.push({
        row: i + 1,
        code: data[i][0],
        date: data[i][1],
        employeeCode: data[i][2],
        employeeName: data[i][3],
        amount: data[i][4] || 0,
        currency: data[i][5] || 'TRY',
        sourceType: data[i][6],
        sourceName: data[i][7],
        purpose: data[i][8],
        status: data[i][9],
        totalExpenses: data[i][10] || 0,
        remaining: data[i][11] || 0
      });
    }
  }
  return advances;
}

// Get Advance Data by Code
function getAdvanceData(advanceCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Advances');
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === advanceCode) {
      return {
        row: i + 1,
        code: data[i][0],
        date: data[i][1],
        employeeCode: data[i][2],
        employeeName: data[i][3],
        amount: data[i][4] || 0,
        currency: data[i][5] || 'TRY',
        sourceType: data[i][6],
        sourceName: data[i][7],
        purpose: data[i][8],
        status: data[i][9],
        totalExpenses: data[i][10] || 0,
        remaining: data[i][11] || 0
      };
    }
  }
  return null;
}

// Get Advance Expenses
function getAdvanceExpenses(advanceCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Advance Expenses');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const expenses = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === advanceCode) {
      expenses.push({
        code: data[i][0],
        advanceCode: data[i][1],
        date: data[i][2],
        description: data[i][3],
        amount: data[i][4] || 0,
        currency: data[i][5] || 'TRY',
        receiptNo: data[i][6],
        vendor: data[i][7],
        category: data[i][8]
      });
    }
  }
  return expenses;
}

// Update Advance Totals
function updateAdvanceTotals(advanceCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const advSheet = ss.getSheetByName('Advances');
  if (!advSheet) return;

  const advData = advSheet.getDataRange().getValues();
  let advRow = -1;
  let advAmount = 0;
  let advCurrency = 'TRY';

  for (let i = 1; i < advData.length; i++) {
    if (advData[i][0] === advanceCode) {
      advRow = i + 1;
      advAmount = advData[i][4] || 0;
      advCurrency = advData[i][5] || 'TRY';
      break;
    }
  }

  if (advRow === -1) return;

  const expenses = getAdvanceExpenses(advanceCode);
  let totalExpenses = 0;
  expenses.forEach(e => {
    if (e.currency === advCurrency) {
      totalExpenses += e.amount;
    }
  });

  const remaining = advAmount - totalExpenses;

  advSheet.getRange(advRow, 11).setValue(totalExpenses);
  advSheet.getRange(advRow, 12).setValue(remaining);
}

// ==================== 7. VIEW FUNCTIONS ====================

// Show Advances Sheet
function showAdvances() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Advances');
  if (!sheet) {
    sheet = createAdvancesSheet(ss);
  }
  ss.setActiveSheet(sheet);
}

// Show Advance Expenses Sheet
function showAdvanceExpenses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Advance Expenses');
  if (!sheet) {
    sheet = createAdvanceExpensesSheet(ss);
  }
  ss.setActiveSheet(sheet);
}

// Show Advance Statement
function showAdvanceStatement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const advSheet = ss.getSheetByName('Advances');
  if (!advSheet || advSheet.getLastRow() < 2) {
    ui.alert('❌ No advances found!');
    return;
  }

  // Get all advances
  const data = advSheet.getDataRange().getValues();
  const advances = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      advances.push({
        code: data[i][0],
        employeeName: data[i][3],
        amount: data[i][4],
        currency: data[i][5],
        status: data[i][9],
        remaining: data[i][11]
      });
    }
  }

  if (advances.length === 0) {
    ui.alert('❌ No advances found!');
    return;
  }

  // Select advance
  const advList = advances.map((a, i) =>
    (i + 1) + '. ' + a.code + ' - ' + a.employeeName + ' [' + a.status + ']'
  ).join('\n');

  const response = ui.prompt(
    '📊 Advance Statement\n\nكشف حساب العهدة',
    'Select advance number:\n\n' + advList,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const idx = parseInt(response.getResponseText()) - 1;
  if (isNaN(idx) || idx < 0 || idx >= advances.length) {
    ui.alert('⚠️ Invalid selection!');
    return;
  }

  const selectedAdvance = getAdvanceData(advances[idx].code);
  const expenses = getAdvanceExpenses(advances[idx].code);

  let statement =
    '📊 ADVANCE STATEMENT (كشف حساب العهدة)\n' +
    '═══════════════════════════════════════════════\n\n' +
    '📋 ADVANCE DETAILS:\n' +
    '───────────────────────────────────────────────\n' +
    'Code: ' + selectedAdvance.code + '\n' +
    'Employee: ' + selectedAdvance.employeeName + '\n' +
    'Date: ' + formatDate(selectedAdvance.date, 'yyyy-MM-dd') + '\n' +
    'Purpose: ' + selectedAdvance.purpose + '\n' +
    'Amount: ' + formatCurrency(selectedAdvance.amount, selectedAdvance.currency) + '\n' +
    'Source: ' + selectedAdvance.sourceType + ' - ' + selectedAdvance.sourceName + '\n' +
    'Status: ' + selectedAdvance.status + '\n\n';

  statement += '📝 EXPENSES (' + expenses.length + '):\n' +
    '───────────────────────────────────────────────\n';

  if (expenses.length > 0) {
    expenses.forEach((e, i) => {
      statement += (i + 1) + '. ' + formatDate(e.date, 'yyyy-MM-dd') + ' | ' +
        e.description + ' | ' + formatCurrency(e.amount, e.currency) + '\n';
    });
  } else {
    statement += 'No expenses recorded.\n';
  }

  statement += '\n═══════════════════════════════════════════════\n' +
    'Total Expenses: ' + formatCurrency(selectedAdvance.totalExpenses, selectedAdvance.currency) + '\n' +
    'Remaining: ' + formatCurrency(selectedAdvance.remaining, selectedAdvance.currency) + '\n' +
    '═══════════════════════════════════════════════';

  ui.alert(statement);
}

// ==================== END OF PART 10 ====================