// ╔════════════════════════════════════════════════════════════════════════════╗
// ║                    DC CONSULTING ACCOUNTING SYSTEM v3.1                     ║
// ║                              Part 1 of 9                                    ║
// ║                    Core + Menu + Config + Security                          ║
// ╚════════════════════════════════════════════════════════════════════════════╝

// ==================== GLOBAL CONSTANTS ====================
const SYSTEM_VERSION = '3.1';
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
      .addItem('🖼️ Add Logo to Template', 'updateInvoiceLogo')
      .addItem('📊 Invoice Log', 'showInvoiceLog'))
    
    // Clients & Parties
    .addSubMenu(ui.createMenu('👥 Clients & Parties (العملاء والأطراف)')
      .addItem('➕ Add Client (إضافة عميل)', 'addNewClient')
      .addItem('➕ Add Vendor (إضافة مورد)', 'addNewVendor')
      .addItem('➕ Add Employee (إضافة موظف)', 'addNewEmployee')
      .addSeparator()
      .addItem('🔢 Generate Missing Codes (توليد الأكواد)', 'generateMissingClientCodes')
      .addSeparator()
      .addItem('📋 Add Client Sector (قطاع عميل)', 'addClientSector')
      .addSeparator()
      .addItem('📄 Client Statement (كشف حساب)', 'showClientStatement')
      .addItem('💹 Client Profitability (ربحية العميل)', 'showClientProfitability'))
    
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
      .addItem('➕ Add Sector (إضافة قطاع)', 'addNewSector')
      .addItem('🏷️ Sector Profiles (ملفات القطاعات)', 'showSectorProfiles')
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
    refreshSectorDropdown();
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
      patterns: ['Clients', 'Client Sector', 'Vendors', 'Employees', 'Items Database', 'Sector Profiles', 'Movement Types', 'Categories', 'Holidays', 'Cash Boxes', 'Bank Accounts']
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
