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
    ['Company Logo URL', 'https://drive.google.com/file/d/1retRm0IhrHep3s4BB0bIAhyvpdBIrSxm/view?usp=sharing'],
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
  
  sheet.getRange(2, 1, holidays2025.length, 1).setNumberFormat('dd.mm.yyyy');
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

// ==================== 6. SECTOR PROFILES SHEET (Unified) ====================
/**
 * Sector Profiles - unified sheet replacing old Activities + Activity Profiles
 * Each sector (Accounting, Consulting, etc.) has:
 * - Sector names (EN/AR/TR) for dropdowns
 * - Company branding (names, logo, website) for invoices
 * - Bank details per sector
 * Shared fields (Address, Phone, Email) come from Settings
 */
function createSectorProfilesSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Sector Profiles');
  if (sheet) ss.deleteSheet(sheet);

  sheet = ss.insertSheet('Sector Profiles');
  sheet.setTabColor('#00695c');

  const headers = [
    'Sector Code',          // A - e.g. ACC, CON, LOG
    'Sector Name (EN)',     // B
    'Sector Name (AR)',     // C
    'Sector Name (TR)',     // D
    'Company Name (EN)',    // E
    'Company Name (AR)',    // F
    'Company Name (TR)',    // G
    'Logo URL',             // H - Google Drive link
    'Website',              // I
    'Bank Name',            // J
    'IBAN TRY',             // K
    'IBAN USD',             // L
    'SWIFT Code',           // M
    'Status',               // N
    'Notes'                 // O
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(COLORS.header)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const widths = [100, 150, 140, 160, 200, 180, 200, 300, 200, 150, 260, 260, 120, 80, 200];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Default data
  const data = [
    ['ACC', 'Accounting',  'محاسبة',      'Muhasebe',      'Dewan Accounting',  'ديوان للمحاسبة',    'DİVAN MUHASEBECİLİK', '', '', 'Kuveyt Türk', '', '', 'KTEFTRIS', 'Active', ''],
    ['CON', 'Consulting',  'استشارات',     'Danışmanlık',   'Dewan Consulting',  'ديوان للاستشارات',   'DİVAN DANIŞMANLIK',   '', '', 'Kuveyt Türk', '', '', 'KTEFTRIS', 'Active', ''],
    ['LOG', 'Logistics',   'لوجستيات',     'Lojistik',      'Dewan Logistics',   'ديوان للوجستيات',    'DİVAN LOJİSTİK',      '', '', 'Kuveyt Türk', '', '', 'KTEFTRIS', 'Active', ''],
    ['TRD', 'Trading',     'تجارة',        'Ticaret',       'Dewan Trading',     'ديوان للتجارة',      'DİVAN TİCARET',        '', '', 'Kuveyt Türk', '', '', 'KTEFTRIS', 'Active', ''],
    ['INS', 'Inspection',  'تفتيش',        'Denetim',       'Dewan Inspection',  'ديوان للتفتيش',      'DİVAN DENETİM',        '', '', 'Kuveyt Türk', '', '', 'KTEFTRIS', 'Active', ''],
    ['TUR', 'Tourism',     'سياحة',        'Turizm',        'Dewan Tourism',     'ديوان للسياحة',      'DİVAN TURİZM',         '', '', 'Kuveyt Türk', '', '', 'KTEFTRIS', 'Active', '']
  ];

  sheet.getRange(2, 1, data.length, headers.length).setValues(data);

  const lastRow = 20;

  // Status validation (column N = 14)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive'], true)
    .build();
  sheet.getRange(2, 14, lastRow, 1).setDataValidation(statusRule);

  // Conditional formatting for Status
  const statusRange = sheet.getRange(2, 14, lastRow, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Active').setBackground(COLORS.success).setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Inactive').setBackground(COLORS.warning).setRanges([statusRange]).build()
  ]);

  sheet.setFrozenRows(1);

  // Notes
  sheet.getRange('A1').setNote('Sector Code: Short code (ACC, CON, LOG, TRD, INS, TUR)');
  sheet.getRange('B1').setNote('Sector Name EN - used in dropdowns and invoices');
  sheet.getRange('H1').setNote('Google Drive sharing link for the logo image');
  sheet.getRange('I1').setNote('Website URL for this sector');
  sheet.getRange('K1').setNote('IBAN for TRY transactions');
  sheet.getRange('L1').setNote('IBAN for USD transactions');

  applyAlternatingColors(sheet, 2, data.length, headers.length);

  return sheet;
}

/**
 * Add a new sector to Sector Profiles
 */
function addNewSector() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Sector Profiles');

  if (!sheet) {
    ui.alert('⚠️ Sector Profiles sheet not found!\n\nRun "Setup System" first.');
    return;
  }

  const lastRow = sheet.getLastRow() + 1;

  // Set defaults
  sheet.getRange(lastRow, 14).setValue('Active');

  sheet.setActiveRange(sheet.getRange(lastRow, 1));
  ss.setActiveSheet(sheet);

  ui.alert(
    '🏭 Add New Sector (إضافة قطاع جديد)\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Row: ' + lastRow + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'Required fields (الحقول المطلوبة):\n' +
    '• Sector Code (e.g. ACC, CON)\n' +
    '• Sector Name (EN/AR/TR)\n' +
    '• Company Name (EN/AR/TR)\n' +
    '• Bank Details'
  );
}

/**
 * Get list of active sectors for dropdowns
 * Replaces old getActivitiesList()
 * @returns {Array} - List of active sectors
 */
function getSectorsList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Sector Profiles');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const sectors = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][13] === 'Active' && data[i][1]) { // Status=N(14), NameEN=B(2)
      sectors.push({
        code: data[i][0],
        nameEN: data[i][1],
        nameAR: data[i][2],
        nameTR: data[i][3],
        companyNameEN: data[i][4],
        companyNameAR: data[i][5],
        companyNameTR: data[i][6],
        display: data[i][1] + ' (' + (data[i][2] || data[i][1]) + ')'
      });
    }
  }
  return sectors;
}

// ==================== 6b. CLIENT SECTOR SHEET ====================
function createClientSectorSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Client Sector');
  if (sheet) ss.deleteSheet(sheet);

  sheet = ss.insertSheet('Client Sector');
  sheet.setTabColor('#00838f');

  const headers = [
    'Client Code',       // A
    'Client Name',       // B
    'Sector',            // C
    'Fee Type',          // D - Monthly / Per-Job
    'Monthly Fee',       // E
    'Currency',          // F
    'Start Date',        // G
    'Status',            // H
    'Notes'              // I
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(COLORS.header)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const widths = [100, 200, 140, 100, 120, 80, 110, 80, 200];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  const lastRow = 500;

  // Fee Type validation (column D)
  const feeTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Monthly', 'Per-Job'], true)
    .build();
  sheet.getRange(2, 4, lastRow, 1).setDataValidation(feeTypeRule);

  // Sector validation (column C)
  const sectorRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Accounting', 'Consulting', 'Logistics', 'Trading', 'Inspection', 'Tourism', 'Other'], true)
    .build();
  sheet.getRange(2, 3, lastRow, 1).setDataValidation(sectorRule);

  // Currency validation (column F)
  const currencyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CURRENCIES, true)
    .build();
  sheet.getRange(2, 6, lastRow, 1).setDataValidation(currencyRule);

  // Status validation (column H)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive'], true)
    .build();
  sheet.getRange(2, 8, lastRow, 1).setDataValidation(statusRule);

  // Number formats
  sheet.getRange(2, 5, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 7, lastRow, 1).setNumberFormat('dd.mm.yyyy');

  // Conditional formatting for Fee Type
  const feeTypeRange = sheet.getRange(2, 4, lastRow, 1);
  const statusRange = sheet.getRange(2, 8, lastRow, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Monthly').setBackground('#bbdefb').setRanges([feeTypeRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Per-Job').setBackground('#e1bee7').setRanges([feeTypeRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Active').setBackground(COLORS.success).setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Inactive').setBackground(COLORS.warning).setRanges([statusRange]).build()
  ]);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  // Add notes
  sheet.getRange('D1').setNote('Monthly = فيز شهري ثابت (Accounting/Consulting)\nPer-Job = حسب المعاملة (Logistics/Inspection/Trading/Tourism)');
  sheet.getRange('E1').setNote('Monthly Fee: Only for Monthly fee type activities (Accounting/Consulting)');

  return sheet;
}

function addClientSector() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Client Sector');

  if (!sheet) {
    ui.alert('⚠️ Client Sector sheet not found!\n\nRun "Setup System" first.');
    return;
  }

  const lastRow = sheet.getLastRow() + 1;

  // Set defaults
  sheet.getRange(lastRow, 4).setValue('Monthly'); // Fee Type
  sheet.getRange(lastRow, 6).setValue('TRY'); // Currency
  sheet.getRange(lastRow, 7).setValue(new Date()); // Start Date
  sheet.getRange(lastRow, 8).setValue('Active'); // Status

  sheet.setActiveRange(sheet.getRange(lastRow, 1));
  ss.setActiveSheet(sheet);

  ui.alert(
    '📋 Add Client Sector (إضافة قطاع عميل)\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Row: ' + lastRow + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'Required fields:\n' +
    '• Client Code\n' +
    '• Sector (Accounting/Consulting/Logistics/...)\n' +
    '• Fee Type (Monthly/Per-Job)\n' +
    '• Monthly Fee (for Monthly type only)'
  );
}

/**
 * Get all client activities, optionally filtered
 * @param {string} [clientCode] - Filter by client code
 * @param {string} [feeType] - Filter by fee type ('Monthly' or 'Per-Job')
 * @returns {Array} - List of client activities
 */
function getClientSectorList(clientCode, feeType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Client Sector');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const activities = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][7] !== 'Active' || !data[i][0]) continue;
    if (clientCode && data[i][0] !== clientCode) continue;
    if (feeType && data[i][3] !== feeType) continue;

    activities.push({
      clientCode: data[i][0],
      clientName: data[i][1],
      activity: data[i][2],
      feeType: data[i][3],
      monthlyFee: data[i][4] || 0,
      currency: data[i][5] || 'TRY',
      startDate: data[i][6],
      status: data[i][7],
      notes: data[i][8] || ''
    });
  }
  return activities;
}

/**
 * Get clients with monthly fees (from Client Sector sheet)
 * Used for monthly invoice generation
 * @returns {Array} - List of {clientCode, clientName, activity, monthlyFee, currency}
 */
function getClientsWithMonthlyFees() {
  return getClientSectorList(null, 'Monthly').filter(a => a.monthlyFee > 0);
}

// ==================== 7. SECTOR PROFILE FUNCTIONS ====================
/**
 * Get sector profile (branding) for a specific sector
 * Falls back to Settings for shared fields (address, phone, email)
 * Replaces old getActivityProfile()
 * @param {string} sectorName - e.g. 'Accounting', 'Consulting'
 * @returns {object} - Sector profile with branding info
 */
function getSectorProfile(sectorName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Sector Profiles');

  // Default: use Settings if no Sector Profiles sheet
  if (!sheet || !sectorName) {
    return getDefaultProfile();
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    // Match by Sector Name EN (col B=1) or Sector Code (col A=0)
    if ((data[i][1] === sectorName || data[i][0] === sectorName) && data[i][13] === 'Active') {
      return {
        sector: data[i][1],
        sectorCode: data[i][0],
        sectorNameAR: data[i][2] || '',
        sectorNameTR: data[i][3] || '',
        companyNameEN: data[i][4] || getSettingValue('Company Name (EN)') || '',
        companyNameAR: data[i][5] || getSettingValue('Company Name (AR)') || '',
        companyNameTR: data[i][6] || getSettingValue('Company Name (TR)') || '',
        logoUrl: data[i][7] || getSettingValue('Company Logo URL') || '',
        website: data[i][8] || '',
        bankName: data[i][9] || getSettingValue('Bank Name') || '',
        ibanTRY: data[i][10] || getSettingValue('IBAN TRY') || '',
        ibanUSD: data[i][11] || getSettingValue('IBAN USD') || '',
        swiftCode: data[i][12] || getSettingValue('SWIFT Code') || '',
        // Shared fields from Settings
        address: getSettingValue('Company Address') || '',
        phone: getSettingValue('Company Phone') || '',
        email: getSettingValue('Company Email') || ''
      };
    }
  }

  // Sector not found - use defaults
  return getDefaultProfile();
}

// Backward-compatible alias
function getActivityProfile(activityName) {
  return getSectorProfile(activityName);
}

/**
 * Get default profile from Settings (fallback)
 */
function getDefaultProfile() {
  return {
    sector: '',
    sectorCode: '',
    sectorNameAR: '',
    sectorNameTR: '',
    companyNameEN: getSettingValue('Company Name (EN)') || 'Dewan Consulting',
    companyNameAR: getSettingValue('Company Name (AR)') || 'ديوان للاستشارات',
    companyNameTR: getSettingValue('Company Name (TR)') || 'DİVAN DANIŞMANLIK',
    logoUrl: getSettingValue('Company Logo URL') || '',
    website: '',
    bankName: getSettingValue('Bank Name') || 'Kuveyt Türk',
    ibanTRY: getSettingValue('IBAN TRY') || '',
    ibanUSD: getSettingValue('IBAN USD') || '',
    swiftCode: getSettingValue('SWIFT Code') || 'KTEFTRIS',
    address: getSettingValue('Company Address') || '',
    phone: getSettingValue('Company Phone') || '',
    email: getSettingValue('Company Email') || ''
  };
}

/**
 * Get client's primary sector from Client Sector sheet
 * @param {string} clientCode - Client code
 * @returns {string} - Sector name (e.g. 'Accounting') or empty string
 */
function getClientPrimarySector(clientCode) {
  const activities = getClientSectorList(clientCode);
  if (activities.length > 0) {
    return activities[0].activity;
  }
  return '';
}

// Backward-compatible alias
function getClientPrimaryActivity(clientCode) {
  return getClientPrimarySector(clientCode);
}

function showSectorProfiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Sector Profiles');
  if (sheet) ss.setActiveSheet(sheet);
  else SpreadsheetApp.getUi().alert('⚠️ Sector Profiles sheet not found!\n\nRun "Setup System" first.');
}

// ==================== 8. HELPER: ALTERNATING COLORS ====================
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

// ==================== 9. GET FUNCTIONS ====================
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

// ==================== END OF PART 2 (v3.1) ====================
