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
