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
  sheet.getRange(2, 2, lastRow, 1).setNumberFormat('dd.mm.yyyy');
  sheet.getRange(2, 11, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 13, lastRow, 1).setNumberFormat('#,##0.0000');
  sheet.getRange(2, 14, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 20, lastRow, 1).setNumberFormat('dd.mm.yyyy');
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
  const clientData = clientsSheet.getRange(2, 1, lastClientRow - 1, 17).getValues();

  const clientCodes = [];
  const clientNames = []; // سيشمل كل الأسماء (EN, AR, TR)

  clientData.forEach(row => {
    const code = row[0];      // A = Code
    const nameEN = row[1];    // B = Name EN
    const nameAR = row[2];    // C = Name AR
    const nameTR = row[3];    // D = Name TR
    const status = row[16];   // Q = Status (column 17)

    if (code && status === 'Active') {
      clientCodes.push(code);

      // إضافة جميع الأسماء المتاحة للـ dropdown
      if (nameEN && nameEN.toString().trim() !== '') {
        clientNames.push(nameEN);
      }
      if (nameAR && nameAR.toString().trim() !== '' && nameAR !== nameEN) {
        clientNames.push(nameAR);
      }
      if (nameTR && nameTR.toString().trim() !== '' && nameTR !== nameEN && nameTR !== nameAR) {
        clientNames.push(nameTR);
      }
    }
  });

  if (clientCodes.length === 0) return;

  // Client Code Dropdown (Column E)
  const codeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(clientCodes, true)
    .setAllowInvalid(true)
    .build();
  transSheet.getRange(2, 5, lastRow, 1).setDataValidation(codeRule);

  // Client Name Dropdown (Column F) - يشمل كل الأسماء
  if (clientNames.length > 0) {
    const nameRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(clientNames, true)
      .setAllowInvalid(true)
      .build();
    transSheet.getRange(2, 6, lastRow, 1).setDataValidation(nameRule);
  }
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

            // تحديث Show in Statement تلقائياً
            updateShowInStatement(sheet, row);

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

            // تحديث Show in Statement تلقائياً
            updateShowInStatement(sheet, row);

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

    // ───── Client Name (F) / Party Name (I) → Show in Statement (Y) ─────
    // التحقق من توافق اسم العميل واسم الطرف
    if (col === 6 || col === 9) {
      updateShowInStatement(sheet, row);
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

// ==================== 8.5 UPDATE SHOW IN STATEMENT ====================
/**
 * تحديث عمود Show in Statement بناءً على توافق اسم العميل واسم الطرف
 * إذا كان هناك توافق → Yes (نعم)
 * إذا كان هناك اختلاف → No (لا)
 */
function updateShowInStatement(sheet, row) {
  const clientName = sheet.getRange(row, 6).getValue(); // Column F - Client Name
  const partyName = sheet.getRange(row, 9).getValue();  // Column I - Party Name

  // إذا لم يكن هناك اسم عميل أو اسم طرف، لا تفعل شيء
  if (!clientName || !partyName) return;

  const clientNameStr = clientName.toString().trim().toLowerCase();
  const partyNameStr = partyName.toString().trim().toLowerCase();

  // التحقق من التوافق
  // يعتبر متوافق إذا:
  // 1. الاسمين متطابقين
  // 2. اسم العميل موجود داخل اسم الطرف (مثل: "ABC" موجود في "ABC (شركة أ ب ج)")
  // 3. اسم الطرف موجود داخل اسم العميل
  const isMatch = clientNameStr === partyNameStr ||
                  partyNameStr.includes(clientNameStr) ||
                  clientNameStr.includes(partyNameStr);

  // تعيين Show in Statement (Column Y = 25)
  if (isMatch) {
    sheet.getRange(row, 25).setValue('Yes (نعم)');
  } else {
    sheet.getRange(row, 25).setValue('No (لا)');
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

  // Ask user for date
  const todayFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
  const dateResponse = ui.prompt(
    '📅 تاريخ المعاملة (Transaction Date)',
    'أدخل التاريخ بصيغة dd.mm.yyyy\n' +
    'مثال: 15.03.2024\n\n' +
    'اتركه فارغاً لاستخدام تاريخ اليوم (' + todayFormatted + ')',
    ui.ButtonSet.OK_CANCEL
  );

  if (dateResponse.getSelectedButton() === ui.Button.CANCEL) {
    return;
  }

  let transactionDate = new Date();
  const dateInput = dateResponse.getResponseText().trim();

  if (dateInput !== '') {
    // Parse custom date in dd.mm.yyyy format
    const parsedDate = parseCustomDate(dateInput);
    if (parsedDate) {
      transactionDate = parsedDate;
    } else {
      ui.alert('❌ صيغة التاريخ غير صحيحة!\n\nالرجاء استخدام الصيغة: dd.mm.yyyy\nمثال: 15.03.2024');
      return;
    }
  }

  // Set auto number
  sheet.getRange(lastRow, 1).setValue(lastRow - 1);

  // Set date with dd.mm.yyyy format
  sheet.getRange(lastRow, 2).setValue(transactionDate).setNumberFormat('dd.mm.yyyy');

  // Set defaults
  sheet.getRange(lastRow, 12).setValue('TRY');
  sheet.getRange(lastRow, 13).setValue(1);
  sheet.getRange(lastRow, 19).setValue('Pending (معلق)');
  sheet.getRange(lastRow, 25).setValue('Yes (نعم)');

  // Select first input cell
  sheet.setActiveRange(sheet.getRange(lastRow, 3));

  const displayDate = Utilities.formatDate(transactionDate, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  ui.alert(
    '➕ Add Transaction (إضافة معاملة)\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Row #' + (lastRow - 1) + ' is ready.\n\n' +
    'Defaults:\n' +
    '• Date: ' + displayDate + '\n' +
    '• Currency: TRY\n' +
    '• Exchange Rate: 1\n' +
    '• Status: Pending\n' +
    '• Show in Statement: Yes\n\n' +
    '💡 Tips:\n' +
    '• اختر Client Code → الاسم يُملأ تلقائياً\n' +
    '• اختر Party Type → يتغير dropdown الأسماء'
  );
}

/**
 * Parse date from dd.mm.yyyy format
 * @param {string} dateStr - Date string in dd.mm.yyyy format (also supports dd.mm.yy)
 * @returns {Date|null} - Parsed date or null if invalid
 */
function parseCustomDate(dateStr) {
  // Support both dd.mm.yyyy and dd.mm.yy formats
  const regex = /^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/;
  const match = dateStr.match(regex);

  if (!match) return null;

  const day = parseInt(match[1], 10);
  const month = parseInt(match[2], 10) - 1; // JavaScript months are 0-indexed
  let year = parseInt(match[3], 10);

  // Handle 2-digit year (assume 20xx for years 00-99)
  if (year < 100) {
    year += 2000;
  }

  // Validate date parts
  if (day < 1 || day > 31 || month < 0 || month > 11) {
    return null;
  }

  const date = new Date(year, month, day);

  // Verify the date is valid (e.g., not Feb 30)
  if (date.getDate() !== day || date.getMonth() !== month || date.getFullYear() !== year) {
    return null;
  }

  return date;
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
