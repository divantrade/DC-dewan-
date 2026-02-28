// ╔════════════════════════════════════════════════════════════════════════════╗
// ║                    DC CONSULTING ACCOUNTING SYSTEM v3.1                     ║
// ║                              Part 11                                       ║
// ║               Import Excel Data (استيراد بيانات من إكسيل)                  ║
// ╚════════════════════════════════════════════════════════════════════════════╝

// ==================== 1. CREATE IMPORT SHEET ====================
/**
 * إنشاء شيت Import Data مع الهيدرات والتنسيق
 * المستخدم يلصق بيانات Excel هنا ثم يضغط "Import"
 */
function createImportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let sheet = ss.getSheetByName('Import Data');
  if (sheet) {
    const confirm = ui.alert(
      '⚠️ شيت Import Data موجود بالفعل\n\n' +
      'هل تريد إعادة إنشائه؟ (البيانات الحالية ستُحذف)',
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet('Import Data');
  sheet.setTabColor('#ff6f00');

  // === Headers (18 columns) ===
  const headers = [
    'Date\nالتاريخ',                    // A (1)
    'Sector\nالقطاع',                   // B (2)
    'Movement Type\nنوع الحركة',        // C (3)
    'Category\nالتصنيف',               // D (4)
    'Client Code\nكود العميل',          // E (5)
    'Client Name\nاسم العميل',          // F (6)
    'Item\nالبند',                      // G (7)
    'Description\nالوصف',               // H (8)
    'Party Name\nاسم الطرف',            // I (9)
    'Party Type\nنوع الطرف',            // J (10)
    'Amount\nالمبلغ',                   // K (11)
    'Currency\nالعملة',                 // L (12)
    'Exchange Rate\nسعر الصرف',         // M (13)
    'Payment Method\nطريقة الدفع',      // N (14)
    'Cash/Bank\nالخزينة/البنك',         // O (15)
    'Reference\nالمرجع',               // P (16)
    'Status\nالحالة',                   // Q (17)
    'Notes\nملاحظات'                    // R (18)
  ];

  // Header row
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#ff6f00')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setRowHeight(1, 55);

  // Column widths
  const widths = [100, 150, 200, 180, 110, 180, 160, 200, 180, 130, 110, 80, 90, 160, 170, 110, 130, 200];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // === Instructions row (row 2) ===
  const instructions = [
    'dd.mm.yyyy\nأو yyyy-mm-dd',
    'مثال:\nAccounting (محاسبة)',
    'مثال:\nOpening Balance (رصيد افتتاحي)',
    'مثال:\nOpening Balance (رصيد افتتاحي)',
    'CLT-001',
    'اسم الشركة',
    'مثال:\nConsulting (استشارات)',
    'وصف المعاملة',
    'اسم الطرف',
    'Client/Vendor/\nEmployee/Internal',
    'رقم فقط\n1000.50',
    'TRY/USD/EUR\nSAR/EGP/AED/GBP',
    'رقم\n1 للـ TRY',
    'Cash/Bank Transfer/\nAccrual',
    'اسم الخزينة أو البنك',
    'رقم المرجع',
    'Paid/Pending/\nPartial',
    'ملاحظات إضافية'
  ];

  sheet.getRange(2, 1, 1, instructions.length)
    .setValues([instructions])
    .setBackground('#fff3e0')
    .setFontColor('#e65100')
    .setFontSize(8)
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setRowHeight(2, 60);

  // Data starts from row 3
  const dataRows = 500;

  // Date format
  sheet.getRange(3, 1, dataRows, 1).setNumberFormat('dd.mm.yyyy');

  // Amount format
  sheet.getRange(3, 11, dataRows, 1).setNumberFormat('#,##0.00');

  // Exchange Rate format
  sheet.getRange(3, 13, dataRows, 1).setNumberFormat('#,##0.0000');

  // === Control section (status/buttons area) ===
  // Add control area on the right side
  sheet.getRange(1, 20).setValue('IMPORT STATUS').setFontWeight('bold').setBackground('#1565c0').setFontColor('#ffffff');
  sheet.getRange(2, 20).setValue('Ready').setBackground('#c8e6c9').setFontWeight('bold');
  sheet.getRange(3, 20).setValue('Rows to import:').setFontWeight('bold');
  sheet.getRange(3, 21).setValue(0);
  sheet.getRange(4, 20).setValue('Errors:').setFontWeight('bold');
  sheet.getRange(4, 21).setValue(0);

  sheet.setColumnWidth(20, 130);
  sheet.setColumnWidth(21, 80);

  sheet.setFrozenRows(2);

  // Add sheet note
  sheet.getRange('A1').setNote(
    '📋 كيفية استخدام شيت الاستيراد:\n\n' +
    '1. الصق بيانات Excel من السطر 3\n' +
    '2. تأكد من ترتيب الأعمدة\n' +
    '3. من القائمة: DC Consulting → Import → Import from Sheet\n\n' +
    '💡 نصائح:\n' +
    '• التاريخ: dd.mm.yyyy أو yyyy-mm-dd\n' +
    '• العملة: TRY, USD, EUR, SAR, EGP, AED, GBP\n' +
    '• الحالة: Paid, Pending, Partial\n' +
    '• نوع الطرف: Client, Vendor, Employee, Internal'
  );

  ss.setActiveSheet(sheet);
  sheet.setActiveRange(sheet.getRange('A3'));

  ui.alert(
    '✅ Import Data Sheet Created!\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'الخطوات:\n\n' +
    '1. الصق بيانات Excel من السطر 3\n' +
    '   (السطر 1 = عناوين، السطر 2 = تعليمات)\n\n' +
    '2. تأكد من ترتيب الأعمدة:\n' +
    '   A=التاريخ, B=القطاع, C=نوع الحركة...\n\n' +
    '3. من القائمة:\n' +
    '   DC Consulting → 📥 Import → Import from Sheet\n\n' +
    '💡 للأرصدة الافتتاحية:\n' +
    '   DC Consulting → 📥 Import → Import Opening Balances'
  );

  return sheet;
}

// ==================== 2. CREATE OPENING BALANCES IMPORT SHEET ====================
/**
 * شيت مخصص لاستيراد الأرصدة الافتتاحية - أبسط من الشيت العام
 */
function createOpeningBalancesImportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let sheet = ss.getSheetByName('Import Opening Balances');
  if (sheet) {
    const confirm = ui.alert(
      '⚠️ شيت Import Opening Balances موجود بالفعل\n\n' +
      'هل تريد إعادة إنشائه؟',
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet('Import Opening Balances');
  sheet.setTabColor('#4caf50');

  // Headers (7 columns - simplified)
  const headers = [
    'Date\nالتاريخ',              // A
    'Account Type\nنوع الحساب',    // B (Cash/Bank/Client/Vendor)
    'Account Name\nاسم الحساب',    // C
    'Amount\nالمبلغ',              // D
    'Currency\nالعملة',            // E
    'Sector\nالقطاع',             // F
    'Notes\nملاحظات'               // G
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#4caf50')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setRowHeight(1, 50);

  // Instructions
  const instructions = [
    'dd.mm.yyyy',
    'Cash / Bank /\nClient / Vendor',
    'اسم الخزينة أو البنك\nأو العميل أو المورد',
    'رقم فقط\nموجب = رصيد دائن',
    'TRY/USD/EUR...',
    'القطاع (اختياري)',
    'ملاحظات'
  ];

  sheet.getRange(2, 1, 1, instructions.length)
    .setValues([instructions])
    .setBackground('#e8f5e9')
    .setFontColor('#2e7d32')
    .setFontSize(8)
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setWrap(true);

  sheet.setRowHeight(2, 50);

  // Column widths
  [100, 130, 200, 120, 80, 150, 200].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Formats
  sheet.getRange(3, 1, 500, 1).setNumberFormat('dd.mm.yyyy');
  sheet.getRange(3, 4, 500, 1).setNumberFormat('#,##0.00');

  // Account Type dropdown
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Cash', 'Bank', 'Client', 'Vendor'], true)
    .setAllowInvalid(false).build();
  sheet.getRange(3, 2, 500, 1).setDataValidation(typeRule);

  // Currency dropdown
  const currRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CURRENCIES, true)
    .setAllowInvalid(false).build();
  sheet.getRange(3, 5, 500, 1).setDataValidation(currRule);

  // Status area
  sheet.getRange(1, 9).setValue('STATUS').setFontWeight('bold').setBackground('#4caf50').setFontColor('#ffffff');
  sheet.getRange(2, 9).setValue('Ready').setBackground('#c8e6c9').setFontWeight('bold');

  sheet.setFrozenRows(2);
  ss.setActiveSheet(sheet);
  sheet.setActiveRange(sheet.getRange('A3'));

  ui.alert(
    '✅ Opening Balances Import Sheet Created!\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'الخطوات:\n' +
    '1. الصق بيانات الأرصدة الافتتاحية من السطر 3\n' +
    '2. حدد نوع الحساب (Cash/Bank/Client/Vendor)\n' +
    '3. من القائمة:\n' +
    '   DC Consulting → 📥 Import → Import Opening Balances\n\n' +
    '💡 مثال:\n' +
    '   01.01.2025 | Cash | Cash TRY - Main | 50000 | TRY\n' +
    '   01.01.2025 | Bank | Kuveyt Turk     | 120000| TRY\n' +
    '   01.01.2025 | Client | ABC Company   | 15000 | USD'
  );

  return sheet;
}

// ==================== 3. VALIDATE IMPORT DATA ====================
/**
 * تحقق من صحة البيانات قبل الاستيراد
 * @param {Array} data - البيانات المراد التحقق منها
 * @param {string} importType - نوع الاستيراد ('transactions' أو 'opening')
 * @returns {Object} - {valid: boolean, errors: [], validRows: [], skippedRows: []}
 */
function validateImportData(data, importType) {
  const errors = [];
  const validRows = [];
  const skippedRows = [];

  const validCurrencies = new Set(CURRENCIES);

  // Valid movement types (English part only for matching)
  const validMovementTypes = new Set([
    'Revenue Accrual', 'Revenue Collection', 'Expense Accrual', 'Expense Payment',
    'Cash Transfer', 'Bank Transfer', 'Cash to Bank', 'Bank to Cash',
    'Currency Exchange', 'Adjustment Add', 'Adjustment Deduct', 'Opening Balance',
    'Advance Issue', 'Advance Return', 'Expense'
  ]);

  const validPartyTypes = new Set(['Client', 'Vendor', 'Employee', 'Internal']);
  const validStatuses = new Set(['Paid', 'Pending', 'Partial', 'Cancelled']);
  const validPaymentMethods = new Set(['Cash', 'Bank Transfer', 'Accrual', 'Credit Card', 'Advance']);

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + 3; // Data starts from row 3
    const rowErrors = [];

    if (importType === 'transactions') {
      // Check if row is empty
      const hasData = row.some(cell => cell !== '' && cell !== null && cell !== undefined);
      if (!hasData) {
        skippedRows.push(rowNum);
        continue;
      }

      // Date (required)
      const dateVal = row[0];
      if (!dateVal) {
        rowErrors.push('Date is required');
      } else {
        const parsed = parseImportDate(dateVal);
        if (!parsed) {
          rowErrors.push('Invalid date format: ' + dateVal);
        }
      }

      // Movement Type (required)
      const movType = row[2];
      if (!movType) {
        rowErrors.push('Movement Type is required');
      } else {
        const movTypeEN = extractEnglishPart(movType.toString());
        if (!validMovementTypes.has(movTypeEN)) {
          rowErrors.push('Invalid Movement Type: ' + movType);
        }
      }

      // Amount (required)
      const amount = row[10];
      if (!amount && amount !== 0) {
        rowErrors.push('Amount is required');
      } else if (isNaN(parseFloat(amount))) {
        rowErrors.push('Invalid amount: ' + amount);
      }

      // Currency (required)
      const currency = row[11];
      if (!currency) {
        rowErrors.push('Currency is required');
      } else if (!validCurrencies.has(currency.toString().toUpperCase().trim())) {
        rowErrors.push('Invalid currency: ' + currency);
      }

      // Party Type (optional but must be valid if provided)
      const partyType = row[9];
      if (partyType) {
        const ptEN = extractEnglishPart(partyType.toString());
        if (!validPartyTypes.has(ptEN)) {
          rowErrors.push('Invalid Party Type: ' + partyType);
        }
      }

      // Status (optional but must be valid if provided)
      const status = row[16];
      if (status) {
        const stEN = extractEnglishPart(status.toString());
        if (!validStatuses.has(stEN)) {
          rowErrors.push('Invalid Status: ' + status);
        }
      }

      // Payment Method (optional but must be valid if provided)
      const payMethod = row[13];
      if (payMethod) {
        const pmEN = extractEnglishPart(payMethod.toString());
        if (!validPaymentMethods.has(pmEN)) {
          rowErrors.push('Invalid Payment Method: ' + payMethod);
        }
      }

    } else if (importType === 'opening') {
      // Opening Balances validation
      const hasData = row.some(cell => cell !== '' && cell !== null && cell !== undefined);
      if (!hasData) {
        skippedRows.push(rowNum);
        continue;
      }

      // Date (required)
      if (!row[0]) {
        rowErrors.push('Date is required');
      } else {
        const parsed = parseImportDate(row[0]);
        if (!parsed) {
          rowErrors.push('Invalid date format: ' + row[0]);
        }
      }

      // Account Type (required)
      const accType = row[1];
      if (!accType) {
        rowErrors.push('Account Type is required');
      } else if (!['Cash', 'Bank', 'Client', 'Vendor'].includes(accType.toString().trim())) {
        rowErrors.push('Invalid Account Type: ' + accType + ' (use: Cash, Bank, Client, or Vendor)');
      }

      // Account Name (required)
      if (!row[2]) {
        rowErrors.push('Account Name is required');
      }

      // Amount (required)
      if (!row[3] && row[3] !== 0) {
        rowErrors.push('Amount is required');
      } else if (isNaN(parseFloat(row[3]))) {
        rowErrors.push('Invalid amount: ' + row[3]);
      }

      // Currency (required)
      if (!row[4]) {
        rowErrors.push('Currency is required');
      } else if (!validCurrencies.has(row[4].toString().toUpperCase().trim())) {
        rowErrors.push('Invalid currency: ' + row[4]);
      }
    }

    if (rowErrors.length > 0) {
      errors.push({ row: rowNum, errors: rowErrors });
    } else {
      validRows.push({ rowIndex: i, rowNum: rowNum, data: row });
    }
  }

  return {
    valid: errors.length === 0,
    errors: errors,
    validRows: validRows,
    skippedRows: skippedRows,
    totalRows: data.length,
    errorCount: errors.length,
    validCount: validRows.length,
    skippedCount: skippedRows.length
  };
}

// ==================== 4. IMPORT TRANSACTIONS FROM SHEET ====================
/**
 * استيراد المعاملات من شيت Import Data إلى شيت Transactions
 */
function importTransactionsFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const importSheet = ss.getSheetByName('Import Data');
  if (!importSheet) {
    ui.alert('❌ شيت Import Data غير موجود!\n\nاستخدم: DC Consulting → 📥 Import → Create Import Sheet');
    return;
  }

  const transSheet = ss.getSheetByName('Transactions');
  if (!transSheet) {
    ui.alert('❌ شيت Transactions غير موجود!\n\nقم بإعداد النظام أولاً.');
    return;
  }

  // Read import data (from row 3, skip headers and instructions)
  const lastRow = importSheet.getLastRow();
  if (lastRow < 3) {
    ui.alert('⚠️ لا توجد بيانات للاستيراد!\n\nالصق البيانات من السطر 3.');
    return;
  }

  const dataRange = importSheet.getRange(3, 1, lastRow - 2, 18);
  const data = dataRange.getValues();

  // Update status
  importSheet.getRange(2, 20).setValue('Validating...').setBackground('#fff9c4');

  // Validate
  const validation = validateImportData(data, 'transactions');

  importSheet.getRange(3, 21).setValue(validation.validCount);
  importSheet.getRange(4, 21).setValue(validation.errorCount);

  // Show validation results
  if (validation.errorCount > 0) {
    let errorMsg = '⚠️ تم العثور على ' + validation.errorCount + ' أخطاء:\n\n';
    validation.errors.forEach(err => {
      errorMsg += '• Row ' + err.row + ': ' + err.errors.join(', ') + '\n';
    });
    errorMsg += '\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
    errorMsg += 'صالح: ' + validation.validCount + ' | أخطاء: ' + validation.errorCount + ' | فارغ: ' + validation.skippedCount;
    errorMsg += '\n\nهل تريد استيراد الصفوف الصالحة فقط؟';

    const proceed = ui.alert('Validation Results', errorMsg, ui.ButtonSet.YES_NO);
    if (proceed !== ui.Button.YES) {
      importSheet.getRange(2, 20).setValue('Cancelled').setBackground('#ffcdd2');
      return;
    }
  }

  if (validation.validCount === 0) {
    ui.alert('❌ لا توجد صفوف صالحة للاستيراد!');
    importSheet.getRange(2, 20).setValue('No valid data').setBackground('#ffcdd2');
    return;
  }

  // Confirm import
  const confirm = ui.alert(
    '📥 تأكيد الاستيراد\n\n' +
    'سيتم استيراد ' + validation.validCount + ' معاملة إلى شيت Transactions.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    importSheet.getRange(2, 20).setValue('Cancelled').setBackground('#ffcdd2');
    return;
  }

  importSheet.getRange(2, 20).setValue('Importing...').setBackground('#bbdefb');

  // Get current last row in Transactions
  let transLastRow = transSheet.getLastRow();
  let imported = 0;

  // Process each valid row
  validation.validRows.forEach(item => {
    const row = item.data;
    transLastRow++;

    // Parse date
    const date = parseImportDate(row[0]);

    // Map Movement Type to full bilingual format
    const movementType = mapToDropdownValue(row[2], DROPDOWN_VALUES.movementTypes);

    // Map Category
    const category = mapToDropdownValue(row[3], DROPDOWN_VALUES.categories);

    // Map Party Type
    const partyType = row[9] ? mapToDropdownValue(row[9], DROPDOWN_VALUES.partyTypes) : '';

    // Map Payment Method
    const paymentMethod = row[13] ? mapToDropdownValue(row[13], DROPDOWN_VALUES.paymentMethods) : '';

    // Map Status
    const statusVal = row[16] ? mapToDropdownValue(row[16], DROPDOWN_VALUES.paymentStatus) : 'Pending (معلق)';

    // Parse amounts
    const amount = parseFloat(row[10]) || 0;
    const currency = (row[11] || 'TRY').toString().toUpperCase().trim();
    const exchangeRate = parseFloat(row[12]) || (currency === 'TRY' ? 1 : 1);
    const amountTRY = currency === 'TRY' ? amount : amount * exchangeRate;

    // Map Sector
    const sector = row[1] ? mapToSectorValue(ss, row[1]) : '';

    // Map Cash/Bank
    const cashBank = row[14] ? mapToCashBankValue(ss, row[14]) : '';

    // Build transaction row (26 columns: A-Z)
    const transRow = [
      transLastRow - 1,       // A: # (auto number)
      date,                   // B: Date
      sector,                 // C: Sector
      movementType,           // D: Movement Type
      category,               // E: Category
      row[4] || '',           // F: Client Code
      row[5] || '',           // G: Client Name
      row[6] || '',           // H: Item
      row[7] || '',           // I: Description
      row[8] || '',           // J: Party Name
      partyType,              // K: Party Type
      amount,                 // L: Amount
      currency,               // M: Currency
      exchangeRate,           // N: Exchange Rate
      amountTRY,              // O: Amount TRY
      paymentMethod,          // P: Payment Method
      cashBank,               // Q: Cash/Bank
      row[15] || '',          // R: Reference
      '',                     // S: Invoice No
      statusVal,              // T: Status
      '',                     // U: Due Date
      amount,                 // V: Paid Amount (defaults to full if Paid)
      0,                      // W: Remaining
      row[17] || '',          // X: Notes
      '',                     // Y: Attachment
      'Yes (نعم)'            // Z: Show in Statement
    ];

    // Adjust Paid/Remaining based on status
    if (statusVal.includes('Paid')) {
      transRow[21] = amount; // V: Paid = Amount
      transRow[22] = 0;      // W: Remaining = 0
    } else if (statusVal.includes('Pending')) {
      transRow[21] = 0;      // V: Paid = 0
      transRow[22] = amount;  // W: Remaining = Amount
    }

    transSheet.getRange(transLastRow, 1, 1, 26).setValues([transRow]);
    transSheet.getRange(transLastRow, 2).setNumberFormat('dd.mm.yyyy');

    imported++;
  });

  // Mark imported rows in Import sheet
  validation.validRows.forEach(item => {
    importSheet.getRange(item.rowNum, 1, 1, 18).setBackground('#c8e6c9');
  });

  // Mark error rows
  validation.errors.forEach(err => {
    importSheet.getRange(err.row, 1, 1, 18).setBackground('#ffcdd2');
  });

  importSheet.getRange(2, 20).setValue('Done ✅').setBackground('#c8e6c9');

  ui.alert(
    '✅ Import Complete!\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Imported: ' + imported + ' transactions\n' +
    'Skipped (empty): ' + validation.skippedCount + '\n' +
    'Errors: ' + validation.errorCount + '\n\n' +
    '🟢 Green rows = imported successfully\n' +
    '🔴 Red rows = errors (not imported)\n\n' +
    '💡 Check Transactions sheet for the imported data.'
  );

  // Navigate to Transactions
  ss.setActiveSheet(transSheet);
}

// ==================== 5. IMPORT OPENING BALANCES ====================
/**
 * استيراد الأرصدة الافتتاحية من شيت Import Opening Balances
 */
function importOpeningBalances() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const importSheet = ss.getSheetByName('Import Opening Balances');
  if (!importSheet) {
    ui.alert('❌ شيت Import Opening Balances غير موجود!\n\nاستخدم: DC Consulting → 📥 Import → Create Opening Balances Sheet');
    return;
  }

  const transSheet = ss.getSheetByName('Transactions');
  if (!transSheet) {
    ui.alert('❌ شيت Transactions غير موجود!');
    return;
  }

  const lastRow = importSheet.getLastRow();
  if (lastRow < 3) {
    ui.alert('⚠️ لا توجد بيانات للاستيراد!');
    return;
  }

  const data = importSheet.getRange(3, 1, lastRow - 2, 7).getValues();

  // Validate
  const validation = validateImportData(data, 'opening');

  importSheet.getRange(2, 9).setValue('Validating...').setBackground('#fff9c4');

  if (validation.errorCount > 0) {
    let errorMsg = '⚠️ أخطاء في البيانات:\n\n';
    validation.errors.forEach(err => {
      errorMsg += '• Row ' + err.row + ': ' + err.errors.join(', ') + '\n';
    });
    errorMsg += '\nهل تريد استيراد الصفوف الصالحة فقط؟';

    const proceed = ui.alert('Validation', errorMsg, ui.ButtonSet.YES_NO);
    if (proceed !== ui.Button.YES) {
      importSheet.getRange(2, 9).setValue('Cancelled').setBackground('#ffcdd2');
      return;
    }
  }

  if (validation.validCount === 0) {
    ui.alert('❌ لا توجد بيانات صالحة!');
    importSheet.getRange(2, 9).setValue('No valid data').setBackground('#ffcdd2');
    return;
  }

  const confirm = ui.alert(
    '📥 تأكيد استيراد الأرصدة الافتتاحية\n\n' +
    'سيتم إضافة ' + validation.validCount + ' رصيد افتتاحي\n' +
    'إلى شيت Transactions.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    importSheet.getRange(2, 9).setValue('Cancelled').setBackground('#ffcdd2');
    return;
  }

  importSheet.getRange(2, 9).setValue('Importing...').setBackground('#bbdefb');

  let transLastRow = transSheet.getLastRow();
  let imported = 0;

  validation.validRows.forEach(item => {
    const row = item.data;
    transLastRow++;

    const date = parseImportDate(row[0]);
    const accType = row[1].toString().trim();
    const accName = row[2].toString().trim();
    const amount = parseFloat(row[3]) || 0;
    const currency = (row[4] || 'TRY').toString().toUpperCase().trim();
    const sector = row[5] ? mapToSectorValue(ss, row[5]) : '';
    const notes = row[6] || '';

    // Determine Party Type and Payment Method based on account type
    let partyType = '';
    let paymentMethod = '';
    let cashBankVal = '';
    let clientCode = '';
    let clientName = '';

    if (accType === 'Cash') {
      partyType = 'Internal (داخلي)';
      paymentMethod = 'Cash (نقدي)';
      cashBankVal = mapToCashBankValue(ss, accName);
    } else if (accType === 'Bank') {
      partyType = 'Internal (داخلي)';
      paymentMethod = 'Bank Transfer (تحويل بنكي)';
      cashBankVal = mapToCashBankValue(ss, accName);
    } else if (accType === 'Client') {
      partyType = 'Client (عميل)';
      paymentMethod = 'Accrual (استحقاق)';
      clientName = accName;
      // Try to find client code
      clientCode = findClientCode(ss, accName);
    } else if (accType === 'Vendor') {
      partyType = 'Vendor (مورد)';
      paymentMethod = 'Accrual (استحقاق)';
    }

    const transRow = [
      transLastRow - 1,                           // A: #
      date,                                        // B: Date
      sector,                                      // C: Sector
      'Opening Balance (رصيد افتتاحي)',           // D: Movement Type
      'Opening Balance (رصيد افتتاحي)',           // E: Category
      clientCode,                                  // F: Client Code
      clientName,                                  // G: Client Name
      '',                                          // H: Item
      'Opening Balance - ' + accName,              // I: Description
      accName,                                     // J: Party Name
      partyType,                                   // K: Party Type
      amount,                                      // L: Amount
      currency,                                    // M: Currency
      currency === 'TRY' ? 1 : 1,                 // N: Exchange Rate
      amount,                                      // O: Amount TRY (approximate)
      paymentMethod,                               // P: Payment Method
      cashBankVal,                                 // Q: Cash/Bank
      'OB-' + formatDate(date, 'yyyyMM'),          // R: Reference
      '',                                          // S: Invoice No
      'Paid (مدفوع)',                              // T: Status
      '',                                          // U: Due Date
      amount,                                      // V: Paid Amount
      0,                                           // W: Remaining
      notes || 'Opening Balance',                  // X: Notes
      '',                                          // Y: Attachment
      'Yes (نعم)'                                  // Z: Show in Statement
    ];

    transSheet.getRange(transLastRow, 1, 1, 26).setValues([transRow]);
    transSheet.getRange(transLastRow, 2).setNumberFormat('dd.mm.yyyy');

    imported++;
  });

  // Color imported rows
  validation.validRows.forEach(item => {
    importSheet.getRange(item.rowNum, 1, 1, 7).setBackground('#c8e6c9');
  });
  validation.errors.forEach(err => {
    importSheet.getRange(err.row, 1, 1, 7).setBackground('#ffcdd2');
  });

  importSheet.getRange(2, 9).setValue('Done ✅').setBackground('#c8e6c9');

  ui.alert(
    '✅ Opening Balances Imported!\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Imported: ' + imported + ' opening balances\n' +
    'Errors: ' + validation.errorCount + '\n\n' +
    '💡 Check Transactions sheet.\n' +
    '💡 Use "Sync to Cash/Bank" to update account sheets.'
  );

  ss.setActiveSheet(transSheet);
}

// ==================== 6. HELPER FUNCTIONS ====================

/**
 * تحويل تاريخ من أي صيغة إلى Date object
 */
function parseImportDate(dateVal) {
  if (!dateVal) return null;

  // Already a Date object
  if (dateVal instanceof Date) {
    if (isNaN(dateVal.getTime())) return null;
    return dateVal;
  }

  const str = dateVal.toString().trim();

  // dd.mm.yyyy or dd/mm/yyyy or dd-mm-yyyy
  let match = str.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})$/);
  if (match) {
    let day = parseInt(match[1], 10);
    let month = parseInt(match[2], 10) - 1;
    let year = parseInt(match[3], 10);
    if (year < 100) year += 2000;

    if (day >= 1 && day <= 31 && month >= 0 && month <= 11) {
      const d = new Date(year, month, day);
      if (d.getDate() === day && d.getMonth() === month) return d;
    }
  }

  // yyyy-mm-dd or yyyy/mm/dd
  match = str.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);
  if (match) {
    const year = parseInt(match[1], 10);
    const month = parseInt(match[2], 10) - 1;
    const day = parseInt(match[3], 10);

    if (day >= 1 && day <= 31 && month >= 0 && month <= 11) {
      const d = new Date(year, month, day);
      if (d.getDate() === day && d.getMonth() === month) return d;
    }
  }

  // Try native parsing as last resort
  const native = new Date(str);
  if (!isNaN(native.getTime())) return native;

  return null;
}

/**
 * استخراج الجزء الإنجليزي من القيمة ثنائية اللغة
 * مثال: "Revenue Accrual (استحقاق إيراد)" → "Revenue Accrual"
 * مثال: "Cash" → "Cash"
 */
function extractEnglishPart(value) {
  if (!value) return '';
  const str = value.toString().trim();

  // Remove Arabic/Turkish part in parentheses
  const match = str.match(/^([^(]+)/);
  if (match) {
    return match[1].trim();
  }
  return str;
}

/**
 * تطابق قيمة الإدخال مع قيم dropdown الموجودة
 * يقبل الاسم الإنجليزي فقط أو الاسم الكامل ثنائي اللغة
 */
function mapToDropdownValue(inputValue, dropdownList) {
  if (!inputValue) return '';
  const input = inputValue.toString().trim();

  // Exact match
  for (const item of dropdownList) {
    if (item === input) return item;
  }

  // Match by English part
  const inputEN = extractEnglishPart(input).toLowerCase();
  for (const item of dropdownList) {
    const itemEN = extractEnglishPart(item).toLowerCase();
    if (itemEN === inputEN) return item;
  }

  // Partial match (contains)
  for (const item of dropdownList) {
    if (item.toLowerCase().includes(inputEN) || inputEN.includes(extractEnglishPart(item).toLowerCase())) {
      return item;
    }
  }

  // Return the input as-is if no match found
  return input;
}

/**
 * تطابق اسم القطاع مع القيم الموجودة في Sector Profiles
 */
function mapToSectorValue(ss, inputValue) {
  if (!inputValue) return '';
  const input = inputValue.toString().trim().toLowerCase();

  const sectorsSheet = ss.getSheetByName('Sector Profiles');
  if (!sectorsSheet || sectorsSheet.getLastRow() < 2) return inputValue;

  const data = sectorsSheet.getRange(2, 1, sectorsSheet.getLastRow() - 1, 14).getValues();

  for (const row of data) {
    const nameEN = (row[1] || '').toString().trim();
    const nameAR = (row[2] || '').toString().trim();
    const nameTR = (row[3] || '').toString().trim();
    const status = row[13];

    if (status !== 'Active') continue;

    const fullValue = nameEN + ' (' + (nameAR || nameEN) + ')';

    // Exact or partial match
    if (nameEN.toLowerCase() === input ||
        nameAR.toLowerCase() === input ||
        nameTR.toLowerCase() === input ||
        fullValue.toLowerCase() === input ||
        input.includes(nameEN.toLowerCase()) ||
        nameEN.toLowerCase().includes(input)) {
      return fullValue;
    }
  }

  return inputValue;
}

/**
 * تطابق اسم الخزينة/البنك مع القيم الموجودة
 */
function mapToCashBankValue(ss, inputValue) {
  if (!inputValue) return '';
  const input = inputValue.toString().trim().toLowerCase();

  // Check Cash Boxes
  // Columns: Cash Code(1), Cash Name(2), Currency(3), Responsible(4), Location(5), Opening Balance(6), Opening Date(7), Status(8)
  const cashSheet = ss.getSheetByName('Cash Boxes');
  if (cashSheet && cashSheet.getLastRow() > 1) {
    const data = cashSheet.getRange(2, 2, cashSheet.getLastRow() - 1, 7).getValues();
    for (const row of data) {
      const name = (row[0] || '').toString().trim();      // Cash Name (col 2)
      const currency = (row[1] || '').toString().trim();   // Currency (col 3)
      const status = (row[6] || '').toString().trim();     // Status (col 8)

      if (name && status === 'Active' && (name.toLowerCase().includes(input) || input.includes(name.toLowerCase()))) {
        return '💰 ' + name + ' (' + currency + ')';
      }
    }
  }

  // Check Bank Accounts
  // Columns: Account Code(1), Account Name(2), Bank Name(3), Currency(4), IBAN(5), SWIFT(6), Holder(7), Branch(8), Opening Balance(9), Opening Date(10), Status(11)
  const bankSheet = ss.getSheetByName('Bank Accounts');
  if (bankSheet && bankSheet.getLastRow() > 1) {
    const data = bankSheet.getRange(2, 2, bankSheet.getLastRow() - 1, 10).getValues();
    for (const row of data) {
      const name = (row[0] || '').toString().trim();       // Account Name (col 2)
      const currency = (row[2] || '').toString().trim();   // Currency (col 4)
      const status = (row[9] || '').toString().trim();     // Status (col 11)

      if (name && status === 'Active' && (name.toLowerCase().includes(input) || input.includes(name.toLowerCase()))) {
        return '🏦 ' + name + ' (' + currency + ')';
      }
    }
  }

  return inputValue;
}

/**
 * البحث عن كود العميل من اسمه
 */
function findClientCode(ss, clientName) {
  if (!clientName) return '';
  const input = clientName.toString().trim().toLowerCase();

  const clientsSheet = ss.getSheetByName('Clients');
  if (!clientsSheet || clientsSheet.getLastRow() < 2) return '';

  const data = clientsSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const code = data[i][0];
    const nameEN = (data[i][1] || '').toString().trim().toLowerCase();
    const nameAR = (data[i][2] || '').toString().trim().toLowerCase();
    const nameTR = (data[i][3] || '').toString().trim().toLowerCase();

    if (nameEN === input || nameAR === input || nameTR === input ||
        nameEN.includes(input) || input.includes(nameEN)) {
      return code;
    }
  }

  return '';
}

// ==================== 7. CLEAR IMPORT SHEET ====================
/**
 * مسح بيانات شيت الاستيراد للاستخدام مرة أخرى
 */
function clearImportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const importSheet = ss.getSheetByName('Import Data');
  if (!importSheet) {
    ui.alert('❌ شيت Import Data غير موجود!');
    return;
  }

  const confirm = ui.alert(
    '🗑️ مسح بيانات الاستيراد\n\n' +
    'هل تريد مسح كل البيانات من شيت Import Data؟\n' +
    '(الهيدرات والتعليمات ستبقى)',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const lastRow = importSheet.getLastRow();
  if (lastRow > 2) {
    importSheet.getRange(3, 1, lastRow - 2, 18).clear();
    importSheet.getRange(3, 1, lastRow - 2, 18).setBackground(null);
  }

  importSheet.getRange(2, 20).setValue('Ready').setBackground('#c8e6c9');
  importSheet.getRange(3, 21).setValue(0);
  importSheet.getRange(4, 21).setValue(0);

  ui.alert('✅ تم مسح بيانات الاستيراد!');
}

/**
 * مسح بيانات شيت الأرصدة الافتتاحية
 */
function clearOpeningBalancesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const sheet = ss.getSheetByName('Import Opening Balances');
  if (!sheet) {
    ui.alert('❌ شيت Import Opening Balances غير موجود!');
    return;
  }

  const confirm = ui.alert(
    '🗑️ مسح بيانات الأرصدة الافتتاحية\n\n' +
    'هل تريد مسح كل البيانات؟',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(3, 1, lastRow - 2, 7).clear();
    sheet.getRange(3, 1, lastRow - 2, 7).setBackground(null);
  }

  sheet.getRange(2, 9).setValue('Ready').setBackground('#c8e6c9');

  ui.alert('✅ تم مسح البيانات!');
}

// Note: refreshAllData() is defined in 08-Reports.gs - no duplicate needed here

// ==================== END OF PART 11 ====================
