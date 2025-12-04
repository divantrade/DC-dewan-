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
  sheet.getRange(2, 2, lastRow, 1).setNumberFormat('dd.mm.yy');
  sheet.getRange(2, 5, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 11, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 12, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 13, lastRow, 1).setNumberFormat('dd.mm.yy');
  sheet.getRange(2, 15, lastRow, 1).setNumberFormat('dd.mm.yy HH:mm');

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
  sheet.getRange(2, 3, lastRow, 1).setNumberFormat('dd.mm.yy');
  sheet.getRange(2, 5, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 11, lastRow, 1).setNumberFormat('dd.mm.yy HH:mm');

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
