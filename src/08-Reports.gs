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
      const movementType = transData[i][3]; // Movement Type
      const amount = parseFloat(transData[i][14]) || 0; // Amount TRY
      
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
      const status = transData[i][19];
      const dueDate = transData[i][20];
      const amount = parseFloat(transData[i][14]) || 0;
      
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
    const code = transData[i][5]; // Client Code
    const name = transData[i][6]; // Client Name
    const showInStatement = transData[i][25]; // Column Z

    if ((code === clientCode || name === clientName) &&
        (!showInStatement || showInStatement.includes('Yes'))) {

      const movementType = transData[i][3] || '';
      const amount = parseFloat(transData[i][11]) || 0;
      const item = transData[i][7] || '';
      const description = transData[i][8] || '';

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

  // Convert Google Drive link to direct image URL if needed
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

  let currentHeaderRow = 1;

  // Row 1: Logo (centered) - if provided
  if (logoUrl) {
    sheet.getRange('A1:F1').merge();
    sheet.getRange('A1').setFormula('=IMAGE("' + logoUrl + '", 1)');
    sheet.setRowHeight(1, 70);
    sheet.getRange('A1').setHorizontalAlignment('center').setVerticalAlignment('middle');
    currentHeaderRow = 2;
  }

  // Company Name EN
  sheet.getRange('A' + currentHeaderRow + ':F' + currentHeaderRow).merge()
    .setValue(companyNameEN)
    .setFontSize(22).setFontWeight('bold').setFontColor('#1565c0')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(currentHeaderRow, 40);
  currentHeaderRow++;

  // Company Name AR
  sheet.getRange('A' + currentHeaderRow + ':F' + currentHeaderRow).merge()
    .setValue(companyNameAR)
    .setFontSize(16).setFontWeight('bold').setFontColor('#424242')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(currentHeaderRow, 30);
  currentHeaderRow++;

  // Address
  sheet.getRange('A' + currentHeaderRow + ':F' + currentHeaderRow).merge()
    .setValue('📍 ' + companyAddress)
    .setFontSize(10).setFontColor('#616161')
    .setHorizontalAlignment('center');
  currentHeaderRow++;

  // Phone & Email
  sheet.getRange('A' + currentHeaderRow + ':F' + currentHeaderRow).merge()
    .setValue('📞 ' + companyPhone + '  |  ✉️ ' + companyEmail)
    .setFontSize(10).setFontColor('#616161')
    .setHorizontalAlignment('center');
  currentHeaderRow++;

  // Decorative line
  sheet.getRange('A' + currentHeaderRow + ':F' + currentHeaderRow).merge()
    .setBackground('#1565c0');
  sheet.setRowHeight(currentHeaderRow, 4);
  currentHeaderRow++;

  // Empty spacer
  sheet.setRowHeight(currentHeaderRow, 15);
  currentHeaderRow++;

  // ═══════════════════════════════════════════════════════════════════════════
  // STATEMENT TITLE - عنوان الكشف
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A' + currentHeaderRow + ':F' + currentHeaderRow).merge()
    .setValue('كشف حساب العميل  |  STATEMENT OF ACCOUNT')
    .setFontSize(14).setFontWeight('bold').setBackground('#e3f2fd').setFontColor('#1565c0')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(currentHeaderRow, 35);
  currentHeaderRow++;

  // Empty spacer
  sheet.setRowHeight(currentHeaderRow, 10);
  currentHeaderRow++;

  // ═══════════════════════════════════════════════════════════════════════════
  // CLIENT INFO SECTION - معلومات العميل
  // ═══════════════════════════════════════════════════════════════════════════

  // Client Name
  sheet.getRange('B' + currentHeaderRow).setValue('Client Name:').setFontWeight('bold').setFontColor('#424242').setHorizontalAlignment('right');
  sheet.getRange('C' + currentHeaderRow + ':F' + currentHeaderRow).merge().setValue(clientName).setFontColor('#1565c0').setFontWeight('bold');
  currentHeaderRow++;

  // Client Code
  sheet.getRange('B' + currentHeaderRow).setValue('Client Code:').setFontWeight('bold').setFontColor('#424242').setHorizontalAlignment('right');
  sheet.getRange('C' + currentHeaderRow + ':F' + currentHeaderRow).merge().setValue(clientCode).setFontColor('#1565c0');
  currentHeaderRow++;

  // Issue Date
  sheet.getRange('B' + currentHeaderRow).setValue('Issue Date:').setFontWeight('bold').setFontColor('#424242').setHorizontalAlignment('right');
  sheet.getRange('C' + currentHeaderRow + ':F' + currentHeaderRow).merge().setValue(formatDate(new Date(), 'yyyy-MM-dd')).setFontColor('#1565c0');
  currentHeaderRow++;

  // Decorative line
  sheet.getRange('A' + currentHeaderRow + ':F' + currentHeaderRow).merge().setBackground('#e0e0e0');
  sheet.setRowHeight(currentHeaderRow, 2);
  currentHeaderRow++;

  // Empty spacer
  sheet.setRowHeight(currentHeaderRow, 10);
  currentHeaderRow++;

  // ═══════════════════════════════════════════════════════════════════════════
  // TABLE SECTION - جدول الحركات
  // ═══════════════════════════════════════════════════════════════════════════

  const tableHeaderRow = currentHeaderRow;

  // Table headers
  const headers = ['#', 'التاريخ\nDate', 'الوصف\nDescription', 'له (دائن)\nCredit', 'عليه (مدين)\nDebit', 'الرصيد\nBalance'];
  sheet.getRange(tableHeaderRow, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1565c0')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(tableHeaderRow, 40);

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

  const dataStartRow = tableHeaderRow + 1;
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
  sheet.setFrozenRows(tableHeaderRow);

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
    const code = transData[i][5];
    const name = transData[i][6];
    const movementType = transData[i][3] || '';
    const item = transData[i][7] || '';
    const description = transData[i][8] || '';
    const amount = parseFloat(transData[i][14]) || 0; // Amount TRY
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
    const status = data[i][19];
    const dueDate = data[i][20];
    const clientName = data[i][6];
    const amount = data[i][11];
    const currency = data[i][12];
    const invoiceNo = data[i][18];
    
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

// ==================== END OF PART 8 ====================
