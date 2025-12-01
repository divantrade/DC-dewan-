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

function generateClientStatement(clientCode, clientName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const transSheet = ss.getSheetByName('Transactions');
  if (!transSheet || transSheet.getLastRow() < 2) {
    ui.alert('⚠️ No transactions found!');
    return;
  }
  
  const transData = transSheet.getDataRange().getValues();
  const headers = transData[0];
  
  // Find column indices
  const cols = {};
  headers.forEach((h, i) => cols[h] = i);
  
  // Filter transactions for this client where Show in Statement = Yes
  const clientTrans = [];
  
  for (let i = 1; i < transData.length; i++) {
    const code = transData[i][4]; // Client Code
    const name = transData[i][5]; // Client Name
    const showInStatement = transData[i][24]; // Column Y
    
    if ((code === clientCode || name === clientName) && 
        (!showInStatement || showInStatement.includes('Yes'))) {
      clientTrans.push({
        date: transData[i][1],
        movementType: transData[i][2],
        item: transData[i][6],
        description: transData[i][7],
        amount: transData[i][10],
        currency: transData[i][11],
        status: transData[i][18],
        invoiceNo: transData[i][17]
      });
    }
  }
  
  if (clientTrans.length === 0) {
    ui.alert('ℹ️ No statement items found for this client.\n\n(Only items with "Show in Statement = Yes" are included)');
    return;
  }
  
  // Calculate totals
  let totalRevenue = 0, totalPaid = 0;
  clientTrans.forEach(t => {
    const amt = parseFloat(t.amount) || 0;
    if (t.movementType && t.movementType.includes('Revenue')) {
      totalRevenue += amt;
    }
    if (t.status && t.status.includes('Paid')) {
      totalPaid += amt;
    }
  });
  
  // Show summary
  const summary = 
    '📄 Client Statement: ' + clientName + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Transactions: ' + clientTrans.length + '\n' +
    'Total Billed: ' + formatCurrency(totalRevenue, 'TRY') + '\n' +
    'Total Paid: ' + formatCurrency(totalPaid, 'TRY') + '\n' +
    'Balance Due: ' + formatCurrency(totalRevenue - totalPaid, 'TRY') + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'Export to sheet?';
  
  const exportConfirm = ui.alert(summary, ui.ButtonSet.YES_NO);
  
  if (exportConfirm === ui.Button.YES) {
    exportClientStatement(clientCode, clientName, clientTrans);
  }
}

function exportClientStatement(clientCode, clientName, transactions) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Statement - ' + clientName.substring(0, 20);
  
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet(sheetName);
  sheet.setTabColor('#4caf50');
  
  // Header
  sheet.getRange('A1:F1').merge()
    .setValue('Client Statement: ' + clientName)
    .setFontSize(14).setFontWeight('bold').setBackground('#4caf50').setFontColor('#ffffff');
  
  sheet.getRange('A2').setValue('Generated: ' + formatDate(new Date(), 'yyyy-MM-dd HH:mm'));
  sheet.getRange('A3').setValue('Client Code: ' + clientCode);
  
  // Table headers
  const headers = ['Date', 'Type', 'Item', 'Description', 'Amount', 'Status'];
  sheet.getRange(5, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold').setBackground('#c8e6c9');
  
  // Data
  const data = transactions.map(t => [
    formatDate(t.date, 'yyyy-MM-dd'),
    t.movementType,
    t.item,
    t.description,
    t.amount,
    t.status
  ]);
  
  if (data.length > 0) {
    sheet.getRange(6, 1, data.length, headers.length).setValues(data);
    sheet.getRange(6, 5, data.length, 1).setNumberFormat('#,##0.00');
  }
  
  // Column widths
  const widths = [100, 150, 150, 200, 100, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  ss.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert('✅ Statement exported to sheet: ' + sheetName);
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

function generateClientProfitability(clientCode, clientName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const transSheet = ss.getSheetByName('Transactions');
  if (!transSheet || transSheet.getLastRow() < 2) {
    ui.alert('⚠️ No transactions found!');
    return;
  }
  
  const transData = transSheet.getDataRange().getValues();
  
  // All transactions for this client (including hidden ones)
  let totalRevenue = 0;
  let totalDirectExpenses = 0;
  let transCount = 0;
  
  for (let i = 1; i < transData.length; i++) {
    const code = transData[i][4];
    const name = transData[i][5];
    const movementType = transData[i][2];
    const amount = parseFloat(transData[i][13]) || 0; // Amount TRY
    
    if (code === clientCode || name === clientName) {
      transCount++;
      
      if (movementType && movementType.includes('Revenue')) {
        totalRevenue += amount;
      }
      if (movementType && (movementType.includes('Expense') || movementType.includes('مصروف'))) {
        totalDirectExpenses += amount;
      }
    }
  }
  
  const grossProfit = totalRevenue - totalDirectExpenses;
  const profitMargin = totalRevenue > 0 ? (grossProfit / totalRevenue * 100).toFixed(1) : 0;
  
  const report = 
    '💹 Profitability Report: ' + clientName + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '📊 FINANCIAL SUMMARY\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Total Revenue: ' + formatCurrency(totalRevenue, 'TRY') + '\n' +
    'Direct Expenses: ' + formatCurrency(totalDirectExpenses, 'TRY') + '\n' +
    '───────────────────────────\n' +
    'Gross Profit: ' + formatCurrency(grossProfit, 'TRY') + '\n' +
    'Profit Margin: ' + profitMargin + '%\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Total Transactions: ' + transCount + '\n\n' +
    '💡 Note: This includes ALL transactions\n(even hidden from statement)';
  
  ui.alert(report);
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

// ==================== END OF PART 8 ====================
