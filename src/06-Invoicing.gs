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
  
  // Set column widths
  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 50);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  
  // Company Header
  sheet.getRange('A1:E1').merge()
    .setValue(getSettingValue('Company Name (EN)') || 'Dewan Consulting')
    .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#1565c0').setFontColor('#ffffff');
  
  sheet.getRange('A2:E2').merge()
    .setValue(getSettingValue('Company Name (AR)') || 'ديوان للاستشارات')
    .setFontSize(14).setHorizontalAlignment('center')
    .setBackground('#1976d2').setFontColor('#ffffff');
  
  sheet.getRange('A3:E3').merge()
    .setValue(getSettingValue('Company Address') || '')
    .setFontSize(10).setHorizontalAlignment('center');
  
  // Invoice Title
  sheet.getRange('A5:E5').merge()
    .setValue('INVOICE / فاتورة')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#e3f2fd');
  
  // Invoice Details Section
  const detailLabels = [
    ['Invoice No:', '', '', 'Date:', ''],
    ['Client:', '', '', '', ''],
    ['Tax Number:', '', '', '', ''],
    ['Address:', '', '', '', ''],
    ['Period:', '', '', '', '']
  ];
  sheet.getRange('A7:E11').setValues(detailLabels);
  sheet.getRange('A7:A11').setFontWeight('bold');
  sheet.getRange('D7').setFontWeight('bold');
  
  // Items Table Header
  sheet.getRange('A13:E13')
    .setValues([['#', 'Description / الوصف', 'Qty', 'Unit Price', 'Total']])
    .setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Totals Section
  sheet.getRange('D24').setValue('Subtotal:').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('D25').setValue('VAT (0%):').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('D26').setValue('TOTAL:').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('right');
  sheet.getRange('E24:E26').setNumberFormat('#,##0.00').setHorizontalAlignment('right');
  sheet.getRange('E26').setFontWeight('bold').setFontSize(12).setBackground('#e3f2fd');
  
  // Bank Details
  sheet.getRange('A28:E28').merge()
    .setValue('Bank Details / التفاصيل البنكية')
    .setFontWeight('bold').setBackground('#f5f5f5');
  
  const bankDetails = [
    ['Bank:', getSettingValue('Bank Name') || 'Kuveyt Türk', '', '', ''],
    ['IBAN (TRY):', getSettingValue('IBAN TRY') || '', '', '', ''],
    ['IBAN (USD):', getSettingValue('IBAN USD') || '', '', '', ''],
    ['SWIFT:', getSettingValue('SWIFT Code') || 'KTEFTRIS', '', '', '']
  ];
  sheet.getRange('A29:E32').setValues(bankDetails);
  sheet.getRange('A29:A32').setFontWeight('bold');
  
  // Footer
  sheet.getRange('A34:E34').merge()
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
    description: (d.item || '') + (d.description ? ' - ' + d.description : ''),
    qty: 1,
    unitPrice: d.amount,
    total: d.amount
  }));
  
  fillInvoiceTemplate(ss, {
    invoiceNo: invoiceNo,
    invoiceDate: invoiceDate,
    clientName: firstClientName || (clientData ? clientData.nameEN : ''),
    clientNameAR: clientData ? clientData.nameAR : '',
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
    taxNumber: clientData.taxNumber || '',
    address: clientData.address || '',
    period: period,
    items: [{
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
  
  // Clear previous data
  sheet.getRange('B7:E11').clearContent();
  sheet.getRange('A14:E22').clearContent();
  sheet.getRange('E24:E26').clearContent();
  
  // Invoice details
  sheet.getRange('B7').setValue(data.invoiceNo);
  sheet.getRange('E7').setValue(formatDate(data.invoiceDate, 'yyyy-MM-dd'));
  sheet.getRange('B8').setValue(data.clientName + (data.clientNameAR ? ' / ' + data.clientNameAR : ''));
  sheet.getRange('B9').setValue(data.taxNumber || '');
  sheet.getRange('B10').setValue(data.address || '');
  sheet.getRange('B11').setValue(data.period || '');
  
  // Items - dynamic numbering
  if (data.items && data.items.length > 0) {
    data.items.forEach((item, i) => {
      const row = 14 + i;
      if (row <= 22) {
        sheet.getRange(row, 1).setValue(i + 1).setHorizontalAlignment('center');
        sheet.getRange(row, 2).setValue(item.description);
        sheet.getRange(row, 3).setValue(item.qty || 1).setHorizontalAlignment('center');
        sheet.getRange(row, 4).setValue(item.unitPrice).setNumberFormat('#,##0.00');
        sheet.getRange(row, 5).setValue(item.total).setNumberFormat('#,##0.00');
      }
    });
  }
  
  // Totals
  const currencySymbol = data.currency === 'TRY' ? '₺' : (data.currency === 'USD' ? '$' : (data.currency === 'EUR' ? '€' : data.currency));
  sheet.getRange('E24').setValue(data.subtotal).setNumberFormat('#,##0.00');
  sheet.getRange('D25').setValue('VAT (' + (data.vatRate || 0) + '%):').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('E25').setValue(data.vat || 0).setNumberFormat('#,##0.00');
  sheet.getRange('E26').setValue(data.total).setNumberFormat('#,##0.00 "' + currencySymbol + '"');
  
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
  
  sheet.getRange('B7:E11').clearContent();
  sheet.getRange('A14:E22').clearContent();
  sheet.getRange('E24:E26').clearContent();
  
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
