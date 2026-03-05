// ╔════════════════════════════════════════════════════════════════════════════╗
// ║                    DC CONSULTING ACCOUNTING SYSTEM v3.1                     ║
// ║                              Part 6 of 9                                    ║
// ║              Invoice System (with Sector Profiles support)                 ║
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
    'Sector',            // E - Accounting/Consulting/etc.
    'Service',           // F
    'Period',            // G
    'Amount',            // H
    'Currency',          // I
    'Status',            // J
    'PDF Link',          // K
    'Send Email',        // L - Yes/No
    'Email Status',      // M - Pending/Sent/Failed
    'Email Sent Date',   // N
    'Trans. Code',       // O
    'Notes',             // P
    'Created Date'       // Q
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#6a1b9a')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  const widths = [100, 100, 90, 180, 120, 150, 100, 100, 70, 90, 250, 80, 100, 100, 120, 200, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  const lastRow = 500;
  
  // Sector validation (column E)
  const sectorRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Accounting', 'Consulting', 'Logistics', 'Trading', 'Inspection', 'Tourism', 'Other'], true).build();
  sheet.getRange(2, 5, lastRow, 1).setDataValidation(sectorRule);

  // Status validation (column J)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Draft', 'Issued', 'Sent', 'Paid', 'Cancelled'], true).build();
  sheet.getRange(2, 10, lastRow, 1).setDataValidation(statusRule);

  // Send Email validation (column L)
  const sendRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true).build();
  sheet.getRange(2, 12, lastRow, 1).setDataValidation(sendRule);

  // Email Status validation (column M)
  const emailStatusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Sent', 'Failed'], true).build();
  sheet.getRange(2, 13, lastRow, 1).setDataValidation(emailStatusRule);

  // Number formats
  sheet.getRange(2, 2, lastRow, 1).setNumberFormat('dd.mm.yyyy');
  sheet.getRange(2, 8, lastRow, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 14, lastRow, 1).setNumberFormat('dd.mm.yyyy HH:mm');
  sheet.getRange(2, 17, lastRow, 1).setNumberFormat('dd.mm.yyyy HH:mm');

  // Conditional formatting
  const statusRange = sheet.getRange(2, 10, lastRow, 1);
  const emailRange = sheet.getRange(2, 13, lastRow, 1);
  
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

  // Get logo URL from settings (default profile)
  const logoUrl = resolveLogoUrl(getSettingValue('Company Logo URL') || '');

  let currentRow = 1;

  // Row 1: Logo (centered) - if provided
  if (logoUrl) {
    sheet.getRange('A1:F1').merge();
    sheet.getRange('A1').setFormula('=IMAGE("' + logoUrl + '", 1)');
    sheet.setRowHeight(1, 90);
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

  // Row 3: City / Country
  sheet.getRange('A' + (detailsStartRow + 2) + ':B' + (detailsStartRow + 2)).merge().setValue('City / Country:').setFontWeight('bold');
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

// ==================== 2.5 UPDATE LOGO IN INVOICE TEMPLATE ====================
/**
 * تحديث/إضافة اللوجو في نموذج الفاتورة الموجود
 */
function updateInvoiceLogo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Invoice Template');

  if (!sheet) {
    ui.alert('❌ Invoice Template not found!\n\nRun Setup System first.');
    return;
  }

  // Get logo URL from settings
  const companyLogo = getSettingValue('Company Logo URL') || '';

  if (!companyLogo || companyLogo.trim() === '') {
    ui.alert('❌ No logo URL found in Settings!\n\nAdd "Company Logo URL" in Settings sheet first.');
    return;
  }

  const logoUrl = resolveLogoUrl(companyLogo);

  // Insert row at top for logo
  sheet.insertRowBefore(1);

  // Add logo
  sheet.getRange('A1:F1').merge();
  sheet.getRange('A1').setFormula('=IMAGE("' + logoUrl + '", 1)');
  sheet.setRowHeight(1, 90);
  sheet.getRange('A1').setHorizontalAlignment('center').setVerticalAlignment('middle');

  ui.alert('✅ Logo added successfully!\n\nتم إضافة اللوجو بنجاح!');
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
  let firstSector = null;
  let totalAmount = 0;
  let currency = 'TRY';

  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const rowData = transSheet.getRange(row, 1, 1, 26).getValues()[0];

    const transCode = rowData[0];
    const transDate = rowData[1];
    const rowSector = rowData[2] || '';  // Column C - Sector
    const clientCode = rowData[5];
    const clientName = rowData[6];
    const item = rowData[7];
    const description = rowData[8];
    const amount = rowData[11] || 0;
    const rowCurrency = rowData[12] || 'TRY';

    if (!amount || amount === 0) continue;

    if (firstClientCode === null) {
      firstClientCode = clientCode;
      firstClientName = clientName;
      firstSector = rowSector;
      currency = rowCurrency;
    } else if (clientCode !== firstClientCode && clientName !== firstClientName) {
      ui.alert('⚠️ All selected rows must be for the SAME client!\n\nكل الصفوف المختارة يجب أن تكون لنفس العميل');
      return;
    }

    if (rowCurrency !== currency) {
      ui.alert('⚠️ All selected rows must have the SAME currency!\n\nكل الصفوف يجب أن تكون بنفس العملة');
      return;
    }

    if (rowSector && firstSector && rowSector !== firstSector) {
      ui.alert('⚠️ All selected rows must be for the SAME sector!\n\nكل الصفوف المختارة يجب أن تكون لنفس القطاع');
      return;
    }

    // Use whichever row has a sector value
    if (rowSector && !firstSector) {
      firstSector = rowSector;
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

  // Use sector from transaction Column C (extract EN name before parenthesis)
  // Dropdown format: "Accounting (محاسبة)" → extract "Accounting"
  let clientActivity = '';
  if (firstSector) {
    const parenIndex = firstSector.indexOf(' (');
    clientActivity = parenIndex > 0 ? firstSector.substring(0, parenIndex) : firstSector;
  } else {
    // Fallback: get sector from Client Sector sheet
    clientActivity = firstClientCode ? getClientPrimaryActivity(firstClientCode) : '';
  }

  const itemsList = selectedData.map((d, i) =>
    (i + 1) + '. ' + (d.item || d.description || 'Item') + ': ' + formatCurrency(d.amount, currency)
  ).join('\n');
  
  const sectorDisplay = clientActivity ? clientActivity : 'Default (افتراضي)';
  const confirm = ui.alert(
    '📄 Generate Invoice (إنشاء فاتورة)\n\n' +
    'Client: ' + (firstClientName || firstClientCode) + '\n' +
    'Sector: ' + sectorDisplay + '\n' +
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
    activity: clientActivity,  // Sector profile for branding
    clientName: firstClientName || (clientData ? clientData.nameEN : ''),
    clientNameAR: clientData ? clientData.nameAR : '',
    cityCountry: clientData ? ((clientData.city || '') + (clientData.city && clientData.country ? ', ' : '') + (clientData.country || '')) : '',
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
    activity: clientActivity,  // Track sector in log
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
    transSheet.getRange(d.row, 19).setValue(invoiceNo);
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
  
  // Get client activities for auto-detection
  const clientActivities = getClientSectorList(clientCode);
  let selectedActivity = '';

  if (clientActivities.length === 1) {
    selectedActivity = clientActivities[0].activity;
  } else if (clientActivities.length > 1) {
    const actList = clientActivities.map(a => a.activity).join(', ');
    const actResponse = ui.prompt(
      '📄 Custom Invoice - Select Sector\n\n' +
      'Client has multiple sectors: ' + actList,
      'Enter sector name (اختر القطاع):\n\n' + actList,
      ui.ButtonSet.OK_CANCEL
    );
    if (actResponse.getSelectedButton() !== ui.Button.OK) return;
    selectedActivity = actResponse.getResponseText().trim();
  }

  const activityProfile = getSectorProfile(selectedActivity);

  const clientConfirm = ui.alert(
    '✅ Client Found (تم العثور على العميل)\n\n' +
    'Code: ' + clientCode + '\n' +
    'Name (EN): ' + clientData.nameEN + '\n' +
    'Name (AR): ' + (clientData.nameAR || '-') + '\n' +
    'Tax Number: ' + (clientData.taxNumber || '-') + '\n' +
    (selectedActivity ? 'Sector: ' + selectedActivity + ' (' + activityProfile.companyNameEN + ')' : '') + '\n\n' +
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
    activity: selectedActivity,  // Sector profile for branding
    clientName: clientData.nameEN,
    clientNameAR: clientData.nameAR || '',
    cityCountry: ((clientData.city || '') + (clientData.city && clientData.country ? ', ' : '') + (clientData.country || '')),
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
    activity: selectedActivity,  // Track sector in log
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
    if (selectedActivity) {
      transSheet.getRange(lastRow, 3).setValue(selectedActivity);  // Column C - Sector
    }
    transSheet.getRange(lastRow, 4).setValue('Revenue Accrual (استحقاق إيراد)');
    transSheet.getRange(lastRow, 5).setValue('Service Revenue (إيرادات خدمات)');
    transSheet.getRange(lastRow, 6).setValue(clientCode);
    transSheet.getRange(lastRow, 7).setValue(clientData.nameEN);
    transSheet.getRange(lastRow, 9).setValue(description);
    transSheet.getRange(lastRow, 10).setValue(clientData.nameEN + ' (' + (clientData.nameAR || clientData.nameEN) + ')');
    transSheet.getRange(lastRow, 11).setValue('Client (عميل)');
    transSheet.getRange(lastRow, 12).setValue(totalAmount);
    transSheet.getRange(lastRow, 13).setValue(currency);
    transSheet.getRange(lastRow, 14).setValue(1);
    transSheet.getRange(lastRow, 15).setValue(totalAmount);
    transSheet.getRange(lastRow, 16).setValue('Accrual (استحقاق)');
    transSheet.getRange(lastRow, 19).setValue(invoiceNo);
    transSheet.getRange(lastRow, 20).setValue('Pending (معلق)');
    transSheet.getRange(lastRow, 26).setValue('Yes (نعم)');
    
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
        transSheet.getRange(transRow, 24).setValue('PDF: ' + pdfResult.url);
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

  // Get clients with monthly fees from Client Sector sheet
  const monthlyActivities = getClientsWithMonthlyFees();

  if (monthlyActivities.length === 0) {
    ui.alert('⚠️ No clients with monthly fees found!\n\nAdd monthly fee activities in "Client Sector" sheet.');
    return;
  }

  const clientsList = monthlyActivities.map(a =>
    '• ' + a.clientName + ' [' + a.activity + ']: ' + formatCurrency(a.monthlyFee, a.currency)
  ).join('\n');

  const confirm = ui.alert(
    '📋 Generate All Monthly Invoices\n\n' +
    'This will create invoices for ' + monthlyActivities.length + ' sector subscriptions:\n\n' +
    clientsList + '\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const invoiceDate = new Date();
  const period = Utilities.formatDate(invoiceDate, Session.getScriptTimeZone(), 'MMMM yyyy');
  let generated = 0;
  let pdfSaved = 0;

  monthlyActivities.forEach(act => {
    const invoiceNo = getNextInvoiceNumber();
    const clientData = getClientData(act.clientCode);

    // Determine service label based on activity
    const serviceLabel = act.activity === 'Accounting'
      ? 'Monthly Accounting (محاسبة شهرية)'
      : act.activity === 'Consulting'
        ? 'Monthly Consulting (استشارات شهرية)'
        : 'Monthly ' + act.activity;

    // Fill template with sector-specific branding
    fillInvoiceTemplate(ss, {
      invoiceNo: invoiceNo,
      invoiceDate: invoiceDate,
      activity: act.activity,  // Sector profile for branding
      clientName: act.clientName,
      clientNameAR: clientData ? clientData.nameAR : '',
      cityCountry: clientData ? ((clientData.city || '') + (clientData.city && clientData.country ? ', ' : '') + (clientData.country || '')) : '',
      taxNumber: clientData ? clientData.taxNumber : '',
      address: clientData ? clientData.address : '',
      period: period,
      items: [{
        item: act.activity,
        description: serviceLabel,
        qty: 1,
        unitPrice: act.monthlyFee,
        total: act.monthlyFee
      }],
      currency: act.currency,
      subtotal: act.monthlyFee,
      vat: 0,
      vatRate: 0,
      total: act.monthlyFee
    });

    // Save PDF
    let pdfUrl = '';
    if (clientData && clientData.folderId) {
      try {
        const pdfResult = createInvoicePDF(invoiceNo, clientData.folderId);
        pdfUrl = pdfResult.url;
        pdfSaved++;
      } catch (e) {
        console.log('PDF error for ' + act.clientCode + ': ' + e.message);
      }
    }

    // Log invoice with sector
    logInvoice({
      invoiceNo: invoiceNo,
      invoiceDate: invoiceDate,
      clientCode: act.clientCode,
      clientName: act.clientName,
      activity: act.activity,  // Track sector in log
      service: serviceLabel,
      period: period,
      amount: act.monthlyFee,
      currency: act.currency,
      status: 'Issued',
      pdfLink: pdfUrl,
      sendEmail: 'Yes',
      emailStatus: 'Pending',
      transCode: ''
    });

    // Record transaction with sector
    recordInvoiceTransaction(invoiceNo, act.clientCode, act.clientName, act.monthlyFee, act.currency, serviceLabel, act.activity);

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
/**
 * Fill invoice template with data and sector-specific branding
 * @param {Spreadsheet} ss
 * @param {object} data - Invoice data including items, amounts, client info
 * data.activity - Sector name (e.g. 'Accounting') for per-sector branding
 */
function fillInvoiceTemplate(ss, data) {
  let sheet = ss.getSheetByName('Invoice Template');
  if (!sheet) {
    sheet = createInvoiceTemplateSheet(ss);
  }

  // Get sector profile for branding (falls back to Settings if no sector)
  const profile = getSectorProfile(data.activity || '');

  // === Update branding: Logo, Company Name, Bank Details ===
  let logoUrl = resolveLogoUrl(profile.logoUrl);
  const hasLogo = logoUrl !== '';

  // Row 1: Logo (mode 1 = fit to cell, maintains aspect ratio)
  if (hasLogo) {
    sheet.getRange('A1').setFormula('=IMAGE("' + logoUrl + '", 1)');
    sheet.setRowHeight(1, 90);
  } else {
    sheet.getRange('A1').clearContent();
    sheet.setRowHeight(1, 20);
  }

  const rowOffset = hasLogo ? 1 : 0;

  // Row 2 (or 1): Company Name EN
  sheet.getRange('A' + (1 + rowOffset) + ':F' + (1 + rowOffset)).getMergedRanges().length ||
    sheet.getRange('A' + (1 + rowOffset) + ':F' + (1 + rowOffset));
  sheet.getRange('A' + (1 + rowOffset)).setValue(profile.companyNameEN);

  // Row 3 (or 2): Company Name AR
  sheet.getRange('A' + (2 + rowOffset)).setValue(profile.companyNameAR);

  // Row 4 (or 3): Address
  sheet.getRange('A' + (3 + rowOffset)).setValue(profile.address);

  // Details start row (7 without logo, 8 with logo)
  const detailsStartRow = 7 + rowOffset;
  // Items start row (14 without logo, 15 with logo)
  const itemsStartRow = 14 + rowOffset;
  // Totals row (26 without logo, 27 with logo)
  const totalsStartRow = 26 + rowOffset;

  // Clear previous data
  sheet.getRange('C' + detailsStartRow + ':F' + (detailsStartRow + 4)).clearContent();
  sheet.getRange('A' + itemsStartRow + ':F' + (itemsStartRow + 9)).clearContent();
  sheet.getRange('F' + totalsStartRow + ':F' + (totalsStartRow + 2)).clearContent();

  // Invoice details
  sheet.getRange('C' + detailsStartRow).setValue(data.invoiceNo);
  sheet.getRange('F' + detailsStartRow).setValue(formatDate(data.invoiceDate, 'yyyy-MM-dd'));
  sheet.getRange('C' + (detailsStartRow + 1)).setValue(data.clientName + (data.clientNameAR ? ' / ' + data.clientNameAR : ''));
  sheet.getRange('C' + (detailsStartRow + 2)).setValue(data.cityCountry || '');
  sheet.getRange('C' + (detailsStartRow + 3)).setValue(data.taxNumber || '');
  sheet.getRange('C' + (detailsStartRow + 4)).setValue(data.address || '');

  // Items - 6 columns: #, Item, Description, Qty, Unit Price, Total
  if (data.items && data.items.length > 0) {
    data.items.forEach((item, i) => {
      const row = itemsStartRow + i;
      if (row < itemsStartRow + 10) {
        sheet.getRange(row, 1).setValue(i + 1).setHorizontalAlignment('center');
        sheet.getRange(row, 2).setValue(item.item || '');
        sheet.getRange(row, 3).setValue(item.description || '');
        sheet.getRange(row, 4).setValue(item.qty || 1).setHorizontalAlignment('center');
        sheet.getRange(row, 5).setValue(item.unitPrice).setNumberFormat('#,##0.00');
        sheet.getRange(row, 6).setValue(item.total).setNumberFormat('#,##0.00');
      }
    });
  }

  // Totals
  const currencySymbol = data.currency === 'TRY' ? '₺' : (data.currency === 'USD' ? '$' : (data.currency === 'EUR' ? '€' : data.currency));
  sheet.getRange('F' + totalsStartRow).setValue(data.subtotal).setNumberFormat('#,##0.00');
  sheet.getRange('E' + (totalsStartRow + 1)).setValue('VAT (' + (data.vatRate || 0) + '%):').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('F' + (totalsStartRow + 1)).setValue(data.vat || 0).setNumberFormat('#,##0.00');
  sheet.getRange('F' + (totalsStartRow + 2)).setValue(data.total).setNumberFormat('#,##0.00 "' + currencySymbol + '"');

  // === Update Bank Details with sector profile ===
  const bankRow = totalsStartRow + 4;
  sheet.getRange('C' + (bankRow + 1)).setValue(profile.bankName);
  sheet.getRange('C' + (bankRow + 2)).setValue(profile.ibanTRY);
  sheet.getRange('C' + (bankRow + 3)).setValue(profile.ibanUSD);
  sheet.getRange('C' + (bankRow + 4)).setValue(profile.swiftCode);

  // === Add website to footer if available ===
  const footerRow = bankRow + 6;
  let footerText = 'Thank you for your business! / شكراً لتعاملكم معنا';
  if (profile.website) {
    footerText += '\n' + profile.website;
  }
  sheet.getRange('A' + footerRow).setValue(footerText);

  return sheet;
}

/**
 * Resolve logo URL - convert Google Drive sharing links to direct image URLs
 */
function resolveLogoUrl(rawUrl) {
  if (!rawUrl || rawUrl.trim() === '') return '';

  let logoUrl = rawUrl.trim();

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

  return logoUrl;
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
  sheet.getRange(lastRow, 5).setValue(data.activity || '');       // Sector
  sheet.getRange(lastRow, 6).setValue(data.service);
  sheet.getRange(lastRow, 7).setValue(data.period);
  sheet.getRange(lastRow, 8).setValue(data.amount);
  sheet.getRange(lastRow, 9).setValue(data.currency);
  sheet.getRange(lastRow, 10).setValue(data.status || 'Issued');
  sheet.getRange(lastRow, 11).setValue(data.pdfLink || '');
  sheet.getRange(lastRow, 12).setValue(data.sendEmail || 'Yes');
  sheet.getRange(lastRow, 13).setValue(data.emailStatus || 'Pending');
  sheet.getRange(lastRow, 15).setValue(data.transCode || '');
  sheet.getRange(lastRow, 16).setValue(data.notes || '');
  sheet.getRange(lastRow, 17).setValue(new Date());
  
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
      logSheet.getRange(i + 1, 11).setValue(pdfUrl);   // PDF Link (column K)
      logSheet.getRange(i + 1, 16).setValue('PDF saved to client folder');  // Notes (column P)
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
function recordInvoiceTransaction(invoiceNo, clientCode, clientName, amount, currency, item, sectorName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName('Transactions');

  if (!transSheet) return null;

  const lastRow = transSheet.getLastRow() + 1;

  transSheet.getRange(lastRow, 1).setValue(lastRow - 1);
  transSheet.getRange(lastRow, 2).setValue(new Date());
  if (sectorName) {
    transSheet.getRange(lastRow, 3).setValue(sectorName);  // Column C - Sector
  }
  transSheet.getRange(lastRow, 4).setValue('Revenue Accrual (استحقاق إيراد)');
  transSheet.getRange(lastRow, 5).setValue('Service Revenue (إيرادات خدمات)');
  transSheet.getRange(lastRow, 6).setValue(clientCode);
  transSheet.getRange(lastRow, 7).setValue(clientName);
  transSheet.getRange(lastRow, 9).setValue(item);
  transSheet.getRange(lastRow, 11).setValue('Client (عميل)');
  transSheet.getRange(lastRow, 12).setValue(amount);
  transSheet.getRange(lastRow, 13).setValue(currency);
  transSheet.getRange(lastRow, 14).setValue(1);
  transSheet.getRange(lastRow, 16).setValue('Accrual (استحقاق)');
  transSheet.getRange(lastRow, 19).setValue(invoiceNo);
  transSheet.getRange(lastRow, 20).setValue('Pending (معلق)');
  transSheet.getRange(lastRow, 26).setValue('Yes (نعم)');

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
