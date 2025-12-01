// ╔════════════════════════════════════════════════════════════════════════════╗
// ║                    DC CONSULTING ACCOUNTING SYSTEM v3.0                     ║
// ║                              Part 7 of 9                                    ║
// ║                    Email System + Triggers + Scheduling                     ║
// ╚════════════════════════════════════════════════════════════════════════════╝

// ==================== 1. CREATE EMAIL LOG SHEET ====================
function createEmailLogSheet(ss) {
  let sheet = ss.getSheetByName('Email Log');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Email Log');
  sheet.setTabColor('#f44336');
  
  const headers = [
    'Date/Time',       // A
    'Invoice No',      // B
    'Client Code',     // C
    'Client Name',     // D
    'Email',           // E
    'Language',        // F
    'Status',          // G - Sent/Failed
    'Error Message',   // H
    'Sent By',         // I - Auto/Manual
    'PDF Link'         // J
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#c62828')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  const widths = [150, 120, 90, 180, 200, 80, 80, 300, 80, 250];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  sheet.getRange(2, 1, 500, 1).setNumberFormat('yyyy-mm-dd HH:mm:ss');
  
  // Conditional formatting
  const statusRange = sheet.getRange(2, 7, 500, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Sent').setBackground('#c8e6c9').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Failed').setBackground('#ffcdd2').setRanges([statusRange]).build()
  ]);
  
  sheet.setFrozenRows(1);
  
  return sheet;
}

// ==================== 2. LOG EMAIL ====================
function logEmail(invoiceNo, clientCode, clientName, email, language, status, errorMsg, sentBy, pdfLink) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Email Log');
  if (!sheet) {
    sheet = createEmailLogSheet(ss);
  }
  
  const lastRow = sheet.getLastRow() + 1;
  
  sheet.getRange(lastRow, 1).setValue(new Date());
  sheet.getRange(lastRow, 2).setValue(invoiceNo);
  sheet.getRange(lastRow, 3).setValue(clientCode);
  sheet.getRange(lastRow, 4).setValue(clientName);
  sheet.getRange(lastRow, 5).setValue(email);
  sheet.getRange(lastRow, 6).setValue(language);
  sheet.getRange(lastRow, 7).setValue(status);
  sheet.getRange(lastRow, 8).setValue(errorMsg || '');
  sheet.getRange(lastRow, 9).setValue(sentBy || 'Manual');
  sheet.getRange(lastRow, 10).setValue(pdfLink || '');
}

// ==================== 3. WORKING DAYS CALCULATOR ====================

/**
 * التحقق هل اليوم عطلة رسمية
 */
function isHoliday(date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Holidays');
  if (!sheet) return false;
  
  const data = sheet.getDataRange().getValues();
  const checkDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      const holidayDate = Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (holidayDate === checkDate) {
        return true;
      }
    }
  }
  return false;
}

/**
 * التحقق هل اليوم عطلة أو نهاية أسبوع
 */
function isHolidayOrWeekend(date) {
  const day = date.getDay();
  // السبت = 6، الأحد = 0
  if (day === 0 || day === 6) return true;
  if (isHoliday(date)) return true;
  return false;
}

/**
 * التحقق هل اليوم يوم عمل
 */
function isWorkingDay(date) {
  return !isHolidayOrWeekend(date);
}

/**
 * إضافة أيام عمل لتاريخ معين
 */
function addWorkingDays(startDate, workingDays) {
  let currentDate = new Date(startDate);
  let addedDays = 0;
  
  while (addedDays < workingDays) {
    currentDate.setDate(currentDate.getDate() + 1);
    if (isWorkingDay(currentDate)) {
      addedDays++;
    }
  }
  
  return currentDate;
}

/**
 * حساب تاريخ الإرسال (3 أيام عمل بعد الإصدار)
 */
function calculateSendDate(issueDate) {
  return addWorkingDays(issueDate, 3);
}

/**
 * الحصول على تاريخ إصدار الفواتير (يوم 25 أو أقرب يوم عمل)
 */
function getInvoiceGenerationDate(year, month) {
  let genDate = new Date(year, month, 25);
  
  // إذا كان يوم 25 عطلة، انتقل لأقرب يوم عمل
  while (isHolidayOrWeekend(genDate)) {
    genDate.setDate(genDate.getDate() + 1);
  }
  
  genDate.setHours(9, 0, 0, 0);
  return genDate;
}

/**
 * الحصول على تاريخ إرسال الفواتير (3 أيام عمل بعد الإصدار)
 */
function getInvoiceSendDate(generationDate) {
  return addWorkingDays(generationDate, 3);
}

/**
 * اختبار جدول الفواتير
 */
function testInvoiceSchedule() {
  const ui = SpreadsheetApp.getUi();
  const now = new Date();
  
  const thisMonthGen = getInvoiceGenerationDate(now.getFullYear(), now.getMonth());
  const thisMonthSend = getInvoiceSendDate(thisMonthGen);
  
  const nextMonth = now.getMonth() === 11 ? 0 : now.getMonth() + 1;
  const nextYear = now.getMonth() === 11 ? now.getFullYear() + 1 : now.getFullYear();
  const nextMonthGen = getInvoiceGenerationDate(nextYear, nextMonth);
  const nextMonthSend = getInvoiceSendDate(nextMonthGen);
  
  ui.alert(
    '📅 Invoice Schedule Test\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'This Month:\n' +
    '• Generation: ' + Utilities.formatDate(thisMonthGen, Session.getScriptTimeZone(), 'EEEE, dd MMM yyyy') + '\n' +
    '• Send: ' + Utilities.formatDate(thisMonthSend, Session.getScriptTimeZone(), 'EEEE, dd MMM yyyy') + '\n\n' +
    'Next Month:\n' +
    '• Generation: ' + Utilities.formatDate(nextMonthGen, Session.getScriptTimeZone(), 'EEEE, dd MMM yyyy') + '\n' +
    '• Send: ' + Utilities.formatDate(nextMonthSend, Session.getScriptTimeZone(), 'EEEE, dd MMM yyyy')
  );
}

// ==================== 4. EMAIL TEMPLATES (3 LANGUAGES) ====================

/**
 * الحصول على قالب الإيميل حسب اللغة
 */
function getEmailTemplate(language, data) {
  const templates = {
    'EN': {
      subject: 'Invoice ' + data.invoiceNo + ' - ' + (getSettingValue('Company Name (EN)') || 'Dewan Consulting'),
      body: `
        <div style="font-family: Arial, sans-serif; padding: 20px; max-width: 600px;">
          <div style="background: #1565c0; color: white; padding: 20px; text-align: center;">
            <h1 style="margin: 0;">${getSettingValue('Company Name (EN)') || 'Dewan Consulting'}</h1>
            <p style="margin: 5px 0 0 0;">${getSettingValue('Company Name (AR)') || 'ديوان للاستشارات'}</p>
          </div>
          
          <div style="padding: 20px; background: #f5f5f5;">
            <p>Dear <strong>${data.contactPerson || data.clientName}</strong>,</p>
            
            <p>Please find attached your invoice for <strong>${data.period}</strong>.</p>
            
            <div style="background: white; padding: 15px; border-radius: 5px; margin: 20px 0;">
              <h3 style="margin-top: 0; color: #1565c0;">Invoice Details:</h3>
              <table style="width: 100%;">
                <tr><td><strong>Invoice No:</strong></td><td>${data.invoiceNo}</td></tr>
                <tr><td><strong>Date:</strong></td><td>${data.invoiceDate}</td></tr>
                <tr><td><strong>Amount:</strong></td><td style="font-size: 18px; color: #1565c0;"><strong>${data.amount}</strong></td></tr>
              </table>
            </div>
            
            <div style="background: white; padding: 15px; border-radius: 5px;">
              <h3 style="margin-top: 0; color: #1565c0;">Payment Details:</h3>
              <p><strong>Bank:</strong> ${getSettingValue('Bank Name') || 'Kuveyt Türk'}</p>
              <p><strong>IBAN (TRY):</strong> ${getSettingValue('IBAN TRY') || ''}</p>
              <p><strong>IBAN (USD):</strong> ${getSettingValue('IBAN USD') || ''}</p>
              <p><strong>SWIFT:</strong> ${getSettingValue('SWIFT Code') || 'KTEFTRIS'}</p>
            </div>
            
            <p style="margin-top: 20px;">If you have any questions, please don't hesitate to contact us.</p>
            
            <p>Best regards,<br>
            <strong>${getSettingValue('Company Name (EN)') || 'Dewan Consulting'}</strong><br>
            <a href="mailto:sales@aldewan.net">sales@aldewan.net</a></p>
          </div>
          
          <div style="background: #333; color: white; padding: 10px; text-align: center; font-size: 12px;">
            Thank you for your business!
          </div>
        </div>
      `
    },
    
    'TR': {
      subject: 'Fatura ' + data.invoiceNo + ' - ' + (getSettingValue('Company Name (EN)') || 'Dewan Consulting'),
      body: `
        <div style="font-family: Arial, sans-serif; padding: 20px; max-width: 600px;">
          <div style="background: #1565c0; color: white; padding: 20px; text-align: center;">
            <h1 style="margin: 0;">${getSettingValue('Company Name (EN)') || 'Dewan Consulting'}</h1>
            <p style="margin: 5px 0 0 0;">${getSettingValue('Company Name (AR)') || 'ديوان للاستشارات'}</p>
          </div>
          
          <div style="padding: 20px; background: #f5f5f5;">
            <p>Sayın <strong>${data.contactPerson || data.clientName}</strong>,</p>
            
            <p><strong>${data.period}</strong> dönemine ait faturanız ekte sunulmuştur.</p>
            
            <div style="background: white; padding: 15px; border-radius: 5px; margin: 20px 0;">
              <h3 style="margin-top: 0; color: #1565c0;">Fatura Detayları:</h3>
              <table style="width: 100%;">
                <tr><td><strong>Fatura No:</strong></td><td>${data.invoiceNo}</td></tr>
                <tr><td><strong>Tarih:</strong></td><td>${data.invoiceDate}</td></tr>
                <tr><td><strong>Tutar:</strong></td><td style="font-size: 18px; color: #1565c0;"><strong>${data.amount}</strong></td></tr>
              </table>
            </div>
            
            <div style="background: white; padding: 15px; border-radius: 5px;">
              <h3 style="margin-top: 0; color: #1565c0;">Ödeme Bilgileri:</h3>
              <p><strong>Banka:</strong> ${getSettingValue('Bank Name') || 'Kuveyt Türk'}</p>
              <p><strong>IBAN (TRY):</strong> ${getSettingValue('IBAN TRY') || ''}</p>
              <p><strong>IBAN (USD):</strong> ${getSettingValue('IBAN USD') || ''}</p>
              <p><strong>SWIFT:</strong> ${getSettingValue('SWIFT Code') || 'KTEFTRIS'}</p>
            </div>
            
            <p style="margin-top: 20px;">Herhangi bir sorunuz varsa lütfen bizimle iletişime geçin.</p>
            
            <p>Saygılarımızla,<br>
            <strong>${getSettingValue('Company Name (EN)') || 'Dewan Consulting'}</strong><br>
            <a href="mailto:sales@aldewan.net">sales@aldewan.net</a></p>
          </div>
          
          <div style="background: #333; color: white; padding: 10px; text-align: center; font-size: 12px;">
            İlginiz için teşekkür ederiz!
          </div>
        </div>
      `
    },
    
    'AR': {
      subject: 'فاتورة ' + data.invoiceNo + ' - ' + (getSettingValue('Company Name (AR)') || 'ديوان للاستشارات'),
      body: `
        <div dir="rtl" style="font-family: Arial, sans-serif; padding: 20px; max-width: 600px;">
          <div style="background: #1565c0; color: white; padding: 20px; text-align: center;">
            <h1 style="margin: 0;">${getSettingValue('Company Name (AR)') || 'ديوان للاستشارات'}</h1>
            <p style="margin: 5px 0 0 0;">${getSettingValue('Company Name (EN)') || 'Dewan Consulting'}</p>
          </div>
          
          <div style="padding: 20px; background: #f5f5f5;">
            <p>السيد/السيدة <strong>${data.contactPerson || data.clientName}</strong> المحترم/ة،</p>
            
            <p>مرفق لكم فاتورة <strong>${data.period}</strong>.</p>
            
            <div style="background: white; padding: 15px; border-radius: 5px; margin: 20px 0;">
              <h3 style="margin-top: 0; color: #1565c0;">تفاصيل الفاتورة:</h3>
              <table style="width: 100%;">
                <tr><td><strong>رقم الفاتورة:</strong></td><td>${data.invoiceNo}</td></tr>
                <tr><td><strong>التاريخ:</strong></td><td>${data.invoiceDate}</td></tr>
                <tr><td><strong>المبلغ:</strong></td><td style="font-size: 18px; color: #1565c0;"><strong>${data.amount}</strong></td></tr>
              </table>
            </div>
            
            <div style="background: white; padding: 15px; border-radius: 5px;">
              <h3 style="margin-top: 0; color: #1565c0;">معلومات الدفع:</h3>
              <p><strong>البنك:</strong> ${getSettingValue('Bank Name') || 'Kuveyt Türk'}</p>
              <p><strong>IBAN (TRY):</strong> ${getSettingValue('IBAN TRY') || ''}</p>
              <p><strong>IBAN (USD):</strong> ${getSettingValue('IBAN USD') || ''}</p>
              <p><strong>SWIFT:</strong> ${getSettingValue('SWIFT Code') || 'KTEFTRIS'}</p>
            </div>
            
            <p style="margin-top: 20px;">في حال وجود أي استفسار، لا تترددوا في التواصل معنا.</p>
            
            <p>مع أطيب التحيات،<br>
            <strong>${getSettingValue('Company Name (AR)') || 'ديوان للاستشارات'}</strong><br>
            <a href="mailto:sales@aldewan.net">sales@aldewan.net</a></p>
          </div>
          
          <div style="background: #333; color: white; padding: 10px; text-align: center; font-size: 12px;">
            شكراً لتعاملكم معنا!
          </div>
        </div>
      `
    }
  };
  
  return templates[language] || templates['EN'];
}

// ==================== 5. SEND SINGLE INVOICE EMAIL ====================

/**
 * إرسال فاتورة واحدة بالإيميل
 */
function sendInvoiceEmail(invoiceNo, sentBy) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get invoice data from Invoice Log
  const logSheet = ss.getSheetByName('Invoice Log');
  if (!logSheet) return { success: false, error: 'Invoice Log not found' };
  
  const logData = logSheet.getDataRange().getValues();
  let invoiceRow = -1;
  let invoiceData = null;
  
  for (let i = 1; i < logData.length; i++) {
    if (logData[i][0] === invoiceNo) {
      invoiceRow = i + 1;
      invoiceData = {
        invoiceNo: logData[i][0],
        invoiceDate: logData[i][1],
        clientCode: logData[i][2],
        clientName: logData[i][3],
        service: logData[i][4],
        period: logData[i][5],
        amount: logData[i][6],
        currency: logData[i][7],
        pdfLink: logData[i][9]
      };
      break;
    }
  }
  
  if (!invoiceData) {
    return { success: false, error: 'Invoice not found: ' + invoiceNo };
  }
  
  // Get client data
  const clientData = getClientData(invoiceData.clientCode);
  if (!clientData) {
    logEmail(invoiceNo, invoiceData.clientCode, invoiceData.clientName, '', '', 'Failed', 'Client not found', sentBy, '');
    return { success: false, error: 'Client not found: ' + invoiceData.clientCode };
  }
  
  const clientEmail = clientData.email;
  const clientLanguage = clientData.language || 'EN';
  const contactPerson = clientData.contactPerson || '';
  
  if (!clientEmail) {
    logEmail(invoiceNo, invoiceData.clientCode, invoiceData.clientName, '', clientLanguage, 'Failed', 'No email address', sentBy, '');
    return { success: false, error: 'No email for client: ' + invoiceData.clientCode };
  }
  
  // Get email template
  const template = getEmailTemplate(clientLanguage, {
    invoiceNo: invoiceData.invoiceNo,
    invoiceDate: Utilities.formatDate(new Date(invoiceData.invoiceDate), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    clientName: invoiceData.clientName,
    contactPerson: contactPerson,
    period: invoiceData.period,
    amount: formatCurrency(invoiceData.amount, invoiceData.currency)
  });
  
  try {
    // Get PDF file if exists
    let attachments = [];
    if (invoiceData.pdfLink) {
      try {
        const fileId = extractFileIdFromUrl(invoiceData.pdfLink);
        if (fileId) {
          const file = DriveApp.getFileById(fileId);
          attachments.push(file.getAs(MimeType.PDF));
        }
      } catch (e) {
        console.log('Could not attach PDF: ' + e.message);
      }
    }
    
    // Send email
    GmailApp.sendEmail(clientEmail, template.subject, '', {
      htmlBody: template.body,
      name: getSettingValue('Company Name (EN)') || 'Dewan Consulting',
      replyTo: 'sales@aldewan.net',
      attachments: attachments
    });
    
    // Update Invoice Log
    logSheet.getRange(invoiceRow, 9).setValue('Sent');
    logSheet.getRange(invoiceRow, 12).setValue('Sent');
    logSheet.getRange(invoiceRow, 13).setValue(new Date());
    
    // Log to Email Log
    logEmail(invoiceData.invoiceNo, invoiceData.clientCode, invoiceData.clientName, 
             clientEmail, clientLanguage, 'Sent', '', sentBy, invoiceData.pdfLink);
    
    return { success: true, email: clientEmail };
    
  } catch (e) {
    // Log failure
    logEmail(invoiceData.invoiceNo, invoiceData.clientCode, invoiceData.clientName,
             clientEmail, clientLanguage, 'Failed', e.message, sentBy, invoiceData.pdfLink);
    
    // Update Invoice Log
    logSheet.getRange(invoiceRow, 12).setValue('Failed');
    
    return { success: false, error: e.message };
  }
}

/**
 * استخراج File ID من رابط Drive
 */
function extractFileIdFromUrl(url) {
  if (!url) return null;
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

// ==================== 6. SEND PENDING INVOICES ====================

/**
 * إرسال كل الفواتير المعلقة (بعد 3 أيام عمل)
 */
function sendPendingInvoices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const logSheet = ss.getSheetByName('Invoice Log');
  
  if (!logSheet || logSheet.getLastRow() < 2) {
    ui.alert('⚠️ No invoices in Invoice Log!');
    return;
  }
  
  const data = logSheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  let readyToSend = [];
  let notReady = [];
  
  for (let i = 1; i < data.length; i++) {
    const invoiceNo = data[i][0];
    const invoiceDate = new Date(data[i][1]);
    const sendEmail = data[i][10];
    const emailStatus = data[i][11];
    
    if (sendEmail !== 'Yes' || emailStatus === 'Sent') continue;
    
    // Calculate send date (3 working days after invoice date)
    const sendDate = calculateSendDate(invoiceDate);
    sendDate.setHours(0, 0, 0, 0);
    
    if (today >= sendDate) {
      readyToSend.push({
        row: i + 1,
        invoiceNo: invoiceNo,
        clientName: data[i][3],
        sendDate: sendDate
      });
    } else {
      notReady.push({
        invoiceNo: invoiceNo,
        clientName: data[i][3],
        sendDate: sendDate
      });
    }
  }
  
  if (readyToSend.length === 0 && notReady.length === 0) {
    ui.alert('✅ No pending invoices to send!');
    return;
  }
  
  let message = '📧 Pending Invoices Report\n\n';
  
  if (readyToSend.length > 0) {
    message += '✅ Ready to Send (' + readyToSend.length + '):\n';
    readyToSend.forEach(inv => {
      message += '• ' + inv.invoiceNo + ' - ' + inv.clientName + '\n';
    });
    message += '\n';
  }
  
  if (notReady.length > 0) {
    message += '⏳ Not Ready Yet (' + notReady.length + '):\n';
    notReady.slice(0, 5).forEach(inv => {
      message += '• ' + inv.invoiceNo + ' - Send: ' + Utilities.formatDate(inv.sendDate, Session.getScriptTimeZone(), 'dd/MM/yyyy') + '\n';
    });
    if (notReady.length > 5) {
      message += '... and ' + (notReady.length - 5) + ' more\n';
    }
    message += '\n';
  }
  
  if (readyToSend.length === 0) {
    ui.alert(message + 'No invoices ready to send yet.');
    return;
  }
  
  const confirm = ui.alert(
    message + 'Send ' + readyToSend.length + ' invoices now?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  let sent = 0, failed = 0;
  const results = [];
  
  readyToSend.forEach(inv => {
    const result = sendInvoiceEmail(inv.invoiceNo, 'Manual');
    
    if (result.success) {
      sent++;
      results.push({ invoice: inv.invoiceNo, status: '✅ Sent', detail: result.email });
    } else {
      failed++;
      results.push({ invoice: inv.invoiceNo, status: '❌ Failed', detail: result.error });
    }
  });
  
  let report = '📧 Send Complete!\n\n';
  report += '✅ Sent: ' + sent + '\n';
  report += '❌ Failed: ' + failed + '\n\n';
  
  if (results.length <= 10) {
    report += 'Details:\n';
    results.forEach(r => {
      report += r.status + ' ' + r.invoice + ': ' + r.detail + '\n';
    });
  }
  
  ui.alert(report);
}

/**
 * إرسال فاتورة محددة يدوياً
 */
function sendSelectedInvoice() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();
  
  if (sheet.getName() !== 'Invoice Log') {
    ui.alert('⚠️ Please go to Invoice Log sheet and select an invoice row.');
    return;
  }
  
  const row = sheet.getActiveCell().getRow();
  if (row < 2) {
    ui.alert('⚠️ Please select an invoice row (not header).');
    return;
  }
  
  const invoiceNo = sheet.getRange(row, 1).getValue();
  const clientName = sheet.getRange(row, 4).getValue();
  const emailStatus = sheet.getRange(row, 12).getValue();
  
  if (emailStatus === 'Sent') {
    const resend = ui.alert(
      '⚠️ Invoice Already Sent!\n\n' +
      'Invoice: ' + invoiceNo + '\n' +
      'Client: ' + clientName + '\n\n' +
      'Send again?',
      ui.ButtonSet.YES_NO
    );
    if (resend !== ui.Button.YES) return;
  }
  
  const confirm = ui.alert(
    '📧 Send Invoice Email\n\n' +
    'Invoice: ' + invoiceNo + '\n' +
    'Client: ' + clientName + '\n\n' +
    'Send now?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  const result = sendInvoiceEmail(invoiceNo, 'Manual');
  
  if (result.success) {
    ui.alert('✅ Email sent successfully!\n\nTo: ' + result.email);
  } else {
    ui.alert('❌ Failed to send email:\n\n' + result.error);
  }
}

// ==================== 7. SHOW EMAIL LOG ====================
function showEmailLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Email Log');
  if (!sheet) {
    sheet = createEmailLogSheet(ss);
  }
  ss.setActiveSheet(sheet);
}

// ==================== 8. EMAIL STATISTICS ====================
function showEmailStatistics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Email Log');
  
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('📊 No email data yet.');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  let sent = 0, failed = 0;
  const byLanguage = { EN: 0, TR: 0, AR: 0 };
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] === 'Sent') {
      sent++;
      const lang = data[i][5] || 'EN';
      byLanguage[lang] = (byLanguage[lang] || 0) + 1;
    } else if (data[i][6] === 'Failed') {
      failed++;
    }
  }
  
  ui.alert(
    '📊 Email Statistics\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '✅ Sent: ' + sent + '\n' +
    '❌ Failed: ' + failed + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'By Language:\n' +
    '🇬🇧 English: ' + (byLanguage.EN || 0) + '\n' +
    '🇹🇷 Turkish: ' + (byLanguage.TR || 0) + '\n' +
    '🇸🇦 Arabic: ' + (byLanguage.AR || 0)
  );
}

// ==================== 9. TRIGGERS MANAGEMENT ====================

/**
 * إعداد الـ Triggers التلقائية
 */
function setupAutoTriggers() {
  const ui = SpreadsheetApp.getUi();
  
  const confirm = ui.alert(
    '⏰ Setup Automatic Triggers\n\n' +
    'This will create:\n\n' +
    '1. 📅 Monthly Invoice Generation\n' +
    '   Day 25 at 9:00 AM (or next working day)\n\n' +
    '2. 📧 Daily Email Check\n' +
    '   Every day at 10:00 AM\n' +
    '   (Sends invoices after 3 working days)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  // Remove existing triggers
  removeAllTriggers(true);
  
  try {
    // Monthly invoice generation - day 25 at 9:00 AM
    ScriptApp.newTrigger('autoGenerateMonthlyInvoices')
      .timeBased()
      .onMonthDay(25)
      .atHour(9)
      .create();
    
    // Daily email check - 10:00 AM
    ScriptApp.newTrigger('autoSendPendingInvoices')
      .timeBased()
      .everyDays(1)
      .atHour(10)
      .create();
    
    ui.alert(
      '✅ Triggers Created!\n\n' +
      '📅 Monthly Invoices: Day 25 at 9:00 AM\n' +
      '📧 Email Check: Daily at 10:00 AM\n\n' +
      'The system will:\n' +
      '• Generate invoices on day 25\n' +
      '• Wait 3 working days\n' +
      '• Send emails automatically\n\n' +
      'Holidays and weekends are respected.'
    );
    
  } catch (error) {
    ui.alert('❌ Error creating triggers:\n\n' + error.message);
  }
}

/**
 * إزالة كل الـ Triggers
 */
function removeAllTriggers(silent) {
  const triggers = ScriptApp.getProjectTriggers();
  
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  
  if (!silent) {
    SpreadsheetApp.getUi().alert('✅ All triggers removed!\n\nRemoved: ' + triggers.length + ' triggers');
  }
}

/**
 * عرض حالة الـ Triggers
 */
function showTriggersStatus() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  
  if (triggers.length === 0) {
    ui.alert('⚠️ No active triggers.\n\nUse "Setup Auto Triggers" to create them.');
    return;
  }
  
  let status = '⏰ Active Triggers:\n\n';
  
  triggers.forEach(trigger => {
    const funcName = trigger.getHandlerFunction();
    const type = trigger.getEventType();
    status += '• ' + funcName + ' (' + type + ')\n';
  });
  
  ui.alert(status);
}

// ==================== 10. AUTO TRIGGER FUNCTIONS ====================

/**
 * إصدار الفواتير الشهرية تلقائياً
 */
function autoGenerateMonthlyInvoices() {
  const today = new Date();
  
  // Check if today is a working day
  if (isHolidayOrWeekend(today)) {
    return;
  }
  
  // Check if this is the correct generation day
  const expectedDate = getInvoiceGenerationDate(today.getFullYear(), today.getMonth());
  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const expectedStr = Utilities.formatDate(expectedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  if (todayStr !== expectedStr) {
    return;
  }
  
  // Generate all monthly invoices
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const clients = getActiveClients().filter(c => c.monthlyFee > 0);
    
    if (clients.length === 0) return;
    
    const invoiceDate = new Date();
    const period = Utilities.formatDate(invoiceDate, Session.getScriptTimeZone(), 'MMMM yyyy');
    
    clients.forEach(client => {
      const invoiceNo = getNextInvoiceNumber();
      const clientData = getClientData(client.code);
      
      // Fill template
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
        } catch (e) {
          console.log('PDF error: ' + e.message);
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
    });
    
    // Log the action
    logAlert('Auto Invoice', clients.length + ' monthly invoices generated for ' + period, 'Info');
    
  } catch (error) {
    console.error('Auto generate error: ' + error);
    logAlert('Auto Invoice Error', error.message, 'Error');
  }
}

/**
 * إرسال الفواتير المعلقة تلقائياً
 */
function autoSendPendingInvoices() {
  const today = new Date();
  
  // Skip weekends
  if (today.getDay() === 0 || today.getDay() === 6) {
    return;
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('Invoice Log');
    
    if (!logSheet || logSheet.getLastRow() < 2) return;
    
    const data = logSheet.getDataRange().getValues();
    let sent = 0, failed = 0;
    
    for (let i = 1; i < data.length; i++) {
      const invoiceNo = data[i][0];
      const invoiceDate = new Date(data[i][1]);
      const sendEmail = data[i][10];
      const emailStatus = data[i][11];
      
      if (sendEmail !== 'Yes' || emailStatus === 'Sent') continue;
      
      // Check if 3 working days have passed
      const sendDate = calculateSendDate(invoiceDate);
      
      if (today >= sendDate) {
        const result = sendInvoiceEmail(invoiceNo, 'Auto');
        
        if (result.success) {
          sent++;
        } else {
          failed++;
        }
      }
    }
    
    if (sent > 0 || failed > 0) {
      logAlert('Auto Email', 'Sent: ' + sent + ', Failed: ' + failed, sent > 0 ? 'Info' : 'Warning');
    }
    
  } catch (error) {
    console.error('Auto send error: ' + error);
    logAlert('Auto Email Error', error.message, 'Error');
  }
}

// ==================== 11. ALERTS LOG ====================

function createAlertsLogSheet(ss) {
  let sheet = ss.getSheetByName('Alerts Log');
  if (sheet) return sheet;
  
  sheet = ss.insertSheet('Alerts Log');
  sheet.setTabColor('#ff9800');
  
  const headers = ['Date/Time', 'Type', 'Message', 'Severity', 'Acknowledged'];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#e65100')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  const widths = [150, 120, 400, 80, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  // Conditional formatting
  const severityRange = sheet.getRange(2, 4, 500, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Error').setBackground('#ffcdd2').setRanges([severityRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Warning').setBackground('#fff9c4').setRanges([severityRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Info').setBackground('#c8e6c9').setRanges([severityRange]).build()
  ]);
  
  sheet.setFrozenRows(1);
  
  return sheet;
}

function logAlert(type, message, severity) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Alerts Log');
  if (!sheet) {
    sheet = createAlertsLogSheet(ss);
  }
  
  const lastRow = sheet.getLastRow() + 1;
  
  sheet.getRange(lastRow, 1).setValue(new Date()).setNumberFormat('yyyy-mm-dd HH:mm');
  sheet.getRange(lastRow, 2).setValue(type);
  sheet.getRange(lastRow, 3).setValue(message);
  sheet.getRange(lastRow, 4).setValue(severity || 'Info');
  sheet.getRange(lastRow, 5).setValue('No');
}

function showAlertsLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Alerts Log');
  if (!sheet) {
    sheet = createAlertsLogSheet(ss);
  }
  ss.setActiveSheet(sheet);
}

// ==================== 12. OVERDUE REMINDERS ====================

/**
 * التحقق من المدفوعات المتأخرة
 */
function checkOverduePayments() {
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
  
  const reminderDays = parseInt(getSettingValue('First Reminder (Days)')) || 30;
  
  for (let i = 1; i < data.length; i++) {
    const status = data[i][18]; // Status
    const dueDate = data[i][19]; // Due Date
    const amount = data[i][10]; // Amount
    const clientName = data[i][5]; // Client Name
    const invoiceNo = data[i][17]; // Invoice No
    const currency = data[i][11]; // Currency
    
    if (status && status.includes('Pending') && dueDate) {
      const due = new Date(dueDate);
      const diffDays = Math.floor((today - due) / (1000 * 60 * 60 * 24));
      
      if (diffDays > 0) {
        overdueList.push({
          row: i + 1,
          clientName: clientName,
          invoiceNo: invoiceNo || 'N/A',
          amount: formatCurrency(amount, currency),
          dueDate: Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          daysOverdue: diffDays
        });
      }
    }
  }
  
  if (overdueList.length === 0) {
    ui.alert('✅ No overdue payments!\n\nAll pending payments are still within their due dates.');
    return;
  }
  
  // Sort by days overdue
  overdueList.sort((a, b) => b.daysOverdue - a.daysOverdue);
  
  let report = '⚠️ Overdue Payments Report\n\n';
  report += 'Found ' + overdueList.length + ' overdue payments:\n\n';
  
  overdueList.slice(0, 10).forEach(o => {
    report += '• ' + o.clientName + '\n';
    report += '  Invoice: ' + o.invoiceNo + ' | ' + o.amount + '\n';
    report += '  Overdue: ' + o.daysOverdue + ' days\n\n';
  });
  
  if (overdueList.length > 10) {
    report += '... and ' + (overdueList.length - 10) + ' more\n';
  }
  
  ui.alert(report);
}

// ==================== END OF PART 7 ====================
