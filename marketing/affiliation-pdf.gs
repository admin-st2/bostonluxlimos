/**
 * BostonLuxLimos affiliation intake: private PDF + owner email.
 * Bound to the affiliation response spreadsheet. No web deployment needed.
 * Run setupAffiliationPdf once, approve Google's permissions, then run
 * testAffiliationPdf to process ONLY the existing clearly marked QA row.
 */
const BLL_PDF = Object.freeze({
  spreadsheetId: '18u6C4XOewDcUydCiAwkkDd3sUOPdLyrQb5B6R8ZR_Ko',
  sheetId: 8253552,
  folderName: 'BostonLuxLimos — Private Affiliation PDFs',
  timeZone: 'America/New_York',
  handler: 'affiliationPdfOnSubmit',
  responseColumns: 12
});

function setupAffiliationPdf() {
  const ss = SpreadsheetApp.openById(BLL_PDF.spreadsheetId);
  const ownerEmail = Session.getEffectiveUser().getEmail();
  if (!ownerEmail || !ownerEmail.endsWith('@bostonluxlimos.com')) {
    throw new Error('Run setup using the BostonLuxLimos Workspace owner account.');
  }
  PropertiesService.getScriptProperties().setProperty('BLL_PDF_OWNER', ownerEmail);
  const folder = affiliationPdfFolder_();
  const found = ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === BLL_PDF.handler &&
    t.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT &&
    t.getTriggerSourceId() === BLL_PDF.spreadsheetId);
  if (!found) ScriptApp.newTrigger(BLL_PDF.handler).forSpreadsheet(ss).onFormSubmit().create();
  console.log('Configured. Private PDF folder: ' + folder.getUrl());
  console.log('Owner notification: ' + ownerEmail);
  console.log('Next: run testAffiliationPdf, then verify one new form submission.');
}

function affiliationPdfOnSubmit(e) {
  if (!e || !e.range) throw new Error('This function requires the spreadsheet form-submit trigger.');
  const sheet = e.range.getSheet();
  if (sheet.getParent().getId() !== BLL_PDF.spreadsheetId ||
      sheet.getSheetId() !== BLL_PDF.sheetId || e.range.getRow() < 2) {
    throw new Error('Unexpected submission source; no PDF was created.');
  }
  createAffiliationPdf_(sheet, e.range.getRow());
}

function testAffiliationPdf() {
  const sheet = SpreadsheetApp.openById(BLL_PDF.spreadsheetId).getSheetById(BLL_PDF.sheetId);
  const rows = sheet.getDataRange().getDisplayValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (/^TEST — BostonLuxLimos website integration/.test(rows[i][2] || '')) {
      createAffiliationPdf_(sheet, i + 1);
      return;
    }
  }
  throw new Error('No clearly marked affiliation QA row found; no customer record was processed.');
}

function affiliationPdfFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('BLL_PDF_FOLDER');
  if (id) return DriveApp.getFolderById(id);
  // New folders in the owner's My Drive are private; never enable link sharing.
  const folder = DriveApp.createFolder(BLL_PDF.folderName);
  props.setProperty('BLL_PDF_FOLDER', folder.getId());
  return folder;
}

function createAffiliationPdf_(sheet, row) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  const props = PropertiesService.getScriptProperties();
  const key = 'BLL_PDF_ROW_' + row;
  try {
    const ownerEmail = props.getProperty('BLL_PDF_OWNER');
    if (!ownerEmail) throw new Error('Run setupAffiliationPdf first.');
    const data = sheet.getRange(row, 1, 1, BLL_PDF.responseColumns).getDisplayValues()[0];
    const headers = sheet.getRange(1, 1, 1, BLL_PDF.responseColumns).getDisplayValues()[0];
    if (!data[0] || !data[2]) throw new Error('Submission is missing its timestamp or company.');
    const previous = JSON.parse(props.getProperty(key) || '{}');
    if (previous.emailSent) {
      console.log('Already completed: row ' + row + '. No duplicate PDF or email.');
      return;
    }
    const reference = 'BLL-AFF-' + String(row - 1).padStart(5, '0');
    const company = String(data[2]).replace(/[\r\n]+/g, ' ').slice(0, 150);
    const safeName = company.replace(/[^a-zA-Z0-9 ._-]/g, '').trim().slice(0, 75) || 'Company';
    const html = affiliationPdfHtml_(reference, headers, data);
    let file;
    if (previous.fileId) {
      file = DriveApp.getFileById(previous.fileId);
    } else {
      const pdf = HtmlService.createHtmlOutput(html).getAs(MimeType.PDF)
        .setName(reference + ' - ' + safeName + '.pdf');
      file = affiliationPdfFolder_().createFile(pdf);
      props.setProperty(key, JSON.stringify({fileId: file.getId(), emailSent: false}));
    }
    const auditSheet = affiliationPdfAudit_(sheet.getParent());
    let auditRow = previous.auditRow;
    if (!auditRow) {
      auditRow = auditSheet.getLastRow() + 1;
      auditSheet.getRange(auditRow, 1, 1, 6).setValues([[
        reference, row, new Date(), file.getUrl(), 'PDF saved; email pending', ''
      ]]);
      props.setProperty(key, JSON.stringify({fileId: file.getId(), auditRow, emailSent: false}));
    }
    if (MailApp.getRemainingDailyQuota() < 1) throw new Error('Email quota unavailable; PDF saved. Retry after quota resets.');
    MailApp.sendEmail({
      to: ownerEmail,
      subject: 'BostonLuxLimos | Partner Introduction | ' + reference,
      name: 'BostonLuxLimos Affiliation Desk',
      body: 'A partner introduction has been recorded.\n\nReference: ' + reference +
        '\nCompany: ' + company + '\nPrivate PDF: ' + file.getUrl() +
        '\n\nThe attached PDF contains the submitted coverage, fee structure and operating details. ' +
        'Please review and reply personally to the business contact. No affiliation, rate or trip has been accepted. ' +
        'The PDF is sent only to the Workspace owner, not to the applicant.',
      attachments: [file.getBlob()]
    });
    props.setProperty(key, JSON.stringify({fileId: file.getId(), auditRow, emailSent: true}));
    auditSheet.getRange(auditRow, 5, 1, 2).setValues([['PDF saved; owner email sent', new Date()]]);
    console.log(reference + ' complete. Private PDF: ' + file.getUrl());
  } catch (err) {
    console.error('Affiliation PDF row ' + row + ': ' + err.message);
    throw err; // Apps Script records failure; original form response remains intact.
  } finally {
    lock.releaseLock();
  }
}

function affiliationPdfAudit_(ss) {
  let sheet = ss.getSheetByName('PDF Delivery Log');
  if (!sheet) {
    sheet = ss.insertSheet('PDF Delivery Log');
    sheet.getRange(1, 1, 1, 6).setValues([[
      'Reference', 'Response row', 'PDF created', 'Private PDF URL', 'Delivery status', 'Owner emailed'
    ]]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 6).setBackground('#171717').setFontColor('#d4ba7c').setFontWeight('bold');
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(4, 360);
    sheet.setColumnWidth(5, 230);
    sheet.getRange('C:C').setNumberFormat('yyyy-mm-dd hh:mm');
    sheet.getRange('F:F').setNumberFormat('yyyy-mm-dd hh:mm');
  }
  return sheet;
}

function affiliationPdfHtml_(reference, headers, data) {
  const esc = v => String(v == null ? '' : v).replace(/[&<>"']/g, c =>
    ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const fields = headers.map((label, i) => {
    if (!data[i] || !label) return '';
    return '<section class="field"><h2>' + esc(label) + '</h2><p>' +
      esc(data[i]).replace(/\n/g, '<br>') + '</p></section>';
  }).join('');
  const generated = Utilities.formatDate(new Date(), BLL_PDF.timeZone, 'MMMM d, yyyy h:mm a z');
  return '<!doctype html><html><head><meta charset="UTF-8"><style>' +
    '@page{size:A4;margin:19mm 18mm}body{font-family:Georgia,"Times New Roman",serif;color:#202020;font-size:10.5pt;line-height:1.55}' +
    '.brand{font-family:"Times New Roman",serif;font-style:italic;font-size:25pt;margin:0}.tagline{font-size:10pt;color:#82632e;margin:0 0 18pt}' +
    '.rule{border-top:2pt solid #b99b60;margin:12pt 0 18pt}.eyebrow{font:8pt Arial,sans-serif;letter-spacing:2pt;color:#82632e;text-transform:uppercase}' +
    'h1{font-size:25pt;font-weight:normal;margin:7pt 0 5pt}.meta{font:9pt Arial,sans-serif;color:#555}.intro{margin:16pt 0 22pt}' +
    '.field{margin:0 0 14pt;padding-top:8pt;border-top:.5pt solid #d9d3c7}h2{font: bold 9pt Arial,sans-serif;color:#82632e;margin:0 0 5pt}' +
    'p{margin:0;overflow-wrap:anywhere;white-space:normal}.notice{background:#f3f0e9;padding:12pt;font-size:9pt;margin-top:20pt}' +
    '.footer{border-top:.5pt solid #b99b60;padding-top:9pt;margin-top:22pt;font-size:8pt;color:#666}</style></head><body>' +
    '<p class="brand">BostonLuxLimos</p><p class="tagline">Elegance in every journey</p><div class="rule"></div>' +
    '<p class="eyebrow">Worldwide affiliations</p><h1>Partner Introduction</h1>' +
    '<p class="meta">' + esc(reference) + ' · Prepared ' + esc(generated) + '</p>' +
    '<p class="intro">Different destinations. The same personal attention.<br>' +
    'An introduction to explore how we can support each other’s clients.</p>' + fields +
    '<div class="notice"><strong>For partnership review.</strong> This document records information provided by the applicant. ' +
    'Details have not been independently verified. It is not an agreement, an acceptance of rates, an approved affiliation or a confirmed reservation. ' +
    'Commercial arrangements and individual journeys require separate review and written confirmation.</div>' +
    '<p class="footer">BostonLuxLimos · Operated by STRUT2 LLC<br>bostonluxlimos.com · Private business correspondence</p>' +
    '</body></html>';
}
