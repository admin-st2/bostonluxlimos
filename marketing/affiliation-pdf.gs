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
  return affiliationPdfOnePageHtml_(reference, headers, data);
}

function previewAffiliationPdfRedesign() {
  const sheet = SpreadsheetApp.openById(BLL_PDF.spreadsheetId).getSheetById(BLL_PDF.sheetId);
  const data = sheet.getRange(3, 1, 1, BLL_PDF.responseColumns).getDisplayValues()[0];
  if (!/^TEST/.test(data[2] || '')) throw new Error('Preview is restricted to the internal test row.');
  const html = affiliationPdfOnePageHtml_('DESIGN PREVIEW', [], data);
  const file = affiliationPdfFolder_().createFile(HtmlService.createHtmlOutput(html).getAs(MimeType.PDF).setName('BostonLuxLimos - Partner Introduction - One Page Preview.pdf'));
  console.log('One-page design preview: ' + file.getUrl());
}

function affiliationPdfOnePageHtml_(reference, headers, data) {
  return affiliationPdfRefinedHtml_(reference, headers, data);
}

function previewLongAffiliationPdfRedesign() {
  const data = Array(12).fill('LAYOUT TEST ONLY - Longer sample response to verify that the single-page summary remains legible. '.repeat(30));
  data[0] = 'November 1, 2026';
  data[1] = 'layout-test@example.invalid';
  data[2] = 'LAYOUT TEST ONLY - maximum-length partner introduction';
  const html = affiliationPdfOnePageHtml_('LONG RESPONSE TEST', [], data);
  const file = affiliationPdfFolder_().createFile(HtmlService.createHtmlOutput(html).getAs(MimeType.PDF).setName('BostonLuxLimos - One Page Long Response Layout Test.pdf'));
  console.log('Long response design preview: ' + file.getUrl());
}

function affiliationPdfRefinedHtml_(reference, headers, data) {
  const dense = data.join('').length > 1000;
  const esc = v => String(v == null ? '' : v).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  let shortened = false;
  const value = (i, limit) => {
    let text = String(data[i] || '').replace(/\s+/g, ' ').trim();
    if (!text) return '<span class="empty">Not supplied</span>';
    if (text.length > limit) { text = text.slice(0, limit).replace(/\s+\S*$/, '') + '... [full response in inquiry]'; shortened = true; }
    // Break long unspaced strings so a URL cannot widen the table.
    return esc(text).replace(/([^\s<>]{35})(?=[^\s<>])/g, '$1<wbr>');
  };
  const box = (label, text) => '<td><p class="label">' + label + '</p><p class="answer">' + text + '</p></td>';
  const row = (a, b) => '<tr>' + a + b + '</tr>';
  const profile = row(box('Business contact',value(3,120)), box('Email',value(1,110))) +
    row(box('Destinations & airports',value(5,230)),box('Dispatch & availability',value(6,230))) +
    row(box('Vehicles & accessibility',value(7,230)),box('Client support & referrals',value(10,280)));
  const commercial = row(box('Rates, currency & additional fees',value(8,350)),box('Cancellations, waiting & payment',value(9,350)));
  const standards = value(11,330);
  const company = value(2,170);
  const source = 'https://docs.google.com/spreadsheets/d/' + BLL_PDF.spreadsheetId + '/edit#gid=' + BLL_PDF.sheetId;
  return '<!doctype html><html><head><meta charset="UTF-8"><style>' +
    '@page{size:A4;margin:13mm 15mm}*{box-sizing:border-box}body{margin:0;color:#222;font:9.2pt Georgia,"Times New Roman",serif;line-height:1.35}' +
    'table{border-collapse:collapse;table-layout:fixed;width:100%}td{vertical-align:top;width:50%;padding:9pt 11pt;border:1px solid #ded9cf}' +
    '.masthead td{border:0;padding:0 0 12pt}.masthead .right{text-align:right;vertical-align:middle}.brand{font:italic 25pt "Times New Roman",serif;margin:0}.tag{font-size:9pt;color:#82632e;margin:3pt 0 0}' +
    '.kicker{font:7.2pt Arial,sans-serif;letter-spacing:1.4pt;text-transform:uppercase;color:#82632e;margin:0 0 5pt}.ref{font:8pt Arial,sans-serif;color:#666;margin:0}' +
    '.hero{border-top:2pt solid #b99b60;padding:15pt 0 12pt}h1{font-size:26pt;font-weight:normal;line-height:1.1;margin:0 0 5pt}.sub{font-size:10pt;color:#5d5140;margin:0}' +
    '.company{background:#f2efe8;border-left:3pt solid #b99b60;padding:10pt 12pt;margin-bottom:12pt}.company .answer{font-size:12pt;line-height:1.25}' +
    '.label{font:bold 7.2pt Arial,sans-serif;letter-spacing:.5pt;text-transform:uppercase;color:#786037;margin:0 0 5pt}.answer{margin:0;word-wrap:break-word;overflow-wrap:break-word}' +
    '.section{font:8pt Arial,sans-serif;letter-spacing:1.2pt;text-transform:uppercase;color:#82632e;margin:13pt 0 6pt}.standards{padding:10pt 11pt;border:1px solid #ded9cf}' +
    '.notice{font:7.5pt Arial,sans-serif;line-height:1.4;color:#68635c;margin:13pt 0 10pt}.footer{border-top:1pt solid #b99b60;padding-top:8pt;font:7.2pt Arial,sans-serif;color:#666}.footer a{color:#786037}.empty{color:#888;font-style:italic}' +
    'body.dense{font-size:8.3pt;line-height:1.22}.dense td{padding:6pt 9pt}.dense .masthead td{padding:0 0 9pt}.dense .brand{font-size:23pt}.dense .hero{padding:10pt 0}.dense h1{font-size:23pt}.dense .company{padding:8pt 10pt;margin-bottom:8pt}.dense .company .answer{font-size:10pt}.dense .section{margin:9pt 0 5pt}.dense .standards{padding:7pt 9pt}.dense .notice{font-size:7pt;margin:9pt 0 8pt}' +
    'body{font-size:9.5pt;line-height:1.4}td{border:0;border-bottom:.5pt solid #d9d3c8;padding:10pt 0}td+td{padding-left:17pt}.masthead td{padding:0 0 12pt;border:0}.masthead td+td{padding-left:0}.brand{font-size:27pt}.tag{font-size:9.5pt}.hero{padding:14pt 0 13pt}h1{font-size:28pt;letter-spacing:-.4pt}.company{padding:10pt 12pt;margin-bottom:10pt;background:#f4f1eb}.company .answer{font-size:13pt}.label{font-size:7pt;letter-spacing:.45pt;color:#7b6743}.section{font:italic 14pt Georgia,serif;text-transform:none;letter-spacing:0;color:#222;border-bottom:.6pt solid #b99b60;padding-bottom:4pt;margin:12pt 0 3pt}.commercial td{background:#f3f0e9;padding:10pt 11pt;border-bottom:0}.commercial td+td{border-left:7pt solid white}.standards{border:0;border-bottom:.5pt solid #d9d3c8;padding:8pt 0 10pt}.notice{margin:11pt 0 9pt;color:#77726a}.footer{font-size:7.5pt;line-height:1.45}.dense td{padding:6pt 0}.dense td+td{padding-left:14pt}.dense .masthead td{padding:0 0 8pt}.dense .masthead td+td{padding-left:0}.dense .section{font-size:12pt;margin:8pt 0 3pt}.dense .commercial td{padding:7pt 9pt}.dense .standards{padding:7pt 0}.dense .company .answer{font-size:10pt}' +
    '</style></head><body class="' + (dense ? 'dense' : '') + '"><table class="masthead"><tr><td><p class="brand">BostonLuxLimos</p><p class="tag">Elegance in every journey</p></td><td class="right"><p class="kicker">Worldwide affiliations</p><p class="ref">' + esc(reference) + '</p></td></tr></table>' +
    '<div class="hero"><h1>Partner Introduction</h1><p class="sub">Different destinations. The same personal attention.</p></div>' +
    '<div class="company"><p class="label">Company & website</p><p class="answer">' + company + '</p></div>' +
    '<p class="section">Company profile</p><table>' + profile + '</table>' +
    '<p class="section">Commercial details</p><table class="commercial">' + commercial + '</table>' +
    '<p class="section">Service standards & discretion</p><div class="standards"><p class="answer">' + standards + '</p></div>' +
    '<p class="notice">For review only. Information is supplied by the applicant and has not been independently verified. This introduction does not approve an affiliation, accept rates or confirm a journey.' +
    (shortened ? ' Longer answers are excerpted; the complete submission remains in the private inquiry record.' : '') + '</p>' +
    '<div class="footer">Operated by STRUT2 LLC &nbsp; | &nbsp; bostonluxlimos.com<br>Received: ' + esc(data[0]) + ' &nbsp; | &nbsp; <a href="' + source + '">Open complete private inquiry</a></div></body></html>';
}

function previewRefinedAffiliationPdf() {
 const sheet = SpreadsheetApp.openById(BLL_PDF.spreadsheetId).getSheetById(BLL_PDF.sheetId);
 const data = sheet.getRange(4,1,1,BLL_PDF.responseColumns).getDisplayValues()[0];
 if (!/^TEST/.test(data[2] || '')) throw new Error('Only internal QA records may be previewed.');
 const long = Array(12).fill('LAYOUT TEST ONLY - A longer response verifies the refined single-page document. '.repeat(30));
 long[0]='November 1, 2026';long[1]='layout-test@example.invalid';long[2]='LAYOUT TEST ONLY - long-response verification';
 [[data,'Refined Partner Introduction'],[long,'Refined Long Response Test']].forEach(pair => {
  const pdf=HtmlService.createHtmlOutput(affiliationPdfRefinedHtml_('DESIGN REVIEW',[],pair[0])).getAs(MimeType.PDF).setName('BostonLuxLimos - '+pair[1]+'.pdf');
  const file=affiliationPdfFolder_().createFile(pdf);console.log(pair[1]+': '+file.getUrl());
 });
}
