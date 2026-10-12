// ============================================================
// THRYVE PULSE — UPDATED GOOGLE APPS SCRIPT BACKEND
// Preserves existing registration email template and Results/admin
// email workflow; adds verified ₹99 payment fields to registrations.
// Deploy as Web App: Execute as Me; access: Anyone.
// ============================================================

// ============================================================

//  THRYVE PULSE - Google Apps Script (Backend)

//  Paste into Apps Script editor

//  Deploy as Web App -> Execute as: Me -> Anyone can access

//

//  IMPORTANT: Before deploying, run testGmailAuth() ONCE from

//  the editor to grant Gmail permission. Without this step,

//  emails silently fail.

// ============================================================



var ADMIN_EMAIL   = 'thryvemeeraki@gmail.com';

var SHEET_NAME    = 'Registrations';

var RESULTS_SHEET = 'Results';



// ============================================================

//  RUN THIS FIRST — authorizes Gmail and tests email sending

//  In the editor: select testGmailAuth -> click Run

// ============================================================

function testGmailAuth() {

  try {

    GmailApp.sendEmail(

      ADMIN_EMAIL,

      'Thryve Pulse - Gmail Auth Test',

      'Gmail is authorized and working correctly. Registration emails will now be delivered.'

    );

    Logger.log('SUCCESS: Test email sent to ' + ADMIN_EMAIL);

  } catch(err) {

    Logger.log('FAILED: ' + err.toString());

  }

}



// ------------------------------------------------------------

//  ENTRY POINT

// ------------------------------------------------------------

// ------------------------------------------------------------
// ENTRY POINT — supports existing form posts and the paid flow
// ------------------------------------------------------------
function doPost(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  Logger.log('doPost received. Parameters: ' + JSON.stringify(params));
  var payload = {};
  try {
    if (params.payload) payload = JSON.parse(params.payload);
    else if (e && e.postData && e.postData.contents) payload = JSON.parse(e.postData.contents);
    else throw new Error('Missing payload.');
  } catch (parseErr) {
    Logger.log('JSON parse error: ' + parseErr);
    return ContentService.createTextOutput('ERROR: invalid_payload').setMimeType(ContentService.MimeType.TEXT);
  }
  Logger.log('Action: ' + (payload.action || 'none') + ' | Email: ' + (payload.email || payload.student_email || 'none'));
  try {
    if (payload.action === 'admin') {
      handleAdminResult(payload);
      return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
    }
    if (payload.action === 'register') {
      handleRegistration(payload);
      return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
    }
    if (payload.action === 'redeemCashCoupon') {
      var couponResult = handleCashCouponRegistration(payload);
      return ContentService.createTextOutput(JSON.stringify(couponResult)).setMimeType(ContentService.MimeType.JSON);
    }
    throw new Error('Unsupported action: ' + payload.action);
  } catch (err) {
    Logger.log('ERROR in handler: ' + err.toString() + ' | Stack: ' + (err.stack || ''));
    if (payload.action === 'redeemCashCoupon') {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.publicMessage || 'Coupon redemption failed. Please check the code or contact support.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput('ERROR: handler_failed').setMimeType(ContentService.MimeType.TEXT);
  }
}


function doGet(e) {

  var action   = e.parameter.action   || '';

  var callback = e.parameter.callback || 'cb';

  Logger.log('doGet action: ' + action);

  try {

    if (action === 'admin') { handleAdminResult(e.parameter); }

  } catch (err) {

    Logger.log('doGet error: ' + err.toString());

  }

  return ContentService

    .createTextOutput(callback + '({"status":"ok"})')

    .setMimeType(ContentService.MimeType.JAVASCRIPT);

}



// ------------------------------------------------------------
// STEP 1 — Paid registration: Sheet + email assessment link
// ------------------------------------------------------------
function handleRegistration(data) {
  Logger.log('handleRegistration for: ' + (data.email || 'unknown'));
  var paymentId = String(data.razorpayPaymentId || '').trim();
  if (String(data.paymentStatus || '').toLowerCase() !== 'paid' || !paymentId ||
      Number(data.amountInr) !== 99 || !data.razorpayOrderId) {
    throw new Error('Registration rejected: verified ₹99 payment fields are missing.');
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME, ['Timestamp','Name','Email','Phone','Age Category','Link','Link Expiry']);
  ensurePaymentHeaders(sheet);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var paymentIdCol = headers.indexOf('Razorpay Payment ID') + 1;
  if (paymentIdCol > 0 && sheet.getLastRow() > 1) {
    var existing = sheet.getRange(2, paymentIdCol, sheet.getLastRow() - 1, 1).createTextFinder(paymentId).matchEntireCell(true).findNext();
    if (existing) { Logger.log('Duplicate payment callback ignored: ' + paymentId); return; }
  }
  var values = {
    'Timestamp': new Date(), 'Name': data.name || '', 'Email': data.email || '',
    'Phone': data.phone || '', 'Age Category': data.category || '', 'Link': data.link || '',
    'Link Expiry': data.expTime || '', 'Payment Status': 'Paid', 'Payment Method': 'Razorpay',
    'Coupon Code': '', 'Razorpay Order ID': data.razorpayOrderId || '',
    'Razorpay Payment ID': paymentId, 'Amount (INR)': 99
  };
  sheet.appendRow(headers.map(function(header) { return Object.prototype.hasOwnProperty.call(values, header) ? values[header] : ''; }));
  sendRegistrationEmail(data);
  Logger.log('Paid registration row added and assessment email sent to: ' + data.email);
}

function handleCashCouponRegistration(data) {
  var code = String(data.couponCode || '').trim().toUpperCase();
  var name = String(data.name || '').trim();
  var email = String(data.email || '').trim().toLowerCase();
  var phone = String(data.phone || '').trim();
  if (!/^TP-CASH-[A-Z0-9]{12}$/.test(code) || !name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || !phone || !data.link || Number(data.amountInr) !== 99) {
    var invalid = new Error('Invalid coupon registration details.'); invalid.publicMessage = 'Please check your details and coupon code.'; throw invalid;
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var registered = false;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var couponSheet = getOrCreateSheet(ss, 'Cash Coupons', ['Coupon Code','Status','Created At','Redeemed At','Redeemed By Name','Redeemed By Email','Registration Link','Payment Method']);
    ensureCouponHeaders(couponSheet);
    var couponHeaders = couponSheet.getRange(1, 1, 1, couponSheet.getLastColumn()).getDisplayValues()[0];
    var codeCol = couponHeaders.indexOf('Coupon Code') + 1;
    var statusCol = couponHeaders.indexOf('Status') + 1;
    var found = null;
    if (couponSheet.getLastRow() > 1) {
      var foundCell = couponSheet.getRange(2, codeCol, couponSheet.getLastRow() - 1, 1).createTextFinder(code).matchEntireCell(true).matchCase(false).findNext();
      if (foundCell) found = foundCell.getRow();
    }
    if (!found) {
      var notFound = new Error('Coupon not found.'); notFound.publicMessage = 'This coupon code is not valid. Please check it or contact Thryve Pulse support.'; throw notFound;
    }
    var status = String(couponSheet.getRange(found, statusCol).getDisplayValue() || '').trim().toLowerCase();
    if (status !== 'active' && status !== 'available' && status !== 'issued') {
      var used = new Error('Coupon already redeemed or inactive.'); used.publicMessage = 'This coupon has already been used or is inactive. Please contact Thryve Pulse support if you believe this is a mistake.'; throw used;
    }

    // Redeem and create registration while holding the script lock so the same code cannot be used twice.
    var cvals = {
      'Status': 'Redeemed', 'Redeemed At': new Date(), 'Redeemed By Name': name,
      'Redeemed By Email': email, 'Registration Link': data.link, 'Payment Method': 'Cash Coupon'
    };
    couponHeaders.forEach(function(header, i) {
      if (Object.prototype.hasOwnProperty.call(cvals, header)) couponSheet.getRange(found, i + 1).setValue(cvals[header]);
    });

    var regSheet = getOrCreateSheet(ss, SHEET_NAME, ['Timestamp','Name','Email','Phone','Age Category','Link','Link Expiry']);
    ensurePaymentHeaders(regSheet);
    var headers = regSheet.getRange(1, 1, 1, regSheet.getLastColumn()).getDisplayValues()[0];
    var values = {
      'Timestamp': new Date(), 'Name': name, 'Email': email, 'Phone': phone,
      'Age Category': data.category || '', 'Link': data.link, 'Link Expiry': data.expTime || '',
      'Payment Status': 'Cash Paid', 'Payment Method': 'Cash Coupon', 'Coupon Code': code,
      'Razorpay Order ID': '', 'Razorpay Payment ID': '', 'Amount (INR)': 99
    };
    regSheet.appendRow(headers.map(function(header) { return Object.prototype.hasOwnProperty.call(values, header) ? values[header] : ''; }));
    registered = true;
  } finally {
    lock.releaseLock();
  }
  if (registered) sendRegistrationEmail(data);
  return { ok: true, email: email };
}

function ensurePaymentHeaders(sheet) {
  var required = ['Timestamp','Name','Email','Phone','Age Category','Link','Link Expiry','Payment Status','Payment Method','Coupon Code','Razorpay Order ID','Razorpay Payment ID','Amount (INR)'];
  if (sheet.getLastRow() === 0 || (sheet.getLastRow() === 1 && sheet.getLastColumn() === 1 && !String(sheet.getRange(1, 1).getDisplayValue()).trim())) {
    sheet.getRange(1, 1, 1, required.length).setValues([required]);
    sheet.getRange(1, 1, 1, required.length).setFontWeight('bold').setBackground('#1A1A2E').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1); return;
  }
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var current = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  required.forEach(function(header) {
    if (current.indexOf(header) === -1) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(header).setFontWeight('bold').setBackground('#1A1A2E').setFontColor('#FFFFFF');
      current.push(header);
    }
  });
  sheet.setFrozenRows(1);
}

function ensureCouponHeaders(sheet) {
  var required = ['Coupon Code','Status','Created At','Redeemed At','Redeemed By Name','Redeemed By Email','Registration Link','Payment Method'];
  if (sheet.getLastRow() === 0 || (sheet.getLastRow() === 1 && sheet.getLastColumn() === 1 && !String(sheet.getRange(1, 1).getDisplayValue()).trim())) {
    sheet.getRange(1, 1, 1, required.length).setValues([required]);
    sheet.getRange(1, 1, 1, required.length).setFontWeight('bold').setBackground('#1A1A2E').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1); return;
  }
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var current = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  required.forEach(function(header) {
    if (current.indexOf(header) === -1) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(header).setFontWeight('bold').setBackground('#1A1A2E').setFontColor('#FFFFFF');
      current.push(header);
    }
  });
  sheet.setFrozenRows(1);
}

// Run manually from the Apps Script editor to create a batch of unique codes.
// Only issue a code to a customer after the cash payment has actually been received.
function generateCashCoupons(count) {
  count = Math.max(1, Math.min(Number(count) || 50, 500));
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, 'Cash Coupons', ['Coupon Code','Status','Created At','Redeemed At','Redeemed By Name','Redeemed By Email','Registration Link','Payment Method']);
  ensureCouponHeaders(sheet);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var codeCol = headers.indexOf('Coupon Code') + 1;
  var existing = {};
  if (sheet.getLastRow() > 1) sheet.getRange(2, codeCol, sheet.getLastRow() - 1, 1).getDisplayValues().forEach(function(r) { existing[String(r[0]).toUpperCase()] = true; });
  var rows = [];
  while (rows.length < count) {
    var code = 'TP-CASH-' + Utilities.getUuid().replace(/-/g, '').substring(0, 12).toUpperCase();
    if (existing[code]) continue;
    existing[code] = true;
    var values = {'Coupon Code':code,'Status':'Active','Created At':new Date(),'Redeemed At':'','Redeemed By Name':'','Redeemed By Email':'','Registration Link':'','Payment Method':'Cash Coupon'};
    rows.push(headers.map(function(header) { return Object.prototype.hasOwnProperty.call(values, header) ? values[header] : ''; }));
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  return rows.length + ' unique cash coupons created in the Cash Coupons sheet.';
}

function sendRegistrationEmail(data) {

  var name    = data.name    || 'Friend';

  var expTime = data.expTime || '1 hour from now';

  var link    = data.link    || '#';

  var subject = 'Your Thryve Pulse Assessment Link is Ready';



  var html =

  '\<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http\://www\.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">' +

  '\<html xmlns="http\://www\.w3.org/1999/xhtml">\<head>' +

  '\<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />' +

  '\<meta name="viewport" content="width=device-width, initial-scale=1.0"/>' +

  '\</head>\<body style="margin:0;padding:0;background-color:#ECEAE4;">' +

  '\<table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color:#ECEAE4;">' +

  '\<tr>\<td align="center" style="padding:32px 16px;">' +

  '\<table border="0" cellpadding="0" cellspacing="0" width="560" style="max-width:560px;background-color:#FAF8F3;border-radius:20px;">' +

  '\<tr>\<td style="background-color:#2D6A4F;height:6px;border-radius:20px 20px 0 0;font-size:0;line-height:0;">&nbsp;\</td>\</tr>' +

  '\<tr>\<td style="background-color:#1A1A2E;padding:44px 48px 36px;text-align:center;">' +

  '\<table border="0" cellpadding="0" cellspacing="0" width="100%">\<tr>\<td align="center">' +

  '\<table border="0" cellpadding="0" cellspacing="0">\<tr>' +

  '\<td style="background-color:#2D6A4F;border-radius:30px;padding:5px 18px;">' +

  '\<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:10px;font-weight:bold;letter-spacing:3px;color:#A8DFC9;text-transform:uppercase;">THRYVE PULSE ASSESSMENT\</p>' +

  '\</td>\</tr>\</table>' +

  '\<p style="margin:20px 0 6px;font-family:Georgia,serif;font-size:34px;font-weight:bold;color:#FFFFFF;line-height:1.2;">Discover Your\</p>' +

  '\<p style="margin:0 0 24px;font-family:Georgia,serif;font-size:34px;font-weight:bold;color:#D4A853;line-height:1.2;">True Self\</p>' +

  '\<table border="0" cellpadding="0" cellspacing="0">\<tr>' +

  '\<td style="border:1px solid rgba(255,255,255,0.2);border-radius:30px;padding:10px 26px;">' +

  '\<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#FFFFFF;">Hello, \<strong>' + name + '\</strong>\</p>' +

  '\</td>\</tr>\</table>' +

  '\</td>\</tr>\</table>' +

  '\</td>\</tr>' +

  '\<tr>\<td style="padding:40px 48px 32px;">' +

  '\<table border="0" cellpadding="0" cellspacing="0" width="100%">' +

  '\<tr>\<td style="padding-bottom:28px;">' +

  '\<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:15px;color:#3A3A3A;line-height:1.8;">Your personal Thryve Pulse assessment link is ready. Answer 18 scenario-based questions to reveal your unique archetype and unlock insights about how you think, lead, and connect.\</p>' +

  '\</td>\</tr>' +

  '\<tr>\<td style="padding-bottom:28px;">' +

  '\<table border="0" cellpadding="0" cellspacing="0" width="100%">' +

  '\<tr>\<td style="background-color:#FFFBEF;border-left:4px solid #D4A853;border-radius:0 10px 10px 0;padding:16px 20px;">' +

  '\<p style="margin:0 0 4px;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:bold;color:#7A5C00;">LINK EXPIRES AT\</p>' +

  '\<p style="margin:0 0 6px;font-family:Arial,Helvetica,sans-serif;font-size:15px;font-weight:bold;color:#5A4000;">' + expTime + '\</p>' +

  '\<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#9A8040;">This link is personal to you. Please do not share it with anyone.\</p>' +

  '\</td>\</tr>\</table>' +

  '\</td>\</tr>' +

  '\<tr>\<td style="padding-bottom:36px;" align="center">' +

  '\<table border="0" cellpadding="0" cellspacing="0">\<tr>' +

  '\<td style="background-color:#1A1A2E;border-radius:50px;padding:0;">' +

  '\<a href="' + link + '" style="display:block;font-family:Arial,Helvetica,sans-serif;font-size:16px;font-weight:bold;color:#FFFFFF;text-decoration:none;padding:18px 56px;letter-spacing:0.5px;">Begin My Assessment\</a>' +

  '\</td>\</tr>\</table>' +

  '\</td>\</tr>' +

  '\<tr>\<td style="padding-bottom:28px;">' +

  '\<table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color:#FFFFFF;border-radius:14px;border:1px solid #E4E0D8;">' +

  '\<tr>\<td style="padding:24px 26px 20px;">' +

  '\<p style="margin:0 0 20px;font-family:Arial,Helvetica,sans-serif;font-size:10px;font-weight:bold;letter-spacing:2px;color:#9A9A9A;text-transform:uppercase;border-bottom:1px solid #EEEBE4;padding-bottom:12px;">HOW IT WORKS\</p>' +

  stepRow(1, 'Open this email and click Begin My Assessment') +

  stepRow(2, 'Answer 18 scenario-based questions honestly') +

  stepRow(3, 'Your personal archetype is revealed on screen') +

  stepRow(4, 'A full report is sent to your coach or admin') +

  '\</td>\</tr>\</table>' +

  '\</td>\</tr>' +

  '\<tr>\<td>' +

  '\<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#9A9A9A;line-height:1.7;">If the button above does not work, copy and paste this URL into your browser:\</p>' +

  '\<p style="margin:6px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#2D6A4F;word-break:break-all;">' + link + '\</p>' +

  '\</td>\</tr>' +

  '\</table>' +

  '\</td>\</tr>' +

  '\<tr>\<td style="background-color:#1A1A2E;border-radius:0 0 20px 20px;padding:20px 48px;text-align:center;">' +

  '\<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6A7A8A;">Thryve Pulse &nbsp;|&nbsp; Powered by Meeraki\</p>' +

  '\</td>\</tr>' +

  '\</table>' +

  '\</td>\</tr>\</table>' +

  '\</body>\</html>';



  GmailApp.sendEmail(data.email, subject, 'Your Thryve Pulse assessment link: ' + link, {

    htmlBody: html,

    name: 'Thryve Pulse'

  });

}



function stepRow(n, text) {

  return '\<table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:14px;">' +

    '\<tr>' +

    '\<td width="36" valign="top">' +

    '\<table border="0" cellpadding="0" cellspacing="0">\<tr>' +

    '\<td style="background-color:#1A1A2E;width:28px;height:28px;border-radius:14px;text-align:center;vertical-align:middle;">' +

    '\<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:12px;font-weight:bold;color:#FFFFFF;line-height:28px;">' + n + '\</p>' +

    '\</td>\</tr>\</table>' +

    '\</td>' +

    '\<td valign="middle" style="padding-left:12px;padding-top:4px;">' +

    '\<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#3A3A3A;line-height:1.6;">' + text + '\</p>' +

    '\</td>' +

    '\</tr>' +

    '\</table>';

}



// ------------------------------------------------------------

//  STEP 2 - Assessment Result: Sheet + Email admin

// ------------------------------------------------------------

function handleAdminResult(data) {

  Logger.log('handleAdminResult for: ' + (data.student_name || 'unknown'));

  var ss    = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = getOrCreateSheet(ss, RESULTS_SHEET, [

    'Completed At','Name','Email','Phone','Age','Archetype','Classification',

    'Archetype Blend','Keywords','Core Orientation','Strengths','Growth Areas','Pathways','Breakdown'

  ]);

  sheet.appendRow([

    data.completed_at  || new Date().toLocaleString(),

    data.student_name  || '',

    data.student_email || '',

    data.student_phone || '',

    data.student_age   || '',

    data.archetype     || '',

    data.classification|| '',

    data.blend         || '',

    data.keywords      || '',

    data.core          || '',

    data.strengths     || '',

    data.growth        || '',

    data.pathways      || '',

    data.breakdown     || ''

  ]);

  sendAdminResultEmail(data);

}



function sendAdminResultEmail(d) {

  var name    = d.student_name || 'Student';

  var arch    = (d.archetype || 'Unknown').replace(/[^\x00-\x7F]/g, '').trim();

  var subject = 'Thryve Result: ' + name + ' - ' + arch;

  var color   = (d.color && /^#[0-9A-Fa-f]{6}$/.test(d.color)) ? d.color : '#1A1A2E';



  var strengths = (d.strengths || '').split(' | ');

  var growth    = (d.growth    || '').split(' | ');

  var pathways  = (d.pathways  || '').split(' | ');

  var blend     = (d.blend     || '').split(' | ');



  var html =

  '\<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http\://www\.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">' +

  '\<html xmlns="http\://www\.w3.org/1999/xhtml">\<head>' +

  '\<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />' +

  '\</head>\<body style="margin:0;padding:0;background-color:#ECEAE4;">' +

  '\<table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color:#ECEAE4;">' +

  '\<tr>\<td align="center" style="padding:32px 16px;">' +

  '\<table border="0" cellpadding="0" cellspacing="0" width="600" style="max-width:600px;background-color:#FAF8F3;border-radius:20px;">' +

  '\<tr>\<td style="background-color:' + color + ';height:6px;border-radius:20px 20px 0 0;font-size:0;line-height:0;">&nbsp;\</td>\</tr>' +

  '\<tr>\<td style="background-color:#1A1A2E;padding:36px 48px;text-align:center;">' +

  '\<p style="margin:0 0 10px;font-family:Arial,Helvetica,sans-serif;font-size:10px;font-weight:bold;letter-spacing:3px;color:#52B788;text-transform:uppercase;">THRYVE PULSE - RESULTS REPORT\</p>' +

  '\<p style="margin:0 0 6px;font-family:Georgia,serif;font-size:28px;font-weight:bold;color:#FFFFFF;">' + name + '\</p>' +

  '\<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6A8A7A;">Completed: ' + (d.completed_at || '') + '\</p>' +

  '\</td>\</tr>' +

  '\<tr>\<td style="background-color:' + color + ';padding:32px 48px;text-align:center;">' +

  '\<p style="margin:0 0 6px;font-family:Arial,Helvetica,sans-serif;font-size:10px;font-weight:bold;letter-spacing:3px;color:rgba(255,255,255,0.7);text-transform:uppercase;">ARCHETYPE\</p>' +

  '\<p style="margin:0 0 8px;font-family:Georgia,serif;font-size:38px;font-weight:bold;color:#FFFFFF;">' + arch + '\</p>' +

  '\<p style="margin:0 0 16px;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:rgba(255,255,255,0.85);font-style:italic;">' + (d.archetype_tag || '') + '\</p>' +

  '\<table border="0" cellpadding="0" cellspacing="0" align="center">\<tr>' +

  '\<td style="background-color:rgba(255,255,255,0.2);border-radius:30px;padding:8px 24px;">' +

  '\<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:bold;color:#FFFFFF;">' + (d.classification || '') + '\</p>' +

  '\</td>\</tr>\</table>' +

  '\</td>\</tr>' +

  '\<tr>\<td style="padding:36px 48px;">' +

  '\<table border="0" cellpadding="0" cellspacing="0" width="100%">' +

  adminSection('STUDENT DETAILS',

    '\<table border="0" cellpadding="0" cellspacing="0" width="100%">' +

    adminDetailRow('Email',    d.student_email || '') +

    adminDetailRow('Phone',    d.student_phone || '') +

    adminDetailRow('Category', d.student_age   || '') +

    '\</table>'

  ) +

  adminSection('ARCHETYPE BLEND', blendBars(blend)) +

  adminSection('CORE ORIENTATION',

    '\<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#3A3A3A;line-height:1.8;">' + (d.core || '') + '\</p>'

  ) +

  adminSection('KEYWORDS',

    '\<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#2D6A4F;font-weight:bold;">' + (d.keywords || '') + '\</p>'

  ) +

  adminSection('STRENGTHS',    bulletList(strengths, '#2D6A4F')) +

  adminSection('GROWTH AREAS', bulletList(growth,    '#C0652B')) +

  adminSection('PATHWAYS',     bulletList(pathways,   '#2C3E8C')) +

  adminSection('ANSWER BREAKDOWN',

    '\<table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color:#F4F2EC;border-radius:10px;">\<tr>\<td style="padding:14px 18px;">' +

    '\<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#3A3A3A;line-height:1.8;">' + (d.breakdown || '') + '\</p>' +

    '\</td>\</tr>\</table>'

  ) +

  '\</table>' +

  '\</td>\</tr>' +

  '\<tr>\<td style="background-color:#1A1A2E;border-radius:0 0 20px 20px;padding:20px 48px;text-align:center;">' +

  '\<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6A7A8A;">Thryve Pulse &nbsp;|&nbsp; Admin Report &nbsp;|&nbsp; ' + ADMIN_EMAIL + '\</p>' +

  '\</td>\</tr>' +

  '\</table>\</td>\</tr>\</table>\</body>\</html>';



  GmailApp.sendEmail(ADMIN_EMAIL, subject, 'New Thryve Pulse result for ' + name, {

    htmlBody: html,

    name: 'Thryve Pulse System'

  });

}



function adminSection(title, content) {

  return '\<tr>\<td style="padding-bottom:26px;">' +

    '\<p style="margin:0 0 10px;font-family:Arial,Helvetica,sans-serif;font-size:10px;font-weight:bold;letter-spacing:2px;color:#9A9A9A;text-transform:uppercase;border-bottom:2px solid #EEEBE4;padding-bottom:8px;">' + title + '\</p>' +

    content + '\</td>\</tr>';

}



function adminDetailRow(label, value) {

  return '\<tr>' +

    '\<td width="100" valign="top" style="padding:5px 14px 5px 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:bold;color:#9A9A9A;">' + label + '\</td>' +

    '\<td style="padding:5px 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#3A3A3A;">' + value + '\</td>' +

    '\</tr>';

}



function bulletList(items, dotColor) {

  var dc = dotColor || '#2D6A4F';

  return items.map(function(item) {

    return '\<table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:10px;">\<tr>' +

      '\<td width="20" valign="top" style="padding-top:5px;">' +

      '\<table border="0" cellpadding="0" cellspacing="0">\<tr>' +

      '\<td style="background-color:' + dc + ';width:8px;height:8px;border-radius:4px;">&nbsp;\</td>' +

      '\</tr>\</table>\</td>' +

      '\<td style="padding-left:8px;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#3A3A3A;line-height:1.6;">' + item.trim() + '\</td>' +

      '\</tr>\</table>';

  }).join('');

}



function blendBars(items) {

  var barColor = { 'Driver':'#E63946','Explorer':'#2D6A4F','Thinker':'#4361EE','Connector':'#E76F51' };

  return items.map(function(item) {

    var m = item.trim().match(/^(\w+)\s+(\d+)%$/);

    var label = m ? m[1] : item.trim();

    var pct   = m ? parseInt(m[2], 10) : 0;

    var color = barColor[label] || '#1A1A2E';

    if (pct < 1) pct = 1; if (pct > 100) pct = 100;

    return '\<table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:14px;">' +

      '\<tr>\<td>\<table border="0" cellpadding="0" cellspacing="0" width="100%">\<tr>' +

      '\<td style="font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:bold;color:#3A3A3A;">' + label + '\</td>' +

      '\<td align="right" style="font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:bold;color:#3A3A3A;">' + pct + '%\</td>' +

      '\</tr>\</table>\</td>\</tr>' +

      '\<tr>\<td style="padding-top:5px;">' +

      '\<table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color:#EEEBE4;border-radius:6px;">\<tr>' +

      '\<td width="' + pct + '%" style="background-color:' + color + ';border-radius:6px;height:8px;font-size:0;line-height:0;">&nbsp;\</td>' +

      '\<td style="font-size:0;line-height:0;">&nbsp;\</td>' +

      '\</tr>\</table>\</td>\</tr>\</table>';

  }).join('');

}



function getOrCreateSheet(ss, name, headers) {

  var sheet = ss.getSheetByName(name);

  if (!sheet) {

    sheet = ss.insertSheet(name);

    sheet.appendRow(headers);

    sheet.getRange(1, 1, 1, headers.length)

      .setFontWeight('bold').setBackground('#1A1A2E').setFontColor('#FFFFFF');

    sheet.setFrozenRows(1);

  }

  return sheet;

}