/**
 * Jomo Home Interiors — Enquiry System backend.
 *
 * SETUP:
 * 1. Go to https://sheets.google.com and create a new blank spreadsheet.
 *    Name it e.g. "Jomo Enquiries".
 * 2. In the sheet, go to Extensions > Apps Script.
 * 3. Delete any starter code in Code.gs and paste this entire file in its place.
 * 4. Click Deploy > New deployment.
 *    - Click the gear icon next to "Select type" and choose "Web app".
 *    - Description: anything, e.g. "Enquiry form endpoint".
 *    - Execute as: Me.
 *    - Who has access: Anyone.
 *    - Click Deploy, then "Authorize access" and approve (it's your own script
 *      on your own sheet, so this is safe).
 * 5. Copy the "Web app URL" it gives you (ends in /exec).
 * 6. Open enquiry-form.html, find the line near the top of the <script> that says:
 *        const SHEET_WEB_APP_URL = "PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE";
 *    and replace the placeholder text with the URL you copied.
 * 7. Save, re-upload enquiry-form.html to your site.
 *
 * HOW MATCHING WORKS: every submission carries a phone number. Each sheet
 * tab is searched for an existing row with that same phone number — if found,
 * that row is updated in place; if not, a new row is appended. This means:
 *   - A client filling in their details, then later clicking "Edit My Details"
 *     and resubmitting, updates the same row (same phone number).
 *   - Staff filling in Management details types the client's phone number;
 *     as long as it matches, it fills into that same client's row rather than
 *     creating a new one — no dependency on refreshing the page or browser
 *     storage.
 *   - A genuinely new phone number always creates a new row.
 * Vendor Walk-In and Candidate Entry work the same way, each in their own tab.
 *
 * If you ever change a field in enquiry-form.html, add the matching entry to
 * the relevant section's fieldMap below so it has somewhere to land.
 *
 * NOTE ON COLUMN ORDER: a tab's header row is only created once, the first
 * time that tab is empty. If you're updating this script on a tab that
 * already has data, the column order below won't rearrange existing columns
 * automatically — delete that tab entirely (right-click its tab → Delete)
 * and the next submission will recreate it fresh with the new order.
 */

const SECTION_CONFIG = {
  client: {
    sheetName: 'Enquiries',
    phoneHeader: 'Your Contact Number',
    // Every client-facing field together, in natural reading order, then
    // the staff-only Management fields at the very end.
    headers: [
      'Last Updated',
      'Your Name', 'Your Contact Number', 'WhatsApp Number', 'Your Email', 'Your City',
      'Property Type', 'Configuration', 'Configuration (Other)', 'Preferred Style',
      'Project Location', 'Site Status', 'Expected Handover Date',
      'Carpet Area (sq.ft)', 'Floorplan Available', 'Budget',
      'Scope of Work', 'Other Requirements',
      'Lead Source', 'Referral Name', 'Lead Source (Other)',
      // --- Management (staff-filled) ---
      'Handled By', 'Date of Visit', 'Next Follow-Up Date', 'Remarks'
    ],
    fieldMap: {
      clientName: 'Your Name',
      contactNumber: 'Your Contact Number',
      propertyType: 'Property Type',
      configuration: 'Configuration',
      configurationOtherText: 'Configuration (Other)',
      projectLocation: 'Project Location',
      possessionStatus: 'Site Status',
      possessionDate: 'Expected Handover Date',
      scopeOfWork: 'Scope of Work',
      scopeOtherText: 'Other Requirements',
      leadSource: 'Lead Source',
      referralName: 'Referral Name',
      leadSourceOtherText: 'Lead Source (Other)',
      budgetRange: 'Budget',
      whatsappNumber: 'WhatsApp Number',
      email: 'Your Email',
      currentCity: 'Your City',
      preferredStyle: 'Preferred Style',
      carpetArea: 'Carpet Area (sq.ft)',
      floorplanAvailable: 'Floorplan Available'
    }
  },
  management: {
    sheetName: 'Enquiries',
    phoneHeader: 'Your Contact Number',
    headers: null, // same tab as client, headers already created by it
    fieldMap: {
      handledBy: 'Handled By',
      walkInDate: 'Date of Visit',
      followUpDate: 'Next Follow-Up Date',
      remarks: 'Remarks'
    }
  },
  vendor: {
    sheetName: 'Vendor Walk-In',
    phoneHeader: 'Phone Number',
    headers: ['Last Updated', 'Name', 'Phone Number', 'Date of Visit', 'Brand Name', 'Handled By'],
    fieldMap: {
      vendorName: 'Name',
      vendorPhone: 'Phone Number',
      vendorVisitDate: 'Date of Visit',
      brandName: 'Brand Name',
      vendorHandledBy: 'Handled By'
    }
  },
  candidate: {
    sheetName: 'Candidate Entry',
    phoneHeader: 'Phone Number',
    headers: ['Last Updated', 'Name', 'Phone Number', 'Date of Visit', 'Position'],
    fieldMap: {
      candidateName: 'Name',
      candidatePhone: 'Phone Number',
      candidateVisitDate: 'Date of Visit',
      candidatePosition: 'Position'
    }
  }
};

function getSheet_(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  if (sheet.getLastRow() === 0 && headers) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Strips everything but digits, and — since Indian mobile numbers are always
// 10 digits — drops any country code prefix so "+91 98765-43210", "91 9876543210"
// and "9876543210" are all treated as the same number for matching purposes.
function normalizePhone_(v) {
  const digits = String(v == null ? '' : v).replace(/\D/g, '');
  return digits.length > 10 ? digits.slice(-10) : digits;
}

function findRow_(sheet, headerRow, phoneCol, normPhone) {
  if (phoneCol <= 0 || sheet.getLastRow() <= 1) return -1;
  const phones = sheet.getRange(2, phoneCol, sheet.getLastRow() - 1, 1)
    .getValues().flat().map(normalizePhone_);
  const idx = phones.indexOf(normPhone);
  return idx === -1 ? -1 : idx + 2;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const config = SECTION_CONFIG[data.section];
    if (!config) {
      return jsonOutput_({ status: 'error', message: 'unknown section' });
    }

    const rawPhone = String(data.phone || '').trim();
    const normPhone = normalizePhone_(rawPhone);
    if (!normPhone) {
      return jsonOutput_({ status: 'error', message: 'missing phone number' });
    }

    const sheet = getSheet_(config.sheetName, config.headers);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const phoneCol = headerRow.indexOf(config.phoneHeader) + 1;
    const updatedCol = headerRow.indexOf('Last Updated') + 1;

    let targetRow = findRow_(sheet, headerRow, phoneCol, normPhone);
    const isNewRow = targetRow === -1;
    if (isNewRow) {
      targetRow = sheet.getLastRow() + 1;
    }

    // Only write the phone number into its cell for a brand-new row — an
    // update (e.g. Management typing the client's number to attach their
    // notes) shouldn't overwrite the client's own original formatting.
    if (phoneCol > 0 && isNewRow) sheet.getRange(targetRow, phoneCol).setValue(rawPhone);
    if (updatedCol > 0) sheet.getRange(targetRow, updatedCol).setValue(new Date());

    const fields = data.fields || {};
    Object.keys(fields).forEach(function (key) {
      const headerName = config.fieldMap[key];
      if (!headerName) return;
      const colIndex = headerRow.indexOf(headerName) + 1;
      if (colIndex > 0) {
        let val = fields[key];
        if (Array.isArray(val)) val = val.join(', ');
        sheet.getRange(targetRow, colIndex).setValue(val);
      }
    });

    return jsonOutput_({ status: 'ok', row: targetRow, wasNew: isNewRow });
  } catch (err) {
    return jsonOutput_({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// GET .../exec?phone=9876543210 — searches Enquiries, then Vendor Walk-In,
// then Candidate Entry (in that order) for a row with a matching phone
// number, and returns every column for that row as one flat object.
function doGet(e) {
  const phone = e && e.parameter && e.parameter.phone;
  if (!phone) {
    return ContentService.createTextOutput('Jomo Home Interiors enquiry system endpoint is live.');
  }

  const normPhone = normalizePhone_(phone);
  const searchOrder = [
    { key: 'client', label: 'Client / Management' },
    { key: 'vendor', label: 'Vendor Walk-In' },
    { key: 'candidate', label: 'Candidate Entry' }
  ];

  for (let i = 0; i < searchOrder.length; i++) {
    const config = SECTION_CONFIG[searchOrder[i].key];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(config.sheetName);
    if (!sheet || sheet.getLastRow() <= 1) continue;

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const phoneCol = headerRow.indexOf(config.phoneHeader) + 1;
    const row = findRow_(sheet, headerRow, phoneCol, normPhone);
    if (row === -1) continue;

    const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const record = {};
    headerRow.forEach(function (h, idx) {
      if (!h) return;
      const v = values[idx];
      record[h] = (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd MMM yyyy') : v;
    });

    return jsonOutput_({ found: true, type: searchOrder[i].key, typeLabel: searchOrder[i].label, record: record });
  }

  return jsonOutput_({ found: false });
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
