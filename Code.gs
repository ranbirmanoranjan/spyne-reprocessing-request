/**
 * Reprocessing Request - Apps Script backend
 * ---------------------------------------------------------------------------
 * Deploy this as a Web App (Deploy > New deployment > type: Web app,
 * execute as: Me, who has access: Anyone) and paste the resulting URL into
 * CONFIG.GAS_URL at the top of reprocessing-request.html.
 *
 * It expects to run bound to a Google Sheet. On first run it will create a
 * "Requests" tab with headers if one doesn't already exist.
 * ---------------------------------------------------------------------------
 */

const SHEET_NAME = 'Requests';
const HEADERS = [
  'Request ID', 'Timestamp', 'Requester Name', 'Requester Team', 'Reason',
  'Responsible Team', 'VIN Count', 'Criticality', 'Manager', 'Manager Approved',
  'Estimated Hours', 'Status',
];

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function pad(n, width) {
  n = String(n);
  while (n.length < width) n = '0' + n;
  return n;
}

/**
 * GET actions:
 *   ?action=nextId&date=YYYYMMDD
 *     -> { requestId: "RPR-YYYYMMDD-001" }
 *   ?action=checkDuplicate&requester=...&team=...&vinCount=...
 *     -> { duplicate: true/false, match: {...} | null }
 */
function doGet(e) {
  const action = e.parameter.action;
  const sheet = getOrCreateSheet();

  if (action === 'nextId') {
    const dateStr = e.parameter.date || '';
    const data = sheet.getDataRange().getValues();
    let count = 0;
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][0] || '');
      if (id.indexOf('RPR-' + dateStr + '-') === 0) count++;
    }
    return respond({ requestId: 'RPR-' + dateStr + '-' + pad(count + 1, 3) });
  }

  if (action === 'checkDuplicate') {
    const requester = (e.parameter.requester || '').toLowerCase();
    const team = (e.parameter.team || '').toLowerCase();
    const vinCount = String(e.parameter.vinCount || '');
    const data = sheet.getDataRange().getValues();
    const cutoff = new Date(Date.now() - 24 * 60 * 60 * 1000); // last 24h

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const rowTimestamp = new Date(row[1]);
      if (rowTimestamp < cutoff) continue;
      const rowRequester = String(row[2] || '').toLowerCase();
      const rowTeam = String(row[5] || '').toLowerCase();
      const rowVinCount = String(row[6] || '');
      if (rowRequester === requester && rowTeam === team && rowVinCount === vinCount) {
        return respond({
          duplicate: true,
          match: { requestId: row[0], timestamp: row[1] },
        });
      }
    }
    return respond({ duplicate: false, match: null });
  }

  return respond({ error: 'Unknown action: ' + action });
}

/**
 * POST body (JSON): the full request object built by the front-end, e.g.
 * {
 *   requestId, requesterName, requesterTeam, reason, responsibleTeam,
 *   vinCount, criticality, managerName, managerApproved, estimatedHours
 * }
 * Appends one row to the Requests sheet.
 */
function doPost(e) {
  const sheet = getOrCreateSheet();
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return respond({ success: false, error: 'Invalid JSON body' });
  }

  sheet.appendRow([
    data.requestId || '',
    new Date(),
    data.requesterName || '',
    data.requesterTeam || '',
    data.reason || '',
    data.responsibleTeam || '',
    data.vinCount || '',
    data.criticality || '',
    data.managerName || '',
    data.managerApproved ? 'Yes' : 'No',
    data.estimatedHours || '',
    'Submitted',
  ]);

  return respond({ success: true, requestId: data.requestId || '' });
}
