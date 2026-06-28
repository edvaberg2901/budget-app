const SHEET_NAME = 'Sheet1'; // Change if your sheet tab has a different name
const AUTH_PIN   = '3783';   // PIN required for add/delete operations

function doGet(e) {
  try {
    const action = e.parameter.action || 'get';
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    if (action === 'get') {
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) {
        return jsonResponse({ data: [] });
      }
      const rows = data.slice(1).map((row, i) => ({
        _row:     i + 2,
        manad:    formatMonth(row[0]),
        kr:       Number(row[1]) || 0,
        vem:      row[2] || '',
        kategori: row[3] || '',
        vad:      row[4] || ''
      })).filter(r => r.kr > 0 || r.vad);
      return jsonResponse({ data: rows });
    }

    if (action === 'verify') {
      if (AUTH_PIN && e.parameter.pin !== AUTH_PIN) {
        return jsonResponse({ error: 'Unauthorized' });
      }
      return jsonResponse({ success: true });
    }

    if (action === 'add') {
      if (AUTH_PIN && e.parameter.pin !== AUTH_PIN) {
        return jsonResponse({ error: 'Unauthorized' });
      }
      const manad    = e.parameter.manad    || currentMonth();
      const kr       = Number(e.parameter.kr) || 0;
      const vem      = e.parameter.vem      || '';
      const kategori = e.parameter.kategori || '';
      const vad      = e.parameter.vad      || '';
      sheet.appendRow([manad, kr, vem, kategori, vad]);
      return jsonResponse({ success: true });
    }

    if (action === 'delete') {
      if (AUTH_PIN && e.parameter.pin !== AUTH_PIN) {
        return jsonResponse({ error: 'Unauthorized' });
      }
      const row = parseInt(e.parameter.row, 10);
      if (row >= 2) sheet.deleteRow(row);
      return jsonResponse({ success: true });
    }

    return jsonResponse({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// Always return manad as YYYY-MM-DD string regardless of how Sheets stores the cell.
function formatMonth(value) {
  if (!value) return '';
  if (value instanceof Date) {
    const y  = value.getFullYear();
    const mo = String(value.getMonth() + 1).padStart(2, '0');
    const d  = String(value.getDate()).padStart(2, '0');
    return y + '-' + mo + '-' + d;
  }
  return String(value);
}

function currentMonth() {
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
