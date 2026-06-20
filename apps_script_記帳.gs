// ════════════════════════════════════════════
//  瑞士記帳 2026 — Google Apps Script Web App
//  貼到 Apps Script 編輯器後部署為 Web App
// ════════════════════════════════════════════

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'getAll';
  let result;
  try {
    if      (action === 'getAll') result = getAll();
    else if (action === 'add')    result = add(JSON.parse(e.parameter.data));
    else if (action === 'delete') result = del(e.parameter.id);
    else result = { error: 'unknown action' };
  } catch(err) {
    result = { error: err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 取得全部使用者記帳（不含 prepaid） ──
function getAll() {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(r => {
    // r[5] is a Date object when Sheets auto-converts YYYY-MM-DD strings
    // Use try/catch since instanceof Date behaves unexpectedly in Apps Script V8
    let dateVal;
    try {
      dateVal = Utilities.formatDate(new Date(r[5]), 'Asia/Taipei', 'yyyy-MM-dd');
    } catch(e) {
      dateVal = String(r[5]).substring(0, 10);
    }
    return {
      id:      String(r[0]),
      cat:     String(r[1]),
      nt:      Number(r[2]),
      chf:     Number(r[3]),
      note:    String(r[4]),
      date:    dateVal,
      prepaid: false
    };
  });
}

// ── 新增一筆（防重複） ──
function add(data) {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      return { success: true, duplicate: true };
    }
  }
  sheet.appendRow([
    data.id, data.cat,
    data.nt  || 0,
    data.chf || 0,
    data.note || '',
    data.date,
    false
  ]);
  return { success: true };
}

// ── 刪除一筆 ──
function del(id) {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'not found' };
}

// ── 取得（或建立）Expenses 工作表 ──
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Expenses');
  if (!sheet) {
    sheet = ss.insertSheet('Expenses');
    sheet.appendRow(['id','cat','nt','chf','note','date','prepaid']);
  }
  return sheet;
}
