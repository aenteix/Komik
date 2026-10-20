// ============================================================
// Code.gs — Google Apps Script untuk Reading List Komik
// Cara pakai:
// 1. Buka https://sheets.new untuk buat Google Sheet baru
// 2. Menu: Ekstensi → Apps Script, hapus isi file, paste kode ini
// 3. Ganti TOKEN di bawah dengan kode rahasia buatanmu
// 4. Deploy → Deployment baru → Web app:
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Salin URL Web App ke reading-list.html (konstanta SYNC_URL)
// ============================================================

const SHEET_NAME = 'ReadingList';
const TOKEN = '1qjbFXDttyP1EqtG6oo-8WQ-_HGz7DY5Cx9RvBBXNgmRaWSdjvtE4_0FK'; // wajib diganti!

function doGet(e) {
  if (!e || e.parameter.token !== TOKEN) {
    return jsonOut({ error: 'unauthorized' });
  }
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues();
  const items = rows.slice(1)
    .filter(r => r[1])
    .map(r => ({
      id: Number(r[0]),
      title: String(r[1]),
      type: String(r[2]),
      chapter: Number(r[3]),
      status: String(r[4]),
      updated: String(r[5] || '')
    }));
  return jsonOut(items);
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return jsonOut({ error: 'bad json' }); }

  if (body.token !== TOKEN) {
    return jsonOut({ error: 'unauthorized' });
  }

  const sheet = getSheet_();
  const items = Array.isArray(body.items) ? body.items : [];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 6).setValues([['id', 'title', 'type', 'chapter', 'status', 'updated']]);

  if (items.length) {
    sheet.getRange(2, 1, items.length, 6).setValues(items.map(i => [
      i.id,
      String(i.title || ''),
      String(i.type || 'Komik'),
      Number(i.chapter) || 0,
      String(i.status || 'ongoing'),
      new Date().toLocaleString('id-ID')
    ]));
  }
  return jsonOut({ ok: true, count: items.length });
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 6).setValues([['id', 'title', 'type', 'chapter', 'status', 'updated']]);
  }
  return sheet;
}
