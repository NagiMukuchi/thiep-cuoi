/**
 * Sổ lưu bút cho thiệp cưới — dán toàn bộ file này vào Apps Script của Google Sheet.
 *
 * - doGet  : trả về danh sách lời chúc (chỉ tên, lời chúc, thời gian, tham dự) cho trang thiệp
 * - doPost : nhận lời chúc mới từ trang thiệp và ghi thêm 1 dòng vào sheet "LoiChuc"
 *
 * Muốn ẩn một lời chúc: gõ chữ x vào cột "Ẩn" của dòng đó (hoặc xóa luôn dòng).
 * Cột "Số người" chỉ bạn thấy trong sheet, không hiện trên thiệp.
 */

const SHEET_NAME = 'LoiChuc';
const HEADERS = ['Thời gian', 'Tên', 'Tham dự', 'Số người', 'Lời chúc', 'Ẩn (gõ x để ẩn)'];
const ATTEND_LABELS = ['Sẽ đến', 'Đang sắp xếp', 'Không thể đến'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setColumnWidth(5, 420);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Chặn chữ bắt đầu bằng = + - @ để không bị Google Sheets hiểu là công thức
function clean_(value, maxLen) {
  let s = String(value == null ? '' : value).replace(/\s+/g, ' ').trim().slice(0, maxLen);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function doGet() {
  const rows = getSheet_().getDataRange().getValues().slice(1);
  const wishes = rows
    .filter(r => r[1] && r[4] && String(r[5]).trim() === '')
    .map(r => ({
      time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      name: String(r[1]).replace(/^'/, ''),
      attend: String(r[2]),
      message: String(r[4]).replace(/^'/, ''),
    }))
    .reverse()          // mới nhất lên đầu
    .slice(0, 300);
  return json_({ ok: true, wishes: wishes });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.website) return json_({ ok: true });            // bẫy chống spam bot
    const name = clean_(d.name, 60);
    const message = clean_(d.message, 500);
    if (!name || !message) return json_({ ok: false, error: 'Thiếu tên hoặc lời chúc' });
    const attend = ATTEND_LABELS.indexOf(d.attend) >= 0 ? d.attend : '';
    const guests = String(d.guests || '').replace(/[^0-9+]/g, '').slice(0, 3);
    getSheet_().appendRow([new Date(), name, attend, guests, message, '']);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
