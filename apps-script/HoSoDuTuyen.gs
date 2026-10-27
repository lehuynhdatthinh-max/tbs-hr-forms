/**
 * HỒ SƠ DỰ TUYỂN (TBSL-NS-QT01-BM02) — máy chủ cho trang ho-so-du-tuyen/ trong repo tbs-hr-forms.
 *
 * Cài: tạo Google Sheet mới › Tiện ích mở rộng › Apps Script › dán file này › chạy setup() một lần
 * › Triển khai › Triển khai mới › Ứng dụng web (Thực thi: Tôi, Ai có quyền truy cập: Bất kỳ ai).
 * Xem apps-script/HUONG-DAN.md.
 *
 * - Mỗi hồ sơ là một dòng trong tab "Hồ sơ dự tuyển"; toàn bộ nội dung nằm ở cột "Dữ liệu (JSON)".
 * - Ảnh 3x4 và chữ ký lưu trong thư mục Drive riêng (chỉ chủ sheet xem được).
 * - Cột "Xem / in hồ sơ" mở bản in đúng bố cục mẫu Word (có phần VII để bộ phận tuyển dụng ký trên giấy).
 * - Nộp lại cùng submissionId (trang tự thử lại khi mất mạng) không tạo hồ sơ trùng.
 */
var HSDT = {
  VERSION: '1.1.0',
  SHEET: 'Hồ sơ dự tuyển',
  FOLDER_NAME: 'Hồ sơ dự tuyển - Ảnh & chữ ký',
  FOLDER_PROP: 'HSDT_FOLDER_ID',
  NOTIFY_PROP: 'HSDT_NOTIFY_EMAILS',
  // Trang quản lý HCNS (quan-ly/): đăng nhập bằng email + mật khẩu đặt trong Thuộc tính tập lệnh (không ghi vào mã).
  ADMIN_EMAILS_PROP: 'HSDT_ADMIN_EMAILS',
  ADMIN_PASSCODE_PROP: 'HSDT_ADMIN_PASSCODE',
  LOGIN_MAX_FAILS: 5,
  LOGIN_LOCK_SECONDS: 900,
  LIST_LIMIT: 1000,
  TIME_ZONE: 'Asia/Ho_Chi_Minh',
  LOGO_URL: 'https://lehuynhdatthinh-max.github.io/tbs-hr-forms/images/logo.png',
  MAX_IMAGE_BYTES: 1500000,
  MAX_JSON_IN_CELL: 45000
};

var HSDT_HEADERS = [
  'Thời gian nộp', 'Mã hồ sơ', 'Họ và tên', 'Ngày sinh', 'Giới tính', 'Vị trí ưu tiên 1', 'Vị trí ưu tiên 2',
  'Lương mong đợi', 'Ngày có thể nhận việc', 'Di động', 'Email', 'Số CMND/CCCD', 'Địa chỉ thường trú',
  'Từng làm ở TBS', 'Biết tin qua', 'Ảnh 3x4', 'Xem / in hồ sơ', 'Trạng thái', 'Ghi chú HCNS', 'ID', 'Dữ liệu (JSON)'
];

/* =================== Cài đặt =================== */

function setup() {
  var sheet = hsdtSheet_();
  var folder = hsdtFolder_();
  var message = 'Đã sẵn sàng. Tab: "' + sheet.getName() + '", thư mục ảnh: ' + folder.getUrl() +
    '\nBước tiếp: Triển khai › Triển khai mới › Ứng dụng web (Thực thi: Tôi; Truy cập: Bất kỳ ai), rồi dán URL /exec vào ho-so-du-tuyen/index.html.';
  console.log(message);
  return message;
}

/** Tùy chọn: gửi email báo có hồ sơ mới. Sửa danh sách email rồi chạy một lần (để trống = tắt). */
function setNotifyEmails() {
  var emails = ''; // ví dụ: 'tuyendung@tbslogistics.com.vn, hcns@tbslogistics.com.vn'
  PropertiesService.getScriptProperties().setProperty(HSDT.NOTIFY_PROP, emails);
  console.log(emails ? 'Sẽ báo hồ sơ mới tới: ' + emails : 'Đã tắt email báo hồ sơ mới.');
}

function hsdtSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HSDT.SHEET) || ss.insertSheet(HSDT.SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HSDT_HEADERS.length).setValues([HSDT_HEADERS]).setFontWeight('bold').setBackground('#E45620').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(HSDT_HEADERS.indexOf('Dữ liệu (JSON)') + 1, 120);
  } else {
    var current = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    var missing = HSDT_HEADERS.filter(function (h) { return current.indexOf(h) === -1; });
    if (missing.length) sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
  }
  return sheet;
}

function hsdtFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(HSDT.FOLDER_PROP);
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (error) {}
  }
  var folder = DriveApp.createFolder(HSDT.FOLDER_NAME);
  props.setProperty(HSDT.FOLDER_PROP, folder.getId());
  return folder;
}

function hsdtHeaderMap_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var map = {};
  headers.forEach(function (h, i) { if (h) map[h] = i + 1; });
  return { headers: headers, col: map };
}

/* =================== Web app =================== */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'view') return hsdtView_(p.id);
  return hsdtJson_({ ok: true, app: 'Hồ sơ dự tuyển', version: HSDT.VERSION });
}

function doPost(e) {
  var p = (e && e.parameter) || {};
  try {
    var data = JSON.parse(p.data || '{}');
    if (p.action === 'submit') return hsdtJson_(hsdtSubmit_(data));
    if (p.action === 'adminLogin') {
      hsdtVerifyAdmin_(data.email, data.passcode);
      return hsdtJson_({ ok: true });
    }
    if (p.action === 'listApplications') {
      hsdtVerifyAdmin_(data.email, data.passcode);
      return hsdtJson_({ ok: true, result: hsdtList_(data.filters || {}) });
    }
    return hsdtJson_({ ok: false, message: 'Yêu cầu không hợp lệ.' });
  } catch (error) {
    if (error && error.adminError) return hsdtJson_({ ok: false, message: error.message });
    console.error(error && error.stack || error);
    return hsdtJson_({ ok: false, retry: true, message: 'Máy chủ gặp lỗi, vui lòng thử lại. (' + (error && error.message || error) + ')' });
  }
}

function hsdtJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* =================== Trang quản lý HCNS =================== */

function hsdtAdminError_(message) {
  var error = new Error(message);
  error.adminError = true;
  return error;
}

/** Email nằm trong HSDT_ADMIN_EMAILS (cách nhau dấu phẩy) và đúng HSDT_ADMIN_PASSCODE. Sai 5 lần / 15 phút thì khóa email đó. */
function hsdtVerifyAdmin_(email, passcode) {
  var props = PropertiesService.getScriptProperties();
  var allowed = String(props.getProperty(HSDT.ADMIN_EMAILS_PROP) || '').toLowerCase().split(/[,;\s]+/).filter(Boolean);
  var secret = String(props.getProperty(HSDT.ADMIN_PASSCODE_PROP) || '');
  if (!allowed.length || !secret) throw hsdtAdminError_('Máy chủ hồ sơ dự tuyển chưa đặt tài khoản quản trị (HSDT_ADMIN_EMAILS / HSDT_ADMIN_PASSCODE).');
  var e = String(email || '').toLowerCase().trim();
  if (!e) throw hsdtAdminError_('Vui lòng nhập email quản trị.');
  var cache = CacheService.getScriptCache();
  var failKey = 'hsdtAdminFail_' + Utilities.base64EncodeWebSafe(e).slice(0, 200);
  var fails = Number(cache.get(failKey) || 0);
  if (fails >= HSDT.LOGIN_MAX_FAILS) throw hsdtAdminError_('Đăng nhập sai quá nhiều lần. Vui lòng thử lại sau 15 phút.');
  if (allowed.indexOf(e) === -1 || String(passcode || '').trim() !== secret.trim()) {
    cache.put(failKey, String(fails + 1), HSDT.LOGIN_LOCK_SECONDS);
    throw hsdtAdminError_(allowed.indexOf(e) === -1 ? 'Email này không có quyền quản trị.' : 'Mật khẩu quản trị không đúng.');
  }
  if (fails) cache.remove(failKey);
}

/** Danh sách hồ sơ cho trang quản lý: mới nhất trước. filters = { search, fromDate, toDate (yyyy-MM-dd), limit }. */
function hsdtList_(filters) {
  var sheet = hsdtSheet_();
  var info = hsdtHeaderMap_(sheet);
  var last = sheet.getLastRow();
  if (last < 2) return { total: 0, returned: 0, rows: [] };
  var values = sheet.getRange(2, 1, last - 1, info.headers.length).getValues();
  var raw = function (row, name) { var c = info.col[name]; return c ? row[c - 1] : ''; };
  var col = function (row, name) { return hsdtCellString_(raw(row, name)); };
  var fold = function (v) {
    return String(v || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd');
  };
  var search = fold(filters.search).trim();
  var from = String(filters.fromDate || ''), to = String(filters.toDate || '');
  var limit = Math.max(1, Math.min(HSDT.LIST_LIMIT, Number(filters.limit) || HSDT.LIST_LIMIT));
  var base = ScriptApp.getService().getUrl();

  var rows = values.map(function (row) {
    var id = String(col(row, 'ID') || '').trim();
    var ts = raw(row, 'Thời gian nộp');
    var when = ts instanceof Date ? ts : new Date(ts);
    var ok = !isNaN(when.getTime());
    return {
      id: id,
      code: String(col(row, 'Mã hồ sơ') || ''),
      submittedDate: ok ? Utilities.formatDate(when, HSDT.TIME_ZONE, 'yyyy-MM-dd') : '',
      submittedAt: ok ? Utilities.formatDate(when, HSDT.TIME_ZONE, 'dd/MM/yyyy HH:mm') : String(ts || ''),
      fullName: String(col(row, 'Họ và tên') || ''),
      dob: String(col(row, 'Ngày sinh') || ''),
      gender: String(col(row, 'Giới tính') || ''),
      position1: String(col(row, 'Vị trí ưu tiên 1') || ''),
      position2: String(col(row, 'Vị trí ưu tiên 2') || ''),
      salary: String(col(row, 'Lương mong đợi') || ''),
      startDate: String(col(row, 'Ngày có thể nhận việc') || ''),
      mobile: String(col(row, 'Di động') || ''),
      email: String(col(row, 'Email') || ''),
      idNumber: String(col(row, 'Số CMND/CCCD') || ''),
      address: String(col(row, 'Địa chỉ thường trú') || ''),
      everTBS: String(col(row, 'Từng làm ở TBS') || ''),
      sources: String(col(row, 'Biết tin qua') || ''),
      status: String(col(row, 'Trạng thái') || ''),
      note: String(col(row, 'Ghi chú HCNS') || ''),
      viewUrl: id ? base + '?action=view&id=' + encodeURIComponent(id) : ''
    };
  }).filter(function (r) {
    if (!r.id) return false;
    if (from && (!r.submittedDate || r.submittedDate < from)) return false;
    if (to && (!r.submittedDate || r.submittedDate > to)) return false;
    if (search && fold([r.code, r.fullName, r.position1, r.position2, r.mobile, r.email, r.idNumber, r.status].join(' ')).indexOf(search) === -1) return false;
    return true;
  });
  rows.reverse(); // dòng mới nộp nằm cuối sheet
  rows.sort(function (a, b) { return String(b.submittedDate).localeCompare(String(a.submittedDate)); });
  return { total: rows.length, returned: Math.min(rows.length, limit), rows: rows.slice(0, limit) };
}

/* =================== Nhận hồ sơ =================== */

function hsdtSubmit_(raw) {
  var id = String(raw.submissionId || '').trim();
  if (!/^[0-9a-f-]{16,40}$/i.test(id)) return { ok: false, message: 'Thiếu mã lần nộp. Vui lòng tải lại trang.' };

  var photo = raw.photo, signature = raw.signature;
  var data = hsdtClean_(raw, 0);
  delete data.photo; delete data.signature;

  var problems = [];
  if (!data.fullName) problems.push('Họ & tên');
  if (!data.dob) problems.push('Ngày sinh');
  if (!data.gender) problems.push('Giới tính');
  if (!data.position1) problems.push('Vị trí ứng tuyển');
  if (!data.idNumber) problems.push('Số CMND/CCCD');
  if (!data.permanentAddress) problems.push('Địa chỉ thường trú');
  if (String(data.mobile || '').replace(/\D/g, '').length < 9) problems.push('Di động');
  if (data.consent !== true) problems.push('Xác nhận cam kết');
  if (problems.length) return { ok: false, message: 'Còn thiếu: ' + problems.join(', ') + '.' };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, retry: true, message: 'Máy chủ đang bận, vui lòng thử lại.' };
  try {
    var sheet = hsdtSheet_();
    var info = hsdtHeaderMap_(sheet);
    var existing = hsdtFindRow_(sheet, info, id);
    if (existing) {
      var existingCode = sheet.getRange(existing, info.col['Mã hồ sơ']).getDisplayValue();
      return { ok: true, duplicate: true, id: id, code: existingCode, viewUrl: hsdtViewUrl_(id) };
    }

    var now = new Date();
    var code = 'HS' + Utilities.formatDate(now, HSDT.TIME_ZONE, 'yyMMdd') + '-' + id.replace(/-/g, '').slice(0, 5).toUpperCase();
    var safeName = String(data.fullName).replace(/[\\/:*?"<>|]/g, ' ').slice(0, 60);
    var folder = hsdtFolder_();
    data.photoFileId = hsdtSaveImage_(folder, photo, code + ' - ' + safeName + ' - anh 3x4');
    data.signatureFileId = hsdtSaveImage_(folder, signature, code + ' - ' + safeName + ' - chu ky');
    data.code = code;
    data.submittedAt = now.toISOString();

    var json = JSON.stringify(data);
    if (json.length > HSDT.MAX_JSON_IN_CELL) {
      json = 'drive:' + folder.createFile(code + ' - ' + safeName + '.json', json, 'application/json').getId();
    }

    var salary = [data.salary1Probation, data.salary1After].some(Boolean)
      ? 'Thử việc: ' + (data.salary1Probation || '-') + ' / Sau: ' + (data.salary1After || '-') : '';
    var sources = (data.sources || []).map(function (s) { return s === 'Khác' && data.sourceOther ? 'Khác: ' + data.sourceOther : s; }).join(', ');
    var values = {
      'Thời gian nộp': now,
      'Mã hồ sơ': code,
      'Họ và tên': data.fullName,
      'Ngày sinh': hsdtDate_(data.dob),
      'Giới tính': data.gender,
      'Vị trí ưu tiên 1': data.position1,
      'Vị trí ưu tiên 2': data.position2,
      'Lương mong đợi': salary,
      'Ngày có thể nhận việc': hsdtDate_(data.startDate),
      'Di động': data.mobile,
      'Email': data.email,
      'Số CMND/CCCD': data.idNumber,
      'Địa chỉ thường trú': data.permanentAddress,
      'Từng làm ở TBS': data.everTBS,
      'Biết tin qua': sources,
      'Ảnh 3x4': data.photoFileId ? hsdtLinkFormula_('https://drive.google.com/file/d/' + data.photoFileId + '/view', 'Ảnh') : '',
      'Xem / in hồ sơ': hsdtLinkFormula_(hsdtViewUrl_(id), 'Xem / in'),
      'Trạng thái': 'Mới',
      'ID': id,
      'Dữ liệu (JSON)': json
    };
    var formulaKeys = { 'Ảnh 3x4': true, 'Xem / in hồ sơ': true };
    var row = info.headers.map(function (h) {
      var v = values.hasOwnProperty(h) ? values[h] : '';
      return formulaKeys[h] || v instanceof Date ? v : hsdtCellText_(v);
    });
    sheet.appendRow(row);

    hsdtNotify_(data, code, id);
    return { ok: true, id: id, code: code, viewUrl: hsdtViewUrl_(id) };
  } finally {
    lock.releaseLock();
  }
}

function hsdtFindRow_(sheet, info, id) {
  var col = info.col['ID'];
  var last = sheet.getLastRow();
  if (!col || last < 2) return 0;
  var hit = sheet.getRange(2, col, last - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

/** Giữ đúng kiểu dữ liệu mong đợi, cắt chuỗi quá dài, giới hạn số dòng của mỗi bảng. */
function hsdtClean_(value, depth) {
  if (depth > 3) return '';
  if (value === null || value === undefined) return '';
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return isFinite(value) ? value : '';
  if (typeof value === 'string') return value.trim().slice(0, 1000);
  if (Array.isArray(value)) return value.slice(0, 12).map(function (v) { return hsdtClean_(v, depth + 1); });
  if (typeof value === 'object') {
    var out = {};
    Object.keys(value).slice(0, 80).forEach(function (k) {
      if (k === 'photo' || k === 'signature') return;
      out[String(k).slice(0, 40)] = hsdtClean_(value[k], depth + 1);
    });
    return out;
  }
  return '';
}

function hsdtSaveImage_(folder, dataUrl, name) {
  var m = /^data:(image\/(?:jpeg|png));base64,([A-Za-z0-9+\/=]+)$/.exec(String(dataUrl || ''));
  if (!m) return '';
  var bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > HSDT.MAX_IMAGE_BYTES) return '';
  var ext = m[1] === 'image/png' ? '.png' : '.jpg';
  return folder.createFile(Utilities.newBlob(bytes, m[1], name + ext)).getId();
}

/**
 * Ghi mọi chuỗi dưới dạng chữ (thêm dấu ' — Sheets không hiển thị): chống chèn công thức (= + - @) và giữ nguyên
 * số 0 đầu của số điện thoại / CMND / BHXH, ngày dạng dd/MM/yyyy không bị đổi thành kiểu ngày.
 */
function hsdtCellText_(value) {
  var s = value === null || value === undefined ? '' : String(value);
  return s ? "'" + s : s;
}

/** Ô ngày cũ (nếu Sheets đã tự đổi) → dd/MM/yyyy; ô khác → chuỗi. */
function hsdtCellString_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, HSDT.TIME_ZONE, 'dd/MM/yyyy');
  return value === null || value === undefined ? '' : String(value);
}

function hsdtLinkFormula_(url, label) {
  return '=HYPERLINK("' + String(url).replace(/"/g, '""') + '","' + String(label).replace(/"/g, '""') + '")';
}

function hsdtViewUrl_(id) {
  return ScriptApp.getService().getUrl() + '?action=view&id=' + encodeURIComponent(id);
}

function hsdtNotify_(data, code, id) {
  var to = PropertiesService.getScriptProperties().getProperty(HSDT.NOTIFY_PROP);
  if (!to) return;
  try {
    MailApp.sendEmail({
      to: to,
      subject: 'Hồ sơ dự tuyển mới: ' + data.fullName + ' – ' + data.position1 + ' (' + code + ')',
      htmlBody: '<p>Có hồ sơ dự tuyển mới.</p><ul>' +
        '<li><b>Họ tên:</b> ' + hsdtEsc_(data.fullName) + '</li>' +
        '<li><b>Vị trí:</b> ' + hsdtEsc_(data.position1) + (data.position2 ? ' / ' + hsdtEsc_(data.position2) : '') + '</li>' +
        '<li><b>Di động:</b> ' + hsdtEsc_(data.mobile) + '</li></ul>' +
        '<p><a href="' + hsdtEsc_(hsdtViewUrl_(id)) + '">Xem / in hồ sơ</a></p>'
    });
  } catch (error) {
    console.warn('Không gửi được email báo hồ sơ mới: ' + error);
  }
}

/* =================== Bản xem / in =================== */

function hsdtView_(id) {
  id = String(id || '');
  var sheet = hsdtSheet_();
  var info = hsdtHeaderMap_(sheet);
  var row = /^[0-9a-f-]{16,40}$/i.test(id) ? hsdtFindRow_(sheet, info, id) : 0;
  if (!row) return HtmlService.createHtmlOutput('<p style="font-family:Arial;padding:24px">Không tìm thấy hồ sơ.</p>').setTitle('Hồ sơ dự tuyển');
  var raw = sheet.getRange(row, info.col['Dữ liệu (JSON)']).getValue();
  if (/^drive:/.test(raw)) raw = DriveApp.getFileById(raw.slice(6)).getBlob().getDataAsString();
  var data = JSON.parse(raw);
  var html = hsdtRenderPrint_(data, hsdtImageDataUrl_(data.photoFileId), hsdtImageDataUrl_(data.signatureFileId));
  return HtmlService.createHtmlOutput(html)
    .setTitle('Hồ sơ dự tuyển - ' + (data.fullName || ''))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function hsdtImageDataUrl_(fileId) {
  if (!fileId) return '';
  try {
    var blob = DriveApp.getFileById(fileId).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (error) {
    return '';
  }
}

function hsdtEsc_(value) {
  return String(value === null || value === undefined ? '' : value).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function hsdtDate_(value) {
  var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(value || ''));
  return m ? m[3] + '/' + m[2] + '/' + m[1] : String(value || '');
}

function hsdtMonth_(value) {
  var m = /^(\d{4})-(\d{2})$/.exec(String(value || ''));
  return m ? m[2] + '/' + m[1] : String(value || '');
}

/** Bản in theo bố cục mẫu Word TBSL-NS-QT01-BM02 (A4, Times New Roman). */
function hsdtRenderPrint_(d, photoUrl, signUrl) {
  var e = hsdtEsc_;
  var v = function (x) { return x ? e(x) : '<span class="dots"></span>'; };
  var box = function (on) { return '<span class="cb">' + (on ? '✓' : '') + '</span>'; };
  var opts = function (list, chosen) {
    return list.map(function (o) { return '<span class="opt">' + box(chosen === o) + ' ' + e(o) + '</span>'; }).join('');
  };
  var pad = function (rows, min) { rows = (rows || []).slice(); while (rows.length < min) rows.push({}); return rows; };
  var sec = function (title, note) { return '<div class="sec">' + title + (note ? ' <i>(' + note + ')</i>' : '') + '</div>'; };
  var levels = ['Tốt', 'Khá', 'Căn bản', 'Không biết'];
  var skillRows = function (rows) {
    return rows.map(function (r, i) {
      return '<tr><td class="c">' + (i + 1) + '</td><td>' + e(r.name || '') + '</td>' +
        levels.map(function (l) { return '<td class="c">' + (r.level === l ? '✓' : '') + '</td>'; }).join('') +
        '<td>' + e(r.cert || '') + '</td><td></td></tr>';
    }).join('');
  };
  var skillHead = function (first) {
    return '<tr><th></th><th>' + first + '</th><th>Tốt</th><th>Khá</th><th>Căn bản</th><th>Không biết</th><th>Văn bằng</th><th>TBS đánh giá</th></tr>';
  };
  var signedDate = hsdtDate_(String(d.submittedAt || '').slice(0, 10)).split('/');

  var languages = pad(d.languages && d.languages.length ? d.languages : [{ name: 'Tiếng Anh' }], 3);
  var computer = d.computer && d.computer.length ? d.computer : [];
  var work = d.work && d.work.length ? d.work : [{}];
  var phones = pad(d.phones, 3);
  var sourceList = ['Quảng cáo', 'Mạng XH', 'TT giới thiệu việc làm', 'Người thân, bạn bè', 'Khác'];

  var h = [];
  h.push('<!DOCTYPE html><html lang="vi"><head><meta charset="utf-8"><style>' +
    '@page{size:A4;margin:10mm 11mm}' +
    '*{box-sizing:border-box}body{margin:0;background:#e9e9eb;font-family:"Times New Roman",Times,serif;font-size:10.5pt;color:#000}' +
    '.page{width:190mm;margin:12px auto;padding:0;background:#fff}' +
    '.bar{display:flex;justify-content:center;gap:10px;padding:12px;font-family:Arial,sans-serif}' +
    '.bar button{min-height:42px;padding:0 18px;border:0;border-radius:10px;background:#E45620;color:#fff;font-weight:700;font-size:15px;cursor:pointer}' +
    'table{width:100%;border-collapse:collapse}td,th{border:1px solid #000;padding:2px 5px;vertical-align:top;font-weight:normal}' +
    'th{text-align:center}.c{text-align:center}.hdr td{vertical-align:middle}.hdr .logo{width:22%;text-align:center}.hdr .logo img{height:40px}' +
    '.hdr .title{text-align:center;font-size:15pt;font-weight:bold}.hdr .meta{width:27%;font-size:9pt;line-height:1.5}' +
    '.sec{margin-top:0;padding:1px 5px;border:1px solid #000;border-top:0;background:#7f7f7f;color:#fff;font-weight:bold;font-size:10.5pt}' +
    '.sec i{font-weight:bold}.frame{border:1px solid #000;border-top:0;padding:4px 6px}' +
    '.frame p{margin:3px 0;line-height:1.55}.u{text-decoration:underline}.val{font-weight:bold}' +
    '.dots{display:inline-block;min-width:110px;border-bottom:1px dotted #000;height:1em}' +
    '.cb{display:inline-block;width:13px;height:13px;border:1px solid #000;margin:0 2px -2px 6px;font-size:10px;line-height:12px;text-align:center}' +
    '.opt{margin-right:14px;white-space:nowrap}.photo{float:right;width:30mm;height:40mm;border:1px solid #000;margin:2px 0 4px 8px;display:flex;align-items:center;justify-content:center;text-align:center}' +
    '.photo img{width:100%;height:100%;object-fit:cover}.cap{font-style:italic;padding:1px 5px;border:1px solid #000;border-top:0}' +
    '.t{border-top:0}.t td{height:18px}.grid2{display:flex;gap:10px}.grid2>div{flex:1}' +
    '.consent{font-weight:bold;font-style:italic;text-align:justify;line-height:1.45}.sign{display:flex;justify-content:flex-end}' +
    '.sign div{width:62mm;text-align:center}.sign img{height:22mm;max-width:60mm;display:block;margin:2px auto}' +
    '.verify td{border:0;padding:3px 4px}.lines div{border-bottom:1px dotted #000;height:20px}.break{page-break-before:always}' +
    '@media print{body{background:#fff}.page{margin:0}.bar{display:none}}' +
    '</style></head><body>');
  h.push('<div class="bar"><button onclick="window.print()">In / Lưu PDF</button></div><div class="page">');

  // Tiêu đề
  h.push('<table class="hdr"><tr><td class="logo"><img src="' + e(HSDT.LOGO_URL) + '" alt="TBS Logistics"></td>' +
    '<td class="title">HỒ SƠ DỰ TUYỂN</td><td class="meta">Mã tài liệu: TBSL-NS-QT01-BM02<br>Ban hành lần: 02<br>Ngày ban hành: 08/08/2022</td></tr></table>');

  // I
  h.push(sec('I. THÔNG TIN CÁ NHÂN'));
  h.push('<div class="frame"><div class="photo">' + (photoUrl ? '<img src="' + photoUrl + '" alt="Ảnh 3x4">' : 'Ảnh<br>3x4') + '</div>' +
    '<p><span class="u">Họ &amp; tên</span>: <span class="val">' + v((d.fullName || '').toUpperCase()) + '</span></p>' +
    '<p><span class="u">Ngày/tháng/năm sinh</span>: ' + v(hsdtDate_(d.dob)) + ' &nbsp; Tại: ' + v(d.birthPlace) +
    ' &nbsp;&nbsp; <span class="u">Giới tính</span>: ' + opts(['Nam', 'Nữ'], d.gender) + '</p>' +
    '<p style="clear:none"><span class="u">Vị trí ứng tuyển</span>: 1. Ưu tiên 1: ' + v(d.position1) + ' &nbsp; 2. Ưu tiên 2: ' + v(d.position2) + '</p>' +
    '<p><span class="u">Mức lương mong đợi</span>: 1. Thử việc: ' + v(d.salary1Probation) + ' &nbsp; Sau thử việc: ' + v(d.salary1After) + '</p>' +
    '<p style="padding-left:132px">2. Thử việc: ' + v(d.salary2Probation) + ' &nbsp; Sau thử việc: ' + v(d.salary2After) + '</p>' +
    '<p style="clear:both">Ngày có thể bắt đầu nhận việc: ' + v(hsdtDate_(d.startDate)) + '</p></div>');
  h.push('<table class="t"><tr><td>Quốc tịch:<br>' + e(d.nationality || '') + '</td><td>Dân tộc:<br>' + e(d.ethnicity || '') +
    '</td><td>Tôn giáo:<br>' + e(d.religion || '') + '</td><td>Số CMND/CCCD:<br><b>' + e(d.idNumber || '') + '</b></td><td>Ngày cấp:<br>' +
    e(hsdtDate_(d.idIssueDate)) + '</td><td>Nơi cấp:<br>' + e(d.idIssuePlace || '') + '</td></tr></table>');
  h.push('<table class="t"><tr><td style="width:20%">Địa chỉ thường trú</td><td>' + e(d.permanentAddress || '') + '</td><td style="width:30%">Email: ' + e(d.email || '') + '</td></tr>' +
    '<tr><td>Địa chỉ tạm trú</td><td>' + e(d.tempAddress || '') + '</td><td>Di động: <b>' + e(d.mobile || '') + '</b></td></tr>' +
    '<tr><td>Tình trạng gia đình</td><td colspan="2">' + opts(['Độc thân', 'Có gia đình', 'Ly dị', 'Góa'], d.marital) + '</td></tr></table>');

  // II
  h.push(sec('II. THÔNG TIN GIA ĐÌNH:'));
  h.push('<table class="t"><tr><th style="width:5%"></th><th>Họ &amp; tên</th><th>Quan hệ</th><th>Năm sinh</th><th>Nghề nghiệp</th><th>Địa chỉ thường trú</th></tr>' +
    pad(d.family, 7).map(function (r, i) {
      return '<tr><td class="c">' + (r.name ? i + 1 : '') + '</td><td>' + e(r.name || '') + '</td><td>' + e(r.relation || '') + '</td><td class="c">' +
        e(r.birthYear || '') + '</td><td>' + e(r.job || '') + '</td><td>' + e(r.address || '') + '</td></tr>';
    }).join('') + '</table>');

  // III
  h.push(sec('III. TRÌNH ĐỘ HỌC VẤN'));
  h.push('<div class="cap">1. Các văn bằng đã đạt được (vui lòng liệt kê các bằng cấp chuyên môn mà Anh/ Chị đạt được)</div>');
  h.push('<table class="t"><tr><th style="width:5%"></th><th>Tên trường</th><th>Chuyên ngành</th><th>Năm tốt nghiệp</th><th>Bằng cấp</th><th>Xếp loại</th></tr>' +
    pad(d.degrees, 3).map(function (r, i) {
      return '<tr><td class="c">' + (r.school ? i + 1 : '') + '</td><td>' + e(r.school || '') + '</td><td>' + e(r.major || '') + '</td><td class="c">' +
        e(r.gradYear || '') + '</td><td>' + e(r.degree || '') + '</td><td>' + e(r.rank || '') + '</td></tr>';
    }).join('') + '</table>');
  h.push('<div class="cap">2. Các khóa huấn luyện (vui lòng liệt kê các khóa học tập/ huấn luyện chuyên môn mà Anh/ Chị đã dự)</div>');
  h.push('<table class="t"><tr><th style="width:5%"></th><th>Đơn vị tổ chức</th><th>Tên khóa học</th><th>Thời gian học</th><th>Chứng chỉ</th></tr>' +
    pad(d.courses, 4).map(function (r, i) {
      return '<tr><td class="c">' + (r.course || r.org ? i + 1 : '') + '</td><td>' + e(r.org || '') + '</td><td>' + e(r.course || '') + '</td><td>' +
        e(r.time || '') + '</td><td>' + e(r.certificate || '') + '</td></tr>';
    }).join('') + '</table>');
  h.push('<div class="cap">3. Kỹ năng ngoại ngữ: (đánh dấu ✓)</div><table class="t">' + skillHead('Ngôn ngữ') + skillRows(languages) + '</table>');
  h.push('<div class="cap">4. Kỹ năng vi tính: (đánh dấu ✓)</div><table class="t">' + skillHead('Chương trình') +
    skillRows(['Email-Internet', 'Ms Word', 'Ms Excel', 'Ms PowerPoint', 'CorelDraw', 'AutoCad', ''].map(function (name, i) {
      var r = computer[i] || {};
      return { name: name || ('Khác: ' + (r.name || '')), level: r.level, cert: r.cert };
    })) + '</table>');

  // IV
  h.push(sec('IV. KINH NGHIỆM LÀM VIỆC', 'Vui lòng liệt kê từ thời gian gần đây về trước'));
  h.push('<div class="frame">' + work.map(function (w, i) {
    return '<p><b>' + (i + 1) + '. Tên Công ty: ' + e(w.company || '') + '</b></p>' +
      '<p>Địa chỉ: ' + v(w.address) + '</p>' +
      '<p>Ngành nghề kinh doanh: ' + v(w.industry) + ' &nbsp; Số lượng nhân viên: ' + v(w.staffCount) + '</p>' +
      '<p>Thời gian công tác: từ ' + v(hsdtMonth_(w.from)) + ' đến ' + v(hsdtMonth_(w.to)) + ' &nbsp; Chức vụ: ' + v(w.title) + '</p>' +
      '<p>Họ &amp; Tên cấp trên trực tiếp: ' + v(w.manager) + ' &nbsp; Chức vụ: ' + v(w.managerTitle) + '</p>' +
      '<p>Lý do nghỉ việc: ' + v(w.reason) + ' &nbsp; Mức lương cuối cùng: ' + v(w.lastSalary) + '</p>';
  }).join('<div style="height:6px"></div>') + '</div>');

  // V
  h.push(sec('V. THÔNG TIN YÊU CẦU KHÁC', 'Xin điền đầy đủ thông tin trong bản sau'));
  h.push('<div class="frame">' +
    '<p>Bạn đã từng làm việc ở TBS chưa? ' + opts(['Có', 'Không'], d.everTBS) + '</p>' +
    '<p>Nếu có, xin cho biết bộ phận làm việc và thời gian cụ thể: ' + v(d.everTBSDetail) + '</p>' +
    '<p>Bạn có người thân, bạn bè làm việc cho TBS không? Nếu có, vui lòng ghi chi tiết</p>' +
    '<p>Có/ Không: <b>' + e(d.relativeTBS || '') + '</b> &nbsp; Tên/ Chức vụ: ' + v(d.relativeTBSDetail) + '</p>' +
    '<p>Bạn có người thân, bạn bè làm việc cho công ty đối thủ của TBS không? Nếu có, vui lòng ghi chi tiết</p>' +
    '<p>Có/ Không: <b>' + e(d.relativeCompetitor || '') + '</b> &nbsp; Tên/ Chức vụ: ' + v(d.relativeCompetitorDetail) + '</p>' +
    '<p>Tại sao bạn nộp đơn vào TBS? ' + v(d.whyTBS) + '</p>' +
    '<p>Những hoạt động ưa thích? ' + v(d.hobbies) + '</p>' +
    '<p>Môn thể thao bạn ưa thích nhất? ' + v(d.sport) + ' &nbsp; Mức độ tập luyện: ' + opts(['Thường xuyên', 'Thỉnh thoảng', 'Khi thuận tiện'], d.sportFreq) + '</p>' +
    '<p>Vui lòng liệt kê 03 số điện thoại người thân/ bạn bè của bạn:</p>' +
    phones.map(function (p, i) { return '<p style="padding-left:12px">Số điện thoại ' + (i + 1) + ': ' + v(p.phone) + ' &nbsp; Quan hệ: ' + v(p.relation) + '</p>'; }).join('') +
    '<p>Bạn biết thông tin tuyển dụng của công ty chúng tôi qua:</p><p>' +
    sourceList.map(function (s) {
      return '<span class="opt">' + box((d.sources || []).indexOf(s) !== -1) + ' ' + e(s) + (s === 'Khác' && d.sourceOther ? ': ' + e(d.sourceOther) : '') + '</span>';
    }).join('') + '</p></div>');

  // VI
  h.push(sec('VI. THAM CHIẾU', 'Xin điền đầy đủ thông tin trong bản sau'));
  h.push('<table class="t"><tr><th colspan="2">Tên &amp; địa chỉ công ty</th><th>Người liên hệ</th><th>Chức vụ</th><th>Điện thoại</th></tr>' +
    pad(d.references, 5).map(function (r) {
      return '<tr><td>' + e(r.companyName || '') + '</td><td>' + e(r.companyAddress || '') + '</td><td>' + e(r.contact || '') + '</td><td>' +
        e(r.title || '') + '</td><td>' + e(r.phone || '') + '</td></tr>';
    }).join('') + '</table>');
  h.push('<div class="frame"><p>Số sổ BHXH: ' + v(d.bhxh) + ' &nbsp;&nbsp; Mã số thuế: ' + v(d.taxCode) + '</p>' +
    '<p class="consent">' + box(d.consent === true) + ' Tôi đồng ý cho Phòng HCNS ICD TBS xác minh tất cả các chi tiết cá nhân được kê khai trên đây. Tôi xác nhận đã cung cấp các thông tin chính xác và đúng sự thật trong bản khai này, và tôi sẽ chịu trách nhiệm thông báo ngay lập tức nếu có bất cứ sự thay đổi nào. Tôi hoàn toàn chịu trách nhiệm theo Nội quy lao động của công ty nếu có bất cứ chi tiết nào sai sự thật.</p>' +
    '<div class="sign"><div>Ngày ' + (signedDate[0] || '……') + '/ ' + (signedDate[1] || '……') + '/ ' + (signedDate[2] || '……') +
    '<br>Chữ ký ứng viên' + (signUrl ? '<img src="' + signUrl + '" alt="Chữ ký">' : '<div style="height:22mm"></div>') +
    '<b>' + e(d.fullName || '') + '</b></div></div>' +
    '<p style="font-size:9pt;color:#444">Mã hồ sơ: ' + e(d.code || '') + ' · Nộp trực tuyến lúc ' +
    e(d.submittedAt ? Utilities.formatDate(new Date(d.submittedAt), HSDT.TIME_ZONE, 'HH:mm dd/MM/yyyy') : '') + '</p></div>');

  // VII — để trống cho bộ phận tuyển dụng
  h.push(sec('VII. KIỂM TRA/XÁC MINH THÔNG TIN', 'Phần dành cho bộ phận tuyển dụng'));
  h.push('<div class="frame"><table class="verify">' +
    ['Đối chiếu với CMND/CCCD gốc', 'Đối chiếu với văn bằng, chứng chỉ gốc', 'Đối chiếu sơ yếu lý lịch với các hồ sơ khác', 'Xác minh quá trình làm việc', 'Xác minh hạnh kiểm ứng viên']
      .map(function (t, i) { return '<tr><td>' + (i + 1) + '. ' + t + '</td><td style="width:22%">' + box(false) + ' Đạt</td><td style="width:22%">' + box(false) + ' Không đạt</td></tr>'; }).join('') +
    '</table><p>Ý kiến khác:</p><div class="lines"><div></div><div></div><div></div><div></div></div>' +
    '<div class="sign"><div>Người kiểm tra<div style="height:22mm"></div>……………………………</div></div></div>');

  h.push('</div></body></html>');
  return h.join('');
}
