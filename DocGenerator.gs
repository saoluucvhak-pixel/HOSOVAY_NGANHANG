/**
 * DOCGENERATOR.GS
 * Các hàm định dạng dùng chung + sinh báo cáo tiền vay (Google Doc/Excel).
 * Phần sinh 5 văn bản hồ sơ vay bằng cách dựng Document từ đầu (Giấy nhận nợ, Văn bản đề nghị
 * giải ngân, Uỷ nhiệm chi, Bảng kê tài liệu, Phụ lục 02) đã được thay thế hoàn toàn bởi
 * TemplateGenerator.gs (copy mẫu thật rồi điền số liệu) nên đã được xoá khỏi file này.
 */

// ---------- Tiện ích định dạng ----------

function formatSo_(n) {
  n = Math.round(Number(n) || 0);
  var s = String(Math.abs(n));
  var out = '';
  var count = 0;
  for (var i = s.length - 1; i >= 0; i--) {
    out = s.charAt(i) + out;
    count++;
    if (count % 3 === 0 && i !== 0) out = '.' + out;
  }
  return (n < 0 ? '-' : '') + out;
}

function formatNgay_(ngayGoc) {
  // Nếu ô dữ liệu trống, trả về chuỗi rỗng để không bị hiện "Invalid Date"
  if (!ngayGoc || ngayGoc === '') return '';

  var d;
  // Kiểm tra xem dữ liệu truyền vào đã là kiểu Date chưa
  if (ngayGoc instanceof Date) {
    d = ngayGoc;
  } else {
    // Nếu từ nền tảng khác truyền sang dạng chuỗi, ép nó về Date
    d = new Date(ngayGoc);
  }

  // Nếu ép kiểu thất bại (dữ liệu rác), trả lại nguyên gốc
  if (isNaN(d.getTime())) {
    return String(ngayGoc);
  }

  // Chuyển đổi an toàn với múi giờ của Việt Nam
  return Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy');
}

function formatNgayDayDu_(v) {
  if (!v) return 'ngày .... tháng .... năm ........';
  var d = (v instanceof Date) ? v : new Date(v);
  if (isNaN(d.getTime())) return String(v);
  return 'ngày ' + Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'dd') +
    ' tháng ' + Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'MM') +
    ' năm ' + Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'yyyy');
}

// ---------- Drive: thư mục gốc ----------

function layThuMucGoc_() {
  return DriveApp.getFolderById(THU_MUC_GOC_ID);
}

// =========================================================================
// BÁO CÁO TIỀN VAY (tab "Báo cáo tiền vay")
// =========================================================================

function layThuMucBaoCao_() {
  var goc = layThuMucGoc_();
  var it = goc.getFoldersByName('BaoCao_TienVay');
  if (it.hasNext()) return it.next();
  return goc.createFolder('BaoCao_TienVay');
}

/**
 * Tạo Google Doc "BẢNG KÊ CHI TIẾT TIỀN VAY THEO KHẾ ƯỚC" (để in), kèm xuất PDF.
 * Lưu vào thư mục con "BaoCao_TienVay" bên trong thư mục gốc (THU_MUC_GOC_ID).
 * Trả về { docUrl, pdfUrl }.
 */
function taoBaoCaoTienVayDoc(tuNgay, denNgay, tenKhachHang, nguoiLapBieu, trangThai) {
  var rows = layBaoCaoTienVay(tuNgay, denNgay, tenKhachHang, trangThai);

  var doc = DocumentApp.create('BaoCaoTienVay_' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyyMMdd_HHmmss'));
  var body = doc.getBody();
  body.setPageWidth(842).setPageHeight(595); // A4 ngang, cho bảng nhiều cột dễ đọc
  body.setMarginLeft(30).setMarginRight(30).setMarginTop(30).setMarginBottom(30);

  var t = body.appendParagraph('BẢNG KÊ CHI TIẾT TIỀN VAY THEO KHẾ ƯỚC');
  t.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(14);
  var p = body.appendParagraph('Từ ngày ' + (tuNgay ? formatNgay_(new Date(tuNgay)) : '……………') +
    ' đến ngày ' + (denNgay ? formatNgay_(new Date(denNgay)) : '……………'));
  p.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(11);
  if (tenKhachHang) {
    var kh = body.appendParagraph('Khách hàng: ' + tenKhachHang);
    kh.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setItalic(true);
  }
  if (trangThai) {
    var tt = body.appendParagraph('Trạng thái hồ sơ: ' + trangThai);
    tt.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setItalic(true);
  }
  body.appendParagraph('');

  var header = ['STT', 'Mã hồ sơ', 'Số hợp đồng', 'Số giấy nhận nợ', 'Ngày nhận nợ', 'Ngày giải ngân', 'Tên người hưởng',
    'Số tiền vay', 'Số tài liệu', 'Ngày tài liệu', 'Ngày đến hạn', 'Lãi suất trong hạn (%/năm)', 'Trạng thái'];
  var dataRows = [header];
  var tong = 0;
  rows.forEach(function (r, i) {
    dataRows.push([
      String(i + 1), 
      String(r.maHoSo || ''), 
      String(r.soHopDong || ''), 
      String(r.soGiayNhanNo || ''), 
      
      // SỬA Ở ĐÂY: Dùng formatNgay_ cho các cột ngày tháng
      (r.ngayNhanNo ? formatNgay_(new Date(r.ngayNhanNo)) : ''), 
      (r.ngayGiaiNgan ? formatNgay_(new Date(r.ngayGiaiNgan)) : ''), 
      
      String(r.tenNguoiHuong || ''), 
      formatSo_(r.soTienVay),
      String(r.soTaiLieu || ''), 
      
      // SỬA Ở ĐÂY: Dùng formatNgay_ cho các cột ngày tháng
      (r.ngayTaiLieu ? formatNgay_(new Date(r.ngayTaiLieu)) : ''), 
      (r.ngayDenHan ? formatNgay_(new Date(r.ngayDenHan)) : ''), 
      
      String(r.laiSuatTrongHan || ''), 
      String(r.trangThai || '')
    ]);
    tong += Number(r.soTienVay) || 0;
  });;
  dataRows.push(['', '', '', '', '', '', 'TỔNG CỘNG', formatSo_(tong), '', '', '', '', '']);

  var table = body.appendTable(dataRows);
  table.setBorderWidth(0.75);
  for (var rr = 0; rr < table.getNumRows(); rr++) {
    for (var cc = 0; cc < header.length; cc++) {
      table.getRow(rr).getCell(cc).editAsText().setFontSize(9);
    }
  }
  for (var c = 0; c < header.length; c++) table.getRow(0).getCell(c).editAsText().setBold(true);
  var lastR = table.getNumRows() - 1;
  table.getRow(lastR).getCell(6).editAsText().setBold(true);
  table.getRow(lastR).getCell(7).editAsText().setBold(true);

  body.appendParagraph('');
  var ngayLap = body.appendParagraph('Ngày lập ' + formatNgay_(new Date()));
  ngayLap.setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setItalic(true);
  var lap = body.appendParagraph('Người lập biểu');
  lap.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph('');
  body.appendParagraph('');
  var tenLap = body.appendParagraph(nguoiLapBieu || '');
  tenLap.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  doc.saveAndClose();

  var folder = layThuMucBaoCao_();
  var file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch (e) { /* bỏ qua */ }

  var pdfBlob = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF).setName(doc.getName() + '.pdf');
  var pdfFile = DriveApp.createFile(pdfBlob);
  folder.addFile(pdfFile);
  try { DriveApp.getRootFolder().removeFile(pdfFile); } catch (e) { /* bỏ qua */ }

  return { docUrl: doc.getUrl(), pdfUrl: pdfFile.getUrl() };
}

/**
 * Xuất TOÀN BỘ dữ liệu tiền vay ra Excel, KHÔNG áp dụng bộ lọc Ngày/Khách hàng/Trạng thái —
 * file có sẵn AutoFilter để tự lọc/sắp xếp ngay trong Excel. Đây là cách đáng tin cậy hơn lọc
 * "live" nhiều lần trên Apps Script (không phụ thuộc độ trễ đọc Google Sheets).
 */
function taoBaoCaoTienVayExcelToanBo() {
  return taoBaoCaoTienVayExcel('', '', '', '');
}

function taoBaoCaoTienVayExcel(tuNgay, denNgay, tenKhachHang, trangThai) {
  var rows = layBaoCaoTienVay(tuNgay, denNgay, tenKhachHang, trangThai);

  var ss = SpreadsheetApp.create('BaoCaoTienVay_' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyyMMdd_HHmmss'));
  var sh = ss.getSheets()[0];
  sh.setName('BaoCao');

  sh.getRange(1, 1).setValue('BẢNG KÊ CHI TIẾT TIỀN VAY THEO KHẾ ƯỚC').setFontWeight('bold').setFontSize(13);
  var coLoc = !!(tuNgay || denNgay || tenKhachHang || trangThai);
  sh.getRange(2, 1).setValue(coLoc
    ? ('Từ ngày ' + (tuNgay || '') + ' đến ngày ' + (denNgay || '') +
        (tenKhachHang ? ('  |  Khách hàng: ' + tenKhachHang) : '') +
        (trangThai ? ('  |  Trạng thái: ' + trangThai) : ''))
    : 'TOÀN BỘ dữ liệu (chưa lọc) — dùng nút lọc (▼) trên dòng tiêu đề bảng để tự lọc theo Ngày/Trạng thái/Khách hàng ngay trong Excel.');

  var header = ['STT', 'Mã hồ sơ', 'Số hợp đồng', 'Số giấy nhận nợ', 'Ngày nhận nợ', 'Ngày giải ngân', 'Tên người hưởng',
    'Số tiền vay', 'Số tài liệu', 'Ngày tài liệu', 'Ngày đến hạn', 'Lãi suất trong hạn (%/năm)', 'Trạng thái'];
  var data = [header];
  var tong = 0;
  rows.forEach(function (r, i) {
    data.push([i + 1, r.maHoSo, r.soHopDong, r.soGiayNhanNo, r.ngayNhanNo, r.ngayGiaiNgan, r.tenNguoiHuong,
      r.soTienVay, r.soTaiLieu, r.ngayTaiLieu, r.ngayDenHan, r.laiSuatTrongHan, r.trangThai]);
    tong += Number(r.soTienVay) || 0;
  });
  data.push(['', '', '', '', '', '', 'TỔNG CỘNG', tong, '', '', '', '', '']);

  var startRow = 4;
  // Đặt định dạng TEXT (@) cho các cột ngày TRƯỚC khi ghi dữ liệu, để Google Sheets không tự
  // suy đoán/định dạng lại kiểu Date theo locale máy (tránh hiển thị sai khác dd/mm/yyyy).
  var cotNgay = [5, 6, 10, 11]; // Ngày nhận nợ, Ngày giải ngân, Ngày tài liệu, Ngày đến hạn
  cotNgay.forEach(function (c) {
    sh.getRange(startRow, c, data.length, 1).setNumberFormat('@');
  });
  sh.getRange(startRow, 1, data.length, header.length).setValues(data);
  sh.getRange(startRow, 1, 1, header.length).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  sh.getRange(startRow + data.length - 1, 7, 1, 2).setFontWeight('bold');
  sh.autoResizeColumns(1, header.length);
  try {
    sh.getRange(startRow, 1, data.length - 1, header.length).createFilter();
  } catch (e) { /* không chặn nếu tạo filter lỗi */ }
  sh.setFrozenRows(startRow);
  SpreadsheetApp.flush();

  var folder = layThuMucBaoCao_();
  var file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch (e) { /* bỏ qua */ }

  var exportUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  var resp = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    return { xlsxUrl: '', sheetUrl: ss.getUrl() };
  }
  var xlsxBlob = resp.getBlob().setName(ss.getName() + '.xlsx');
  var xlsxFile = DriveApp.createFile(xlsxBlob);
  folder.addFile(xlsxFile);
  try { DriveApp.getRootFolder().removeFile(xlsxFile); } catch (e) { /* bỏ qua */ }

  return { xlsxUrl: xlsxFile.getUrl(), sheetUrl: ss.getUrl() };
}
