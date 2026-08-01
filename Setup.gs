/**
 * SETUP.GS
 * Khởi tạo cấu trúc các tab nếu chưa tồn tại (an toàn khi chạy nhiều lần),
 * và lưu ID Spreadsheet vào Script Properties để Web App độc lập dùng được.
 */

function initializeSpreadsheet() {
  khoiTaoCauTrucSheet_();
  SpreadsheetApp.getUi().alert('Đã kiểm tra/khởi tạo xong cấu trúc Sheet. Bạn có thể dùng menu "Mở ứng dụng".');
}

/**
 * Gọi được từ tab "Cài đặt" trên Web App (không có SpreadsheetApp.getUi() ở ngữ cảnh Web App
 * nên không thể gọi thẳng initializeSpreadsheet() — dùng hàm này thay thế, trả về kết quả dạng
 * object để giao diện hiển thị thông báo thay vì alert().
 */
function khoiTaoDanhMucTuWebApp() {
  khoiTaoCauTrucSheet_();
  return { ok: true, thongBao: 'Đã kiểm tra/khởi tạo xong cấu trúc Sheet (DM_CongTy, DM_HopDongVay, DM_KhachHang, HoSoGiaiNgan, ChiTietThuHuong, BaoCaoTienVay_Draft).' };
}

/** Phần lõi khởi tạo/kiểm tra cấu trúc Sheet — an toàn khi chạy nhiều lần, không đụng dữ liệu cũ. */
function khoiTaoCauTrucSheet_() {
  var ss = getSS_();

  ensureSheet_(ss, SHEET_CONGTY, [
    'MaCty', 'TenCty', 'MaCIF', 'DiaChiTruSo', 'DienThoai', 'Fax',
    'NguoiDaiDien', 'ChucVu', 'GiayUyQuyenSo', 'GiayUyQuyenNgay', 'NguoiLapBieu'
  ]);
  suaLoiCotNguoiLapBieu_(); // dọn lại cột "Người lập biểu" nếu bị lệch/trùng do sheet cũ

  ensureSheet_(ss, SHEET_HOPDONG, [
    'MaHD', 'MaCty', 'SoHopDong', 'NgayHopDong', 'TenNganHang', 'ChiNhanh',
    'DiaChiChiNhanh', 'MST_ChiNhanh', 'HanMucVay', 'LaiSuatTrongHan_%',
    'MoTaLaiSuatQuaHan', 'LaiSuatLaiChamTra_%', 'KyHanTraGoc', 'KyHanTraLai',
    'HauToThamChieu'
  ]);

  ensureSheet_(ss, SHEET_KHACHHANG, [
    'MaKH', 'TenKhachHang', 'MaSoThue', 'SoTaiKhoan', 'TaiNganHang', 'DiaChi', 'GhiChu'
  ]);

  ensureSheet_(ss, SHEET_HOSO, [
    'MaHoSo', 'MaHD', 'SoGiayNhanNo', 'NgayGiayNhanNo', 'DuNoHienTai',
    'SoTienNhanNoLanNay', 'SoTienBangChu', 'PhuongThucThanhToan',
    'MucDichSuDungVon', 'ThoiHanChoVay_Ngay', 'NgayGiaiNgan', 'NgayDenHan',
    'TaiLieuChungMinhMucDich', 'NguoiLapBieu', 'SoThamChieu', 'TrangThai',
    'NgayTao',
    'Link_GiayNhanNo', 'Link_GiayNhanNo_PDF', 'Link_VanBanDeNghi', 'Link_VanBanDeNghi_PDF',
    'Link_UNC', 'Link_UNC_PDF', 'Link_BangKeTaiLieu', 'Link_BangKeTaiLieu_PDF',
    'Link_BangKeUNC', 'Link_BangKeUNC_PDF'
  ]);
  xoaCotThua_(ss.getSheetByName(SHEET_HOSO), [
    'LinkGiayNhanNo', 'LinkDeNghiGiaiNgan', 'LinkUyNhiemChi',
    'LinkBangKeTaiLieu', 'LinkPhuLuc02', 'LinkPDF_TrongBo'
  ]);

  ensureSheet_(ss, SHEET_CHITIET, [
    'MaHoSo', 'STT', 'TenNguoiHuong', 'SoTaiKhoan', 'TaiNganHang', 'LoaiTien',
    'NoiDungThanhToan', 'SoTien', 'ChuyenTienNhanh_24_7', 'TaiLieuSo',
    'NgayTaiLieu', 'DonViLapTaiLieu', 'NgayCapGiayToTuyThan', 'GhiChu'
  ]);

  ensureSheet_(ss, SHEET_BAOCAO_DRAFT, [
    'MaHoSo', 'TrangThai', 'SoHopDong', 'SoGiayNhanNo', 'NgayNhanNo', 'NgayGiaiNgan', 'TenNguoiHuong',
    'SoTienVay', 'SoTaiLieu', 'NgayTaiLieu', 'NgayDenHan', 'LaiSuatTrongHan'
  ]);
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  } else {
    // Sheet đã có dữ liệu: chỉ bổ sung các cột còn THIẾU vào cuối, không đụng tới cột/dữ liệu đã có.
    var soCotHienTai = sh.getLastColumn();
    var headerHienTai = soCotHienTai > 0 ? sh.getRange(1, 1, 1, soCotHienTai).getValues()[0] : [];
    // So khớp bỏ khoảng trắng thừa đầu/cuối, tránh nhận nhầm là "thiếu cột" chỉ vì tiêu đề dính dấu cách.
    var headerHienTaiChuan = headerHienTai.map(function (h) { return String(h || '').trim(); });
    var thieu = headers.filter(function (h) { return headerHienTaiChuan.indexOf(h) < 0; });
    if (thieu.length) {
      var batDau = soCotHienTai + 1;
      sh.getRange(1, batDau, 1, thieu.length).setValues([thieu]);
      sh.getRange(1, batDau, 1, thieu.length).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
    }
  }
  return sh;
}

/** Xoá các cột không còn dùng nữa khỏi 1 sheet, tra theo TÊN cột (an toàn nếu cột không tồn tại). */
function xoaCotThua_(sh, tenCotCanXoa) {
  if (!sh) return;
  var soCot = sh.getLastColumn();
  if (soCot === 0) return;
  var header = sh.getRange(1, 1, 1, soCot).getValues()[0];
  // Xoá từ cột cuối về đầu để không bị lệch chỉ số khi xoá dần.
  for (var c = header.length - 1; c >= 0; c--) {
    if (tenCotCanXoa.indexOf(header[c]) >= 0) {
      sh.deleteColumn(c + 1);
    }
  }
}

/**
 * Tự sửa lỗi lệch cột "Người lập biểu" ở DM_CongTy — trường hợp cột K (vị trí chuẩn) từng có
 * tiêu đề rỗng hoặc mang tên cũ khác (VD "NguoiLap"), khiến ensureSheet_ tưởng cột "NguoiLapBieu"
 * chưa tồn tại và tự thêm 1 cột MỚI ở cuối — dữ liệu cũ ở cột K bị "mồ côi", không đọc được.
 * Hàm này gộp mọi dữ liệu ứng viên về đúng 1 cột "NguoiLapBieu" tại vị trí chuẩn (cột K = cột 11),
 * rồi xoá các cột trùng thừa. Chạy an toàn nhiều lần (không làm gì nếu đã đúng).
 */
function suaLoiCotNguoiLapBieu_() {
  var sh = getSS_().getSheetByName(SHEET_CONGTY);
  if (!sh) return;
  var soCot = sh.getLastColumn();
  if (soCot === 0) return;
  var header = sh.getRange(1, 1, 1, soCot).getValues()[0];
  var VI_TRI_CHUAN = 11; // cột K: 10 cột đầu (MaCty..GiayUyQuyenNgay) + cột 11 = NguoiLapBieu

  var cacCotUngVien = [];
  for (var c = 0; c < header.length; c++) {
    var ten = String(header[c] || '').trim();
    if (ten === 'NguoiLapBieu' || ten === 'NguoiLap') cacCotUngVien.push(c + 1);
  }
  if (String(header[VI_TRI_CHUAN - 1] || '').trim() === '' && cacCotUngVien.indexOf(VI_TRI_CHUAN) < 0) {
    cacCotUngVien.push(VI_TRI_CHUAN);
  }
  if (!cacCotUngVien.length) return;

  var lastRow = Math.max(sh.getLastRow(), 1);
  var soDong = Math.max(lastRow - 1, 0);

  var giaTriGop = [];
  for (var r = 0; r < soDong; r++) giaTriGop.push('');
  cacCotUngVien.forEach(function (colIdx) {
    if (soDong === 0) return;
    var vals = sh.getRange(2, colIdx, soDong, 1).getValues();
    for (var r = 0; r < soDong; r++) {
      if (!giaTriGop[r] && vals[r][0]) giaTriGop[r] = vals[r][0];
    }
  });

  sh.getRange(1, VI_TRI_CHUAN).setValue('NguoiLapBieu').setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  if (soDong > 0) sh.getRange(2, VI_TRI_CHUAN, soDong, 1).setValues(giaTriGop.map(function (v) { return [v]; }));

  var cotCanXoa = cacCotUngVien.filter(function (c) { return c !== VI_TRI_CHUAN; }).sort(function (a, b) { return b - a; });
  cotCanXoa.forEach(function (c) { sh.deleteColumn(c); });
}
