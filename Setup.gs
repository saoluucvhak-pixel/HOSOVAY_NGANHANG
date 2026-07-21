/**
 * SETUP.GS
 * Khởi tạo cấu trúc các tab nếu chưa tồn tại (an toàn khi chạy nhiều lần),
 * và lưu ID Spreadsheet vào Script Properties để Web App độc lập dùng được.
 */

function initializeSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  ensureSheet_(ss, SHEET_CONGTY, [
    'MaCty', 'TenCty', 'MaCIF', 'DiaChiTruSo', 'DienThoai', 'Fax',
    'NguoiDaiDien', 'ChucVu', 'GiayUyQuyenSo', 'GiayUyQuyenNgay', 'NguoiLapBieu'
  ]);

  ensureSheet_(ss, SHEET_HOPDONG, [
    'MaHD', 'MaCty', 'SoHopDong', 'NgayHopDong', 'TenNganHang', 'ChiNhanh',
    'DiaChiChiNhanh', 'MST_ChiNhanh', 'HanMucVay', 'LaiSuatTrongHan_%',
    'MoTaLaiSuatQuaHan', 'LaiSuatLaiChamTra_%', 'KyHanTraGoc', 'KyHanTraLai'
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

  SpreadsheetApp.getUi().alert('Đã kiểm tra/khởi tạo xong cấu trúc Sheet. Bạn có thể dùng menu "Mở ứng dụng".');
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
    var thieu = headers.filter(function (h) { return headerHienTai.indexOf(h) < 0; });
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
