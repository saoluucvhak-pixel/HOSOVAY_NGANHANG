/**
 * DATASERVICE.GS
 * Lớp truy xuất dữ liệu: đọc/ghi các tab DM_CongTy, DM_HopDongVay,
 * HoSoGiaiNgan, ChiTietThuHuong.
 */

function sheetToObjects_(sheetName) {
  var sh = getSS_().getSheetByName(sheetName);
  if (!sh) throw new Error('Không tìm thấy tab "' + sheetName + '". Hãy chạy "Khởi tạo / kiểm tra cấu trúc Sheet".');
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (row.join('') === '') continue; // bỏ dòng trống
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      obj[headers[c]] = row[c];
    }
    obj._row = r + 1; // vị trí dòng thật trên sheet (1-indexed)
    out.push(obj);
  }
  return out;
}

/** Danh sách công ty vay vốn (cho dropdown). */
/** Bộ (set) các MaHD đã có ít nhất 1 Hồ sơ giải ngân gắn vào — dùng để xác định "đang sử dụng". */
function layBoMaHopDongCoHoSo_() {
  var bo = {};
  var sh = getSS_().getSheetByName(SHEET_HOSO);
  if (!sh) return bo;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return bo;
  var soCot = sh.getLastColumn();
  var values = sh.getRange(1, 1, lastRow, soCot).getValues();
  var idx = values[0].indexOf('MaHD');
  if (idx < 0) return bo;
  for (var r = 1; r < values.length; r++) {
    var v = values[r][idx];
    if (v) bo[String(v)] = true;
  }
  return bo;
}

/** true nếu công ty này có ít nhất 1 hợp đồng đang có hồ sơ giải ngân gắn vào. */
function congTyDangSuDung_(maCty) {
  var maHDCoHoSo = layBoMaHopDongCoHoSo_();
  var hopDongCuaCty = sheetToObjects_(SHEET_HOPDONG).filter(function (hd) { return String(hd.MaCty) === String(maCty); });
  return hopDongCuaCty.some(function (hd) { return maHDCoHoSo[String(hd.MaHD)]; });
}

/** Danh sách hồ sơ giải ngân (rút gọn) đang gắn với 1 hợp đồng cụ thể — dùng để hiện khi Hợp đồng bị khoá. */
function layDanhSachHoSoTheoHopDong(maHD) {
  if (!maHD) return [];
  var sh = getSS_().getSheetByName(SHEET_HOSO);
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var soCot = sh.getLastColumn();
  var values = sh.getRange(1, 1, lastRow, soCot).getValues();
  var headers = values[0];
  var idxMaHD = headers.indexOf('MaHD');
  var idxMaHoSo = headers.indexOf('MaHoSo');
  var idxSoGNN = headers.indexOf('SoGiayNhanNo');
  var idxTrangThai = headers.indexOf('TrangThai');
  var idxNgayGiaiNgan = headers.indexOf('NgayGiaiNgan');
  var ket = [];
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idxMaHD]) !== String(maHD)) continue;
    ket.push({
      maHoSo: values[r][idxMaHoSo],
      soGiayNhanNo: values[r][idxSoGNN],
      trangThai: values[r][idxTrangThai],
      ngayGiaiNgan: (values[r][idxNgayGiaiNgan] instanceof Date)
        ? Utilities.formatDate(values[r][idxNgayGiaiNgan], 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy')
        : String(values[r][idxNgayGiaiNgan] || '')
    });
  }
  return ket;
}

/** Danh sách hồ sơ giải ngân (rút gọn) đang gắn với 1 công ty — gộp từ TẤT CẢ hợp đồng của công ty đó. */
function layDanhSachHoSoTheoCongTy(maCty) {
  if (!maCty) return [];
  var hopDongCuaCty = sheetToObjects_(SHEET_HOPDONG).filter(function (hd) { return String(hd.MaCty) === String(maCty); });
  var ket = [];
  hopDongCuaCty.forEach(function (hd) {
    layDanhSachHoSoTheoHopDong(hd.MaHD).forEach(function (hs) {
      hs.soHopDong = hd.SoHopDong;
      ket.push(hs);
    });
  });
  return ket;
}

function layDanhSachCongTy() {
  var all = sheetToObjects_(SHEET_CONGTY);
  all = chuyenDoiDateThanhChuoi_(all);
  var maHDCoHoSo = layBoMaHopDongCoHoSo_();
  var hopDongAll = sheetToObjects_(SHEET_HOPDONG);
  var maCtyDangSuDung = {};
  hopDongAll.forEach(function (hd) {
    if (maHDCoHoSo[String(hd.MaHD)]) maCtyDangSuDung[String(hd.MaCty)] = true;
  });
  all.forEach(function (c) { c.dangSuDung = !!maCtyDangSuDung[String(c.MaCty)]; });
  return all;
}

/** Danh sách hợp đồng vay (cho dropdown), lọc theo MaCty nếu truyền vào. */
function layDanhSachHopDong(maCty) {
  var all = sheetToObjects_(SHEET_HOPDONG);
  var ketQua = all;
  if (maCty) {
    ketQua = all.filter(function (x) { return String(x.MaCty) === String(maCty); });
  }
  ketQua = chuyenDoiDateThanhChuoi_(ketQua);
  var maHDCoHoSo = layBoMaHopDongCoHoSo_();
  ketQua.forEach(function (hd) { hd.dangSuDung = !!maHDCoHoSo[String(hd.MaHD)]; });
  return ketQua;
}

/** Danh sách khách hàng/nhà cung cấp (thụ hưởng) để chọn khi lập UNC thanh toán. */
function layDanhSachKhachHang() {
  var all = sheetToObjects_(SHEET_KHACHHANG);
  return chuyenDoiDateThanhChuoi_(all);
}

/**
 * Lọc danh sách hồ sơ dạng bảng — có tự động thử lại 1 lần nếu lần đọc đầu tiên trả về rỗng bất
 * thường (Google Sheets đôi khi đọc lệch/thiếu dữ liệu tạm thời sau khi vừa ghi/sửa dữ liệu).
 * boLoc = { maHoSo, tuNgay, denNgay (lọc theo Ngày giải ngân, yyyy-mm-dd), tenNguoiHuong, maCty, maHD }
 * Trả về mảng: [{ maHoSo, tenCty, soHopDong, soGiayNhanNo, ngayGiaiNgan, soTien, trangThai }, ...]
 * Không truyền bộ lọc nào (boLoc = {}) sẽ trả về TOÀN BỘ hồ sơ.
 */
function locDanhSachHoSo(boLoc) {
  var out = locDanhSachHoSo_(boLoc);
  var lanThu = 0;
  while (out.length === 0 && lanThu < 2) {
    Utilities.sleep(400 + lanThu * 300);
    out = locDanhSachHoSo_(boLoc);
    lanThu++;
  }
  return out;
}

function locDanhSachHoSo_(boLoc) {
  boLoc = boLoc || {};
  var dsHoSo = sheetToObjects_(SHEET_HOSO);
  var dsHopDong = sheetToObjects_(SHEET_HOPDONG);
  var dsCongTy = sheetToObjects_(SHEET_CONGTY);
  var dsChiTiet = sheetToObjects_(SHEET_CHITIET);

  var hopDongByMa = {};
  dsHopDong.forEach(function (h) { hopDongByMa[h.MaHD] = h; });
  var congTyByMa = {};
  dsCongTy.forEach(function (c) { congTyByMa[c.MaCty] = c; });
  var chiTietByHoSo = {};
  dsChiTiet.forEach(function (ct) {
    if (!chiTietByHoSo[ct.MaHoSo]) chiTietByHoSo[ct.MaHoSo] = [];
    chiTietByHoSo[ct.MaHoSo].push(ct);
  });

  var tu = boLoc.tuNgay ? new Date(boLoc.tuNgay + 'T00:00:00') : null;
  var den = boLoc.denNgay ? new Date(boLoc.denNgay + 'T23:59:59') : null;
  var maHoSoLoc = chuanHoaChuoi_(boLoc.maHoSo || '');
  var tenNguoiHuongLoc = chuanHoaChuoi_(boLoc.tenNguoiHuong || '');

  var out = [];
  dsHoSo.forEach(function (hs) {
    if (maHoSoLoc && chuanHoaChuoi_(hs.MaHoSo).indexOf(maHoSoLoc) < 0) return;

    var hd = hopDongByMa[hs.MaHD];
    if (boLoc.maCty && (!hd || hd.MaCty !== boLoc.maCty)) return;
    if (boLoc.maHD && hs.MaHD !== boLoc.maHD) return;

    var ngayGN = hs.NgayGiaiNgan instanceof Date ? hs.NgayGiaiNgan : (hs.NgayGiaiNgan ? new Date(hs.NgayGiaiNgan) : null);
    if (tu && ngayGN && ngayGN < tu) return;
    if (den && ngayGN && ngayGN > den) return;

    var dsCt = chiTietByHoSo[hs.MaHoSo] || [];
    if (tenNguoiHuongLoc) {
      var khop = dsCt.some(function (ct) { return chuanHoaChuoi_(ct.TenNguoiHuong).indexOf(tenNguoiHuongLoc) >= 0; });
      if (!khop) return;
    }

    var cty = hd ? congTyByMa[hd.MaCty] : null;
    out.push({
      maHoSo: hs.MaHoSo,
      tenCty: cty ? cty.TenCty : '(không rõ công ty)',
      soHopDong: hd ? hd.SoHopDong : '',
      soGiayNhanNo: hs.SoGiayNhanNo || '',
      ngayGiaiNgan: formatNgay_(hs.NgayGiaiNgan),
      soTien: Number(hs.SoTienNhanNoLanNay) || 0,
      trangThai: hs.TrangThai || '',
      _row: hs._row || 0
    });
  });

  out.sort(function (a, b) { return b._row - a._row; });
  return out;
}

function timTheoTruong_(list, field, value) {
  for (var i = 0; i < list.length; i++) {
    if (String(list[i][field]) === String(value)) return list[i];
  }
  return null;
}

function layCongTyTheoMa(maCty) {
  return timTheoTruong_(sheetToObjects_(SHEET_CONGTY), 'MaCty', maCty);
}

function layHopDongTheoMa(maHD) {
  return timTheoTruong_(sheetToObjects_(SHEET_HOPDONG), 'MaHD', maHD);
}

function layHoSoTheoMa(maHoSo) {
  return timTheoTruong_(sheetToObjects_(SHEET_HOSO), 'MaHoSo', maHoSo);
}

/** Chi tiết các dòng thụ hưởng của 1 hồ sơ giải ngân. */
function layChiTietTheoHoSo(maHoSo) {
  var all = sheetToObjects_(SHEET_CHITIET);
  var rows = all.filter(function (x) { return String(x.MaHoSo) === String(maHoSo); });
  rows.sort(function (a, b) { return (Number(a.STT) || 0) - (Number(b.STT) || 0); });
  return rows;
}

/** Danh sách khách hàng/nhà cung cấp (thụ hưởng) để chọn khi lập UNC thanh toán. */
function layDanhSachKhachHang() {
  return sheetToObjects_(SHEET_KHACHHANG);
}

/**
 * Lưu (tạo mới/cập nhật) 1 khách hàng ở tab DM_KhachHang.
 * payload = { maKH (rỗng nếu tạo mới), tenKhachHang, maSoThue, soTaiKhoan, taiNganHang, diaChi, ghiChu }
 */
function luuKhachHang(payload) {
  if (!payload.tenKhachHang) throw new Error('Vui lòng nhập Tên khách hàng.');
  if (!payload.soTaiKhoan) throw new Error('Vui lòng nhập Số tài khoản.');
  var sh = getSS_().getSheetByName(SHEET_KHACHHANG);
  var maKH = payload.maKH;
  var isNew = !maKH;
  if (isNew) maKH = sinhMaTuDong_(SHEET_KHACHHANG, 'KH', 3);

  var row = [
    maKH, payload.tenKhachHang, payload.maSoThue || '', String(payload.soTaiKhoan || ''),
    payload.taiNganHang || '', payload.diaChi || '', payload.ghiChu || ''
  ];

  if (isNew) {
    sh.appendRow(row);
  } else {
    var r = timDongTheoMa_(SHEET_KHACHHANG, 'MaKH', maKH);
    if (r < 0) throw new Error('Không tìm thấy khách hàng ' + maKH + ' để cập nhật.');
    sh.getRange(r, 1, 1, row.length).setValues([row]);
  }
  return { maKH: maKH };
}

/** Sinh mã tự động dạng PREFIX + số, dựa trên cột đầu tiên của 1 sheet. */
function sinhMaTuDong_(sheetName, prefix, doRong) {
  var sh = getSS_().getSheetByName(sheetName);
  var lastRow = sh.getLastRow();
  var max = 0;
  if (lastRow >= 2) {
    var col = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    var re = new RegExp('^' + prefix + '(\\d+)$');
    col.forEach(function (r) {
      var m = String(r[0] || '').match(re);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  var next = max + 1;
  var so = String(next);
  while (so.length < (doRong || 3)) so = '0' + so;
  return prefix + so;
}

function timDongTheoMa_(sheetName, maCot, giaTri) {
  var sh = getSS_().getSheetByName(sheetName);
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var idx = headers.indexOf(maCot);
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idx]) === String(giaTri)) return r + 1; // 1-indexed row
  }
  return -1;
}

/**
 * Lưu (tạo mới/cập nhật) 1 công ty vay vốn ở tab DM_CongTy.
 * payload = { maCty (rỗng nếu tạo mới), tenCty, maCIF, diaChiTruSo, dienThoai, fax,
 *             nguoiDaiDien, chucVu, giayUyQuyenSo, giayUyQuyenNgay, nguoiLapBieu }
 */
function luuCongTy(payload) {
  if (!payload.tenCty) throw new Error('Vui lòng nhập Tên công ty.');
  var sh = getSS_().getSheetByName(SHEET_CONGTY);
  var maCty = payload.maCty;
  var isNew = !maCty;
  if (isNew) {
    maCty = sinhMaTuDong_(SHEET_CONGTY, 'CTY', 3);
  } else if (congTyDangSuDung_(maCty)) {
    throw new Error('Công ty ' + maCty + ' đã có hồ sơ giải ngân gắn với hợp đồng của công ty này — không thể sửa. Chỉ có thể xem.');
  }

  var row = [
    maCty, payload.tenCty, payload.maCIF || '', payload.diaChiTruSo || '',
    payload.dienThoai || '', payload.fax || '', payload.nguoiDaiDien || '',
    payload.chucVu || '', payload.giayUyQuyenSo || '', payload.giayUyQuyenNgay || '',
    payload.nguoiLapBieu || ''
  ];

  if (isNew) {
    sh.appendRow(row);
  } else {
    var r = timDongTheoMa_(SHEET_CONGTY, 'MaCty', maCty);
    if (r < 0) throw new Error('Không tìm thấy công ty ' + maCty + ' để cập nhật.');
    sh.getRange(r, 1, 1, row.length).setValues([row]);
  }
  return { maCty: maCty };
}

/**
 * Lưu (tạo mới/cập nhật) 1 hợp đồng cho vay theo hạn mức ở tab DM_HopDongVay.
 * payload = { maHD (rỗng nếu tạo mới), maCty, soHopDong, ngayHopDong, tenNganHang, chiNhanh,
 *             diaChiChiNhanh, mstChiNhanh, hanMucVay, laiSuatTrongHan, moTaLaiSuatQuaHan,
 *             laiSuatLaiChamTra, kyHanTraGoc, kyHanTraLai, hauToThamChieu }
 */
function luuHopDong(payload) {
  if (!payload.maCty) throw new Error('Vui lòng chọn công ty vay vốn.');
  if (!payload.soHopDong) throw new Error('Vui lòng nhập Số hợp đồng.');
  var sh = getSS_().getSheetByName(SHEET_HOPDONG);
  var maHD = payload.maHD;
  var isNew = !maHD;
  if (isNew) {
    maHD = sinhMaTuDong_(SHEET_HOPDONG, 'HD', 3);
  } else if (layBoMaHopDongCoHoSo_()[String(maHD)]) {
    throw new Error('Hợp đồng ' + maHD + ' đã có hồ sơ giải ngân gắn vào — không thể sửa. Chỉ có thể xem.');
  }

  var row = [
    maHD, payload.maCty, payload.soHopDong, toDateOrEmpty_(payload.ngayHopDong),
    payload.tenNganHang || '', payload.chiNhanh || '', payload.diaChiChiNhanh || '',
    payload.mstChiNhanh || '', Number(payload.hanMucVay) || 0,
    Number(payload.laiSuatTrongHan) || 0, payload.moTaLaiSuatQuaHan || '',
    Number(payload.laiSuatLaiChamTra) || 0, payload.kyHanTraGoc || 'Cuối kỳ',
    payload.kyHanTraLai || 'Hàng tháng', payload.hauToThamChieu || ''
  ];

  if (isNew) {
    sh.appendRow(row);
  } else {
    var r = timDongTheoMa_(SHEET_HOPDONG, 'MaHD', maHD);
    if (r < 0) throw new Error('Không tìm thấy hợp đồng ' + maHD + ' để cập nhật.');
    sh.getRange(r, 1, 1, row.length).setValues([row]);
  }
  return { maHD: maHD };
}

function toDateOrEmpty_(v) {
  if (!v) return '';
  if (v instanceof Date) return v;
  var d = new Date(v);
  return isNaN(d.getTime()) ? v : d;
}

/**
 * Lưu (tạo mới hoặc cập nhật) một hồ sơ giải ngân + danh sách chi tiết thụ hưởng.
 */
function luuHoSoGiaiNgan(payload) {
  var ss = getSS_();
  var shHoSo = ss.getSheetByName(SHEET_HOSO);
  var shChiTiet = ss.getSheetByName(SHEET_CHITIET);

  var maHoSo = payload.maHoSo;
  var isNew = !maHoSo;
  if (isNew) maHoSo = sinhMaTuDong_(SHEET_HOSO, 'HS', 4);

  // ===== Xác thực dữ liệu (hàng rào cuối cùng phía server — phòng trường hợp gọi thẳng qua API,
  // bỏ qua kiểm tra phía giao diện ở JavaScript.html) =====
  var chiTietKiemTra = payload.chiTiet || [];
  var tongThuHuongKiemTra = chiTietKiemTra.reduce(function (s, ct) { return s + (Number(ct.soTien) || 0); }, 0);
  if (!chiTietKiemTra.length || tongThuHuongKiemTra <= 0) {
    throw new Error('Hồ sơ phải có ít nhất 1 dòng thụ hưởng với Số tiền lớn hơn 0 ở mục 4.');
  }

  var soTienBangChu = soThanhChuVN(payload.soTienNhanNoLanNay);
  var headers = shHoSo.getRange(1, 1, 1, shHoSo.getLastColumn()).getValues()[0];

  var giaTriTheoTen = {
    MaHoSo: maHoSo,
    MaHD: payload.maHD || '',
    SoGiayNhanNo: payload.soGiayNhanNo || '',
    NgayGiayNhanNo: toDateOrEmpty_(payload.ngayGiayNhanNo),
    DuNoHienTai: Number(payload.duNoHienTai) || 0,
    SoTienNhanNoLanNay: Number(payload.soTienNhanNoLanNay) || 0,
    SoTienBangChu: soTienBangChu,
    PhuongThucThanhToan: payload.phuongThucThanhToan || 'Chuyển khoản',
    MucDichSuDungVon: payload.mucDichSuDungVon || '',
    ThoiHanChoVay_Ngay: payload.thoiHanChoVayNgay || '',
    NgayGiaiNgan: toDateOrEmpty_(payload.ngayGiaiNgan),
    NgayDenHan: toDateOrEmpty_(payload.ngayDenHan),
    TaiLieuChungMinhMucDich: payload.taiLieuChungMinhMucDich || '',
    NguoiLapBieu: payload.nguoiLapBieu || '',
    SoThamChieu: payload.soThamChieu || '',
    TrangThai: payload.trangThai || 'Chờ tạo hồ sơ',
    NgayTao: new Date()
  };

  var existing = isNew ? null : layHoSoTheoMa(maHoSo);
  if (!isNew && !existing) throw new Error('Không tìm thấy hồ sơ ' + maHoSo + ' để cập nhật.');
  if (existing && existing.TrangThai === 'Đã giải ngân') {
    throw new Error('Hồ sơ ' + maHoSo + ' đã ở trạng thái "Đã giải ngân", không thể sửa nữa.');
  }

  var rowData = headers.map(function (h) {
    if (giaTriTheoTen.hasOwnProperty(h)) return giaTriTheoTen[h];
    return existing && existing.hasOwnProperty(h) ? existing[h] : '';
  });

  if (isNew) {
    shHoSo.appendRow(rowData);
  } else {
    shHoSo.getRange(existing._row, 1, 1, rowData.length).setValues([rowData]);
  }

  var allChiTiet = shChiTiet.getDataRange().getValues();
  var headerLen = allChiTiet[0].length;
  var rowsToDelete = [];
  for (var r = allChiTiet.length - 1; r >= 1; r--) {
    if (String(allChiTiet[r][0]) === String(maHoSo)) rowsToDelete.push(r + 1);
  }
  rowsToDelete.forEach(function (rowIdx) { shChiTiet.deleteRow(rowIdx); });

  var chiTiet = payload.chiTiet || [];
  chiTiet.forEach(function (ct, idx) {
    shChiTiet.appendRow([
      maHoSo,
      idx + 1,
      ct.tenNguoiHuong || '',
      ct.soTaiKhoan || '',
      ct.taiNganHang || '',
      ct.loaiTien || 'VND',
      ct.noiDungThanhToan || '',
      Number(ct.soTien) || 0,
      ct.chuyenTienNhanh || 'N',
      ct.taiLieuSo || '',
      ct.ngayTaiLieu || '',
      ct.donViLapTaiLieu || ct.tenNguoiHuong || '',
      ct.ngayCapGiayToTuyThan || '',
      ct.ghiChu || ''
    ]);
  });

  try { capNhatBaoCaoDraft_(); } catch (e) { /* không chặn việc lưu hồ sơ nếu bước này lỗi */ }

  return { maHoSo: maHoSo, trangThai: giaTriTheoTen.TrangThai };
}

function capNhatLinkHoSo_(maHoSo, cot, url) {
  var sh = getSS_().getSheetByName(SHEET_HOSO);
  var hs = layHoSoTheoMa(maHoSo);
  if (!hs) return;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var colIdx = headers.indexOf(cot) + 1;
  if (colIdx > 0) sh.getRange(hs._row, colIdx).setValue(url);
}

function layNgayHomNay() {
  return Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
}

function chuanHoaChuoi_(s) {
  return String(s || '').trim().toUpperCase().replace(/\s+/g, ' ');
}

function timCotTheoTenHoacViTri_(headers, tenCot, viTriMacDinh) {
  var idx = -1;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim().toLowerCase() === tenCot) { idx = i; break; }
  }
  return idx >= 0 ? idx : viTriMacDinh;
}

function soTuChuoiHoacSo_(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined || v === '') return 0;
  var s = String(v).replace(/[^0-9.\-]/g, '');
  if (!s) return 0;
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function layTapSoHDDaSuDung_(maHoSoBoQua) {
  var sh = getSS_().getSheetByName(SHEET_CHITIET);
  if (!sh) return {};
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return {};
  var headers = values[0];
  var idxMaHoSo = headers.indexOf('MaHoSo');
  var idxTaiLieuSo = headers.indexOf('TaiLieuSo');
  if (idxTaiLieuSo < 0) return {};
  var tap = {};
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (maHoSoBoQua && idxMaHoSo >= 0 && String(row[idxMaHoSo]) === String(maHoSoBoQua)) continue;
    var raw = String(row[idxTaiLieuSo] || '');
    if (!raw) continue;
    raw.split(',').forEach(function (s) {
      var v = s.trim();
      if (v) tap[v] = true;
    });
  }
  return tap;
}

/**
 * Lấy danh sách hoá đơn của 1 nhà cung cấp/khách hàng từ Sổ chi tiết mua hàng, ĐÃ GỘP NHÓM theo
 * (Nhà cung cấp + Số hoá đơn + Ngày hoá đơn) và CỘNG DỒN số tiền của các dòng cùng nhóm — vì
 * thanh toán được thực hiện theo NGUYÊN hoá đơn (không theo từng mặt hàng/dòng lẻ trong hoá đơn),
 * nên tránh để 1 hoá đơn xuất hiện lặp lại nhiều lần với số tiền từng dòng nhỏ lẻ.
 */
function layDanhSachTaiLieuTheoNhaCungCap(tenNguoiHuong, maHoSoHienTai) {
  if (!tenNguoiHuong) return [];

  var ch = CauHinh_();
  var ss = SpreadsheetApp.openById(ch.ssMuaHangId);
  var sh = ss.getSheetByName(ch.sheetMuaHang);
  if (!sh) {
    throw new Error('Không tìm thấy tab "' + ch.sheetMuaHang + '" trong file Sổ chi tiết mua hàng.');
  }

  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var idxSoHD = timCotTheoTenHoacViTri_(headers, 'sohd', 7);
  var idxNgay = timCotTheoTenHoacViTri_(headers, 'ngay', 3);
  var idxNCC = timCotTheoTenHoacViTri_(headers, 'nhacungcap', 8);
  var idxDienGiai = timCotTheoTenHoacViTri_(headers, 'diengiai', 1);
  var idxGiaTri = timCotTheoTenHoacViTri_(headers, 'giatri', 14);

  var chuan = chuanHoaChuoi_(tenNguoiHuong);

  var nhomChinhXac = {}, thuTuChinhXac = [];
  var nhomGanDung = {}, thuTuGanDung = [];

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var nccRaw = row[idxNCC];
    if (!nccRaw) continue;
    var soHD = row[idxSoHD];
    var ngay = row[idxNgay];
    if (soHD === '' || soHD === null || soHD === undefined || !ngay) continue;

    var nccChuan = chuanHoaChuoi_(nccRaw);
    var laKhopChinhXac = (nccChuan === chuan);
    var laKhopGanDung = !laKhopChinhXac && (nccChuan.indexOf(chuan) >= 0 || chuan.indexOf(nccChuan) >= 0);
    if (!laKhopChinhXac && !laKhopGanDung) continue;

    var giaTriDong = soTuChuoiHoacSo_(row[idxGiaTri]);
    var ngayStr = (ngay instanceof Date) ? Utilities.formatDate(ngay, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy') : String(ngay);
    var soHDStr = String(soHD).trim();
    var dienGiaiDong = idxDienGiai >= 0 ? String(row[idxDienGiai] || '') : '';
    var key = nccChuan + '|' + soHDStr + '|' + ngayStr;

    var nhom = laKhopChinhXac ? nhomChinhXac : nhomGanDung;
    var thuTu = laKhopChinhXac ? thuTuChinhXac : thuTuGanDung;

    if (!nhom[key]) {
      nhom[key] = {
        sohd: soHDStr,
        ngay: ngayStr,
        ngaySort: (ngay instanceof Date) ? ngay.getTime() : 0,
        dienGiai: dienGiaiDong,
        nhaCungCap: String(nccRaw).trim(),
        giaTri: 0,
        soDong: 0
      };
      thuTu.push(key);
    }
    nhom[key].giaTri += giaTriDong;
    nhom[key].soDong++;
  }

  var dungNhom = thuTuChinhXac.length ? nhomChinhXac : nhomGanDung;
  var dungThuTu = thuTuChinhXac.length ? thuTuChinhXac : thuTuGanDung;

  var out = dungThuTu.map(function (key) {
    var it = dungNhom[key];
    return {
      sohd: it.sohd,
      ngay: it.ngay,
      ngaySort: it.ngaySort,
      dienGiai: it.soDong > 1 ? (it.dienGiai + ' (gộp ' + it.soDong + ' dòng)') : it.dienGiai,
      nhaCungCap: it.nhaCungCap,
      giaTri: it.giaTri
    };
  });

  var daDung = layTapSoHDDaSuDung_(maHoSoHienTai);
  out = out.filter(function (it) { return !daDung[it.sohd]; });

  out.sort(function (a, b) { return a.ngaySort - b.ngaySort; });
  return out;
}

function chuyenDoiDateThanhChuoi_(obj) {
  if (Array.isArray(obj)) {
    return obj.map(chuyenDoiDateThanhChuoi_);
  }
  if (obj && typeof obj === 'object') {
    var out = {};
    Object.keys(obj).forEach(function (k) {
      var v = obj[k];
      if (v instanceof Date) {
        out[k] = isNaN(v.getTime()) ? '' : Utilities.formatDate(v, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
      } else {
        out[k] = v;
      }
    });
    return out;
  }
  return obj;
}

function layDuLieuDayDu(maHoSo) {
  var hoSo = layHoSoTheoMa(maHoSo);
  if (!hoSo) throw new Error('Không tìm thấy hồ sơ: ' + maHoSo);
  var hopDong = layHopDongTheoMa(hoSo.MaHD);
  if (!hopDong) throw new Error('Không tìm thấy hợp đồng vay: ' + hoSo.MaHD);
  var congTy = layCongTyTheoMa(hopDong.MaCty);
  if (!congTy) throw new Error('Không tìm thấy công ty vay: ' + hopDong.MaCty);
  var chiTiet = layChiTietTheoHoSo(maHoSo);
  return {
    hoSo: chuyenDoiDateThanhChuoi_(hoSo),
    hopDong: chuyenDoiDateThanhChuoi_(hopDong),
    congTy: chuyenDoiDateThanhChuoi_(congTy),
    chiTiet: chuyenDoiDateThanhChuoi_(chiTiet)
  };
}

function danhDauDaGiaiNgan(maHoSo) {
  var sh = getSS_().getSheetByName(SHEET_HOSO);
  var hs = layHoSoTheoMa(maHoSo);
  if (!hs) throw new Error('Không tìm thấy hồ sơ ' + maHoSo + '.');
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var colTrangThai = headers.indexOf('TrangThai') + 1;
  if (colTrangThai <= 0) throw new Error('Không tìm thấy cột TrangThai trong sheet ' + SHEET_HOSO + '.');
  sh.getRange(hs._row, colTrangThai).setValue('Đã giải ngân');
  try { capNhatBaoCaoDraft_(); } catch (e) { /* không chặn việc đánh dấu nếu bước này lỗi */ }
  return { maHoSo: maHoSo, trangThai: 'Đã giải ngân' };
}

function xoaHoSo(maHoSo) {
  var ss = getSS_();
  var shHoSo = ss.getSheetByName(SHEET_HOSO);
  var hs = layHoSoTheoMa(maHoSo);
  if (!hs) throw new Error('Không tìm thấy hồ sơ ' + maHoSo + '.');
  if (hs.TrangThai === 'Đã giải ngân') {
    throw new Error('Hồ sơ đã ở trạng thái "Đã giải ngân", không thể xoá.');
  }
  shHoSo.deleteRow(hs._row);

  var shChiTiet = ss.getSheetByName(SHEET_CHITIET);
  var allChiTiet = shChiTiet.getDataRange().getValues();
  var rowsToDelete = [];
  for (var r = allChiTiet.length - 1; r >= 1; r--) {
    if (String(allChiTiet[r][0]) === String(maHoSo)) rowsToDelete.push(r + 1);
  }
  rowsToDelete.forEach(function (rowIdx) { shChiTiet.deleteRow(rowIdx); });

  return { maHoSo: maHoSo };
}

/** Đếm số dòng ở 1 sheet có giá trị `giaTri` tại cột `tenCot` — dùng để kiểm tra ràng buộc trước khi xoá. */
function demSoDongThamChieu_(sheetName, tenCot, giaTri) {
  var sh = getSS_().getSheetByName(sheetName);
  if (!sh) return 0;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  var soCot = sh.getLastColumn();
  var values = sh.getRange(1, 1, lastRow, soCot).getValues();
  var headers = values[0];
  var idx = headers.indexOf(tenCot);
  if (idx < 0) return 0;
  var dem = 0;
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idx]) === String(giaTri)) dem++;
  }
  return dem;
}

/** Xoá nhiều dòng cùng lúc ở 1 sheet, dựa theo danh sách giá trị khớp cột `maCot`. */
function xoaNhieuDongTheoMa_(sheetName, maCot, danhSachMa) {
  if (!danhSachMa || !danhSachMa.length) return 0;
  var sh = getSS_().getSheetByName(sheetName);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  var soCot = sh.getLastColumn();
  var values = sh.getRange(1, 1, lastRow, soCot).getValues();
  var headers = values[0];
  var idx = headers.indexOf(maCot);
  if (idx < 0) return 0;
  var boMa = {};
  danhSachMa.forEach(function (m) { boMa[String(m)] = true; });
  var soDongDaXoa = 0;
  // Xoá từ dòng cuối lên đầu để không bị lệch chỉ số khi xoá dần.
  for (var r = values.length - 1; r >= 1; r--) {
    if (boMa[String(values[r][idx])]) {
      sh.deleteRow(r + 1);
      soDongDaXoa++;
    }
  }
  return soDongDaXoa;
}

/**
 * Xoá nhiều công ty vay vốn cùng lúc (chọn từ bảng danh sách, dùng nút "Xoá đã chọn").
 * An toàn dữ liệu: BỎ QUA (không xoá) công ty nào đang có ít nhất 1 Hợp đồng vay gắn với nó,
 * để tránh làm "mồ côi" dữ liệu hợp đồng/hồ sơ đã tạo trước đó.
 * @param {string[]} danhSachMa - danh sách MaCty cần xoá
 * @return {{daXoa: string[], boQua: {ma:string, lyDo:string}[]}}
 */
function xoaCongTyNhieu(danhSachMa) {
  danhSachMa = danhSachMa || [];
  var daXoa = [], boQua = [];
  var maHDCoHoSo = layBoMaHopDongCoHoSo_();
  var hopDongAll = sheetToObjects_(SHEET_HOPDONG);
  danhSachMa.forEach(function (ma) {
    var hopDongCuaCty = hopDongAll.filter(function (hd) { return String(hd.MaCty) === String(ma); });
    var soHoSo = hopDongCuaCty.filter(function (hd) { return maHDCoHoSo[String(hd.MaHD)]; }).length;
    if (soHoSo > 0) {
      boQua.push({ ma: ma, lyDo: 'đang có ' + soHoSo + ' hồ sơ giải ngân gắn với hợp đồng của công ty này' });
    } else {
      daXoa.push(ma);
    }
  });
  xoaNhieuDongTheoMa_(SHEET_CONGTY, 'MaCty', daXoa);
  return { daXoa: daXoa, boQua: boQua };
}

/**
 * Xoá nhiều hợp đồng vay cùng lúc. BỎ QUA hợp đồng nào đang có ít nhất 1 Hồ sơ giải ngân gắn với nó.
 * @param {string[]} danhSachMa - danh sách MaHD cần xoá
 * @return {{daXoa: string[], boQua: {ma:string, lyDo:string}[]}}
 */
function xoaHopDongNhieu(danhSachMa) {
  danhSachMa = danhSachMa || [];
  var daXoa = [], boQua = [];
  danhSachMa.forEach(function (ma) {
    var soHoSo = demSoDongThamChieu_(SHEET_HOSO, 'MaHD', ma);
    if (soHoSo > 0) {
      boQua.push({ ma: ma, lyDo: 'đang có ' + soHoSo + ' hồ sơ giải ngân gắn với hợp đồng này' });
    } else {
      daXoa.push(ma);
    }
  });
  xoaNhieuDongTheoMa_(SHEET_HOPDONG, 'MaHD', daXoa);
  return { daXoa: daXoa, boQua: boQua };
}

/**
 * Xoá nhiều khách hàng / nhà cung cấp cùng lúc. Khách hàng chỉ là danh mục gợi ý điền nhanh
 * (không phải khoá ngoại của hồ sơ đã lưu — hồ sơ lưu trực tiếp tên/STK vào ChiTietThuHuong),
 * nên không cần kiểm tra ràng buộc, xoá được ngay.
 * @param {string[]} danhSachMa - danh sách MaKH cần xoá
 * @return {{daXoa: string[], boQua: {ma:string, lyDo:string}[]}}
 */
function xoaKhachHangNhieu(danhSachMa) {
  danhSachMa = danhSachMa || [];
  xoaNhieuDongTheoMa_(SHEET_KHACHHANG, 'MaKH', danhSachMa);
  return { daXoa: danhSachMa, boQua: [] };
}

function capNhatBaoCaoDraft_() {


  var rows = layBaoCaoTienVay('', '', '', '');

  var ss = getSS_();
  var sh = ss.getSheetByName(SHEET_BAOCAO_DRAFT);
  if (!sh) {
    sh = ss.insertSheet(SHEET_BAOCAO_DRAFT);
  } else {
    sh.clearContents();
  }

  var header = ['MaHoSo', 'TrangThai', 'SoHopDong', 'SoGiayNhanNo', 'NgayNhanNo', 'NgayGiaiNgan', 'TenNguoiHuong',
    'SoTienVay', 'SoTaiLieu', 'NgayTaiLieu', 'NgayDenHan', 'LaiSuatTrongHan'];
  var data = [header];
  rows.forEach(function (r) {
    data.push([
      String(r.maHoSo || ''), String(r.trangThai || ''), String(r.soHopDong || ''), String(r.soGiayNhanNo || ''),
      String(r.ngayNhanNo || ''), String(r.ngayGiaiNgan || ''), String(r.tenNguoiHuong || ''),
      Number(r.soTienVay) || 0, String(r.soTaiLieu || ''), String(r.ngayTaiLieu || ''),
      String(r.ngayDenHan || ''), String(r.laiSuatTrongHan || '')
    ]);
  });

  // Đặt định dạng TEXT (@) cho các cột ngày TRƯỚC khi ghi dữ liệu, để Google Sheets không tự
  // suy đoán/định dạng lại kiểu Date theo locale máy (tránh hiển thị sai khác dd/mm/yyyy).
  var cotNgayDraft = [5, 6, 10, 11]; // NgayNhanNo, NgayGiaiNgan, NgayTaiLieu, NgayDenHan
  cotNgayDraft.forEach(function (c) {
    sh.getRange(1, c, data.length, 1).setNumberFormat('@');
  });

  sh.getRange(1, 1, data.length, header.length).setValues(data);
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  SpreadsheetApp.flush();
  return sh;
}

function docBaoCaoDraft_() {
  var ss = getSS_();
  var sh = ss.getSheetByName(SHEET_BAOCAO_DRAFT);
  if (!sh || sh.getLastRow() < 1) {
    sh = capNhatBaoCaoDraft_();
  }
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = sh.getLastColumn();
  var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return values.map(function (row) {
    return {
      maHoSo: row[0], trangThai: row[1], soHopDong: row[2], soGiayNhanNo: row[3], ngayNhanNo: row[4],
      ngayGiaiNgan: row[5], tenNguoiHuong: row[6], soTienVay: Number(row[7]) || 0,
      soTaiLieu: row[8], ngayTaiLieu: row[9], ngayDenHan: row[10], laiSuatTrongHan: row[11]
    };
  });
}

function parseNgayVN_(s) {
  if (!s || typeof s !== 'string') return null;
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

/**
 * Tối ưu: Lọc báo cáo từ Sheet Draft với khả năng chuẩn hóa chuỗi tuyệt đối và xử lý "Chọn tất cả"
 */
function locBaoCaoTienVayQuaSheet(tuNgay, denNgay, tenKhachHang, trangThai) {
  capNhatBaoCaoDraft_();
  var tatCa = docBaoCaoDraft_();

  var tu = tuNgay ? new Date(tuNgay + 'T00:00:00') : null;
  var den = denNgay ? new Date(denNgay + 'T23:59:59') : null;
  var tuKhoa = chuanHoaChuoi_(tenKhachHang || '');
  
  var trangThaiLoc = String(trangThai || '').trim().normalize('NFC');
  var isAllTrangThai = (!trangThaiLoc || trangThaiLoc === 'ALL' || trangThaiLoc.toLowerCase() === 'chọn tất cả' || trangThaiLoc.toLowerCase() === '-- tất cả --');

  return tatCa.filter(function (r) {
    if (!isAllTrangThai) {
      var statusRow = String(r.trangThai || '').trim().normalize('NFC');
      if (statusRow.toLowerCase() !== trangThaiLoc.toLowerCase()) return false;
    }

    if (tuKhoa && chuanHoaChuoi_(r.tenNguoiHuong).indexOf(tuKhoa) < 0) return false;

    if (tu || den) {
      var ngayGN = parseNgayVN_(r.ngayGiaiNgan);
      if (ngayGN && !isNaN(ngayGN.getTime())) {
        if (tu && ngayGN < tu) return false;
        if (den && ngayGN > den) return false;
      }
    }
    return true;
  });
}

function layBaoCaoTienVay(tuNgay, denNgay, tenKhachHang, trangThai) {
  try {
    // Gọi thẳng vào hàm gốc bên dưới, không qua trung gian nào khác để chống tràn stack
    return layBaoCaoTienVay_(tuNgay, denNgay, tenKhachHang, trangThai);
  } catch (e) {
    // Nếu có lỗi, trả về mảng rỗng để giao diện không bị sập
    Logger.log("Lỗi tại layBaoCaoTienVay: " + e.message);
    return [];
  }
}










/**
 * Tối ưu: Lọc báo cáo trực tiếp từ dữ liệu gốc với logic kiểm tra an toàn
 */
function layBaoCaoTienVay_(tuNgay, denNgay, tenKhachHang, trangThai) {
  var dsHoSo = sheetToObjects_(SHEET_HOSO);
  var dsChiTiet = sheetToObjects_(SHEET_CHITIET);
  var dsHopDong = sheetToObjects_(SHEET_HOPDONG);

  var tu = tuNgay ? new Date(tuNgay + 'T00:00:00') : null;
  var den = denNgay ? new Date(denNgay + 'T23:59:59') : null;
  
  // Dùng hàm xóa dấu riêng biệt ở server
  var tuKhoa = boDauTiengViet_(tenKhachHang); 
  
  var trangThaiLoc = String(trangThai || '').trim().normalize('NFC');
  var isAllTrangThai = (!trangThaiLoc || trangThaiLoc === 'ALL' || trangThaiLoc.toLowerCase() === 'chọn tất cả' || trangThaiLoc.toLowerCase() === '-- tất cả --');

  var hopDongByMa = {};
  dsHopDong.forEach(function (h) { hopDongByMa[h.MaHD] = h; });

  var chiTietByHoSo = {};
  dsChiTiet.forEach(function (ct) {
    var key = String(ct.MaHoSo || '').trim();
    if (!key) return;
    if (!chiTietByHoSo[key]) chiTietByHoSo[key] = [];
    chiTietByHoSo[key].push(ct);
  });

  var out = [];
  dsHoSo.forEach(function (hs) {
    // 1. Lọc theo trạng thái
    if (!isAllTrangThai) {
      var statusRow = String(hs.TrangThai || '').trim().normalize('NFC');
      if (statusRow.toLowerCase() !== trangThaiLoc.toLowerCase()) return;
    }

    // 2. Lọc theo Ngày Giải Ngân (Xử lý thông minh cả kiểu String và Date của Google Sheets)
    var ngayGN = null;
    if (hs.NgayGiaiNgan instanceof Date) {
        ngayGN = hs.NgayGiaiNgan;
    } else if (typeof hs.NgayGiaiNgan === 'string' && hs.NgayGiaiNgan.trim() !== '') {
        // Tự động phân tích ngày dạng dd/MM/yyyy nếu Google Sheets hiểu lầm là Text
        var m = hs.NgayGiaiNgan.trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
        if (m) {
            ngayGN = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
        } else {
            ngayGN = new Date(hs.NgayGiaiNgan);
        }
    }
    
    // Nếu có chọn Từ Ngày / Đến Ngày thì mới đối chiếu
    if (tu || den) {
        if (!ngayGN || isNaN(ngayGN.getTime())) return; 
        if (tu && ngayGN < tu) return;
        if (den && ngayGN > den) return;
    }

    var hd = hopDongByMa[hs.MaHD];
    var laiSuat = hd ? hd['LaiSuatTrongHan_%'] : '';
    var soHopDong = hd ? hd.SoHopDong : '';

    // 3. Lọc theo Tên Khách Hàng (Tương đối, không phân biệt dấu)
    var dsCt = chiTietByHoSo[String(hs.MaHoSo || '').trim()] || [];
    dsCt.forEach(function (ct) {
      if (tuKhoa && boDauTiengViet_(ct.TenNguoiHuong).indexOf(tuKhoa) < 0) return;
      
      out.push({
        maHoSo: hs.MaHoSo,
        trangThai: hs.TrangThai || '',
        soHopDong: soHopDong || '',
        soGiayNhanNo: hs.SoGiayNhanNo || '',
        ngayNhanNo: formatNgay_(hs.NgayGiayNhanNo),
        ngayGiaiNgan: formatNgay_(hs.NgayGiaiNgan),
        tenNguoiHuong: ct.TenNguoiHuong || '',
        soTienVay: Number(ct.SoTien) || 0,
        soTaiLieu: ct.TaiLieuSo || '',
        ngayTaiLieu: ct.NgayTaiLieu || '',
        ngayDenHan: formatNgay_(hs.NgayDenHan),
        laiSuatTrongHan: laiSuat || ''
      });
    });
  });

  out.sort(function (a, b) { return String(b.ngayGiaiNgan).localeCompare(String(a.ngayGiaiNgan)); });
  return out;
}

// Cần dán thêm hàm hỗ trợ này vào file DATASERVICE.GS
function boDauTiengViet_(str) {
  if (!str) return '';
  return String(str).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().trim();
}

function testLoiHopDong() {
  var tatCa = layDanhSachHopDong();
  Logger.log("Tổng số hợp đồng tìm thấy: " + tatCa.length);
  Logger.log(tatCa);
}

// 1. Hàm chuẩn hóa ngày tháng chống lệch định dạng
function chuanHoaNgay_(ngayStr) {
  if (!ngayStr) return "";
  if (typeof ngayStr === 'number') {
    var d = new Date(Math.round((ngayStr - 25569) * 86400 * 1000));
    var day = ('0' + d.getDate()).slice(-2);
    var month = ('0' + (d.getMonth() + 1)).slice(-2);
    var year = d.getFullYear();
    return day + '/' + month + '/' + year;
  }
  var str = String(ngayStr).trim();
  if (str.indexOf(',') !== -1) {
    return str.split(',').map(function(item) { return item.trim(); }).filter(Boolean).join(', ');
  }
  return str;
}

// 2. Hàm lấy và bọc dữ liệu báo cáo trả về giao diện
function layBaoCaoTienVay(tuNgay, denNgay, tenKhachHang, trangThai) {
  // Lấy dữ liệu từ hàm gốc của bạn
  var danhSach = layBaoCaoTienVay_(tuNgay, denNgay, tenKhachHang, trangThai);
  
  // Tự động quét và chuẩn hóa toàn bộ ngày tháng để không bị lỗi hiển thị
  if (danhSach && danhSach.forEach) {
    danhSach.forEach(function(r) {
      if (r.ngayNhanNo) r.ngayNhanNo = chuanHoaNgay_(r.ngayNhanNo);
      if (r.ngayGiaiNgan) r.ngayGiaiNgan = chuanHoaNgay_(r.ngayGiaiNgan);
      if (r.ngayTaiLieu) r.ngayTaiLieu = chuanHoaNgay_(r.ngayTaiLieu);
      if (r.ngayDenHan) r.ngayDenHan = chuanHoaNgay_(r.ngayDenHan);
    });
  }
  
  return danhSach;
}
