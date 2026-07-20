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
function layDanhSachCongTy() {
  return sheetToObjects_(SHEET_CONGTY);
}

/** Danh sách hợp đồng vay (cho dropdown), lọc theo MaCty nếu truyền vào. */
function layDanhSachHopDong(maCty) {
  var all = sheetToObjects_(SHEET_HOPDONG);
  if (!maCty) return all;
  return all.filter(function (x) { return String(x.MaCty) === String(maCty); });
}

/** Toàn bộ hồ sơ giải ngân đã tạo (cho danh sách/lịch sử). */
function layDanhSachHoSo() {
  var list = sheetToObjects_(SHEET_HOSO);
  list.sort(function (a, b) { return (b._row || 0) - (a._row || 0); });
  return list;
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
    // Chỉ loại theo khoảng ngày khi hồ sơ CÓ Ngày giải ngân; hồ sơ chưa điền ngày này vẫn được giữ lại
    // nếu chỉ lọc theo Mã hồ sơ/Công ty/Hợp đồng/Tên người thụ hưởng.
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
 *             nguoiDaiDien, chucVu, giayUyQuyenSo, giayUyQuyenNgay }
 */
function luuCongTy(payload) {
  if (!payload.tenCty) throw new Error('Vui lòng nhập Tên công ty.');
  var sh = getSS_().getSheetByName(SHEET_CONGTY);
  var maCty = payload.maCty;
  var isNew = !maCty;
  if (isNew) maCty = sinhMaTuDong_(SHEET_CONGTY, 'CTY', 3);

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
 *             laiSuatLaiChamTra, kyHanTraGoc, kyHanTraLai }
 */
function luuHopDong(payload) {
  if (!payload.maCty) throw new Error('Vui lòng chọn công ty vay vốn.');
  if (!payload.soHopDong) throw new Error('Vui lòng nhập Số hợp đồng.');
  var sh = getSS_().getSheetByName(SHEET_HOPDONG);
  var maHD = payload.maHD;
  var isNew = !maHD;
  if (isNew) maHD = sinhMaTuDong_(SHEET_HOPDONG, 'HD', 3);

  var row = [
    maHD, payload.maCty, payload.soHopDong, toDateOrEmpty_(payload.ngayHopDong),
    payload.tenNganHang || '', payload.chiNhanh || '', payload.diaChiChiNhanh || '',
    payload.mstChiNhanh || '', Number(payload.hanMucVay) || 0,
    Number(payload.laiSuatTrongHan) || 0, payload.moTaLaiSuatQuaHan || '',
    Number(payload.laiSuatLaiChamTra) || 0, payload.kyHanTraGoc || 'Cuối kỳ',
    payload.kyHanTraLai || 'Hàng tháng'
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
 * payload = {
 *   maHoSo (rỗng nếu tạo mới), maHD, soGiayNhanNo, ngayGiayNhanNo, duNoHienTai,
 *   soTienNhanNoLanNay, phuongThucThanhToan, mucDichSuDungVon, thoiHanChoVayNgay,
 *   ngayGiaiNgan, ngayDenHan, taiLieuChungMinhMucDich, nguoiLapBieu, soThamChieu, trangThai,
 *   chiTiet: [{ tenNguoiHuong, soTaiKhoan, taiNganHang, loaiTien, noiDungThanhToan,
 *               soTien, chuyenTienNhanh, taiLieuSo, ngayTaiLieu, donViLapTaiLieu,
 *               ngayCapGiayToTuyThan, ghiChu }, ...]
 * }
 * Lưu ý: nếu payload.trangThai để trống, hồ sơ MỚI sẽ mặc định "Đã tạo hồ sơ"; khi cập nhật hồ sơ
 * đã có, client nên truyền lại trạng thái hiện tại để không bị vô tình reset trạng thái.
 */
function luuHoSoGiaiNgan(payload) {
  var ss = getSS_();
  var shHoSo = ss.getSheetByName(SHEET_HOSO);
  var shChiTiet = ss.getSheetByName(SHEET_CHITIET);

  var maHoSo = payload.maHoSo;
  var isNew = !maHoSo;
  if (isNew) maHoSo = sinhMaTuDong_(SHEET_HOSO, 'HS', 4);

  var soTienBangChu = soThanhChuVN(payload.soTienNhanNoLanNay);
  var headers = shHoSo.getRange(1, 1, 1, shHoSo.getLastColumn()).getValues()[0];

  // Các giá trị hồ sơ TỰ NHẬP — cột nào không có trong đây (VD: các cột Link_...) sẽ được GIỮ NGUYÊN
  // giá trị cũ khi cập nhật (tra theo TÊN cột, không theo vị trí, để không lệch khi thêm cột mới).
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
    TrangThai: payload.trangThai || 'Nháp',
    NgayTao: new Date()
  };

  var existing = isNew ? null : layHoSoTheoMa(maHoSo);
  if (!isNew && !existing) throw new Error('Không tìm thấy hồ sơ ' + maHoSo + ' để cập nhật.');
  if (existing && existing.TrangThai === 'Đã giải ngân') {
    throw new Error('Hồ sơ ' + maHoSo + ' đã ở trạng thái "Đã giải ngân", không thể sửa nữa.');
  }

  var rowData = headers.map(function (h) {
    if (giaTriTheoTen.hasOwnProperty(h)) return giaTriTheoTen[h];
    // Cột không tự nhập (VD: Link_...) — giữ nguyên giá trị cũ nếu đang cập nhật, để trống nếu tạo mới.
    return existing && existing.hasOwnProperty(h) ? existing[h] : '';
  });

  if (isNew) {
    shHoSo.appendRow(rowData);
  } else {
    shHoSo.getRange(existing._row, 1, 1, rowData.length).setValues([rowData]);
  }

  // Xoá các dòng chi tiết cũ của hồ sơ này rồi ghi lại toàn bộ (đơn giản & an toàn)
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

/**
 * Lấy ngày hôm nay theo múi giờ Việt Nam (Asia/Ho_Chi_Minh), dạng yyyy-MM-dd.
 * Dùng để đặt mặc định các ô lọc ngày ở client — tránh lệch ngày do đồng hồ/múi giờ trình duyệt
 * khác múi giờ Việt Nam (VD: máy tính đặt giờ UTC) khiến bộ lọc loại nhầm hồ sơ mới nhất.
 */
function layNgayHomNay() {
  return Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
}

/**
 * Chuẩn hoá chuỗi để so khớp tên (bỏ khoảng trắng thừa, chuyển hoa).
 */
function chuanHoaChuoi_(s) {
  return String(s || '').trim().toUpperCase().replace(/\s+/g, ' ');
}

/**
 * Tìm chỉ số cột theo tên header (không phân biệt hoa/thường, bỏ khoảng trắng thừa);
 * nếu không tìm thấy tên cột thì dùng vị trí mặc định (0-indexed) truyền vào — để vẫn hoạt động
 * được kể cả khi header trong Sheet gõ khác đôi chút (hoa/thường, dấu cách...).
 */
function timCotTheoTenHoacViTri_(headers, tenCot, viTriMacDinh) {
  var idx = -1;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim().toLowerCase() === tenCot) { idx = i; break; }
  }
  return idx >= 0 ? idx : viTriMacDinh;
}

/**
 * Chuyển 1 giá trị (số hoặc chuỗi có định dạng dấu phẩy/chấm ngăn cách hàng nghìn) thành số thuần.
 */
function soTuChuoiHoacSo_(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined || v === '') return 0;
  var s = String(v).replace(/[^0-9.\-]/g, '');
  if (!s) return 0;
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

/**
 * Tập hợp các Số hoá đơn (TaiLieuSo) đã được dùng ở các hồ sơ KHÁC trong ChiTietThuHuong,
 * để loại khỏi danh sách cho chọn (tránh dùng trùng 1 hoá đơn cho nhiều hồ sơ).
 * maHoSoBoQua: mã hồ sơ đang sửa — các dòng thuộc chính hồ sơ này KHÔNG bị loại.
 */
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
 * Tra cứu danh sách hoá đơn (Số hoá đơn + Ngày) trong file "Sổ chi tiết mua hàng"
 * (sheet SO_CHI_TIET_MUA_HANG, cột I "nhacungcap") theo Tên người hưởng, để chọn
 * nhanh vào ô "Tài liệu số" / "Ngày tài liệu" ở mục 3 - Danh sách thụ hưởng.
 * Trả về mảng: [{ sohd, ngay (dd/MM/yyyy), ngaySort, dienGiai, nhaCungCap, giaTri }, ...]
 * đã loại trùng, loại các hoá đơn đã dùng ở hồ sơ khác, và sắp xếp theo ngày tăng dần.
 * giaTri lấy từ cột "giatri" (cột O) trong SO_CHI_TIET_MUA_HANG, dùng để gợi ý "Số tiền"
 * của dòng thụ hưởng (vẫn sửa tay được).
 * maHoSoHienTai: mã hồ sơ đang sửa (để không loại các hoá đơn đã gán cho chính hồ sơ này).
 */
function layDanhSachTaiLieuTheoNhaCungCap(tenNguoiHuong, maHoSoHienTai) {
  if (!tenNguoiHuong) return [];

  var ss = SpreadsheetApp.openById(SS_MUA_HANG_ID);
  var sh = ss.getSheetByName(SHEET_MUA_HANG);
  if (!sh) {
    throw new Error('Không tìm thấy tab "' + SHEET_MUA_HANG + '" trong file Sổ chi tiết mua hàng.');
  }

  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var idxSoHD = timCotTheoTenHoacViTri_(headers, 'sohd', 7);        // cột H
  var idxNgay = timCotTheoTenHoacViTri_(headers, 'ngay', 3);        // cột D
  var idxNCC = timCotTheoTenHoacViTri_(headers, 'nhacungcap', 8);   // cột I
  var idxDienGiai = timCotTheoTenHoacViTri_(headers, 'diengiai', 1);// cột B
  var idxGiaTri = timCotTheoTenHoacViTri_(headers, 'giatri', 14);   // cột O

  var chuan = chuanHoaChuoi_(tenNguoiHuong);
  var khopChinhXac = [];
  var khopGanDung = [];

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var nccRaw = row[idxNCC];
    if (!nccRaw) continue;
    var soHD = row[idxSoHD];
    var ngay = row[idxNgay];
    if (soHD === '' || soHD === null || soHD === undefined || !ngay) continue;

    var nccChuan = chuanHoaChuoi_(nccRaw);
    var giaTri = soTuChuoiHoacSo_(row[idxGiaTri]);
    var item = {
      sohd: String(soHD).trim(),
      ngay: (ngay instanceof Date) ? Utilities.formatDate(ngay, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy') : String(ngay),
      ngaySort: (ngay instanceof Date) ? ngay.getTime() : 0,
      dienGiai: idxDienGiai >= 0 ? String(row[idxDienGiai] || '') : '',
      nhaCungCap: String(nccRaw).trim(),
      giaTri: giaTri
    };

    if (nccChuan === chuan) {
      khopChinhXac.push(item);
    } else if (nccChuan.indexOf(chuan) >= 0 || chuan.indexOf(nccChuan) >= 0) {
      khopGanDung.push(item);
    }
  }

  // Ưu tiên khớp chính xác tên nhà cung cấp; nếu không có thì dùng khớp gần đúng (chứa nhau).
  var ketQua = khopChinhXac.length ? khopChinhXac : khopGanDung;

  // Loại bỏ trùng lặp: mỗi hoá đơn thường có 2 dòng hạch toán (Nợ/Có) giống hệt Số hoá đơn +
  // Ngày + Giá trị. Gộp theo cả 3 để không mất các dòng có Giá trị khác nhau trong cùng 1 hoá đơn.
  var daThay = {};
  var out = [];
  ketQua.forEach(function (it) {
    var key = it.sohd + '|' + it.ngay + '|' + it.giaTri;
    if (daThay[key]) return;
    daThay[key] = true;
    out.push(it);
  });

  // Loại các hoá đơn đã được dùng (Tài liệu số) ở hồ sơ khác.
  var daDung = layTapSoHDDaSuDung_(maHoSoHienTai);
  out = out.filter(function (it) { return !daDung[it.sohd]; });

  out.sort(function (a, b) { return a.ngaySort - b.ngaySort; });
  return out;
}

/** Dữ liệu tổng hợp đầy đủ cho 1 hồ sơ: hồ sơ + hợp đồng + công ty + chi tiết. */
/**
 * Chuyển mọi giá trị kiểu Date trong 1 object (hoặc mảng object) phẳng thành chuỗi yyyy-MM-dd.
 * google.script.run đôi khi serialize Date không ổn định (đặc biệt Date "rỗng"/không hợp lệ do ô
 * trống trong Sheet) khiến cả object trả về bị null phía client — chuyển hết sang chuỗi để tránh lỗi.
 */
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

/**
 * Đánh dấu 1 hồ sơ là "Đã giải ngân" (cột TrangThai - cột P sheet HoSoGiaiNgan).
 * Sau khi đánh dấu, hồ sơ này không thể sửa/lưu lại được nữa (xem thêm guard trong luuHoSoGiaiNgan).
 */
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

/**
 * Xoá 1 hồ sơ (và toàn bộ chi tiết thụ hưởng liên quan). Chỉ cho phép xoá khi hồ sơ đang ở
 * trạng thái "Đã tạo hồ sơ" (chưa tạo văn bản/giải ngân), để tránh xoá nhầm hồ sơ đã xử lý.
 */
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

/**
 * Báo cáo chi tiết tiền vay theo khế ước, lọc theo khoảng "Ngày giải ngân" [tuNgay, denNgay]
 * (yyyy-mm-dd) và (tuỳ chọn) theo Tên khách hàng/người hưởng (khớp gần đúng, không phân biệt hoa/thường).
 * Mỗi dòng kết quả tương ứng 1 dòng thụ hưởng (ChiTietThuHuong) của các hồ sơ khớp điều kiện.
 * Trả về mảng: [{ soGiayNhanNo, ngayNhanNo, ngayGiaiNgan, tenNguoiHuong, soTienVay,
 *                  soTaiLieu, ngayTaiLieu, ngayDenHan, laiSuatTrongHan }, ...]
 */
/**
 * Báo cáo chi tiết tiền vay theo khế ước — có tự động thử lại 1 lần nếu lần đọc đầu tiên trả về
 * rỗng bất thường (Google Sheets đôi khi đọc lệch/thiếu dữ liệu tạm thời, đặc biệt ngay sau khi
 * vừa ghi/sửa dữ liệu ở 1 lệnh gọi khác trước đó).
 */
/**
 * Lọc báo cáo tiền vay theo cách "ghi ra sheet nháp rồi đọc lại ngay": tính kết quả, GHI vào 1 sheet
 * riêng (BaoCaoTienVay_Draft) trong CHÍNH file gốc — xoá sạch dữ liệu cũ trước khi ghi mới — rồi ĐỌC
 * LẠI NGAY LẬP TỨC từ chính sheet đó trong CÙNG 1 lần thực thi để trả về client.
 *
 * Vì ghi và đọc diễn ra trong cùng 1 lần gọi (không phải 2 lần google.script.run tách biệt như
 * layBaoCaoTienVay), cách này tránh hoàn toàn hiện tượng Google Sheets đọc dữ liệu bị trễ/lệch giữa
 * các lần gọi khác nhau — đây là phương án đáng tin cậy nhất, dùng cho bảng xem trực tiếp trên web.
 * Sheet nháp này cũng là 1 "bằng chứng" bạn có thể mở trực tiếp trong Google Sheets để tự kiểm tra.
 */
/**
 * Xây dựng lại TOÀN BỘ (không lọc) sheet nháp BaoCaoTienVay_Draft. Gọi mỗi khi Lưu hồ sơ hoặc bấm
 * "Đã hoàn thành giải ngân" để sheet này luôn mới nhất — không cần chờ người dùng vào tab Báo cáo.
 * "Lọc báo cáo" (locBaoCaoTienVayQuaSheet) CHỈ đọc + lọc trong bộ nhớ từ sheet này, không tính lại.
 */
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

  sh.getRange(1, 1, data.length, header.length).setValues(data);
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  SpreadsheetApp.flush();
  return sh;
}

/** Đọc toàn bộ dữ liệu từ sheet nháp BaoCaoTienVay_Draft (tự xây dựng nếu chưa có/còn trống). */
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

/** Chuyển chuỗi "dd/MM/yyyy" (định dạng lưu trong BaoCaoTienVay_Draft) thành Date để so sánh khoảng ngày. */
function parseNgayVN_(s) {
  if (!s || typeof s !== 'string') return null;
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

/**
 * Lọc báo cáo tiền vay — CHỈ đọc và lọc (trong bộ nhớ) từ sheet nháp BaoCaoTienVay_Draft, KHÔNG
 * tính toán lại từ HoSoGiaiNgan/ChiTietThuHuong mỗi lần lọc. Sheet nháp tự cập nhật khi Lưu hồ sơ
 * hoặc bấm "Đã hoàn thành giải ngân" — xem capNhatBaoCaoDraft_().
 */
function locBaoCaoTienVayQuaSheet(tuNgay, denNgay, tenKhachHang, trangThai) {
  // Luôn làm mới sheet nháp trước khi lọc — đảm bảo dữ liệu mới nhất kể cả khi hồ sơ được tạo/sửa
  // từ trước khi có cơ chế tự cập nhật draft (Lưu hồ sơ / Đã giải ngân), tránh đọc phải dữ liệu cũ.
  capNhatBaoCaoDraft_();
  var tatCa = docBaoCaoDraft_();

  var tu = tuNgay ? new Date(tuNgay + 'T00:00:00') : null;
  var den = denNgay ? new Date(denNgay + 'T23:59:59') : null;
  var tuKhoa = chuanHoaChuoi_(tenKhachHang || '');

  return tatCa.filter(function (r) {
    if (trangThai && r.trangThai !== trangThai) return false;
    if (tuKhoa && chuanHoaChuoi_(r.tenNguoiHuong).indexOf(tuKhoa) < 0) return false;
    if (tu || den) {
      var ngayGN = parseNgayVN_(r.ngayGiaiNgan);
      if (tu && ngayGN && ngayGN < tu) return false;
      if (den && ngayGN && ngayGN > den) return false;
    }
    return true;
  });
}

function layBaoCaoTienVay(tuNgay, denNgay, tenKhachHang, trangThai) {
  var out = layBaoCaoTienVay_(tuNgay, denNgay, tenKhachHang, trangThai);
  var lanThu = 0;
  while (out.length === 0 && lanThu < 2) {
    Utilities.sleep(400 + lanThu * 300);
    out = layBaoCaoTienVay_(tuNgay, denNgay, tenKhachHang, trangThai);
    lanThu++;
  }
  return out;
}

function layBaoCaoTienVay_(tuNgay, denNgay, tenKhachHang, trangThai) {
  var dsHoSo = sheetToObjects_(SHEET_HOSO);
  var dsChiTiet = sheetToObjects_(SHEET_CHITIET);
  var dsHopDong = sheetToObjects_(SHEET_HOPDONG);

  var tu = tuNgay ? new Date(tuNgay + 'T00:00:00') : null;
  var den = denNgay ? new Date(denNgay + 'T23:59:59') : null;
  var tuKhoa = chuanHoaChuoi_(tenKhachHang || '');

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
    if (trangThai && hs.TrangThai !== trangThai) return;

    var ngayGN = hs.NgayGiaiNgan instanceof Date ? hs.NgayGiaiNgan : (hs.NgayGiaiNgan ? new Date(hs.NgayGiaiNgan) : null);
    // Chỉ loại theo khoảng ngày khi hồ sơ CÓ Ngày giải ngân; hồ sơ chưa điền ngày này (VD: vừa đánh dấu
    // "Đã giải ngân" nhưng quên điền ngày) vẫn được giữ lại nếu chỉ lọc theo Trạng thái/Tên khách hàng.
    if (tu && ngayGN && ngayGN < tu) return;
    if (den && ngayGN && ngayGN > den) return;

    var hd = hopDongByMa[hs.MaHD];
    var laiSuat = hd ? hd['LaiSuatTrongHan_%'] : '';
    var soHopDong = hd ? hd.SoHopDong : '';

    var dsCt = chiTietByHoSo[String(hs.MaHoSo || '').trim()] || [];
    dsCt.forEach(function (ct) {
      if (tuKhoa && chuanHoaChuoi_(ct.TenNguoiHuong).indexOf(tuKhoa) < 0) return;
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

  out.sort(function (a, b) { return String(a.ngayGiaiNgan).localeCompare(String(b.ngayGiaiNgan)); });
  return out;
}
