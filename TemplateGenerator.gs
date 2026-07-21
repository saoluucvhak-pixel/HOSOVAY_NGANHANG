/**
 * TEMPLATEGENERATOR.GS
 * Sinh 5 văn bản hồ sơ vay bằng cách COPY file mẫu thật (đúng định dạng ngân hàng đang dùng) rồi
 * điền số liệu vào — thay cho cách dựng văn bản từ đầu bằng code (DocGenerator.gs, giữ lại làm
 * phương án dự phòng). 2 file .docx dùng placeholder {{TenBien}}; 3 file .xlsx điền trực tiếp theo
 * toạ độ ô cố định trong mẫu.
 */

// ---------- Tiện ích dùng chung ----------

/** Tên file mới = tênGốc_MãHồSơ_ngàyLưu (ddMMyyyy). */
function tenFileTheoMau_(ten, maHoSo) {
  var ngay = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'ddMMyyyy');
  return ten + '_' + maHoSo + '_' + ngay;
}

/** Thư mục con theo Mã hồ sơ, nằm trong thư mục xuất báo cáo theo mẫu (THU_MUC_XUAT_THEO_MAU_ID). */
function layThuMucXuatTheoMau_(maHoSo) {
  var goc = DriveApp.getFolderById(THU_MUC_XUAT_THEO_MAU_ID);
  var it = goc.getFoldersByName(maHoSo);
  if (it.hasNext()) return it.next();
  return goc.createFolder(maHoSo);
}

/** Thay {{key}} bằng value trong toàn bộ Document, tự escape ký tự $ (replaceText hiểu $ là backreference). */
function thayTheAnToan_(body, key, value) {
  var v = (value === undefined || value === null) ? '' : String(value);
  v = v.replace(/\$/g, '$$$$');
  body.replaceText('\\{\\{' + key + '\\}\\}', v);
}

function thayNhieuTrongDoc_(body, thayThe) {
  Object.keys(thayThe).forEach(function (key) {
    thayTheAnToan_(body, key, thayThe[key]);
  });
}

/** "dd tháng MM năm yyyy" — KHÔNG có chữ "ngày" phía trước (mẫu Giấy nhận nợ đã có sẵn chữ "ngày"). */
function ngayThangNamKhongTien_(v) {
  if (!v) return '.... tháng .... năm ........';
  var d = (v instanceof Date) ? v : new Date(v);
  if (isNaN(d.getTime())) return String(v);
  return Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'dd') +
    ' tháng ' + Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'MM') +
    ' năm ' + Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'yyyy');
}

/** Chuyển 1 file Doc/Sheet vừa tạo thành PDF, lưu cùng thư mục, trả về URL PDF. */
function xuatPdfTheoMau_(fileId, tenGoc, folder) {
  var blob = DriveApp.getFileById(fileId).getAs(MimeType.PDF).setName(tenGoc + '.pdf');
  var pdfFile = folder.createFile(blob);
  return pdfFile.getUrl();
}

/**
 * Chỉnh số dòng của 1 bảng dữ liệu trong Sheet cho khớp số dòng thật cần có — chèn thêm dòng nếu
 * cần nhiều hơn mẫu, xoá bớt nếu cần ít hơn — để các nội dung PHÍA DƯỚI bảng (cam đoan, chữ ký...)
 * tự dịch chuyển theo đúng vị trí, không bị đè lên nhau.
 */
function dieuChinhSoDongBang_(sh, dongDauDuLieu, soDongMauHienTai, soDongCanCo) {
  var chenhLech = soDongCanCo - soDongMauHienTai;
  if (chenhLech > 0) {
    sh.insertRowsBefore(dongDauDuLieu + soDongMauHienTai, chenhLech);
  } else if (chenhLech < 0) {
    sh.deleteRows(dongDauDuLieu + soDongCanCo, -chenhLech);
  }
  return chenhLech;
}

// =========================================================================
// 1. GIẤY NHẬN NỢ (Docs, placeholder {{...}})
// =========================================================================
function taoGiayNhanNoTheoMau(maHoSo) {
  var data = layDuLieuDayDu(maHoSo);
  var hoSo = data.hoSo, hopDong = data.hopDong, congTy = data.congTy;

  var folder = layThuMucXuatTheoMau_(maHoSo);
  var tenFile = tenFileTheoMau_('GIAYNHANNO_VCB', maHoSo);
  var copy = DriveApp.getFileById(TEMPLATE_GIAYNHANNO_ID).makeCopy(tenFile, folder);
  var doc = DocumentApp.openById(copy.getId());
  var body = doc.getBody();

  thayNhieuTrongDoc_(body, {
    SoGiayNhanNo: hoSo.SoGiayNhanNo || '',
    SoHopDong: hopDong.SoHopDong || '',
    NgayHopDong: formatNgay_(hopDong.NgayHopDong),
    TenCongTy: congTy.TenCty || '',
    MaCIF: congTy.MaCIF || '',
    DiaChi: congTy.DiaChiTruSo || '',
    DienThoai: congTy.DienThoai || '',
    NguoiDaiDien: congTy.NguoiDaiDien || '',
    ChucVu: congTy.ChucVu || '',
    DuNoHienTai: formatSo_(hoSo.DuNoHienTai),
    SoTienVay_So: formatSo_(hoSo.SoTienNhanNoLanNay),
    SoTienVay_Chu: soThanhChuVN(hoSo.SoTienNhanNoLanNay),
    MucDichVay: hoSo.MucDichSuDungVon || '',
    ThoiHanVay: hoSo.ThoiHanChoVay_Ngay || '',
    NgayGiaiNgan: formatNgay_(hoSo.NgayGiaiNgan),
    LaiSuat: (hopDong['LaiSuatTrongHan_%'] || '') + '%',
    HanMucVay: formatSo_(hopDong.HanMucVay),
    NgayDenHan: formatNgay_(hoSo.NgayDenHan),
    NgayGiayNhanNo: ngayThangNamKhongTien_(hoSo.NgayGiayNhanNo)
  });

  doc.saveAndClose();
  var pdfUrl = xuatPdfTheoMau_(copy.getId(), tenFile, folder);
  return { docUrl: copy.getUrl(), pdfUrl: pdfUrl, ten: tenFile };
}

// =========================================================================
// 2. VĂN BẢN ĐỀ NGHỊ GIẢI NGÂN (Docs, placeholder {{...}} + bảng nhiều dòng thụ hưởng)
// =========================================================================
function taoVanBanDeNghiTheoMau(maHoSo) {
  var data = layDuLieuDayDu(maHoSo);
  var hoSo = data.hoSo, hopDong = data.hopDong, chiTiet = data.chiTiet || [];

  var folder = layThuMucXuatTheoMau_(maHoSo);
  var tenFile = tenFileTheoMau_('VANBANDENGHIGIAINGAN', maHoSo);
  var copy = DriveApp.getFileById(TEMPLATE_VANBAN_DENGHI_ID).makeCopy(tenFile, folder);
  var doc = DocumentApp.openById(copy.getId());
  var body = doc.getBody();

  // Các placeholder không nằm trong bảng thụ hưởng — thay trực tiếp.
  thayNhieuTrongDoc_(body, {
    SoHopDong: hopDong.SoHopDong || '',
    NgayHopDong: formatNgay_(hopDong.NgayHopDong),
    SoTienVay_So: formatSo_(hoSo.SoTienNhanNoLanNay),
    SoTienVay_Chu: soThanhChuVN(hoSo.SoTienNhanNoLanNay),
    NgayGiaiNgan: formatNgay_(hoSo.NgayGiaiNgan)
  });

  // Bảng thụ hưởng: tìm dòng còn placeholder {{TenNguoiHuong}}, nhân dòng theo số người thụ hưởng.
  var tongTien = dienBangThuHuongVanBanDeNghi_(body, chiTiet);
  // Placeholder tổng tiền còn sót lại trong bảng (nếu mẫu có) — không bắt buộc có, bỏ qua nếu không tìm thấy.
  thayTheAnToan_(body, 'TongTienGiaiNgan', formatSo_(tongTien));

  doc.saveAndClose();
  var pdfUrl = xuatPdfTheoMau_(copy.getId(), tenFile, folder);
  return { docUrl: copy.getUrl(), pdfUrl: pdfUrl, ten: tenFile };
}

function dienBangThuHuongVanBanDeNghi_(body, chiTiet) {
  var tables = body.getTables();
  var bang = null, soHangMau = -1;
  for (var t = 0; t < tables.length; t++) {
    var tbl = tables[t];
    for (var r = 0; r < tbl.getNumRows(); r++) {
      if (tbl.getRow(r).getText().indexOf('{{TenNguoiHuong}}') >= 0) {
        bang = tbl; soHangMau = r; break;
      }
    }
    if (bang) break;
  }
  if (!bang || !chiTiet.length) return 0;

  var soCotMau = bang.getRow(soHangMau).getNumCells();
  var tongTien = 0;

  chiTiet.forEach(function (ct, idx) {
    var hang;
    if (idx === 0) {
      hang = bang.getRow(soHangMau);
    } else {
      hang = bang.insertTableRow(soHangMau + idx);
      for (var cc = 0; cc < soCotMau; cc++) hang.appendTableCell('');
    }
    if (soCotMau > 0) hang.getCell(0).setText(ct.TenNguoiHuong || '');
    if (soCotMau > 1) hang.getCell(1).setText(ct.SoTaiKhoan || '');
    if (soCotMau > 2) hang.getCell(2).setText(ct.TaiNganHang || '');
    if (soCotMau > 3) hang.getCell(3).setText(ct.NoiDungThanhToan || '');
    if (soCotMau > 4) hang.getCell(4).setText(formatSo_(ct.SoTien));
    tongTien += Number(ct.SoTien) || 0;
  });

  return tongTien;
}

// =========================================================================
// 3. UỶ NHIỆM CHI 4 LIÊN (Sheets — 2 khối giống hệt nhau trong cùng 1 sheet, điền cả 2)
// =========================================================================
function taoUNCTheoMau(maHoSo) {
  var data = layDuLieuDayDu(maHoSo);
  var hoSo = data.hoSo, congTy = data.congTy;

  var folder = layThuMucXuatTheoMau_(maHoSo);
  var tenFile = tenFileTheoMau_('UNC_4lien', maHoSo);
  var copy = DriveApp.getFileById(TEMPLATE_UNC_4LIEN_ID).makeCopy(tenFile, folder);
  var ss = SpreadsheetApp.openById(copy.getId());
  var sh = ss.getSheets()[0];

  var ngayGN = 'Ngày ( Date): ' + formatNgay_(hoSo.NgayGiaiNgan);
  var soThamChieu = 'Số tham chiếu LTT: ' + (hoSo.SoThamChieu || '');
  var soTienChu = soThanhChuVN(hoSo.SoTienNhanNoLanNay);
  var soTien = Number(hoSo.SoTienNhanNoLanNay) || 0;

  [{ ngay: 'G3', tk: 'B5', ten: 'D8', tien: 'J7', chu: 'J8', diaChi: 'D9' },
    { ngay: 'G27', tk: 'B29', ten: 'D32', tien: 'J31', chu: 'J32', diaChi: 'D33' }
  ].forEach(function (o) {
    sh.getRange(o.ngay).setValue(ngayGN);
    sh.getRange(o.tk).setValue(soThamChieu);
    sh.getRange(o.ten).setValue(congTy.TenCty || '');
    sh.getRange(o.tien).setValue(soTien);
    sh.getRange(o.chu).setValue(soTienChu);
    sh.getRange(o.diaChi).setValue(congTy.DiaChiTruSo || '');
  });

  SpreadsheetApp.flush();
  var pdfUrl = xuatPdfTheoMauSheet_(copy.getId(), sh.getSheetId(), tenFile, folder);
  return { sheetUrl: copy.getUrl(), pdfUrl: pdfUrl, ten: tenFile };
}

// =========================================================================
// 4. BẢNG KÊ TÀI LIỆU CHỨNG MINH MỤC ĐÍCH SỬ DỤNG VỐN VAY (Sheets, bảng nhiều dòng)
// =========================================================================
function taoBangKeTaiLieuTheoMau(maHoSo) {
  var data = layDuLieuDayDu(maHoSo);
  var hoSo = data.hoSo, congTy = data.congTy, chiTiet = data.chiTiet || [];

  var folder = layThuMucXuatTheoMau_(maHoSo);
  var tenFile = tenFileTheoMau_('BangKe_HD_GiaiNgan', maHoSo);
  var copy = DriveApp.getFileById(TEMPLATE_BANGKE_HD_GIAINGAN_ID).makeCopy(tenFile, folder);
  var ss = SpreadsheetApp.openById(copy.getId());
  var sh = ss.getSheetByName('Sheet1 (2)') || ss.getSheets()[0];

  sh.getRange('A1').setValue(congTy.TenCty || '');
  sh.getRange('D3').setValue('Đà Nẵng, ' + formatNgayDayDu_(new Date()));
  sh.getRange('A5').setValue('BẢNG KÊ TÀI LIỆU CHỨNG MINH MỤC ĐÍCH SỬ DỤNG VỐN VAY\nTheo Giấy nhận nợ số ' +
    (hoSo.SoGiayNhanNo || '') + ' ngày ' + formatNgay_(hoSo.NgayGiayNhanNo));

  var dongDauDuLieu = 7, soDongMau = 2;
  var soDongCanCo = Math.max(chiTiet.length, 1);
  var chenhLech = dieuChinhSoDongBang_(sh, dongDauDuLieu, soDongMau, soDongCanCo);

  var rows = [];
  var tong = 0;
  chiTiet.forEach(function (ct, i) {
    var tien = Number(ct.SoTien) || 0;
    tong += tien;
    rows.push([i + 1, ct.TaiLieuSo || '', ct.NgayTaiLieu || '', ct.DonViLapTaiLieu || ct.TenNguoiHuong || '',
      ct.NoiDungThanhToan || '', tien, ct.GhiChu || '']);
  });
  if (!rows.length) rows.push(['', '', '', '', '', 0, '']);
  sh.getRange(dongDauDuLieu, 1, rows.length, 7).setValues(rows);

  var dongTongCong = dongDauDuLieu + soDongCanCo;
  sh.getRange(dongTongCong, 1).setValue('TỔNG CỘNG');
  sh.getRange(dongTongCong, 6).setValue(tong);

  var dongNguoiLap = 22 + chenhLech;
  sh.getRange('B' + dongNguoiLap).setValue(hoSo.NguoiLapBieu || '');

  SpreadsheetApp.flush();
  var pdfUrl = xuatPdfTheoMauSheet_(copy.getId(), sh.getSheetId(), tenFile, folder);
  return { sheetUrl: copy.getUrl(), pdfUrl: pdfUrl, ten: tenFile };
}

// =========================================================================
// 5. PHỤ LỤC 02 - BẢNG KÊ DANH SÁCH THANH TOÁN / UỶ NHIỆM CHI (Sheets, bảng nhiều dòng)
// =========================================================================
function taoBangKeUNCTheoMau(maHoSo) {
  var data = layDuLieuDayDu(maHoSo);
  var hoSo = data.hoSo, chiTiet = data.chiTiet || [];

  var folder = layThuMucXuatTheoMau_(maHoSo);
  var tenFile = tenFileTheoMau_('BangKe_UNC_4lien', maHoSo);
  var copy = DriveApp.getFileById(TEMPLATE_BANGKE_UNC_4LIEN_ID).makeCopy(tenFile, folder);
  var ss = SpreadsheetApp.openById(copy.getId());
  var sh = ss.getSheets()[0];

  sh.getRange('A2').setValue('Số tham chiếu LTT: ' + (hoSo.SoThamChieu || ''));

  var dongDauDuLieu = 7, soDongMau = 2;
  var soDongCanCo = Math.max(chiTiet.length, 1);
  var chenhLech = dieuChinhSoDongBang_(sh, dongDauDuLieu, soDongMau, soDongCanCo);

  var rows = [];
  var tong = 0;
  chiTiet.forEach(function (ct, i) {
    var tien = Number(ct.SoTien) || 0;
    tong += tien;
    rows.push([i + 1, ct.SoTaiKhoan || '', 'VND', ct.TenNguoiHuong || '', tien, ct.NoiDungThanhToan || '',
      ct.ChuyenTienNhanh_24_7 === 'Y' ? 'Y' : 'N', ct.TaiNganHang || '', ct.NgayCapGiayToTuyThan || '', ct.GhiChu || '']);
  });
  if (!rows.length) rows.push(['', '', 'VND', '', 0, '', 'N', '', '', '']);
  sh.getRange(dongDauDuLieu, 1, rows.length, 10).setValues(rows);

  var dongTongCong = dongDauDuLieu + soDongCanCo;
  sh.getRange(dongTongCong, 1).setValue('Tổng cộng');
  sh.getRange(dongTongCong, 5).setValue(tong);

  var dongNguoiLap = 18 + chenhLech;
  sh.getRange('B' + dongNguoiLap).setValue(hoSo.NguoiLapBieu || '');

  SpreadsheetApp.flush();
  var pdfUrl = xuatPdfTheoMauSheet_(copy.getId(), sh.getSheetId(), tenFile, folder);
  return { sheetUrl: copy.getUrl(), pdfUrl: pdfUrl, ten: tenFile };
}

/** Xuất 1 sheet (theo gid) trong file Google Sheet thành PDF qua URL export, lưu cùng thư mục. */
function xuatPdfTheoMauSheet_(spreadsheetId, sheetId, tenGoc, folder) {
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=pdf&gid=' + sheetId +
    '&portrait=false&fitw=true&gridlines=false';
  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) return '';
  var pdfFile = folder.createFile(resp.getBlob().setName(tenGoc + '.pdf'));
  return pdfFile.getUrl();
}

// =========================================================================
// TẠO TRỌN BỘ 5 VĂN BẢN (theo mẫu thật) CHO 1 HỒ SƠ
// =========================================================================
function taoTronBoTheoMau(maHoSo) {
  var giayNhanNo = taoGiayNhanNoTheoMau(maHoSo);
  var vanBanDeNghi = taoVanBanDeNghiTheoMau(maHoSo);
  var unc = taoUNCTheoMau(maHoSo);
  var bangKeTaiLieu = taoBangKeTaiLieuTheoMau(maHoSo);
  var bangKeUNC = taoBangKeUNCTheoMau(maHoSo);

  // Lưu lại link để xem/in lại sau này (không cần tạo lại mỗi lần mở hồ sơ).
  capNhatLinkHoSo_(maHoSo, 'Link_GiayNhanNo', giayNhanNo.docUrl);
  capNhatLinkHoSo_(maHoSo, 'Link_GiayNhanNo_PDF', giayNhanNo.pdfUrl);
  capNhatLinkHoSo_(maHoSo, 'Link_VanBanDeNghi', vanBanDeNghi.docUrl);
  capNhatLinkHoSo_(maHoSo, 'Link_VanBanDeNghi_PDF', vanBanDeNghi.pdfUrl);
  capNhatLinkHoSo_(maHoSo, 'Link_UNC', unc.sheetUrl);
  capNhatLinkHoSo_(maHoSo, 'Link_UNC_PDF', unc.pdfUrl);
  capNhatLinkHoSo_(maHoSo, 'Link_BangKeTaiLieu', bangKeTaiLieu.sheetUrl);
  capNhatLinkHoSo_(maHoSo, 'Link_BangKeTaiLieu_PDF', bangKeTaiLieu.pdfUrl);
  capNhatLinkHoSo_(maHoSo, 'Link_BangKeUNC', bangKeUNC.sheetUrl);
  capNhatLinkHoSo_(maHoSo, 'Link_BangKeUNC_PDF', bangKeUNC.pdfUrl);

  // Chuyển trạng thái sang "Đã tạo hồ sơ" (không hạ cấp nếu lỡ đã "Đã giải ngân").
  var hs = layHoSoTheoMa(maHoSo);
  var trangThaiMoi = (hs && hs.TrangThai === 'Đã giải ngân') ? 'Đã giải ngân' : 'Đã tạo hồ sơ';
  capNhatLinkHoSo_(maHoSo, 'TrangThai', trangThaiMoi);

  return {
    giayNhanNo: giayNhanNo,
    vanBanDeNghi: vanBanDeNghi,
    unc: unc,
    bangKeTaiLieu: bangKeTaiLieu,
    bangKeUNC: bangKeUNC,
    thuMucUrl: layThuMucXuatTheoMau_(maHoSo).getUrl(),
    trangThai: trangThaiMoi
  };
}
