/**
 * DOCGENERATOR.GS
 * Sinh các văn bản hồ sơ vay ngân hàng (Google Docs) từ dữ liệu trên Sheet,
 * theo đúng 5 mẫu: Giấy nhận nợ, Văn bản đề nghị giải ngân (Mẫu 13/PN-ĐNGN),
 * Ủy nhiệm chi, Bảng kê tài liệu chứng minh mục đích sử dụng vốn vay, Phụ lục 02.
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

function formatNgay_(v) {
  if (!v) return '…../…../20….';
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy');
  }
  return String(v);
}

function formatNgayDayDu_(v) {
  if (!v) return 'ngày .... tháng .... năm ........';
  var d = (v instanceof Date) ? v : new Date(v);
  if (isNaN(d.getTime())) return String(v);
  return 'ngày ' + Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'dd') +
    ' tháng ' + Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'MM') +
    ' năm ' + Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'yyyy');
}

// ---------- Khối tiêu đề quốc hiệu dùng chung ----------

function themQuocHieu_(body) {
  var p1 = body.appendParagraph('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM');
  p1.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(13);
  var p2 = body.appendParagraph('Độc lập - Tự do - Hạnh phúc');
  p2.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(12);
  var p3 = body.appendParagraph('--------------------');
  p3.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

function themTieuDeChinh_(body, text, soVanBan) {
  var t = body.appendParagraph(text);
  t.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(14);
  if (soVanBan) {
    var s = body.appendParagraph(soVanBan);
    s.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }
}

/** Gom các dòng [label, value] rồi tạo Table 1 lần duy nhất (tránh dòng rỗng đầu bảng). */
function BangKV_() {
  this.rows = [];
}
BangKV_.prototype.them = function (label, value) {
  this.rows.push([label, value == null ? '' : String(value)]);
  return this;
};
BangKV_.prototype.xuatRaBody = function (body) {
  if (!this.rows.length) return null;
  var table = body.appendTable(this.rows);
  table.setBorderWidth(0.75);
  for (var r = 0; r < table.getNumRows(); r++) {
    var cell0 = table.getRow(r).getCell(0);
    cell0.editAsText().setBold(true);
    cell0.setWidth(230);
  }
  return table;
};

// =========================================================================
// 1) GIẤY NHẬN NỢ
// =========================================================================

function addGiayNhanNoToBody_(body, d) {
  var hoSo = d.hoSo, hopDong = d.hopDong, congTy = d.congTy;

  themQuocHieu_(body);
  themTieuDeChinh_(body, 'GIẤY NHẬN NỢ', 'Số: ' + (hoSo.SoGiayNhanNo || '……………'));

  var kem = body.appendParagraph('Kèm theo Hợp đồng cho vay theo hạn mức số ' + hopDong.SoHopDong +
    ' ngày ' + formatNgay_(hopDong.NgayHopDong) + ' và các hợp đồng sửa đổi bổ sung (nếu có)');
  kem.setItalic(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('');

  var t1 = new BangKV_();
  t1.them('Chúng tôi là:', congTy.TenCty);
  t1.them('Mã CIF:', congTy.MaCIF);
  t1.them('Địa chỉ trụ sở chính:', congTy.DiaChiTruSo);
  t1.them('Điện thoại / Fax:', (congTy.DienThoai || '') + ' / ' + (congTy.Fax || ''));
  t1.them('Người đại diện:', congTy.NguoiDaiDien);
  t1.them('Chức vụ:', congTy.ChucVu);
  t1.them('Giấy ủy quyền (nếu có) số / ngày:',
    (congTy.GiayUyQuyenSo || '') + (congTy.GiayUyQuyenNgay ? (' - ' + formatNgay_(congTy.GiayUyQuyenNgay)) : ''));
  t1.xuatRaBody(body);

  body.appendParagraph('');
  var pgg = body.appendParagraph('Chúng tôi đồng ý nhận nợ với ' + hopDong.TenNganHang + ' – ' + hopDong.ChiNhanh +
    ' (sau đây gọi tắt là "Ngân hàng") theo Hợp đồng cho vay theo hạn mức số ' + hopDong.SoHopDong +
    ' ngày ' + formatNgay_(hopDong.NgayHopDong) + ' và các hợp đồng sửa đổi bổ sung (nếu có) giữa chúng tôi và ' +
    'Ngân hàng (sau đây gọi tắt là "Hợp đồng cho vay") với các nội dung như sau:');
  pgg.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
  body.appendParagraph('');

  var t2 = new BangKV_();
  t2.them('Số tiền vay theo Hợp đồng cho vay:',
    formatSo_(hopDong.HanMucVay) + ' VND (' + soThanhChuVN(hopDong.HanMucVay) + ')');
  t2.them('Dư nợ tới thời điểm hiện tại:', formatSo_(hoSo.DuNoHienTai) + ' VND');
  t2.them('Số tiền nhận nợ lần này:',
    formatSo_(hoSo.SoTienNhanNoLanNay) + ' VND (' + (hoSo.SoTienBangChu || soThanhChuVN(hoSo.SoTienNhanNoLanNay)) + ')');
  t2.them('Phương tiện thanh toán để giải ngân vốn cho vay:',
    (hoSo.PhuongThucThanhToan || 'Chuyển khoản') + ' – số tiền ' + formatSo_(hoSo.SoTienNhanNoLanNay) + ' VND ' +
    '(theo UNC/Bảng kê UNC đính kèm)');
  t2.them('*Mục đích sử dụng vốn vay:', hoSo.MucDichSuDungVon);
  t2.them('Thời hạn cho vay:', (hoSo.ThoiHanChoVay_Ngay || '……') + ' ngày');
  t2.them('Ngày giải ngân vốn cho vay:', formatNgay_(hoSo.NgayGiaiNgan));
  t2.them('Ngày đến hạn:', formatNgay_(hoSo.NgayDenHan));
  t2.them('Lãi suất cho vay trong hạn:', 'Cố định: ' + (hopDong['LaiSuatTrongHan_%'] || '……') + ' %/năm');
  t2.them('Lãi suất áp dụng đối với dư nợ gốc bị quá hạn:', hopDong.MoTaLaiSuatQuaHan);
  t2.them('Lãi suất áp dụng đối với lãi chậm trả:', (hopDong['LaiSuatLaiChamTra_%'] || 0) + ' %/năm trên số dư lãi chậm trả');
  t2.them('Kỳ hạn trả nợ:', 'Trả nợ gốc: ' + hopDong.KyHanTraGoc + '   |   Trả nợ lãi: ' + hopDong.KyHanTraLai);
  t2.them('Chứng từ kèm theo:', 'Ủy nhiệm chi ngày ' + formatNgay_(hoSo.NgayGiaiNgan));
  t2.them('Tài liệu chứng minh mục đích sử dụng vốn vay:', hoSo.TaiLieuChungMinhMucDich);
  t2.xuatRaBody(body);

  body.appendParagraph('');
  var ghino = body.appendParagraph('Số tiền chúng tôi nhận nợ theo các nội dung nêu trên được Ngân hàng hạch toán ' +
    'ghi nợ vào tài khoản vay của chúng tôi số: .................................................... tại Ngân hàng.');
  ghino.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

  var camket1 = body.appendParagraph('Chúng tôi cam kết sử dụng vốn vay đúng mục đích theo thỏa thuận với Ngân hàng; ' +
    'trả đầy đủ, đúng hạn nợ gốc, lãi tiền vay, phí và nghĩa vụ tài chính khác (nếu có) phù hợp với quy định tại Hợp đồng cho vay.');
  camket1.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

  body.appendPageBreak();

  var camdoan = body.appendParagraph('Chúng tôi cam đoan và cam kết các thông tin, tài liệu chứng minh mục đích sử dụng ' +
    'vốn vay trong Giấy nhận nợ này mà Chúng tôi cung cấp cho Ngân hàng là chính xác, phản ánh trung thực và hợp lý ' +
    'tình hình hoạt động của Chúng tôi. Việc nhận nợ theo Giấy nhận nợ này không vi phạm quy định nội bộ của chúng tôi ' +
    'cũng như quy định của Hợp đồng cho vay và quy định của pháp luật.');
  camdoan.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

  var camket2 = body.appendParagraph('Chúng tôi cam kết (i) Nội dung của (các) bản in hóa đơn điện tử/TKHQ điện tử ' +
    'và/hoặc hóa đơn điện tử chuyển đổi thành hóa đơn giấy khớp đúng, toàn vẹn các nội dung của (các) hóa đơn điện tử/' +
    'TKHQ điện tử; (ii) Sử dụng (các) hóa đơn điện tử/TKHQ điện tử này làm chứng từ chứng minh mục đích sử dụng vốn vay ' +
    'tại Ngân hàng đảm bảo không trùng lặp; (iii) Chịu trách nhiệm trước pháp luật và bồi thường tổn thất, thiệt hại ' +
    'phát sinh cho Ngân hàng trong trường hợp thông tin hóa đơn điện tử/TKHQ điện tử không trung thực, chính xác và ' +
    'đầy đủ và (iv) Trả nợ cho Ngân hàng trong trường hợp Ngân hàng sau tra cứu phát hiện hóa đơn/TKHQ không hợp lệ.');
  camket2.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

  var lapthanh = body.appendParagraph('Giấy nhận nợ này là một bộ phận không tách rời của Hợp đồng cho vay. Giấy nhận ' +
    'nợ này được lập thành 03 bản gốc, Ngân hàng giữ 02 bản gốc, Bên vay giữ 01 bản gốc, các bản gốc có giá trị pháp lý ngang nhau.');
  lapthanh.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

  body.appendParagraph('');
  var diaDiemNgay = body.appendParagraph('Đà Nẵng, ' + formatNgayDayDu_(hoSo.NgayGiayNhanNo));
  diaDiemNgay.setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setItalic(true);

  var kyTable = body.appendTable([
    ['XÁC NHẬN CỦA NGÂN HÀNG\nĐẠI DIỆN CỦA NGÂN HÀNG', 'NGƯỜI ĐẠI DIỆN HỢP PHÁP CỦA BÊN VAY\n(Ký, ghi rõ họ tên, chức vụ, đóng dấu)'],
  ]);
  kyTable.setBorderWidth(0);
  for (var i = 0; i < 2; i++) {
    kyTable.getRow(0).getCell(i).editAsText().setBold(true);
    kyTable.getRow(0).getCell(i).setPaddingTop(30);
  }
}

function taoGiayNhanNo(maHoSo) {
  var d = layDuLieuDayDu(maHoSo);
  var doc = DocumentApp.create('GiayNhanNo_' + maHoSo + '_' + d.hoSo.SoGiayNhanNo);
  addGiayNhanNoToBody_(doc.getBody(), d);
  doc.saveAndClose();
  dichChuyenVaoThuMuc_(doc.getId(), maHoSo);
  var url = doc.getUrl();
  capNhatLinkHoSo_(maHoSo, 'LinkGiayNhanNo', url);
  return url;
}

// =========================================================================
// 2) VĂN BẢN ĐỀ NGHỊ GIẢI NGÂN VỐN CHO VAY (Mẫu 13/PN-ĐNGN) — "Đề nghị thanh toán"
// =========================================================================

function addDeNghiGiaiNganToBody_(body, d) {
  var hoSo = d.hoSo, hopDong = d.hopDong, congTy = d.congTy, chiTiet = d.chiTiet;

  var mau = body.appendParagraph('Mẫu số 13/PN-ĐNGN: Văn bản của khách hàng là pháp nhân đề nghị giải ngân vốn cho vay');
  mau.setBold(true).setFontSize(10);
  var mauNote = body.appendParagraph('Ban hành kèm theo Công văn số 5110/VCB-PC ngày 04/04/2024 của Tổng giám đốc về ' +
    'việc Ban hành một số biểu mẫu liên quan đến nghiệp vụ cho vay');
  mauNote.setItalic(true).setFontSize(9);
  body.appendParagraph('');

  var headTable = body.appendTable([[congTy.TenCty, 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc']]);
  headTable.setBorderWidth(0);
  headTable.getRow(0).getCell(0).editAsText().setBold(true);
  headTable.getRow(0).getCell(1).editAsText().setBold(true);
  headTable.getRow(0).getCell(1).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);

  var soTable = body.appendTable([['Số: ..........', 'Đà Nẵng, ' + formatNgayDayDu_(hoSo.NgayGiaiNgan)]]);
  soTable.setBorderWidth(0);
  body.appendParagraph('V/v: Đề nghị giải ngân vốn cho vay').setBold(true);
  body.appendParagraph('');

  var kg = body.appendParagraph('Kính gửi: ' + hopDong.TenNganHang + ' – ' + hopDong.ChiNhanh);
  kg.setBold(true);
  var noiDung1 = body.appendParagraph('Thực hiện quy định tại Hợp đồng cho vay theo hạn mức số ' + hopDong.SoHopDong +
    ' ngày ' + formatNgay_(hopDong.NgayHopDong) + ' giữa ' + congTy.TenCty + ' và ' + hopDong.TenNganHang + ' – ' +
    hopDong.ChiNhanh + ' (sau đây gọi tắt là "Hợp đồng cho vay"), Chúng tôi đề nghị Quý Ngân hàng tiến hành giải ngân ' +
    'vốn cho vay như sau:');
  noiDung1.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

  body.appendParagraph('1. Tổng số tiền đề nghị giải ngân vốn cho vay là: ' + formatSo_(hoSo.SoTienNhanNoLanNay) +
    ' VND (' + (hoSo.SoTienBangChu || soThanhChuVN(hoSo.SoTienNhanNoLanNay)) + ')');
  body.appendParagraph('2. Ngày giải ngân vốn cho vay là: ' + formatNgay_(hoSo.NgayGiaiNgan));
  body.appendParagraph('3. Phương thức giải ngân vốn cho vay:');
  var laBangChuyenKhoan = String(hoSo.PhuongThucThanhToan || '').toLowerCase().indexOf('tiền mặt') === -1;
  body.appendParagraph('   ' + (laBangChuyenKhoan ? '☐' : '☒') + ' Bằng tiền mặt: số tiền giải ngân là [•] (Bằng chữ: [•]). ' +
    'Mục đích sử dụng: ' + hoSo.MucDichSuDungVon);
  body.appendParagraph('   ' + (laBangChuyenKhoan ? '☒' : '☐') + ' Bằng chuyển tiền vào tài khoản: chuyển tiền giải ngân ' +
    'vốn cho vay vào (các) tài khoản sau:');

  var rows = [['Tên tài khoản', 'Số tài khoản', 'Tại Ngân hàng', 'Mục đích sử dụng số tiền vay được giải ngân', 'Số tiền giải ngân']];
  var tong = 0;
  chiTiet.forEach(function (ct) {
    rows.push([ct.TenNguoiHuong, String(ct.SoTaiKhoan), ct.TaiNganHang, ct.NoiDungThanhToan, formatSo_(ct.SoTien)]);
    tong += Number(ct.SoTien) || 0;
  });
  rows.push(['Tổng số tiền đề nghị giải ngân', '', '', '', formatSo_(tong)]);
  var bTable = body.appendTable(rows);
  bTable.setBorderWidth(0.75);
  for (var c = 0; c < 5; c++) bTable.getRow(0).getCell(c).editAsText().setBold(true);
  var lastRowIdx = bTable.getNumRows() - 1;
  bTable.getRow(lastRowIdx).getCell(0).editAsText().setBold(true);
  bTable.getRow(lastRowIdx).getCell(4).editAsText().setBold(true);

  body.appendParagraph('');
  body.appendParagraph('4. Tài liệu chứng minh mục đích sử dụng vốn vay: ' + hoSo.TaiLieuChungMinhMucDich);
  var camdoan = body.appendParagraph('Chúng tôi cam đoan và cam kết các thông tin, tài liệu chứng minh mục đích sử ' +
    'dụng vốn vay kèm theo văn bản đề nghị giải ngân vốn vay lần này mà chúng tôi cung cấp cho Quý Ngân hàng là chính ' +
    'xác, phản ánh trung thực nhu cầu giải ngân vốn vay của Chúng tôi. Việc đề nghị giải ngân vốn cho vay theo văn bản ' +
    'này không vi phạm quy định nội bộ của chúng tôi cũng như quy định của Hợp đồng cho vay và quy định của pháp luật.');
  camdoan.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

  body.appendParagraph('');
  body.appendParagraph('Trân trọng.');
  var noiNhan = body.appendParagraph('Nơi nhận:\n- Như trên;\n- Lưu:');
  noiNhan.setItalic(true).setFontSize(10);

  body.appendParagraph('');
  var kyTen = body.appendParagraph('NGƯỜI ĐẠI DIỆN HỢP PHÁP\n(Ký, ghi rõ họ tên, chức vụ, đóng dấu (nếu có))');
  kyTen.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph('');
  body.appendParagraph('');
  body.appendParagraph('');
  var ten = body.appendParagraph(congTy.NguoiDaiDien + ' - ' + congTy.ChucVu);
  ten.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
}

function taoDeNghiGiaiNgan(maHoSo) {
  var d = layDuLieuDayDu(maHoSo);
  var doc = DocumentApp.create('DeNghiGiaiNgan_' + maHoSo);
  addDeNghiGiaiNganToBody_(doc.getBody(), d);
  doc.saveAndClose();
  dichChuyenVaoThuMuc_(doc.getId(), maHoSo);
  var url = doc.getUrl();
  capNhatLinkHoSo_(maHoSo, 'LinkDeNghiGiaiNgan', url);
  return url;
}

// =========================================================================
// 3) ỦY NHIỆM CHI (1 liên / 1 thụ hưởng)
// =========================================================================

function addUyNhiemChiMotLienToBody_(body, d, ct) {
  var hopDong = d.hopDong, congTy = d.congTy, hoSo = d.hoSo;

  var headTable = body.appendTable([[
    hopDong.ChiNhanh + '\n' + hopDong.DiaChiChiNhanh + '\nMST: ' + hopDong.MST_ChiNhanh +
    '\nSố tham chiếu: ' + (hoSo.SoThamChieu || ''),
    'CHỨNG TỪ GIAO DỊCH\nỦY NHIỆM CHI - PAYMENT ORDER\nNgày (Date): ' + formatNgay_(hoSo.NgayGiaiNgan)
  ]]);
  headTable.setBorderWidth(0.75);
  headTable.getRow(0).getCell(1).editAsText().setBold(true);

  var t = new BangKV_();
  t.them('ĐỀ NGHỊ GHI NỢ TÀI KHOẢN (Please debit account)', '');
  t.them('Số TK (A/c No.):', '');
  t.them('Tên TK (A/c name):', congTy.TenCty);
  t.them('Địa chỉ (address):', congTy.DiaChiTruSo);
  t.them('Tại NH (With bank):', hopDong.TenNganHang + ' - ' + hopDong.ChiNhanh);
  t.them('SỐ TIỀN (With amount) - Bằng số:', formatSo_(ct.SoTien));
  t.them('Bằng chữ (In words):', soThanhChuVN(ct.SoTien));
  t.them('& GHI CÓ TÀI KHOẢN (& Credit account)', '');
  t.them('Số TK (A/c No.):', ct.SoTaiKhoan);
  t.them('Tên TK (A/c name):', ct.TenNguoiHuong);
  t.them('Tại NH (With bank):', ct.TaiNganHang);
  t.them('Nội dung (Details of payment):', ct.NoiDungThanhToan);
  t.xuatRaBody(body);

  var kyTable = body.appendTable([['Kế toán trưởng ký\n(Chief accountant)', 'Chủ tài khoản ký & đóng dấu\n(A/c holder & stamp)']]);
  kyTable.setBorderWidth(0);
  for (var i = 0; i < 2; i++) kyTable.getRow(0).getCell(i).setPaddingTop(30);
}

function addUyNhiemChiToBody_(body, d) {
  d.chiTiet.forEach(function (ct, idx) {
    addUyNhiemChiMotLienToBody_(body, d, ct);
    if (idx < d.chiTiet.length - 1) {
      body.appendParagraph('').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      body.appendHorizontalRule();
    }
  });
}

function taoUyNhiemChi(maHoSo) {
  var d = layDuLieuDayDu(maHoSo);
  if (!d.chiTiet.length) throw new Error('Hồ sơ chưa có dòng thụ hưởng nào để lập Ủy nhiệm chi.');
  var doc = DocumentApp.create('UyNhiemChi_' + maHoSo);
  addUyNhiemChiToBody_(doc.getBody(), d);
  doc.saveAndClose();
  dichChuyenVaoThuMuc_(doc.getId(), maHoSo);
  var url = doc.getUrl();
  capNhatLinkHoSo_(maHoSo, 'LinkUyNhiemChi', url);
  return url;
}

// =========================================================================
// 4) BẢNG KÊ TÀI LIỆU CHỨNG MINH MỤC ĐÍCH SỬ DỤNG VỐN VAY
// =========================================================================

function addBangKeTaiLieuToBody_(body, d) {
  var congTy = d.congTy, hoSo = d.hoSo, chiTiet = d.chiTiet;

  var headTable = body.appendTable([[congTy.TenCty, 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n' +
    'Đà Nẵng, ' + formatNgayDayDu_(hoSo.NgayGiayNhanNo)]]);
  headTable.setBorderWidth(0);
  headTable.getRow(0).getCell(0).editAsText().setBold(true);
  headTable.getRow(0).getCell(1).editAsText().setBold(true);

  body.appendParagraph('');
  var title = body.appendParagraph('BẢNG KÊ TÀI LIỆU CHỨNG MINH MỤC ĐÍCH SỬ DỤNG VỐN VAY');
  title.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(13);
  var sub = body.appendParagraph('Theo Giấy nhận nợ số ' + hoSo.SoGiayNhanNo + ' ngày ' + formatNgay_(hoSo.NgayGiayNhanNo));
  sub.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('');

  var rows = [['STT', 'Tài liệu số', 'Ngày', 'Đơn vị lập', 'Nội dung', 'Số tiền vay', 'Ghi chú']];
  var tong = 0;
  chiTiet.forEach(function (ct, i) {
    rows.push([String(i + 1), ct.TaiLieuSo, ct.NgayTaiLieu, ct.DonViLapTaiLieu || ct.TenNguoiHuong,
      ct.NoiDungThanhToan, formatSo_(ct.SoTien), ct.GhiChu || '']);
    tong += Number(ct.SoTien) || 0;
  });
  rows.push(['', '', '', '', 'TỔNG CỘNG', formatSo_(tong), '']);
  var table = body.appendTable(rows);
  table.setBorderWidth(0.75);
  for (var c = 0; c < 7; c++) table.getRow(0).getCell(c).editAsText().setBold(true);
  var last = table.getNumRows() - 1;
  table.getRow(last).getCell(4).editAsText().setBold(true);
  table.getRow(last).getCell(5).editAsText().setBold(true);

  body.appendParagraph('');
  var camdoan = body.appendParagraph('Chúng tôi cam đoan và cam kết các thông tin, tài liệu chứng minh mục đích sử ' +
    'dụng vốn vay mà chúng tôi cung cấp cho Quý Ngân hàng là chính xác, phản ánh trung thực và hợp lý tình hình hoạt ' +
    'động của chúng tôi. Đối với khoản giải ngân bằng phương tiện tiền mặt, chúng tôi cam kết sẽ bổ sung tài liệu ' +
    'chứng minh mục đích sử dụng vốn vay theo yêu cầu của Quý Ngân hàng.');
  camdoan.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

  body.appendParagraph('');
  body.appendParagraph('Trân trọng.');
  var lap = body.appendParagraph('Người lập biểu');
  lap.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('');
  body.appendParagraph('');
  var ten = body.appendParagraph(hoSo.NguoiLapBieu || '');
  ten.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

function taoBangKeTaiLieu(maHoSo) {
  var d = layDuLieuDayDu(maHoSo);
  var doc = DocumentApp.create('BangKeTaiLieu_' + maHoSo);
  addBangKeTaiLieuToBody_(doc.getBody(), d);
  doc.saveAndClose();
  dichChuyenVaoThuMuc_(doc.getId(), maHoSo);
  var url = doc.getUrl();
  capNhatLinkHoSo_(maHoSo, 'LinkBangKeTaiLieu', url);
  return url;
}

// =========================================================================
// 5) PHỤ LỤC 02 - MẪU DANH SÁCH THANH TOÁN
// =========================================================================

function addPhuLuc02ToBody_(body, d) {
  var hoSo = d.hoSo, chiTiet = d.chiTiet;

  var pl = body.appendParagraph('PHỤ LỤC 02');
  pl.setBold(true);
  var stc = body.appendParagraph('Số tham chiếu: ' + (hoSo.SoThamChieu || ''));
  stc.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  var title = body.appendParagraph('MẪU DANH SÁCH THANH TOÁN');
  title.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(13);
  body.appendParagraph('');

  var rows = [['STT', 'Số tài khoản / Số giấy tờ tùy thân', 'Loại tiền', 'Tên người hưởng', 'Số tiền',
    'Nội dung', 'Chuyển tiền nhanh 24/7', 'Tại Ngân hàng', 'Ngày cấp giấy tờ tùy thân', 'Ghi chú']];
  var tong = 0;
  chiTiet.forEach(function (ct, i) {
    rows.push([String(i + 1), String(ct.SoTaiKhoan), ct.LoaiTien || 'VND', ct.TenNguoiHuong,
      formatSo_(ct.SoTien), ct.NoiDungThanhToan, ct.ChuyenTienNhanh_24_7 || 'N', ct.TaiNganHang,
      ct.NgayCapGiayToTuyThan || '', ct.GhiChu || '']);
    tong += Number(ct.SoTien) || 0;
  });
  rows.push(['', '', '', 'Tổng cộng', formatSo_(tong), '', '', '', '', '']);
  var table = body.appendTable(rows);
  table.setBorderWidth(0.75);
  for (var c = 0; c < 10; c++) {
    table.getRow(0).getCell(c).editAsText().setBold(true).setFontSize(9);
  }
  var last = table.getNumRows() - 1;
  table.getRow(last).getCell(3).editAsText().setBold(true);
  table.getRow(last).getCell(4).editAsText().setBold(true);

  body.appendParagraph('');
  var kyTable = body.appendTable([['NGƯỜI LẬP', 'CHỦ TÀI KHOẢN']]);
  kyTable.setBorderWidth(0);
  for (var i = 0; i < 2; i++) {
    kyTable.getRow(0).getCell(i).editAsText().setBold(true);
    kyTable.getRow(0).getCell(i).setPaddingTop(30);
  }
  body.appendParagraph('');
  var ten2 = body.appendParagraph(hoSo.NguoiLapBieu || '');
  ten2.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
}

function taoPhuLuc02(maHoSo) {
  var d = layDuLieuDayDu(maHoSo);
  var doc = DocumentApp.create('PhuLuc02_DanhSachThanhToan_' + maHoSo);
  addPhuLuc02ToBody_(doc.getBody(), d);
  doc.saveAndClose();
  dichChuyenVaoThuMuc_(doc.getId(), maHoSo);
  var url = doc.getUrl();
  capNhatLinkHoSo_(maHoSo, 'LinkPhuLuc02', url);
  return url;
}

// =========================================================================
// TẠO TRỌN BỘ HỒ SƠ (5 mẫu) — sinh riêng từng Doc + 1 PDF gộp
// =========================================================================

function taoTronBoHoSo(maHoSo) {
  var d = layDuLieuDayDu(maHoSo);

  var linkGiayNhanNo = taoGiayNhanNo(maHoSo);
  var linkDeNghi = taoDeNghiGiaiNgan(maHoSo);
  var linkUNC = d.chiTiet.length ? taoUyNhiemChi(maHoSo) : '';
  var linkBangKe = taoBangKeTaiLieu(maHoSo);
  var linkPhuLuc02 = taoPhuLuc02(maHoSo);

  // Doc gộp để in trọn bộ 1 lần
  var docGop = DocumentApp.create('HoSoVay_TRONBO_' + maHoSo);
  var body = docGop.getBody();
  addGiayNhanNoToBody_(body, d);
  body.appendPageBreak();
  addDeNghiGiaiNganToBody_(body, d);
  if (d.chiTiet.length) {
    body.appendPageBreak();
    addUyNhiemChiToBody_(body, d);
  }
  body.appendPageBreak();
  addBangKeTaiLieuToBody_(body, d);
  body.appendPageBreak();
  addPhuLuc02ToBody_(body, d);
  docGop.saveAndClose();
  dichChuyenVaoThuMuc_(docGop.getId(), maHoSo);

  var pdfUrl = xuatPdf_(docGop.getId(), 'HoSoVay_TRONBO_' + maHoSo, maHoSo);
  capNhatLinkHoSo_(maHoSo, 'LinkPDF_TrongBo', pdfUrl);

  return {
    giayNhanNo: linkGiayNhanNo,
    deNghiGiaiNgan: linkDeNghi,
    uyNhiemChi: linkUNC,
    bangKeTaiLieu: linkBangKe,
    phuLuc02: linkPhuLuc02,
    docGop: docGop.getUrl(),
    pdf: pdfUrl
  };
}

// ---------- Drive: thư mục & xuất PDF ----------

function layThuMucGoc_() {
  return DriveApp.getFolderById(THU_MUC_GOC_ID);
}

function layThuMucHoSo_(maHoSo) {
  var goc = layThuMucGoc_();
  var it = goc.getFoldersByName(maHoSo);
  if (it.hasNext()) return it.next();
  return goc.createFolder(maHoSo);
}

function dichChuyenVaoThuMuc_(fileId, maHoSo) {
  var folder = layThuMucHoSo_(maHoSo);
  var file = DriveApp.getFileById(fileId);
  folder.addFile(file);
  var root = DriveApp.getRootFolder();
  try { root.removeFile(file); } catch (e) { /* file có thể đã không còn ở root, bỏ qua */ }
}

function xuatPdf_(docId, tenGoi, maHoSo) {
  var file = DriveApp.getFileById(docId);
  var blob = file.getAs(MimeType.PDF).setName(tenGoi + '.pdf');
  var pdfFile = DriveApp.createFile(blob);
  var folder = layThuMucHoSo_(maHoSo);
  folder.addFile(pdfFile);
  DriveApp.getRootFolder().removeFile(pdfFile);
  return pdfFile.getUrl();
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
      String(i + 1), String(r.maHoSo || ''), String(r.soHopDong || ''), String(r.soGiayNhanNo || ''), String(r.ngayNhanNo || ''),
      String(r.ngayGiaiNgan || ''), String(r.tenNguoiHuong || ''), formatSo_(r.soTienVay),
      String(r.soTaiLieu || ''), String(r.ngayTaiLieu || ''), String(r.ngayDenHan || ''),
      String(r.laiSuatTrongHan || ''), String(r.trangThai || '')
    ]);
    tong += Number(r.soTienVay) || 0;
  });
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
 * Xuất báo cáo tiền vay thành 1 file Excel (.xlsx) thật, lưu vào thư mục con "BaoCao_TienVay".
 * Trả về { xlsxUrl, sheetUrl }.
 */
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
  sh.getRange(startRow, 1, data.length, header.length).setValues(data);
  sh.getRange(startRow, 1, 1, header.length).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  sh.getRange(startRow + data.length - 1, 7, 1, 2).setFontWeight('bold');
  sh.autoResizeColumns(1, header.length);
  // Bật bộ lọc (AutoFilter) ngay trên file — lọc/sắp xếp trực tiếp trong Excel/Google Sheets sẽ
  // đáng tin cậy hơn lọc "live" nhiều lần qua Apps Script (không phụ thuộc độ trễ đọc Google Sheets).
  try {
    sh.getRange(startRow, 1, data.length - 1, header.length).createFilter();
  } catch (e) { /* không chặn nếu tạo filter lỗi */ }
  sh.setFrozenRows(startRow);
  SpreadsheetApp.flush();

  var folder = layThuMucBaoCao_();
  var file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch (e) { /* bỏ qua */ }

  // Xuất thành file .xlsx thật (không chỉ Google Sheet) rồi lưu cùng thư mục.
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
