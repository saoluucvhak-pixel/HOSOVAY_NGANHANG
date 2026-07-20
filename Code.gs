/**
 * CODE.GS
 * Điểm vào chính của ứng dụng: menu trên Google Sheet + Web App (doGet).
 */

var SHEET_CONGTY = 'DM_CongTy';
var SHEET_HOPDONG = 'DM_HopDongVay';
var SHEET_KHACHHANG = 'DM_KhachHang';
var SHEET_HOSO = 'HoSoGiaiNgan';
var SHEET_CHITIET = 'ChiTietThuHuong';
var SHEET_BAOCAO_DRAFT = 'BaoCaoTienVay_Draft';

// File Google Sheet "Sổ chi tiết mua hàng" (nguồn tra cứu Số hoá đơn / Ngày tài liệu
// theo Nhà cung cấp ở mục 3 - Danh sách thụ hưởng).
// https://docs.google.com/spreadsheets/d/1KdbvfpXI3EBaFlq4bfv9ZSala-D3Pe-6tZZYQArt-vU/edit
var SS_MUA_HANG_ID = '1KdbvfpXI3EBaFlq4bfv9ZSala-D3Pe-6tZZYQArt-vU';
var SHEET_MUA_HANG = 'SO_CHI_TIET_MUA_HANG';

// Thư mục Google Drive gốc để lưu TẤT CẢ hồ sơ vay được tạo (Google Doc/PDF/báo cáo).
// https://drive.google.com/drive/folders/1OmF33Lb03Iu2Tzk0-08nBYFBY9u_OM3t
var THU_MUC_GOC_ID = '1OmF33Lb03Iu2Tzk0-08nBYFBY9u_OM3t';

// ===== File MẪU THẬT (theo đúng biểu mẫu VCB) — dùng để "copy mẫu rồi điền số liệu" thay vì dựng
// văn bản từ đầu bằng code. 2 file .docx có sẵn placeholder {{TenBien}}; 3 file .xlsx điền theo toạ
// độ ô cố định (không có placeholder, ghi trực tiếp vào ô tương ứng). =====
var TEMPLATE_VANBAN_DENGHI_ID = '14QYSWm3y3B2ijRcHD6bjCeJGOI9JMgSTzsVb3jBSOIc';   // Văn bản đề nghị giải ngân (Docs)
var TEMPLATE_GIAYNHANNO_ID = '1pUHGl27tJQFqJo40-62Nv6e3J_l5JLxEyy3XXn9SpCY';      // Giấy nhận nợ VCB (Docs)
var TEMPLATE_UNC_4LIEN_ID = '1Ia7Vvq6FAijahlCE2yV8YEz3bfE0Q-HvmL_p-vlpsdo';       // Uỷ nhiệm chi 4 liên (Sheets)
var TEMPLATE_BANGKE_HD_GIAINGAN_ID = '1krPKN68CQqjujx4ACHrpGAvFPsYfKYSLTacipm0WQcg'; // Bảng kê tài liệu chứng minh MĐSDV (Sheets)
var TEMPLATE_BANGKE_UNC_4LIEN_ID = '1PR7a2ccZWQUwW-FuVChcwIetUhoUAYIr4oguIZYCLEM'; // Phụ lục 02 - Bảng kê UNC (Sheets)

// Thư mục riêng để lưu 5 văn bản sinh ra từ mẫu thật (khác thư mục gốc ở trên theo yêu cầu).
// https://drive.google.com/drive/folders/1Pqt9uTZ5SIkS3RagpBUwKVtkzxofSYjl
var THU_MUC_XUAT_THEO_MAU_ID = '1Pqt9uTZ5SIkS3RagpBUwKVtkzxofSYjl';

/** Chạy khi mở Google Sheet: thêm menu tuỳ chỉnh. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Hồ Sơ Vay NH')
    .addItem('Mở ứng dụng (tạo hồ sơ giải ngân)', 'showApp')
    .addSeparator()
    .addItem('Khởi tạo / kiểm tra cấu trúc Sheet', 'initializeSpreadsheet')
    .addToUi();
}

/** Mở giao diện ứng dụng dạng cửa sổ (dialog) ngay trong Google Sheet. */
function showApp() {
  var html = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setWidth(1100)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'Tạo hồ sơ giải ngân & in ấn hồ sơ vay ngân hàng');
}

/**
 * doGet: cho phép truy cập như một Web App độc lập (khi deploy Publish > Deploy as web app).
 * Dùng chung 1 giao diện Index.html.
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Hồ sơ vay & Đề nghị giải ngân ngân hàng')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Cho phép include file .html khác (CSS/JS) vào Index.html bằng <?!= include('X'); ?> */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSS_() {
  // Ưu tiên SPREADSHEET_ID đã lưu cố định (được set khi chạy "Khởi tạo / kiểm tra cấu trúc Sheet").
  // Đây là cách xác định spreadsheet ĐÁNG TIN CẬY DUY NHẤT khi chạy qua Web App (doGet) —
  // SpreadsheetApp.getActiveSpreadsheet() có thể trả về kết quả KHÔNG NHẤT QUÁN (lúc có lúc không,
  // hoặc trỏ nhầm sang spreadsheet khác đang mở cùng tài khoản) khi chạy ở chế độ Web App độc lập,
  // dẫn tới tình trạng lưu vào 1 nơi nhưng đọc lại ở 1 nơi khác (hoặc rỗng).
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) {
    return SpreadsheetApp.openById(id);
  }
  // Chưa từng chạy "Khởi tạo" lần nào: thử dùng active spreadsheet (chỉ đúng khi đang chạy dạng
  // dialog gắn kèm trong Sheet), đồng thời tự lưu lại SPREADSHEET_ID để các lần gọi sau nhất quán.
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', active.getId());
    return active;
  }
  throw new Error('Chưa thiết lập SPREADSHEET_ID. Hãy mở Google Sheet, vào menu ' +
    '"Hồ Sơ Vay NH" > "Khởi tạo / kiểm tra cấu trúc Sheet" một lần trước khi dùng Web App độc lập.');
}
