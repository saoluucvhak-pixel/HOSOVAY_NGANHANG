/**
 * BACKUP.GS
 * Sao lưu định kỳ file Sheet chính ra 1 bản copy trong thư mục "Backup_Sheet".
 * Cách kích hoạt: Trình kích hoạt (⏰) > Thêm trình kích hoạt > chọn hàm saoLuuHangTuan
 * > Time-driven > Week timer > chọn ngày/giờ chạy (VD: Thứ Hai, 1h-2h sáng).
 */
function saoLuuHangTuan() {
  var ss = getSS_();
  var tenBanSao = 'BACKUP_' + ss.getName() + '_' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyyMMdd_HHmmss');
  var file = DriveApp.getFileById(ss.getId()).makeCopy(tenBanSao);

  var goc = DriveApp.getFolderById(THU_MUC_GOC_ID);
  var itBk = goc.getFoldersByName('Backup_Sheet');
  var thuMuc = itBk.hasNext() ? itBk.next() : goc.createFolder('Backup_Sheet');
  thuMuc.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch (e) { /* bỏ qua */ }

  don_DepBanSaoCu_(thuMuc, 12); // giữ tối đa 12 bản gần nhất, xoá bản cũ hơn
}

/** Chỉ giữ lại N bản backup gần nhất trong thư mục, xoá các bản cũ hơn để không đầy Drive. */
function don_DepBanSaoCu_(thuMuc, soBanGiuLai) {
  var files = [];
  var it = thuMuc.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    files.push({ file: f, ngay: f.getDateCreated().getTime() });
  }
  files.sort(function (a, b) { return b.ngay - a.ngay; }); // mới nhất trước
  for (var i = soBanGiuLai; i < files.length; i++) {
    files[i].file.setTrashed(true);
  }
}
