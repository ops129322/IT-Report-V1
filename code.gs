// --- CẤU HÌNH HỆ THỐNG ---
const ADMIN_LIST = [
  'huythanh.pham@spxpress.com', // Email chính của bạn
  'admin2@spxpress.com'         // Thêm các email admin khác vào đây
];

function doGet(e) {
  var userEmail = Session.getActiveUser().getEmail().toLowerCase();
  var template = HtmlService.createTemplateFromFile('Index');
  
  // Kiểm tra quyền Admin
  template.isAdmin = (ADMIN_LIST.indexOf(userEmail) !== -1);
  template.userEmail = userEmail;

  return template.evaluate()
      .setTitle('BMT SOC IT - Hệ Thống Hoàn Thiện')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. Tra cứu tên nhân viên
function getEmployeeName(nvId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Admin_NhanVien');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === nvId.toString().trim()) return data[i][1];
  }
  return "Vô danh tiểu tốt";
}

// 2. Ghi nhật ký Nhận/Trả
function submitLog(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_NhatKy');
  var displayName = data.nvId + " (" + data.nvName + ")";
  sheet.appendRow([new Date(), displayName, data.tbId, data.tinhTrang, data.hinhAnh, data.hanhDong, Session.getActiveUser().getEmail()]);
  return "Đã ghi nhận thành công!";
}

// 3. Lấy lịch sử (Chỉ Admin mới có dữ liệu)
function getHistory() {
  var userEmail = Session.getActiveUser().getEmail().toLowerCase();
  if (ADMIN_LIST.indexOf(userEmail) === -1) return [];
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_NhatKy');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).reverse().slice(0, 50).map(row => ({
    time: Utilities.formatDate(new Date(row[0]), "GMT+7", "dd/MM HH:mm"),
    nv: row[1], tb: row[2], action: row[5]
  }));
}

// 4. Các hàm quản lý Admin (getData, update, delete)
function getData(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getDataRange().getValues();
}

function adminUpdateRow(sheetName, rowData) {
  if (ADMIN_LIST.indexOf(Session.getActiveUser().getEmail().toLowerCase()) === -1) return "Từ chối!";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var rowFound = -1;
  for (var i = 1; i < data.length; i++) { if (data[i][0].toString() === rowData[0].toString()) { rowFound = i + 1; break; } }
  if (rowFound !== -1) { sheet.getRange(rowFound, 1, 1, rowData.length).setValues([rowData]); return "Đã cập nhật!"; }
  else { sheet.appendRow(rowData); return "Đã thêm mới!"; }
}

function deleteDataRow(sheetName, idToDelete) {
  if (ADMIN_LIST.indexOf(Session.getActiveUser().getEmail().toLowerCase()) === -1) return "Từ chối!";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (data[i][0].toString() === idToDelete.toString()) { sheet.deleteRow(i + 1); return "Đã xóa!"; } }
}
