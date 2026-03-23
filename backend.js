// --- CẤU HÌNH ADMIN ---
const ADMIN_LIST = ['huythanh.pham@spxexpress.com']; //

// Hàm xử lý yêu cầu GET (Tra cứu & Lấy dữ liệu)
function doGet(e) {
  var action = e.parameter.action;
  var userEmail = Session.getActiveUser().getEmail().toLowerCase();
  
  if (action === 'checkAdmin') return content(ADMIN_LIST.indexOf(userEmail) !== -1);
  if (action === 'getName') return content(getEmployeeName(e.parameter.id));
  if (action === 'getHistory') return content(getHistory(userEmail));
  if (action === 'getAdminData') return content(getData(e.parameter.sheet, userEmail));
  
  // Trả về giao diện chính nếu truy cập trực tiếp link Script
  var template = HtmlService.createTemplateFromFile('Index');
  template.isAdmin = (ADMIN_LIST.indexOf(userEmail) !== -1);
  template.userEmail = userEmail;
  return template.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Hàm xử lý yêu cầu POST (Ghi/Sửa/Xóa dữ liệu)
function doPost(e) {
  var postData = JSON.parse(e.postData.contents);
  var userEmail = Session.getActiveUser().getEmail().toLowerCase();
  
  if (postData.action === 'submitLog') return content(submitLog(postData.data, userEmail));
  if (postData.action === 'adminUpdate') return content(adminUpdateRow(postData.sheet, postData.row, userEmail));
  if (postData.action === 'adminDelete') return content(deleteDataRow(postData.sheet, postData.id, userEmail));
}

// --- CÁC HÀM BỔ TRỢ (GIỮ NGUYÊN LOGIC CŨ) ---
function content(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getEmployeeName(id) {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Admin_NhanVien').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) return data[i][1];
  }
  return "Vô danh tiểu tốt"; //
}

function submitLog(data, email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_NhatKy');
  sheet.appendRow([new Date(), data.nvId + " (" + data.nvName + ")", data.tbId, data.tinhTrang, data.hinhAnh, data.hanhDong, email]);
  return "Ghi nhận thành công!";
}

function getHistory(email) {
  if (ADMIN_LIST.indexOf(email) === -1) return [];
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_NhatKy').getDataRange().getValues();
  return data.slice(1).reverse().slice(0, 50).map(r => ({ time: Utilities.formatDate(new Date(r[0]), "GMT+7", "HH:mm"), nv: r[1], tb: r[2], action: r[5] }));
}

function getData(name, email) {
  if (ADMIN_LIST.indexOf(email) === -1) return [];
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).getDataRange().getValues();
}

function adminUpdateRow(name, row, email) {
  if (ADMIN_LIST.indexOf(email) === -1) return "Từ chối!";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === row[0].toString()) {
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return "Đã cập nhật!";
    }
  }
  sheet.appendRow(row);
  return "Đã thêm mới!";
}

function deleteDataRow(name, id, email) {
  if (ADMIN_LIST.indexOf(email) === -1) return "Từ chối!";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) { sheet.deleteRow(i + 1); return "Đã xóa!"; }
  }
}
