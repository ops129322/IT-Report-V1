/**
 * HỆ THỐNG QUẢN LÝ THIẾT BỊ IT - BMT SOC
 * Các tính năng: Báo cáo Nhận/Trả, Tra cứu tên NV, Lịch sử Admin, Quản lý danh mục.
 */

// 1. CẤU HÌNH DANH SÁCH ADMIN
// Chỉ những email này mới thấy được tab "Lịch sử" và "Admin"
const ADMIN_LIST = [
  'huythanh.pham@spxexpress.com', // Email của bạn
  'admin2@spxpress.com'         // Thêm các email admin khác tại đây
];

/**
 * Hàm khởi tạo giao diện App
 */
function doGet(e) {
  var userEmail = Session.getActiveUser().getEmail().toLowerCase();
  var template = HtmlService.createTemplateFromFile('Index');
  
  // Kiểm tra quyền Admin dựa trên danh sách ADMIN_LIST
  template.isAdmin = (ADMIN_LIST.indexOf(userEmail) !== -1);
  template.userEmail = userEmail;

  return template.evaluate()
      .setTitle('BMT SOC IT REPORT')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 2. CHỨC NĂNG NGƯỜI DÙNG (BÁO CÁO)
 */

// Tra cứu tên nhân viên từ mã ID
function getEmployeeName(nvId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Admin_NhanVien');
  var data = sheet.getDataRange().getValues();
  
  // Duyệt cột A tìm mã ID, nếu thấy trả về cột B (Tên)
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === nvId.toString().trim()) {
      return data[i][1];
    }
  }
  return "Vô danh tiểu tốt"; // Mặc định nếu chưa cập nhật thông tin
}

// Ghi dữ liệu Nhận/Trả vào Sheet Nhật Ký
function submitLog(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_NhatKy');
  // Lưu mã NV kèm theo tên để dễ truy soát
  var displayName = data.nvId + " (" + data.nvName + ")";
  
  sheet.appendRow([
    new Date(), 
    displayName, 
    data.tbId, 
    data.tinhTrang, 
    data.hinhAnh, 
    data.hanhDong, 
    Session.getActiveUser().getEmail()
  ]);
  return "Đã ghi nhận thành công cho: " + data.nvName;
}

/**
 * 3. CHỨC NĂNG ADMIN (BẢO MẬT)
 */

// Lấy 50 dòng lịch sử mới nhất (Chỉ Admin mới có quyền lấy dữ liệu)
function getHistory() {
  var userEmail = Session.getActiveUser().getEmail().toLowerCase();
  if (ADMIN_LIST.indexOf(userEmail) === -1) return []; // Chặn truy cập trái phép

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_NhatKy');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  // Đảo ngược danh sách để hiện cái mới nhất lên đầu
  return data.slice(1).reverse().slice(0, 50).map(row => ({
    time: Utilities.formatDate(new Date(row[0]), "GMT+7", "dd/MM HH:mm"),
    nv: row[1],
    tb: row[2],
    action: row[5]
  }));
}

// Lấy dữ liệu danh mục (Nhân viên hoặc Thiết bị)
function getData(sheetName) {
  var userEmail = Session.getActiveUser().getEmail().toLowerCase();
  if (ADMIN_LIST.indexOf(userEmail) === -1) return [];
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getDataRange().getValues();
}

// Cập nhật hoặc Thêm mới dòng dữ liệu trong danh mục
function adminUpdateRow(sheetName, rowData) {
  var userEmail = Session.getActiveUser().getEmail().toLowerCase();
  if (ADMIN_LIST.indexOf(userEmail) === -1) return "Từ chối truy cập!";
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var rowFound = -1;

  // Tìm kiếm theo ID (Cột A) để biết là sửa hay thêm mới
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === rowData[0].toString()) {
      rowFound = i + 1;
      break;
    }
  }

  if (rowFound !== -1) {
    sheet.getRange(rowFound, 1, 1, rowData.length).setValues([rowData]);
    return "Đã cập nhật thông tin cũ!";
  } else {
    sheet.appendRow(rowData);
    return "Đã thêm dữ liệu mới!";
  }
}

// Xóa một dòng dữ liệu theo ID
function deleteDataRow(sheetName, idToDelete) {
  var userEmail = Session.getActiveUser().getEmail().toLowerCase();
  if (ADMIN_LIST.indexOf(userEmail) === -1) return "Từ chối!";
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === idToDelete.toString()) {
      sheet.deleteRow(i + 1);
      return "Xóa thành công!";
    }
  }
  return "Không tìm thấy mã để xóa.";
}
