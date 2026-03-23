// --- CẤU HÌNH ADMIN ---
// Thay đổi email này thành email thật của bạn
const MY_ADMIN_EMAIL = 'huythanh.pham@spxpress.com'; 

// 1. Hàm khởi tạo Web App
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  
  // Lấy email người đăng nhập
  var userEmail = Session.getActiveUser().getEmail();
  
  // Phân quyền Admin
  template.isAdmin = (userEmail.toLowerCase() === MY_ADMIN_EMAIL.toLowerCase());
  template.userEmail = userEmail;

  return template.evaluate()
      .setTitle('BMT SOC IT REPORT - Đã Hoàn Thiện')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. Hàm xử lý khi User bấm NHẬN/TRẢ (Lưu nhật ký)
function submitLog(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Data_NhatKy');
  
  // Lưu dòng mới vào nhật ký
  sheet.appendRow([
    new Date(), 
    data.nvId, 
    data.tbId, 
    data.tinhTrang, 
    data.hinhAnh, // base64 string
    data.hanhDong,
    Session.getActiveUser().getEmail() // Email người thực hiện
  ]);
  
  return "Đã ghi nhận thành công!";
}

// 3. Hàm lấy dữ liệu (cho Admin)
function getData(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var values = sheet.getDataRange().getValues();
  // Bỏ dòng tiêu đề đầu tiên
  values.shift(); 
  return values;
}

// 4. Hàm cập nhật dữ liệu (cho Admin)
function adminUpdateRow(sheetName, rowData) {
  // Kiểm tra quyền Admin
  if (Session.getActiveUser().getEmail().toLowerCase() !== MY_ADMIN_EMAIL.toLowerCase()) {
    return "Lỗi: Bạn không có quyền Admin.";
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var idToFind = rowData[0];
  
  var rowFound = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == idToFind) { rowFound = i + 1; break; }
  }
  
  if (rowFound !== -1) {
    sheet.getRange(rowFound, 1, 1, rowData.length).setValues([rowData]);
    return "Đã cập nhật dữ liệu!";
  } else {
    sheet.appendRow(rowData);
    return "Đã thêm dữ liệu mới!";
  }
}
