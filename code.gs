// 1. Hàm khởi tạo Web App
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  
  // Lấy email người đang truy cập để phân quyền
  var userEmail = Session.getActiveUser().getEmail();
  
  // ĐỊNH NGHĨA EMAIL ADMIN TẠI ĐÂY
  var adminEmails = ['huythanh.pham@spxepxress.com', '*#*#*#'];
  
  // Kiểm tra xem người dùng có phải Admin không
  template.isAdmin = adminEmails.includes(userEmail);
  template.userEmail = userEmail;

  return template.evaluate()
      .setTitle('BMT SOC IT REPORT - Trọn Gói')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. Hàm lấy dữ liệu (dùng chung cho User và Admin)
function getData(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  // Lấy toàn bộ dữ liệu, trừ dòng tiêu đề đầu tiên
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  var values = range.getValues();
  
  return values;
}

// 3. Hàm xử lý khi User bấm NHẬN/TRẢ (Lưu nhật ký)
function submitLog(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetData = ss.getSheetByName('Data_NhatKy');
  var sheetNV = ss.getSheetByName('Admin_NhanVien');
  var sheetTB = ss.getSheetByName('Admin_ThietBi');
  
  // Lấy tên NV và tên TB dựa trên ID quét được
  var tenNV = "Không tìm thấy";
  var dataNV = getData('Admin_NhanVien');
  for (var i = 0; i < dataNV.length; i++) {
    if (dataNV[i][0] == data.nvId) { tenNV = dataNV[i][1]; break; }
  }
  
  var tenTB = "Không tìm thấy";
  var dataTB = getData('Admin_ThietBi');
  for (var i = 0; i < dataTB.length; i++) {
    if (dataTB[i][0] == data.tbId) { tenTB = dataTB[i][1]; break; }
  }

  // Lưu dòng mới vào nhật ký
  sheetData.appendRow([
    new Date(), 
    data.nvId, 
    tenNV,
    data.tbId, 
    tenTB,
    data.tinhTrang, 
    data.hinhAnh, // Đây sẽ là base64 image string
    data.hanhDong,
    Session.getActiveUser().getEmail() // Email người thực hiện
  ]);
  
  return "Đã ghi nhận thành công!";
}

// 4. Hàm dành cho Admin: Cập nhật dữ liệu (Thêm hoặc Sửa)
// Nếu ID đã tồn tại -> Sửa. Nếu ID chưa có -> Thêm mới.
function adminUpdateRow(sheetName, rowData) {
  // Kiểm tra quyền Admin lần nữa cho chắc chắn
  var userEmail = Session.getActiveUser().getEmail();
  var adminEmails = ['huythanh.pham@spxexpress.com']; // Phải khớp với doGet
  if (!adminEmails.includes(userEmail)) return "Lỗi: Bạn không có quyền Admin.";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var idToFind = rowData[0]; // Giả định cột đầu tiên luôn là ID
  
  var rowFound = -1;
  // Tìm xem ID đã tồn tại chưa (bắt đầu từ dòng 2)
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == idToFind) {
      rowFound = i + 1; // Apps Script tính dòng từ 1
      break;
    }
  }
  
  if (rowFound !== -1) {
    // Sửa dòng đã có
    sheet.getRange(rowFound, 1, 1, rowData.length).setValues([rowData]);
    return "Đã cập nhật dữ liệu!";
  } else {
    // Thêm dòng mới
    sheet.appendRow(rowData);
    return "Đã thêm dữ liệu mới!";
  }
}
