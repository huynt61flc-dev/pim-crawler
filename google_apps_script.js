/**
 * HƯỚNG DẪN CÀI ĐẶT TRÊN GOOGLE SHEETS:
 * 1. Mở file Google Sheet của bạn.
 * 2. Trên thanh menu, chọn: Tiện ích mở rộng (Extensions) -> Apps Script.
 * 3. Xóa hết code cũ (nếu có) và dán toàn bộ đoạn code dưới đây vào.
 * 4. Nhấn biểu tượng Lưu (Save) hoặc phím Ctrl + S.
 * 5. Nhấn nút "Triển khai" (Deploy) ở góc trên bên phải -> chọn "Tùy chọn triển khai mới" (New deployment).
 * 6. Mục "Chọn loại" (Select type): Nhấn biểu tượng bánh răng ⚙️ -> chọn "Ứng dụng web" (Web app).
 * 7. Cấu hình triển khai:
 *    - Mô tả (Description): Bot PIM
 *    - Người thực thi (Execute as): Tôi (Me)
 *    - Ai có quyền truy cập (Who has access): Bất kỳ ai (Anyone)  <--- QUAN TRỌNG!
 * 8. Nhấn "Triển khai" (Deploy), cấp quyền truy cập nếu Google yêu cầu.
 * 9. Copy "URL ứng dụng web" (Web App URL có dạng https://script.google.com/macros/s/.../exec) 
 *    và dán vào "google_apps_script_url" trong file config.json của bot.
 */

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var action = e.parameter.action || "read";

    // 1. Cập nhật tên sản phẩm vào Cột C
    if (action === "update") {
      var row = parseInt(e.parameter.row);
      var name = e.parameter.name || "";
      if (row && row >= 2) {
        // Cột 3 là Cột C
        sheet.getRange(row, 3).setValue(name);
        return ContentService.createTextOutput(JSON.stringify({
          status: "success",
          message: "Đã cập nhật hàng " + row,
          row: row,
          name: name
        })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({
          status: "error",
          message: "Hàng không hợp lệ: " + row
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // 2. Đọc toàn bộ danh sách mã từ Cột B và kết quả Cột C
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        total: 0,
        rows: []
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Đọc cột B (cột 2) và cột C (cột 3) từ dòng 2 đến hết
    var range = sheet.getRange(2, 2, lastRow - 1, 2);
    var values = range.getValues();
    var rows = [];

    for (var i = 0; i < values.length; i++) {
      var rowNumber = i + 2;
      var code = values[i][0];
      var name = values[i][1];

      rows.push({
        row: rowNumber,
        code: code !== null && code !== undefined ? String(code).trim() : "",
        name: name !== null && name !== undefined ? String(name).trim() : ""
      });
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      total: rows.length,
      rows: rows
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  return doGet(e);
}
