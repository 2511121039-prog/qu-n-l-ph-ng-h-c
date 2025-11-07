# qu-n-l-ph-ng-h-c
phân loại các thời gian sử dụng phòng trong 15 ngày 
/* =====================================================
 HÀM 1: KIỂM TRA TRÙNG LẶP & SẮP XẾP
 Chạy khi người dùng GỬI FORM
=====================================================
*/
function checkBookingConflicts(e) {
  
  // --- PHẦN 1: Lấy dữ liệu từ Form vừa được gửi ---
  var formResponse = e.namedValues;
  
  // <<< KIỂM TRA TIÊU ĐỀ CÂU HỎI CỦA BẠN >>>
  var userEmail = formResponse['Email'][0]; 
  var newDateStr = formResponse['Ngày đăng ký mượn phòng:'][0];
  var newStartTimeStr = formResponse['Thời gian bắt đầu:'][0];
  var newEndTimeStr = formResponse['Thời gian kết thúc:'][0];
  var newRoom = formResponse['Mã phòng đăng ký mượn'][0]; 

  // --- PHẦN 2: Chuyển đổi Ngày/Giờ ---
  // Giả sử định dạng ngày của bạn là mm/dd/yyyy
  var newDateParts = newDateStr.split('/');
  var newDate = new Date(newDateParts[2], newDateParts[0] - 1, newDateParts[1]); 
  newDate.setHours(0, 0, 0, 0); // Chuẩn hóa ngày

  var newStartParts = newStartTimeStr.split(':');
  var newStartDateTime = new Date(newDate.getFullYear(), newDate.getMonth(), newDate.getDate(), newStartParts[0], newStartParts[1]);
  
  var newEndParts = newEndTimeStr.split(':');
  var newEndDateTime = new Date(newDate.getFullYear(), newDate.getMonth(), newDate.getDate(), newEndParts[0], newEndParts[1]);

  // --- PHẦN 3: Mở Sheet và quét dữ liệu cũ ---
  // <<< THAY ĐỔI TÊN SHEET NẾU CẦN >>>
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1"); 
  var data = sheet.getDataRange().getValues(); 

  var conflictFound = false;
  var conflictDetails = "";

  // Bắt đầu quét từ hàng thứ 2 (index 1)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // --- PHẦN 4: Logic kiểm tra xung đột (ĐÃ SỬA INDEX) ---
    // Cột I (index 8) là 'Mã phòng'
    var oldRoom = row[8]; 
    
    if (oldRoom == newRoom) {
      // Cột E (index 4) là 'Ngày đăng ký'
      var oldDate = new Date(row[4]); 
      oldDate.setHours(0, 0, 0, 0); // Chuẩn hóa ngày cũ
      
      if (oldDate.getTime() == newDate.getTime()) {
        // Cột G (index 6) là 'Giờ Bắt đầu'
        var oldStartTime = new Date(row[6]); 
        // Cột H (index 7) là 'Giờ Kết thúc'
        var oldEndTime = new Date(row[7]); 

        var oldStartDateTime = new Date(oldDate.getFullYear(), oldDate.getMonth(), oldDate.getDate(), oldStartTime.getHours(), oldStartTime.getMinutes());
        var oldEndDateTime = new Date(oldDate.getFullYear(), oldDate.getMonth(), oldDate.getDate(), oldEndTime.getHours(), oldEndTime.getMinutes());

        if (newStartDateTime < oldEndDateTime && newEndDateTime > oldStartDateTime) {
          conflictFound = true;
          conflictDetails = "Phòng '" + newRoom + "' đã được đặt từ " + 
                            oldStartTime.toLocaleTimeString('vi-VN', {hour: '2-digit', minute:'2-digit'}) + " đến " + 
                            oldEndTime.toLocaleTimeString('vi-VN', {hour: '2-digit', minute:'2-digit'}) + 
                            " vào ngày " + newDateStr + ".";
          break; 
        }
      }
    }
  } // Kết thúc vòng lặp

  // --- PHẦN 5: Gửi Email VÀ BÔI MÀU (NẾU CẦN) ---
  var subject = "";
  var body = "";
  
  // Lấy hàng (range) của câu trả lời này để bôi màu
  var newRowRange = sheet.getRange(e.range.getRow(), 1, 1, sheet.getLastColumn());

  if (conflictFound) {
    subject = "[TỪ CHỐI] Đăng ký mượn phòng " + newRoom;
    body = "Chào bạn,\n\n" +
           "Yêu cầu đăng ký mượn phòng của bạn đã bị TỪ CHỐI do trùng lịch.\n\n" +
           "Chi tiết trùng lặp: " + conflictDetails + "\n\n" +
           "Vui lòng kiểm tra lại và gửi một yêu cầu khác với thời gian hoặc phòng khác.\n\n" +
           "Cảm ơn.";
    
    // Bôi màu ĐỎ/HỒNG cho hàng bị từ chối
    newRowRange.setBackground("#f4cccc"); // Màu hồng nhạt
    
  } else {
    subject = "[XÁC NHẬN] Đăng ký mượn phòng " + newRoom;
    body = "Chào bạn,\n\n" +
           "Yêu cầu đăng ký mượn phòng của bạn đã được XÁC NHẬN.\n\n" +
           "Chi tiết:\n" +
           "- Phòng: " + newRoom + "\n" +
           "- Ngày: " + newDateStr + "\n" +
           "- Thời gian: " + newStartTimeStr + " - " + newEndTimeStr + "\n\n" +
           "Cảm ơn.";
           
    // Kiểm tra 15 ngày
    var today = new Date();
    today.setHours(0, 0, 0, 0); 
    var fifteenDaysInMillis = 15 * 24 * 60 * 60 * 1000;
    var fifteenDaysFromNow = new Date(today.getTime() + fifteenDaysInMillis);
    
    if (newDate >= today && newDate <= fifteenDaysFromNow) {
      newRowRange.setBackground("#d9ead3"); // Màu xanh lá nhạt
    } else {
      newRowRange.setBackground("#ffffff"); // Màu trắng
    }
  }
  
  MailApp.sendEmail(userEmail, subject, body);
  
  // --- PHẦN 6: TỰ ĐỘNG SẮP XẾP SHEET (ĐÃ SỬA SỐ CỘT) ---
  var rangeToSort = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  
  // Cột 5 (Cột E - 'Ngày đăng ký')
  // Cột 7 (Cột G - 'Thời gian bắt đầu')
  rangeToSort.sort([
    { column: 5, ascending: true }, 
    { column: 7, ascending: true } 
  ]);
  
} // <-- KẾT THÚC HÀM 1

/* =====================================================
 HÀM 2: TỰ ĐỘNG CẬP NHẬT MÀU SẮC HÀNG NGÀY
 Chạy mỗi ngày một lần
=====================================================
*/
function updateBookingHighlights() {
  
  // <<< THAY ĐỔI TÊN SHEET NẾU CẦN >>>
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  var COLOR_UPCOMING = "#d9ead3"; // Xanh lá nhạt (sắp tới trong 15 ngày)
  var COLOR_PAST = "#f3f3f3";     // Xám nhạt (đã qua)
  var COLOR_FUTURE = "#ffffff";    // Trắng (còn xa)

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  var fifteenDaysInMillis = 15 * 24 * 60 * 60 * 1000;
  var fifteenDaysFromNow = new Date(today.getTime() + fifteenDaysInMillis);

  var backgroundColors = [];

  // Lặp qua tất cả các hàng, BỎ QUA hàng 1 (tiêu đề)
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    
    // Cột E (index 4) là 'Ngày đăng ký'
    var bookingDateCell = row[4]; 
    
    if (!bookingDateCell || !(bookingDateCell instanceof Date)) {
      backgroundColors.push(COLOR_FUTURE); // Màu trắng
      continue;
    }
    
    var bookingDate = new Date(bookingDateCell);
    bookingDate.setHours(0, 0, 0, 0); 
    
    if (bookingDate < today) {
      backgroundColors.push(COLOR_PAST);
    } else if (bookingDate >= today && bookingDate <= fifteenDaysFromNow) {
      backgroundColors.push(COLOR_UPCOMING);
    } else {
      backgroundColors.push(COLOR_FUTURE);
    }
  }
  
  // Cập nhật màu sắc cho tất cả các hàng
  var startRow = 2; // Bắt đầu từ hàng 2
  for (var j = 0; j < backgroundColors.length; j++) {
    var color = backgroundColors[j];
    sheet.getRange(startRow + j, 1, 1, sheet.getLastColumn()).setBackground(color);
  }
  
} // <-- KẾT THÚC HÀM 2
