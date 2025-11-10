/**
 * @OnlyCurrentDoc
 *
 * Tệp mã Google Apps Script này quản lý việc đăng ký mượn phòng từ Google Form.
 * PHIÊN BẢN CẬP NHẬT (v12 - Thêm "Lệnh Đồng ý")
 *
 * Chức năng:
 * 1. Tách biệt logic: 'processBookings' giờ chỉ
 * sắp xếp và "đánh dấu" (flag) các hàng chờ xử lý
 * (ví dụ: 'Chờ từ chối...').
 * 2. KHÔNG tự động gửi email từ 'processBookings'
 * hoặc 'autoReject...'.
 * 3. Thêm một nút menu mới 'Gửi mail tự động (Đồng ý)'
 * (hàm 'sendQueuedEmails').
 * 4. Nút này là "lệnh đồng ý", sẽ quét và gửi
 * tất cả các email đã được "đánh dấu" chờ.
 */

// --- CẤU HÌNH ---
const SHEET_BOOKINGS = 'Bookings'; // Sheet nhận dữ liệu từ Form
const SHEET_WAITING_LIST = 'Danh sách chờ';
const SHEET_ARCHIVE = 'Lưu trữ';

// CHỈ MỤC CỘT (Cấu trúc A-P)
const COL_TIMESTAMP = 1;     // A
const COL_EMAIL = 2;         // B
const COL_NAME = 3;          // C
const COL_BOOK_DATE = 7;     // G
const COL_START_TIME = 8;    // H
const COL_END_TIME = 9;      // I
const COL_ROOM = 10;         // J
const COL_PURPOSE = 11;      // K

const COL_STATUS = 17;       // Q
const COL_RESPONSE_TYPE = 18; // R

// --- HÀM KHỞI TẠO (THÊM NÚT MỚI) ---
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Quản lý Phòng')
    .addItem('1. Xử lý & Sắp xếp (Đánh dấu)', 'processBookings')
    .addItem('2. Gửi mail tự động (Đồng ý)', 'sendQueuedEmails') // NÚT MỚI (LỆNH ĐỒNG Ý)
    .addSeparator()
    .addItem('Dọn dẹp DS chờ (Thủ công)', 'autoRejectPassedWaitingList')
    .addToUi();
}

// --- HÀM CHÍNH (ĐÃ XÓA TỰ ĐỘNG GỬI MAIL) ---
function processBookings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bookingSheet = ss.getSheetByName(SHEET_BOOKINGS);
  const waitingSheet = ss.getSheetByName(SHEET_WAITING_LIST);
  const archiveSheet = ss.getSheetByName(SHEET_ARCHIVE);

  if (!bookingSheet || !waitingSheet || !archiveSheet) {
    SpreadsheetApp.getUi().alert('Lỗi: Không tìm thấy một trong các sheet cần thiết.');
    return;
  }

  let allPendingRows = [];
  
  // 1. GOM
  const bookingLastRow = bookingSheet.getLastRow();
  if (bookingLastRow > 1) {
    allPendingRows = allPendingRows.concat(
      bookingSheet.getRange(2, 1, bookingLastRow - 1, bookingSheet.getLastColumn()).getValues()
    );
  }
  const waitingLastRow = waitingSheet.getLastRow();
  if (waitingLastRow > 1) {
     allPendingRows = allPendingRows.concat(
      waitingSheet.getRange(2, 1, waitingLastRow - 1, waitingSheet.getLastColumn()).getValues()
    );
  }
  if (allPendingRows.length === 0) {
    SpreadsheetApp.getUi().alert('Không có dữ liệu mới hoặc dữ liệu chờ để xử lý.');
    return;
  }

  // 2. XÓA SẠCH
  if (bookingLastRow > 1) {
    bookingSheet.getRange(2, 1, bookingLastRow - 1, bookingSheet.getLastColumn()).clearContent().clearFormat();
  }
  if (waitingLastRow > 1) {
    waitingSheet.getRange(2, 1, waitingLastRow - 1, waitingSheet.getLastColumn()).clearContent().clearFormat();
  }
  
  // 3. PHÂN LOẠI
  const validRows = [], waitingRows = [], archiveRows = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayPlus15 = new Date(today.getTime());
  todayPlus15.setDate(today.getDate() + 15);

  allPendingRows.forEach(row => {
    const currentStatus = row[COL_STATUS - 1] ? row[COL_STATUS - 1].toString().toLowerCase() : '';

    // Bỏ qua tất cả các hàng đã được xử lý xong
    if (currentStatus.includes('đã gửi mail') || currentStatus.includes('đã từ chối')) {
      if (currentStatus.includes('đã từ chối')) {
        archiveRows.push(row);
      } else {
        validRows.push(row);
      }
      return; 
    }
    
    // Giữ nguyên trạng thái "Chờ" nếu nó đã được đánh dấu
    if (currentStatus.includes('chờ')) {
       if (currentStatus.includes('chờ từ chối')) {
         archiveRows.push(row);
       } else {
         waitingRows.push(row);
       }
       return;
    }
    
    // Xử lý các hàng mới (chưa có trạng thái)
    const bookDate = new Date(row[COL_BOOK_DATE - 1]);
    if (!bookDate || !(bookDate instanceof Date) || isNaN(bookDate.getTime())) return; 
    
    const startTime = new Date(row[COL_START_TIME - 1]);
    const dayOfWeek = bookDate.getDay();

    let status = '';
    let reason = ''; 
    let destination = 'valid'; 

    if (bookDate.getTime() < today.getTime()) {
      reason = 'Ngày đăng ký đã qua';
      destination = 'archive';
    } else if (dayOfWeek === 0) {
      reason = 'Thư viện không làm việc vào Chủ Nhật';
      destination = 'archive';
    } else if (dayOfWeek === 6 && (startTime.getHours() > 11 || (startTime.getHours() === 11 && startTime.getMinutes() >= 30))) {
      reason = 'Thư viện không làm việc vào chiều Thứ 7 (sau 11:30)';
      destination = 'archive';
    } else if (bookDate.getTime() > todayPlus15.getTime()) {
      destination = 'waiting';
    }
    
    // V12 THAY ĐỔI: Chỉ "đánh dấu" (flag), KHÔNG gửi email
    if (destination === 'archive' && reason) {
      status = `Chờ từ chối (${reason})`;
    } else if (destination === 'waiting') {
      status = 'Chờ gửi mail (Đang chờ)';
    }

    if (status) {
      row[COL_STATUS - 1] = status;
    }

    switch (destination) {
      case 'valid': validRows.push(row); break;
      case 'waiting': waitingRows.push(row); break;
      case 'archive': archiveRows.push(row); break;
    }
  });

  // 4. SẮP XẾP
  validRows.sort((a, b) => {
    const dateA = new Date(a[COL_BOOK_DATE - 1]);
    const dateB = new Date(b[COL_BOOK_DATE - 1]);
    if (dateA.getTime() !== dateB.getTime()) return dateA.getTime() - dateB.getTime();
    
    const timeA = new Date(a[COL_START_TIME - 1]).getTime();
    const timeB = new Date(b[COL_START_TIME - 1]).getTime();
    if (timeA !== timeB) return timeA - timeB;
    
    const timestampA = new Date(a[COL_TIMESTAMP - 1]);
    const timestampB = new Date(b[COL_TIMESTAMP - 1]);
    return timestampA.getTime() - timestampB.getTime(); 
  });

  // 5. GHI LẠI
  if (validRows.length > 0) {
    bookingSheet.getRange(2, 1, validRows.length, validRows[0].length).setValues(validRows);
  }
  if (waitingRows.length > 0) {
    waitingSheet.getRange(2, 1, waitingRows.length, waitingRows[0].length).setValues(waitingRows);
  }
  if (archiveRows.length > 0) {
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, archiveRows.length, archiveRows[0].length).setValues(archiveRows);
  }
  
  // 6. TÔ MÀU VÀ ĐỊNH DẠNG
  applyFormatting(bookingSheet);
  applyFormatting(waitingSheet);

  SpreadsheetApp.getUi().alert('Hoàn tất! Đã sắp xếp và "đánh dấu" các hàng. \n\nSử dụng nút "Gửi mail tự động (Đồng ý)" để gửi email hàng loạt.');
}

// --- (HÀM MỚI v12) - LỆNH ĐỒNG Ý ---
function sendQueuedEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const waitingSheet = ss.getSheetByName(SHEET_WAITING_LIST);
  const archiveSheet = ss.getSheetByName(SHEET_ARCHIVE);
  
  let sentCount = 0;
  
  // 1. Quét Danh sách chờ
  const waitingLastRow = waitingSheet.getLastRow();
  if (waitingLastRow > 1) {
    const range = waitingSheet.getRange(2, 1, waitingLastRow - 1, waitingSheet.getLastColumn());
    const data = range.getValues();
    const statuses = range.getBackgrounds(); // Dùng tạm để check thay đổi
    
    for (let i = 0; i < data.length; i++) {
      let row = data[i];
      let status = row[COL_STATUS - 1] ? row[COL_STATUS - 1].toString() : '';
      
      if (status === 'Chờ gửi mail (Đang chờ)') {
        try {
          let bookDate = new Date(row[COL_BOOK_DATE - 1]);
          let reviewDate = new Date(bookDate.getTime());
          reviewDate.setDate(reviewDate.getDate() - 15);
          
          sendWaitingListEmail(row, reviewDate); 
          
          waitingSheet.getRange(i + 2, COL_STATUS).setValue('Chờ xử lý (Đã gửi mail chờ)');
          sentCount++;
        } catch (e) {
          waitingSheet.getRange(i + 2, COL_STATUS).setValue(`Lỗi gửi mail: ${e.message}`);
          Logger.log(`Lỗi gửi mail CHỜ (Hàng ${i+2}): ${e.message}`);
        }
      }
    }
  }

  // 2. Quét Lưu trữ
  const archiveLastRow = archiveSheet.getLastRow();
  if (archiveLastRow > 1) {
    const range = archiveSheet.getRange(2, 1, archiveLastRow - 1, archiveSheet.getLastColumn());
    const data = range.getValues();
    
    for (let i = 0; i < data.length; i++) {
      let row = data[i];
      let status = row[COL_STATUS - 1] ? row[COL_STATUS - 1].toString() : '';
      
      if (status.startsWith('Chờ từ chối (')) {
        try {
          // Trích xuất lý do
          let reason = status.substring('Chờ từ chối ('.length, status.length - 1);
          
          sendRejectionEmail(row, reason); 
          
          let newStatus = status.replace('Chờ từ chối', 'Đã tự động từ chối');
          archiveSheet.getRange(i + 2, COL_STATUS).setValue(newStatus);
          sentCount++;
        } catch (e) {
          archiveSheet.getRange(i + 2, COL_STATUS).setValue(`Lỗi gửi mail: ${e.message}`);
          Logger.log(`Lỗi gửi mail TỪ CHỐI (Hàng ${i+2}): ${e.message}`);
        }
      }
    }
  }

  SpreadsheetApp.getUi().alert(`Hoàn tất! Đã gửi ${sentCount} email tự động.`);
}


// --- HÀM TÔ MÀU VÀ ĐỊNH DẠNG (GIỮ NGUYÊN) ---
function applyFormatting(sheet) {
  const numRows = sheet.getLastRow() - 1;
  if (numRows < 1) return; 
  
  const dataRange = sheet.getRange(2, 1, numRows, sheet.getLastColumn());
  const data = dataRange.getValues();
  const colors = [];

  const colorLightGray = '#f3f3f3'; 
  const colorWhite = '#ffffff';     
  
  let currentColor = colorLightGray;
  let lastDateStr = ''; 

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row[COL_BOOK_DATE - 1]) {
      colors.push(new Array(sheet.getLastColumn()).fill(colorWhite)); 
      continue; 
    }
    
    const bookDate = new Date(row[COL_BOOK_DATE - 1]);
    const currentDateStr = bookDate.toDateString(); 

    if (currentDateStr !== lastDateStr && lastDateStr !== '') {
      currentColor = (currentColor === colorLightGray) ? colorWhite : colorLightGray;
    }
    
    const rowColors = new Array(sheet.getLastColumn()).fill(currentColor);
    colors.push(rowColors);
    
    lastDateStr = currentDateStr; 
  }
  
  if (colors.length > 0) {
    dataRange.setBackgrounds(colors);
  }
  
  sheet.getRange(2, COL_START_TIME, numRows, 2).setNumberFormat("HH:mm");
}

// --- TRIGGER GỬI MAIL BẰNG TAY (KHÔNG THAY ĐỔI) ---
function handleEditTrigger(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();
  const value = e.value ? e.value.trim() : '';
  const lowerValue = value.toLowerCase();

  // Chỉ chạy khi chỉnh sửa cột R (Loại Phản Hồi) ở sheet Bookings
  if (sheet.getName() === SHEET_BOOKINGS && col === COL_RESPONSE_TYPE && row > 1) {
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusCell = sheet.getRange(row, COL_STATUS); 
    const responseCell = range; 

    if (value === '') {
      statusCell.setValue('');
      responseCell.clearFormat();
      return;
    }
    
    const currentStatus = statusCell.getValue().toString().toLowerCase();
    if (currentStatus.includes('đã gửi mail') && lowerValue === 'xác nhận') return;
    if (currentStatus.includes('đã từ chối') && lowerValue.startsWith('từ chối:')) return;

    // Đây là email GỬI BẰNG TAY, nó sẽ chạy ngay lập tức
    try {
      if (lowerValue === 'xác nhận') {
        sendConfirmationEmail(rowData); // Gửi mail xác nhận
        statusCell.setValue('Đã gửi mail');
        responseCell.setFontColor('#008000'); 
      } 
      else if (lowerValue.startsWith('từ chối:')) {
        const reason = value.substring(value.indexOf(':') + 1).trim(); 
        sendRejectionEmail(rowData, reason); // Gửi mail từ chối
        statusCell.setValue(`Đã từ chối`);
        responseCell.setFontColor('#ff0000'); 
      }
      statusCell.setFontWeight('bold');
    } catch (err) {
      Logger.log('Lỗi gửi email bằng tay: ' + err.message);
      statusCell.setValue('Lỗi: ' + err.message);
    }
  }
}

// --- CÁC HÀM GỬI EMAIL (KHÔNG THAY ĐỔI) ---

function sendWaitingListEmail(rowData, reviewDate) {
  const recipientEmail = rowData[COL_EMAIL - 1]; 
  if (!recipientEmail || recipientEmail.indexOf('@') === -1) {
    throw new Error('Email không hợp lệ (cột B).');
  }
  const name = rowData[COL_NAME - 1]; 
  const bookDate = new Date(rowData[COL_BOOK_DATE - 1]); 
  const formattedBookDate = Utilities.formatDate(bookDate, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const formattedReviewDate = Utilities.formatDate(reviewDate, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const room = rowData[COL_ROOM - 1]; 

  const subject = '[THÔNG BÁO] Đăng ký mượn phòng của bạn đang trong danh sách chờ';
  const body = `
Chào bạn ${name},
Chúng tôi đã nhận được đăng ký mượn phòng của bạn:
- Ngày mượn: ${formattedBookDate}
- Mã phòng: ${room}
Do bạn đăng ký sớm (trước hơn 15 ngày), đăng ký của bạn đã được đưa vào DANH SÁCH CHỜ.
Chúng tôi sẽ tự động xem xét và duyệt đăng ký này (nếu hợp lệ) vào khoảng ngày ${formattedReviewDate}.
Bạn không cần thực hiện thêm thao tác nào.
Trân trọng,
(Tên đơn vị/Thư viện của bạn)
  `;

  MailApp.sendEmail({ to: recipientEmail, subject: subject, body: body.trim() });
  Logger.log('Đã gửi email CHỜ thành công đến: ' + recipientEmail);
}


function sendConfirmationEmail(rowData) {
  const recipientEmail = rowData[COL_EMAIL - 1]; 
  if (!recipientEmail || recipientEmail.indexOf('@') === -1) {
    throw new Error('Email không hợp lệ (cột B).');
  }
  const name = rowData[COL_NAME - 1]; 
  const bookDate = new Date(rowData[COL_BOOK_DATE - 1]); 
  const startTime = new Date(rowData[COL_START_TIME - 1]); 
  const formattedDate = Utilities.formatDate(bookDate, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const formattedTime = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'HH:mm');
  const room = rowData[COL_ROOM - 1]; 
  const purpose = rowData[COL_PURPOSE - 1]; 

  const subject = '[XÁC NHẬN] Đăng ký mượn phòng thư viện thành công';
  const body = `
Chào bạn ${name},
Chúng tôi xác nhận đăng ký mượn phòng của bạn đã được duyệt:
- Ngày mượn: ${formattedDate}
- Thời gian: ${formattedTime}
- Mã phòng: ${room}
- Mục đích: ${purpose}
Vui lòng có mặt đúng giờ.
Trân trọng,
(Tên đơn vị/Thư viện của bạn)
  `;

  MailApp.sendEmail({ to: recipientEmail, subject: subject, body: body.trim() });
  Logger.log('Đã gửi email XÁC NHẬN thành công đến: ' + recipientEmail);
}

function sendRejectionEmail(rowData, reason) {
  const recipientEmail = rowData[COL_EMAIL - 1]; 
  if (!recipientEmail || recipientEmail.indexOf('@') === -1) {
    throw new Error('Email không hợp lệ (cột B).');
  }
  const name = rowData[COL_NAME - 1]; 
  const bookDate = new Date(rowData[COL_BOOK_DATE - 1]); 
  const startTime = new Date(rowData[COL_START_TIME - 1]); 
  const formattedDate = Utilities.formatDate(bookDate, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const formattedTime = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'HH:mm');
  const room = rowData[COL_ROOM - 1]; 

  const subject = '[THÔNG BÁO] Đăng ký mượn phòng thư viện KHÔNG thành công';
  const body = `
Chào bạn ${name},
Chúng tôi rất tiếc phải thông báo đăng ký mượn phòng của bạn đã không được duyệt:
- Ngày mượn: ${formattedDate}
- Thời gian: ${formattedTime}
- Mã phòng: ${room}
Lý do: ${reason}
Vui lòng kiểm tra lại lịch và thực hiện đăng ký mới nếu cần.
Trân trọng,
(Tên đơn vị/Thư viện của bạn)
  `;

  MailApp.sendEmail({ to: recipientEmail, subject: subject, body: body.trim() });
  Logger.log('Đã gửi email TỪ CHỐI thành công đến: ' + recipientEmail);
}

// --- HÀM TỰ ĐỘNG DỌN DẸP (ĐÃ SỬA, CHỈ ĐÁNH DẤU) ---
function autoRejectPassedWaitingList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const waitingSheet = ss.getSheetByName(SHEET_WAITING_LIST);
  const archiveSheet = ss.getSheetByName(SHEET_ARCHIVE);

  if (!waitingSheet || !archiveSheet) {
    Logger.log('Không tìm thấy sheet Danh sách chờ hoặc Lưu trữ.');
    return;
  }
  const lastRow = waitingSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('Không có gì trong danh sách chờ.');
    return; 
  }
  const dataRange = waitingSheet.getRange(2, 1, lastRow - 1, waitingSheet.getLastColumn());
  const data = dataRange.getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0); 
  const rowsToDelete = []; 
  const rowsToArchive = []; 

  data.forEach((row, index) => {
    const bookDate = new Date(row[COL_BOOK_DATE - 1]);
    
    if (bookDate.getTime() < today.getTime()) {
      Logger.log(`Đánh dấu hàng ${index + 2} (quá hạn)`);
      
      // V12 THAY ĐỔI: Chỉ "đánh dấu", KHÔNG gửi mail
      const reason = "Quá hạn trên danh sách chờ, không được duyệt";
      row[COL_STATUS - 1] = `Chờ từ chối (${reason})`; // Đánh dấu để hàm sendQueuedEmails xử lý
      
      rowsToArchive.push(row);
      rowsToDelete.push(index + 2); 
    }
  });

  if (rowsToArchive.length > 0) {
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);
    Logger.log(`Đã chuyển ${rowsToArchive.length} hàng (chờ từ chối) sang Lưu trữ.`);
  }

  // Xóa hàng từ dưới lên
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    waitingSheet.deleteRow(rowsToDelete[i]);
  }
  
  Logger.log('Hoàn tất dọn dẹp danh sách chờ.');
}
