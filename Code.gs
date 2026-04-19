function doGet() { 
  return HtmlService.createTemplateFromFile('Index') 
    .evaluate() 
    .setTitle('Lịch Làm Việc T4/2026') // Đã đổi thành T4
    // ĐÂY LÀ DÒNG LỆNH QUAN TRỌNG NHẤT ĐỂ ÉP GOOGLE HIỂN THỊ FULL MÀN HÌNH ĐIỆN THOẠI
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
} 

function getScheduleDataForWeb() { 
  const CONFIG = { 
    YEAR: 2026, 
    MONTH: 3, // Tháng 4 (JS Date tính từ 0 -> 3 là tháng 4)
    SHEET_NAME_KEYWORD: "Attendance", 
    DATA_START_ROW: 3, 
    // Ngày bắt đầu hiển thị: 25/03/2026
    START_DISPLAY_DATE: new Date(2026, 2, 25), // Tháng 3 (index 2)
    // Ngày kết thúc hiển thị: 24/04/2026
    END_DISPLAY_DATE: new Date(2026, 3, 24)    // Tháng 4 (index 3)
  }; 
  
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  let sheets = ss.getSheets(); 
  let sheet = null; 
  for(let s of sheets) { 
    if(s.getName().includes(CONFIG.SHEET_NAME_KEYWORD)) { sheet = s; break; } 
  } 
  if (!sheet) sheet = ss.getActiveSheet(); 
  
  const data = sheet.getDataRange().getValues(); 
  const lastRow = sheet.getLastRow(); 
  
  // 1. Tìm cột ngày bắt đầu (25/03)
  let startCol = -1; 
  const dateRow = data[0]; // Hàng 1 chứa ngày
  for (let i = 0; i < dateRow.length; i++) { 
    let d = dateRow[i]; 
    if (d instanceof Date && 
        d.getDate() === CONFIG.START_DISPLAY_DATE.getDate() && 
        d.getMonth() === CONFIG.START_DISPLAY_DATE.getMonth()) { 
      startCol = i; break; 
    } 
  } 
  
  if (startCol === -1) return { error: "Không tìm thấy ngày 25/03/2026 trên dòng 1" }; 
  
  // 2. Tính số cột cần lấy (từ 25/03 đến 24/04)
  let endCol = -1; 
  for (let i = startCol; i < dateRow.length; i++) { 
    let d = dateRow[i]; 
    if (d instanceof Date && 
        d.getDate() === CONFIG.END_DISPLAY_DATE.getDate() && 
        d.getMonth() === CONFIG.END_DISPLAY_DATE.getMonth()) { 
      endCol = i; break; 
    } 
  } 
  
  if (endCol === -1) endCol = startCol + 30; // Fallback: 25/03 -> 24/04 là 31 ngày (start + 30)
  
  const numCols = endCol - startCol + 1; 
  
  // 3. Lấy Header (Ngày & Thứ)
  let dates = []; 
  let days = []; 
  
  for (let i = 0; i < numCols; i++) { 
    let colIdx = startCol + i; 
    let d = data[0][colIdx]; 
    let dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM"); 
    dates.push(dateStr); 
    days.push(data[1][colIdx]); // Hàng 2 chứa Thứ
  } 
  
  // 4. Tìm hàng kết thúc nhân viên
  let endEmpRow = -1; 
  for (let i = CONFIG.DATA_START_ROW; i < lastRow; i++) { 
    if (String(data[i][0]).startsWith("SL Ca")) { endEmpRow = i; break; } 
  } 
  if (endEmpRow === -1) endEmpRow = lastRow; 
  
  // 5. Lấy dữ liệu Nhân viên & Chèn khoảng trắng
  let employees = []; 
  for (let i = CONFIG.DATA_START_ROW - 1; i < endEmpRow; i++) { 
    let name = String(data[i][0]).trim(); 
    if (!name) continue; 
    let shifts = data[i].slice(startCol, endCol + 1); 
    
    employees.push({ name: name, shifts: shifts, isSeparator: false }); 
    
    // Logic chèn dòng trống phân cách các nhóm
    // Sau "Nguyễn Văn Nam" -> Chèn trống
    if (name.includes("Nguyễn Văn Nam")) { 
      employees.push({ name: "", shifts: [], isSeparator: true }); 
    } 
    // Sau "Phạm Thái Toản" -> Chèn trống
    if (name.includes("Phạm Thái Toản")) { 
      employees.push({ name: "", shifts: [], isSeparator: true }); 
    } 
  } 
  
  // Lấy giá trị từ ô AN4 trong Sheet (Nơi lưu thời gian chạy tool xếp lịch)
  let updateTime = sheet.getRange("AN4").getDisplayValue().replace("Cập nhật: ", ""); 
  
  // Nếu ô AN4 trống, mới lấy giờ hiện tại làm dự phòng
  if (!updateTime) { 
    updateTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm dd/MM/yyyy"); 
  } 
  
  return { 
    dates: dates, 
    days: days, 
    employees: employees, 
    updateTime: updateTime
  }; 
}
