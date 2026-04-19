/**
 * HỆ THỐNG XẾP LỊCH TỰ ĐỘNG - NHÀ MÁY (VERSION 15.1 - HOTFIX THÁNG 4/2026)
 * - Fix lỗi "substring of null" do thiếu tên Nguyễn Đức Hiếu trong nhóm Admin
 * - Bổ sung cơ chế an toàn: Nhân viên không xác định sẽ tự động gán OFF để chống crash.
 */
const CONFIG = { 
  YEAR: 2026, 
  MONTH: 3, // Tháng 4 (Javascript Date bắt đầu từ 0)
  SHEET_NAME_KEYWORD: "Attendance", 
  HEADER_ROW: 2, 
  DATA_START_ROW: 3, 
  OT_COLUMN_INDEX: 39, // Cột AM
  
  GROUPS: { 
    VH_LEADERS: ["Nguyễn Văn Luân", "Nguyễn Trường An", "Nguyễn Hải Nguyên"], 
    VH_BACKUP: "Trần Mậu Thìn", 
    VH_FLEX: ["Nguyễn Hữu Quyết", "Nguyễn Văn Sáng", "Trần Tuấn Khoa", "Nguyễn Văn Nam", "Lê Hoàng Phúc", "Nguyễn Duy Triệu"], 
    VH_SPECIAL_DIEN: "Hoàng Văn Nam Điền", 
    VH_ADMIN: ["Tô Minh Hà", "Nguyễn Đức Hiếu"], // ĐÃ FIX: Thêm Nguyễn Đức Hiếu
    
    DG_LEADERS: ["Trần Nguyễn Khánh Linh", "Nguyễn Đình Thuận", "Jang Đan Tùng", "Nguyễn Văn Hải"], 
    DG_SUPPORT: ["Nguyễn Văn Linh", "Phan Thị Thanh Lý", "Bùi Văn Hoan", "Phạm Thái Toản"], 
    DG_REGULAR: ["Lê Thị Lụa", "Đỗ Thị Cài", "Nguyễn Thị Mai Thảo", "Hoàng Thị Việt", "Phạm Thị Minh Xuân", "Bùi Thị Minh Nam"], 
    DG_SPECIAL_PHUNG: "Võ Thị Thanh Phụng"
  }, 
  SYMBOLS: { 
    OFF: "O", ADM: "ADM", CN: "CN",
    S1: "S1", S2: "S2", S3: "S3", 
    S1_OT: "S1+4", S3_OT: "S3+4",
    PH: "PH", AL: "AL", UL: "UL"
  }, 
  COLORS: { 
    OFF: "#ff0000", OT: "#ffff00", ADM: "#ffffff", 
    PH: "#ff8c00", LEAVE: "#87cefa", 
    DEFAULT: "#ffffff", TEXT_OFF: "#ffffff", TEXT_STD: "#000000"
  }, 
  RULES: { 
    MAX_CONSECUTIVE_WORK: 6, 
    MAX_SAME_SHIFT: 12, 
    OT_LIMIT_MONTH: 40, 
    TARGET_VH: 3, 
    TARGET_DG: 5 
  } 
}; 

// ============================================================================
// CLASS EMPLOYEE
// ============================================================================
class Employee { 
  constructor(name, rowIndex) { 
    this.name = name; 
    this.rowIndex = rowIndex; 
    this.role = this.identifyRole(); 
    
    this.currentStreak = 0; 
    this.shiftStreak = 0; 
    this.lastShift = null; 
    
    this.currentPhase = 'S3'; 
    this.daysInPhase = 0; 
    this.lastCycleShift = 'S3'; 
    this.isPhaseTruncated = false; 
    
    this.otHours = 0; 
    this.schedule = {}; 
    this.isLocked = false; 
    this.todaysShift = null; 
    this.allowFlex = false; 
  } 
  
  identifyRole() { 
    const g = CONFIG.GROUPS; 
    if (g.VH_LEADERS.includes(this.name)) return 'VH_LEADER'; 
    if (this.name === g.VH_BACKUP) return 'VH_BACKUP'; 
    if (g.VH_FLEX.includes(this.name)) return 'VH_FLEX'; 
    if (this.name === g.VH_SPECIAL_DIEN) return 'VH_DIEN'; 
    if (g.VH_ADMIN.includes(this.name)) return 'VH_ADMIN'; // ĐÃ FIX
    
    if (g.DG_LEADERS.includes(this.name)) return 'DG_LEADER'; 
    if (g.DG_SUPPORT.includes(this.name)) return 'DG_SUPPORT'; 
    if (g.DG_REGULAR.includes(this.name)) return 'DG_REGULAR'; 
    if (this.name === g.DG_SPECIAL_PHUNG) return 'DG_PHUNG'; 
    return 'UNKNOWN'; 
  } 
  
  analyzeHistory(pastData) { 
    let lastDayVal = String(pastData[pastData.length - 1]).trim(); 
    let isLastOff = ['O', 'OFF', 'CN'].includes(lastDayVal) || !lastDayVal; 
    
    if (isLastOff) { 
      this.lastShift = 'O'; 
    } else { 
      this.lastShift = lastDayVal.substring(0, 2); 
    } 
    
    let workStreak = 0; 
    for (let i = pastData.length - 1; i >= 0; i--) { 
      let s = String(pastData[i]).trim(); 
      if (['O', 'OFF', 'CN'].includes(s) || !s) break; 
      workStreak++; 
    } 
    this.currentStreak = workStreak; 
    
    let phaseCount = 0; 
    let detectedPhase = null; 
    let realLastShift = null; 
    let reachedEnd = false; 
    
    for (let i = pastData.length - 1; i >= 0; i--) { 
      let s = String(pastData[i]).trim(); 
      if (!['O', 'OFF', 'CN'].includes(s) && s) { 
        let base = s.substring(0, 2); 
        if (!realLastShift) { 
          realLastShift = base; 
          detectedPhase = (base === 'S3') ? 'S3' : 'FLEX'; 
        } 
        break; 
      } 
    } 
    
    if (!detectedPhase) detectedPhase = 'S3'; 
    
    for (let i = pastData.length - 1; i >= 0; i--) { 
      let s = String(pastData[i]).trim(); 
      let isOff = ['O', 'OFF', 'CN'].includes(s) || !s; 
      
      if (!isOff) { 
        let base = s.substring(0, 2); 
        let currentShiftPhase = (base === 'S3') ? 'S3' : 'FLEX'; 
        
        if (currentShiftPhase === detectedPhase) { 
          phaseCount++; 
          if (i === 0) reachedEnd = true; 
        } else { 
          break; 
        } 
      } 
    } 
    
    this.currentPhase = detectedPhase; 
    this.daysInPhase = phaseCount; 
    this.lastCycleShift = realLastShift || 'S3'; 
    this.isPhaseTruncated = reachedEnd; 
  } 
} 

function rotateShift(shift) { 
  if (shift === 'S3') return 'S2'; 
  if (shift === 'S2') return 'S1'; 
  if (shift === 'S1') return 'S3'; 
  return 'S3'; 
} 

function isBioSafe(prev, next) { 
  if (prev === 'S3' && (next === 'S1' || next === 'S2')) return false; 
  return true; 
} 

// ============================================================================
// MAIN LOGIC
// ============================================================================
function runScheduleApril2026() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  let sheets = ss.getSheets(); 
  let sheet = null; 
  for(let s of sheets) { 
    if(s.getName().includes(CONFIG.SHEET_NAME_KEYWORD)) { sheet = s; break; } 
  } 
  if (!sheet) sheet = ss.getActiveSheet(); 
  
  const data = sheet.getDataRange().getValues(); 
  const lastRow = sheet.getLastRow(); 
  
  // 1. Tìm ngày 01/04
  let startCol = -1; 
  const dateRow = data[0]; 
  for (let i = 0; i < dateRow.length; i++) { 
    if (dateRow[i] instanceof Date && dateRow[i].getDate() === 1 && 
        dateRow[i].getMonth() === CONFIG.MONTH && dateRow[i].getFullYear() === CONFIG.YEAR) { 
      startCol = i; break; 
    } 
  } 
  if (startCol === -1) { SpreadsheetApp.getUi().alert("Không tìm thấy ngày 01/04/2026!"); return; } 
  
  // 2. Load Nhân viên
  let endEmpRow = -1; 
  for (let i = CONFIG.DATA_START_ROW; i < lastRow; i++) { 
    if (String(data[i][0]).startsWith("SL Ca")) { endEmpRow = i; break; } 
  } 
  if (endEmpRow === -1) endEmpRow = lastRow; 
  
  let employees = []; 
  for (let i = CONFIG.DATA_START_ROW - 1; i < endEmpRow; i++) { 
    let name = String(data[i][0]).trim(); 
    if (!name) continue; 
    let emp = new Employee(name, i); 
    let pastData = data[i].slice(Math.max(1, startCol - 30), startCol); 
    emp.analyzeHistory(pastData); 
    employees.push(emp); 
  } 
  
  // 3. Loop Ngày
  const daysInMonth = new Date(CONFIG.YEAR, CONFIG.MONTH + 1, 0).getDate(); 
  let dailyStats = []; 
  
  for (let d = 0; d < daysInMonth; d++) { 
    let currentDate = new Date(CONFIG.YEAR, CONFIG.MONTH, d + 1); 
    let dayOfWeek = currentDate.getDay(); 
    let dateKey = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "dd/MM"); 
    
    employees.forEach(e => { 
      e.isLocked = false; 
      e.todaysShift = null; 
      e.allowFlex = false; 
    }); 
    
    // --- BƯỚC 1: NGHỈ BẮT BUỘC & CỐ ĐỊNH ---
    employees.forEach(e => { 
      if (e.currentStreak >= CONFIG.RULES.MAX_CONSECUTIVE_WORK) { 
        e.todaysShift = CONFIG.SYMBOLS.OFF; e.isLocked = true; 
      } 
    }); 
    
    // ĐÃ FIX: Xếp lịch cho nhóm Admin
    let admins = employees.filter(e => e.role === 'VH_ADMIN'); 
    admins.forEach(admin => {
      admin.todaysShift = (dayOfWeek === 0) ? CONFIG.SYMBOLS.OFF : CONFIG.SYMBOLS.ADM; 
      admin.isLocked = true; 
    });
    
    let dien = employees.find(e => e.role === 'VH_DIEN'); 
    if (dien) { 
      if (dayOfWeek === 0) dien.todaysShift = CONFIG.SYMBOLS.OFF; 
      else if (dayOfWeek === 6) dien.todaysShift = CONFIG.SYMBOLS.ADM; 
      else dien.todaysShift = CONFIG.SYMBOLS.S2; 
      dien.isLocked = true; 
    } 
    
    // --- BƯỚC 2: XẾP NHÓM ĐÓNG GÓI (DG) ---
    let dgGroup = employees.filter(e => e.role.startsWith('DG')); 
    
    let dgFixed = dgGroup.filter(e => e.role === 'DG_LEADER' || e.role === 'DG_SUPPORT'); 
    dgFixed.forEach(emp => { 
      if (emp.isLocked) return; 
      let next = emp.lastShift; 
      if (emp.lastShift === 'O' || emp.shiftStreak >= CONFIG.RULES.MAX_SAME_SHIFT) { 
        next = rotateShift(emp.lastCycleShift); 
      } 
      if (!isBioSafe(emp.lastShift, next)) next = CONFIG.SYMBOLS.S3; 
      
      if (emp.role === 'DG_LEADER' && next === 'S1' && (dayOfWeek === 1 || dayOfWeek === 2)) { 
        next = CONFIG.SYMBOLS.S1_OT; 
        emp.otHours += 4; 
      } 
      emp.todaysShift = next; emp.isLocked = true; 
    }); 
    
    let phung = dgGroup.find(e => e.role === 'DG_PHUNG'); 
    if (phung && !phung.isLocked) { 
      let next = phung.lastShift; 
      if (phung.lastShift === 'O' || phung.shiftStreak >= CONFIG.RULES.MAX_SAME_SHIFT) { 
        if (phung.lastCycleShift === 'S3') next = 'S2'; 
        else if (phung.lastCycleShift === 'S2') { 
          if (phung.daysInPhase > 10) next = 'S3'; else next = 'S2'; 
        } 
      } 
      if (!isBioSafe(phung.lastShift, next)) next = CONFIG.SYMBOLS.S3; 
      if (next === 'S1') next = 'S2'; 
      phung.todaysShift = next; phung.isLocked = true; 
    } 
    
    let dgRegular = dgGroup.filter(e => e.role === 'DG_REGULAR'); 
    dgRegular.forEach(emp => { 
      if (emp.isLocked) return; 
      let next = emp.lastShift; 
      
      if (emp.lastShift === 'O') { 
        if (emp.currentPhase === 'S3') { 
          if (emp.daysInPhase >= 4 || emp.isPhaseTruncated) { 
            emp.currentPhase = 'FLEX'; 
            emp.daysInPhase = 0; 
            emp.isPhaseTruncated = false; 
            next = 'S2'; 
          } else { 
            next = 'S3'; 
          } 
        } else { 
          if (emp.daysInPhase >= 10 || emp.isPhaseTruncated) { 
            emp.currentPhase = 'S3'; 
            emp.daysInPhase = 0; 
            emp.isPhaseTruncated = false; 
            next = 'S3'; 
          } else { 
            next = 'S2'; 
          } 
        } 
      } else { 
        if (emp.currentPhase === 'S3') next = 'S3'; 
        else next = 'S2'; 
      } 
      
      if (!isBioSafe(emp.lastShift, next)) next = CONFIG.SYMBOLS.S3; 
      emp.todaysShift = next; 
      
      if (emp.currentPhase === 'S3' || next === 'S3') { 
        emp.isLocked = true; 
        emp.allowFlex = false; 
      } else { 
        emp.isLocked = false; 
        emp.allowFlex = true; 
      } 
    }); 
    
    balanceDG_StrictFlex(dgGroup, CONFIG.RULES.TARGET_DG); 
    
    // --- BƯỚC 3: CHECK RULE 6 (HẢI - LINH) ---
    let hai = dgGroup.find(e => e.name.includes("Nguyễn Văn Hải")); 
    let linh = dgGroup.find(e => e.name.includes("Trần Nguyễn Khánh Linh")); 
    let haiIsVH = false; 
    if (hai && linh && hai.todaysShift && linh.todaysShift) { 
      let hS = String(hai.todaysShift).substring(0, 2); 
      let lS = String(linh.todaysShift).substring(0, 2); 
      if (hS === lS && (hS === 'S1' || hS === 'S2')) haiIsVH = true; 
    } 
    
    // --- BƯỚC 4: XẾP NHÓM VẬN HÀNH (VH) ---
    let vhGroup = employees.filter(e => e.role.startsWith('VH')); 
    let vhLeaders = vhGroup.filter(e => e.role === 'VH_LEADER'); 
    let vhBackup = vhGroup.find(e => e.role === 'VH_BACKUP'); 
    let vhFlex = vhGroup.filter(e => e.role === 'VH_FLEX'); 
    
    vhLeaders.forEach(lead => { 
      if (lead.isLocked) return; 
      let next = lead.lastShift; 
      if (lead.lastShift === 'O' || lead.shiftStreak >= CONFIG.RULES.MAX_SAME_SHIFT) { 
        next = rotateShift(lead.lastCycleShift); 
      } 
      if (!isBioSafe(lead.lastShift, next)) next = CONFIG.SYMBOLS.S3; 
      lead.todaysShift = next; lead.isLocked = true; 
    }); 
    
    [vhBackup, ...vhFlex].forEach(emp => { 
      if (!emp || emp.isLocked) return; 
      let next = emp.lastShift; 
      if (emp.lastShift === 'O' || emp.shiftStreak >= CONFIG.RULES.MAX_SAME_SHIFT) { 
        next = rotateShift(emp.lastCycleShift); 
      } 
      if (!isBioSafe(emp.lastShift, next)) next = CONFIG.SYMBOLS.S3; 
      emp.todaysShift = next; 
    }); 
    
    let targetS1 = CONFIG.RULES.TARGET_VH; 
    let targetS3 = CONFIG.RULES.TARGET_VH; 
    if (haiIsVH && hai.todaysShift && hai.todaysShift.startsWith('S1')) targetS1 = 2; 
    let targetS2_Base = CONFIG.RULES.TARGET_VH; 
    if (haiIsVH && hai.todaysShift === 'S2') targetS2_Base = 2; 
    
    balanceVH_Aggressive(vhGroup, targetS1, targetS3); 
    
    let s2Count = vhGroup.filter(e => e.todaysShift === 'S2').length; 
    let missingS2 = targetS2_Base - s2Count; 
    let pairsCreated = 0; 
    
    if (missingS2 > 0) { 
      let s1C = vhGroup.filter(e => e.todaysShift === 'S1' && e.otHours < CONFIG.RULES.OT_LIMIT_MONTH); 
      let s3C = vhGroup.filter(e => e.todaysShift === 'S3' && e.otHours < CONFIG.RULES.OT_LIMIT_MONTH); 
      s1C.sort((a,b) => a.otHours - b.otHours); s3C.sort((a,b) => a.otHours - b.otHours); 
      
      while (missingS2 > 0 && s1C.length > 0 && s3C.length > 0) { 
        let p1 = s1C.shift(); let p3 = s3C.shift(); 
        p1.todaysShift = CONFIG.SYMBOLS.S1_OT; p1.otHours += 4; 
        p3.todaysShift = CONFIG.SYMBOLS.S3_OT; p3.otHours += 4; 
        missingS2--; pairsCreated++; 
      } 
    } 
    
    // --- BƯỚC 5: UPDATE STATE & STATS ---
    let vhS1 = vhGroup.filter(e => e.todaysShift && e.todaysShift.startsWith('S1')).length; 
    let vhS2 = vhGroup.filter(e => e.todaysShift === 'S2').length + pairsCreated; 
    let vhS3 = vhGroup.filter(e => e.todaysShift && e.todaysShift.startsWith('S3')).length; 
    
    let dgS1 = dgGroup.filter(e => e.todaysShift && e.todaysShift.startsWith('S1')).length; 
    let dgS2 = dgGroup.filter(e => e.todaysShift === 'S2').length; 
    let dgS3 = dgGroup.filter(e => e.todaysShift && e.todaysShift.startsWith('S3')).length; 
    
    if (haiIsVH && hai.todaysShift) { 
      if (hai.todaysShift.startsWith('S1')) { vhS1++; dgS1--; } 
      else if (hai.todaysShift === 'S2') { vhS2++; dgS2--; } 
    } 
    
    dailyStats.push({ vh: [vhS1, vhS2, vhS3], dg: [dgS1, dgS2, dgS3] }); 
    
    employees.forEach(e => { 
      // ĐÃ FIX: Safety check, nếu ai đó bị rớt khỏi mọi logic thì mặc định cho nghỉ để không bị lỗi null
      if (!e.todaysShift) e.todaysShift = CONFIG.SYMBOLS.OFF; 
      
      e.schedule[dateKey] = e.todaysShift; 
      if (e.todaysShift === CONFIG.SYMBOLS.OFF || e.todaysShift === CONFIG.SYMBOLS.ADM) { 
        e.currentStreak = 0; e.shiftStreak = 0; 
        if (['S1', 'S2', 'S3'].includes(e.lastShift)) e.lastCycleShift = e.lastShift; 
        e.lastShift = 'O'; 
      } else { 
        e.currentStreak++; 
        let raw = String(e.todaysShift).substring(0, 2); 
        if (raw === e.lastShift) e.shiftStreak++; else e.shiftStreak = 1; 
        
        let isFlex = ['S1', 'S2'].includes(raw); 
        if (e.currentPhase === 'FLEX' && isFlex) { 
          e.daysInPhase++; 
        } else if (e.currentPhase === 'S3' && raw === 'S3') { 
          e.daysInPhase++; 
        } else { 
          e.currentPhase = (raw === 'S3') ? 'S3' : 'FLEX'; 
          e.daysInPhase = 1; 
          e.isPhaseTruncated = false; 
        } 
        e.lastShift = raw; 
      } 
    }); 
  } 
  
  // --- GHI SHEET ---
  let output = employees.map(e => { 
    let row = []; 
    for (let d = 0; d < daysInMonth; d++) { 
      row.push(e.schedule[Utilities.formatDate(new Date(CONFIG.YEAR, CONFIG.MONTH, d + 1), Session.getScriptTimeZone(), "dd/MM")]); 
    } 
    return row; 
  }); 
  
  let range = sheet.getRange(CONFIG.DATA_START_ROW, startCol + 1, output.length, daysInMonth); 
  range.setValues(output); 
  applyFormatting(range, output); 
  
  let otData = employees.map(e => [e.otHours]); 
  sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.OT_COLUMN_INDEX, otData.length, 1).setValues(otData); 
  
  let counterStartRow = endEmpRow + 1; 
  for(let i=endEmpRow; i<lastRow; i++) { 
    if(String(data[i][0]).includes("SL Ca 1 (VH)")) { counterStartRow = i + 1; break; } 
  } 
  
  let statsOut = [[],[],[],[],[],[]]; 
  for(let d=0; d<daysInMonth; d++) { 
    statsOut[0].push(dailyStats[d].vh[0]); statsOut[1].push(dailyStats[d].vh[1]); statsOut[2].push(dailyStats[d].vh[2]); 
    statsOut[3].push(dailyStats[d].dg[0]); statsOut[4].push(dailyStats[d].dg[1]); statsOut[5].push(dailyStats[d].dg[2]); 
  } 
  sheet.getRange(counterStartRow, startCol + 1, 6, daysInMonth).setValues(statsOut).setFontWeight("bold"); 

  // --- GHI THỜI GIAN CẬP NHẬT TẠI Ô AN4 ---
  let timeString = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm dd/MM");
  sheet.getRange("AN4").setValue("Cập nhật: " + timeString).setFontColor("red").setFontWeight("bold");
  
  SpreadsheetApp.getUi().alert("✅ Đã xếp lịch thành công cho Tháng 4/2026 (Version 15.1)!"); 
} 

function balanceVH_Aggressive(group, targetS1, targetS3) { 
  let s1 = group.filter(e => e.todaysShift && e.todaysShift.startsWith('S1')); 
  let s3 = group.filter(e => e.todaysShift && e.todaysShift.startsWith('S3')); 
  
  if (s1.length < targetS1) { 
    let candidates = group.filter(e => !e.isLocked && e.todaysShift !== 'S1' && isBioSafe(e.lastShift, 'S1')); 
    candidates.sort((a,b) => (a.todaysShift === 'S3' ? -1 : 1)); 
    for (let emp of candidates) { 
      if (s1.length >= targetS1) break; 
      emp.todaysShift = 'S1'; s1.push(emp); 
    } 
  } 
  if (s3.length < targetS3) { 
    let candidates = group.filter(e => !e.isLocked && e.todaysShift !== 'S3' && isBioSafe(e.lastShift, 'S3')); 
    candidates.sort((a,b) => (a.todaysShift === 'S2' ? -1 : 1)); 
    for (let emp of candidates) { 
      if (s3.length >= targetS3) break; 
      emp.todaysShift = 'S3'; s3.push(emp); 
    } 
  } 
} 

function balanceDG_StrictFlex(group, target) { 
  let s1 = group.filter(e => e.todaysShift === 'S1'); 
  
  if (s1.length < target) { 
    let candidates = group.filter(e => e.allowFlex && e.todaysShift === 'S2' && isBioSafe(e.lastShift, 'S1')); 
    for (let emp of candidates) { 
      if (s1.length >= target) break; 
      emp.todaysShift = 'S1'; s1.push(emp); 
    } 
  } else if (s1.length > target) { 
    let candidates = group.filter(e => e.allowFlex && e.todaysShift === 'S1' && isBioSafe(e.lastShift, 'S2')); 
    for (let emp of candidates) { 
      if (s1.length <= target) break; 
      emp.todaysShift = 'S2'; 
      s1.pop(); 
    } 
  } 
} 

function applyFormatting(range, values) { 
  let bgs = [], fcs = [], fws = []; 
  for (let i = 0; i < values.length; i++) { 
    let rB = [], rF = [], rW = []; 
    for (let j = 0; j < values[i].length; j++) { 
      let val = values[i][j]; 
      let bg = CONFIG.COLORS.DEFAULT, fc = CONFIG.COLORS.TEXT_STD, fw = "normal"; 
      
      if (val === CONFIG.SYMBOLS.OFF || val === 'CN') { bg = CONFIG.COLORS.OFF; fc = CONFIG.COLORS.TEXT_OFF; } 
      else if (String(val).includes('+4')) { bg = CONFIG.COLORS.OT; } 
      else if (val === CONFIG.SYMBOLS.ADM) { bg = CONFIG.COLORS.ADM; fw = "bold"; } 
      else if (val === CONFIG.SYMBOLS.PH) { bg = CONFIG.COLORS.PH; fc = CONFIG.COLORS.TEXT_OFF; } 
      else if (val === CONFIG.SYMBOLS.AL || val === CONFIG.SYMBOLS.UL) { bg = CONFIG.COLORS.LEAVE; } 
      
      rB.push(bg); rF.push(fc); rW.push(fw); 
    } 
    bgs.push(rB); fcs.push(rF); fws.push(rW); 
  } 
  range.setBackgrounds(bgs).setFontColors(fcs).setFontWeights(fws); 
} 

function onOpen() { 
  SpreadsheetApp.getUi().createMenu('🏭 Admin Factory V15.1') 
    .addItem('🚀 Xếp Lịch Tháng 4/2026', 'runScheduleApril2026') 
    .addToUi(); 
}
