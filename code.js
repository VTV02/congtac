// ============================================
// HỆ THỐNG QUẢN LÝ CÔNG TÁC - THACO AGRI
// Version: 2.0.0 - FULL FEATURES
// Author: Development Team
// ============================================

// ============================================
// CẤU HÌNH HỆ THỐNG
// ============================================
var CONFIG = {
  sheetNames: {
    dangKy: 'Đăng ký công tác',
    cauHinhEmail: 'Cấu hình Email',
    danhMuc: 'Danh mục',
    phanQuyen: 'Phân quyền',
    lichSuDuyet: 'Lịch sử duyệt',
    thongKe: 'Thống kê'
  },
  columns: {
    timestamp: 0,       // A
    maDangKy: 1,        // B
    hoTen: 2,           // C
    chucVu: 3,          // D
    phongBan: 4,        // E
    thongTin: 5,        // F
    ngayDen: 6,         // G
    ngayDi: 7,          // H
    phuongTien: 8,      // I
    nhaAn: 9,           // J
    diaDiem: 10,        // K
    email: 11,          // L
    trangThai: 12,      // M
    emailDaGui: 13,     // N
    thoiGianGui: 14,    // O
    nguoiDuyet: 15,     // P
    ngayDuyet: 16,      // Q
    lyDoTuChoi: 17,     // R
    lichSu: 18,         // S
    fileDinhKem: 19,    // T
    daXoa: 20           // U
  },
  emailSubject: '[THACO AGRI] Thông báo công tác',
  brandColor: '#00A86B',
  webAppUrl: '' // Sẽ cập nhật sau khi deploy
};

// ============================================
// HÀM TẠO MENU CUSTOM
// ============================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Quản lý Công tác')
    .addItem('🚀 Mở Dashboard', 'openDashboard')
    .addSeparator()
    .addItem('🔧 Cài đặt hệ thống', 'setupSystem')
    .addItem('📊 Xem thống kê', 'viewStatistics')
    .addSeparator()
    .addItem('📥 Export Excel', 'exportToExcel')
    .addItem('🗑️ Dọn dẹp dữ liệu cũ', 'cleanOldData')
    .addToUi();
}

// ============================================
// MỞ DASHBOARD WEB APP
// ============================================
function openDashboard() {
  var html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Quản lý Công tác - THACO AGRI')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard Quản lý Công tác');
}

// ============================================
// SERVE DASHBOARD KHI TRUY CẬP URL
// ============================================
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Dashboard');
  return template.evaluate()
    .setTitle('Hệ thống Quản lý Công tác - THACO AGRI')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================
// INCLUDE CSS/JS FILES
// ============================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// API: LẤY DỮ LIỆU DANH MỤC
// ============================================
function getDanhMuc() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetNames.danhMuc);
    var data = sheet.getDataRange().getValues();
    
    return {
      success: true,
      data: {
        chucVu: getColumnData(data, 0),
        phongBan: getColumnData(data, 2),
        phuongTien: getColumnData(data, 4),
        nhaAn: getColumnData(data, 6),
        diaDiem: getColumnData(data, 8),
        trangThai: getColumnData(data, 10)
      }
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

function getColumnData(data, colIndex) {
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][colIndex]) {
      result.push(data[i][colIndex]);
    }
  }
  return result;
}

// ============================================
// API: LẤY DỮ LIỆU ĐĂNG KÝ (CÓ PHÂN TRANG & LỌC)
// ============================================
function getDangKyList(filters) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetNames.dangKy);
    var data = sheet.getDataRange().getValues();
    
    var result = [];
    var userEmail = Session.getActiveUser().getEmail();
    var userRole = getUserRole(userEmail);
    
    // Bỏ qua header
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Bỏ qua dòng đã xóa
      if (row[CONFIG.columns.daXoa] === true) continue;
      
      // Kiểm tra quyền xem
      if (!canViewRecord(userRole, userEmail, row)) continue;
      
      // Áp dụng filters
      if (filters) {
        if (filters.search && !matchSearch(row, filters.search)) continue;
        if (filters.trangThai && row[CONFIG.columns.trangThai] !== filters.trangThai) continue;
        if (filters.diaDiem && row[CONFIG.columns.diaDiem] !== filters.diaDiem) continue;
        if (filters.fromDate && new Date(row[CONFIG.columns.ngayDen]) < new Date(filters.fromDate)) continue;
        if (filters.toDate && new Date(row[CONFIG.columns.ngayDen]) > new Date(filters.toDate)) continue;
      }
      
      result.push({
        rowIndex: i + 1,
        timestamp: formatDateTime(row[CONFIG.columns.timestamp]),
        maDangKy: row[CONFIG.columns.maDangKy],
        hoTen: row[CONFIG.columns.hoTen],
        chucVu: row[CONFIG.columns.chucVu],
        phongBan: row[CONFIG.columns.phongBan],
        thongTin: row[CONFIG.columns.thongTin],
        ngayDen: formatDate(row[CONFIG.columns.ngayDen]),
        ngayDi: formatDate(row[CONFIG.columns.ngayDi]),
        phuongTien: row[CONFIG.columns.phuongTien],
        nhaAn: row[CONFIG.columns.nhaAn],
        diaDiem: row[CONFIG.columns.diaDiem],
        email: row[CONFIG.columns.email],
        trangThai: row[CONFIG.columns.trangThai],
        nguoiDuyet: row[CONFIG.columns.nguoiDuyet],
        ngayDuyet: row[CONFIG.columns.ngayDuyet] ? formatDateTime(row[CONFIG.columns.ngayDuyet]) : '',
        lyDoTuChoi: row[CONFIG.columns.lyDoTuChoi]
      });
    }
    
    return {
      success: true,
      data: result,
      total: result.length
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// API: THÊM ĐĂNG KÝ MỚI
// ============================================
function submitDangKy(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetNames.dangKy);
    var userEmail = Session.getActiveUser().getEmail();
    
    // Validate dữ liệu
    var validation = validateFormData(formData);
    if (!validation.valid) {
      return {
        success: false,
        error: validation.error
      };
    }
    
    // Tạo mã đăng ký tự động
    var maDangKy = generateMaDangKy();
    
    // Tạo timestamp
    var now = new Date();
    
    // Tạo log lịch sử
    var lichSu = JSON.stringify([{
      time: formatDateTime(now),
      action: 'Tạo mới',
      user: userEmail
    }]);
    
    // Thêm dòng mới
    var newRow = [
      now,                                    // A: Timestamp
      maDangKy,                               // B: Mã đăng ký
      formData.hoTen,                         // C: Họ tên
      formData.chucVu,                        // D: Chức vụ
      formData.phongBan,                      // E: Loại phòng
      formData.thongTin || '',                // F: Thông tin
      new Date(formData.ngayDen),             // G: Ngày đến
      new Date(formData.ngayDi),              // H: Ngày đi
      formData.phuongTien,                    // I: Phương tiện
      formData.nhaAn ? formData.nhaAn.join(', ') : '', // J: Nhà ăn
      formData.diaDiem,                       // K: Địa điểm
      formData.email,                         // L: Email
      'Chờ duyệt',                            // M: Trạng thái
      '',                                     // N: Email đã gửi
      '',                                     // O: Thời gian gửi
      '',                                     // P: Người duyệt
      '',                                     // Q: Ngày duyệt
      '',                                     // R: Lý do từ chối
      lichSu,                                 // S: Lịch sử
      '',                                     // T: File đính kèm
      false                                   // U: Đã xóa
    ];
    
    sheet.appendRow(newRow);
    var newRowIndex = sheet.getLastRow();
    
    // Format dòng mới
    formatNewRow(sheet, newRowIndex);
    
    // Gửi email thông báo
    var emailResult = sendEmailThongBao(maDangKy, formData, 'Chờ duyệt');
    
    // Cập nhật thông tin email đã gửi
    if (emailResult.success) {
      sheet.getRange(newRowIndex, CONFIG.columns.emailDaGui + 1).setValue(emailResult.sentTo);
      sheet.getRange(newRowIndex, CONFIG.columns.thoiGianGui + 1).setValue(new Date());
    }
    
    return {
      success: true,
      message: 'Đăng ký thành công! Mã đăng ký: ' + maDangKy,
      maDangKy: maDangKy,
      rowIndex: newRowIndex
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// API: CẬP NHẬT ĐĂNG KÝ
// ============================================
function updateDangKy(rowIndex, formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetNames.dangKy);
    var userEmail = Session.getActiveUser().getEmail();
    var userRole = getUserRole(userEmail);
    
    // Kiểm tra quyền
    var currentData = sheet.getRange(rowIndex, 1, 1, 21).getValues()[0];
    if (!canEditRecord(userRole, userEmail, currentData)) {
      return {
        success: false,
        error: 'Bạn không có quyền chỉnh sửa đăng ký này!'
      };
    }
    
    // Validate
    var validation = validateFormData(formData);
    if (!validation.valid) {
      return {
        success: false,
        error: validation.error
      };
    }
    
    // Lấy lịch sử cũ và thêm log mới
    var oldLichSu = currentData[CONFIG.columns.lichSu];
    var lichSuArray = oldLichSu ? JSON.parse(oldLichSu) : [];
    lichSuArray.push({
      time: formatDateTime(new Date()),
      action: 'Chỉnh sửa',
      user: userEmail
    });
    
    // Cập nhật dữ liệu
    sheet.getRange(rowIndex, CONFIG.columns.hoTen + 1).setValue(formData.hoTen);
    sheet.getRange(rowIndex, CONFIG.columns.chucVu + 1).setValue(formData.chucVu);
    sheet.getRange(rowIndex, CONFIG.columns.phongBan + 1).setValue(formData.phongBan);
    sheet.getRange(rowIndex, CONFIG.columns.thongTin + 1).setValue(formData.thongTin || '');
    sheet.getRange(rowIndex, CONFIG.columns.ngayDen + 1).setValue(new Date(formData.ngayDen));
    sheet.getRange(rowIndex, CONFIG.columns.ngayDi + 1).setValue(new Date(formData.ngayDi));
    sheet.getRange(rowIndex, CONFIG.columns.phuongTien + 1).setValue(formData.phuongTien);
    sheet.getRange(rowIndex, CONFIG.columns.nhaAn + 1).setValue(formData.nhaAn ? formData.nhaAn.join(', ') : '');
    sheet.getRange(rowIndex, CONFIG.columns.diaDiem + 1).setValue(formData.diaDiem);
    sheet.getRange(rowIndex, CONFIG.columns.email + 1).setValue(formData.email);
    sheet.getRange(rowIndex, CONFIG.columns.lichSu + 1).setValue(JSON.stringify(lichSuArray));
    
    return {
      success: true,
      message: 'Cập nhật thành công!'
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// API: XÓA ĐĂNG KÝ (XÓA MỀM)
// ============================================
function deleteDangKy(rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetNames.dangKy);
    var userEmail = Session.getActiveUser().getEmail();
    var userRole = getUserRole(userEmail);
    
    // Kiểm tra quyền
    var currentData = sheet.getRange(rowIndex, 1, 1, 21).getValues()[0];
    if (!canDeleteRecord(userRole, userEmail, currentData)) {
      return {
        success: false,
        error: 'Bạn không có quyền xóa đăng ký này!'
      };
    }
    
    // Xóa mềm - đánh dấu đã xóa
    sheet.getRange(rowIndex, CONFIG.columns.daXoa + 1).setValue(true);
    
    // Thêm log
    var oldLichSu = currentData[CONFIG.columns.lichSu];
    var lichSuArray = oldLichSu ? JSON.parse(oldLichSu) : [];
    lichSuArray.push({
      time: formatDateTime(new Date()),
      action: 'Xóa',
      user: userEmail
    });
    sheet.getRange(rowIndex, CONFIG.columns.lichSu + 1).setValue(JSON.stringify(lichSuArray));
    
    // Tô màu xám dòng đã xóa
    sheet.getRange(rowIndex, 1, 1, 21).setBackground('#f0f0f0');
    
    return {
      success: true,
      message: 'Xóa thành công!'
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// API: DUYỆT ĐĂNG KÝ
// ============================================
function approveDangKy(rowIndex, ghiChu) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetNames.dangKy);
    var userEmail = Session.getActiveUser().getEmail();
    var userRole = getUserRole(userEmail);
    
    // Kiểm tra quyền duyệt
    if (userRole !== 'Admin' && userRole !== 'Approver') {
      return {
        success: false,
        error: 'Bạn không có quyền duyệt đăng ký!'
      };
    }
    
    var currentData = sheet.getRange(rowIndex, 1, 1, 21).getValues()[0];
    var diaDiem = currentData[CONFIG.columns.diaDiem];
    
    // Kiểm tra Approver chỉ được duyệt địa điểm của mình
    if (userRole === 'Approver' && !canApproveLocation(userEmail, diaDiem)) {
      return {
        success: false,
        error: 'Bạn chỉ có thể duyệt đăng ký thuộc địa điểm bạn quản lý!'
      };
    }
    
    var now = new Date();
    
    // Cập nhật trạng thái
    sheet.getRange(rowIndex, CONFIG.columns.trangThai + 1).setValue('Đã duyệt');
    sheet.getRange(rowIndex, CONFIG.columns.nguoiDuyet + 1).setValue(userEmail);
    sheet.getRange(rowIndex, CONFIG.columns.ngayDuyet + 1).setValue(now);
    
    // Tô màu xanh
    sheet.getRange(rowIndex, CONFIG.columns.trangThai + 1).setBackground('#d9ead3');
    
    // Thêm log lịch sử
    var oldLichSu = currentData[CONFIG.columns.lichSu];
    var lichSuArray = oldLichSu ? JSON.parse(oldLichSu) : [];
    lichSuArray.push({
      time: formatDateTime(now),
      action: 'Duyệt',
      user: userEmail,
      note: ghiChu || ''
    });
    sheet.getRange(rowIndex, CONFIG.columns.lichSu + 1).setValue(JSON.stringify(lichSuArray));
    
    // Lưu vào lịch sử duyệt
    saveApprovalHistory(currentData[CONFIG.columns.maDangKy], userEmail, 'Duyệt', ghiChu);
    
    // Gửi email thông báo
    var formData = rowToFormData(currentData);
    sendEmailThongBao(currentData[CONFIG.columns.maDangKy], formData, 'Đã duyệt', ghiChu);
    
    return {
      success: true,
      message: 'Duyệt thành công!'
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// API: TỪ CHỐI ĐĂNG KÝ
// ============================================
function rejectDangKy(rowIndex, lyDo) {
  try {
    if (!lyDo || lyDo.trim() === '') {
      return {
        success: false,
        error: 'Vui lòng nhập lý do từ chối!'
      };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetNames.dangKy);
    var userEmail = Session.getActiveUser().getEmail();
    var userRole = getUserRole(userEmail);
    
    // Kiểm tra quyền
    if (userRole !== 'Admin' && userRole !== 'Approver') {
      return {
        success: false,
        error: 'Bạn không có quyền từ chối đăng ký!'
      };
    }
    
    var currentData = sheet.getRange(rowIndex, 1, 1, 21).getValues()[0];
    var diaDiem = currentData[CONFIG.columns.diaDiem];
    
    if (userRole === 'Approver' && !canApproveLocation(userEmail, diaDiem)) {
      return {
        success: false,
        error: 'Bạn chỉ có thể từ chối đăng ký thuộc địa điểm bạn quản lý!'
      };
    }
    
    var now = new Date();
    
    // Cập nhật trạng thái
    sheet.getRange(rowIndex, CONFIG.columns.trangThai + 1).setValue('Từ chối');
    sheet.getRange(rowIndex, CONFIG.columns.nguoiDuyet + 1).setValue(userEmail);
    sheet.getRange(rowIndex, CONFIG.columns.ngayDuyet + 1).setValue(now);
    sheet.getRange(rowIndex, CONFIG.columns.lyDoTuChoi + 1).setValue(lyDo);
    
    // Tô màu đỏ
    sheet.getRange(rowIndex, CONFIG.columns.trangThai + 1).setBackground('#f4cccc');
    
    // Thêm log
    var oldLichSu = currentData[CONFIG.columns.lichSu];
    var lichSuArray = oldLichSu ? JSON.parse(oldLichSu) : [];
    lichSuArray.push({
      time: formatDateTime(now),
      action: 'Từ chối',
      user: userEmail,
      note: lyDo
    });
    sheet.getRange(rowIndex, CONFIG.columns.lichSu + 1).setValue(JSON.stringify(lichSuArray));
    
    // Lưu lịch sử duyệt
    saveApprovalHistory(currentData[CONFIG.columns.maDangKy], userEmail, 'Từ chối', lyDo);
    
    // Gửi email
    var formData = rowToFormData(currentData);
    sendEmailThongBao(currentData[CONFIG.columns.maDangKy], formData, 'Từ chối', lyDo);
    
    return {
      success: true,
      message: 'Từ chối thành công!'
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}
async function init() {
      showLoading();
      try {
        // Setup charts first (empty)
        setupCharts();
        
        // Load user info
        await loadUserInfo();
        
        // Load danh mục
        await loadDanhMuc();
        
        // Load data
        await loadData();
        
        // Load statistics (will update charts)
        await loadStatistics();
        
      } catch (error) {
        console.error('Init error:', error);
        showToast('Lỗi khởi tạo: ' + error.message, 'error');
      } finally {
        hideLoading();
      }
    }
function getThongKe() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetThongKe = ss.getSheetByName(CONFIG.sheetNames.thongKe);
    
    // Lấy 1 lần duy nhất từ B2 đến B13
    var values = sheetThongKe.getRange('B2:B13').getValues();
    
    var data = {
      tongSo: values[0][0],      // B2
      thangNay: values[1][0],    // B3
      choDuyet: values[2][0],    // B4
      daDuyet: values[3][0],     // B5
      tuChoi: values[4][0],      // B6
      theoDiaDiem: {
        'Văn phòng 55': values[7][0],   // B9
        'Bình Phước 1': values[8][0],   // B10
        'Bình Phước 2': values[9][0],   // B11
        'ERC': values[10][0],           // B12
        'Xi nghiệp Bò': values[11][0]   // B13
      }
    };
    return { success: true, data: data };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}
// ============================================
// API: LẤY THÔNG TIN USER
// ============================================
function getUserInfo() {
  try {
    var userEmail = Session.getActiveUser().getEmail();
    var userRole = getUserRole(userEmail);
    var managedLocations = getManagedLocations(userEmail);
    
    return {
      success: true,
      data: {
        email: userEmail,
        role: userRole,
        managedLocations: managedLocations
      }
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// HELPER: LẤY VAI TRÒ USER
// ============================================
function getUserRole(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheetNames.phanQuyen);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][4] === 'Active') {
      return data[i][2]; // Cột C: Vai trò
    }
  }
  
  return 'User'; // Mặc định
}

// ============================================
// HELPER: LẤY ĐỊA ĐIỂM QUẢN LÝ
// ============================================
function getManagedLocations(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheetNames.phanQuyen);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][4] === 'Active') {
      var locations = data[i][3]; // Cột D: Địa điểm quản lý
      if (locations === 'Tất cả') {
        return ['Tất cả'];
      }
      return locations.split(',').map(function(loc) { return loc.trim(); });
    }
  }
  
  return [];
}

// ============================================
// HELPER: KIỂM TRA QUYỀN XEM
// ============================================
function canViewRecord(role, email, rowData) {
  if (role === 'Admin') return true;
  if (role === 'Approver') {
    var managedLocations = getManagedLocations(email);
    return managedLocations.indexOf('Tất cả') !== -1 || 
           managedLocations.indexOf(rowData[CONFIG.columns.diaDiem]) !== -1;
  }
  // User chỉ xem của mình
  return rowData[CONFIG.columns.email] === email;
}

// ============================================
// HELPER: KIỂM TRA QUYỀN SỬA
// ============================================
function canEditRecord(role, email, rowData) {
  if (role === 'Admin') return true;
  
  // Chỉ sửa được khi đang Chờ duyệt và là người tạo
  return rowData[CONFIG.columns.trangThai] === 'Chờ duyệt' && 
         rowData[CONFIG.columns.email] === email;
}

// ============================================
// HELPER: KIỂM TRA QUYỀN XÓA
// ============================================
function canDeleteRecord(role, email, rowData) {
  if (role === 'Admin') return true;
  
  // Chỉ xóa được khi đang Chờ duyệt và là người tạo
  return rowData[CONFIG.columns.trangThai] === 'Chờ duyệt' && 
         rowData[CONFIG.columns.email] === email;
}

// ============================================
// HELPER: KIỂM TRA QUYỀN DUYỆT ĐỊA ĐIỂM
// ============================================
function canApproveLocation(email, diaDiem) {
  var managedLocations = getManagedLocations(email);
  return managedLocations.indexOf('Tất cả') !== -1 || 
         managedLocations.indexOf(diaDiem) !== -1;
}

// ============================================
// HELPER: VALIDATE FORM DATA
// ============================================
function validateFormData(data) {
  if (!data.hoTen || data.hoTen.trim() === '') {
    return { valid: false, error: 'Vui lòng nhập họ tên!' };
  }
  if (!data.chucVu) {
    return { valid: false, error: 'Vui lòng chọn chức vụ!' };
  }
  if (!data.phongBan) {
    return { valid: false, error: 'Vui lòng chọn Loại phòng!' };
  }
  if (!data.ngayDen) {
    return { valid: false, error: 'Vui lòng chọn ngày đến!' };
  }
  if (!data.ngayDi) {
    return { valid: false, error: 'Vui lòng chọn ngày đi!' };
  }
  if (!data.diaDiem) {
    return { valid: false, error: 'Vui lòng chọn địa điểm công tác!' };
  }
  if (!data.email || !isValidEmail(data.email)) {
    return { valid: false, error: 'Email không hợp lệ!' };
  }
  
  // Kiểm tra ngày đi > ngày đến
  var ngayDen = new Date(data.ngayDen);
  var ngayDi = new Date(data.ngayDi);
  if (ngayDi < ngayDen) {
    return { valid: false, error: 'Ngày đi phải sau ngày đến!' };
  }
  
  return { valid: true };
}

// ============================================
// HELPER: VALIDATE EMAIL
// ============================================
function isValidEmail(email) {
  var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(email);
}

// ============================================
// HELPER: TẠO MÃ ĐĂNG KÝ TỰ ĐỘNG
// ============================================
function generateMaDangKy() {
  var now = new Date();
  var dateStr = Utilities.formatDate(now, 'GMT+7', 'yyyyMMdd');
  var prefix = 'DK-' + dateStr + '-';
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheetNames.dangKy);
  var data = sheet.getRange('B:B').getValues();
  
  var maxNum = 0;
  for (var i = 1; i < data.length; i++) {
    var ma = data[i][0];
    if (ma && ma.toString().indexOf(prefix) === 0) {
      var num = parseInt(ma.toString().split('-')[2]);
      if (num > maxNum) maxNum = num;
    }
  }
  
  var newNum = (maxNum + 1).toString().padStart(3, '0');
  return prefix + newNum;
}

// ============================================
// HELPER: FORMAT DATE
// ============================================
function formatDate(date) {
  if (!date) return '';
  var d = new Date(date);
  return Utilities.formatDate(d, 'GMT+7', 'dd/MM/yyyy');
}

function formatDateTime(date) {
  if (!date) return '';
  var d = new Date(date);
  return Utilities.formatDate(d, 'GMT+7', 'dd/MM/yyyy HH:mm:ss');
}

// ============================================
// HELPER: FORMAT DÒNG MỚI
// ============================================
function formatNewRow(sheet, rowIndex) {
  // Format ngày tháng
  sheet.getRange(rowIndex, CONFIG.columns.timestamp + 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
  sheet.getRange(rowIndex, CONFIG.columns.ngayDen + 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(rowIndex, CONFIG.columns.ngayDi + 1).setNumberFormat('dd/mm/yyyy');
  
  // Tô màu vàng cho Chờ duyệt
  sheet.getRange(rowIndex, CONFIG.columns.trangThai + 1).setBackground('#fff2cc');
}

// ============================================
// HELPER: MATCH SEARCH
// ============================================
function matchSearch(row, searchTerm) {
  searchTerm = searchTerm.toLowerCase();
  var searchFields = [
    row[CONFIG.columns.maDangKy],
    row[CONFIG.columns.hoTen],
    row[CONFIG.columns.email],
    row[CONFIG.columns.diaDiem],
    row[CONFIG.columns.phongBan]
  ];
  
  for (var i = 0; i < searchFields.length; i++) {
    if (searchFields[i] && searchFields[i].toString().toLowerCase().indexOf(searchTerm) !== -1) {
      return true;
    }
  }
  return false;
}

// ============================================
// HELPER: ROW TO FORM DATA
// ============================================
function rowToFormData(row) {
  return {
    hoTen: row[CONFIG.columns.hoTen],
    chucVu: row[CONFIG.columns.chucVu],
    phongBan: row[CONFIG.columns.phongBan],
    thongTin: row[CONFIG.columns.thongTin],
    ngayDen: row[CONFIG.columns.ngayDen],
    ngayDi: row[CONFIG.columns.ngayDi],
    phuongTien: row[CONFIG.columns.phuongTien],
    nhaAn: row[CONFIG.columns.nhaAn],
    diaDiem: row[CONFIG.columns.diaDiem],
    email: row[CONFIG.columns.email]
  };
}

// ============================================
// HELPER: LƯU LỊCH SỬ DUYỆT
// ============================================
function saveApprovalHistory(maDangKy, nguoiThaoTac, hanhDong, ghiChu) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheetNames.lichSuDuyet);
  
  var lastRow = sheet.getLastRow();
  var newId = lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() + 1 : 1;
  
  sheet.appendRow([
    newId,
    maDangKy,
    nguoiThaoTac,
    hanhDong,
    ghiChu || '',
    new Date(),
    '' // IP Address (có thể thêm sau)
  ]);
}

// ============================================
// GỬI EMAIL THÔNG BÁO
// ============================================
function sendEmailThongBao(maDangKy, formData, trangThai, ghiChu) {
  try {
    // Lấy email người nhận theo địa điểm
    var emailNguoiNhan = layEmailNguoiNhan(formData.diaDiem);
    
    if (!emailNguoiNhan) {
      return { success: false, error: 'Không tìm thấy email người nhận' };
    }
    
    // Tạo subject theo trạng thái
    var subject = CONFIG.emailSubject;
    if (trangThai === 'Đã duyệt') {
      subject = '[THACO AGRI]  Đăng ký công tác đã được duyệt - ' + maDangKy;
    } else if (trangThai === 'Từ chối') {
      subject = '[THACO AGRI] ❌ Đăng ký công tác bị từ chối - ' + maDangKy;
    }
    
    // Tạo nội dung email
    var emailBody = taoNoiDungEmail(maDangKy, formData, trangThai, ghiChu);
    
    // Gửi email
    var recipients = emailNguoiNhan;
    var cc = formData.email;
    
    if (trangThai === 'Đã duyệt' || trangThai === 'Từ chối') {
      // Chỉ gửi cho người đăng ký và người duyệt
      recipients = formData.email;
      cc = emailNguoiNhan;
    }
    
    MailApp.sendEmail({
      to: recipients,
      cc: cc,
      subject: subject,
      htmlBody: emailBody
    });
    
    return {
      success: true,
      sentTo: recipients
    };
    
  } catch (error) {
    Logger.log('Lỗi gửi email: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// LẤY EMAIL NGƯỜI NHẬN THEO ĐỊA ĐIỂM
// ============================================
function layEmailNguoiNhan(diaDiem) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheetNames.cauHinhEmail);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === diaDiem && data[i][4] === 'Active') {
      return data[i][1];
    }
  }
  
  return null;
}

// ============================================
// TẠO NỘI DUNG EMAIL HTML
// ============================================
function taoNoiDungEmail(maDangKy, formData, trangThai, ghiChu) {
  var statusBadge = '';
  var statusColor = '';
  var statusText = '';
  
  if (trangThai === 'Chờ duyệt') {
    statusBadge = '⏳';
    statusColor = '#ff9800';
    statusText = 'CHỜ DUYỆT';
  } else if (trangThai === 'Đã duyệt') {
    statusBadge = '';
    statusColor = '#4caf50';
    statusText = 'ĐÃ DUYỆT';
  } else if (trangThai === 'Từ chối') {
    statusBadge = '';
    statusColor = '#f44336';
    statusText = 'TỪ CHỐI';
  }
  
  var html = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
      body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        line-height: 1.6;
        color: #333;
        margin: 0;
        padding: 0;
        background-color: #f5f5f5;
      }
      .container {
        max-width: 650px;
        margin: 20px auto;
        background: white;
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      }
      .header {
        background: linear-gradient(135deg, ${CONFIG.brandColor} 0%, #008556 100%);
        color: white;
        padding: 30px;
        text-align: center;
      }
      .header h1 {
        margin: 0;
        font-size: 24px;
        font-weight: 600;
      }
      .status-badge {
        display: inline-block;
        padding: 8px 20px;
        background: ${statusColor};
        color: white;
        border-radius: 20px;
        font-weight: bold;
        margin-top: 10px;
      }
      .content {
        padding: 30px;
      }
      .info-table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
      }
      .info-table td {
        padding: 12px;
        border-bottom: 1px solid #eee;
      }
      .info-table td:first-child {
        font-weight: 600;
        color: ${CONFIG.brandColor};
        width: 40%;
      }
      .highlight-box {
        background: #f0f9f5;
        border-left: 4px solid ${CONFIG.brandColor};
        padding: 15px;
        margin: 20px 0;
        border-radius: 4px;
      }
      .warning-box {
        background: #fff3e0;
        border-left: 4px solid #ff9800;
        padding: 15px;
        margin: 20px 0;
        border-radius: 4px;
      }
      .danger-box {
        background: #ffebee;
        border-left: 4px solid #f44336;
        padding: 15px;
        margin: 20px 0;
        border-radius: 4px;
      }
      .footer {
        background: #f9f9f9;
        padding: 20px 30px;
        text-align: center;
        font-size: 12px;
        color: #666;
      }
      @media only screen and (max-width: 600px) {
        .container {
          margin: 0;
          border-radius: 0;
        }
        .info-table td {
          display: block;
          width: 100% !important;
        }
        .info-table td:first-child {
          padding-bottom: 5px;
        }
      }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="header">
        <h1>${statusBadge} THÔNG BÁO CÔNG TÁC</h1>
        <div class="status-badge">${statusText}</div>
        <p style="margin: 10px 0 0 0; opacity: 0.9;">Mã đăng ký: ${maDangKy}</p>
      </div>
      
      <div class="content">
        <p style="font-size: 16px; color: #333;">Kính gửi Quý Anh/Chị,</p>
        
        <p>Thông tin đăng ký công tác ${trangThai === 'Chờ duyệt' ? 'mới' : ''}:</p>
        
        <table class="info-table">
          <tr>
            <td>👤 Họ tên</td>
            <td><strong>${formData.hoTen}</strong></td>
          </tr>
          <tr>
            <td>💼 Chức vụ</td>
            <td>${formData.chucVu}</td>
          </tr>
          <tr>
            <td>🏢 Loại phòng</td>
            <td>${formData.phongBan}</td>
          </tr>
          <tr>
            <td>📅 Ngày đến</td>
            <td><strong style="color: ${CONFIG.brandColor}">${formatDate(formData.ngayDen)}</strong></td>
          </tr>
          <tr>
            <td>📅 Ngày đi</td>
            <td><strong style="color: ${CONFIG.brandColor}">${formatDate(formData.ngayDi)}</strong></td>
          </tr>
          <tr>
            <td>🚗 Phương tiện</td>
            <td>${formData.phuongTien}</td>
          </tr>
          <tr>
            <td>🍽️ Nhà ăn</td>
            <td>${formData.nhaAn || 'Không'}</td>
          </tr>
          <tr>
            <td>📍 Địa điểm</td>
            <td><strong>${formData.diaDiem}</strong></td>
          </tr>
        </table>
        
        ${formData.thongTin ? `
        <div class="highlight-box">
          <strong style="color: ${CONFIG.brandColor};">📋 Thông tin cần thiết:</strong>
          <p style="margin: 10px 0 0 0;">${formData.thongTin}</p>
        </div>
        ` : ''}
        
        ${trangThai === 'Chờ duyệt' ? `
        <div class="warning-box">
          <strong style="color: #ff9800;">⏳ Đăng ký đang chờ duyệt</strong>
          <p style="margin: 10px 0 0 0;">Vui lòng kiểm tra và duyệt đăng ký này trên hệ thống.</p>
        </div>
        ` : ''}
        
        ${trangThai === 'Đã duyệt' && ghiChu ? `
        <div class="highlight-box">
          <strong style="color: ${CONFIG.brandColor};"> Ghi chú từ người duyệt:</strong>
          <p style="margin: 10px 0 0 0;">${ghiChu}</p>
        </div>
        ` : ''}
        
        ${trangThai === 'Từ chối' ? `
        <div class="danger-box">
          <strong style="color: #f44336;"> Lý do từ chối:</strong>
          <p style="margin: 10px 0 0 0;">${ghiChu || 'Không có lý do cụ thể'}</p>
          <p style="margin: 10px 0 0 0;"><em>Bạn có thể đăng ký lại sau khi điều chỉnh thông tin.</em></p>
        </div>
        ` : ''}
        
        <p style="margin-top: 30px; color: #666;">
          ${trangThai === 'Chờ duyệt' ? 'Vui lòng sắp xếp và chuẩn bị đón tiếp theo thông tin trên.' : ''}
          ${trangThai === 'Đã duyệt' ? 'Đăng ký của bạn đã được xác nhận. Chúc bạn có chuyến công tác hiệu quả!' : ''}
        </p>
      </div>
      
      <div class="footer">
        <p style="margin: 0;">Email này được gửi tự động từ <strong>Hệ thống quản lý công tác THACO AGRI</strong></p>
        <p style="margin: 5px 0 0 0;">© ${new Date().getFullYear()} THACO AGRI. All rights reserved.</p>
      </div>
    </div>
  </body>
  </html>
  `;
  
  return html;
}

// ============================================
// EXPORT EXCEL
// ============================================
function exportToExcelData(filters) {
  var result = getDangKyList(filters);
  if (!result.success) {
    return result;
  }
  
  return {
    success: true,
    data: result.data,
    sheetName: 'Danh sách công tác',
    filename: 'DanhSachCongTac_' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss') + '.xlsx'
  };
}

// ============================================
// SETUP HỆ THỐNG LẦN ĐẦU
// ============================================
function setupSystem() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Cài đặt hệ thống',
    'Bạn có muốn thiết lập các sheet mẫu và công thức tính toán không?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    try {
      // Setup các sheet nếu chưa có
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      
      // Kiểm tra và tạo sheets
      var sheetNames = Object.values(CONFIG.sheetNames);
      for (var i = 0; i < sheetNames.length; i++) {
        if (!ss.getSheetByName(sheetNames[i])) {
          ss.insertSheet(sheetNames[i]);
        }
      }
      
      ui.alert(' Cài đặt thành công!\n\nVui lòng thiết lập cấu trúc dữ liệu theo hướng dẫn.');
      
    } catch (error) {
      ui.alert(' Lỗi: ' + error.toString());
    }
  }
}

// ============================================
// XEM THỐNG KÊ
// ============================================
function viewStatistics() {
  var result = getThongKe();
  if (result.success) {
    var data = result.data;
    var message = 
      '📊 THỐNG KÊ HỆ THỐNG\n\n' +
      '📝 Tổng số đăng ký: ' + data.tongSo + '\n' +
      '📅 Đăng ký tháng này: ' + data.thangNay + '\n\n' +
      '⏳ Chờ duyệt: ' + data.choDuyet + '\n' +
      '✅ Đã duyệt: ' + data.daDuyet + '\n' +
      '❌ Từ chối: ' + data.tuChoi + '\n\n' +
      '📍 THEO ĐỊA ĐIỂM:\n' +
      '- Văn phòng 55: ' + data.theoDiaDiem['Văn phòng 55'] + '\n' +
      '- Bình Phước 1: ' + data.theoDiaDiem['Bình Phước 1'] + '\n' +
      '- Bình Phước 2: ' + data.theoDiaDiem['Bình Phước 2'] + '\n' +
      '- ERC: ' + data.theoDiaDiem['ERC'] + '\n' +
      '- Xi nghiệp Bò: ' + data.theoDiaDiem['Xi nghiệp Bò'];
    
    SpreadsheetApp.getUi().alert(message);
  }
}

// ============================================
// DỌN DẸP DỮ LIỆU CŨ (>6 THÁNG)
// ============================================
function cleanOldData() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Dọn dẹp dữ liệu',
    'Bạn có muốn xóa các đăng ký cũ hơn 6 tháng không?\n(Chỉ xóa dữ liệu đã đánh dấu xóa)',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(CONFIG.sheetNames.dangKy);
      var data = sheet.getDataRange().getValues();
      
      var sixMonthsAgo = new Date();
      sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
      
      var deletedCount = 0;
      
      // Duyệt từ dưới lên để xóa không ảnh hưởng index
      for (var i = data.length - 1; i > 0; i--) {
        var row = data[i];
        var timestamp = new Date(row[CONFIG.columns.timestamp]);
        var daXoa = row[CONFIG.columns.daXoa];
        
        if (daXoa === true && timestamp < sixMonthsAgo) {
          sheet.deleteRow(i + 1);
          deletedCount++;
        }
      }
      
      ui.alert(' Đã xóa ' + deletedCount + ' dòng dữ liệu cũ!');
      
    } catch (error) {
      ui.alert(' Lỗi: ' + error.toString());
    }
  }
}


function testSystem() {
  try {
    Logger.log('=== TEST BẮT ĐẦU ===');
    
    // Test 1: Kiểm tra sheets
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetNames = ['Đăng ký công tác', 'Cấu hình Email', 'Danh mục', 'Phân quyền', 'Lịch sử duyệt', 'Thống kê'];
    
    Logger.log('Test 1: Checking sheets...');
    sheetNames.forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      if (sheet) {
        Logger.log(' Sheet "' + name + '" OK');
      } else {
        Logger.log(' Sheet "' + name + '" MISSING!');
      }
    });
    
    // Test 2: Kiểm tra user
    Logger.log('\nTest 2: Checking user...');
    var userEmail = Session.getActiveUser().getEmail();
    Logger.log('User email: ' + userEmail);
    
    // Test 3: Kiểm tra functions
    Logger.log('\nTest 3: Checking functions...');
    
    var danhMuc = getDanhMuc();
    Logger.log('getDanhMuc: ' + (danhMuc.success ? ' OK' : ' FAILED'));
    
    var thongKe = getThongKe();
    Logger.log('getThongKe: ' + (thongKe.success ? ' OK' : ' FAILED'));
    
    var userInfo = getUserInfo();
    Logger.log('getUserInfo: ' + (userInfo.success ? ' OK' : ' FAILED'));
    
    Logger.log('\n=== TEST HOÀN THÀNH ===');
    
  } catch (error) {
    Logger.log(' LỖI: ' + error.toString());
  }
}