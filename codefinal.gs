// ============================================
// HỆ THỐNG QUẢN LÝ CÔNG TÁC - THACO AGRI
// Version 2.0 - Viết lại hoàn toàn
// ============================================

// CONFIGURATION
var CONFIG = {
  SHEET_NAMES: {
    RECEPTION: 'Đón tiếp khách',
    BUSINESS_TRIP: 'KLH Đi công tác',
    CONFIG_EMAIL: 'Cấu hình Email',
    CATEGORIES: 'Danh mục',
    PERMISSIONS: 'Phân quyền'
  },
  
  BRAND_COLOR: '#00682B',
  
  EMAIL_RECEPTION: ['openaibku@gmail.com'],
  EMAIL_BUSINESS_TRIP: ['vovantrungphone2002@gmail.com', 'trung@thagrico.vn', 'phu@thagrico.vn']
};

// ============================================
// MENU
// ============================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏢 Quản lý Công tác')
    .addItem('📊 Mở Dashboard', 'openDashboard')
    .addSeparator()
    .addItem('🧪 Test System', 'testSystem')
    .addToUi();
}

function openDashboard() {
  var html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Quản lý Công tác - THACO AGRI')
    .setWidth(1400)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard Quản lý Công tác');
}

// ============================================
// WEB APP
// ============================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Hệ thống Quản lý Công tác - THACO AGRI')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================
// API: GET USER INFO
// ============================================
function getUserInfo() {
  try {
    var email = Session.getActiveUser().getEmail();
    var role = getUserRole(email);
    
    return {
      success: true,
      email: email,
      role: role
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

function getUserRole(email) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PERMISSIONS);
    
    if (!sheet) return 'User';
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === email && data[i][4] === 'Active') {
        return data[i][2] || 'User';
      }
    }
    return 'User';
  } catch (error) {
    return 'User';
  }
}

// ============================================
// API: GET CATEGORIES
// ============================================
function getDanhMuc() {
  return {
    success: true,
    data: {
      loaiKhach: ['VPDH', 'Địa phương', 'VIP'],
      gioiTinh: ['Nam', 'Nữ'],
      noiAnO: ['Nhà khách VP55', 'XN BP1', 'XN BP2', 'XN ERC', 'XN BÒ SS', 'Tổng kho'],
      trangThai: ['Chờ xử lý', 'Đã xử lý'],
      diaDiem: ['Phnom Penh', 'Kratie', 'VPDH', 'Lào', 'Kounmom', 'Thaco']
    }
  };
}

// ============================================
// API: RECEPTION (ĐÓN TIẾP KHÁCH)
// ============================================
function getReceptionList(filters) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.RECEPTION);
    
    if (!sheet) {
      return { success: false, error: 'Sheet không tồn tại' };
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, data: [], total: 0 };
    }
    
    var data = sheet.getRange(2, 1, lastRow - 1, 23).getValues();
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      // Skip deleted rows
      if (row[22] === true || row[22] === 'TRUE' || row[22] === 'true') continue;
      
      // Parse guest list
      var guestList = [];
      try {
        if (row[8]) {
          guestList = typeof row[8] === 'string' ? JSON.parse(row[8]) : row[8];
        }
      } catch (e) {
        guestList = [];
      }
      
      var record = {
        rowIndex: i + 2,
        timestamp: formatDateTime(row[0]),
        maDonTiep: row[1] || '',
        doanKhach: row[2] || '',
        loaiKhach: row[3] || '',
        gioDen: row[4] || '',
        ngayDen: formatDate(row[5]),
        ngayDi: formatDate(row[6]),
        soLuong: row[7] || 0,
        danhSachKhach: guestList,
        nguoiDangKy: row[9] || '',
        emailNguoiDangKy: row[10] || '',
        ngayDangKy: formatDateTime(row[11]),
        trangThai: row[12] || '',
        noiAnO: row[13] || '',
        phongO: row[14] || '',
        phuongTien: row[15] || '',
        hoiHop: row[16] || '',
        nguoiXuLy: row[17] || '',
        ngayXuLy: formatDateTime(row[19])
      };
      
      // Apply filters
      if (filters) {
        if (filters.trangThai && record.trangThai !== filters.trangThai) continue;
        if (filters.loaiKhach && record.loaiKhach !== filters.loaiKhach) continue;
        if (filters.search) {
          var searchLower = filters.search.toLowerCase();
          var match = false;
          if (record.maDonTiep.toLowerCase().indexOf(searchLower) >= 0) match = true;
          if (record.doanKhach.toLowerCase().indexOf(searchLower) >= 0) match = true;
          if (!match) continue;
        }
      }
      
      result.push(record);
    }
    
    return {
      success: true,
      data: result,
      total: result.length
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      stack: error.stack
    };
  }
}

function submitReception(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.RECEPTION);
    
    if (!sheet) {
      return { success: false, error: 'Sheet không tồn tại' };
    }
    
    var user = Session.getActiveUser().getEmail();
    var userName = user.split('@')[0];
    var timestamp = new Date();
    var maDonTiep = 'DT-' + Utilities.formatDate(timestamp, 'GMT+7', 'yyyyMMdd') + '-' + String(sheet.getLastRow()).padStart(3, '0');
    
    var rowData = [
      timestamp,                          // A: Timestamp
      maDonTiep,                          // B: Mã đón tiếp
      formData.doanKhach,                 // C: Đoàn khách
      formData.loaiKhach,                 // D: Loại khách
      formData.gioDen,                    // E: Giờ đến
      new Date(formData.ngayDen),         // F: Ngày đến
      new Date(formData.ngayDi),          // G: Ngày đi
      formData.soLuong,                   // H: Số lượng
      JSON.stringify(formData.danhSachKhach), // I: Danh sách khách
      userName,                           // J: Người đăng ký
      user,                               // K: Email người đăng ký
      timestamp,                          // L: Ngày đăng ký
      'Chờ xử lý',                        // M: Trạng thái
      '',                                 // N: Nơi ăn ở
      '',                                 // O: Phòng ở
      '',                                 // P: Phương tiện
      '',                                 // Q: Hội họp
      '',                                 // R: Người xử lý
      '',                                 // S: Email người xử lý
      '',                                 // T: Ngày xử lý
      false,                              // U: Email đã gửi
      '',                                 // V: Thời gian gửi
      false                               // W: Đã xóa
    ];
    
    sheet.appendRow(rowData);
    
    // Send email notification
    sendReceptionNotificationEmail(maDonTiep, formData, user, userName);
    sendReceptionConfirmationEmail(maDonTiep, formData, user, userName);
    
    return {
      success: true,
      message: 'Đăng ký thành công! Mã đón tiếp: ' + maDonTiep,
      maDonTiep: maDonTiep
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

function updateReceptionProcessing(rowIndex, formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.RECEPTION);
    
    if (!sheet) {
      return { success: false, error: 'Sheet không tồn tại' };
    }
    
    var user = Session.getActiveUser().getEmail();
    var userName = user.split('@')[0];
    var timestamp = new Date();
    
    sheet.getRange(rowIndex, 14).setValue(formData.noiAnO);
    sheet.getRange(rowIndex, 15).setValue(formData.phongO);
    sheet.getRange(rowIndex, 16).setValue(formData.phuongTien);
    sheet.getRange(rowIndex, 17).setValue(formData.hoiHop);
    sheet.getRange(rowIndex, 18).setValue(userName);
    sheet.getRange(rowIndex, 19).setValue(user);
    sheet.getRange(rowIndex, 20).setValue(timestamp);
    sheet.getRange(rowIndex, 13).setValue('Đã xử lý');
    
    return {
      success: true,
      message: 'Cập nhật xử lý thành công!'
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// API: BUSINESS TRIP (ĐI CÔNG TÁC)
// ============================================
function getBusinessTripList(filters) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.BUSINESS_TRIP);
    
    if (!sheet) {
      return { success: false, error: 'Sheet không tồn tại' };
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, data: [], total: 0 };
    }
    
    var data = sheet.getRange(2, 1, lastRow - 1, 18).getValues();
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      // Skip deleted rows
      if (row[17] === true || row[17] === 'TRUE' || row[17] === 'true') continue;
      
      // Parse member list
      var memberList = [];
      try {
        if (row[6]) {
          memberList = typeof row[6] === 'string' ? JSON.parse(row[6]) : row[6];
        }
      } catch (e) {
        memberList = [];
      }
      
      var record = {
        rowIndex: i + 2,
        timestamp: formatDateTime(row[0]),
        maDoan: row[1] || '',
        diaDiem: row[2] || '',
        ngayDi: formatDate(row[3]),
        ngayVe: formatDate(row[4]),
        soLuong: row[5] || 0,
        danhSachThanhVien: memberList,
        truongDoan: row[7] || '',
        datPhong: row[8] || false,
        comTrua: row[9] || false,
        xeDuaDon: row[10] || false,
        hoTroKhac: row[11] || '',
        nguoiTao: row[12] || '',
        emailNguoiTao: row[13] || '',
        ngayTao: formatDateTime(row[14])
      };
      
      // Apply filters
      if (filters) {
        if (filters.diaDiem && record.diaDiem !== filters.diaDiem) continue;
        if (filters.search) {
          var searchLower = filters.search.toLowerCase();
          var match = false;
          if (record.maDoan.toLowerCase().indexOf(searchLower) >= 0) match = true;
          if (record.diaDiem.toLowerCase().indexOf(searchLower) >= 0) match = true;
          if (record.truongDoan.toLowerCase().indexOf(searchLower) >= 0) match = true;
          if (!match) continue;
        }
      }
      
      result.push(record);
    }
    
    return {
      success: true,
      data: result,
      total: result.length
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      stack: error.stack
    };
  }
}

function submitBusinessTrip(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.BUSINESS_TRIP);
    
    if (!sheet) {
      return { success: false, error: 'Sheet không tồn tại' };
    }
    
    var user = Session.getActiveUser().getEmail();
    var userName = user.split('@')[0];
    var timestamp = new Date();
    var maDoan = 'DCT-' + Utilities.formatDate(timestamp, 'GMT+7', 'yyyyMMdd') + '-' + String(sheet.getLastRow()).padStart(3, '0');
    
    var rowData = [
      timestamp,                          // A: Timestamp
      maDoan,                             // B: Mã đoàn
      formData.diaDiem,                   // C: Địa điểm
      new Date(formData.ngayDi),          // D: Ngày đi
      new Date(formData.ngayVe),          // E: Ngày về
      formData.soLuong,                   // F: Số lượng
      JSON.stringify(formData.danhSachThanhVien), // G: Danh sách thành viên
      formData.truongDoan,                // H: Trưởng đoàn
      formData.datPhong || false,         // I: Đặt phòng
      formData.comTrua || false,          // J: Cơm trưa
      formData.xeDuaDon || false,         // K: Xe đưa đón
      formData.hoTroKhac || '',           // L: Hỗ trợ khác
      userName,                           // M: Người tạo
      user,                               // N: Email người tạo
      timestamp,                          // O: Ngày tạo
      false,                              // P: Email đã gửi
      '',                                 // Q: Thời gian gửi
      false                               // R: Đã xóa
    ];
    
    sheet.appendRow(rowData);
    
    // Send email notification
    sendBusinessTripNotificationEmail(maDoan, formData, user, userName);
    sendBusinessTripConfirmationEmail(maDoan, formData, user, userName);
    
    return {
      success: true,
      message: 'Đăng ký thành công! Mã đoàn: ' + maDoan,
      maDoan: maDoan
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// EMAIL FUNCTIONS
// ============================================
function sendReceptionNotificationEmail(maDonTiep, formData, userEmail, userName) {
  try {
    var guestListHtml = '<table style="width:100%;border-collapse:collapse;margin:15px 0"><thead><tr><th style="padding:10px;background:#00682B;color:white;border:1px solid #ddd">STT</th><th style="padding:10px;background:#00682B;color:white;border:1px solid #ddd">Họ tên</th><th style="padding:10px;background:#00682B;color:white;border:1px solid #ddd">Giới tính</th><th style="padding:10px;background:#00682B;color:white;border:1px solid #ddd">Chức danh</th></tr></thead><tbody>';
    
    for (var i = 0; i < formData.danhSachKhach.length; i++) {
      var guest = formData.danhSachKhach[i];
      guestListHtml += '<tr><td style="padding:8px;border:1px solid #ddd;text-align:center">' + (i + 1) + '</td><td style="padding:8px;border:1px solid #ddd">' + guest.ten + '</td><td style="padding:8px;border:1px solid #ddd;text-align:center">' + guest.gioiTinh + '</td><td style="padding:8px;border:1px solid #ddd">' + guest.chucDanh + '</td></tr>';
    }
    guestListHtml += '</tbody></table>';
    
    var htmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="font-family:Arial,sans-serif;line-height:1.6;color:#333"><div style="max-width:800px;margin:0 auto;padding:20px"><div style="background:linear-gradient(135deg,#00682B 0%,#004d1f 100%);color:white;padding:30px;border-radius:10px 10px 0 0;text-align:center"><h1 style="margin:0">🔔 THÔNG BÁO ĐOÀN KHÁCH ĐẾN</h1><p style="margin:10px 0 0 0">THACO AGRI - KLH SNUOL</p></div><div style="background:#fff;padding:30px;border:1px solid #ddd;border-top:none"><div style="background:#fff3cd;border-left:4px solid #ffc107;padding:15px;margin:15px 0;border-radius:5px"><strong>⚠️ CẦN XỬ LÝ:</strong> Đoàn khách sắp đến, vui lòng chuẩn bị và cập nhật thông tin.</div><h3 style="color:#00682B">📋 Thông tin đoàn khách</h3><table style="width:100%;margin:15px 0"><tr><td style="padding:8px;font-weight:bold;width:200px">Mã đón tiếp:</td><td style="padding:8px"><strong style="color:#00682B">' + maDonTiep + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Đoàn khách:</td><td style="padding:8px"><strong>' + formData.doanKhach + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Loại khách:</td><td style="padding:8px"><span style="padding:4px 12px;background:#17a2b8;color:white;border-radius:4px;font-size:12px">' + formData.loaiKhach + '</span></td></tr><tr><td style="padding:8px;font-weight:bold">Giờ đến:</td><td style="padding:8px">' + formData.gioDen + '</td></tr><tr><td style="padding:8px;font-weight:bold">Ngày đến:</td><td style="padding:8px"><strong>' + formatDate(formData.ngayDen) + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Ngày đi:</td><td style="padding:8px"><strong>' + formatDate(formData.ngayDi) + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Số lượng:</td><td style="padding:8px">' + formData.soLuong + ' người</td></tr><tr><td style="padding:8px;font-weight:bold">Người đăng ký:</td><td style="padding:8px">' + userName + '</td></tr></table><h3 style="color:#00682B">👥 Danh sách khách</h3>' + guestListHtml + '<p style="margin-top:30px;padding:15px;background:#e7f3ff;border-left:4px solid #2196f3;border-radius:5px"><strong>📌 Lưu ý:</strong> Vui lòng cập nhật thông tin <strong>Nơi ăn ở</strong>, <strong>Phòng ở</strong>, <strong>Phương tiện</strong>, <strong>Hội họp</strong> trong hệ thống.</p></div><div style="background:#f8f9fa;padding:20px;border-radius:0 0 10px 10px;text-align:center;font-size:12px;color:#666"><p><strong>THACO AGRI - KLH SNUOL</strong></p><p>Email tự động, vui lòng không trả lời</p></div></div></body></html>';
    
    MailApp.sendEmail({
      to: CONFIG.EMAIL_RECEPTION.join(','),
      subject: '[KLH] Đoàn khách đến - ' + maDonTiep + ' - ' + formData.doanKhach,
      htmlBody: htmlBody
    });
    
    return true;
  } catch (error) {
    Logger.log('Error sending reception notification email: ' + error.toString());
    return false;
  }
}

function sendReceptionConfirmationEmail(maDonTiep, formData, userEmail, userName) {
  try {
    var guestListHtml = '<table style="width:100%;border-collapse:collapse;margin:15px 0"><thead><tr><th style="padding:10px;background:#00682B;color:white;border:1px solid #ddd">STT</th><th style="padding:10px;background:#00682B;color:white;border:1px solid #ddd">Họ tên</th><th style="padding:10px;background:#00682B;color:white;border:1px solid #ddd">Giới tính</th><th style="padding:10px;background:#00682B;color:white;border:1px solid #ddd">Chức danh</th></tr></thead><tbody>';
    
    for (var i = 0; i < formData.danhSachKhach.length; i++) {
      var guest = formData.danhSachKhach[i];
      guestListHtml += '<tr><td style="padding:8px;border:1px solid #ddd;text-align:center">' + (i + 1) + '</td><td style="padding:8px;border:1px solid #ddd">' + guest.ten + '</td><td style="padding:8px;border:1px solid #ddd;text-align:center">' + guest.gioiTinh + '</td><td style="padding:8px;border:1px solid #ddd">' + guest.chucDanh + '</td></tr>';
    }
    guestListHtml += '</tbody></table>';
    
    var htmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="font-family:Arial,sans-serif;line-height:1.6;color:#333"><div style="max-width:800px;margin:0 auto;padding:20px"><div style="background:linear-gradient(135deg,#00682B 0%,#004d1f 100%);color:white;padding:30px;border-radius:10px 10px 0 0;text-align:center"><h1 style="margin:0">✅ XÁC NHẬN ĐĂNG KÝ ĐÓN TIẾP</h1><p style="margin:10px 0 0 0">THACO AGRI - KLH SNUOL</p></div><div style="background:#fff;padding:30px;border:1px solid #ddd;border-top:none"><div style="background:#d4edda;border-left:4px solid #28a745;padding:15px;margin:15px 0;border-radius:5px"><h3 style="margin-top:0;color:#28a745">🎉 Đăng ký đón tiếp thành công!</h3><p style="margin:5px 0">Cảm ơn bạn đã đăng ký. Thông tin đoàn khách của bạn đã được ghi nhận.</p></div><h3 style="color:#00682B">📋 Thông tin đoàn khách</h3><table style="width:100%;margin:15px 0"><tr><td style="padding:8px;font-weight:bold;width:200px">Mã đón tiếp:</td><td style="padding:8px"><strong style="color:#00682B">' + maDonTiep + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Đoàn khách:</td><td style="padding:8px"><strong>' + formData.doanKhach + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Loại khách:</td><td style="padding:8px">' + formData.loaiKhach + '</td></tr><tr><td style="padding:8px;font-weight:bold">Giờ đến:</td><td style="padding:8px">' + formData.gioDen + '</td></tr><tr><td style="padding:8px;font-weight:bold">Ngày đến:</td><td style="padding:8px"><strong>' + formatDate(formData.ngayDen) + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Ngày đi:</td><td style="padding:8px"><strong>' + formatDate(formData.ngayDi) + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Số lượng:</td><td style="padding:8px">' + formData.soLuong + ' người</td></tr></table><h3 style="color:#00682B">👥 Danh sách khách</h3>' + guestListHtml + '</div><div style="background:#f8f9fa;padding:20px;border-radius:0 0 10px 10px;text-align:center;font-size:12px;color:#666"><p><strong>THACO AGRI - KLH SNUOL</strong></p><p>Email xác nhận tự động</p></div></div></body></html>';
    
    MailApp.sendEmail({
      to: userEmail,
      subject: '[XÁC NHẬN] Đăng ký đón tiếp - ' + maDonTiep,
      htmlBody: htmlBody
    });
    
    return true;
  } catch (error) {
    Logger.log('Error sending reception confirmation email: ' + error.toString());
    return false;
  }
}

function sendBusinessTripNotificationEmail(maDoan, formData, userEmail, userName) {
  try {
    var memberListHtml = '<table style="width:100%;border-collapse:collapse;margin:15px 0"><thead><tr><th style="padding:10px;background:#00682B;color:white;border:1px solid #ddd">STT</th><th style="padding:10px;background:#00682B;color:white;border:1px solid #ddd">Họ tên</th><th style="padding:10px;background:#00682B;color:white;border:1px solid #ddd">Chức danh</th></tr></thead><tbody>';
    
    for (var i = 0; i < formData.danhSachThanhVien.length; i++) {
      var member = formData.danhSachThanhVien[i];
      memberListHtml += '<tr><td style="padding:8px;border:1px solid #ddd;text-align:center">' + (i + 1) + '</td><td style="padding:8px;border:1px solid #ddd">' + member.ten + '</td><td style="padding:8px;border:1px solid #ddd">' + member.chucDanh + '</td></tr>';
    }
    memberListHtml += '</tbody></table>';
    
    var supportList = [];
    if (formData.datPhong) supportList.push('🏨 Đặt phòng');
    if (formData.comTrua) supportList.push('🍽️ Cơm trưa');
    if (formData.xeDuaDon) supportList.push('🚗 Xe đưa đón');
    if (formData.hoTroKhac) supportList.push('📝 Khác: ' + formData.hoTroKhac);
    var supportHtml = supportList.length > 0 ? supportList.join('<br>') : 'Không yêu cầu hỗ trợ';
    
    var htmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="font-family:Arial,sans-serif;line-height:1.6;color:#333"><div style="max-width:800px;margin:0 auto;padding:20px"><div style="background:linear-gradient(135deg,#00682B 0%,#004d1f 100%);color:white;padding:30px;border-radius:10px 10px 0 0;text-align:center"><h1 style="margin:0">🔔 THÔNG BÁO ĐOÀN CÔNG TÁC</h1><p style="margin:10px 0 0 0">THACO AGRI - KLH SNUOL</p></div><div style="background:#fff;padding:30px;border:1px solid #ddd;border-top:none"><h3 style="color:#00682B">📋 Thông tin đoàn công tác</h3><table style="width:100%;margin:15px 0"><tr><td style="padding:8px;font-weight:bold;width:200px">Mã đoàn:</td><td style="padding:8px"><strong style="color:#00682B">' + maDoan + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Địa điểm:</td><td style="padding:8px"><strong>' + formData.diaDiem + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Trưởng đoàn:</td><td style="padding:8px"><strong>' + formData.truongDoan + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Ngày đi:</td><td style="padding:8px"><strong>' + formatDate(formData.ngayDi) + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Ngày về:</td><td style="padding:8px"><strong>' + formatDate(formData.ngayVe) + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Số lượng:</td><td style="padding:8px">' + formData.soLuong + ' người</td></tr><tr><td style="padding:8px;font-weight:bold">Người tạo:</td><td style="padding:8px">' + userName + '</td></tr></table><h3 style="color:#00682B">👥 Danh sách thành viên</h3>' + memberListHtml + '<h3 style="color:#00682B">🎯 Yêu cầu hỗ trợ</h3><div style="padding:15px;background:#f8f9fa;border-left:4px solid #00682B;border-radius:5px">' + supportHtml + '</div></div><div style="background:#f8f9fa;padding:20px;border-radius:0 0 10px 10px;text-align:center;font-size:12px;color:#666"><p><strong>THACO AGRI - KLH SNUOL</strong></p><p>Email tự động, vui lòng không trả lời</p></div></div></body></html>';
    
    MailApp.sendEmail({
      to: CONFIG.EMAIL_BUSINESS_TRIP.join(','),
      subject: '[KLH] Đoàn công tác - ' + maDoan + ' - ' + formData.diaDiem,
      htmlBody: htmlBody
    });
    
    return true;
  } catch (error) {
    Logger.log('Error sending business trip notification email: ' + error.toString());
    return false;
  }
}

function sendBusinessTripConfirmationEmail(maDoan, formData, userEmail, userName) {
  try {
    var memberListHtml = '<table style="width:100%;border-collapse:collapse;margin:15px 0"><thead><tr><th style="padding:10px;background:#00682B;color:white;border:1px solid #ddd">STT</th><th style="padding:10px;background:#00682B;color:white;border:1px solid #ddd">Họ tên</th><th style="padding:10px;background:#00682B;color:white;border:1px solid #ddd">Chức danh</th></tr></thead><tbody>';
    
    for (var i = 0; i < formData.danhSachThanhVien.length; i++) {
      var member = formData.danhSachThanhVien[i];
      memberListHtml += '<tr><td style="padding:8px;border:1px solid #ddd;text-align:center">' + (i + 1) + '</td><td style="padding:8px;border:1px solid #ddd">' + member.ten + '</td><td style="padding:8px;border:1px solid #ddd">' + member.chucDanh + '</td></tr>';
    }
    memberListHtml += '</tbody></table>';
    
    var supportList = [];
    if (formData.datPhong) supportList.push('🏨 Đặt phòng');
    if (formData.comTrua) supportList.push('🍽️ Cơm trưa');
    if (formData.xeDuaDon) supportList.push('🚗 Xe đưa đón');
    if (formData.hoTroKhac) supportList.push('📝 Khác: ' + formData.hoTroKhac);
    var supportHtml = supportList.length > 0 ? supportList.join('<br>') : 'Không yêu cầu hỗ trợ';
    
    var htmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="font-family:Arial,sans-serif;line-height:1.6;color:#333"><div style="max-width:800px;margin:0 auto;padding:20px"><div style="background:linear-gradient(135deg,#00682B 0%,#004d1f 100%);color:white;padding:30px;border-radius:10px 10px 0 0;text-align:center"><h1 style="margin:0">✅ XÁC NHẬN ĐĂNG KÝ CÔNG TÁC</h1><p style="margin:10px 0 0 0">THACO AGRI - KLH SNUOL</p></div><div style="background:#fff;padding:30px;border:1px solid #ddd;border-top:none"><div style="background:#d4edda;border-left:4px solid #28a745;padding:15px;margin:15px 0;border-radius:5px"><h3 style="margin-top:0;color:#28a745">🎉 Đăng ký công tác thành công!</h3><p style="margin:5px 0">Cảm ơn bạn đã đăng ký. Thông tin đoàn của bạn đã được ghi nhận.</p></div><h3 style="color:#00682B">📋 Thông tin đoàn công tác</h3><table style="width:100%;margin:15px 0"><tr><td style="padding:8px;font-weight:bold;width:200px">Mã đoàn:</td><td style="padding:8px"><strong style="color:#00682B">' + maDoan + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Địa điểm:</td><td style="padding:8px"><strong>' + formData.diaDiem + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Trưởng đoàn:</td><td style="padding:8px"><strong>' + formData.truongDoan + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Ngày đi:</td><td style="padding:8px"><strong>' + formatDate(formData.ngayDi) + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Ngày về:</td><td style="padding:8px"><strong>' + formatDate(formData.ngayVe) + '</strong></td></tr><tr><td style="padding:8px;font-weight:bold">Số lượng:</td><td style="padding:8px">' + formData.soLuong + ' người</td></tr></table><h3 style="color:#00682B">👥 Danh sách thành viên</h3>' + memberListHtml + '<h3 style="color:#00682B">🎯 Yêu cầu hỗ trợ</h3><div style="padding:15px;background:#f8f9fa;border-left:4px solid #00682B;border-radius:5px">' + supportHtml + '</div></div><div style="background:#f8f9fa;padding:20px;border-radius:0 0 10px 10px;text-align:center;font-size:12px;color:#666"><p><strong>THACO AGRI - KLH SNUOL</strong></p><p>Email xác nhận tự động</p></div></div></body></html>';
    
    MailApp.sendEmail({
      to: userEmail,
      subject: '[XÁC NHẬN] Đăng ký công tác - ' + maDoan,
      htmlBody: htmlBody
    });
    
    return true;
  } catch (error) {
    Logger.log('Error sending business trip confirmation email: ' + error.toString());
    return false;
  }
}

// ============================================
// UTILITY FUNCTIONS
// ============================================
function formatDateTime(date) {
  if (!date) return '';
  try {
    return Utilities.formatDate(new Date(date), 'GMT+7', 'dd/MM/yyyy HH:mm');
  } catch (e) {
    return '';
  }
}

function formatDate(date) {
  if (!date) return '';
  try {
    return Utilities.formatDate(new Date(date), 'GMT+7', 'dd/MM/yyyy');
  } catch (e) {
    return '';
  }
}

// ============================================
// TEST FUNCTION
// ============================================
function testSystem() {
  var ui = SpreadsheetApp.getUi();
  
  Logger.log('===== SYSTEM TEST START =====');
  
  // Test getUserInfo
  var userInfo = getUserInfo();
  Logger.log('getUserInfo: ' + JSON.stringify(userInfo));
  
  // Test getDanhMuc
  var danhMuc = getDanhMuc();
  Logger.log('getDanhMuc: ' + JSON.stringify(danhMuc));
  
  // Test getReceptionList
  var receptionList = getReceptionList({});
  Logger.log('getReceptionList: ' + JSON.stringify(receptionList));
  
  // Test getBusinessTripList
  var tripList = getBusinessTripList({});
  Logger.log('getBusinessTripList: ' + JSON.stringify(tripList));
  
  Logger.log('===== SYSTEM TEST END =====');
  
  if (userInfo.success && danhMuc.success && receptionList.success && tripList.success) {
    ui.alert('✅ Test thành công!\n\nTất cả functions hoạt động bình thường.\n\nReception: ' + receptionList.total + ' records\nBusiness Trip: ' + tripList.total + ' records');
  } else {
    ui.alert('❌ Test thất bại!\n\nCó lỗi xảy ra. Xem Logs để biết chi tiết.');
  }
}

function testDashboardAPI() {
  Logger.clear();
  
  Logger.log('===== TEST API CALLS =====');
  
  // Test 1: getUserInfo
  var userResult = getUserInfo();
  Logger.log('getUserInfo: ' + JSON.stringify(userResult));
  
  // Test 2: getDanhMuc
  var danhMucResult = getDanhMuc();
  Logger.log('getDanhMuc: ' + JSON.stringify(danhMucResult));
  
  // Test 3: getReceptionList
  var receptionResult = getReceptionList({});
  Logger.log('getReceptionList: ' + JSON.stringify(receptionResult));
  
  if (receptionResult.success) {
    Logger.log('✅ Reception data count: ' + receptionResult.data.length);
  } else {
    Logger.log('❌ Reception error: ' + receptionResult.error);
  }
  
  Logger.log('===== END TEST =====');
}

