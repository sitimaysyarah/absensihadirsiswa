/**
 * APLIKASI ABSENSI SISWA (DAFTAR HADIR & PULANG BARCODE/QR)
 * Sekolah: SMP NEGERI 1 TANA TIDUNG
 * Tahun Ajaran: 2026/2027 - Genap
 */

// CONFIGURATION PARAMETERS
const CONFIG = {
  APP_TITLE: "Aplikasi Absensi Siswa",
  SCHOOL_NAME: "SMP NEGERI 1 TANA TIDUNG",
  ACADEMIC_YEAR: "2026/2027 - Genap",
  ADMIN_PASSWORD: "admin123",
  TIMEZONE: "Asia/Makassar", // WITA (GMT+8)
  TIME_LIMITS: {
    HADIR_START: "06:40",
    HADIR_END: "07:20",
    PULANG_START: "14:00",
    PULANG_END: "16:00"
  },
  SHEETS: {
    MASTER: "Master_Siswa",
    LOG: "Log_Absensi",
    REKAP: "Rekap_Bulanan"
  }
};

/**
 * Entry point Web App (GET HTTP Request)
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle(CONFIG.APP_TITLE + " - " + CONFIG.SCHOOL_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Mendapatkan konfigurasi aplikasi untuk frontend
 */
function getAppConfig() {
  return {
    schoolName: CONFIG.SCHOOL_NAME,
    academicYear: CONFIG.ACADEMIC_YEAR,
    appTitle: CONFIG.APP_TITLE,
    timeLimits: CONFIG.TIME_LIMITS
  };
}

/**
 * FUNGSI SETUP DATABASE AUTOMATIS
 * Jalankan fungsi ini 1 kali dari Google Apps Script Editor saat pertama kali setup project.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Sheet Master_Siswa
  let masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER);
  if (!masterSheet) {
    masterSheet = ss.insertSheet(CONFIG.SHEETS.MASTER);
  }
  masterSheet.clear();
  masterSheet.appendRow(["NISN", "Nama_Siswa", "Kelas", "Jenis_Kelamin", "URL_Barcode_QR"]);
  
  // Sampel Data Siswa
  const sampleStudents = [
    ["1001", "Ahmad Fauzi", "7A", "Laki-laki", "https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=1001"],
    ["1002", "Siti Nurhaliza", "7A", "Perempuan", "https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=1002"],
    ["1003", "Budi Santoso", "7B", "Laki-laki", "https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=1003"],
    ["1004", "Dewi Lestari", "8A", "Perempuan", "https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=1004"],
    ["1005", "Rian Hidayat", "9B", "Laki-laki", "https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=1005"]
  ];
  masterSheet.getRange(2, 1, sampleStudents.length, 5).setValues(sampleStudents);
  masterSheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#4F46E5").setFontColor("#FFFFFF");
  
  // 2. Sheet Log_Absensi
  let logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.SHEETS.LOG);
  }
  logSheet.clear();
  logSheet.appendRow(["Timestamp", "Tanggal", "NISN", "Nama_Siswa", "Kelas", "Tipe_Absen", "Jam_Absen", "Status_Kehadiran", "Petugas_Guru"]);
  logSheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#4F46E5").setFontColor("#FFFFFF");
  
  // 3. Sheet Rekap_Bulanan
  let rekapSheet = ss.getSheetByName(CONFIG.SHEETS.REKAP);
  if (!rekapSheet) {
    rekapSheet = ss.insertSheet(CONFIG.SHEETS.REKAP);
  }
  rekapSheet.clear();
  rekapSheet.appendRow(["NISN", "Nama_Siswa", "Kelas", "Total_Hadir", "Total_Terlambat", "Total_Izin", "Total_Alfa"]);
  rekapSheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#4F46E5").setFontColor("#FFFFFF");

  // Format perataan kolom
  [masterSheet, logSheet, rekapSheet].forEach(sh => {
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, 10);
  });
  
  updateRekapBulanan();
  return "Database Berhasil Diinisialisasi dengan Sample Data!";
}

/**
 * Autentikasi Login Admin / Guru
 */
function verifyAdminLogin(password) {
  if (password === CONFIG.ADMIN_PASSWORD) {
    return { success: true, message: "Login Berhasil" };
  } else {
    return { success: false, message: "Password Admin Salah!" };
  }
}

/**
 * PROSES ABSENSI (HADIR / PULANG)
 */
function processScanAbsen(nisn, tipeAbsen, petugasGuru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER);
  const logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
  
  nisn = String(nisn).trim();
  if (!nisn) {
    return { success: false, message: "NISN tidak boleh kosong!", code: "EMPTY_NISN" };
  }
  
  // Cari siswa di Master Data
  const masterData = masterSheet.getDataRange().getValues();
  let student = null;
  for (let i = 1; i < masterData.length; i++) {
    if (String(masterData[i][0]).trim() === nisn) {
      student = {
        nisn: String(masterData[i][0]),
        nama: masterData[i][1],
        kelas: masterData[i][2],
        jk: masterData[i][3],
        qrUrl: masterData[i][4]
      };
      break;
    }
  }
  
  if (!student) {
    return { 
      success: false, 
      message: "NISN [" + nisn + "] tidak terdaftar dalam database!", 
      code: "NOT_FOUND" 
    };
  }
  
  const now = new Date();
  const dateStr = Utilities.formatDate(now, CONFIG.TIMEZONE, "yyyy-MM-dd");
  const timeStr = Utilities.formatDate(now, CONFIG.TIMEZONE, "HH:mm:ss");
  const shortTime = Utilities.formatDate(now, CONFIG.TIMEZONE, "HH:mm");
  
  // Cek Log Absensi Hari Ini untuk NISN terkait
  const logData = logSheet.getDataRange().getValues();
  let sudahHadir = false;
  let sudahPulang = false;
  
  for (let j = 1; j < logData.length; j++) {
    const logDate = logData[j][1] instanceof Date ? 
      Utilities.formatDate(logData[j][1], CONFIG.TIMEZONE, "yyyy-MM-dd") : String(logData[j][1]);
    const logNisn = String(logData[j][2]).trim();
    const logTipe = logData[j][5];
    
    if (logDate === dateStr && logNisn === nisn) {
      if (logTipe === "Hadir") sudahHadir = true;
      if (logTipe === "Pulang") sudahPulang = true;
    }
  }
  
  let statusKehadiran = "";
  
  if (tipeAbsen === "Hadir") {
    if (sudahHadir) {
      return {
        success: false,
        message: student.nama + " (" + student.nisn + ") SUDAH melakukan absensi MASUK hari ini!",
        code: "DUPLICATE_HADIR",
        student: student
      };
    }
    
    // Tentukan Tepat Waktu / Terlambat
    if (shortTime <= CONFIG.TIME_LIMITS.HADIR_END) {
      statusKehadiran = "Tepat Waktu";
    } else {
      statusKehadiran = "Terlambat";
    }
    
  } else if (tipeAbsen === "Pulang") {
    if (!sudahHadir) {
      return {
        success: false,
        message: student.nama + " BELUM melakukan absensi MASUK hari ini!",
        code: "NO_HADIR_YET",
        student: student
      };
    }
    if (sudahPulang) {
      return {
        success: false,
        message: student.nama + " SUDAH melakukan absensi PULANG hari ini!",
        code: "DUPLICATE_PULANG",
        student: student
      };
    }
    statusKehadiran = "Pulang";
  }
  
  // Simpan ke Log_Absensi
  logSheet.appendRow([
    now,
    dateStr,
    student.nisn,
    student.nama,
    student.kelas,
    tipeAbsen,
    timeStr,
    statusKehadiran,
    petugasGuru || "Guru Piket"
  ]);
  
  // Update Rekap
  updateRekapBulanan();
  
  return {
    success: true,
    message: "Absensi " + tipeAbsen + " Berhasil recorded!",
    student: student,
    tipeAbsen: tipeAbsen,
    statusKehadiran: statusKehadiran,
    jamAbsen: timeStr,
    tanggal: dateStr
  };
}

/**
 * Mendapatkan 10 Log Scan Terbaru Hari Ini
 */
function getTodayLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
  const data = logSheet.getDataRange().getValues();
  
  if (data.length <= 1) return [];
  
  const now = new Date();
  const todayStr = Utilities.formatDate(now, CONFIG.TIMEZONE, "yyyy-MM-dd");
  
  const todayLogs = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const rowDate = data[i][1] instanceof Date ? 
      Utilities.formatDate(data[i][1], CONFIG.TIMEZONE, "yyyy-MM-dd") : String(data[i][1]);
      
    if (rowDate === todayStr) {
      todayLogs.push({
        timestamp: data[i][0],
        tanggal: rowDate,
        nisn: String(data[i][2]),
        nama: data[i][3],
        kelas: data[i][4],
        tipe: data[i][5],
        jam: data[i][6],
        status: data[i][7],
        petugas: data[i][8]
      });
    }
    if (todayLogs.length >= 10) break; // Ambil 10 data terakhir
  }
  return todayLogs;
}

/**
 * Cek Riwayat Kehadiran Siswa berdasarkan NISN (Mode Siswa/Wali)
 */
function getStudentHistory(nisn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER);
  const logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
  
  nisn = String(nisn).trim();
  const masterData = masterSheet.getDataRange().getValues();
  let studentInfo = null;
  
  for (let i = 1; i < masterData.length; i++) {
    if (String(masterData[i][0]).trim() === nisn) {
      studentInfo = {
        nisn: String(masterData[i][0]),
        nama: masterData[i][1],
        kelas: masterData[i][2],
        jk: masterData[i][3]
      };
      break;
    }
  }
  
  if (!studentInfo) {
    return { success: false, message: "NISN tidak ditemukan!" };
  }
  
  const logData = logSheet.getDataRange().getValues();
  const logs = [];
  for (let j = 1; j < logData.length; j++) {
    if (String(logData[j][2]).trim() === nisn) {
      const logDate = logData[j][1] instanceof Date ? 
        Utilities.formatDate(logData[j][1], CONFIG.TIMEZONE, "yyyy-MM-dd") : String(logData[j][1]);
      logs.push({
        tanggal: logDate,
        tipe: logData[j][5],
        jam: logData[j][6],
        status: logData[j][7]
      });
    }
  }
  
  return {
    success: true,
    student: studentInfo,
    history: logs.reverse()
  };
}

/**
 * Mendapatkan seluruh master siswa untuk Cetak Kartu QR Code / Barcode
 */
function getStudentBarcodes(kelasFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER);
  const data = masterSheet.getDataRange().getValues();
  
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const kelas = data[i][2];
    if (!kelasFilter || kelasFilter === "SEMUA" || kelas === kelasFilter) {
      result.push({
        nisn: String(data[i][0]),
        nama: data[i][1],
        kelas: data[i][2],
        jk: data[i][3],
        qrUrl: data[i][4] || ("https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=" + data[i][0])
      });
    }
  }
  return result;
}

/**
 * Mendapatkan Seluruh Log Absensi untuk Export/Rekap Dashboard Guru
 */
function getAllLogs(tanggalFilter, kelasFilter, tipeFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
  const data = logSheet.getDataRange().getValues();
  
  const logs = [];
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][1] instanceof Date ? 
      Utilities.formatDate(data[i][1], CONFIG.TIMEZONE, "yyyy-MM-dd") : String(data[i][1]);
    const kelas = data[i][4];
    const tipe = data[i][5];
    
    let match = true;
    if (tanggalFilter && rowDate !== tanggalFilter) match = false;
    if (kelasFilter && kelasFilter !== "SEMUA" && kelas !== kelasFilter) match = false;
    if (tipeFilter && tipeFilter !== "SEMUA" && tipe !== tipeFilter) match = false;
    
    if (match) {
      logs.push({
        tanggal: rowDate,
        nisn: String(data[i][2]),
        nama: data[i][3],
        kelas: data[i][4],
        tipe: data[i][5],
        jam: data[i][6],
        status: data[i][7],
        petugas: data[i][8]
      });
    }
  }
  return logs.reverse();
}

/**
 * Update Rekapitulasi Bulanan
 */
function updateRekapBulanan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER);
  const logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
  const rekapSheet = ss.getSheetByName(CONFIG.SHEETS.REKAP);
  
  const masterData = masterSheet.getDataRange().getValues();
  const logData = logSheet.getDataRange().getValues();
  
  rekapSheet.clear();
  rekapSheet.appendRow(["NISN", "Nama_Siswa", "Kelas", "Total_Hadir", "Total_Terlambat", "Total_Izin", "Total_Alfa"]);
  rekapSheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#4F46E5").setFontColor("#FFFFFF");
  
  const rekapMap = {};
  for (let i = 1; i < masterData.length; i++) {
    const nisn = String(masterData[i][0]).trim();
    rekapMap[nisn] = {
      nama: masterData[i][1],
      kelas: masterData[i][2],
      hadir: 0,
      terlambat: 0,
      izin: 0,
      alfa: 0
    };
  }
  
  for (let j = 1; j < logData.length; j++) {
    const nisn = String(logData[j][2]).trim();
    const tipe = logData[j][5];
    const status = logData[j][7];
    
    if (rekapMap[nisn] && tipe === "Hadir") {
      rekapMap[nisn].hadir++;
      if (status === "Terlambat") {
        rekapMap[nisn].terlambat++;
      }
    }
  }
  
  const rows = [];
  for (const nisn in rekapMap) {
    const r = rekapMap[nisn];
    rows.push([nisn, r.nama, r.kelas, r.hadir, r.terlambat, r.izin, r.alfa]);
  }
  
  if (rows.length > 0) {
    rekapSheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }
}