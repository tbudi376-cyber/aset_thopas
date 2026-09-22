// Objek untuk menyimpan konfigurasi.
const CONFIG = {
  SHEET_ID: "1vzRLW0e0fcB52BL3AezhrhF4c8MKKNS3bmCsWsbNsMg", 
  DATA_SHEET: "Data Aset Utama",
  USERS_SHEET: "Users",
  LOG_SHEET: "Log Aktivitas",
  REPORT_SHEET_BA: "BA Laporan Aset",
  LOAN_SHEET: "Peminjaman Aset"
};

// --- FUNGSI UTAMA: ROUTING HALAMAN ---
function doGet(e) {
  var idAset = e.parameter.id || ""; 
  var pageMode = e.parameter.page || ""; 

  var template = HtmlService.createTemplateFromFile('Index');
  template.idAset = idAset;

  if (idAset) {
    template.pageMode = 'detail';
  } else if (pageMode === 'general_loan') {
    template.pageMode = 'general_loan';
  } else {
    template.pageMode = 'admin_dashboard'; 
  }
  
  return template.evaluate()
      .setTitle("Sistem Aset Thopas")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- FUNGSI DASHBOARD ---
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.DATA_SHEET);
    const data = sheet.getDataRange().getValues();
    const header = data.shift(); 
    
    const idxStatus = header.indexOf("STATUS");
    
    let stats = { total: data.length, dipinjam: 0, rusak: 0, tersedia: 0 };

    data.forEach(row => {
      if (!row[0]) return; 
      let status = String(row[idxStatus] || "").toLowerCase();
      if (status.includes("pinjam")) stats.dipinjam++;
      else if (status.includes("rusak")) stats.rusak++;
      else stats.tersedia++;
    });

    return { success: true, stats: stats };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- FUNGSI PENCARIAN ASET ADVANCED ---
function searchAssets(keyword, filterAset) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.DATA_SHEET);
    const data = sheet.getDataRange().getValues();
    const header = data.shift(); 
    
    const idxID = header.indexOf("ID ASET");
    const idxAset = header.indexOf("ASET"); 
    const idxNama = header.indexOf("NAMA BARANG");
    const idxMerk = header.indexOf("MEREK & MODEL");
    const idxLokasi = header.indexOf("LOKASI SAAT INI");
    const idxStatus = header.indexOf("STATUS");
    
    if (idxID === -1 || idxNama === -1) return [];

    const terms = keyword ? keyword.toLowerCase().split(" ") : [];
    
    const results = data.filter(row => {
      if (!row[idxID]) return false;

      if (filterAset && filterAset !== "ALL") {
        const rowAset = String(row[idxAset] || "").trim().toUpperCase();
        if (rowAset !== filterAset) return false;
      }

      if (terms.length === 0) return true; 

      const textData = (
        String(row[idxID]) + " " + 
        String(row[idxAset]) + " " + 
        String(row[idxNama]) + " " + 
        String(row[idxMerk]) + " " + 
        String(row[idxLokasi])
      ).toLowerCase();

      return terms.every(term => textData.includes(term));
    });

    return results.slice(0, 50).map(row => {
      return {
        id: row[idxID],
        aset: row[idxAset], 
        nama: row[idxNama],
        merk: row[idxMerk],
        lokasi: row[idxLokasi],
        status: row[idxStatus]
      };
    });
    
  } catch (e) {
    return [{ error: e.message }];
  }
}

// --- FUNGSI LOGIN ---
function authenticateUser(credentials) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.USERS_SHEET);
    const data = sheet.getDataRange().getValues();
    const header = data.shift();
    const uIdx = header.indexOf("USERNAME");
    const pIdx = header.indexOf("PASSWORD");
    const rIdx = header.indexOf("ROLE");

    for (let row of data) {
      if (String(row[uIdx]) === String(credentials.username) && String(row[pIdx]) === String(credentials.password)) {
        return { success: true, role: row[rIdx] };
      }
    }
    return { success: false, message: "Username atau Password salah." };
  } catch (e) {
    return { success: false, message: "Error Login: " + e.message };
  }
}

// --- FUNGSI AMBIL DETAIL ---
function getAssetDetails(request) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.DATA_SHEET);
    const dataRange = sheet.getDataRange();
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const textFinder = dataRange.createTextFinder(request.idAset).matchEntireCell(true);
    const foundCell = textFinder.findNext();

    if (!foundCell) return { error: "Aset tidak ditemukan" };
    
    const rowNum = foundCell.getRow();
    const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    var assetDetails = {};
    const desiredHeaders = ["ID ASET", "NAMA BARANG", "KATEGORI", "MEREK & MODEL", "LOKASI SAAT INI", "PIC", "STATUS"];

    desiredHeaders.forEach(colName => {
      const index = header.indexOf(colName);
      if (index !== -1) {
        const val = rowData[index];
        assetDetails[colName] = (val instanceof Date) ? Utilities.formatDate(val, "Asia/Jakarta", "dd-MM-yyyy") : val;
      }
    });

    if(request.username) logActivity(request.username, request.idAset, "VIEW");
    return { success: true, data: assetDetails };
  } catch (e) {
    return { error: "Error Server: " + e.message };
  }
}

// --- FUNGSI UPDATE ---
function updateAssetData(updateData) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.DATA_SHEET);
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const textFinder = sheet.getDataRange().createTextFinder(updateData.idAset).matchEntireCell(true);
    const foundCell = textFinder.findNext();

    if (!foundCell) return { success: false, message: "Aset tidak ditemukan." };
    
    const colIndex = header.indexOf(updateData.field);
    if (colIndex === -1) return { success: false, message: "Kolom tidak valid." };

    sheet.getRange(foundCell.getRow(), colIndex + 1).setValue(updateData.newValue);
    logActivity(updateData.username, updateData.idAset, "EDIT " + updateData.field);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- FUNGSI PUBLIC DETAIL (UPDATED: RETURN ASET) ---
function getPublicAssetDetails(idAset) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.DATA_SHEET);
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const textFinder = sheet.getDataRange().createTextFinder(idAset).matchEntireCell(true);
    const foundCell = textFinder.findNext();

    if (!foundCell) return null;
    
    const rowData = sheet.getRange(foundCell.getRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
    const getVal = (name) => {
       const idx = header.indexOf(name);
       return idx > -1 ? rowData[idx] : '';
    };

    return {
      id: getVal("ID ASET"),
      aset: getVal("ASET"), // <--- UPDATE: Kirim Kode Aset (RBT, TAN, dll)
      nama_aset: getVal("NAMA BARANG"),
      merk_model: getVal("MEREK & MODEL"),
      no_seri: getVal("NOMOR SERI") || getVal("NO SERI"),
      lokasi_aset: getVal("LOKASI SAAT INI"),
      status_terkini: getVal("STATUS"), 
      pic_terkini: getVal("PIC")
    };
  } catch (e) {
    return null;
  }
}

// --- FUNGSI PEMINJAMAN ---
function handleLoanTransaction(form) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); 

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const loanSheet = ss.getSheetByName(CONFIG.LOAN_SHEET);
    const assetSheet = ss.getSheetByName(CONFIG.DATA_SHEET);
    const timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");

    const header = assetSheet.getRange(1, 1, 1, assetSheet.getLastColumn()).getValues()[0];
    const textFinder = assetSheet.getDataRange().createTextFinder(form.idAset).matchEntireCell(true);
    const foundCell = textFinder.findNext();

    if (!foundCell) return { success: false, message: "Aset tidak ditemukan di database." };
    const row = foundCell.getRow();
    
    const colStatus = header.indexOf("STATUS") + 1;
    const colPIC = header.indexOf("PIC") + 1;

    if (form.tipe === "PINJAM") {
      const currentStatus = assetSheet.getRange(row, colStatus).getValue();
      if (String(currentStatus).toLowerCase() === "dipinjam") {
        return { success: false, message: "Aset ini sedang status DIPINJAM. Harap kembalikan dahulu." };
      }

      loanSheet.appendRow([
        timestamp, form.idAset, form.peminjam, form.departemen, 
        "PINJAM", form.rencana_kembali, "", "Baik", form.keterangan, "OPEN"
      ]);

      assetSheet.getRange(row, colStatus).setValue("Dipinjam");
      assetSheet.getRange(row, colPIC).setValue(form.peminjam);
      logActivity(form.peminjam, form.idAset, "PEMINJAMAN");

    } else if (form.tipe === "KEMBALI") {
      assetSheet.getRange(row, colStatus).setValue(form.kondisi); 
      assetSheet.getRange(row, colPIC).setValue("Standby/Gudang"); 

      loanSheet.appendRow([
        timestamp, form.idAset, form.peminjam, form.departemen, 
        "KEMBALI", "", timestamp, form.kondisi, form.keterangan, "CLOSED"
      ]);
      
      logActivity(form.peminjam, form.idAset, "PENGEMBALIAN");
    }

    return { success: true };

  } catch (e) {
    return { success: false, message: "Error: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

function submitBAReport(form) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.REPORT_SHEET_BA);
    if (!sheet) return { success: false, message: "Sheet Laporan tidak ditemukan." };

    const timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([
      timestamp, form.idAset, form.nama_pelapor, form.id_karyawan, form.jabatan, form.departemen,
      form.nama_aset, form.merk_model, form.no_seri, form.lokasi_aset,
      form.jenis_laporan, form.tanggal_kejadian, form.waktu_kejadian,
      form.kronologi, form.deskripsi_kerusakan, form.tindakan_diambil,
      "Open", form.ttd_pelapor, form.ttd_atasan, form.ttd_ga
    ]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function logActivity(user, id, act) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.LOG_SHEET);
    sheet.appendRow([new Date(), user, id, act]);
  } catch(e){}
}
