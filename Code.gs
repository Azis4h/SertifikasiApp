/**
 * CODE.GS - UPDATED HEADER VERSION
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Sistem Informasi Sertifikasi')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

var SPREADSHEET_ID = '1Zd4plVMj7Z_UczDz8enSMYKI3AgD505AuuQdhNdPGqo'; 
var MAIN_SHEET_NAME_CAP = 'Perencanaan';
var MAIN_SHEET_NAME_LOWER = 'perencanaan';

function connect() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME_CAP) || ss.getSheetByName(MAIN_SHEET_NAME_LOWER);
  return sheet;
}

/* --- 1. DATA SERTIFIKASI (KIRI) --- */
function getData() {
  try {
    var sheet = connect();
    if (!sheet) return []; // Safety check
    
    var rows = sheet.getDataRange().getValues();
    var data = [];
    
    for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        // SKIP HEADER: Jika kolom A bernilai "NO" (case-insensitive)
        if (String(r[0]).toUpperCase() === "NO") continue;
        
        // VALIDASI: Skip jika baris dianggap kosong (tidak ada SAP dan Nama)
        // Kita tidak wajibkan kolom A (No) terisi, agar data tetap muncul meski user lupa isi nomor
        if (!r[1] && !r[2]) continue;

        try {
            var certItemId = r[3]; // Kolom D
            
            // Ambil data jika ada ID atau Nama
            if ((certItemId && String(certItemId).trim() !== "") || (r[1] && r[2])) {
                // Jika ID (r[0]) kosong, gunakan index loop sebagai fallback ID sementara
                var id = r[0] ? String(r[0]) : "ROW_" + i;

                data.push({
                    id: id,          
                    sap: cleanString(r[1]), 
                    nama: String(r[2]),        
                    itemId: String(r[3]),      
                    judul: String(r[4]),       
                    periode: safeParseDate(r[5]), 
                    jumlah: String(r[6]),           // JUMLAH ANGGARAN
                    statusAnggaran: String(r[7]),   // TERSEDIA/TIDAK
                    mandatory: String(r[8]),        // MANDATORY DAN REGULASI
                    resiko: String(r[9]),           // RESIKO
                    type: 'cert'
                });
            }
        } catch (rowErr) {
            Logger.log("Error processing CERT row " + i + ": " + rowErr);
        }
    }
    return data;
  } catch (e) {
    Logger.log('ERROR getData: ' + e.message);
    return []; // Return empty array to keep frontend running
  }
}

/* --- 2. DATA LAT (KANAN - Kolom L ke kanan) --- */
function getLATData() {
  try {
    var sheet = connect();
    if (!sheet) return [];

    var rows = sheet.getDataRange().getValues();
    var data = [];
    
    for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (String(r[0]).toUpperCase() === "NO") continue;
        if (!r[1] && !r[2]) continue; // Skip empty rows

        try {
            // Cek Kolom L (Index 11) - Item ID LAT
            var latItemId = r[11];
            
            // Allow entry if valid LAT item OR if basic data exists (SAP/Nama)
            if ((latItemId && String(latItemId).trim() !== "") || (r[1] && r[2] && r[11])) {
                 var id = r[0] ? String(r[0]) : "ROW_" + i;
                 
                 data.push({
                    id: id + "_LAT", 
                    originalId: id,
                    sap: cleanString(r[1]),
                    nama: String(r[2]),
                    itemId: String(r[11]),     
                    judul: String(r[12]),      
                    instruktur: String(r[13]), 
                    periode: safeParseDate(r[14]),
                    resiko: String(r[15]),     
                    type: 'lat'
                });
            }
        } catch (rowErr) {
             Logger.log("Error processing LAT row " + i + ": " + rowErr);
        }
    }
    return data;
  } catch (e) {
    Logger.log('ERROR getLATData: ' + e.message);
    return [];
  }
}

// HELPER
function cleanString(val) {
  if (!val) return "";
  return String(val).trim().toUpperCase(); 
}

// SAFE PARSE DATE - Handles Indonesian format and returns YYYY-MM-DD
function safeParseDate(dateVal) {
  try {
      if (!dateVal) return "";
      
      // 1. Jika object Date (dari Excel date cell)
      if (Object.prototype.toString.call(dateVal) === '[object Date]') {
        var yyyy = dateVal.getFullYear();
        var mm = String(dateVal.getMonth() + 1).padStart(2, '0');
        var dd = String(dateVal.getDate()).padStart(2, '0');
        return yyyy + "-" + mm + "-" + dd;
      }
      
      var str = String(dateVal).trim();

      // 2. Handle Format "Bulan Tahun" (Contoh: "Maret 2026")
      var monthMap = {
        'JANUARI': '01', 'FEBRUARI': '02', 'MARET': '03', 'APRIL': '04', 'MEI': '05', 'JUNI': '06',
        'JULI': '07', 'AGUSTUS': '08', 'SEPTEMBER': '09', 'OKTOBER': '10', 'NOVEMBER': '11', 'DESEMBER': '12',
        'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04', 'JUN': '06', 'JUL': '07', 'AGU': '08', 'SEP': '09', 'OKT': '10', 'NOV': '11', 'DES': '12'
      };
      
      // Cek apakah format "NamaBulan Tahun"
      var parts = str.split(' ');
      if (parts.length === 2) {
        var mName = parts[0].toUpperCase();
        var yName = parts[1];
        if (monthMap[mName] && !isNaN(yName)) {
           return yName + "-" + monthMap[mName] + "-01";
        }
      }
      
      // 3. Handle Format "D/M/YYYY" atau "M/D/YYYY" (Excel text format kadang begini)
      // Asumsi default Spreadsheet Indonesia: DD/MM/YYYY
      if (str.includes('/')) {
         var p = str.split('/');
         if (p.length === 3) {
            // Cek mana yang tahun (biasanya 4 digit)
            if (p[2].length === 4) return p[2] + "-" + String(p[1]).padStart(2,'0') + "-" + String(p[0]).padStart(2,'0');
            // Jika format english M/D/Y
            if (p[2].length === 2 && p[0].length === 4) return p[0] + "-" + String(p[1]).padStart(2,'0') + "-" + String(p[2]).padStart(2,'0'); 
         }
      }

      return str; 
  } catch (e) {
      return String(dateVal);
  }
}

function parseDate(d) { return safeParseDate(d); }

/* --- 3. CRUD (Update Mapping Save) --- */
/* --- 3. CRUD (DIPERBAIKI AGAR NOMOR BERURUTAN) --- */

// Helper untuk mendapatkan nomor urut selanjutnya
function getNextId(sheet) {
  var lastRow = sheet.getLastRow();
  
  // Jika baris hanya 1 (hanya header), mulai dari 1
  if (lastRow <= 1) return 1;

  // Ambil nilai dari kolom A baris terakhir
  var lastVal = sheet.getRange(lastRow, 1).getValue();

  // Pastikan nilainya angka, jika tidak (misal error), gunakan nomor baris
  var nextNum = parseInt(lastVal);
  if (isNaN(nextNum)) {
    return lastRow; // Fallback jika data berantakan
  }
  
  return nextNum + 1; // Nomor terakhir + 1
}

function addData(formObject) {
  var sheet = connect();
  
  // UBAH DISINI: Pakai getNextId bukan Date().getTime()
  var id = getNextId(sheet); 
  
  var newRow = [
      id, 
      formObject.sap, 
      formObject.nama,
      formObject.itemId, 
      formObject.judul, 
      formObject.periode, 
      formObject.jumlah,         
      formObject.statusAnggaran, 
      formObject.mandatory,      
      formObject.resiko,         
      "", "", "", "", "", "" 
  ];
  sheet.appendRow(newRow);
  return { success: true };
}

function addLATData(formObject) {
    var sheet = connect();
    
    // UBAH DISINI JUGA
    var id = getNextId(sheet);

    var newRow = [
        id, formObject.sap, formObject.nama,
        "", "", "", "", "", "", "", 
        "", 
        formObject.itemId, formObject.judul, formObject.instruktur, 
        formObject.periode, formObject.resiko
    ];
    sheet.appendRow(newRow);
    return { success: true };
}

/* --- UPDATE DATA SERTIFIKASI --- */
function updateData(formData) {
  try {
    var sheet = connect();
    if (!sheet) return { success: false, error: 'Sheet Perencanaan tidak ditemukan' };

    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(formData.id)) {
        rowIndex = i + 1; // 1-indexed untuk getRange
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'Data tidak ditemukan dengan ID: ' + formData.id };

    // Update kolom Cert: B(SAP), C(Nama), D(ItemId), E(Judul), F(Periode), G(Jumlah), H(StatusAnggaran), I(Mandatory), J(Resiko)
    sheet.getRange(rowIndex, 2).setValue(formData.sap || '');
    sheet.getRange(rowIndex, 3).setValue(formData.nama || '');
    sheet.getRange(rowIndex, 4).setValue(formData.itemId || '');
    sheet.getRange(rowIndex, 5).setValue(formData.judul || '');
    sheet.getRange(rowIndex, 6).setValue(formData.periode || '');
    sheet.getRange(rowIndex, 7).setValue(formData.jumlah || '');
    sheet.getRange(rowIndex, 8).setValue(formData.statusAnggaran || '');
    sheet.getRange(rowIndex, 9).setValue(formData.mandatory || '');
    sheet.getRange(rowIndex, 10).setValue(formData.resiko || '');

    return { success: true };
  } catch (e) {
    Logger.log('Error in updateData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/* --- UPDATE DATA LAT --- */
function updateLATData(formData) {
  try {
    var sheet = connect();
    if (!sheet) return { success: false, error: 'Sheet Perencanaan tidak ditemukan' };

    // ID LAT format: "X_LAT" — ambil original ID dengan hapus "_LAT"
    var originalId = String(formData.id).replace('_LAT', '');

    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === originalId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'Data LAT tidak ditemukan dengan ID: ' + originalId };

    // Update kolom LAT: B(SAP), C(Nama), L(ItemId), M(Judul), N(Instruktur), O(Periode), P(Resiko)
    sheet.getRange(rowIndex, 2).setValue(formData.sap || '');
    sheet.getRange(rowIndex, 3).setValue(formData.nama || '');
    sheet.getRange(rowIndex, 12).setValue(formData.itemId || '');
    sheet.getRange(rowIndex, 13).setValue(formData.judul || '');
    sheet.getRange(rowIndex, 14).setValue(formData.instruktur || '');
    sheet.getRange(rowIndex, 15).setValue(formData.periode || '');
    sheet.getRange(rowIndex, 16).setValue(formData.resiko || '');

    return { success: true };
  } catch (e) {
    Logger.log('Error in updateLATData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/* --- DELETE DATA SERTIFIKASI --- */
function deleteData(id) {
  try {
    var sheet = connect();
    if (!sheet) return { success: false, error: 'Sheet Perencanaan tidak ditemukan' };

    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'Data tidak ditemukan dengan ID: ' + id };

    // Cek apakah baris ini juga punya data LAT (kolom L / index 11)
    var hasLAT = rows[rowIndex - 1][11] && String(rows[rowIndex - 1][11]).trim() !== '';

    if (hasLAT) {
      // Baris punya LAT juga — hanya kosongkan kolom Cert (D-J) agar data LAT aman
      sheet.getRange(rowIndex, 4, 1, 7).clearContent(); // D=4 sampai J=10 (7 kolom)
    } else {
      // Baris hanya Cert — hapus seluruh baris
      sheet.deleteRow(rowIndex);
    }

    return { success: true };
  } catch (e) {
    Logger.log('Error in deleteData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/* --- DELETE DATA LAT --- */
function deleteLATData(id) {
  try {
    var sheet = connect();
    if (!sheet) return { success: false, error: 'Sheet Perencanaan tidak ditemukan' };

    // ID LAT format: "X_LAT"
    var originalId = String(id).replace('_LAT', '');

    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === originalId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'Data LAT tidak ditemukan dengan ID: ' + originalId };

    // Cek apakah baris ini juga punya data Cert (kolom D / index 3)
    var hasCert = rows[rowIndex - 1][3] && String(rows[rowIndex - 1][3]).trim() !== '';

    if (hasCert) {
      // Baris punya Cert juga — hanya kosongkan kolom LAT (L-P) agar data Cert aman
      sheet.getRange(rowIndex, 12, 1, 5).clearContent(); // L=12 sampai P=16 (5 kolom)
    } else {
      // Baris hanya LAT — hapus seluruh baris
      sheet.deleteRow(rowIndex);
    }

    return { success: true };
  } catch (e) {
    Logger.log('Error in deleteLATData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/* --- DATA PELAKSANAAN (UPDATED: SESUAI USER HEADERS) --- */
/**
 * Membaca data dari sheet Pelaksanaan.
 * Kolom: NO, SAP, Start, End, Bulan, Tahun, Item ID, Sap Instruktur, Nama Instruktur, 
 * Course Title, SAP, Nama Partisipan, Room, Pesona, Kel, Departemen, Unit Kerja, 
 * Jumlah Hadir, Count Pelatihan, Durasi, Kehadiran, Durasi Peserta, Durasi Instruktur
 */
function getRealizationData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Pelaksanaan') || ss.getSheetByName('pelaksanaan'); 
    
    if (!sheet) {
      Logger.log('Sheet Pelaksanaan/pelaksanaan tidak ditemukan');
      return [];
    }

    var rows = sheet.getDataRange().getValues();
    var data = [];
    
    // Skip header row, mulai dari index 1
    for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        
        try {
            // Ambil NO dari kolom A (index 0)
            var no = r[0] ? String(r[0]) : "";
            
            // Skip jika ini adalah header "NO" (just in case)
            if (no.trim().toUpperCase() === "NO") continue;

            // VALIDASI: Skip jika baris dianggap kosong total (Cek SAP, Nama, Judul, SAP Peserta, Nama Peserta)
            // r[1]=SAP Event, r[9]=Course Title, r[6]=Item ID, r[10]=SAP Peserta, r[11]=Nama Peserta
            if (!r[1] && !r[9] && !r[6] && !r[10] && !r[11]) continue;
            
            // Fallback ID jika kosong
            if (!no || no.trim() === "") no = "REAL_" + i;

            // Safe Date Parsing
            var dateStart = safeParseDate(r[2]);
            var dateEnd = safeParseDate(r[3]);
            
            // Robust Year Extraction
            // 1. Cek kolom Tahun (Index 5)
            var rawTahun = r[5];
            var tahun = "";
            
            if (rawTahun) {
               if (Object.prototype.toString.call(rawTahun) === '[object Date]') {
                  tahun = String(rawTahun.getFullYear());
               } else {
                  var strTahun = String(rawTahun).trim();
                  var match = strTahun.match(/20\d{2}/);
                  if (match) tahun = match[0];
                  else tahun = strTahun;
               }
            }
            
            
            // 2. Fallback: Jika kolom Tahun kosong, ambil dari Start Date (Index 2)
            if ((!tahun || tahun === "") && dateStart) {
                var d = new Date(dateStart);
                if (!isNaN(d.getTime())) {
                    tahun = String(d.getFullYear());
                }
            }

            // 3. Fallback: Ambil dari End Date (Index 3)
            if ((!tahun || tahun === "") && dateEnd) {
                 var d = new Date(dateEnd);
                 if (!isNaN(d.getTime())) {
                     tahun = String(d.getFullYear());
                 }
            }

            // 4. Last Resort: "Uncategorized" atau empty string (biar masuk card "Semua Data")
            if (!tahun) tahun = ""; 

            // --- CRITICAL FIX FOR FRONTEND GROUPING ---
            // Frontend groups by SAP. If Participant SAP (Col K / Index 10) is missing, 
            // we MUST provide a fallback, otherwise it might be grouped under "undefined" or lost.
            
            // PRIORITAS SAP: 1. SAP Peserta (K) -> 2. SAP Event (B) -> 3. "NO_SAP"
            var finalSap = r[10] ? String(r[10]) : (r[1] ? String(r[1]) : "NO_SAP");
            
            // PRIORITAS NAMA: 1. Nama Peserta (L) -> 2. Course Title (J) -> 3. Nama Instruktur -> 4. "No Name"
            var finalNama = r[11] ? String(r[11]) : (r[9] ? String(r[9]) : (r[8] ? String(r[8]) : "No Name"));

            data.push({
                id: no,                                         
                sapEvent: r[1] ? String(r[1]) : "",            
                sapStart: dateStart,         
                end: dateEnd,              
                bulan: r[4] ? String(r[4]) : "",               
                tahun: tahun,                                   
                itemId: r[6] ? String(r[6]) : "",              
                sapInstruktur: r[7] ? String(r[7]) : "",       
                namaInstruktur: r[8] ? String(r[8]) : "",      
                courseTitle: r[9] ? String(r[9]) : "",
                judulPelatihan: r[9] ? String(r[9]) : "", // Alias for frontend compatibility         
                sapPeserta: r[10] ? String(r[10]) : "",        
                namaPeserta: r[11] ? String(r[11]) : "",       
                room: r[12] ? String(r[12]) : "",              
                
                // MAPPING BARU SESUAI GAMBAR USER
                pesona: r[13] ? String(r[13]) : "",          // Ex: Presensi
                kel: r[14] ? String(r[14]) : "",             // Ex: Ket
                
                departemen: r[15] ? String(r[15]) : "",        
                unitKerja: r[16] ? String(r[16]) : "",         
                jumlahHadir: r[17] != null ? String(r[17]) : "",       
                countPelatihan: r[18] != null ? String(r[18]) : "",    
                durasi: r[19] != null ? String(r[19]) : "",            
                kehadiran: r[20] != null ? String(r[20]) : "",         
                durasiPeserta: r[21] != null ? String(r[21]) : "",   
                durasiInstruktur: r[22] != null ? String(r[22]) : "",    
                
                // Compatibility Fields (untuk frontend existing agar tidak error)
                sap: finalSap,       // CRITICAL: Used for grouping in renderRealizationList
                nama: finalNama      // CRITICAL: Used for grouping name
            });
        } catch (errRow) {
            Logger.log("Error processing row " + i + ": " + errRow.message);
        }
    }
    return data;
  } catch (e) {
    Logger.log('ERROR getRealizationData: ' + e.message);
    return []; 
  }
}

/* ─────────────────────────────────────────────────────────────────────────────
 * REALIZATION DATA CRUD OPERATIONS
 * ───────────────────────────────────────────────────────────────────────────── */

/**
 * Add new realization data to Pelaksanaan sheet
 */
function addRealizationData(formData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Pelaksanaan') || ss.getSheetByName('pelaksanaan');
    
    if (!sheet) {
      return { success: false, error: 'Sheet Pelaksanaan not found' };
    }

    // Generate new ID
    var nextId = getNextId(sheet);
    
    // Map formData to row array matching column order
    var newRow = [
      nextId,                            // A: NO
      formData.sapEvent || '',          // B: SAP Event
      formData.sapStart || '',          // C: Start Date
      formData.end || '',               // D: End Date
      formData.bulan || '',             // E: Bulan
      formData.tahun || '',             // F: Tahun
      formData.itemId || '',            // G: Item ID
      formData.sapInstruktur || '',     // H: SAP Instruktur
      formData.namaInstruktur || '',    // I: Nama Instruktur
      formData.courseTitle || '',       // J: Course Title
      formData.sapPeserta || '',        // K: SAP Peserta
      formData.namaPeserta || '',       // L: Nama Peserta
      formData.room || '',              // M: Room
      '',                                // N: Pesona (not in form)
      '',                                // O: Kel (not in form)
      formData.departemen || '',        // P: Departemen
      formData.unitKerja || '',         // Q: Unit Kerja
      formData.jumlahHadir || '',       // R: Jumlah Hadir
      '',                                // S: Count Pelatihan (not in form)
      formData.durasi || '',            // T: Durasi
      formData.kehadiran || '',         // U: Kehadiran
      '',                                // V: Durasi Peserta (not in form)
      ''                                 // W: Durasi Instruktur (not in form)
    ];
    
    sheet.appendRow(newRow);
    
    // Return updated data
    var updatedData = getRealizationData();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in addRealizationData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Update existing realization data
 */
function updateRealizationData(formData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Pelaksanaan') || ss.getSheetByName('pelaksanaan');
    
    if (!sheet) {
      return { success: false, error: 'Sheet Pelaksanaan not found' };
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    // Find row by ID (column A)
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(formData.id)) {
        rowIndex = i + 1; // Row number (1-indexed)
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: 'Data not found with ID: ' + formData.id };
    }
    
    // Update row with new data
    sheet.getRange(rowIndex, 2).setValue(formData.sapEvent || '');      // B
    sheet.getRange(rowIndex, 3).setValue(formData.sapStart || '');      // C
    sheet.getRange(rowIndex, 4).setValue(formData.end || '');           // D
    sheet.getRange(rowIndex, 5).setValue(formData.bulan || '');         // E
    sheet.getRange(rowIndex, 6).setValue(formData.tahun || '');         // F
    sheet.getRange(rowIndex, 7).setValue(formData.itemId || '');        // G
    sheet.getRange(rowIndex, 8).setValue(formData.sapInstruktur || ''); // H
    sheet.getRange(rowIndex, 9).setValue(formData.namaInstruktur || ''); // I
    sheet.getRange(rowIndex, 10).setValue(formData.courseTitle || '');  // J
    sheet.getRange(rowIndex, 11).setValue(formData.sapPeserta || '');   // K
    sheet.getRange(rowIndex, 12).setValue(formData.namaPeserta || '');  // L
    sheet.getRange(rowIndex, 13).setValue(formData.room || '');         // M
    sheet.getRange(rowIndex, 16).setValue(formData.departemen || '');   // P
    sheet.getRange(rowIndex, 17).setValue(formData.unitKerja || '');    // Q
    sheet.getRange(rowIndex, 18).setValue(formData.jumlahHadir || '');  // R
    sheet.getRange(rowIndex, 20).setValue(formData.durasi || '');       // T
    sheet.getRange(rowIndex, 21).setValue(formData.kehadiran || '');    // U
    
    // Return updated data
    var updatedData = getRealizationData();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in updateRealizationData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Delete realization data by ID
 */
function deleteRealizationData(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Pelaksanaan') || ss.getSheetByName('pelaksanaan');
    
    if (!sheet) {
      return { success: false, error: 'Sheet Pelaksanaan not found' };
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    // Find row by ID (column A)
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        rowIndex = i + 1; // Row number (1-indexed)
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: 'Data not found with ID: ' + id };
    }
    
    sheet.deleteRow(rowIndex);
    
    // Return updated data
    var updatedData = getRealizationData();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in deleteRealizationData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/* ─────────────────────────────────────────────────────────────────────────────
 * EVALUASI L1 DATA OPERATIONS
 * ───────────────────────────────────────────────────────────────────────────── */

/**
 * Get all L1 evaluation data from sheet "L1"
 */
function getEvaluasiL1Data() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L1') || ss.getSheetByName('l1');
    
    if (!sheet) {
      Logger.log('Sheet L1 not found');
      return [];
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    if (values.length <= 1) {
      Logger.log('No data in L1 sheet');
      return [];
    }

    var data = [];
    
    // Start from row 2 (skip header)
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      
      // Skip empty rows
      if (!r[0] && !r[1]) continue;
      
        data.push({
        id: r[0] ? String(r[0]) : '',                           
        judulPelatihan: r[1] ? String(r[1]) : '',               
        pelaksanaanId: r[2] ? safeParseDate(r[2]) : '',               
        sap: r[3] ? String(r[3]) : '',                          
        namaPeserta: r[4] ? String(r[4]) : '',                  
        tempatPembelajaran: r[5] ? String(r[5]) : '',           
        fasilitasMedia: r[6] ? String(r[6]) : '',               
        pelayananUmum: r[7] ? String(r[7]) : '',                
        ratapenyelenggaraan: r[8] ? String(r[8]) : '',         
        materi: r[9] ? String(r[9]) : '',                       
        tujuanTercapai: r[10] ? String(r[10]) : '',              
        penyajian: r[11] ? String(r[11]) : '',                   
        disiplin: r[12] ? String(r[12]) : '',                    
        rataPembelajaran: r[13] ? String(r[13]) : '',            
        pengetahuan: r[14] ? String(r[14]) : '',                 
        presentasi: r[15] ? String(r[15]) : '',                  
        perilaku: r[16] ? String(r[16]) : '',                    
        waktu: r[17] ? String(r[17]) : '',                       
        rataInstruktur: r[18] ? String(r[18]) : '',              
        rataKeseluruhan: r[19] ? String(r[19]) : '',             
        komentarPeserta: r[20] ? String(r[20]) : ''              
      });
    }
    
    Logger.log('L1 data loaded: ' + data.length + ' records');
    return data;
    
  } catch (e) {
    Logger.log('ERROR getEvaluasiL1Data: ' + e.message);
    return [];
  }
}

/**
 * Add new L1 evaluation
 */
function addEvaluasiL1(formData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L1') || ss.getSheetByName('l1');
    
    if (!sheet) {
      return { success: false, error: 'Sheet L1 not found' };
    }

    var nextId = getNextId(sheet);
    
    var newRow = [
      nextId,                                     // A
      formData.judulPelatihan || '',              // B
      formData.pelaksanaanId || '',               // C
      formData.sap || '',                         // D
      formData.namaPeserta || '',                 // E
      formData.tempatPembelajaran || '',          // F
      formData.fasilitasMedia || '',              // G
      formData.pelayananUmum || '',               // H
      formData.ratapenyelenggaraan || '',         // I (Manual Input)
      formData.materi || '',                      // J
      formData.tujuanTercapai || '',              // K
      formData.penyajian || '',                   // L
      formData.disiplin || '',                    // M
      formData.rataPembelajaran || '',            // N (Manual Input)
      formData.pengetahuan || '',                 // O
      formData.presentasi || '',                  // P
      formData.perilaku || '',                    // Q
      formData.waktu || '',                       // R
      formData.rataInstruktur || '',              // S (Manual Input)
      formData.rataKeseluruhan || '',             // T (Manual Input)
      formData.komentarPeserta || ''              // U
    ];
    
    sheet.appendRow(newRow);
    
    var updatedData = getEvaluasiL1Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in addEvaluasiL1: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Update L1 evaluation
 */
function updateEvaluasiL1(formData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L1') || ss.getSheetByName('l1');
    
    if (!sheet) {
      return { success: false, error: 'Sheet L1 not found' };
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(formData.id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: 'Data not found' };
    }
    
    // Update all fields
    sheet.getRange(rowIndex, 2).setValue(formData.judulPelatihan || '');
    sheet.getRange(rowIndex, 3).setValue(formData.pelaksanaanId || '');
    sheet.getRange(rowIndex, 4).setValue(formData.sap || '');
    sheet.getRange(rowIndex, 5).setValue(formData.namaPeserta || '');
    sheet.getRange(rowIndex, 6).setValue(formData.tempatPembelajaran || '');
    sheet.getRange(rowIndex, 7).setValue(formData.fasilitasMedia || '');
    sheet.getRange(rowIndex, 8).setValue(formData.pelayananUmum || '');
    sheet.getRange(rowIndex, 9).setValue(formData.ratapenyelenggaraan || ''); // Manual Input
    sheet.getRange(rowIndex, 10).setValue(formData.materi || '');
    sheet.getRange(rowIndex, 11).setValue(formData.tujuanTercapai || '');
    sheet.getRange(rowIndex, 12).setValue(formData.penyajian || '');
    sheet.getRange(rowIndex, 13).setValue(formData.disiplin || '');
    sheet.getRange(rowIndex, 14).setValue(formData.rataPembelajaran || ''); // Manual Input
    sheet.getRange(rowIndex, 15).setValue(formData.pengetahuan || '');
    sheet.getRange(rowIndex, 16).setValue(formData.presentasi || '');
    sheet.getRange(rowIndex, 17).setValue(formData.perilaku || '');
    sheet.getRange(rowIndex, 18).setValue(formData.waktu || '');
    sheet.getRange(rowIndex, 19).setValue(formData.rataInstruktur || ''); // Manual Input
    sheet.getRange(rowIndex, 20).setValue(formData.rataKeseluruhan || ''); // Manual Input
    sheet.getRange(rowIndex, 21).setValue(formData.komentarPeserta || '');
    
    var updatedData = getEvaluasiL1Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in updateEvaluasiL1: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Delete L1 evaluation
 */
function deleteEvaluasiL1(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L1') || ss.getSheetByName('l1');
    
    if (!sheet) {
      return { success: false, error: 'Sheet L1 not found' };
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: 'Data not found' };
    }
    
    sheet.deleteRow(rowIndex);
    
    var updatedData = getEvaluasiL1Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in deleteEvaluasiL1: ' + e.message);
    return { success: false, error: e.message };
  }
}

/** 
 * =================================================================================
 * EVALUASI L2 (LEARNING) - CRUD
 * =================================================================================
 */

function getEvaluasiL2Data() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L2');
    
    if (!sheet) {
      // Auto-create if not exists
      sheet = ss.insertSheet('L2');
      sheet.appendRow([
        'No', 'Judul Pelatihan', 'Pelaksanaan', 'SAP', 'Nama Peserta', 
        'Pre Test', 'Post Test', 'Increase', 'Ket.'
      ]);
      return [];
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    if (values.length <= 1) return [];

    var data = [];
    
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      if (!r[0] && !r[1] && !r[4]) continue; // Check ID, Judul, or Nama
      
      data.push({
        id: r[0] ? String(r[0]) : '',
        judulPelatihan: r[1] ? String(r[1]) : '',
        pelaksanaanId: r[2] ? safeParseDate(r[2]) : '',
        sap: r[3] ? String(r[3]) : '',
        namaPeserta: r[4] ? String(r[4]) : '',
        preTest: r[5] ? String(r[5]) : '0',
        postTest: r[6] ? String(r[6]) : '0',
        increase: r[7] ? String(r[7]) : '0',       
        ket: r[8] ? String(r[8]) : ''            
      });
    }
    
    return data;
    
  } catch (e) {
    Logger.log('ERROR getEvaluasiL2Data: ' + e.message);
    return [];
  }
}

function addEvaluasiL2(formData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L2');
    
    if (!sheet) {
      sheet = ss.insertSheet('L2');
      sheet.appendRow([
        'No', 'Judul Pelatihan', 'Pelaksanaan', 'SAP', 'Nama Peserta', 
        'Pre Test', 'Post Test', 'Increase', 'Ket.'
      ]);
    }

    var nextId = getNextId(sheet);
    var increase = (parseFloat(formData.postTest) || 0) - (parseFloat(formData.preTest) || 0);
    
    var newRow = [
      nextId,
      formData.judulPelatihan || '',
      formData.pelaksanaanId || '',
      formData.sap || '',
      formData.namaPeserta || '',
      formData.preTest || 0,
      formData.postTest || 0,
      increase.toFixed(2),
      formData.ket || ''
    ];
    
    sheet.appendRow(newRow);
    
    var updatedData = getEvaluasiL2Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in addEvaluasiL2: ' + e.message);
    return { success: false, error: e.message };
  }
}

function updateEvaluasiL2(formData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L2');
    if (!sheet) return { success: false, error: 'Sheet L2 not found' };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(formData.id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, error: 'Data not found' };
    
    var increase = (parseFloat(formData.postTest) || 0) - (parseFloat(formData.preTest) || 0);

    // Update fields (Columns 2-9)
    sheet.getRange(rowIndex, 2).setValue(formData.judulPelatihan || '');
    sheet.getRange(rowIndex, 3).setValue(formData.pelaksanaanId || '');
    sheet.getRange(rowIndex, 4).setValue(formData.sap || '');
    sheet.getRange(rowIndex, 5).setValue(formData.namaPeserta || '');
    sheet.getRange(rowIndex, 6).setValue(formData.preTest || 0);
    sheet.getRange(rowIndex, 7).setValue(formData.postTest || 0);
    sheet.getRange(rowIndex, 8).setValue(increase.toFixed(2));
    sheet.getRange(rowIndex, 9).setValue(formData.ket || '');
    
    var updatedData = getEvaluasiL2Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in updateEvaluasiL2: ' + e.message);
    return { success: false, error: e.message };
  }
}

function deleteEvaluasiL2(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L2');
    if (!sheet) return { success: false, error: 'Sheet L2 not found' };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, error: 'Data not found' };
    
    sheet.deleteRow(rowIndex);
    
    var updatedData = getEvaluasiL2Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in deleteEvaluasiL2: ' + e.message);
    return { success: false, error: e.message };
  }
}

/** 
 * =================================================================================
 * EVALUASI L3 (BEHAVIOR) - CRUD
 *Headers: No, Judul Pelatihan, Pelaksanaan Learning, SAP, Nama Peserta, Nilai Evaluasi, Ket., Key Behaviour, Tanggal Eval
 * =================================================================================
 */

function getEvaluasiL3Data() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L3') || ss.getSheetByName('l3');
    
    if (!sheet) {
      sheet = ss.insertSheet('L3');
      sheet.appendRow([
        'No', 'Judul Pelatihan', 'Pelaksanaan Learning', 'SAP', 'Nama Peserta', 
        'Nilai Evaluasi', 'Ket.', 'Key Behaviour', 'Tanggal Eval'
      ]);
      return [];
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    if (values.length <= 1) return [];

    var data = [];
    
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      if (!r[0] && !r[1] && !r[4]) continue;
      
      data.push({
        id: r[0] ? String(r[0]) : '',
        judulPelatihan: r[1] ? String(r[1]) : '',
        pelaksanaanId: r[2] ? String(r[2]) : '',
        sap: r[3] ? String(r[3]) : '',
        namaPeserta: r[4] ? String(r[4]) : '',
        nilaiEvaluasi: r[5] ? String(r[5]) : '',
        ket: r[6] ? String(r[6]) : '',
        keyBehaviour: r[7] ? String(r[7]) : '',
        tanggalEval: r[8] ? Utilities.formatDate(new Date(r[8]), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : ''
      });
    }
    
    return data;
    
  } catch (e) {
    Logger.log('ERROR getEvaluasiL3Data: ' + e.message);
    return [];
  }
}

function addEvaluasiL3(formData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L3') || ss.getSheetByName('l3');
    
    if (!sheet) {
      sheet = ss.insertSheet('L3');
      sheet.appendRow([
        'No', 'Judul Pelatihan', 'Pelaksanaan Learning', 'SAP', 'Nama Peserta', 
        'Nilai Evaluasi', 'Ket.', 'Key Behaviour', 'Tanggal Eval'
      ]);
    }

    var nextId = getNextId(sheet);
    
    var newRow = [
      nextId,
      formData.judulPelatihan || '',
      formData.pelaksanaanId || '',
      formData.sap || '',
      formData.namaPeserta || '',
      formData.nilaiEvaluasi || '',
      formData.ket || '',
      formData.keyBehaviour || '',
      formData.tanggalEval || ''
    ];
    
    sheet.appendRow(newRow);
    
    var updatedData = getEvaluasiL3Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in addEvaluasiL3: ' + e.message);
    return { success: false, error: e.message };
  }
}

function updateEvaluasiL3(formData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L3') || ss.getSheetByName('l3');
    if (!sheet) return { success: false, error: 'Sheet L3 not found' };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(formData.id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, error: 'Data not found' };
    
    // Update fields (Columns 2-9)
    sheet.getRange(rowIndex, 2).setValue(formData.judulPelatihan || '');
    sheet.getRange(rowIndex, 3).setValue(formData.pelaksanaanId || '');
    sheet.getRange(rowIndex, 4).setValue(formData.sap || '');
    sheet.getRange(rowIndex, 5).setValue(formData.namaPeserta || '');
    sheet.getRange(rowIndex, 6).setValue(formData.nilaiEvaluasi || '');
    sheet.getRange(rowIndex, 7).setValue(formData.ket || '');
    sheet.getRange(rowIndex, 8).setValue(formData.keyBehaviour || '');
    sheet.getRange(rowIndex, 9).setValue(formData.tanggalEval || '');
    
    var updatedData = getEvaluasiL3Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in updateEvaluasiL3: ' + e.message);
    return { success: false, error: e.message };
  }
}

function deleteEvaluasiL3(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L3') || ss.getSheetByName('l3');
    if (!sheet) return { success: false, error: 'Sheet L3 not found' };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, error: 'Data not found' };
    
    sheet.deleteRow(rowIndex);
    
    var updatedData = getEvaluasiL3Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in deleteEvaluasiL3: ' + e.message);
    return { success: false, error: e.message };
  }
}
