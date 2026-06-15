/**
 * Project: Logbook IKA (VERSI FINAL SINKRONISASI)
 * Update: Menambahkan pengambilan Dept Pemberi Kerja (J22) ke Kolom G.
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu("🔄 Update Manual")
    .addItem("Perbarui Logbook", "updateLogbookFromModifiedFiles")
    .addToUi();
}

function updateLogbookFromModifiedFiles(e) {
  const parentFolderId = "1Wqt9HGSKtn3DR4N1rJWwutR7GdjQDcdj";
  const logbook = SpreadsheetApp.getActiveSpreadsheet();
  const timezone = Session.getScriptTimeZone();
  const now = new Date();
  
  // SEMENTARA: Gunakan 210 (7 bulan) jika ingin pembersihan total, 
  // atau 30 untuk penggunaan rutin.
  const thresholdMs = 30 * 24 * 60 * 60 * 1000; 

  const currentYearNum = now.getFullYear();
  const currentMonth = now.getMonth(); 
  let totalUpdatedCount = 0;

  let yearsToProcess = [currentYearNum.toString()];
  if (currentMonth <= 2) { 
    yearsToProcess.push((currentYearNum - 1).toString());
  }

  const parentFolder = DriveApp.getFolderById(parentFolderId);

  function normalize(val) {
    if (val instanceof Date) return val.toISOString().slice(0, 19);
    if (val === null || val === undefined) return "";
    return String(val).trim();
  }

  yearsToProcess.forEach(year => {
    const logSheet = logbook.getSheetByName(year);
    const subfolders = parentFolder.getFoldersByName(year);
    if (!logSheet || !subfolders.hasNext()) return; 

    const folder = subfolders.next();
    const files = folder.getFiles();

    const data = logSheet.getDataRange().getValues();
    const linkToRowMap = new Map();
    for (let i = 8; i < data.length; i++) {
      const url = data[i][16]; // Kolom Q
      if (url) {
        const match = url.match(/[-\w]{25,}/);
        if (match) linkToRowMap.set(match[0], i + 1);
      }
    }

    while (files.hasNext()) {
      const file = files.next();
      if (now - file.getLastUpdated() > thresholdMs) continue;

      try {
        const ss = SpreadsheetApp.openById(file.getId());
        const sheet = ss.getSheetByName("IKA");
        const riwayatSheet = ss.getSheetByName("Riwayat Approval");
        if (!sheet) throw new Error("Sheet 'IKA' tidak ditemukan.");

        // --- 1. AMBIL DATA TEKNIS ---
        const nomorEWO = sheet.getRange("J20").getValue();
        const nomorJO = sheet.getRange("J21").getValue();
        
        // BARU: Mengambil Dept Pemberi Kerja dari J22 (Dropdown/Manual)
        const deptPemberiKerja = sheet.getRange("J22").getValue(); 

        const jenisCheckbox = sheet.getRange("D12:D16").getValues();
        const jenisLabels = sheet.getRange("E12:E16").getValues();
        const selectedIKAs = jenisCheckbox.map((row, idx) => row[0] ? jenisLabels[idx][0] : null).filter(v => v).join(", ");
        const departemenAreaKerja = sheet.getRange("J23").getValue();
        const deskripsi = sheet.getRange("J24").getValue();
        const jumlahPekerja = sheet.getRange("J27").getValue();
        const tglMulai = sheet.getRange("J30").getValue();
        const tglSelesai = sheet.getRange("J31").getValue();
        const emailPembuat = sheet.getRange("E96").getValue();

        // --- 2. TIPE PEKERJAAN (KOLOM L) ---
        const tipePekerjaan = sheet.getRange("J12").getValue() ? "Pekerjaan Baru"
                            : sheet.getRange("J13").getValue() ? "Perpanjangan" : "";

        // --- 3. STATUS IZIN (KOLOM O) + PENERJEMAH STATUS ---
        let statusIzinAsli = "";
        let skipStatusUpdate = false;

        if (riwayatSheet) {
          let rawStatus = riwayatSheet.getRange("I1").getValue();

          // Normalisasi "Approval In Progress"
          if (rawStatus === "Approval In Progress") {
            rawStatus = "Approval";
          }

          // Kalau masih Draft / kosong → jangan update
          if (!rawStatus || rawStatus === "Draft") {
            skipStatusUpdate = true;
          } else {
            statusIzinAsli = rawStatus;
          }
        }

        const row = linkToRowMap.get(file.getId());
        if (typeof row === "number") {
          const old = logSheet.getRange(row, 4, 1, 12).getValues()[0]; 
          // Status lama di logbook (kolom O)
          const statusLama = old[11];

          // Tentukan status final yang akan ditulis
          let statusFinal = statusLama;

          // Update hanya jika:
          // 1. bukan Draft
          // 2. status berbeda dari yang lama
          if (!skipStatusUpdate && statusIzinAsli !== statusLama) {
            statusFinal = statusIzinAsli;
          }

          const newVals = [
            nomorEWO,          
            nomorJO,           
            emailPembuat,      
            deptPemberiKerja,  
            departemenAreaKerja,
            selectedIKAs,      
            deskripsi,         
            jumlahPekerja,     
            tipePekerjaan,     
            tglMulai,          
            tglSelesai,        

            // Status sudah aman
            statusFinal        
          ];


          if (newVals.some((val, idx) => normalize(old[idx]) !== normalize(val))) {
            logSheet.getRange(row, 4, 1, 12).setValues([newVals]);
            logSheet.getRange(row, 1, 1, 15).setBackground(null);
            logSheet.getRange(row, 26).clearContent(); // Bersihkan error di Kolom Z
            totalUpdatedCount++;
          }
        }
      } catch (err) {
        let errorRow = linkToRowMap.get(file.getId());

        if (errorRow) {
          // Ambil sheet yang benar berdasarkan tahun yang sedang diproses
          const targetSheet = logbook.getSheetByName(year);

          if (targetSheet) {
            targetSheet.getRange(errorRow, 26)
              .setValue("⚠️ Error: " + err.message)
              .setFontColor("red");
          }
        }
      }
    }
  });

  // Memaksa pembaruan UI agar bar "Running Script" segera hilang
  SpreadsheetApp.flush();

  if (e === undefined) {
    SpreadsheetApp.getUi().alert("✅ Update Selesai. Total baris diperbarui: " + totalUpdatedCount);
  }
}

function createLogbookTrigger() {
  const allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "updateLogbookFromModifiedFiles") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("updateLogbookFromModifiedFiles")
    .timeBased()
    .everyDays(1)
    .atHour(0) 
    .create();

  Logger.log("✅ Trigger harian berhasil dibuat.");
}
