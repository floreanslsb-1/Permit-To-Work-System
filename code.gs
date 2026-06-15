// === KONFIGURASI ===
const TEMPLATE_FILE_ID = "17i79FlZIFw8iEh2xBWKzPt4PxjYDQEhardtgydyRvDQ";
const TARGET_FOLDER_ID = "1Wqt9HGSKtn3DR4N1rJWwutR7GdjQDcdj";
const LOGBOOK_FILE_ID = "196wV40IgrLAVIssGAZXASyuYV3NqOU9w5CwDsNBO-wk";
const WEBAPP_URL = "https://script.google.com/a/macros/wingscorp.com/s/AKfycbxwQZ9MRp5rVDSj2gE5OIAHfz1Gs3y3VabjAar76QydRUfbFJqr-8CggMCH_G-A_4PW/exec"
// PENTING: Ganti URL ini dengan URL Web App Anda setelah melakukan deployment
const SCRIPT_URL = "https://script.google.com/a/macros/wingscorp.com/s/AKfycbxwQZ9MRp5rVDSj2gE5OIAHfz1Gs3y3VabjAar76QydRUfbFJqr-8CggMCH_G-A_4PW/exec"; 
const WINGSGROUP_DOMAIN = "wingscorp.com"; 
const ADMIN_EMAIL = "floreansalsabila.irdana@wingscorp.com";

// === KONFIGURASI BARU UNTUK REMINDER (V12 - Sesuai Permintaan) ===
const MASTER_TIMESTAMP_COL = 21; // Kolom U (Timer Bersama)
const MASTER_STATE_COL = 25;     // Kolom Y (State Bersama: 1, 2, 3, "Expired", "Closing")

// Checklist Approval (V, W, X)
const APPROVAL_H1_COL = 22; // Kolom V
const APPROVAL_H2_COL = 23; // Kolom W
const APPROVAL_H3_COL = 24; // Kolom X

// Checklist Closing (R, S, T) - (H+1, H+3, H+7)
const CLOSING_H1_COL = 18; // Kolom R
const CLOSING_H3_COL = 19; // Kolom S
const CLOSING_H7_COL = 20; // Kolom T
// ====================================================

// --- LOKASI SEL PENTING DI SHEET IKA ---
const CELL_PENGAJU_EMAIL = "E96";
const CELL_PEMBERI_KERJA_L2 = "E99";
const RANGE_APPROVER_L2 = "E99:E101";
const RANGE_APPROVER_L3 = "E104:E105";
const CELL_CLOSING_APPROVER = "E117";
const CELL_PEKERJAAN_BARU_CHECK = "J12";
const CELL_PERPANJANGAN_CHECK = "J13";
const CELL_NO_IKA_LAMA = "J16";
const CELL_NO_REGISTRASI = "J19";
const RANGE_CLOSING_CHECKBOXES = "L110:L114";

/**
 * Fungsi ini berjalan otomatis setiap kali spreadsheet dibuka.
 * Ia bertugas membuat menu dan juga mencegah duplikasi file yang tidak sah.
 */
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentId = ss.getId();
  
  // --- LOGIKA ANTI-DUPLIKAT ---
  const sheet = ss.getSheets()[0];
  const stampedIdCell = sheet.getRange("Z1"); // Asumsikan kita menggunakan sel Z1 untuk stempel
  const stampedId = stampedIdCell.getValue();
  
  // Kondisi untuk duplikat tidak sah:
  // 1. Sel stempel Z1 tidak kosong (artinya ini adalah salinan dari file lain)
  // 2. ID di stempel Z1 TIDAK SAMA dengan ID file ini
  if (stampedId !== "" && stampedId !== currentId) {
    // Ini adalah duplikat ilegal! Hancurkan kontennya.
    ss.getSheets().forEach(s => s.clearContents());
    const firstSheet = ss.getSheets()[0];
    firstSheet.getRange("A1").setValue("DUPLIKASI TIDAK SAH. File ini telah dinonaktifkan. Untuk mendapatkan dokumen baru, silakan gunakan formulir permintaan resmi.");
    SpreadsheetApp.getUi().alert(
      "Peringatan Keamanan",
      "File ini adalah duplikat yang tidak sah dan kontennya telah dihapus. Harap gunakan formulir permintaan resmi untuk membuat dokumen baru.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return; // Hentikan fungsi di sini, jangan buat menu apa pun.
  }
  
  // --- LOGIKA PEMBUATAN MENU ---
  // Menu hanya akan dibuat jika file ini adalah file template asli
  // atau file salinan yang sah (bukan duplikat ilegal).
  if (currentId !== TEMPLATE_FILE_ID) { // Jangan tampilkan menu di file template utama
      SpreadsheetApp.getUi()
        .createMenu("📋 Approval IKA")
        .addItem("Ajukan Approval ✅", "tandaiUntukApproval")
        .addSeparator()
        .addItem("Batalkan Approval ❌", "cancelApproval")
        .addToUi();
  }
}

function tandaiUntukApproval() {
  const ui = SpreadsheetApp.getUi();
  const fileId = SpreadsheetApp.getActiveSpreadsheet().getId();

  // -- LANGKAH 1: KIRIM PERMINTAAN VALIDASI --
  ui.showSidebar(HtmlService.createHtmlOutput("<p>Memvalidasi data...</p>")); // Tampilkan feedback

  const validationPayload = { action: "validateRequest", fileId: fileId };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(validationPayload),
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };

  const validationResponse = UrlFetchApp.fetch(WEBAPP_URL, options);
  const validationResult = JSON.parse(validationResponse.getContentText());

  // Tutup sidebar loading
  ui.showSidebar(HtmlService.createHtmlOutput("<p>Selesai.</p>").setWidth(1));


  // -- LANGKAH 2: PERIKSA HASIL VALIDASI --
  if (validationResult.isValid === false) {
    // JIKA TIDAK VALID, TAMPILKAN ALERT DAN BERHENTI
    ui.alert("Validasi Gagal", validationResult.message, ui.ButtonSet.OK);
    return;
  }

  // -- LANGKAH 3: JIKA VALID, KIRIM PERMINTAAN EKSEKUSI --
  SpreadsheetApp.getActiveSpreadsheet().toast("Data valid. Memulai proses approval...", "Sukses", 5);

  const executionPayload = { action: "startApproval", fileId: fileId };
  options.payload = JSON.stringify(executionPayload);

  // Kirim permintaan eksekusi (tidak perlu menunggu respons)
  UrlFetchApp.fetch(WEBAPP_URL, options);
}

/**
 * Trigger yang berjalan otomatis saat ada editan di spreadsheet.
 * Dijalankan oleh akun pemilik skrip (installable trigger).
 *
 * @param {Object} e Objek event dari trigger OnEdit.
 */

// === FUNGSI WEB APP (Endpoint Utama) ===
function doGet(e) {
  // 1. Logika 'cancel' Anda tetap di sini
  if (e && e.parameter && e.parameter.action === 'cancelByUser') {
    return handleUserCancellation(e.parameter);
  }

  // 2. Buat template dari file HTML yang baru Anda buat
  //    (Ganti "Formulir" jika nama file Anda berbeda)
  const htmlTemplate = HtmlService.createTemplateFromFile("Formulir");

  // 3. (PENTING) Oper variabel 'SCRIPT_URL' dari .gs ke .html
  //    File HTML tidak bisa membaca konstanta SCRIPT_URL Anda secara langsung.
  htmlTemplate.SCRIPT_URL = SCRIPT_URL;

  // 4. Jalankan template untuk menghasilkan HTML akhir dan menampilkannya
  return htmlTemplate.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    if (e.postData && e.postData.type === 'application/json') {
      const params = JSON.parse(e.postData.contents);
            // BARU: Menangani permintaan validasi
      if (params.action === 'validateRequest') {
        const validationResult = validateRequestData(params.fileId);
        return ContentService.createTextOutput(JSON.stringify(validationResult))
          .setMimeType(ContentService.MimeType.JSON);
      }
      if (params.action === 'startApproval') {
        requestApproval(params.fileId);
        // Beri respons kembali ke skrip pemanggil. Teks ini tidak akan dilihat pengguna.
        return ContentService.createTextOutput("OK"); 
      }
    }

    // Jika bukan panggilan JSON, tangani sebagai form biasa dari email
    // Logika lama Anda ditempatkan di sini dan akan berjalan normal.
    if (e.parameter.action === "approval") {
      return processApproval(e.parameter);
    }
    if (e.parameter.action === "initiateClosing") {
      return handleInitiateClosing(e.parameter);
    }
    if (e.parameter.action === "finalizeClosing") {
      return handleFinalizeClosing(e.parameter);
    }
    if (e.parameter.formType === "newRequest") {
      return handleNewRequest(e.parameter);
    }
    
    // Jika tidak ada kondisi yang cocok
    throw new Error("Aksi tidak dikenali atau format data tidak didukung.");

  } catch (err) {
    Logger.log("Error di doPost: " + err.stack);
    // Perbaikan: Kembalikan error dalam format JSON yang bisa dibaca oleh client
    const errorResponse = {
      isValid: false,
      message: "Terjadi kesalahan internal pada server. Silakan hubungi administrator."
    };
    return ContentService.createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// === BAGIAN 1: PEMBUATAN FILE BARU DARI FORM HTML ===
function handleNewRequest(params) {
  // PERUBAHAN: Tidak lagi menggunakan 'nama'
  const dept = params.dept;
  const userEmail = params.email.trim().toLowerCase();

  // PERUBAHAN: Validasi disederhanakan
  if (!dept || !userEmail) return createErrorPage("❌ Semua field harus diisi.");
  if (!/^[^\s@]+@wingscorp\.com$/.test(userEmail)) return createErrorPage("❌ Gunakan email @wingscorp.com.");

  const template = DriveApp.getFileById(TEMPLATE_FILE_ID);
  const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  const now = new Date();
  const timeZone = Session.getScriptTimeZone();
  const year = Utilities.formatDate(now, timeZone, "yyyy");

  // --- PERUBAHAN: Membuat nama file draf sementara ---
  const timestamp = Utilities.formatDate(now, timeZone, "yyyy-MM-dd HH:mm:ss");
  const fileName = `DRAFT IKA - ${dept} - ${userEmail} (${timestamp})`;

  const subfolder = folder.getFoldersByName(year).hasNext() ? folder.getFoldersByName(year).next() : folder.createFolder(year);
  const newFile = template.makeCopy(fileName, subfolder);

  //newFile.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
  newFile.addEditor(userEmail);

  const ss = SpreadsheetApp.openById(newFile.getId());
  const sheet = ss.getSheets()[0];
  sheet.getRange("E96").setValue(userEmail); // Email pengguna dimasukkan ke E96

  // Tulis ID unik dari file baru ini ke dalam sel Z1.
  // Stempel ini akan digunakan oleh onOpen() untuk verifikasi duplikasi.
  const idCell = sheet.getRange("Z1");
  idCell.setValue(newFile.getId());
  
  // Lindungi dan sembunyikan sel stempel agar tidak bisa diubah.
  const protection = idCell.protect().setDescription('ID Protection');
  protection.removeEditors(protection.getEditors());
  sheet.hideColumns(idCell.getColumn()); // Sembunyikan kolom Z

  const logbook = SpreadsheetApp.openById(LOGBOOK_FILE_ID);
  let logSheet = logbook.getSheetByName(year);
  if (!logSheet) {
    const existingSheets = logbook.getSheets();
    let templateSheet = null;
    let latestYear = 0;

    // Cari sheet dengan nama tahun paling baru sebagai template
    existingSheets.forEach(s => {
        const sheetName = s.getName();
        const yearNum = parseInt(sheetName);
        if (!isNaN(yearNum) && sheetName.length === 4 && yearNum > latestYear) {
            latestYear = yearNum;
            templateSheet = s;
        }
    });

    // Jika tidak ada sheet tahun yang valid, gunakan sheet pertama sebagai fallback
    if (!templateSheet && existingSheets.length > 0) {
        templateSheet = existingSheets[0];
    }

    // Buat sheet baru dan salin format header jika template ditemukan
    logSheet = logbook.insertSheet(year);
    if (templateSheet) {
      templateSheet.getRange("1:8").copyTo(logSheet.getRange("1:8"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    }
  }

  const lastRow = logSheet.getLastRow() + 1;
  logSheet.getRange(lastRow, 1, 1, 19).setValues([[
    "", lastRow - 8, "", "", "", userEmail, dept, "", "", "", "", "", "", "", "Draft", "", newFile.getUrl(), "", ""
  ]]);

    // --- PERUBAHAN EMAIL DIMULAI DI SINI ---
  const subject = `✅ File IKA Baru Anda Telah Dibuat (${dept})`;
  const headerText = "File IKA Telah Dibuat";
  const headerColor = "#009980"; // Warna baru pilihan Anda

  const mainContent = `
    Dear Bapak/Ibu,<br><br>
    Dokumen Izin Kerja Aman (IKA) baru telah berhasil dibuat!<br><br>
    &#128193; <b>Link Dokumen:</b> <a href="${newFile.getUrl()}" target="_blank" style="color: #0051a2; text-decoration: none;">Klik di sini untuk membuka dokumen</a><br><br>
    Silakan lengkapi semua informasi yang diperlukan di dalam dokumen tersebut sebelum mengajukan persetujuan.`;

  const htmlBody = createStyledEmailBody(headerText, headerColor, mainContent);

  GmailApp.sendEmail(userEmail, subject, "", {
    htmlBody: htmlBody,
    from: "info.ims@wingscorp.com"
  });
  // --- AKHIR PERUBAHAN EMAIL ---

  return newFile.getUrl();
}

// === BAGIAN 2: LOGIKA APPROVAL ===
/**
 * Memeriksa semua data yang diperlukan untuk persetujuan.
 * @param {string} fileId ID dari spreadsheet yang akan divalidasi.
 * @returns {object} Objek yang berisi status validasi dan pesan.
 */
function validateRequestData(fileId) {
  try {
    const ss = SpreadsheetApp.openById(fileId);
    const sheet = ss.getSheets()[0];
    const riwayatSheet = ensureRiwayatApprovalSheet(ss); // Pastikan riwayat sheet ada

    const currentStatus = riwayatSheet.getRange("I1").getValue();
    if (currentStatus === "Approval In Progress") {
      return { isValid: false, message: "Proses Gagal: Approval sudah sedang berjalan untuk dokumen ini. Tidak bisa diajukan ulang." };
    }

    // Validasi 1: Pastikan hanya ada satu email pengaju
    const pengajuEmails = extractEmails(sheet.getRange(CELL_PENGAJU_EMAIL).getValue());
    if (pengajuEmails.length === 0) {
      return { isValid: false, message: "Validasi Gagal: Email Pengaju (di sel " + CELL_PENGAJU_EMAIL + ") wajib diisi." };
    }
    if (pengajuEmails.length > 1) {
      return { isValid: false, message: "Validasi Gagal: Hanya boleh ada satu email pengaju di sel " + CELL_PENGAJU_EMAIL + ". Mohon perbaiki dan ajukan kembali." };
    }
    
    // PERBAIKAN: Menggunakan konstanta
    const isPerpanjangan = sheet.getRange(CELL_PERPANJANGAN_CHECK).getValue();
    const noIkaLama = sheet.getRange(CELL_NO_IKA_LAMA).getValue();
    if (isPerpanjangan === true && noIkaLama === "") {
      return { isValid: false, message: "Validasi Gagal: Anda mencentang 'Perpanjangan', mohon isi No. Registrasi IKA Lama (di sel " + CELL_NO_IKA_LAMA + ")." };
    }
    
    // PERBAIKAN: Menggunakan konstanta
    const approvers = {
      2: extractEmails(sheet.getRange(RANGE_APPROVER_L2).getValues()),
      3: extractEmails(sheet.getRange(RANGE_APPROVER_L3).getValues())
    };
    if (!approvers[2] || approvers[2].length === 0) {
      return { isValid: false, message: "Validasi Gagal: Email Approver Layer 2 (di sel " + RANGE_APPROVER_L2 + ") wajib diisi." };
    }
    if (!approvers[3] || approvers[3].length === 0) {
      return { isValid: false, message: "Validasi Gagal: Email Approver Layer 3 (di sel " + RANGE_APPROVER_L3 + ") wajib diisi." };
    }

    // Jika semua validasi lolos
    return { isValid: true, message: "Semua data valid." };

  } catch (e) {
    return { isValid: false, message: "Terjadi error saat validasi: " + e.message };
  }
}

// FUNGSI UTAMA YANG DIPANGGIL DARI MENU "AJUKAN APPROVAL"
function requestApproval(fileId) {
  try {
    const ss = SpreadsheetApp.openById(fileId);
    const sheet = ss.getSheets()[0];
    const file = DriveApp.getFileById(fileId);
    const riwayatSheet = ensureRiwayatApprovalSheet(ss);

    // Validasi sudah dilakukan oleh validateRequestData, jadi di sini fokus ke eksekusi
    const pengajuEmail = sheet.getRange(CELL_PENGAJU_EMAIL).getValue();
    const komentarAwal = "Pengajuan Awal via Menu";

    const logbook = SpreadsheetApp.openById(LOGBOOK_FILE_ID);
    const logSheets = logbook.getSheets();
    let dept = null;
    let logSheet = null;
    let logRowIndex = -1;

    for (const lSheet of logSheets) {
      const data = lSheet.getDataRange().getValues();
      for (let i = 8; i < data.length; i++) {
        const linkInLog = data[i][16];
        if (linkInLog && typeof linkInLog === 'string') {
          const idMatch = linkInLog.match(/[-\w]{25,}/);
          if (idMatch && idMatch[0] === fileId) {
            dept = data[i][6];
            logSheet = lSheet;
            logRowIndex = i + 1;
            break;
          }
        }
      }
      // Ini adalah posisi 'break' yang benar
      if (dept) {
        break;
      }
    }

    if (!dept || !logSheet) {
      Logger.log("Proses berhenti: Dokumen tidak ditemukan di Logbook untuk file ID: " + fileId);
      return;
    }

    const now = new Date();
    const timeZone = Session.getScriptTimeZone();
    const yearMonth = Utilities.formatDate(now, timeZone, "yyyy.MM");
    const props = PropertiesService.getScriptProperties();
    const counterKey = `counter_${dept}_${yearMonth}`;
    let counter = parseInt(props.getProperty(counterKey) || "0") + 1;
    props.setProperty(counterKey, counter);
    const noReg = `SMU/IKA/${dept}/${yearMonth}/${String(counter).padStart(3, '0')}`;
    
    sheet.getRange(CELL_NO_REGISTRASI).setValue(noReg);
    logSheet.getRange(logRowIndex, 3).setValue(noReg);
    file.setName(noReg.replaceAll("/", "-"));

    const approvers = {
      1: extractEmails(sheet.getRange(CELL_PENGAJU_EMAIL).getValue()),
      2: extractEmails(sheet.getRange(RANGE_APPROVER_L2).getValues()),
      3: extractEmails(sheet.getRange(RANGE_APPROVER_L3).getValues())
    };

    const newNo = riwayatSheet.getLastRow();
    riwayatSheet.appendRow([newNo, new Date(), 0, "Ajukan Approval", pengajuEmail, komentarAwal]);

    riwayatSheet.getRange("I1").setValue("Approval In Progress");
    updateStatusInLogbook(noReg, "Approval");

    // Memanggil helper baru untuk memulai timer (mengatur U, Y, dan membersihkan R-T & V-X)
    if (logSheet && logRowIndex !== -1) {
      resetMasterTimer(noReg, now, 1); // Reset timer, set state ke 1
    }

    Logger.log("Memulai proses pengaturan akses senyap untuk: " + noReg);

    // 1. Ambil email pemilik file agar tidak terhapus
    const ownerEmail = DriveApp.getFileById(fileId).getOwner().getEmail();

    // 2. Hapus semua akses editor dan commenter yang ada saat ini (kecuali owner)
    const permissions = Drive.Permissions.list(fileId, {
      fields: 'permissions(id, emailAddress, role, permissionDetails, type)' 
    });

    if (permissions.permissions && permissions.permissions.length > 0) {
      permissions.permissions.forEach(permission => {
        const userEmail = permission.emailAddress;
        const userRole = permission.role;
        
        // Cek apakah izin ini warisan (inherited)
        // Di V3, detail ini ada di dalam permissionDetails
        let isInherited = false;
        if (permission.permissionDetails) {
          isInherited = permission.permissionDetails.some(detail => detail.inherited === true);
        }

        // LOGIKA FILTER:
        // 1. Jangan hapus Owner
        // 2. Jangan hapus jika itu akses warisan (Inherited)
        // 3. Hanya hapus yang memiliki role 'writer' (Editor)
        if (userEmail !== ownerEmail && !isInherited && userRole === 'writer') {
          try {
            Drive.Permissions.remove(fileId, permission.id);
            Logger.log(`Berhasil menghapus akses langsung ${userRole} dari ${userEmail}.`);
          } catch (e) {
            Logger.log(`Gagal menghapus akses dari ${userEmail || 'User Tanpa Email'}: ${e.message}`);
          }
        } else if (isInherited) {
          Logger.log(`Melewati ${userEmail || 'System Group'}: Akses warisan (Inherited) dari folder induk.`);
        }
      });
    }

    // 3. Kumpulkan semua approver dari L1, L2, dan L3
    const pengaju = sheet.getRange(CELL_PENGAJU_EMAIL).getValue();
    const approversL2 = extractEmails(sheet.getRange(RANGE_APPROVER_L2).getValues());
    const approversL3 = extractEmails(sheet.getRange(RANGE_APPROVER_L3).getValues());
    
    // Gabungkan semua dan hapus duplikat
    const allApprovers = new Set([pengaju, ...approversL2, ...approversL3]);

    // 4. Tambahkan setiap approver sebagai 'commenter' TANPA mengirim email notifikasi
    allApprovers.forEach(email => {
      if (email && email.includes('@')) {
        const permissionResource = {
          role: 'commenter',
          type: 'user',
          emailAddress: email
        };
        try {
          Drive.Permissions.create(permissionResource, fileId, {
            sendNotificationEmail: false // <-- KUNCI AJAIBNYA DI SINI
          });
          Logger.log(`Memberikan akses commenter (senyap) kepada ${email}.`);
        } catch (e) {
          Logger.log(`Gagal memberikan akses kepada ${email}: ${e.message}`);
        }
      }
    });

    // 5. Atur akses default untuk seluruh domain menjadi 'View'
    //file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
    Logger.log("Pengaturan akses senyap selesai.");
    // ================== AKHIR BLOK BARU ==================

    // Mengirim email approval dari sistem Anda (bukan notifikasi Google Drive)
    sendApprovalEmail(fileId, noReg, pengajuEmail, dept, approvers, 1);
    
  } catch (error) {
    Logger.log("TERJADI ERROR KRITIS di requestApproval: " + error.stack);
  }
}

// FUNGSI YANG MENANGANI CALLBACK DARI EMAIL (VERSI FINAL & BERSIH)
function processApproval(params) {
  const docId = params.docId;
  const layer = parseInt(params.layer);
  const approverEmail = params.approver;
  const komentar = params.komentar;
  const decision = params.decision;

  const ss = SpreadsheetApp.openById(docId);
  const sheet = ss.getSheets()[0];
  const riwayatSheet = ensureRiwayatApprovalSheet(ss);

  // --- GATEKEEPER ---
  const overallStatus = riwayatSheet.getRange("I1").getValue();
  if (overallStatus === "Cancelled" || overallStatus === "Rejected" || overallStatus === "Approved") {
    // <-- PERUBAHAN
    return createSimpleResponsePage("Proses untuk dokumen ini sudah dihentikan atau selesai.", false);
  }

 // Cek apakah approver ini sudah pernah memberikan keputusan DALAM SIKLUS INI
  const riwayatData = riwayatSheet.getDataRange().getValues();
      
  // 1. Temukan titik awal siklus approval saat ini (yaitu "Ajukan Approval" terakhir)
  let cycleStartIndex = 1; // Default mulai dari setelah header
  for (let i = riwayatData.length - 1; i > 0; i--) {
    if (riwayatData[i][3] === "Ajukan Approval") { // Kolom D (indeks 3) adalah 'Tindakan'
      cycleStartIndex = i;
      break; // Hentikan setelah menemukan yang paling terakhir
    }
  }

  // 2. Lakukan iterasi HANYA untuk siklus approval saat ini
  for (let i = cycleStartIndex; i < riwayatData.length; i++) {
    const recordedLayer = riwayatData[i][2];
    const recordedEmail = riwayatData[i][4];

    // Cek jika ada riwayat untuk layer dan approver yang sama persis di siklus ini
    if (recordedLayer == layer && recordedEmail == approverEmail) {
      // <-- PERUBAHAN
      return createSimpleResponsePage("Tindakan Gagal. Anda sudah pernah memberikan approval/rejection untuk siklus pengajuan ini.", false);
    }
  }

  try {
    let daftarApproverSah = [];
    if (layer === 1) {
      daftarApproverSah = extractEmails(sheet.getRange(CELL_PENGAJU_EMAIL).getValue());
    } else if (layer === 2) {
      daftarApproverSah = extractEmails(sheet.getRange(RANGE_APPROVER_L2).getValues());
    } else if (layer === 3) {
      daftarApproverSah = extractEmails(sheet.getRange(RANGE_APPROVER_L3).getValues());
    }

    // Cek apakah approver yang mengklik link masih ada di daftar approver yang sah untuk layer tersebut
    if (!daftarApproverSah.includes(approverEmail)) {
      // Jika tidak ditemukan, berarti link ini sudah tidak valid karena konfigurasi approver telah berubah.
      Logger.log(`Tindakan Ditolak: ${approverEmail} mencoba approve sebagai Layer ${layer}, tetapi tidak lagi terdaftar sebagai approver sah di layer tersebut. Kemungkinan menggunakan link lama.`);
      return createSimpleResponsePage("Tindakan Gagal. Link approval ini sudah tidak valid karena daftar approver telah diperbarui atau Anda tidak lagi ditugaskan di layer ini. Silakan gunakan link dari email terbaru.", false);
    }
  } catch (e) {
     Logger.log(`Error saat validasi otoritas approver: ${e.message}`);
     return createSimpleResponsePage("Terjadi kesalahan saat memvalidasi otoritas Anda. Hubungi administrator.", false);
  }

  // Pencegahan untuk menangani bug jika approver melakukan tindakan secara bersamaan
  const lock = LockService.getScriptLock();
  // Coba kunci selama 30 detik (30000 ms)
  if (!lock.tryLock(30000)) {
    Logger.log(`Gagal mendapatkan lock untuk processApproval L${layer} oleh ${approverEmail}. Server sibuk.`);
    return createSimpleResponsePage("Server sedang sibuk memproses permintaan lain. Silakan coba lagi dalam beberapa saat.", false);
  }

  Logger.log(`Lock diperoleh untuk L${layer} oleh ${approverEmail}.`);

  try {
    // GATEKEEPER 4 (AKSI GANDA - PENGECEKAN KEDUA DI DALAM LOCK)
    // Ini penting untuk menangani kasus jika dua pengguna lolos Gatekeeper 2 dan mengantri untuk mendapatkan lock.
    const riwayatDataPascaLock = riwayatSheet.getDataRange().getValues();
    let cycleStartIndexPascaLock = 1; 
    for (let i = riwayatDataPascaLock.length - 1; i > 0; i--) {
      if (riwayatDataPascaLock[i][3] === "Ajukan Approval") {
        cycleStartIndexPascaLock = i;
        break; 
      }
    }
    for (let i = cycleStartIndexPascaLock; i < riwayatDataPascaLock.length; i++) {
      const recordedLayer = riwayatDataPascaLock[i][2];
      const recordedEmail = riwayatDataPascaLock[i][4];
      if (recordedLayer == layer && recordedEmail == approverEmail) {
        Logger.log(`Tindakan duplikat terdeteksi DI DALAM LOCK untuk ${approverEmail}.`);
        return createSimpleResponsePage("Tindakan Gagal. Anda sudah pernah memberikan approval/rejection untuk siklus pengajuan ini.", false);
      }
    }

  // --- PROSES INTI ---
  // Format 'decision' ("approve" atau "reject") agar huruf pertamanya menjadi kapital sebelum dicatat.
  const formattedDecision = decision.charAt(0).toUpperCase() + decision.slice(1);
  riwayatSheet.appendRow([riwayatSheet.getLastRow(), new Date(), layer, formattedDecision, approverEmail, komentar]);

  const noReg = sheet.getRange(CELL_NO_REGISTRASI).getValue();
  const pengajuEmail = sheet.getRange(CELL_PENGAJU_EMAIL).getValue();
  const dept = noReg.split('/')[2];

  // Pastikan kita selalu menggunakan extractEmails untuk mendapatkan daftar yang bersih.
  const approvers = {
    1: extractEmails(sheet.getRange(CELL_PENGAJU_EMAIL).getValue()),
    2: extractEmails(sheet.getRange(RANGE_APPROVER_L2).getValues()),
    3: extractEmails(sheet.getRange(RANGE_APPROVER_L3).getValues())
  };

  if (decision === "approve") {
    if (layer === 1) {
      sendApprovalEmail(docId, noReg, pengajuEmail, dept, approvers, 2);
      riwayatSheet.appendRow([riwayatSheet.getLastRow(), new Date(), 1, "Notifikasi Terkirim ke Layer 2", "System", ""]);
      return createSimpleResponsePage("Terima kasih. Dokumen ini berhasil Anda setujui.");
    } 
    else if (layer === 2) {
      const uniqueRequiredApproversL2 = [...new Set(approvers[2])];
      const currentHistory = riwayatSheet.getDataRange().getValues();
      const actualApproversL2 = currentHistory.filter(row => row[2] === 2 && row[3] === 'Approve').map(row => row[4]);
      const uniqueActualApproversL2 = new Set(actualApproversL2);
      const allL2HaveApproved = uniqueRequiredApproversL2.every(requiredEmail => uniqueActualApproversL2.has(requiredEmail));

      if (allL2HaveApproved) {
        sendApprovalEmail(docId, noReg, pengajuEmail, dept, approvers, 3);
        riwayatSheet.appendRow([riwayatSheet.getLastRow(), new Date(), 2, "Notifikasi Terkirim ke Layer 3", "System", ""]);
        sendApprovalNotificationToRequester(docId, noReg, pengajuEmail, dept, 2, approverEmail, true, komentar);
        return createSimpleResponsePage("Terima kasih. Dokumen ini berhasil Anda setujui.");
      } else {
        // PANGGIL NOTIFIKASI BARU: Notifikasi SATU L2 telah approve
        sendApprovalNotificationToRequester(docId, noReg, pengajuEmail, dept, 2, approverEmail, true, komentar); 
        return createSimpleResponsePage("Terima kasih, dokumen ini berhasil Anda setujui.");
      }
    }
    else { // Layer 3
      const uniqueRequiredApproversL3 = [...new Set(approvers[3])];
      const currentHistoryL3 = riwayatSheet.getDataRange().getValues();
      const actualApproversL3 = currentHistoryL3.filter(row => row[2] === 3 && row[3] === 'Approve').map(row => row[4]);
      const uniqueActualApproversL3 = new Set(actualApproversL3);
      const allL3HaveApproved = uniqueRequiredApproversL3.every(requiredEmail => uniqueActualApproversL3.has(requiredEmail));

      if (allL3HaveApproved) {
        updateStatusInLogbook(noReg, "Approved");
        riwayatSheet.getRange("I1").setValue("Approved"); 
        riwayatSheet.appendRow([riwayatSheet.getLastRow(), new Date(), 3, "Approval Selesai", "System", ""]);
        
        // --- PERUBAHAN DI SINI: Hapus semua editor untuk mengunci file ---
        const file = DriveApp.getFileById(docId);
        const owner = file.getOwner();
        const editors = file.getEditors();

        editors.forEach(editor => {
          // Jangan hapus pemilik file
          if (owner && editor.getEmail() === owner.getEmail()) {
            return; // Lewati, jangan hapus
          }
          try {
            file.removeEditor(editor.getEmail());
          } catch (e) {
            Logger.log(`Gagal menghapus editor ${editor.getEmail()}: ${e.message}`);
          }
        });

        // --- PERUBAHAN EMAIL "APPROVED" DIMULAI DI SINI ---
        const subject = `✅ [APPROVAL IKA SELESAI] - ${noReg}`;
        const headerText = "Approval IKA Selesai";
        const headerColor = "#004080"; // Warna baru pilihan Anda
        const mainContent = `
          Dear Bapak/Ibu, <br><br>
          Dokumen IKA Anda telah disetujui sepenuhnya.<br><br>
          <b>Detail Dokumen:</b><br>
          &#128196; <b>Nomor Registrasi:</b> ${noReg}<br>
          &#128193; <b>Link Dokumen:</b> <a href="https://docs.google.com/spreadsheets/d/${docId}" target="_blank" style="color: #0051a2; text-decoration: none;">Klik di sini untuk membuka dokumen</a><br><br>
          Dokumen ini sekarang bersifat final dan akses edit telah dikunci.`;
        const htmlBody = createStyledEmailBody(headerText, headerColor, mainContent);
        GmailApp.sendEmail(pengajuEmail, subject, "", { htmlBody: htmlBody, from: "info.ims@wingscorp.com" });
        // --- AKHIR PERUBAHAN ---

        return createSimpleResponsePage("Terima kasih, dokumen ini berhasil Anda setujui.");
      } else {
        // PANGGIL NOTIFIKASI BARU: Notifikasi SATU L3 telah approve
        sendApprovalNotificationToRequester(docId, noReg, pengajuEmail, dept, 3, approverEmail, true, komentar); 
        return createSimpleResponsePage("Terima kasih, dokumen ini berhasil Anda setujui.");
      }
    }
  } else { // decision === 'reject'
      updateStatusInLogbook(noReg, "Rejected");
      riwayatSheet.getRange("I1").setValue("Rejected");
      
      riwayatSheet.appendRow([riwayatSheet.getLastRow(), new Date(), layer, "Proses Ditolak Sistem", "System", "Siklus approval dihentikan karena penolakan."]);

      const file = DriveApp.getFileById(docId);
      file.addEditor(pengajuEmail);

      const subject = `❌ [APPROVAL IKA DITOLAK] - ${noReg}`;
      const headerText = "Approval IKA Ditolak";
      const headerColor = "#80002a";
      const mainContent = `
        Dear Bapak/Ibu, <br><br>
        Pengajuan approval untuk dokumen IKA berikut telah ditolak.<br><br>
        <b>Detail Dokumen:</b><br>
        &#128196; <b>Nomor Registrasi:</b> ${noReg}<br>
        &#128193; <b>Link Dokumen:</b> <a href="https://docs.google.com/spreadsheets/d/${docId}" target="_blank" style="color: #0051a2; text-decoration: none;">Klik di sini untuk membuka dokumen</a><br><br>
        <b>Detail Penolakan:</b><br>
        - <b>Ditolak oleh:</b> ${approverEmail}<br>
        - <b>Komentar:</b> <i>${komentar}</i><br><br>
        Akses untuk memperbaiki dokumen telah diberikan kembali kepada Pengaju.`;

      // --- LOGIKA BARU PENGUMPULAN PENERIMA EMAIL ---
      const allRecipients = new Set([pengajuEmail]);
      const approversL2 = approvers[2] || [];

      // Selalu sertakan approver Layer 2, karena proses pasti sudah melewati mereka.
      approversL2.forEach(email => allRecipients.add(email));

      // HANYA JIKA penolakan terjadi di Layer 3, sertakan juga approver Layer 3.
      if (layer === 3) {
        const approversL3 = approvers[3] || [];
        approversL3.forEach(email => allRecipients.add(email));
      }
      
      const recipientList = [...allRecipients];
      // --- AKHIR LOGIKA BARU ---

      const htmlBody = createStyledEmailBody(headerText, headerColor, mainContent);
      
      GmailApp.sendEmail(recipientList.join(','), subject, "", { htmlBody: htmlBody, from: "info.ims@wingscorp.com" });
      Logger.log(`Email notifikasi penolakan dikirim ke: ${recipientList.join(',')}`);
      
      return createSimpleResponsePage("Dokumen ini berhasil Anda Tolak. Notifikasi telah dikirim kepada pihak terkait.", false);
    }
  } catch (err) { // <-- BARU: Menangkap error yang terjadi DI DALAM lock
      Logger.log(`Error Kritis di dalam processApproval (DI DALAM LOCK): ${err.stack}`);
      return createSimpleResponsePage("Terjadi kesalahan internal saat memproses permintaan Anda. Harap hubungi administrator.", false);
    } finally { // <-- BARU: Blok ini akan SELALU berjalan
      // Ini adalah bagian terpenting: melepaskan kunci
      // agar pengguna lain bisa masuk.
      lock.releaseLock();
      Logger.log(`Lock dilepas untuk L${layer} oleh ${approverEmail}.`);
    }
}

function sendApprovalEmail(fileId, noReg, pengajuEmail, dept, approvers, layer, customSubject = "") {
  const rawEmailData = approvers[layer];
  const approverEmails = extractEmails(rawEmailData);

  if (!approverEmails || approverEmails.length === 0) {
    Logger.log(`Tidak ada approver yang ditemukan untuk Layer ${layer}.`);
    return;
  }

  const subject = customSubject || `[APPROVAL IKA] - ${noReg}`;
  let headerColor, headerText, mainContent, buttonHtml;

  approverEmails.forEach(approverEmail => {
    if (approverEmail && typeof approverEmail === 'string' && approverEmail.includes('@')) {
        
        const cancellationUrl = `${SCRIPT_URL}?action=cancelByUser&docId=${fileId}`;

        if (layer === 1) {
          // --- Tampilan Khusus untuk Email Layer 1 ---
          headerColor = "#0051a2"; // Corporate Blue
          headerText = `Konfirmasi Pengajuan`;
          mainContent = `
            Dear Bapak/Ibu,<br><br>
            Pengajuan IKA dengan nomor registrasi <b>${noReg}</b> siap untuk dimulai. Mohon konfirmasi untuk melanjutkan proses approval ke approver berikutnya.<br><br>
            &#128193; <b>Link Dokumen:</b> <a href="https://docs.google.com/spreadsheets/d/${fileId}" target="_blank" style="color: #0051a2; text-decoration: none;">Klik untuk melihat dokumen</a><br><br>
            <b>PENTING:</b> Tombol <b>"Batalkan Pengajuan"</b> dapat Anda gunakan kapan saja (bahkan setelah konfirmasi) jika pekerjaan batal dilakukan.`;
          buttonHtml = `
            <table style="width:100%; border-spacing:10px; border-collapse: separate;">
              <tr>
                <td style="width:50%; text-align:center;">
                  <form method="POST" action="${SCRIPT_URL}" style="margin:0;">
                    <input type="hidden" name="action" value="approval">
                    <input type="hidden" name="docId" value="${fileId}">
                    <input type="hidden" name="layer" value="${layer}">
                    <input type="hidden" name="approver" value="${approverEmail}">
                    <input type="hidden" name="komentar" value="Pengajuan dikonfirmasi oleh pembuat.">
                    <input type="hidden" name="decision" value="approve">
                    <button type="submit" style="background:#28a745;color:white;padding:12px 22px;border:none;border-radius:5px;cursor:pointer;font-size:16px;width:100%; box-sizing: border-box;">Konfirmasi & Lanjutkan</button>
                  </form>
                </td>
                <td style="width:50%; text-align:center;">
                  <a href="${cancellationUrl}" target="_blank" style="background-color:#dc3545; color:white; padding:12px 22px; border-radius:5px; text-decoration:none; font-size:16px; display:block; width:100%; box-sizing: border-box;">Batalkan Pengajuan</a>
                </td>
              </tr>
            </table>`;
        } else {
          // --- Tampilan Standar untuk Email Layer 2 & 3 ---
          headerColor = (layer === 2) ? "#004d40" : "#8D6E63";
          headerText = (layer === 2) ? "Tinjauan & Persetujuan" : "Persetujuan Final";
          mainContent = `
            Dear Bapak/Ibu,<br><br>
            Sebuah dokumen IKA (${noReg}) membutuhkan persetujuan Anda.<br><br>
            &#128196; <b>Nomor Registrasi:</b> ${noReg}<br>
            &#128100; <b>Diajukan oleh:</b> ${pengajuEmail}<br>
            &#127970; <b>Departemen:</b> ${dept}<br>
            &#128193; <b>Dokumen:</b> <a href="https://docs.google.com/spreadsheets/d/${fileId}" target="_blank" style="color: #0051a2; text-decoration: none;">Klik untuk melihat dokumen</a><br><br>
            Silakan tinjau dokumen tersebut, kemudian berikan komentar dan keputusan Anda melalui menu di bawah ini:`;
          buttonHtml = `
            <form method="POST" action="${SCRIPT_URL}">
              <input type="hidden" name="action" value="approval">
              <input type="hidden" name="docId" value="${fileId}">
              <input type="hidden" name="layer" value="${layer}">
              <input type="hidden" name="approver" value="${approverEmail}">
              <label for="komentar"><b>Komentar Anda:</b></label><br>
              <textarea name="komentar" rows="4" style="width: 98%; border: 1px solid #ccc; border-radius: 4px; padding: 5px; margin-top: 5px;" required></textarea><br><br>
            
              <table style="width:100%; border-spacing:10px; border-collapse: separate;">
                <tr>
                  <td style="width:50%; text-align:center;">
                    <button type="submit" name="decision" value="approve" 
                            style="background:#28a745;color:white;padding:12px 22px;border:none;border-radius:5px;cursor:pointer;font-size:16px;width:100%; box-sizing: border-box;">✅ Approve</button>
                  </td>
                  <td style="width:50%; text-align:center;">
                    <button type="submit" name="decision" value="reject" 
                            style="background:#dc3545;color:white;padding:12px 22px;border:none;border-radius:5px;cursor:pointer;font-size:16px;width:100%; box-sizing: border-box;">❌ Reject</button>
                  </td>
                </tr>
              </table>
            </form>`;
        }
        
        const htmlBody = createStyledEmailBody(headerText, headerColor, mainContent, buttonHtml);
        GmailApp.sendEmail(approverEmail, subject, "", { htmlBody: htmlBody, from: "info.ims@wingscorp.com" });
    }
  });
}

/**
 * Fungsi ini dipanggil saat pengguna mengklik link pembatalan dari email.
 */
function handleUserCancellation(params) {
  const docId = params.docId;
  const userWhoClicked = Session.getActiveUser().getEmail();

  const ss = SpreadsheetApp.openById(docId);
  const sheet = ss.getSheets()[0];
  const pengajuAsli = sheet.getRange(CELL_PENGAJU_EMAIL).getValue();

  // Validasi Keamanan: Pastikan yang mengklik adalah pengaju asli
  if (userWhoClicked.toLowerCase() !== pengajuAsli.toLowerCase()) {
    return createSimpleResponsePage("Akses Ditolak. Hanya pengaju asli dokumen yang dapat membatalkan proses ini.", false);
  }

  // Panggil logika pembatalan inti dan periksa hasilnya
  const isSuccess = performCancellationLogic(docId);

  if (isSuccess) {
    return createSimpleResponsePage("Proses approval telah berhasil dibatalkan.", true);
  } else {
    return createSimpleResponsePage("Tindakan Gagal. Proses approval mungkin sudah selesai atau sudah pernah dibatalkan.", false);
  }
}

// ================================================================
// === FUNGSI INTI LOGIKA PEMBATALAN ===
// ================================================================
/**
 * Fungsi inti yang berisi semua logika untuk membatalkan IKA.
 * Mengembalikan 'true' jika berhasil dan 'false' jika gagal.
 */
function performCancellationLogic(docId) {
  try {
    const ss = SpreadsheetApp.openById(docId);
    const sheet = ss.getSheetByName("IKA"); // Lebih aman menggunakan nama sheet
    const riwayatSheet = ensureRiwayatApprovalSheet(ss);

    const currentStatus = riwayatSheet.getRange("I1").getValue();
    if (currentStatus !== "Approval In Progress") {
      Logger.log("Pembatalan gagal: Proses tidak sedang berjalan.");
      return false;
    }
    
    const noReg = sheet.getRange(CELL_NO_REGISTRASI).getValue();
    const pengajuEmail = sheet.getRange(CELL_PENGAJU_EMAIL).getValue();

    // Lakukan tindakan inti terlebih dahulu
    riwayatSheet.appendRow([riwayatSheet.getLastRow(), new Date(), "", "Dibatalkan oleh Pengaju", pengajuEmail, "Proses dibatalkan via email."]);
    riwayatSheet.getRange("I1").setValue("Cancelled");
    updateStatusInLogbook(noReg, "Cancelled");
    
    const file = DriveApp.getFileById(docId);
    file.addEditor(pengajuEmail);

    // --- LOGIKA BARU: HANYA PERIKSA SIKLUS APPROVAL SAAT INI ---
    const riwayatData = riwayatSheet.getDataRange().getValues();
    
    // 1. Temukan titik awal siklus approval terakhir (looping dari bawah ke atas).
    let cycleStartIndex = 1; // Default adalah baris setelah header
    for (let i = riwayatData.length - 1; i > 0; i--) {
      // Kolom D (indeks 3) adalah kolom 'Tindakan'
      if (riwayatData[i][3] === "Ajukan Approval") { 
        cycleStartIndex = i; // Simpan nomor baris dari "Ajukan Approval" terakhir
        break; // Hentikan pencarian setelah ditemukan
      }
    }

    // 2. Buat array baru yang hanya berisi data dari siklus saat ini.
    const currentCycleData = riwayatData.slice(cycleStartIndex);
    
    // 3. Lakukan pengecekan HANYA pada data siklus saat ini.
    const hasReachedL3 = currentCycleData.some(row => row[2] === 3 || row[3].toString().includes("Layer 3"));
    const hasReachedL2 = currentCycleData.some(row => row[2] >= 2 || row[3].toString().includes("Layer 2"));

    // Jika siklus saat ini (bukan riwayat lama) sudah mencapai L2, baru kirim notif.
    if (hasReachedL2) {
      const approversL2 = extractEmails(sheet.getRange(RANGE_APPROVER_L2).getValues());
      const allRecipients = new Set([...approversL2]);

      if (hasReachedL3) {
        const approversL3 = extractEmails(sheet.getRange(RANGE_APPROVER_L3).getValues());
        approversL3.forEach(email => allRecipients.add(email));
        Logger.log(`Proses IKA ${noReg} (siklus saat ini) sudah mencapai L3, notifikasi dikirim ke L2 & L3.`);
      } else {
        Logger.log(`Proses IKA ${noReg} (siklus saat ini) belum mencapai L3, notifikasi hanya dikirim ke L2.`);
      }
      
      const recipientList = [...allRecipients];

      if (recipientList.length > 0) {
        const subject = `📢 [APPROVAL IKA DIBATALKAN] - ${noReg}`;
        const headerText = "Proses Approval Dibatalkan";
        const headerColor = "#b71c1c"; // Merah Tua
        const mainContent = `
            Dear Bapak/Ibu,<br><br>
            Proses approval untuk dokumen IKA dengan nomor registrasi <b>${noReg}</b> telah <b>dibatalkan</b> oleh pengaju.<br><br>
            - <b>Dibatalkan oleh:</b> ${pengajuEmail}<br>
            - <b>Link Dokumen:</b> <a href="${ss.getUrl()}" target="_blank" style="color: #0051a2; text-decoration: none;">Klik di sini untuk melihat dokumen</a><br><br>
            Tidak ada tindakan lebih lanjut yang diperlukan dari Anda untuk dokumen ini.`;

        const htmlBody = createStyledEmailBody(headerText, headerColor, mainContent);
        GmailApp.sendEmail(recipientList.join(','), subject, "", { htmlBody: htmlBody, from: "info.ims@wingscorp.com" });
        Logger.log(`Email notifikasi pembatalan terkirim ke: ${recipientList.join(',')}`);
      }
    }
    // --- AKHIR BLOK NOTIFIKASI EMAIL BARU ---

    return true; // Laporkan sukses
  } catch (e) {
    Logger.log(`Error kritis saat pembatalan: ${e.stack}`);
    return false; // Laporkan gagal
  }
}

// ================================================================
// === FUNGSI FINAL: cancelApproval (Memanggil Fungsi Inti) ===
// ================================================================
function cancelApproval() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    "Konfirmasi Pembatalan",
    "Apakah Anda yakin ingin membatalkan proses approval yang sedang berjalan?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const docId = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  // Memanggil logika pembatalan inti dan memeriksa hasilnya
  const isSuccess = performCancellationLogic(docId);

  if (isSuccess) {
    ui.alert("Sukses!", "Proses approval telah berhasil dibatalkan.");
  } else {
    ui.alert("Gagal.", "Tidak ada proses approval yang sedang berjalan untuk dibatalkan.");
  }
}

// === BAGIAN 3: FUNGSI UTILITAS & LAIN-LAIN ===

function ensureRiwayatApprovalSheet(ss) {
  let riwayatSheet = ss.getSheetByName("Riwayat Approval");

  // Jika sheet "Riwayat Approval" belum ada, ini adalah satu-satunya saat kita membuat, memformat, dan melindunginya.
  if (!riwayatSheet) {
    riwayatSheet = ss.insertSheet("Riwayat Approval");
    riwayatSheet.appendRow(["No.", "Timestamp", "Layer", "Tindakan", "Email", "Komentar"]);
    
    // --- BLOK FORMATTING BARU ---
    // 1. Atur lebar kolom agar sesuai
    riwayatSheet.setColumnWidth(1, 40);  // Kolom A (No.)
    riwayatSheet.setColumnWidth(2, 150); // Kolom B (Timestamp)
    riwayatSheet.setColumnWidth(3, 60);  // Kolom C (Layer)
    riwayatSheet.setColumnWidth(4, 250); // Kolom D (Tindakan)
    riwayatSheet.setColumnWidth(5, 250); // Kolom E (Email)
    riwayatSheet.setColumnWidth(6, 350); // Kolom F (Komentar)

    // 2. Format header (Baris 1)
    const headerRange = riwayatSheet.getRange("A1:F1");
    headerRange.setBackground("#4a5568"); // Latar belakang abu-abu gelap
    headerRange.setFontColor("#ffffff");   // Teks putih
    headerRange.setFontWeight("bold");     // Teks tebal
    headerRange.setHorizontalAlignment("center");

    // 3. Atur perataan untuk kolom tertentu
    riwayatSheet.getRange("A2:A").setHorizontalAlignment("center"); // Kolom No.
    riwayatSheet.getRange("C2:C").setHorizontalAlignment("center"); // Kolom Layer

    // 4. Terapkan warna baris selang-seling (banded rows)
    const dataRange = riwayatSheet.getRange("A1:F");
    dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

    // Label Status 
    riwayatSheet.getRange("H1").setValue("Status Dokumen:").setFontWeight("bold").setHorizontalAlignment("right");
    riwayatSheet.getRange("I1").setFontWeight("bold");

    // Langsung proteksi sheet yang baru dibuat.
    const protection = riwayatSheet.protect().setDescription('Lock Riwayat Approval');
    
    // --- PERUBAHAN LOGIKA PROTEKSI ---
    // Tentukan email owner project yang selalu boleh mengedit
    const projectOwnerEmail = ADMIN_EMAIL;

    // Dapatkan daftar editor yang ada pada proteksi
    const editors = protection.getEditors();
    
    // Hapus semua editor KECUALI owner project
    editors.forEach(editor => {
      if (editor.getEmail() !== projectOwnerEmail) { 
        protection.removeEditor(editor);
      }
    });

    // Pastikan owner project selalu ada sebagai editor (sebagai pengaman)
    protection.addEditor(projectOwnerEmail);
    
    // Pastikan hanya editor yang ditentukan yang bisa mengedit, bukan seluruh domain.
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
    // --- AKHIR PERUBAHAN LOGIKA PROTEKSI ---
  }

  // Kembalikan object sheet. Jika sheet sudah ada sebelumnya, kita tidak menyentuh proteksinya sama sekali.
  return riwayatSheet;
}


/**
 * FUNGSI PEMBANTU (VERSI v5 - Sederhana)
 */
function findCurrentApprovers(fileId, noReg) {
  try {
    const ss = SpreadsheetApp.openById(fileId);
    const sheet = ss.getSheetByName("IKA");
    if (!sheet) {
      throw new Error("Sheet IKA tidak ditemukan");
    }
    const riwayatSheet = ensureRiwayatApprovalSheet(ss);
    const riwayatData = riwayatSheet.getDataRange().getValues();

    // 1. Dapatkan info utama
    const pengajuEmail = sheet.getRange(CELL_PENGAJU_EMAIL).getValue();
    const dept = noReg.split('/')[2];
    
    // 2. Dapatkan daftar approver WAJIB
    const reqL1 = extractEmails(pengajuEmail);
    const reqL2 = extractEmails(sheet.getRange(RANGE_APPROVER_L2).getValues());
    const reqL3 = extractEmails(sheet.getRange(RANGE_APPROVER_L3).getValues());
    
    // 3. Temukan awal siklus
    let cycleStartIndex = 1; 
    for (let i = riwayatData.length - 1; i > 0; i--) {
      if (riwayatData[i][3] === "Ajukan Approval") { 
        cycleStartIndex = i;
        break;
      }
    }
    const currentCycleData = riwayatData.slice(cycleStartIndex);

    // 4. Dapatkan daftar yang SUDAH approve
    const approvedL1_Set = new Set(currentCycleData.filter(r => r[2] == 1 && r[3] === 'Approve').map(r => r[4]));
    const approvedL2_Set = new Set(currentCycleData.filter(r => r[2] == 2 && r[3] === 'Approve').map(r => r[4]));
    const approvedL3_Set = new Set(currentCycleData.filter(r => r[2] == 3 && r[3] === 'Approve').map(r => r[4]));

    // 5. Tentukan Bottleneck dan kembalikan objek result
    
    // Cek Layer 1
    const allL1HaveApproved = reqL1.every(email => approvedL1_Set.has(email));
    if (!allL1HaveApproved) {
      return {
        pendingLayer: 1, 
        pendingEmails: reqL1.filter(email => !approvedL1_Set.has(email)),
        pengajuEmail: pengajuEmail,
        dept: dept
      };
    }

    // Cek Layer 2
    const allL2HaveApproved = reqL2.every(email => approvedL2_Set.has(email));
    if (!allL2HaveApproved) {
      return {
        pendingLayer: 2, 
        pendingEmails: reqL2.filter(email => !approvedL2_Set.has(email)),
        pengajuEmail: pengajuEmail,
        dept: dept
      };
    }

    // Cek Layer 3
    const allL3HaveApproved = reqL3.every(email => approvedL3_Set.has(email));
    if (!allL3HaveApproved) {
      return {
        pendingLayer: 3, 
        pendingEmails: reqL3.filter(email => !approvedL3_Set.has(email)),
        pengajuEmail: pengajuEmail,
        dept: dept
      };
    }

    return null; // Semua sudah approve

  } catch (e) {
    Logger.log(`Gagal 'findCurrentApprovers' untuk file ${fileId}: ${e.message}`);
    return null; // Gagal
  }
}

/**
 * HELPER V12 (BARU): Mereset MASTER TIMER (U), STATE (Y),
 * dan membersihkan KEDUA checklist (Approval V-X DAN Closing R-S-T).
 * @param {string} noReg Nomor registrasi IKA.
 * @param {Date|null} timestamp Tanggal baru, atau 'null' untuk mematikan.
 * @param {string|number} state State baru (1, 2, "Expired", "Closing", "Close").
 */
function resetMasterTimer(noReg, timestamp, state) {
  try {
    const logbook = SpreadsheetApp.openById(LOGBOOK_FILE_ID);
    const sheets = logbook.getSheets();
    
    for (let sheet of sheets) {
      const data = sheet.getRange(9, 3, sheet.getLastRow() - 8, 1).getValues(); // Cek Kolom C
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === noReg) {
          const targetRow = i + 9;
          
          const timerRange = sheet.getRange(targetRow, MASTER_TIMESTAMP_COL); // Kolom U
          const stateRange = sheet.getRange(targetRow, MASTER_STATE_COL);     // Kolom Y
          const apprCheckRange = sheet.getRange(targetRow, APPROVAL_H1_COL, 1, 3); // Kolom V,W,X
          const closCheckRange = sheet.getRange(targetRow, CLOSING_H1_COL, 1, 3); // Kolom R,S,T

          if (timestamp === null) {
            // Matikan/bersihkan semua
            timerRange.clearContent();
            apprCheckRange.clearContent();
            closCheckRange.clearContent();
            stateRange.setValue(state); // Set state akhir (misal "Close")
            Logger.log(`Semua timer/reminder untuk ${noReg} telah DIMATIKAN.`);
          } else {
            // Reset timer dan bersihkan kedua checklist
            timerRange.setValue(timestamp);
            apprCheckRange.clearContent();
            closCheckRange.clearContent();
            stateRange.setValue(state); // Set state baru (misal "Expired" atau 2)
            Logger.log(`Master Timer untuk ${noReg} telah DI-RESET ke state: ${state}.`);
          }
          return; // Hentikan setelah ditemukan
        }
      }
    }
  } catch (e) {
    Logger.log(`Gagal resetMasterTimer: ${e.message}`);
  }
}

/**
 * FUNGSI HELPER BARU (V12)
 * Mengirim email konfirmasi/reminder penutupan ke Pemilik Area.
 * @param {string} approverEmail Email penerima (Pemilik Area).
 * @param {string} noReg Nomor registrasi IKA.
 * @param {string} pengajuEmail Email pengaju (untuk info).
 * @param {string} dept Departemen.
 * @param {string} fileId ID file IKA.
 * @param {string} [customSubject=""] Subjek email kustom (opsional).
 */
function sendClosingConfirmationEmail(approverEmail, noReg, pengajuEmail, dept, fileId, customSubject = "") {
  
  // Gunakan subjek kustom jika ada, jika tidak, gunakan subjek default
  const subject = customSubject || `[KONFIRMASI PENUTUPAN IKA] - ${noReg}`;
  const headerText = "Konfirmasi Penutupan IKA";
  const headerColor = "#004d40"; // Hijau Tua
  
  const mainContent = `
        Dear Bapak/Ibu,<br><br>
        Sebuah dokumen IKA membutuhkan konfirmasi Anda pada proses penutupan, detail dokumen sebagai berikut:<br><br>
        &#128196; <b>Nomor Registrasi:</b> ${noReg}<br>
        &#128100; <b>Penutupan Diajukan oleh:</b> ${pengajuEmail}<br>
        &#127970; <b>Departemen:</b> ${dept}<br>
        &#128193; <b>Dokumen:</b> <a href="https://docs.google.com/spreadsheets/d/${fileId}" target="_blank" style="color: #0051a2; text-decoration: none;">Klik untuk melihat dokumen</a><br><br>
        Mohon periksa kembali dokumen dan berikan konfirmasi akhir dengan menekan tombol di bawah.`;
  
  const buttonHtml = `
        <div style="text-align: center;">
          <form method="POST" action="${SCRIPT_URL}" style="display: inline-block;">
            <input type="hidden" name="action" value="finalizeClosing">
            <input type="hidden" name="docId" value="${fileId}">
            <input type="hidden" name="approverEmail" value="${approverEmail}">
            <button type="submit" style="background:#004d40;color:white;padding:12px 28px;border:none;border-radius:5px;cursor:pointer;font-size:16px;font-weight:bold;width:100%; box-sizing: border-box;">✅ Konfirmasi Penutupan</button>
          </form>
        </div>`;
        
  const htmlBody = createStyledEmailBody(headerText, headerColor, mainContent, buttonHtml);

  try {
    GmailApp.sendEmail(approverEmail, subject, "", {
      htmlBody: htmlBody,
      from: "info.ims@wingscorp.com"
    });
    Logger.log(`Email konfirmasi closing terkirim ke ${approverEmail} untuk ${noReg}. Subjek: ${subject}`);
  } catch (e) {
    Logger.log(`Gagal kirim email konfirmasi closing ke ${approverEmail}: ${e.message}`);
  }
}

/**
 * Fungsi pembantu untuk mendapatkan alamat email pengaju langsung dari sel E96 di dalam file.
 * @param {string} fileId ID dari file Google Sheet.
 * @returns {string} Alamat email pengaju, atau string kosong jika tidak ditemukan.
 */
function getEditorEmail(fileId) {
  try {
    const ss = SpreadsheetApp.openById(fileId);
    
    // PERBAIKAN: Menggunakan konstanta dan getSheetByName agar lebih andal
    const sheet = ss.getSheetByName("IKA");
    if (!sheet) return ""; // Tambahan pengaman jika sheet IKA tidak ada

    const pengaju = sheet.getRange(CELL_PENGAJU_EMAIL).getValue();
    
    if (pengaju && typeof pengaju === 'string' && pengaju.includes("@")) {
      return pengaju.trim();
    }
  } catch (e) {
    Logger.log(`Gagal mendapatkan email pengaju untuk file ID ${fileId}: ${e.message}`);
  }
  
  return ""; 
}

function updateStatusInLogbook(noReg, status) {
  Logger.log("Mencari No. Reg: '" + noReg + "' untuk diubah menjadi status: '" + status + "'");

  const logbook = SpreadsheetApp.openById(LOGBOOK_FILE_ID);
  const sheets = logbook.getSheets();
  for (let sheet of sheets) {
    const dataRange = sheet.getRange(9, 3, sheet.getLastRow() - 8, 1); // Cukup cek kolom B
    const data = dataRange.getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === noReg) { // Cek kolom B (indeks 0)
        const targetRow = i + 9;
        Logger.log("Kecocokan ditemukan di baris " + targetRow + " pada sheet " + sheet.getName());
        
        // Perbarui Status Izin di Kolom O (kolom ke-15)
        sheet.getRange(targetRow, 15).setValue(status); 

        // JIKA status adalah "Approved", isi timestamp di Kolom P (kolom ke-15)
        if (status === "Approved") {
          sheet.getRange(targetRow, 16).setValue(new Date());
        }
        return;
      }
    }
  }
  Logger.log("PERINGATAN: Tidak ada kecocokan yang ditemukan untuk No. Reg: '" + noReg + "' di seluruh Logbook.");
}

function createErrorPage(message) {
  return HtmlService.createHtmlOutput(`<html><body style="font-family:Arial;text-align:center;padding:50px;color:red;"><h3>${message}</h3></body></html>`);
}

// Fungsi ini mengirim email ke Pengaju untuk memulai proses closing
function sendClosingEmailToPengaju(userEmail, noReg, fileId, notifType, customSubject = "") {
  
  // --- PERUBAHAN DI SINI ---
  // Gunakan subjek kustom jika ada, jika tidak, gunakan subjek default (H+0)
  const subject = customSubject || `[PERMINTAAN PENUTUPAN IKA] - ${noReg}`;
  // --- AKHIR PERUBAHAN ---
  
  const headerText = "Permintaan Penutupan IKA";
  const headerColor = "#0051a2"; // Biru

  const mainContent = `
      Dear Bapak/Ibu,<br><br>
      Masa berlaku IKA dengan nomor dokumen <b>${noReg}</b> telah berakhir.<br><br>
      &#128193; <b>Link Dokumen:</b> <a href="https://docs.google.com/spreadsheets/d/${fileId}" target="_blank" style="color: #0051a2; text-decoration: none;">Klik di sini untuk membuka dokumen</a><br><br>
      Mohon untuk melakukan konfirmasi penutupan IKA. Dengan menekan tombol <b>'Close IKA'</b> di bawah ini, Anda menyatakan telah melaksanakan dan menyetujui seluruh butir pernyataan berikut:<br><br>
      <ul style="margin: 0; padding-left: 20px; line-height: 1.5;">
        <li>Pekerjaan telah selesai dilaksanakan.</li>
        <li>Peralatan telah dirapihkan dan dikembalikan.</li>
        <li>Lokasi kerja telah aman dan dibersihkan, tanpa cemaran.</li>
        <li>Jika diperlukan, LOTO sudah dilepas.</li>
        <li>Seluruh tindakan pengamanan lainnya telah diselesaikan.</li>
      </ul>`;
  const buttonHtml = `
      <div style="text-align: center;">
        <form method="POST" action="${SCRIPT_URL}" style="display: inline-block;">
          <input type="hidden" name="action" value="initiateClosing">
          <input type="hidden" name="docId" value="${fileId}">
          <input type="hidden" name="pengajuEmail" value="${userEmail}">
          <button type="submit" style="background:#0051a2;color:white;padding:12px 28px;border:none;border-radius:5px;cursor:pointer;font-size:16px;font-weight:bold;">Close IKA</button>
        </form>
      </div>`;
      
  const htmlBody = createStyledEmailBody(headerText, headerColor, mainContent, buttonHtml);

  GmailApp.sendEmail(userEmail, subject, "", {
    htmlBody: htmlBody,
    from: "info.ims@wingscorp.com"
  });
}

function handleInitiateClosing(params) {
  Logger.log("DEBUG: handleInitiateClosing DIMULAI. DocID: " + params.docId);
  const docId = params.docId;
  const pengajuEmail = params.pengajuEmail;
  const ss = SpreadsheetApp.openById(docId);
  const sheet = ss.getSheetByName("IKA");
  const riwayatSheet = ensureRiwayatApprovalSheet(ss);
  
  const currentStatus = riwayatSheet.getRange("I1").getValue();
  Logger.log("DEBUG: Status I1 yang dibaca adalah: '" + currentStatus + "'");
  if (currentStatus === "Proses Penutupan" || currentStatus === "Close") {
    return createSimpleResponsePage("Tindakan gagal. Anda sudah melakukan tindakan penutupan.", false);
  }

  // --- LAKUKAN SEMUA TINDAKAN DI SPREADSHEET DULU ---
  riwayatSheet.appendRow([riwayatSheet.getLastRow() + 1, new Date(), "", "Pengajuan penutupan telah dilakukan", pengajuEmail, "Konfirmasi penutupan via email."]);
  riwayatSheet.getRange("I1").setValue("Proses Penutupan");
  sheet.getRange(RANGE_CLOSING_CHECKBOXES).setValue("Ya");

  const noReg = sheet.getRange(CELL_NO_REGISTRASI).getValue();
  const dept = noReg.split('/')[2];
  
  // 1. Update status Logbook
  updateStatusInLogbook(noReg, "Proses Penutupan");
  
  // 2. Reset Master Timer (U) & State (Y)
  resetMasterTimer(noReg, new Date(), "Proses Penutupan");

  // --- SEKARANG, KIRIM EMAIL (MENGGUNAKAN HELPER BARU) ---
  const rawApproverData = sheet.getRange(CELL_CLOSING_APPROVER).getValue();
  const approverEmails = extractEmails(rawApproverData);

  if (approverEmails.length > 0) {
    approverEmails.forEach(approverEmail => {
      // Panggil helper 'sendClosingConfirmationEmail'
      sendClosingConfirmationEmail(
        approverEmail, 
        noReg, 
        pengajuEmail, 
        dept, 
        docId
        // customSubject dibiarkan kosong (untuk H+0)
      );
    });
  }
  
  riwayatSheet.appendRow([riwayatSheet.getLastRow(), new Date(), "", "Notifikasi penutupan kepada pemilik area telah dikirim", "System", `Terkirim ke ${approverEmails.join(", ")}`]);
  return createSimpleResponsePage("Terima kasih. Proses penutupan IKA telah dimulai dan email konfirmasi penutupan telah dikirimkan ke pemilik area.");
}

function handleFinalizeClosing(params) {
  const docId = params.docId;
  const approverEmail = params.approverEmail;
  
  const ss = SpreadsheetApp.openById(docId);
  const sheet = ss.getSheetByName("IKA");
  const riwayatSheet = ensureRiwayatApprovalSheet(ss);

  // Penjaga Gerbang: Cek apakah proses sudah close
  const currentStatus = riwayatSheet.getRange("I1").getValue();
  if (currentStatus === "Close") {
    return createSimpleResponsePage("Tindakan gagal. Anda telah melakukan konfirmasi tindakan penutupan.", false);
  }
  
  // Penjaga Gerbang Anti-Duplikat: Cek apakah approver ini sudah pernah bertindak
  const riwayatData = riwayatSheet.getDataRange().getValues();
  for (let i = 1; i < riwayatData.length; i++) {
    if (riwayatData[i][4] == approverEmail && riwayatData[i][3].includes("Pemilik Area telah melakukan konfirmasi")) {
      return createSimpleResponsePage("Tindakan Gagal. Anda sudah pernah menyetujui penutupan IKA ini.", false);
    }
  }

  // --- PROSES INTI ---
  // 1. Catat tindakan konfirmasi oleh Pemilik Area
  riwayatSheet.appendRow([riwayatSheet.getLastRow(), new Date(), "", "Pemilik Area telah melakukan konfirmasi penutupan IKA", approverEmail, "Konfirmasi akhir via email."]);
  
  // 2. Langsung ubah status menjadi "Close" tanpa menunggu approver lain
  const noReg = sheet.getRange(CELL_NO_REGISTRASI).getValue();
  updateStatusInLogbook(noReg, "Close");
  resetMasterTimer(noReg, null, "Close");
  riwayatSheet.getRange("I1").setValue("Close");

  // 3. TAMBAHAN: Catat tindakan akhir oleh sistem
  riwayatSheet.appendRow([riwayatSheet.getLastRow(), new Date(), "", "IKA Close", "System", "Dokumen telah ditutup sepenuhnya."]);

  // --- PENJELASAN PERUBAHAN ---
  // Blok kode di bawah ini adalah TAMBAHAN BARU.
  // Tujuannya adalah untuk mengirimkan email notifikasi terakhir kepada pemilik dokumen (pengaju)
  // untuk memberitahu bahwa IKA mereka telah resmi ditutup setelah dikonfirmasi oleh Pemilik Area.
  // Email ini menggunakan template modern yang sama dengan email lainnya untuk konsistensi.
  
  // ================================================================
  // === BLOK BARU: KIRIM NOTIFIKASI KE PEMILIK DOKUMEN ===
  // ================================================================
  try {
    const pengajuEmail = sheet.getRange(CELL_PENGAJU_EMAIL).getValue();
    if (pengajuEmail) {
      const subject = `[IKA BERHASIL DITUTUP] - ${noReg}`;
      const headerText = "IKA Telah Berhasil Ditutup";
      const headerColor = "#4A5568"; // Warna Slate Gray
      const mainContent = `
          Dear Bapak/Ibu,<br><br>
          Dokumen IKA dengan nomor registrasi <b>${noReg}</b> telah berhasil ditutup sepenuhnya setelah mendapatkan konfirmasi akhir dari Pemilik Area.<br><br>
          &#128193; <b>Link Dokumen:</b> <a href="${ss.getUrl()}" target="_blank" style="color: #0051a2; text-decoration: none;">Klik di sini untuk membuka dokumen</a><br><br>
          Tidak ada tindakan lebih lanjut yang diperlukan.<br><br>
          Terima kasih atas kerja sama Anda.`;
      
      const htmlBody = createStyledEmailBody(headerText, headerColor, mainContent);
      GmailApp.sendEmail(pengajuEmail, subject, "", { htmlBody: htmlBody, from: "info.ims@wingscorp.com" });
      Logger.log(`Notifikasi penutupan berhasil terkirim ke ${pengajuEmail} untuk IKA ${noReg}.`);
    }
  } catch(e) {
    Logger.log(`Gagal mengirim email notifikasi penutupan ke pengaju: ${e.message}`);
  }
  // ================================================================
  // === AKHIR BLOK BARU ===
  // ================================================================
  
  return createSimpleResponsePage("Terima kasih. IKA telah berhasil ditutup.");
}

/**
 * FUNGSI REMINDER CLOSING (VERSI v13 - Sesuai Permintaan)
 * DIPICU 2x SEHARI (06:00 & 18:00)
 * Menggunakan MASTER TIMER (U) dan MASTER STATE (Y)
 * Checklist Closing H+1 (R), H+3 (S), H+7 (T)
 */
function updateLogbookFromModifiedFiles() {
  const logbook = SpreadsheetApp.openById(LOGBOOK_FILE_ID);
  const now = new Date();
  const startTime = new Date().getTime();
  const TIME_LIMIT_MS = 300000; // Failsafe 5 menit
  const timezone = logbook.getSpreadsheetTimeZone();
  const todayString = Utilities.formatDate(now, timezone, "yyyyMMdd");
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth();
  
  Logger.log("Memulai Pindaian Reminder Closing (v13)...");

  const sheetsToProcessNames = [String(currentYear)];
  if (currentMonth < 6) {
    sheetsToProcessNames.push(String(currentYear - 1));
  }

  sheetsToProcessNames.forEach(sheetName => {
    const sheet = logbook.getSheetByName(sheetName);
    if (!sheet) return;
    
    // Ambil data sampai Kolom Y (25)
    const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), 25); 
    const data = dataRange.getValues();
    
    // Indeks Kolom (berbasis 0)
    const STATUS_COL = 14; // O
    const LINK_COL = 16; // Q
    const NO_REG_COL = 2; // C
    const PENGAJU_EMAIL_COL = 5; // F
    const TGL_SELESAI_COL = 13; // N
    
    // Kolom Bersama
    const TIMER_COL = MASTER_TIMESTAMP_COL - 1; // U (indeks 20)
    
    // Kolom Khusus Closing
    const H1_COL = CLOSING_H1_COL - 1; // R (indeks 17)
    const H3_COL = CLOSING_H3_COL - 1; // S (indeks 18)
    const H7_COL = CLOSING_H7_COL - 1; // T (indeks 19)

    for (let i = 8; i < data.length; i++) { // Mulai dari baris 9
      try {
        const statusIzin = data[i][STATUS_COL];

        if (statusIzin !== "Approved" && 
            statusIzin !== "Expired" && 
            statusIzin !== "Proses Penutupan") {
          continue;
        }

        const link = data[i][LINK_COL];
        const noReg = data[i][NO_REG_COL];
        const pengajuEmail = data[i][PENGAJU_EMAIL_COL];
        const tglSelesai = data[i][TGL_SELESAI_COL];
        if (!link || !noReg || !pengajuEmail || !(tglSelesai instanceof Date)) continue;
        const fileIdMatch = link.match(/[-\w]{25,}/);
        if (!fileIdMatch) continue;
        const fileId = fileIdMatch[0];
        const dept = noReg.split('/')[2];
        const targetRow = i + 1;

        // ================================================================
        // LOGIKA 1: DARI "Approved" -> "Expired" (TIMER START - Per Poin 1)
        // ================================================================
        if (statusIzin === "Approved") {
          const tglSelesaiString = Utilities.formatDate(tglSelesai, timezone, "yyyyMMdd");
          
          if (parseInt(todayString) > parseInt(tglSelesaiString)) {
            Logger.log(`IKA ${noReg} diubah menjadi Expired.`);
            
            // Panggil helper untuk me-reset timer (U), state (Y), dan SEMUA checklist
            resetMasterTimer(noReg, now, "Expired");
            sheet.getRange(targetRow, STATUS_COL + 1).setValue("Expired"); // Update Kolom O
            
            // Kirim email penutupan pertama (H+0) ke Pengaju
            sendClosingEmailToPengaju(pengajuEmail, noReg, fileId, "H+0");
            
            // Sesuai Poin 1, KITA TIDAK MENULIS "Terkirim" di Kolom R
          }
          continue; 
        }

        // ================================================================
        // LOGIKA 2 & 3: MENANGANI "Expired" atau "Proses Penutupan" (Per Poin 2 & 4)
        // ================================================================
        
        const timestamp = data[i][TIMER_COL]; // Baca Timer (U)
        if (!(timestamp instanceof Date)) {
          continue; 
        }

        const diffMs = now.getTime() - timestamp.getTime();
        const daysPassed = Math.floor(diffMs / (1000 * 60 * 60 * 24));

        if (daysPassed > 30) continue; // Filter 30 Hari
        if (daysPassed < 1) continue; // Masa tenggang H+1
        
        // Baca checklist closing dari array 'data'
        const h1_status = data[i][H1_COL];
        const h3_status = data[i][H3_COL];
        const h7_status = data[i][H7_COL];

        let reminderDayToSend = 0;
        let reminderNum = 0; // #1, #2, atau #3
        let colToUpdate = -1;

        if (daysPassed >= 7 && h7_status !== "Terkirim") {
          reminderDayToSend = 7; reminderNum = 3; colToUpdate = CLOSING_H7_COL;
        } else if (daysPassed >= 3 && h3_status !== "Terkirim") {
          reminderDayToSend = 3; reminderNum = 2; colToUpdate = CLOSING_H3_COL;
        } else if (daysPassed >= 1 && h1_status !== "Terkirim") {
          reminderDayToSend = 1; reminderNum = 1; colToUpdate = CLOSING_H1_COL;
        }

        if (reminderDayToSend > 0) {
          
          if (statusIzin === "Expired") {
            // Bottleneck = PENGAJU (Per Poin 2)
            const customSubject = `[REMINDER #${reminderNum} - PENUTUPAN IKA] - ${noReg}`;
            Logger.log(`Mengirim reminder closing H+${reminderDayToSend} (Reminder #${reminderNum}) ke PENGAJU ${pengajuEmail} (IKA: ${noReg})`);
            
            sendClosingEmailToPengaju(
              pengajuEmail, noReg, fileId, `H+${reminderDayToSend}`, 
              customSubject // <-- Kirim subjek kustom
            );
            
          } else if (statusIzin === "Proses Penutupan") {
            // Bottleneck = PEMILIK AREA (Per Poin 4)
            const customSubject = `[REMINDER #${reminderNum} - KONFIRMASI PENUTUPAN IKA] - ${noReg}`;
            
            try {
              const ss = SpreadsheetApp.openById(fileId);
              const sheet = ss.getSheetByName("IKA");
              const pemilikAreaEmails = extractEmails(sheet.getRange(CELL_CLOSING_APPROVER).getValue());
              
              if (pemilikAreaEmails.length > 0) {
                pemilikAreaEmails.forEach(email => {
                  Logger.log(`Mengirim reminder closing H+${reminderDayToSend} (Reminder #${reminderNum}) ke PEMILIK AREA ${email} (IKA: ${noReg})`);
                  
                  // Panggil helper yang TEPAT (sendClosingConfirmationEmail)
                  sendClosingConfirmationEmail(
                    email, noReg, pengajuEmail, dept, fileId, 
                    customSubject // <-- Kirim subjek kustom
                  );
                });
              }
            } catch (e) {
              Logger.log(`Gagal 'openById' untuk reminder closing ${noReg}: ${e.message}`);
            }
          }
          
          sheet.getRange(targetRow, colToUpdate).setValue("Terkirim");
        }
      } catch (e) {
        Logger.log(`Error saat memproses reminder closing (v13) untuk baris ${i + 1}: ${e.message}`);
        sheet.getRange(i + 1, CLOSING_H1_COL).setValue("Error: " + e.message.substring(0, 50)); 
      }

      // Failsafe 5 Menit
      const currentTime = new Date().getTime();
      if (currentTime - startTime > TIME_LIMIT_MS) {
        Logger.log(`Batas waktu 5 menit tercapai (Closing). Menghentikan pindaian untuk ${sheetName}.`);
        break; 
      }
    } // Akhir loop for
  }); 
  Logger.log("Pindaian Reminder Closing (v13) Selesai.");
}

/**
 * =====================================================================================
 * FUNGSI REMINDER APPROVAL (V11 - HYBRID)
 * DIPICU 1x SEHARI (08:00)
 * Menggunakan MASTER TIMER (U) dan MASTER STATE (Y)
 * Tapi menggunakan CHECKLIST APPROVAL terpisah (Kolom V, W, X)
 * =====================================================================================
 */
function checkApprovalReminders() {
  const logbook = SpreadsheetApp.openById(LOGBOOK_FILE_ID);
  const now = new Date();
  const startTime = new Date().getTime();
  const TIME_LIMIT_MS = 300000; // Failsafe 5 menit
  const timezone = logbook.getSpreadsheetTimeZone();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth();
  Logger.log("Memulai Pindaian Reminder Approval (v11)...");
  
  const sheetsToProcessNames = [String(currentYear)];
  if (currentMonth < 6) { 
    sheetsToProcessNames.push(String(currentYear - 1));
  }

  sheetsToProcessNames.forEach(sheetName => {
    const sheet = logbook.getSheetByName(sheetName);
    if (!sheet) return;
    
    // Ambil data sampai Kolom Y (25) untuk mendapatkan semua data reminder kita
    const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), 25);
    const data = dataRange.getValues();
    
    // Indeks Kolom (berbasis 0)
    const STATUS_COL = 14; // O
    const LINK_COL = 16; // Q
    const NO_REG_COL = 2; // C
    
    // Kolom Bersama
    const TIMER_COL = MASTER_TIMESTAMP_COL - 1; // U (indeks 20)
    const STATE_COL = MASTER_STATE_COL - 1;     // Y (indeks 24)

    // Kolom Khusus Approval
    const H1_COL = APPROVAL_H1_COL - 1; // V (indeks 21)
    const H2_COL = APPROVAL_H2_COL - 1; // W (indeks 22)
    const H3_COL = APPROVAL_H3_COL - 1; // X (indeks 23)

    for (let i = 8; i < data.length; i++) { // Mulai dari baris 9
      try {
        const statusIzin = data[i][STATUS_COL];
        
        // --- FILTER 1: HANYA JALANKAN UNTUK STATUS "Approval" ---
        if (statusIzin !== "Approval") {
          continue;
        }

        const timestamp = data[i][TIMER_COL]; // Ambil Timer (U)
        if (!(timestamp instanceof Date)) {
          continue; // Timernya belum diset, abaikan
        }

        const diffMs = now.getTime() - timestamp.getTime();
        const daysPassed = Math.floor(diffMs / (1000 * 60 * 60 * 24));

        // --- FILTER 2: 30 HARI ---
        if (daysPassed > 30) {
          continue;
        }

        // --- PANGGILAN "MAHAL" (DETEKTIF) ---
        const link = data[i][LINK_COL];
        const noReg = data[i][NO_REG_COL];
        const fileIdMatch = link.match(/[-\w]{25,}/);
        if (!fileIdMatch) continue;
        const fileId = fileIdMatch[0];
        
        const result = findCurrentApprovers(fileId, noReg); // (Gunakan v5 dari 'findCurrentApprovers')
        
        if (!result || !result.pendingEmails || result.pendingEmails.length === 0) {
          continue; // Tidak ada yang menahan
        }
        
        const currentLayer = result.pendingLayer; 
        const lastLayer = data[i][STATE_COL]; 
        
        const targetRow = i + 1; // Baris aktual di sheet

        // --- LOGIKA PERUBAHAN LAYER (GRACE PERIOD) ---
        if (currentLayer !== lastLayer) {
          Logger.log(`Layer berubah untuk ${noReg}. Dari: ${lastLayer} -> Ke: ${currentLayer}`);
          // Panggil helper V11 untuk me-reset timer dan SEMUA checklist
          resetMasterTimer(noReg, now, currentLayer);
          continue; // Beri masa tenggang 1 hari
        }

        // --- JIKA LAYER SAMA (Kirim Reminder) ---
        
        // --- FILTER MASA TENGGANG ---
        if (daysPassed < 1) {
          continue; // Belum 1 hari sejak timer dimulai/di-reset
        }

        // Baca checklist APPROVAL dari array 'data'
        const h1_status = data[i][H1_COL];
        const h2_status = data[i][H2_COL];
        const h3_status = data[i][H3_COL];
        let reminderDayToSend = 0;
        let colToUpdate = -1;

        if (daysPassed >= 3 && h3_status !== "Terkirim") {
          reminderDayToSend = 3; colToUpdate = APPROVAL_H3_COL;
        } else if (daysPassed >= 2 && h2_status !== "Terkirim") {
          reminderDayToSend = 2; colToUpdate = APPROVAL_H2_COL;
        } else if (daysPassed >= 1 && h1_status !== "Terkirim") {
          reminderDayToSend = 1; colToUpdate = APPROVAL_H1_COL;
        }

        if (reminderDayToSend > 0) {
          const approversObject = { [result.pendingLayer]: result.pendingEmails };
          Logger.log(`Mengirim reminder APPROVAL H+${reminderDayToSend} (Layer ${result.pendingLayer}) untuk ${noReg}...`);
          
          sendApprovalEmail(
            fileId, noReg, result.pengajuEmail, result.dept,
            approversObject, result.pendingLayer,
            `[REMINDER #${reminderDayToSend} - APPROVAL IKA] - ${noReg}`
          );
          
          // Catat bahwa reminder telah terkirim
          sheet.getRange(targetRow, colToUpdate).setValue("Terkirim");
        }
      } catch (e) {
        Logger.log(`Error saat memproses reminder approval (v11) untuk baris ${i + 1}: ${e.message}`);
        // Catat error di checklist H+1
        sheet.getRange(i + 1, APPROVAL_H1_COL).setValue("Error: " + e.message.substring(0, 50)); 
      }

      // --- Failsafe 5 Menit ---
      const currentTime = new Date().getTime();
      if (currentTime - startTime > TIME_LIMIT_MS) {
        Logger.log(`Batas waktu 5 menit tercapai (Approval). Menghentikan pindaian untuk ${sheetName}.`);
        break; 
      }
    } // Akhir loop for
  }); 
  Logger.log("Pindaian Reminder Approval (v11) Selesai.");
}

/**
 * Fungsi ini berjalan otomatis setiap kali ada sel yang diedit.
 * Tujuannya untuk menangani logika checkbox perpanjangan.
 * @param {Object} e Objek event yang disediakan oleh Google Apps Script.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const cellAddress = range.getA1Notation();
 
  // Hanya jalankan jika editan terjadi di sheet "IKA"
  if (sheet.getName() !== "IKA") {
    return;
  }
  
  // Hanya jalankan jika editan ada di sel yang kita pantau
  if (cellAddress === CELL_PEKERJAAN_BARU_CHECK || cellAddress === CELL_PERPANJANGAN_CHECK || cellAddress === CELL_NO_IKA_LAMA) {
    
    // Jika I12 (Pekerjaan Baru) yang dicentang
    if (cellAddress === CELL_PEKERJAAN_BARU_CHECK && range.getValue() === true) {
      sheet.getRange(CELL_PERPANJANGAN_CHECK).setValue(false);
      sheet.getRange(CELL_NO_IKA_LAMA).setBackground(null).clearContent();
    }
    
    // Jika J13 (Perpanjangan) yang diubah
    if (cellAddress === CELL_PERPANJANGAN_CHECK) {
      if (range.getValue() === true) {
        sheet.getRange(CELL_PEKERJAAN_BARU_CHECK).setValue(false);
        sheet.getRange(CELL_NO_IKA_LAMA).setBackground("#ffcccb");
      } else {
        sheet.getRange(CELL_NO_IKA_LAMA).setBackground(null);
      }
    }

    // Jika I16 (No. IKA Lama) yang diisi
    if (cellAddress === CELL_NO_IKA_LAMA) {
      if (range.getValue() !== "") {
        sheet.getRange(CELL_NO_IKA_LAMA).setBackground(null);
      } 
      else {
        if (sheet.getRange(CELL_PERPANJANGAN_CHECK).getValue() === true) {
          sheet.getRange(CELL_NO_IKA_LAMA).setBackground("#ffcccb");
        }
      }
    }
  }
}

/**
 * Fungsi pembantu untuk mengambil dan membersihkan beberapa alamat email
 * dari satu atau beberapa sel. Bisa menerima nilai tunggal atau array.
 * @param {string|Array<Array<string>>} rawData Data mentah dari .getValue() atau .getValues().
 * @returns {Array<string>} Sebuah array bersih berisi alamat email individual.
 */
function extractEmails(rawData) {
  const emails = [];
  let dataAsArray = [];

  // PERBAIKAN: Cek apakah inputnya adalah array atau bukan.
  if (Array.isArray(rawData)) {
    // Jika dari .getValues(), inputnya sudah berupa array.
    dataAsArray = rawData;
  } else {
    // Jika dari .getValue(), ubah menjadi array agar bisa diproses.
    dataAsArray = [rawData];
  }
  
  // Ubah input menjadi array 1D yang datar.
  const flatData = [].concat.apply([], dataAsArray);

  flatData.forEach(item => {
    if (typeof item === 'string' && item.trim() !== '') {
      // Pisahkan email berdasarkan koma, titik koma, atau spasi.
      const potentialEmails = item.split(/,|;|\s+/); 
      
      potentialEmails.forEach(email => {
        const cleanEmail = email.trim();
        // Pastikan hanya alamat email yang valid yang ditambahkan
        if (cleanEmail && cleanEmail.includes('@')) {
          emails.push(cleanEmail);
        }
      });
    }
  });
  
  return emails;
}

/**
 * Fungsi ini dijalankan manual SATU KALI untuk membuat pemicu (trigger) otomatis.
 */
function createTimeDrivenTrigger() {
  const allTriggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;

  // Hapus SEMUA trigger lama untuk kedua fungsi agar tidak duplikat
  allTriggers.forEach(trigger => {
    const handlerFunction = trigger.getHandlerFunction();
    if (handlerFunction === "updateLogbookFromModifiedFiles" || 
        handlerFunction === "checkApprovalReminders") {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  });
  
  if (deletedCount > 0) {
    Logger.log(`Berhasil menghapus ${deletedCount} trigger lama.`);
  }

  // --- TRIGGER 1: UNTUK CLOSING (2x sehari) ---
  // Buat trigger baru untuk berjalan setiap hari antara jam 6 dan 7 pagi
  ScriptApp.newTrigger("updateLogbookFromModifiedFiles")
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
  
  // Buat trigger baru untuk berjalan setiap hari antara jam 6 dan 7 sore (18:00)
  ScriptApp.newTrigger("updateLogbookFromModifiedFiles")
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();

  // --- TRIGGER 2: UNTUK REMINDER APPROVAL (1x sehari) ---
  ScriptApp.newTrigger("checkApprovalReminders")
    .timeBased()
    .everyDays(1)
    .atHour(8) // Jam 8 pagi
    .create();
  
  Logger.log("✅ Pemicu otomatis berhasil dibuat atau diperbarui (3 trigger baru dibuat).");
}

/**
 * Fungsi pembantu untuk membuat halaman respons HTML yang minimalis.
 * @param {string} message Pesan yang ingin ditampilkan.
 * @param {boolean} isSuccess Apakah ini pesan sukses (true) atau gagal (false).
 * @returns {HtmlOutput} Halaman HTML yang siap ditampilkan.
 */
function createSimpleResponsePage(message, isSuccess = true) {
  const icon = isSuccess ? '✅' : '❌';

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600&display=swap');
        body { font-family: 'Inter', sans-serif; background-color: #f9fafb; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; text-align: center; }
        .card { background-color: white; padding: 32px 40px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
        .icon { font-size: 40px; }
        p { font-size: 18px; color: #1f2937; margin: 16px 0 0 0; font-weight: 600; max-width: 350px; }
      </style>
    </head>
    <body>
      <div class="card">
        <div class="icon">${icon}</div>
        <p>${message}</p>
      </div>
      <script>
        // Opsional: Mencoba menutup tab setelah 5 detik
        setTimeout(function(){ window.close(); }, 5000);
      </script>
    </body>
    </html>
  `;
  return HtmlService.createHtmlOutput(html);
}

/**
 * Fungsi ini dijalankan manual DARI EDITOR untuk mereset semua penomoran dokumen.
 * PERHATIAN: Jalankan fungsi ini hanya jika Anda benar-benar ingin
 * memulai ulang penomoran dari 1 untuk semua departemen.
 */
function resetNumbering() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const allKeys = scriptProperties.getKeys();
    let deletedCount = 0;

    allKeys.forEach(key => {
      // Hapus semua kunci yang dimulai dengan "counter_"
      if (key.startsWith("counter_")) {
        scriptProperties.deleteProperty(key);
        deletedCount++;
      }
    });

    // Gunakan Logger.log sebagai ganti ui.alert untuk memberikan feedback
    Logger.log(`Reset Berhasil: Berhasil menghapus ${deletedCount} data penomoran. Penomoran akan dimulai dari 1 pada permintaan berikutnya.`);

  } catch (e) {
    Logger.log("Gagal mereset penomoran: " + e.message);
  }
}
/**
 * FUNGSI INTI untuk mengatur counter secara manual.
 * Jangan jalankan fungsi ini secara langsung.
 */
function setManualCounter(dept, year, month, lastNumberUsed) {
  const yearMonth = `${year}.${String(month).padStart(2, '0')}`;
  const counterKey = `counter_${dept}_${yearMonth}`;
  
  PropertiesService.getScriptProperties().setProperty(counterKey, lastNumberUsed);
  
  Logger.log(`Penomoran untuk departemen ${dept} pada bulan ${yearMonth} telah diatur secara manual ke ${lastNumberUsed}.`);
}

/**
 * FUNGSI PEMBANTU untuk menjalankan pengaturan manual.
 * Ubah data di dalam fungsi ini, lalu jalankan dari editor.
 */
function jalankanPengaturanManual() {
  // --- GANTI DATA DI BAWAH INI SESUAI KEBUTUHAN ANDA ---
  
  const departemen = "IMS";    // Tulis nama departemen yang sesuai
  const tahun = 2025;         // Tulis tahun (contoh: 2025)
  const bulan = 7;            // Tulis nomor bulan (contoh: 7 untuk Juli)
  const nomorTerakhir = 0;    // Tulis NOMOR TERAKHIR yang sudah terpakai
  
  // --- Jangan ubah baris di bawah ini ---
  setManualCounter(departemen, tahun, bulan, nomorTerakhir);
}

/**
 * FUNGSI INI HANYA UNTUK DIJALANKAN MANUAL JIKA PROSES APPROVAL MACET
 * SETELAH INPUT MANUAL DILAKUKAN DI SHEET RIWAYAT APPROVAL.
 */
function forceCompleteApproval() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const docId = ss.getId();
    const sheet = ss.getSheetByName("IKA");
    const riwayatSheet = ss.getSheetByName("Riwayat Approval");

    if (!sheet || !riwayatSheet) {
      throw new Error("Sheet IKA atau Riwayat Approval tidak ditemukan.");
    }

    // 1. Dapatkan daftar approver L3 yang WAJIB dari sheet IKA
    const requiredApproversL3 = extractEmails(sheet.getRange(RANGE_APPROVER_L3).getValues());
    const uniqueRequiredApproversL3 = [...new Set(requiredApproversL3)];

    // 2. Dapatkan daftar approver L3 yang SUDAH APPROVE dari sheet Riwayat
    const historyData = riwayatSheet.getDataRange().getValues();
    const actualApproversL3 = historyData
      .filter(row => row[2] === 3 && row[3] === 'Approve') // Cek Layer 3 & Tindakan 'Approve'
      .map(row => row[4]); // Ambil emailnya
    const uniqueActualApproversL3 = new Set(actualApproversL3);

    // 3. Bandingkan keduanya
    const allL3HaveApproved = uniqueRequiredApproversL3.every(requiredEmail => uniqueActualApproversL3.has(requiredEmail));

    if (allL3HaveApproved) {
      // --- BLOK FINALISASI PROSES ---
      const noReg = sheet.getRange(CELL_NO_REGISTRASI).getValue();
      const pengajuEmail = sheet.getRange(CELL_PENGAJU_EMAIL).getValue();

      updateStatusInLogbook(noReg, "Approved");
      riwayatSheet.getRange("I1").setValue("Approved");
      riwayatSheet.appendRow([riwayatSheet.getLastRow(), new Date(), 3, "Approval Selesai (Manual Trigger)", "System", "Proses diselesaikan oleh admin."]);

      // Mengunci file kembali
      const file = DriveApp.getFileById(docId);
      const editors = file.getEditors();
      const owner = file.getOwner();
      editors.forEach(editor => {
        if (owner && editor.getEmail() === owner.getEmail()) return;
        try { file.removeEditor(editor.getEmail()); } catch (e) { Logger.log(e); }
      });

      // Kirim email notifikasi ke pengaju
      GmailApp.sendEmail(
        pengajuEmail,
        `✅ [APPROVAL IKA SELESAI] - ${noReg}`,
        "",
        {
          htmlBody: `Dokumen IKA Anda dengan nomor registrasi <b>${noReg}</b> telah disetujui sepenuhnya dan prosesnya telah diselesaikan. <br><br>Link Dokumen: <a href="${ss.getUrl()}">Klik di sini</a>.`,
          from: "info.ims@wingscorp.com"
        }
      );

      ui.alert("Sukses!", "Proses approval telah berhasil diselesaikan secara manual.", ui.ButtonSet.OK);

    } else {
      const missingApprovers = uniqueRequiredApproversL3.filter(email => !uniqueActualApproversL3.has(email));
      ui.alert("Belum Selesai", "Sistem mendeteksi masih ada approver Layer 3 yang belum approve: " + missingApprovers.join(", "), ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert("Error", "Terjadi kesalahan: " + e.message, ui.ButtonSet.OK);
    Logger.log(e);
  }
}

// ================================================================
// === FUNGSI MASTER TEMPLATE UNTUK SEMUA EMAIL ===
// ================================================================
function createStyledEmailBody(headerText, headerColor, mainContentHtml, buttonHtml = "") {
  return `
    <div style="border: 1px solid #e2e8f0; border-radius: 8px; font-family: sans-serif; max-width: 600px; margin: auto;">
      <div style="background-color:${headerColor}; color:white; padding:15px; font-size:18px; text-align:center; border-radius: 8px 8px 0 0;">
          <b>${headerText}</b>
      </div>
      <div style="padding: 20px; line-height: 1.6; color: #333;">
          ${mainContentHtml}
          ${buttonHtml ? `<br>${buttonHtml}` : ''}
      </div>
      <div style="padding: 15px; font-size: 12px; color: #718096; border-top: 1px solid #e2e8f0; text-align: center;">
          <i>Integrated Management System (IMS) Department</i><br>
          <p style="font-size:10px;color:#a0aec0;"><i>Email ini dikirim secara otomatis. Mohon untuk tidak membalas.</i></p>
      </div>
    </div>`;
}

/**
 * Mengirim email notifikasi kepada pengaju saat ada keputusan Approve/Reject L2/L3.
 */
function sendApprovalNotificationToRequester(docId, noReg, pengajuEmail, dept, layer, approverEmail, isApproved, komentar) {
    const action = isApproved ? "disetujui" : "ditolak";
    const subject = isApproved ? `[UPDATE APPROVAL IKA] - ${noReg}` : `[APPROVAL IKA DITOLAK] - ${noReg}`;
    const headerText = isApproved ? `UPDATE APPROVAL IKA` : `APPROVAL IKA DITOLAK`;
    const headerColor = isApproved ? "#0051a2" : "#80002a"; // Biru untuk Approve, Merah untuk Reject
    const komentarDisplay = komentar ? `<b>Komentar:</b> <i>${komentar}</i><br><br>` : '';

    const mainContent = `
        Dear Bapak/Ibu, <br><br>
        Dokumen IKA Anda telah <b>${action}</b> oleh <b>${approverEmail}</b> (<b>Layer ${layer}</b>).<br><br>
        ${komentarDisplay}
        &#128196; <b>Nomor Registrasi:</b> ${noReg}<br>
        &#128193; <b>Link Dokumen:</b> <a href="https://docs.google.com/spreadsheets/d/${docId}" target="_blank" style="color: #0051a2; text-decoration: none;">Klik di sini untuk membuka dokumen</a><br><br>
    `;
    
    const htmlBody = createStyledEmailBody(headerText, headerColor, mainContent);
    
    try {
        GmailApp.sendEmail(pengajuEmail, subject, "", { htmlBody: htmlBody, from: "info.ims@wingscorp.com" });
        Logger.log(`Email notifikasi ${action} berhasil dikirim ke Pengaju: ${pengajuEmail}.`);
    } catch (e) {
        Logger.log(`Gagal mengirim notifikasi ke Pengaju ${pengajuEmail}: ${e.message}`);
    }
}

/**
 * =====================================================================================
 * FUNGSI MANUAL TRIGER CLOSING
 * =====================================================================================
 * Fungsi ini digunakan untuk memicu pengiriman email "Permintaan Penutupan IKA" secara manual.
 * 1. Buka file "Logbook IKA".
 * 2. Cari baris data IKA yang ingin dikirim ulang email closing-nya.
 * 3. Dari baris tersebut, catat dua informasi:
 * - Nomor Registrasi (contoh: "SMU/IKA/DEPT/2025.09/001")
 * - ID File (bisa didapat dari link di kolom "Link Dokumen")
 * 4. Buka skrip ini dan edit 3 VARIABEL di bawah (di dalam blok "GANTI DATA DI BAWAH INI").
 * - Masukkan ID File ke `idFileTes`.
 * - Masukkan No. Registrasi ke `noRegTes`.
 * - Masukkan email pengguna yang akan menerima email ke `emailPenerimaTes`.
 * 5. Simpan skrip.
 * 6. Di menu dropdown fungsi di atas, pilih "manualTriggerClosing".
 * 7. Klik tombol "▶️ Run".
 * 8. Pengguna akan menerima email "Permintaan Penutupan IKA" yang baru.
 * * =====================================================================================
 */
function manualTriggerClosing() {

  // Ganti dengan ID file IKA yang ingin Anda jadikan target.
  const idFileTes = "1_uZil56d6XeW5w2xJ4JprlxinFoMNogRQ58JnVN5qg8"; // <-- GANTI INI

  // Ganti dengan No. Registrasi yang sesuai dengan file di atas.
  const noRegTes = "SMU/IKA/IMS/2025.09/002"; // <-- GANTI INI

  // Ganti dengan alamat email pengguna yang akan menerima email.
  const emailPenerimaTes = "floreansalsabila.irdana@wingscorp.com"; // <-- GANTI INI

  // --- Jangan ubah kode di bawah ini ---
  Logger.log(`Memicu pengiriman email closing manual untuk IKA ${noRegTes} ke ${emailPenerimaTes}...`);

  // Memanggil fungsi pengiriman email closing secara langsung
  sendClosingEmailToPengaju(emailPenerimaTes, noRegTes, idFileTes, "MANUAL_RESEND");

  SpreadsheetApp.getUi().alert("Trigger manual untuk email closing berhasil dijalankan!");
  Logger.log("Email berhasil dikirim.");
}


/**
 * =====================================================================================
 * FUNGSI KIRIM ULANG EMAIL APPROVAL
 * =====================================================================================
 * Fungsi ini digunakan untuk mengirim ulang email persetujuan (approval) ke satu approver spesifik
 * 1. Dapatkan informasi dari pengguna:
 * - ID File IKA 
 * - Layer Approver (1, 2, atau 3)
 * - Alamat email approver yang akan dikirimi email ulang.
 * 2. Buka skrip ini dan edit 3 VARIABEL di bawah (di dalam blok "GANTI DATA DI BAWAH INI").
 * 3. Simpan skrip.
 * 4. Di menu dropdown fungsi di atas, pilih "kirimUlangEmailApproval".
 * 5. Klik tombol "▶️ Run".
 * * =====================================================================================
 */
function kirimUlangEmailApproval() {

  // Ganti dengan ID file IKA yang emailnya akan dikirim ulang.
  const idFileTes = "1lmxnlFeYxx234_fUml7aWZuSRtmKX7Yc0EcfDYm2g58"; // <-- GANTI INI

  // Ganti dengan nomor Layer (1, 2, atau 3) yang akan dikirim ulang.
  const layerTes = 2; // <-- GANTI INI

  // Ganti dengan alamat email approver yang akan menerima email.
  const emailPenerimaTes = "fajarnur.muhammad@wingscorp.com"; // <-- GANTI INI

  // --- Jangan ubah kode di bawah ini ---
  try {
    Logger.log(`Memicu pengiriman ulang email approval Layer ${layerTes} untuk file ${idFileTes} ke ${emailPenerimaTes}...`);

    const ss = SpreadsheetApp.openById(idFileTes);
    const sheet = ss.getSheetByName("IKA");
    if (!sheet) {
      throw new Error("Sheet IKA tidak ditemukan di file target.");
    }

    const noReg = sheet.getRange(CELL_NO_REGISTRASI).getValue();
    const pengajuEmail = sheet.getRange(CELL_PENGAJU_EMAIL).getValue();
    const dept = noReg.split('/')[2];
    
    // Kita membuat objek 'approvers' palsu yang hanya berisi data yang kita butuhkan
    // agar fungsi sendApprovalEmail bisa berjalan dengan benar untuk satu penerima.
    const approvers = {
      [layerTes]: [emailPenerimaTes]
    };

    const customSubject = `[MANUAL - APPROVAL IKA] - ${noReg}`;

    // Memanggil fungsi pengiriman email approval yang sudah ada
    sendApprovalEmail(
      idFileTes, noReg, pengajuEmail, dept, 
      approvers, layerTes, 
      customSubject // <-- Oper Subjek Kustom
    );

    SpreadsheetApp.getUi().alert("Email approval berhasil dikirim ulang ke: " + emailPenerimaTes);
    Logger.log("Email berhasil dikirim ulang.");

  } catch (e) {
    Logger.log(`Gagal mengirim ulang email: ${e.stack}`);
    SpreadsheetApp.getUi().alert("Gagal mengirim ulang email. Cek Logs untuk detail error.");
  }
}

/**
 * FUNGSI DEBUGGING: Mengirim email ke pengguna aktif untuk menguji design HTML.
 * JALANKAN SECARA MANUAL DARI EDITOR SCRIPT.
 */
function testEmailDesign() {
  const userEmail = Session.getActiveUser().getEmail();
  const testDocId = "ABCDEFG1234567890"; // Dummy ID
  const testNoReg = "SMU/IKA/TEST/2025.10/000";
  const testDept = "TEST";
  const testApprover = "approver.l2.test@wingscorp.com";
  const testKomentar = "Ini adalah komentar uji coba dari Layer 2.";

  // --- OPSI 1: UJI COBA EMAIL APPROVE (LAYER 2) ---
  
  // Panggil fungsi notifikasi approve
  sendApprovalNotificationToRequester(
    testDocId, 
    testNoReg, 
    userEmail, // Kirim ke diri sendiri
    testDept, 
    2, // Layer 2
    testApprover, 
    true, // isApproved = true
    testKomentar
  );

  // --- OPSI 2: UJI COBA EMAIL REJECT (LAYER 3) ---

  // Panggil fungsi notifikasi reject
  sendApprovalNotificationToRequester(
    testDocId, 
    testNoReg, 
    userEmail, // Kirim ke diri sendiri
    testDept, 
    3, // Layer 3
    testApprover, 
    false, // isApproved = false
    testKomentar
  );

  Logger.log(`Email uji coba notifikasi Approve (L2) dan Reject (L3) telah dikirim ke ${userEmail}.`);
}

function debugAccess() {
  const file = DriveApp.getFileById("17i79FlZIFw8iEh2xBWKzPt4PxjYDQEhardtgydyRvDQ");
  Logger.log(file.getName());
}

// Fungsi ini HANYA untuk testing manual
function testRequestApproval() {
  // Masukkan ID Spreadsheet yang ingin kamu tes di sini
  // Ambil dari URL file IKA contoh: https://docs.google.com/spreadsheets/d/[ID_DI_SINI]/edit
  const idContoh = "1piSOpv9hhmmKssRinBJVNoLvulWJiHOE38rTdkmCtQ4"; 
  
  // Panggil fungsi asli dengan ID contoh
  requestApproval(idContoh);
}
