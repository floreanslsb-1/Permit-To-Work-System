// --- KONFIGURASI ---
const LOGBOOK_ID = "196wV40IgrLAVIssGAZXASyuYV3NqOU9w5CwDsNBO-wk"; // ID Logbook IKA Anda
const SCRIPT_URL = "https://script.google.com/a/macros/wingscorp.com/s/AKfycbxwQZ9MRp5rVDSj2gE5OIAHfz1Gs3y3VabjAar76QydRUfbFJqr-8CggMCH_G-A_4PW/exec"; // Ganti dengan URL Web App Anda setelah deploy

// Lokasi sel penting di dalam file IKA
const CELL_PENGAJU_EMAIL = "E96";
const RANGE_APPROVER_L2 = "E99:E101";
const RANGE_APPROVER_L3 = "E104:E105";
const CELL_NO_REGISTRASI = "J19";

// --- FUNGSI UTAMA ---

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Kirim Ulang Email Approval')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * LANGKAH 3-6: Mencari IKA dan me-list approver yang gilirannya.
 * @param {string} regNumber Nomor registrasi IKA yang akan dicari.
 * @returns {object} Objek berisi status sukses/gagal dan daftar approver.
*/
function findPendingApprovers(regNumber) {
  try {
    if (!regNumber || regNumber.trim() === "") {
      return { success: false, message: "Error: Mohon masukkan Nomor Registrasi IKA." };
    }

    const ikaData = findIkaInLogbook(regNumber.trim());
    if (!ikaData) {
      return { success: false, message: "Error: Nomor Registrasi tidak ditemukan di Logbook." };
    }
    if (ikaData.status !== "Approval") {
      return { success: false, message: `Error: Status IKA saat ini adalah "${ikaData.status}", bukan dalam proses approval.` };
    }

    const ss = SpreadsheetApp.openById(ikaData.fileId);
    const sheet = ss.getSheetByName("IKA");
    const riwayatSheet = ss.getSheetByName("Riwayat Approval");
    if (!riwayatSheet) {
      return { success: false, message: "Error: Sheet 'Riwayat Approval' tidak ditemukan." };
    }
    
    const { pendingLayer, pendingApprovers } = determinePendingApprovers(sheet, riwayatSheet);

    if (pendingLayer === null || pendingApprovers.length === 0) {
      return { success: false, message: "Info: Semua approver di layer saat ini sudah memberikan persetujuan. Menunggu langkah selanjutnya." };
    }

    // Jika berhasil, kembalikan daftar approver dan data IKA
    return { 
      success: true, 
      pendingApprovers: pendingApprovers,
      ikaData: ikaData // Kirim data IKA untuk digunakan di langkah selanjutnya
    };

  } catch (e) {
    Logger.log(e);
    return { success: false, message: "Error: Terjadi kesalahan internal. Silakan hubungi administrator." };
  }
}

/**
 * LANGKAH 8: Mengirim email ke approver yang dipilih.
 * @param {object} ikaData Objek data IKA dari fungsi sebelumnya.
 * @param {string[]} selectedEmails Array berisi email approver yang dipilih.
 * @returns {string} Pesan sukses.
*/
function resendToSelectedApprovers(ikaData, selectedEmails) {
  if (!selectedEmails || selectedEmails.length === 0) {
    return "Tidak ada approver yang dipilih.";
  }
  
  const ss = SpreadsheetApp.openById(ikaData.fileId);
  const sheet = ss.getSheetByName("IKA");
  const { pendingLayer } = determinePendingApprovers(sheet, ss.getSheetByName("Riwayat Approval"));

  selectedEmails.forEach(email => {
    // Kirim email satu per satu
    sendApprovalEmail(ikaData.fileId, ikaData.noReg, ikaData.pengaju, ikaData.departemen, { [pendingLayer]: [email] }, pendingLayer);
  });
  
  return `Sukses! Email approval telah dikirim ulang ke ${selectedEmails.length} approver.`;
}


// --- [MODIFIKASI] FUNGSI UNTUK KIRIM ULANG EMAIL CLOSING ---

/**
 * FUNGSI BARU (Bagian 1):
 * HANYA mencari dan memvalidasi IKA untuk proses closing.
 * Tidak mengirim email.
 * @param {string} regNumber Nomor registrasi IKA yang akan dicari.
 * @returns {object} Objek berisi status sukses/gagal dan data IKA jika sukses.
*/
function findIkaForClosing(regNumber) {
  try {
    if (!regNumber || regNumber.trim() === "") {
      return { success: false, message: "Error: Mohon masukkan Nomor Registrasi IKA." };
    }

    const ikaData = findIkaInLogbook(regNumber.trim());
    if (!ikaData) {
      return { success: false, message: "Error: Nomor Registrasi tidak ditemukan di Logbook." };
    }

    const validStatus = ["Approved", "Expired", "Proses Penutupan"];
    if (!validStatus.includes(ikaData.status)) {
      return { success: false, message: `Error: Status IKA saat ini adalah "${ikaData.status}".` };
    }

    // --- LOGIKA EKSTRAKSI EMAIL ASLI ---
    const ss = SpreadsheetApp.openById(ikaData.fileId);
    const sheet = ss.getSheetByName("IKA");
    let emailTujuan = "";
    let rolePenerima = "";

    if (ikaData.status === "Proses Penutupan") {
      // Ambil email Pemilik Area dari E117
      const rawE117 = sheet.getRange("E117").getValue();
      emailTujuan = extractEmails(rawE117).join(", ");
      rolePenerima = "Pemilik Area";
    } else {
      // Ambil email Pengaju dari E96
      const rawE96 = sheet.getRange("E96").getValue();
      emailTujuan = extractEmails(rawE96).join(", ");
      rolePenerima = "Pengaju";
    }

    // Tambahkan data email asli ke objek ikaData
    ikaData.emailTujuan = emailTujuan || "Email tidak ditemukan di file!";
    ikaData.rolePenerima = rolePenerima;

    return { 
      success: true, 
      ikaData: ikaData 
    };

  } catch (e) {
    Logger.log(e);
    return { success: false, message: "Error: " + e.message };
  }
}

/**
 * FUNGSI BARU (Bagian 2):
 * Mengirim email closing berdasarkan status IKA.
 * Jika status "Proses Penutupan", email dikirim ke Pemilik Area (E117).
 */
function sendClosingEmail(ikaData) {
  try {
    if (!ikaData || !ikaData.fileId) {
      return { success: false, message: "Error: Data IKA tidak valid atau hilang." };
    }
    
    // Buka spreadsheet IKA untuk mengambil data terbaru dari sel E117
    const ss = SpreadsheetApp.openById(ikaData.fileId);
    const sheet = ss.getSheetByName("IKA");

    if (ikaData.status === "Proses Penutupan") {
      // --- ALUR PEMILIK AREA (Sesuai Skrip Utama) ---
      const rawApproverData = sheet.getRange("E117").getValue();
      const approverEmails = extractEmails(rawApproverData);

      if (approverEmails.length === 0) {
        return { success: false, message: "Error: Email Pemilik Area di sel E117 kosong." };
      }

      const subject = `[MANUAL - KONFIRMASI PENUTUPAN IKA] - ${ikaData.noReg}`;
      
      approverEmails.forEach(email => {
        // Gunakan fungsi helper sendClosingConfirmationEmail
        sendClosingConfirmationEmail(
          email, 
          ikaData.noReg, 
          ikaData.pengaju, 
          ikaData.departemen, 
          ikaData.fileId, 
          subject
        );
      });

      return { 
        success: true, 
        message: `Sukses! Email konfirmasi penutupan dikirim ulang ke Pemilik Area: ${approverEmails.join(", ")}.` 
      };

    } else {
      // --- ALUR PENGAJU (Approved/Expired) ---
      sendClosingEmailToPengaju(ikaData.pengaju, ikaData.noReg, ikaData.fileId, "MANUAL_RESEND");
      
      return { 
        success: true, 
        message: `Sukses! Email permintaan penutupan dikirim ulang ke Pengaju: ${ikaData.pengaju}.` 
      };
    }

  } catch (e) {
    Logger.log(e);
    return { success: false, message: "Error: " + e.message };
  }
}


// --- FUNGSI PENCARIAN & PENENTU APPROVER ---

function findIkaInLogbook(regNumber) {
  const logbook = SpreadsheetApp.openById(LOGBOOK_ID);
  const sheets = logbook.getSheets();
  for (let sheet of sheets) {
    const dataRange = sheet.getRange("C9:Q" + sheet.getLastRow());
    const data = dataRange.getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === regNumber) {
        const fileIdMatch = data[i][14].match(/[-\w]{25,}/);
        if (fileIdMatch) {
          return {
            noReg: data[i][0],
            pengaju: data[i][3],
            departemen: data[i][4],
            status: data[i][12],
            fileId: fileIdMatch[0]
          };
        }
      }
    }
  }
  return null;
}

function determinePendingApprovers(sheet, riwayatSheet) {
  const riwayatData = riwayatSheet.getDataRange().getValues();
  const allApprovers = {
    1: extractEmails(sheet.getRange(CELL_PENGAJU_EMAIL).getValue()),
    2: extractEmails(sheet.getRange(RANGE_APPROVER_L2).getValues()),
    3: extractEmails(sheet.getRange(RANGE_APPROVER_L3).getValues())
  };

  const approvedBy = { 1: new Set(), 2: new Set(), 3: new Set() };
  riwayatData.forEach(row => {
    if (row[3] === 'Approve' || row[3] === 'Konfirmasi & Lanjutkan') {
      approvedBy[row[2]].add(row[4]);
    }
  });

  if (approvedBy[1].size < allApprovers[1].length) return { pendingLayer: 1, pendingApprovers: allApprovers[1] };
  
  if (approvedBy[2].size < allApprovers[2].length) {
    const pending = allApprovers[2].filter(email => !approvedBy[2].has(email));
    return { pendingLayer: 2, pendingApprovers: pending };
  }

  if (approvedBy[3].size < allApprovers[3].length) {
    const pending = allApprovers[3].filter(email => !approvedBy[3].has(email));
    return { pendingLayer: 3, pendingApprovers: pending };
  }

  return { pendingLayer: null, pendingApprovers: [] };
}


// --- FUNGSI PEMBANTU (Helper) ---
// ⚠️ PENTING: Salin 3 fungsi di bawah ini dari skrip IKA utama Anda
// dan tempelkan di sini.

function sendApprovalEmail(fileId, noReg, pengajuEmail, dept, approvers, layer) {
  const rawEmailData = approvers[layer];
  const approverEmails = extractEmails(rawEmailData);

  if (!approverEmails || approverEmails.length === 0) {
    Logger.log(`Tidak ada approver yang ditemukan untuk Layer ${layer}.`);
    return;
  }

  const subject = `[MANUAL - APPROVAL IKA] - ${noReg}`;
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
        // --- AKHIR PERUBAHAN ---
      }
      
      const htmlBody = createStyledEmailBody(headerText, headerColor, mainContent, buttonHtml);
      GmailApp.sendEmail(approverEmail, subject, "", { htmlBody: htmlBody, from: "info.ims@wingscorp.com" });
    }
  });
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

// Fungsi ini mengirim email ke Pengaju untuk memulai proses closing
function sendClosingEmailToPengaju(userEmail, noReg, fileId, notifType) {
  const subject = `[MANUAL - PERMINTAAN PENUTUPAN IKA] - ${noReg}`;
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

function sendClosingConfirmationEmail(approverEmail, noReg, pengajuEmail, dept, fileId, customSubject = "") {
  const subject = customSubject || `[MANUAL - KONFIRMASI PENUTUPAN IKA] - ${noReg}`;
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
  
  // PENTING: SCRIPT_URL di sini harus URL Web App SKRIP UTAMA agar tombolnya berfungsi
  const SCRIPT_UTAMA_URL = "https://script.google.com/a/macros/wingscorp.com/s/AKfycbxwQZ9MRp5rVDSj2gE5OIAHfz1Gs3y3VabjAar76QydRUfbFJqr-8CggMCH_G-A_4PW/exec";

  const buttonHtml = `
        <div style="text-align: center;">
          <form method="POST" action="${SCRIPT_UTAMA_URL}" style="display: inline-block;">
            <input type="hidden" name="action" value="finalizeClosing">
            <input type="hidden" name="docId" value="${fileId}">
            <input type="hidden" name="approverEmail" value="${approverEmail}">
            <button type="submit" style="background:#004d40;color:white;padding:12px 28px;border:none;border-radius:5px;cursor:pointer;font-size:16px;font-weight:bold;width:100%; box-sizing: border-box;">✅ Konfirmasi Penutupan</button>
          </form>
        </div>`;
        
  const htmlBody = createStyledEmailBody(headerText, headerColor, mainContent, buttonHtml);
  GmailApp.sendEmail(approverEmail, subject, "", { htmlBody: htmlBody, from: "info.ims@wingscorp.com" });
}

function pancingIzinBaru() {
  // Panggil layanan yang baru kita tambahkan di JSON
  const sisa = MailApp.getRemainingDailyQuota();
  const files = Drive.Files.list({maxResults: 1});
  console.log("Izin berhasil dipancing");
}
