// === KONFIGURASI ===
const LOGBOOK_ID = "196wV40IgrLAVIssGAZXASyuYV3NqOU9w5CwDsNBO-wk"; 
const SCRIPT_URL = "https://script.google.com/a/macros/wingscorp.com/s/AKfycbzbFE5c4vuZRywIXdFAWHbYoRSXI3cdUcPZplsRFShA1PTPPFOs7cKE1BtotK7LM1Cm/exec"; 

const CELL_PENGAJU_EMAIL = "E96";
const CELL_NO_REGISTRASI = "J19";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('HandoverUI')
    .setTitle('Tool Pergantian PIC IKA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Validasi awal: Cek No Reg dan Status Dokumen
 */
function findIkaForHandover(regNumber) {
  try {
    if (!regNumber) return { success: false, message: "Nomor Registrasi wajib diisi." };

    const ikaData = findIkaInLogbook(regNumber.trim());
    if (!ikaData) return { success: false, message: "Nomor Registrasi tidak ditemukan di Logbook." };

    // Validasi Status: Hanya Approved atau Expired
    const allowedStatus = ["Approved", "Expired"];
    if (!allowedStatus.includes(ikaData.status)) {
      return { 
        success: false, 
        message: `Ditolak: Perubahan PIC hanya bisa untuk status Approved/Expired. (Status saat ini: ${ikaData.status})` 
      };
    }

    return { success: true, ikaData: ikaData };
  } catch (e) {
    return { success: false, message: "Error: " + e.message };
  }
}

/**
 * Eksekusi perubahan PIC (Tanpa akses editor Drive)
 */
function executeHandoverPIC(ikaData, newPicEmail, reason) {
  try {
    const ss = SpreadsheetApp.openById(ikaData.fileId);
    const sheet = ss.getSheetByName("IKA");
    const riwayatSheet = ss.getSheetByName("Riwayat Approval");

    // 1. Ganti email pengaju di sel E96
    sheet.getRange(CELL_PENGAJU_EMAIL).setValue(newPicEmail.trim().toLowerCase());

    // 2. Catat di Riwayat Approval (Audit Trail)
    riwayatSheet.appendRow([
      riwayatSheet.getLastRow(), 
      new Date(), 
      "", 
      "Pergantian Pengaju/PIC Dokumen IKA", 
      newPicEmail.trim().toLowerCase(), 
      `PIC lama: ${ikaData.pengaju}. Alasan: ${reason}`
    ]);

    // 3. Update di Logbook agar sinkron
    updatePengajuInLogbook(ikaData.noReg, newPicEmail.trim().toLowerCase());

    // 4. Kirim email instruksi penutupan ke PIC baru (Admin yang memproses)
    sendClosingEmailToPicBaru(newPicEmail.trim().toLowerCase(), ikaData.noReg, ikaData.fileId);

    return { success: true, message: `Berhasil! PIC IKA ${ikaData.noReg} telah diganti ke ${newPicEmail}.` };
  } catch (e) {
    return { success: false, message: "Gagal eksekusi: " + e.message };
  }
}

// --- HELPERS ---

function findIkaInLogbook(regNumber) {
  const logbook = SpreadsheetApp.openById(LOGBOOK_ID);
  const sheets = logbook.getSheets();
  for (let sheet of sheets) {
    const data = sheet.getRange("C9:Q" + sheet.getLastRow()).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === regNumber) {
        const fileIdMatch = data[i][14].match(/[-\w]{25,}/);
        if (fileIdMatch) {
          return { noReg: data[i][0], pengaju: data[i][3], status: data[i][12], fileId: fileIdMatch[0] };
        }
      }
    }
  }
  return null;
}

function updatePengajuInLogbook(regNumber, newEmail) {
  const logbook = SpreadsheetApp.openById(LOGBOOK_ID);
  const sheets = logbook.getSheets();
  for (let sheet of sheets) {
    const data = sheet.getRange("C9:F" + sheet.getLastRow()).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === regNumber) {
        sheet.getRange(i + 9, 6).setValue(newEmail); // Kolom F
        return;
      }
    }
  }
}

function sendClosingEmailToPicBaru(userEmail, noReg, fileId) {
  const subject = `[SERAH TERIMA PIC IKA] - ${noReg}`;
  const headerColor = "#0051a2"; // Corporate Blue
  
  const mainContent = `
      Dear Bapak/Ibu,<br><br>
      Anda telah didaftarkan sebagai <b>PIC Pengganti</b> untuk dokumen IKA <b>${noReg}</b>.<br><br>
      Jika pekerjaan telah selesai, mohon lakukan konfirmasi penutupan melalui link di bawah ini agar dokumen dapat ditutup secara resmi.
  `;

  const buttonHtml = `
      <div style="text-align: center; margin-top: 20px;">
        <form method="POST" action="${SCRIPT_URL}">
          <input type="hidden" name="action" value="initiateClosing">
          <input type="hidden" name="docId" value="${fileId}">
          <input type="hidden" name="pengajuEmail" value="${userEmail}">
          <button type="submit" style="background:#0051a2; color:white; padding:12px 28px; border:none; border-radius:5px; cursor:pointer; font-size:16px; font-weight:bold;">Close IKA</button>
        </form>
      </div>
  `;
  
  // Gunakan struktur template profesional yang "Full-Width"
  const htmlBody = `
    <div style="border: 1px solid #e2e8f0; border-radius: 8px; font-family: sans-serif; max-width: 600px; margin: auto; overflow: hidden;">
      <div style="background-color:${headerColor}; color:white; padding:20px; font-size:18px; text-align:center;">
          <b>Pembaruan PIC IKA</b>
      </div>
      
      <div style="padding: 30px; line-height: 1.6; color: #333;">
          ${mainContent}
          ${buttonHtml}
      </div>
      
      <div style="padding: 15px; font-size: 12px; color: #718096; border-top: 1px solid #e2e8f0; text-align: center; background-color: #f8fafc;">
          <i>Integrated Management System (IMS) Department</i><br>
          <p style="font-size:10px;color:#a0aec0;"><i>Email ini dikirim secara otomatis. Mohon untuk tidak membalas.</i></p>
      </div>
    </div>
  `;
  
  GmailApp.sendEmail(userEmail, subject, "", { 
    htmlBody: htmlBody, 
    from: "info.ims@wingscorp.com" 
  });
}
