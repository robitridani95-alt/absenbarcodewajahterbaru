/**
 * ============================================================
 *  SISTEM ABSENSI BARCODE SISWA - SMKN RAKIT KULIM
 *  Backend: Google Apps Script (API ONLY, tanpa tampilan HTML)
 *  Database: Google Spreadsheet
 *
 *  Arsitektur:
 *  - File ini di-deploy sebagai Web App (Execute as: Me,
 *    Access: Anyone) dan HANYA berfungsi sebagai API (JSON).
 *  - Frontend (index.html) berjalan TERPISAH (dibuka langsung
 *    di browser / dihosting di tempat lain / Github Pages),
 *    lalu memanggil API ini via fetch(). Ini sengaja dipisah
 *    supaya izin kamera untuk scan barcode tidak terbentur
 *    sandbox iframe Apps Script.
 * ============================================================
 */

// =================== KONFIGURASI ===================
const SHEET_ID          = "GANTI_DENGAN_ID_SPREADSHEET_ANDA"; // ID spreadsheet
const SHEET_SISWA       = "Siswa";
const SHEET_ABSENSI     = "Absensi";
const SHEET_JADWAL      = "Jadwal";
const ZONA_WAKTU        = "Asia/Jakarta";
const ADMIN_PIN         = "202408"; // PIN admin default, ganti sesuai kebutuhan

// Jadwal default (dipakai kalau sheet Jadwal kosong/belum di-setup)
const JADWAL_DEFAULT = {
  Senin : { jamMasuk:"07:00", jamTerlambat:"07:15", jamPulang:"15:00" },
  Selasa: { jamMasuk:"07:00", jamTerlambat:"07:15", jamPulang:"15:00" },
  Rabu  : { jamMasuk:"07:00", jamTerlambat:"07:15", jamPulang:"15:00" },
  Kamis : { jamMasuk:"07:00", jamTerlambat:"07:15", jamPulang:"15:00" },
  Jumat : { jamMasuk:"07:00", jamTerlambat:"07:15", jamPulang:"11:30" }
};
const HARI_URUT = ["Senin","Selasa","Rabu","Kamis","Jumat"];

// =================== ENTRY POINT ===================
function doGet(e) {
  try {
    const action = e.parameter.action;

    switch (action) {
      case "getStudent":
        return jsonResponse(getStudentByBarcode(e.parameter.barcode));
      case "submitAttendance":
        return jsonResponse(submitAttendance(e.parameter.barcode));
      case "getTodayLog":
        return jsonResponse(getTodayLog());
      case "getAllStudents":
        return jsonResponse(getAllStudents());
      case "getJadwal":
        return jsonResponse(getJadwal());
      case "saveJadwal":
        return jsonResponse(saveJadwal(e.parameter.pin, e.parameter.data));
      case "getKelasList":
        return jsonResponse(getKelasList());
      case "getAbsensiByKelas":
        return jsonResponse(getAbsensiByKelas(e.parameter.kelas));
      case "checkPin":
        return jsonResponse({ success: e.parameter.pin === ADMIN_PIN });
      case "ping":
        return jsonResponse({ success: true, message: "API aktif", time: new Date() });
      default:
        return jsonResponse({ success: false, message: "Aksi tidak dikenali." });
    }
  } catch (err) {
    return jsonResponse({ success: false, message: "Error server: " + err.message });
  }
}

// Semua request memakai GET (query string) supaya tidak kena
// CORS preflight (OPTIONS) yang tidak didukung Apps Script.
function doPost(e) {
  return doGet(e);
}

// =================== HELPER RESPONSE ===================
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =================== NORMALISASI LINK FOTO GOOGLE DRIVE ===================
// Menerima link Drive dalam berbagai format (link share, link uc?export=view,
// atau bahkan cuma ID file-nya saja) dan mengembalikan format "thumbnail"
// yang paling stabil untuk ditampilkan langsung di <img>. Link biasa
// (drive.google.com/uc?export=view&id=...) sering diblokir/redirect oleh
// Google saat diakses berkali-kali dari luar, sedangkan format thumbnail jauh
// lebih konsisten untuk kebutuhan hotlink seperti ini.
function normalizeFotoUrl(url) {
  if (!url) return "";
  url = url.toString().trim();
  if (!url) return "";

  let id = null;
  let m = url.match(/\/d\/([a-zA-Z0-9_-]{20,})/);      // .../file/d/ID/view
  if (m) id = m[1];
  if (!id) { m = url.match(/[?&]id=([a-zA-Z0-9_-]{20,})/); if (m) id = m[1]; } // ...?id=ID
  if (!id && /^[a-zA-Z0-9_-]{20,}$/.test(url)) id = url; // hanya ID Drive yang ditempel langsung

  if (id) {
    return "https://drive.google.com/thumbnail?id=" + id + "&sz=w500";
  }
  return url; // bukan link Drive (mis. link foto dari luar) -> pakai apa adanya
}

// =================== AMBIL DATA SISWA ===================
function getStudentByBarcode(barcode) {
  if (!barcode) {
    return { success: false, message: "Barcode kosong." };
  }
  barcode = barcode.toString().trim();

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_SISWA);
  const data = sheet.getDataRange().getValues();
  // Header: [0]ID_Barcode [1]NIS [2]Nama [3]Kelas [4]Foto_URL [5]Status_Aktif

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const idBarcode = row[0] ? row[0].toString().trim() : "";

    if (idBarcode === barcode) {
      const statusAktif = row[5];
      if (statusAktif === false || String(statusAktif).toUpperCase() === "TIDAK") {
        return { success: false, message: "Kartu siswa ini tidak aktif." };
      }
      return {
        success: true,
        nis: row[1],
        nama: row[2],
        kelas: row[3],
        foto: normalizeFotoUrl(row[4]),
        barcode: idBarcode
      };
    }
  }
  return { success: false, message: "Barcode tidak terdaftar di database siswa." };
}

// =================== SIMPAN ABSENSI ===================
function submitAttendance(barcode) {
  const lock = LockService.getScriptLock();
  try {
    // Cegah tabrakan data saat banyak siswa scan bersamaan
    lock.waitLock(15000);

    const siswa = getStudentByBarcode(barcode);
    if (!siswa.success) {
      return siswa; // pesan error sudah sesuai
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetAbsensi = ss.getSheetByName(SHEET_ABSENSI);

    const now = new Date();
    const tanggalHariIni = Utilities.formatDate(now, ZONA_WAKTU, "yyyy-MM-dd");
    const jamSekarang = Utilities.formatDate(now, ZONA_WAKTU, "HH:mm:ss");

    // Cek apakah siswa ini sudah absen hari ini (cegah duplikat)
    const dataAbsensi = sheetAbsensi.getDataRange().getValues();
    for (let i = 1; i < dataAbsensi.length; i++) {
      const row = dataAbsensi[i];
      const tglRow = row[1] ? row[1].toString() : "";
      const barcodeRow = row[3] ? row[3].toString().trim() : "";
      if (tglRow === tanggalHariIni && barcodeRow === siswa.barcode) {
        return {
          success: false,
          duplicate: true,
          message: siswa.nama + " sudah tercatat hadir hari ini pukul " + row[2],
          nama: siswa.nama,
          nis: siswa.nis,
          kelas: siswa.kelas,
          foto: siswa.foto,
          jam: row[2]
        };
      }
    }

    // Tentukan status: Hadir / Terlambat, berdasarkan jadwal hari ini
    const hariIni = getNamaHari(now);
    const jadwalHariIni = getJadwalHari(hariIni);
    const batasTerlambat = jadwalHariIni.jamTerlambat;
    const status = jamSekarang <= (batasTerlambat + ":59") ? "Hadir" : "Terlambat";

    sheetAbsensi.appendRow([
      now,
      tanggalHariIni,
      jamSekarang,
      siswa.barcode,
      siswa.nis,
      siswa.nama,
      siswa.kelas,
      status,
      siswa.foto
    ]);
    // Pastikan kolom tanggal & jam tersimpan sebagai teks, bukan
    // otomatis diubah jadi Date oleh Google Sheets.
    const lastRow = sheetAbsensi.getLastRow();
    sheetAbsensi.getRange(lastRow, 2, 1, 2).setNumberFormat("@");

    return {
      success: true,
      duplicate: false,
      nama: siswa.nama,
      nis: siswa.nis,
      kelas: siswa.kelas,
      foto: siswa.foto,
      status: status,
      jam: jamSekarang,
      message: "Absensi berhasil dicatat."
    };

  } catch (err) {
    return { success: false, message: "Gagal menyimpan absensi: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

// =================== LOG HARI INI (untuk kiosk feed) ===================
function getTodayLog() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_ABSENSI);
  const data = sheet.getDataRange().getValues();
  const tanggalHariIni = Utilities.formatDate(new Date(), ZONA_WAKTU, "yyyy-MM-dd");

  // "log" (untuk kiosk, dibatasi 15 item terbaru) tetap dipertahankan,
  // tapi statistik hadir/terlambat/total dihitung dari SELURUH data
  // hari ini supaya panel admin tidak salah hitung.
  const log = [];
  let totalHadir = 0, totalTerlambat = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (row[1] && row[1].toString() === tanggalHariIni) {
      if (row[7] === "Terlambat") totalTerlambat++; else totalHadir++;
      if (log.length < 15) {
        log.push({
          nama: row[5],
          kelas: row[6],
          jam: row[2],
          status: row[7],
          foto: normalizeFotoUrl(row[8])
        });
      }
    }
  }
  return {
    success: true,
    log: log,
    totalHadir: totalHadir,
    totalTerlambat: totalTerlambat,
    totalScan: totalHadir + totalTerlambat
  };
}

// =================== DAFTAR SEMUA SISWA (untuk halaman cetak barcode) ===================
function getAllStudents() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_SISWA);
  const data = sheet.getDataRange().getValues();
  // Header: [0]ID_Barcode [1]NIS [2]Nama [3]Kelas [4]Foto_URL [5]Status_Aktif

  const siswa = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // lewati baris kosong
    siswa.push({
      barcode: row[0].toString().trim(),
      nis: row[1],
      nama: row[2],
      kelas: row[3],
      foto: normalizeFotoUrl(row[4]),
      aktif: !(row[5] === false || String(row[5]).toUpperCase() === "TIDAK")
    });
  }
  return { success: true, siswa: siswa };
}

// =================== JADWAL (HARI, JAM MASUK/TERLAMBAT/PULANG) ===================
function getNamaHari(date) {
  const hari = ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
  return hari[date.getDay()];
}

// Ambil jadwal satu hari tertentu. Fallback ke default kalau sheet
// belum di-setup atau hari itu tidak ada baris di sheet Jadwal
// (mis. Sabtu/Minggu, dianggap tidak ada sekolah tapi tetap dikasih
// nilai default supaya sistem tidak error kalau ada yang scan).
function getJadwalHari(namaHari) {
  const semua = getJadwal();
  if (semua.jadwal && semua.jadwal[namaHari]) return semua.jadwal[namaHari];
  return JADWAL_DEFAULT[namaHari] || { jamMasuk:"07:00", jamTerlambat:"07:15", jamPulang:"15:00" };
}

function getJadwal() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_JADWAL);
  if (!sheet) return { success: true, jadwal: JADWAL_DEFAULT };

  const data = sheet.getDataRange().getValues();
  // Header: [0]Hari [1]JamMasuk [2]JamTerlambat [3]JamPulang
  const jadwal = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    jadwal[row[0].toString().trim()] = {
      jamMasuk: formatJamCell(row[1]),
      jamTerlambat: formatJamCell(row[2]),
      jamPulang: formatJamCell(row[3])
    };
  }
  // Lengkapi hari yang belum ada barisnya dengan default
  HARI_URUT.forEach(h => { if (!jadwal[h]) jadwal[h] = JADWAL_DEFAULT[h]; });
  return { success: true, jadwal: jadwal };
}

// Sheet bisa saja menyimpan jam sebagai objek Date (kalau diketik manual
// oleh guru di spreadsheet), jadi dinormalisasi ke string "HH:mm".
function formatJamCell(val) {
  if (val instanceof Date) return Utilities.formatDate(val, ZONA_WAKTU, "HH:mm");
  return val ? val.toString().trim() : "";
}

// data: JSON string { Senin:{jamMasuk,jamTerlambat,jamPulang}, ... }
function saveJadwal(pin, dataStr) {
  if (pin !== ADMIN_PIN) {
    return { success: false, message: "PIN admin salah." };
  }
  let parsed;
  try {
    parsed = JSON.parse(dataStr);
  } catch (err) {
    return { success: false, message: "Data jadwal tidak valid." };
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_JADWAL);
  if (!sheet) sheet = ss.insertSheet(SHEET_JADWAL);
  sheet.clear();
  sheet.appendRow(["Hari", "JamMasuk", "JamTerlambat", "JamPulang"]);
  sheet.setFrozenRows(1);

  HARI_URUT.forEach(hari => {
    const j = parsed[hari] || JADWAL_DEFAULT[hari];
    sheet.appendRow([hari, j.jamMasuk, j.jamTerlambat, j.jamPulang]);
  });
  sheet.getRange(2, 2, HARI_URUT.length, 3).setNumberFormat("@"); // simpan sebagai teks

  return { success: true, message: "Jadwal berhasil disimpan." };
}

// =================== DAFTAR KELAS (untuk dropdown admin) ===================
function getKelasList() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_SISWA);
  const data = sheet.getDataRange().getValues();
  const set = {};
  for (let i = 1; i < data.length; i++) {
    const kelas = data[i][3] ? data[i][3].toString().trim() : "";
    if (kelas) set[kelas] = true;
  }
  const kelasList = Object.keys(set).sort();
  return { success: true, kelasList: kelasList };
}

// =================== REKAP ABSENSI PER KELAS (untuk download) ===================
function getAbsensiByKelas(kelas) {
  if (!kelas) return { success: false, message: "Kelas belum dipilih." };
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_ABSENSI);
  const data = sheet.getDataRange().getValues();
  // Header: [0]Timestamp [1]Tanggal [2]Jam [3]ID_Barcode [4]NIS [5]Nama [6]Kelas [7]Status [8]Foto_URL

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[6] && row[6].toString().trim() === kelas.trim()) {
      rows.push({
        tanggal: row[1] ? row[1].toString() : "",
        jam: row[2] ? row[2].toString() : "",
        nis: row[4],
        nama: row[5],
        kelas: row[6],
        status: row[7]
      });
    }
  }
  // Urutkan terbaru dulu
  rows.sort((a, b) => (a.tanggal + a.jam < b.tanggal + b.jam) ? 1 : -1);
  return { success: true, kelas: kelas, data: rows };
}

/**
 * =================== SETUP AWAL SPREADSHEET ===================
 * Jalankan fungsi ini SEKALI dari editor Apps Script untuk
 * membuat struktur sheet otomatis (header + contoh baris).
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  let siswa = ss.getSheetByName(SHEET_SISWA);
  if (!siswa) siswa = ss.insertSheet(SHEET_SISWA);
  siswa.clear();
  siswa.appendRow(["ID_Barcode", "NIS", "Nama", "Kelas", "Foto_URL", "Status_Aktif"]);
  siswa.appendRow(["SISWA001", "2024001", "Contoh Nama Siswa", "X TKJ 1",
    "TEMPEL_ID_ATAU_LINK_FOTO_DARI_DRIVE", "Aktif"]);
  siswa.setFrozenRows(1);

  let absensi = ss.getSheetByName(SHEET_ABSENSI);
  if (!absensi) absensi = ss.insertSheet(SHEET_ABSENSI);
  absensi.clear();
  absensi.appendRow(["Timestamp", "Tanggal", "Jam", "ID_Barcode", "NIS", "Nama", "Kelas", "Status", "Foto_URL"]);
  absensi.setFrozenRows(1);

  let jadwal = ss.getSheetByName(SHEET_JADWAL);
  if (!jadwal) jadwal = ss.insertSheet(SHEET_JADWAL);
  jadwal.clear();
  jadwal.appendRow(["Hari", "JamMasuk", "JamTerlambat", "JamPulang"]);
  HARI_URUT.forEach(hari => {
    const j = JADWAL_DEFAULT[hari];
    jadwal.appendRow([hari, j.jamMasuk, j.jamTerlambat, j.jamPulang]);
  });
  jadwal.getRange(2, 2, HARI_URUT.length, 3).setNumberFormat("@");
  jadwal.setFrozenRows(1);

  SpreadsheetApp.getUi().alert("Setup selesai! Sheet 'Siswa', 'Absensi' & 'Jadwal' sudah siap dipakai.");
}
