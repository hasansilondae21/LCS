// ═══════════════════════════════════════════════════════════
//  Code.gs — LCS Screenshot Rekap, BPP Unaaha
//  Struktur: Google Sheets (data) + Google Drive (foto)
// ═══════════════════════════════════════════════════════════

// ── KONFIGURASI ─────────────────────────────────────────────
// Ganti nilai di bawah setelah pertama kali deploy:
//   SHEET_ID  → ID Google Spreadsheet Anda
//   FOLDER_ID → ID folder Google Drive tempat foto disimpan
// Cara dapat ID: lihat URL sheets/drive, bagian setelah /d/ atau /folders/

var CONFIG = {
  SHEET_ID  : '1RgNsJ-t791WvpzWA5YUBnUZtxOoiAdvar0kMSjaxXYw',
  FOLDER_ID : '1PMVoDFXhVbAhr-WAtNpnnUDNtpQa6BKU',
  SHEET_NAME: 'DataLCS'
};

// ── DAFTAR PENYULUH ─────────────────────────────────────────
var PENYULUH = [
  'LASARUS, S.P',
  'PAULUS, SP., MM',
  'RASNI, S.P.',
  'NURNININGSI, S.P., M.P',
  'MUH. HASAN AL JUFRI SILONDAE, S.P.',
  'MARHADI MAMIR',
  'HJ. SUBRI WAHIDA, S.Pt',
  'NYOMAN ALIT SUTRIADI, S.P.',
  'YULIAS, S.P',
  'RAHMAYANTI, S.P.',
  'MILLA HASMA JELLINOVARISNA, S.P',
  'NUNUNG ULFAYANTI DJAWIE, SP'
];

// ── ENTRY POINT ─────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('LCS Screenshot — BPP Unaaha')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── INIT: pastikan sheet & header ada ───────────────────────
function getOrCreateSheet_() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  } catch(e) {
    throw new Error('SHEET_ID tidak valid. Periksa konfigurasi Code.gs.');
  }

  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(['ID', 'Nama', 'Tanggal', 'Waktu', 'URL Foto', 'Nama File']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── GET DAFTAR PENYULUH ─────────────────────────────────────
function getDaftarPenyuluh() {
  return PENYULUH;
}

// ── SIMPAN SUBMISSION ───────────────────────────────────────
function simpanData(payload) {
  // payload: { nama, base64, mimeType, fileName }
  try {
    var sheet  = getOrCreateSheet_();
    var folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);

    // Decode base64 → blob → simpan ke Drive
    var blob = Utilities.newBlob(
      Utilities.base64Decode(payload.base64),
      payload.mimeType,
      payload.fileName
    );
    var file    = folder.createFile(blob);
    var fileUrl = file.getUrl();

    // Buat ID unik
    var uid = Utilities.getUuid();

    // Tanggal & waktu WIB (UTC+8)
    var now  = new Date();
    var wib  = new Date(now.getTime() + (8 * 60 * 60 * 1000));
    var tgl  = Utilities.formatDate(wib, 'UTC', 'dd/MM/yyyy');
    var jam  = Utilities.formatDate(wib, 'UTC', 'HH:mm:ss');

    sheet.appendRow([uid, payload.nama, tgl, jam, fileUrl, payload.fileName]);

    return { ok: true, id: uid, tanggal: tgl, waktu: jam };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── AMBIL SEMUA DATA (untuk rekap) ──────────────────────────
function getAllData(filter) {
  try {
    var sheet  = getOrCreateSheet_();
    var rows   = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { ok: true, data: [] };

    var result = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];

      // Normalkan tanggal (bisa Date object atau string)
      var tglStr = (r[2] instanceof Date)
        ? Utilities.formatDate(r[2], 'Asia/Makassar', 'dd/MM/yyyy')
        : String(r[2]).trim();

      var jamStr = (r[3] instanceof Date)
        ? Utilities.formatDate(r[3], 'Asia/Makassar', 'HH:mm:ss')
        : String(r[3]).trim();

      if (filter.nama    && r[1] !== filter.nama)    continue;
      if (filter.tanggal && tglStr !== filter.tanggal.trim()) continue;

      result.push({
        id      : r[0],
        nama    : r[1],
        tanggal : tglStr,
        waktu   : jamStr,
        url     : r[4],
        fileName: r[5]
      });
    }
    return { ok: true, data: result };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── AMBIL RINGKASAN STATUS HARI INI ─────────────────────────
function getStatusHariIni() {
  try {
    var sheet = getOrCreateSheet_();
    var rows  = sheet.getDataRange().getValues();

    var hari = Utilities.formatDate(new Date(), 'Asia/Makassar', 'dd/MM/yyyy');

    var sudah = {};
    for (var i = 1; i < rows.length; i++) {
      var selTgl = rows[i][2];
      // Normalkan: jika berupa Date object, format dulu
      var tglStr = (selTgl instanceof Date)
        ? Utilities.formatDate(selTgl, 'Asia/Makassar', 'dd/MM/yyyy')
        : String(selTgl).trim();
      if (tglStr === hari) sudah[rows[i][1]] = true;
    }

    var status = PENYULUH.map(function(n) {
      return { nama: n, sudah: !!sudah[n] };
    });

    return { ok: true, tanggal: hari, status: status };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── AMBIL GAMBAR SEBAGAI BASE64 (untuk PDF client-side) ─────
function getImageBase64(fileUrl) {
  try {
    // Extract file ID dari URL Drive
    var match = fileUrl.match(/[-\w]{25,}/);
    if (!match) return { ok: false, error: 'URL tidak valid' };
    var fileId = match[0];
    var file   = DriveApp.getFileById(fileId);
    var blob   = file.getBlob();
    var b64    = Utilities.base64Encode(blob.getBytes());
    var mime   = blob.getContentType();
    return { ok: true, base64: b64, mime: mime };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function getDaftarTanggal() {
  try {
    var sheet = getOrCreateSheet_();
    var rows  = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { ok: true, data: [] };

    var seen = {};
    var list = [];
    for (var i = 1; i < rows.length; i++) {
      var tglStr = (rows[i][2] instanceof Date)
        ? Utilities.formatDate(rows[i][2], 'Asia/Makassar', 'dd/MM/yyyy')
        : String(rows[i][2]).trim();
      if (tglStr && !seen[tglStr]) {
        seen[tglStr] = true;
        list.push(tglStr);
      }
    }

    // Urutkan terbaru di atas
    list.sort(function(a, b) {
      var toMs = function(d) {
        var p = d.split('/');
        return new Date(p[2], p[1]-1, p[0]).getTime();
      };
      return toMs(b) - toMs(a);
    });

    return { ok: true, data: list };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── HAPUS SATU DATA (opsional, untuk admin) ─────────────────
function hapusData(uid) {
  try {
    var sheet = getOrCreateSheet_();
    var rows  = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === uid) {
        sheet.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: 'Data tidak ditemukan' };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── VERIFIKASI PIN ADMIN ─────────────────────────────────────
function verifikasiPin(pin) {
  return pin === '2026';
}

// ── HAPUS DATA DENGAN VERIFIKASI PIN ────────────────────────
function hapusDataAdmin(uid, pin) {
  if (!verifikasiPin(pin)) return { ok: false, error: 'PIN salah!' };
  try {
    var sheet = getOrCreateSheet_();
    var rows  = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === uid) {
        // Hapus file dari Drive juga
        try {
          var match = rows[i][4].match(/[-\w]{25,}/);
          if (match) DriveApp.getFileById(match[0]).setTrashed(true);
        } catch(e) {}
        sheet.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: 'Data tidak ditemukan' };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}
