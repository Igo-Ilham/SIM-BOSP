/** 
 * Web App: SIMBOSP (Single-Page)
 * Fitur utama: Multi-user login, registrasi dengan approval admin, dashboard laporan PDF,
 * master data (tahun/sumber dana/kategori), struktur folder Drive otomatis,
 * notifikasi email, log aktivitas, session timeout 5 menit.
 */

// --- Setup Config ---
const TZ = "Asia/Jakarta"; // Setting TimeZone APP
const TS_FMT = "yyyy-MM-dd HH:mm:ss"; // Setting Format Date
const SESSION_TTL_SEC = 300; // 5 menit (300 detik) AutoLogOut
const CACHE_PREFIX = "BOSP_SESS_"; // Prefix Setiap Sesi Cache
const APP_FOLDER_NAME = "SIM_BOSP_Manager_Files"; // folder root aplikasi di parent folder Spreadsheet

// --- Sheet names ---
const SH_USER = "user";
const SH_APPROVAL = "approval";
const SH_LAPORAN = "laporan";
const SH_KAT = "kategori_laporan";
const SH_SD = "sumber_dana";
const SH_BULAN = "bulan" ;
const SH_TAHUN = "tahun_anggaran";
const SH_LOG = "log_aktivitas";

// --- Headers / Schemas ---
const HDR_USER = ["timestamp", "role", "status", "jenjang", "negeri_swasta", "sekolah", "email", "password"];
const HDR_APPROVAL = ["timestamp", "status", "jenjang", "negeri_swasta", "sekolah", "email", "password"];
const HDR_LAPORAN = ["id", "timestamp", "owner_email", "sekolah", "tahun", "sumber_dana", "kategori", "bulan", "keterangan", "fileId", "fileUrl", "fileName"];
const HDR_KAT = ["kategori"];
const HDR_SD = ["sumber_dana"];
const HDR_BULAN = ["bulan"];
const HDR_TAHUN = ["tahun"];
const HDR_LOG = ["timestamp", "email", "aksi", "detail"];

// --- Util tanggal ---
function nowStr() {
  // Pastikan selalu disimpan sebagai TEKS agar tidak diubah ke Date oleh Sheets
  return "'" + Utilities.formatDate(new Date(), TZ, TS_FMT);
}

// --- Dapatkan spreadsheet & sheet ---
function SS() { return SpreadsheetApp.getActiveSpreadsheet(); }
function getSheetByName(name) { return SS().getSheetByName(name); }
function getOrCreateSheet(name, headers) {
  let sh = SS().getSheetByName(name);
  if (!sh) {
    sh = SS().insertSheet(name);
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    // pastikan header
    const firstRow = sh.getRange(1,1,1,headers.length).getValues()[0];
    const same = JSON.stringify(firstRow) === JSON.stringify(headers);
    if (!same) {
      sh.clear();
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

// --- Hapus Sheet1 bila ada ---
function removeDefaultSheet() {
  const s = SS().getSheetByName("Sheet1");
  if (s && SS().getSheets().length > 1) {
    SS().deleteSheet(s);
  }
}

// --- Inisialisasi sekali pada saat pertama kali launch ---
function ensureSetup_() {
  getOrCreateSheet(SH_USER, HDR_USER);
  getOrCreateSheet(SH_APPROVAL, HDR_APPROVAL);
  getOrCreateSheet(SH_LAPORAN, HDR_LAPORAN);
  getOrCreateSheet(SH_KAT, HDR_KAT);
  getOrCreateSheet(SH_SD, HDR_SD);
  getOrCreateSheet(SH_BULAN, HDR_BULAN);
  getOrCreateSheet(SH_TAHUN, HDR_TAHUN);
  getOrCreateSheet(SH_LOG, HDR_LOG);
  removeDefaultSheet();

  // Seed akun admin otomatis (email owner, password 'admin123'), jika belum ada
  const ownerEmail = SS().getOwner().getEmail ? SS().getOwner().getEmail() : Session.getEffectiveUser().getEmail();
  const shU = getSheetByName(SH_USER);
  const data = shU.getRange(2,1, shU.getLastRow()-0, HDR_USER.length).getValues().filter(r=>r[0]);
  const already = data.some(r => (r[6]+"").toLowerCase() === (ownerEmail+"").toLowerCase());
  if (!already) {
    const row = [nowStr(), "ADMIN", "AKTIF", "Lainnya", "Swasta", "MASTER", ownerEmail, "admin123"];
    shU.appendRow(row);
    logAktivitas_(ownerEmail, "INIT_ADMIN", "Generate akun ADMIN otomatis untuk owner spreadsheet");
  }

  // Seed master data awal jika kosong
  // Generate Sumber Dana
  const shSD = getSheetByName(SH_SD); 
  if (shSD.getLastRow() < 2) {
    [["BOSP Reguler"], ["BOSP Daerah"], ["BOSP Kinerja"], ["SILPA BOSP Kinerja"], ["BOSP Afirmasi"], ["Lainnya"]].forEach(v => shSD.appendRow(v));
  }
  // Generate Kategori Laporan
  const shK = getSheetByName(SH_KAT);
  if (shK.getLastRow() < 2) {
    [["Rincian Kertas Kerja (Tahunan)"], ["Rincian Kertas Kerja (Tahapan)"], ["Rincian Kertas Kerja (Triwulan)"], ["Rincian Kertas Kerja (Bulanan)"], ["Lembar Kertas Kerja Unit Tahap"], ["Lembar Kertas Kerja Unit Triwulan"], ["Lembar Kertas Kerja (Unit 2.2)"], ["Lembar Kertas Kerja (Unit 2.2.1)"], ["Buku Kas Umum (Bulanan)"], ["Buku Kas Umum (Tahunan)"], ["Buku Kas Pembantu Bank (Bulanan)"], ["Buku Pembantu Pajak (Bulanan)"], ["Buku Kas Pembantu Tunai (Bulanan)"], ["Rekapitulasi Realisasi (Bulanan)"], ["Rekapitulasi Realisasi (Tahapan)"], ["Rekapitulasi Realisasi (Tahunan)"], ["Rekapitulasi Realisasi Barang Habis Pakai (Bulanan)"], ["Rekapitulasi Realisasi Barang Modal/Aset (Bulanan)"], ["Buku Pembantu Rincian Objek Belanja (Bulanan)"], ["SPTJM (Bulanan)"], ["SPTJM (Semester)"], ["Laporan Penggunaan Hibah Dana (Semester)"], ["Laporan Penggunaan Hibah Dana (Tahunan)"], ["Mutasi Rekening Koran (Bulanan)"], ["Register Penutupan Kas (Bulanan)"], ["Bukti Pengeluaran Non Tunai (BNU)"], ["Bukti Pengeluaran Tunai (BPU)"], ["Bukti Penarikan (BBU)"], ["Berita Acara Rekonsiliasi"]].forEach(v => shK.appendRow(v));
  }
  //Generate Bulan
  const shB = getSheetByName(SH_BULAN);
  if (shB.getLastRow() < 2) {
    [["Januari"], ["Februari"], ["Maret"], ["April"], ["Mei"], ["Juni"], ["Juli"], ["Agustus"], ["September"], ["Oktober"], ["November"], ["Desember"]].forEach(v => shB.appendRow(v));
  }
  // Generate Tahun
  const shT = getSheetByName(SH_TAHUN);
  if (shT.getLastRow() < 2) {
    const y = new Date().getFullYear();
    [[String(y-1)], [String(y)], [String(y+1)]].forEach(v => shT.appendRow(v));
  }
  // Pastikan folder root aplikasi ada di parent folder Spreadsheet
  ensureAppRootFolder_();
}

// --- Folder helpers ---
function getSpreadsheetParentFolder_() {
  const file = DriveApp.getFileById(SS().getId());
  const parents = file.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}
// Membuat Root Folder
function ensureAppRootFolder_() {
  const parent = getSpreadsheetParentFolder_();
  const folders = parent.getFoldersByName(APP_FOLDER_NAME);
  if (!folders.hasNext()) {
    parent.createFolder(APP_FOLDER_NAME);
  }
}
// Tarik folder Root Aplikasi
function getAppRootFolder_() {
  ensureAppRootFolder_();
  const parent = getSpreadsheetParentFolder_();
  return parent.getFoldersByName(APP_FOLDER_NAME).next();
}
// Tarik atau buat baru folder nama sekolah
function getOrCreateSchoolFolder_(schoolName) {
  const root = getAppRootFolder_();
  let it = root.getFoldersByName(schoolName);
  if (it.hasNext()) return it.next();
  return root.createFolder(schoolName);
}
// Tarik atau buat baru sub folder kategori
function getOrCreateSubFolder_(parent, name) {
  let it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

// --- Logging ---
function logAktivitas_(email, aksi, detail) {
  const sh = getSheetByName(SH_LOG);
  sh.appendRow([nowStr(), email || "-", aksi, detail || ""]);
}

// --- Session (Cache) ---
function createSession_(email, role, sekolah) {
  const token = Utilities.getUuid();
  const payload = { email, role, sekolah, ts: Date.now() };
  CacheService.getUserCache().put(CACHE_PREFIX + token, JSON.stringify(payload), SESSION_TTL_SEC);
  return token;
}
function getSession_(token) {
  if (!token) return null;
  const raw = CacheService.getUserCache().get(CACHE_PREFIX + token);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch(e) { return null; }
}
function refreshSession_(token) {
  const sess = getSession_(token);
  if (!sess) return false;
  // extend TTL
  CacheService.getUserCache().put(CACHE_PREFIX + token, JSON.stringify(sess), SESSION_TTL_SEC);
  return true;
}
function destroySession_(token) {
  if (token) CacheService.getUserCache().remove(CACHE_PREFIX + token);
}

// --- doGet: render halaman tunggal ---
function doGet() {
  ensureSetup_();
  const t = HtmlService.createTemplateFromFile("Index"); // single HTML file
  const page = t.evaluate()
    .setTitle("SIMBOSP")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  return page;
}

// --- HTML include helper (opsional jika Anda pecah file) ---
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- Data akses & util ---
function getMasterData() {
  const sd = getSheetByName(SH_SD).getRange(2,1,getSheetByName(SH_SD).getLastRow()-1,1).getValues().flat().filter(String);
  const kat = getSheetByName(SH_KAT).getRange(2,1,getSheetByName(SH_KAT).getLastRow()-1,1).getValues().flat().filter(String);
  const bln = getSheetByName(SH_BULAN).getRange(2,1, getSheetByName(SH_BULAN).getLastRow()-1,1).getValues().flat().filter(String);
  const th = getSheetByName(SH_TAHUN).getRange(2,1,getSheetByName(SH_TAHUN).getLastRow()-1,1).getValues().flat().filter(String);
  return { sumberDana: sd, kategori: kat, bulan: bln, tahun: th };
}
// Menambah Master Data Baru
function addMasterItem(type, value, token) {
  const sess = getSession_(token);
  if (!sess || sess.role !== "ADMIN") throw new Error("Unauthorized");
  if (!value || !type) throw new Error("Data tidak lengkap");
  let sh, header;
  if (type === "sumber_dana") { sh = getSheetByName(SH_SD); header = HDR_SD; }
  else if (type === "kategori") { sh = getSheetByName(SH_KAT); header = HDR_KAT; }
  else if (type === "tahun") { sh = getSheetByName(SH_TAHUN); header = HDR_TAHUN; }
  else throw new Error("Tipe master tidak dikenali");
  sh.appendRow([value]);
  logAktivitas_(sess.email, "ADD_MASTER_"+type.toUpperCase(), value);
  return getMasterData();
}
// Menghapus Master Data
function removeMasterItem(type, value, token) {
  const sess = getSession_(token);
  if (!sess || sess.role !== "ADMIN") throw new Error("Unauthorized");
  if (!value || !type) throw new Error("Data tidak lengkap");
  let sh;
  if (type === "sumber_dana") sh = getSheetByName(SH_SD);
  else if (type === "kategori") sh = getSheetByName(SH_KAT);
  else if (type === "tahun") sh = getSheetByName(SH_TAHUN);
  else throw new Error("Tipe master tidak dikenali");

  const last = sh.getLastRow();
  if (last < 2) return getMasterData();
  const vals = sh.getRange(2,1,last-1,1).getValues();
  for (let i=0;i<vals.length;i++) {
    if ((vals[i][0]+"").trim().toLowerCase() === (value+"").trim().toLowerCase()) {
      sh.deleteRow(i+2);
      logAktivitas_(sess.email, "DEL_MASTER_"+type.toUpperCase(), value);
      break;
    }
  }
  return getMasterData();
}

// --- Auth: login & register & approval ---
function login(email, password) {
  ensureSetup_();
  if (!email || !password) throw new Error("Email/password wajib diisi.");
  const sh = getSheetByName(SH_USER);
  const last = sh.getLastRow();
  const vals = last > 1 ? sh.getRange(2,1,last-1,HDR_USER.length).getValues() : [];
  let found = null;
  for (let r of vals) {
    const rEmail = (r[6]+"").trim().toLowerCase();
    const rPass = (r[7]+"").trim();
    if (rEmail === (email+"").trim().toLowerCase() && rPass === password) {
      found = r;
      break;
    }
  }
  if (!found) throw new Error("Email/password tidak cocok atau belum disetujui.");
  if ((found[2]+"").toUpperCase() !== "AKTIF") throw new Error("Akun diblokir/nonaktif. Hubungi admin.");
  const token = createSession_(found[6], found[1], found[5]); // email, role, sekolah
  logAktivitas_(found[6], "LOGIN", "Berhasil login");
  return { token, role: found[1], email: found[6], sekolah: found[5] };
}
// --- Logout ---
function logout(token) {
  const sess = getSession_(token);
  if (sess) logAktivitas_(sess.email, "LOGOUT", "Keluar");
  destroySession_(token);
  return true;
}
// --- Register User Baru ---
function registerUser(payload) {
  ensureSetup_();
  // payload: {nama, jenjang, statusSekolah, sekolah, email, password}
  const { nama, jenjang, statusSekolah, sekolah, email, password } = payload || {};
  if (!nama || !jenjang || !statusSekolah || !sekolah || !email || !password) {
    throw new Error("Semua field registrasi wajib diisi.");
  }
  const shA = getSheetByName(SH_APPROVAL);
  const row = [nowStr(), "PENDING", jenjang, statusSekolah, sekolah, email, password];
  shA.appendRow(row);

  // notif admin
  const adminEmail = SS().getOwner().getEmail ? SS().getOwner().getEmail() : Session.getEffectiveUser().getEmail();
  try {
    MailApp.sendEmail({
      to: adminEmail,
      subject: "[BOSP] Registrasi baru menunggu approval",
      htmlBody: `
        <p>Registrasi baru:</p>
        <ul>
          <li>Nama: ${escapeHtml_(nama)}</li>
          <li>Jenjang: ${escapeHtml_(jenjang)}</li>
          <li>Status: ${escapeHtml_(statusSekolah)}</li>
          <li>Sekolah: ${escapeHtml_(sekolah)}</li>
          <li>Email: ${escapeHtml_(email)}</li>
        </ul>
        <p>Silakan buka Panel Admin → Approval untuk menyetujui pada link : s.id/SIMBOSP.</p>
      `
    });
  } catch(e) {}
  logAktivitas_(email, "REGISTER", `Registrasi ${sekolah} (${jenjang}/${statusSekolah})`);
  return true;
}
// --- Register Menunggu Approval ---
function listPendingApprovals(token) {
  const sess = getSession_(token);
  if (!sess || sess.role !== "ADMIN") throw new Error("Unauthorized");
  const sh = getSheetByName(SH_APPROVAL);
  const last = sh.getLastRow();
  const rows = [];
  if (last > 1) {
    const vals = sh.getRange(2,1,last-1,HDR_APPROVAL.length).getValues();
    vals.forEach((r, idx) => {
      rows.push({
        i: idx+2,
        timestamp: r[0],
        status: r[1],
        jenjang: r[2],
        statusSekolah: r[3],
        sekolah: r[4],
        email: r[5],
        password: r[6]
      });
    });
  }
  return rows;
}
/// --- User approval ---
function approveUser(rowIndex, token) {
  const sess = getSession_(token);
  if (!sess || sess.role !== "ADMIN") throw new Error("Unauthorized");
  const shA = getSheetByName(SH_APPROVAL);
  const r = shA.getRange(rowIndex, 1, 1, HDR_APPROVAL.length).getValues()[0];
  if ((r[1]+"").toUpperCase() !== "PENDING") throw new Error("Data bukan status PENDING.");

  // tulis ke user sheet
  const shU = getSheetByName(SH_USER);
  shU.appendRow([nowStr(), "USER", "AKTIF", r[2], r[3], r[4], r[5], r[6]]);

  // ubah status approval
  shA.getRange(rowIndex, 2).setValue("APPROVED");

  // notif user
  try {
    MailApp.sendEmail({
      to: r[5],
      subject: "[BOSP] Akun Anda telah disetujui",
      htmlBody: `<p>Akun untuk sekolah <b>${escapeHtml_(r[4])}</b> telah <b>Disetujui</b>. Silakan login pada link : s.id/SIMBOSP.</p>`
    });
  } catch(e) {}

  // siapkan folder sekolah (jika belum ada)
  getOrCreateSchoolFolder_(r[4]);

  logAktivitas_(sess.email, "APPROVE_USER", `Menyetujui ${r[5]} (${r[4]})`);
  return true;
}

// --- User tidak terdaftar ditolak ---
function rejectUser(rowIndex, token) {
  const sess = getSession_(token);
  if (!sess || sess.role !== "ADMIN") throw new Error("Unauthorized");
  const shA = getSheetByName(SH_APPROVAL);
  const email = shA.getRange(rowIndex, 6).getValue();
  shA.getRange(rowIndex, 2).setValue("REJECTED");
  try {
    MailApp.sendEmail({
      to: email,
      subject: "[BOSP] Registrasi ditolak",
      htmlBody: `<p>Maaf, registrasi Anda ditolak. Silakan hubungi admin bila perlu.</p>`
    });
  } catch(e) {}
  logAktivitas_(sess.email, "REJECT_USER", `Menolak ${email}`);
  return true;
}

// --- User management (block / unblock) ---
function listUsers(token) {
  const sess = getSession_(token);
  if (!sess || sess.role !== "ADMIN") throw new Error("Unauthorized");
  const sh = getSheetByName(SH_USER);
  const last = sh.getLastRow();
  const rows = [];
  if (last>1) {
    const vals = sh.getRange(2,1,last-1,HDR_USER.length).getValues();
    vals.forEach((r, i) => {
      rows.push({
        i: i+2,
        timestamp: r[0], role: r[1], status: r[2], jenjang: r[3],
        negeri_swasta: r[4], sekolah: r[5], email: r[6], password: r[7]
      });
    });
  }
  return rows;
}
function setUserStatus(rowIndex, newStatus, token) {
  const sess = getSession_(token);
  if (!sess || sess.role !== "ADMIN") throw new Error("Unauthorized");
  const sh = getSheetByName(SH_USER);
  sh.getRange(rowIndex, 3).setValue(newStatus); // status
  const email = sh.getRange(rowIndex, 7).getValue();
  logAktivitas_(sess.email, "SET_USER_STATUS", `${email} => ${newStatus}`);
  return true;
}
function updateUserPassword(rowIndex, newPassword, token) {
  const sess = getSession_(token);
  if (!sess || sess.role !== "ADMIN") throw new Error("Unauthorized");
  const sh = getSheetByName(SH_USER);
  sh.getRange(rowIndex, 8).setValue(newPassword);
  const email = sh.getRange(rowIndex, 7).getValue();
  logAktivitas_(sess.email, "UPDATE_PWD", `${email} password diubah`);
  return true;
}

// --- Laporan CRUD ---
function listLaporan(filters, token) {
  const sess = getSession_(token);
  if (!sess) throw new Error("Session habis / tidak valid.");
  const sh = getSheetByName(SH_LAPORAN);
  const last = sh.getLastRow();
  const rows = [];
  if (last>1) {
    const vals = sh.getRange(2,1,last-1,HDR_LAPORAN.length).getValues();
    vals.forEach((r) => {
      // normalisasi semua field penting ke string agar filter tidak gagal 
      // ketika Sheets mengembalikan number atau date
      const rec = {
        id: String(r[0] || "").trim(),
        timestamp: String(r[1] || "").trim(),
        owner_email: String(r[2] || "").trim(),
        sekolah: String(r[3] || "").trim(),
        tahun: String(r[4] || "").trim(),
        sumber_dana: String(r[5] || "").trim(),
        kategori: String(r[6] || "").trim(),
        bulan: String(r[7] || "").trim(),
        keterangan: String(r[8] || "").trim(),
        fileId: String(r[9] || "").trim(),
        fileUrl: String(r[10] || "").trim(),
        fileName: String(r[11] || "").trim()
      };
      // bagi USER biasa, hanya sekolahnya sendiri
      if (sess.role !== "ADMIN") {
        if ((rec.sekolah+"").toLowerCase() !== (sess.sekolah+"").toLowerCase()) return;
      } else {
        // ADMIN: jika filter sekolah di set, terapkan
        // if (filters && filters.sekolah && String(filters.sekolah).trim() && (rec.sekolah+"").toLowerCase() !== (String(filters.sekolah).trim()+"").toLowerCase()) return;
      }
      // filter lain (gunakan perbandingan string yang dinormalisasi)
      if (filters) {
        if (filters.sekolah && String(filters.sekolah).trim() !== "Semua Sekolah" && rec.sekolah !== String(filters.sekolah).trim()) return;
        if (filters.tahun && String(filters.tahun).trim() !== "Semua Tahun Anggaran" && rec.tahun !== String(filters.tahun).trim()) return;
        if (filters.sumber_dana && String(filters.sumber_dana).trim() !== "Semua Sumber Dana" && rec.sumber_dana !== String(filters.sumber_dana).trim()) return;
        if (filters.kategori && String(filters.kategori).trim() !== "Semua Jenis Laporan" && rec.kategori !== String(filters.kategori).trim()) return;
        if (filters.bulan && String(filters.bulan).trim() !== "Semua Bulan" && rec.bulan !== String(filters.bulan).trim()) return;
        if (filters.q && String(filters.q).trim() && !(`${rec.keterangan}`.toLowerCase().includes(String(filters.q).toLowerCase()))) return;
      }
      rows.push(rec);
    });
  }
  return rows;
}
// Tambah Laporan
function addOrUpdateLaporan(form, token) {
  const sess = getSession_(token);
  if (!sess) throw new Error("Session habis / tidak valid.");

  // form: {id?, sekolah?, tahun, sumber_dana, kategori, bulan, keterangan, fileDataUrl?}
  const { id, tahun, sumber_dana, kategori, bulan, keterangan, fileDataUrl } = form || {};
  const sekolah = (sess.role === "ADMIN" && form.sekolah) ? form.sekolah : sess.sekolah;
  if (!sekolah || !tahun || !sumber_dana || !kategori || !bulan) {
    throw new Error("Field wajib (sekolah/tahun/sumber dana/kategori/bulan) belum lengkap.");
  }

  const sh = getSheetByName(SH_LAPORAN);

  let fileId = "", fileUrl = "", fileName = "";
  if (fileDataUrl) {
    const parts = fileDataUrl.split(',');
    if (parts.length < 2) throw new Error("File upload tidak valid.");
    const contentType = (fileDataUrl.match(/^data:(.*?);base64,/)||[])[1] || "application/pdf";
    const bytes = Utilities.base64Decode(parts[1]);
    let blob = Utilities.newBlob(bytes, contentType, "upload.pdf");

    // susun folder: Sekolah / Sumber Dana / Kategori
    const fSchool = getOrCreateSchoolFolder_(sekolah);
    const fDana = getOrCreateSubFolder_(fSchool, sumber_dana);
    const fTh = getOrCreateSubFolder_(fDana, tahun) ;
    const fKat  = getOrCreateSubFolder_(fTh, kategori);

    // nama file: Sekolah - SumberDana - JenisLaporan - Bulan Tahun.pdf
    const safeSek = sekolah.replace(/[\\/:*?"<>|]/g,'-');
    const safeDana = sumber_dana.replace(/[\\/:*?"<>|]/g,'-');
    const safeKat = kategori.replace(/[\\/:*?"<>|]/g,'-');
    const safeBulan = bulan.replace(/[\\/:*?"<>|]/g,'-');
    const safeTahun = tahun.replace(/[\\/:*?"<>|]/g,'-');
    fileName = `${safeSek} - ${safeDana} - ${safeKat} - ${safeBulan} ${safeTahun}.pdf`;

    blob.setName(fileName);
    const file = fKat.createFile(blob);
    fileId = file.getId();
    fileUrl = file.getUrl();
  }

  if (id) {
    // update
    const idx = findRowById_(id);
    if (!idx) throw new Error("Data tidak ditemukan.");
    if (fileId) { // jika ada file baru, update metadata
      sh.getRange(idx, 10, 1, 3).setValues([[fileId, fileUrl, fileName]]);
    }
    sh.getRange(idx, 5, 1, 5).setValues([[tahun, sumber_dana, kategori, bulan, keterangan || ""]]);
    logAktivitas_(sess.email, "UPDATE_LAPORAN", `id=${id}`);
    return { updated: true, id };
  } else {
    // insert
    const newId = Utilities.getUuid();
    const row = [newId, nowStr(), sess.email, sekolah, tahun, sumber_dana, kategori, bulan, keterangan || "", fileId, fileUrl, fileName];
    sh.appendRow(row);
    logAktivitas_(sess.email, "ADD_LAPORAN", `id=${newId} ${sekolah}/${sumber_dana}/${kategori}/${bulan}/${tahun}`);
    return { inserted: true, id: newId };
  }
}
// --- Hapus Data Laporan ---
function deleteLaporan(id, token) {
  const sess = getSession_(token);
  if (!sess) throw new Error("Session habis / tidak valid.");
  const idx = findRowById_(id);
  if (!idx) throw new Error("Data tidak ditemukan.");
  const sh = getSheetByName(SH_LAPORAN);
  const row = sh.getRange(idx,1,1,HDR_LAPORAN.length).getValues()[0];
  // hapus file di Drive juga (opsional)
  const fid = row[9];
  if (fid) {
    try { DriveApp.getFileById(fid).setTrashed(true); } catch(e) {}
  }
  sh.deleteRow(idx);
  logAktivitas_(sess.email, "DEL_LAPORAN", `id=${id}`);
  return true;
}
// --- Pencarian Data ---
function findRowById_(id) {
  const sh = getSheetByName(SH_LAPORAN);
  const last = sh.getLastRow();
  if (last<2) return 0;
  const vals = sh.getRange(2,1,last-1,1).getValues();
  for (let i=0;i<vals.length;i++) if (vals[i][0] === id) return i+2;
  return 0;
}

// --- Admin: daftar sekolah untuk filter global (ambil dari user sheet unik) ---
function listSekolahAll(token) {
  const sess = getSession_(token);
  if (!sess || sess.role !== "ADMIN") throw new Error("Unauthorized");
  const sh = getSheetByName(SH_USER);
  const last = sh.getLastRow();
  const set = {};
  if (last>1) {
    const vals = sh.getRange(2,1,last-1,HDR_USER.length).getValues();
    vals.forEach(r => {
      const s = (r[5]||"").toString().trim();
      if (s) set[s] = true;
    });
  }
  return Object.keys(set).sort();
}

// --- Session helper untuk client ---
function pingSession(token) {
  return refreshSession_(token);
}
function getSessionInfo(token) {
  const sess = getSession_(token);
  if (!sess) return null;
  return sess;
}

// --- Helpers ---
function escapeHtml_(str) {
  if (str === null || str === undefined) return "";
  return String(str).replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
}
