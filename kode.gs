const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const FOLDER_NAME_FOTO = "SIPBAN_Foto_Kerusakan";
const SHEET_NAMES = {
  USERS: 'users',
  DATA_RUMAH: 'data_rumah',
  DATA_BANTUAN: 'data_bantuan',
  LOG_PERBAIKAN: 'log_perbaikan',
  SETTINGS: 'settings',
  DATA_GANDA: 'data_ganda',
  DATA_RUMAH_DINAS: 'data_rumah_dinas'
};

function getTargetSheetName(kategori) {
  return kategori === 'dinas' ? SHEET_NAMES.DATA_RUMAH_DINAS : SHEET_NAMES.DATA_RUMAH;
}

// ==========================================
// 1. SETUP DATABASE
// ==========================================
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  createSheetIfNotExists(ss, SHEET_NAMES.USERS, ['id', 'username', 'password', 'role', 'created_at']);
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (usersSheet.getLastRow() <= 1) {
    usersSheet.appendRow([Date.now().toString(), 'admin', 'admin123', 'admin', new Date().toISOString()]);
    usersSheet.appendRow([Date.now().toString()+1, 'petugas', 'petugas123', 'petugas', new Date().toISOString()]);
    usersSheet.appendRow([Date.now().toString()+2, 'surveyor', 'surveyor123', 'surveyor', new Date().toISOString()]);
    usersSheet.appendRow([Date.now().toString()+3, 'kecamatan', 'kecamatan123', 'petugas_kecamatan', new Date().toISOString()]);
  }

  createSheetIfNotExists(ss, SHEET_NAMES.SETTINGS, ['key', 'value']);
  const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (settingsSheet.getLastRow() <= 1) {
    settingsSheet.appendRow(['active_phase', 'BNBA 1']);
    settingsSheet.appendRow(['phase_status', 'open']);
  }
  
  const headersRumah = [
    'id', 'petugas_username', 'surveyor_username', 'nama_pemilik', 'nik', 'no_kk', 
    'dusun', 'desa', 'kecamatan', 'koordinat', 'tingkat_kerusakan', 'kebutuhan_mendesak', 
    'foto_1', 'foto_2', 'foto_3', 'foto_4', 'foto_5', 'foto_6', 'status', 'created_at', 'fase_bnba', 'catatan_tolak', 'nama_penghuni'
  ];
  
  createSheetIfNotExists(ss, SHEET_NAMES.DATA_RUMAH, headersRumah);
  createSheetIfNotExists(ss, SHEET_NAMES.DATA_RUMAH_DINAS, headersRumah);
  createSheetIfNotExists(ss, SHEET_NAMES.DATA_BANTUAN, ['id_laporan', 'jenis_bantuan', 'no_sk', 'keterangan', 'updated_at']);
  createSheetIfNotExists(ss, SHEET_NAMES.LOG_PERBAIKAN, ['id_laporan', 'field', 'old_value', 'new_value', 'admin_user', 'timestamp']);
  createSheetIfNotExists(ss, SHEET_NAMES.DATA_GANDA, ['ID Laporan', 'Nama Pemilik', 'NIK', 'No KK', 'Desa', 'Keterangan Ganda', 'Status', 'Fase BNBA', 'Terakhir Dicek']);
  
  let folders = DriveApp.getFoldersByName(FOLDER_NAME_FOTO);
  if (!folders.hasNext()) { DriveApp.createFolder(FOLDER_NAME_FOTO); }

  Logger.log('✅ Database Ter-setup!');
}

function createSheetIfNotExists(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f1f5f9');
  }
  return sheet;
}

// ==========================================
// 2. ROUTER WEB APP
// ==========================================
function doGet(e) { 
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: 'API Aktif' })).setMimeType(ContentService.MimeType.JSON);
}

/*function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const payload = data.payload;
    let result = {};

    switch(action) {
      case 'login': result = handleLogin(payload); break;
      case 'getConfig': result = handleGetConfig(); break;
      case 'updateConfig': result = handleUpdateConfig(payload); break;
      case 'getPublicStats': result = handleGetPublicStats(); break;
      case 'checkDataWarga': result = handleCheckDataWarga(payload); break;
      case 'getData': result = handleGetData(payload); break;
      case 'createLaporan': result = handleCreateLaporan(payload); break;
      case 'submitSurvey': result = handleSubmitSurvey(payload); break;
      case 'editLaporan': result = handleEditLaporan(payload); break;
      case 'deleteLaporan': result = handleDeleteLaporan(payload); break;
      case 'submitBantuan': result = handleSubmitBantuan(payload); break;
      case 'getLogPerbaikan': result = handleGetLogPerbaikan(payload); break;
      case 'getUsers': result = handleGetUsers(); break;
      case 'createUser': result = handleCreateUser(payload); break;
      case 'deleteUser': result = handleDeleteUser(payload); break;
      case 'importData': result = handleImportData(payload); break;
      case 'getDuplicatesData': result = handleGetDuplicatesData(); break;
      case 'syncDuplicatesToSheet': result = handleSyncDuplicatesToSheet(); break;
      default: throw new Error("Aksi tidak valid");
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: error.message})).setMimeType(ContentService.MimeType.JSON);
  }
}
*/
//perbaikan
function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) {
      throw new Error("Tidak ada data yang dikirim");
    }
    
    let p = JSON.parse(e.postData.contents);
    let action = p.action;
    let payload = p.payload || {};
    let res = {};

    // ROUTER UTAMA
    if (action === 'login') res = handleLogin(payload);
    else if (action === 'getData') res = handleGetData(payload);
    else if (action === 'getConfig') res = handleGetConfig();
    else if (action === 'updateConfig') res = handleUpdateConfig(payload);
    else if (action === 'getPublicStats') res = handleGetPublicStats();
    else if (action === 'checkDataWarga') res = handleCheckDataWarga(payload);
    else if (action === 'createLaporan') res = handleCreateLaporan(payload);
    else if (action === 'editLaporan') res = handleEditLaporan(payload);
    else if (action === 'deleteLaporan') res = handleDeleteLaporan(payload);
    else if (action === 'submitSurvey') res = handleSubmitSurvey(payload);
    else if (action === 'submitBantuan') res = handleSubmitBantuan(payload);
    else if (action === 'getUsers') res = handleGetUsers();
    else if (action === 'createUser') res = handleCreateUser(payload);
    else if (action === 'deleteUser') res = handleDeleteUser(payload);
    else if (action === 'importData') res = handleImportData(payload);
    else if (action === 'getDuplicatesData') res = handleGetDuplicatesData();
    else if (action === 'syncDuplicatesToSheet') res = handleSyncDuplicatesToSheet();
    else if (action === 'getLogPerbaikan') res = handleGetLogPerbaikan(payload);
    else throw new Error("Aksi tidak ditemukan");

    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 3. AUTENTIKASI & KONFIGURASI
// ==========================================
function handleLogin(p) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.USERS).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === p.username && data[i][2] === p.password) return { status: 'success', data: { username: data[i][1], role: data[i][3] } };
  }
  throw new Error("Username atau password salah");
}

function fetchConfigData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
  const data = sheet.getDataRange().getValues();
  let conf = { phase: 'BNBA 1', status: 'closed' };
  for(let i=1; i<data.length; i++){
    if(data[i][0] === 'active_phase') conf.phase = data[i][1];
    if(data[i][0] === 'phase_status') conf.status = data[i][1];
  }
  return conf;
}

function handleGetConfig() { 
  return { status: 'success', data: fetchConfigData() }; 
}

function handleUpdateConfig(p) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){
    if(data[i][0] === 'active_phase') sheet.getRange(i+1, 2).setValue(p.phase);
    if(data[i][0] === 'phase_status') sheet.getRange(i+1, 2).setValue(p.status);
  }
  return { status: 'success' };
}

// ==========================================
// 4. API PUBLIK (STATISTIK & CARI WARGA)
// ==========================================
function handleGetPublicStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Mengambil data laporan warga (dari data_rumah)
  const sheet = ss.getSheetByName('data_rumah'); 
  if (!sheet) return { status: 'error', message: 'Sheet data_rumah tidak ditemukan' };
  
  const data = sheet.getDataRange().getValues();
  
  let total = 0, berat = 0, sedang = 0, ringan = 0, verified = 0;
  let villages = {};

  // Mulai dari baris ke-1 untuk melewati Header
  for (let i = 1; i < data.length; i++) {
    // ANTI-BUG: Lewati baris jika ID Laporannya kosong (mencegah error baris kosong Google Sheets)
    if (!data[i][0] || data[i][0] === '') continue; 
    
    let desa = data[i][7] || 'Lainnya'; // Kolom Desa
    let rusak = data[i][10] || 'Belum Dinilai'; // Kolom Kerusakan
    let status = data[i][18] || ''; // Kolom Status Verifikasi

    total++;
    if (rusak === 'Rusak Berat') berat++;
    if (rusak === 'Rusak Sedang') sedang++;
    if (rusak === 'Rusak Ringan') ringan++;
    if (status === 'Terverifikasi') verified++;

    // Rekap per Desa
    if (!villages[desa]) {
      villages[desa] = { total: 0, berat: 0, sedang: 0, ringan: 0 };
    }
    villages[desa].total++;
    if (rusak === 'Rusak Berat') villages[desa].berat++;
    if (rusak === 'Rusak Sedang') villages[desa].sedang++;
    if (rusak === 'Rusak Ringan') villages[desa].ringan++;
  }

  return {
    status: 'success',
    data: {
      total: total,
      berat: berat,
      sedang: sedang,
      ringan: ringan,
      verified: verified,
      villages: villages
    }
  };
}
function handleCheckDataWarga(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.DATA_RUMAH);
  
  const keyword = p.keyword ? p.keyword.toString().trim() : '';
  if(keyword.length < 5) throw new Error("Nomor NIK/KK tidak valid (terlalu pendek).");

  // OPTIMASI: TextFinder untuk performa cepat tanpa load seluruh data ke memory
  const textFinder = sheet.createTextFinder(keyword).matchEntireCell(true);
  const searchResults = textFinder.findAll();

  if (searchResults.length === 0) {
     throw new Error("Data dengan NIK/KK tersebut belum terdaftar di sistem kami.");
  }

  const bantuanSheet = ss.getSheetByName(SHEET_NAMES.DATA_BANTUAN);
  let bantuanMap = {};
  
  if (bantuanSheet) {
    const bData = bantuanSheet.getDataRange().getValues();
    for (let i = 1; i < bData.length; i++) {
      bantuanMap[bData[i][0]] = bData[i][1];
    }
  }

  let hasilPencarian = [];
  let processedRows = {};
  const maxCol = sheet.getLastColumn();

  for (let i = 0; i < searchResults.length; i++) {
    const row = searchResults[i].getRow();
    if (row === 1) continue; // Abaikan header
    
    if (processedRows[row]) continue;
    processedRows[row] = true;
    
    const rowData = sheet.getRange(row, 1, 1, maxCol).getValues()[0];
    const idLaporan = rowData[0];
    const nik = rowData[4] ? rowData[4].toString().trim() : '';
    const kk = rowData[5] ? rowData[5].toString().trim() : '';
    
    if (nik !== keyword && kk !== keyword) continue;

    let namaAsli = rowData[3] ? rowData[3].toString() : '';
    let namaMasked = namaAsli.substring(0,2) + "*** " + (namaAsli.split(' ').length > 1 ? namaAsli.split(' ').pop().substring(0,1) + "." : "");
    
    let item = {
      id: idLaporan,
      nama: namaMasked,
      desa: `Desa ${rowData[7]}, Kec. ${rowData[8]}`,
      status: rowData[18],
      catatan_tolak: rowData[21] || '',
      nik: nik.length > 4 ? "******" + nik.substring(nik.length - 4) : nik,
      kk: kk.length > 4 ? "******" + kk.substring(kk.length - 4) : kk
    };
    
    if (bantuanMap[idLaporan] && bantuanMap[idLaporan] !== 'Belum Ada') {
      item.bantuan = bantuanMap[idLaporan];
    }
    
    hasilPencarian.push(item);
  }

  if (hasilPencarian.length === 0) {
     throw new Error("Data dengan NIK/KK tersebut belum terdaftar di sistem kami.");
  }

  return { status: 'success', data: hasilPencarian };
}

// ==========================================
// 5. DASHBOARD & MANAJEMEN DATA
// ==========================================
function handleGetData(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // PERBAIKAN: Gunakan fungsi pembaca kategori dinamis (Warga vs Dinas)
  const sheetName = getTargetSheetName(p.kategori);
  const data = ss.getSheetByName(sheetName).getDataRange().getValues();
  
  const bantuanData = ss.getSheetByName(SHEET_NAMES.DATA_BANTUAN).getDataRange().getValues();
  
  let mapBantuan = {};
  for(let i=1; i<bantuanData.length; i++) { 
    mapBantuan[bantuanData[i][0]] = { 
      jenis: bantuanData[i][1], 
      no_sk: bantuanData[i][2], 
      ket: bantuanData[i][3] 
    }; 
  }

  let result = [];

  for (let i = 1; i < data.length; i++) {
    let b = mapBantuan[data[i][0]]; 
    
    let item = {
      id: data[i][0], petugas_username: data[i][1], surveyor_username: data[i][2],
      nama_pemilik: data[i][3], nik: data[i][4], no_kk: data[i][5],
      dusun: data[i][6], desa: data[i][7], kecamatan: data[i][8],
      koordinat: data[i][9], tingkat_kerusakan: data[i][10], kebutuhan_mendesak: data[i][11],
      foto_1: data[i][12], foto_2: data[i][13], foto_3: data[i][14], foto_4: data[i][15], foto_5: data[i][16], foto_6: data[i][17],
      status: data[i][18], created_at: data[i][19], fase_bnba: data[i][20] || '-', catatan_tolak: data[i][21] || '',
      nama_penghuni: data[i][22] || '',
      jenis_bantuan: b ? b.jenis : null,
      no_sk_bantuan: b ? b.no_sk : null,
      keterangan_bantuan: b ? b.ket : null
    };

    // Filter berdasarkan Hak Akses (Role)
    if(p.role === 'petugas' && item.petugas_username === p.username) result.push(item);
    else if(p.role === 'surveyor' && item.status === 'Menunggu Survey') result.push(item); 
    else if(p.role === 'surveyor' && item.surveyor_username === p.username) result.push(item); 
    else if(['admin', 'petugas_kecamatan'].includes(p.role)) result.push(item); 
  }
  
  return { status: 'success', data: result.reverse() }; 
}

function uploadFotosToDrive(fotosArray, rowId) {
  let uploadedUrls = {1:'', 2:'', 3:'', 4:'', 5:'', 6:''};
  if(!fotosArray || fotosArray.length === 0) return uploadedUrls;
  let folders = DriveApp.getFoldersByName(FOLDER_NAME_FOTO);
  let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME_FOTO);
  
  fotosArray.forEach(f => {
    let blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.mime, `F${f.index}_${rowId}_${f.name}`);
    let file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    uploadedUrls[f.index] = file.getUrl();
  });
  return uploadedUrls;
}

function handleCreateLaporan(p) {
  const conf = fetchConfigData();
  if(conf.status !== 'open' && p.role === 'petugas') throw new Error("Fase Pendataan sedang ditutup.");

  const sheetName = getTargetSheetName(p.kategori);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  const prefix = p.kategori === 'dinas' ? 'D' : 'W';
  const newId = prefix + Date.now().toString().substring(3);
  const d = p.identitas || {};
  const s = p.survey || {};
  
  // Panggil fungsi Folder NIK (Gunakan NIK, jika kosong gunakan ID Laporan)
  const folderName = d.nik || newId; 
  let fotoUrls = simpanFotoKeFolderNIK(p.fotos, folderName);

  const status = p.status || 'Menunggu Survey';
  const catatan = p.catatan_tolak || '';
  const namaPenghuni = d.nama_penghuni || '';

  // FIX: Mengurutkan susunan 23 Kolom dengan benar agar tidak meleset
  sheet.appendRow([
    newId, p.username, '', d.nama || '', d.nik || '', d.kk || '', 
    d.dusun || '', d.desa || '', d.kec || '', 
    s.koordinat || '', s.rusak || '', s.kebutuhan || '', 
    fotoUrls[1], fotoUrls[2], fotoUrls[3], fotoUrls[4], fotoUrls[5], fotoUrls[6], 
    status, new Date().toISOString(), conf.phase, catatan, namaPenghuni
  ]);
  
  return { status: 'success' };
}

function handleSubmitSurvey(p) {
  const sheetName = getTargetSheetName(p.kategori);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  let nikDariSheet = p.id;
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0] == p.id) {
      rowIndex = i + 1;
      nikDariSheet = data[i][4]; // Kolom NIK
      break;
    }
  }
  if(rowIndex === -1) throw new Error("Data tidak ditemukan");

  const folderName = nikDariSheet || p.id;
  let fotoUrls = simpanFotoKeFolderNIK(p.fotos, folderName);

  const s = p.survey;
  // FIX: Penyesuaian nomor urut index kolom
  sheet.getRange(rowIndex, 3).setValue(p.username); // Kolom 3 = surveyor_username
  sheet.getRange(rowIndex, 10).setValue(s.koordinat); // Kolom 10 = koordinat
  sheet.getRange(rowIndex, 11).setValue(s.rusak); // Kolom 11 = tingkat_kerusakan
  sheet.getRange(rowIndex, 12).setValue(s.kebutuhan); // Kolom 12 = kebutuhan_mendesak
  sheet.getRange(rowIndex, 19).setValue('Pending'); // Kolom 19 = status
  
  for(let i=1; i<=6; i++) {
    if(fotoUrls[i]) sheet.getRange(rowIndex, 12 + i).setValue(fotoUrls[i]); // Mulai Kolom 13 s.d 18
  }

  return { status: 'success' };
}

function handleEditLaporan(p) {
  const sheetName = getTargetSheetName(p.kategori);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  let nikDariSheet = p.id;
  let oldData = [];
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0] == p.id) {
      rowIndex = i + 1;
      oldData = data[i]; // Merekam data lama sebelum diubah
      nikDariSheet = data[i][4];
      break;
    }
  }
  if(rowIndex === -1) throw new Error("Data tidak ditemukan");

  const d = p.identitas || {};
  const s = p.survey || {};
  
  // Panggil fungsi Folder NIK
  const folderName = d.nik || nikDariSheet || p.id;
  let fotoUrls = simpanFotoKeFolderNIK(p.fotos, folderName);

  // --- LOGIKA PENCATATAN LOG PERBAIKAN ---
  const logSheet = ss.getSheetByName(SHEET_NAMES.LOG_PERBAIKAN);
  const timestamp = new Date().toISOString();
  let logsToAppend = [];

  // Fungsi helper untuk membandingkan data lama vs data baru
  function checkLog(fieldIndex, fieldName, newValue) {
    if (newValue !== undefined && newValue !== null) {
      const oldValue = oldData[fieldIndex] || "";
      if (oldValue.toString().trim() !== newValue.toString().trim()) {
        logsToAppend.push([p.id, fieldName, oldValue, newValue, p.username, timestamp]);
      }
    }
  }

  // Update Data Identitas & Cek Log (HANYA NAMA, NIK, DAN KK YANG DICATAT)
  if(Object.keys(d).length > 0) {
    checkLog(3, 'Nama Pemilik', d.nama);
    checkLog(4, 'NIK', d.nik);
    checkLog(5, 'No KK', d.kk);

    sheet.getRange(rowIndex, 4, 1, 6).setValues([[d.nama, d.nik, d.kk, d.dusun, d.desa, d.kec]]);
    if (d.nama_penghuni !== undefined) sheet.getRange(rowIndex, 23).setValue(d.nama_penghuni);
  }
  
  // Update Data Survey & Foto (Tanpa Log)
  if(Object.keys(s).length > 0) {
    sheet.getRange(rowIndex, 10).setValue(s.koordinat); 
    sheet.getRange(rowIndex, 11).setValue(s.rusak); 
    sheet.getRange(rowIndex, 12).setValue(s.kebutuhan); 
    for(let i=1; i<=6; i++) {
      if(fotoUrls[i]) sheet.getRange(rowIndex, 12 + i).setValue(fotoUrls[i]); 
    }
  }

  // Update Status (Tanpa Log)
  if(p.status) {
    sheet.getRange(rowIndex, 19).setValue(p.status);
    sheet.getRange(rowIndex, 22).setValue(p.catatan_tolak || '');
  }

  // Eksekusi Simpan Log ke Database
  if (logsToAppend.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logsToAppend.length, logsToAppend[0].length).setValues(logsToAppend);
  }

  return { status: 'success' };
}

function handleDeleteLaporan(p) {
  const sheetName = getTargetSheetName(p.kategori);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == p.id) {
      sheet.deleteRow(i + 1);
      return { status: 'success' };
    }
  }
  throw new Error("Gagal menghapus");
}

function handleImportData(p) {
  const sheetName = getTargetSheetName(p.kategori);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const records = p.data;
  if(!records || records.length === 0) throw new Error("Data kosong");
  
  const rowsToAppend = [];
  const now = new Date().toISOString();
  const prefix = p.kategori === 'dinas' ? 'D' : 'W';
  
  records.forEach((d, index) => {
    const newId = prefix + (Date.now() + index).toString().substring(3); 
    rowsToAppend.push([
      newId, p.username, '', d.nama, d.nik, d.kk, d.dusun, d.desa, d.kecamatan, 
      '', '', '', '', '', '', '', '', '', 'Menunggu Survey', now, d.fase, '', d.nama_penghuni || ''
    ]);
  });
  
  sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  return { status: 'success', count: rowsToAppend.length };
}

// ==========================================
// 6. BANTUAN & MANAJEMEN PENGGUNA
// ==========================================
function handleSubmitBantuan(p) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.DATA_BANTUAN);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == p.id_laporan) {
      sheet.getRange(i+1, 2, 1, 4).setValues([[p.jenis, p.no_sk, p.keterangan, new Date().toISOString()]]);
      return { status: 'success' };
    }
  }
  sheet.appendRow([p.id_laporan, p.jenis, p.no_sk, p.keterangan, new Date().toISOString()]);
  return { status: 'success' };
}

function handleGetLogPerbaikan(p) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.LOG_PERBAIKAN).getDataRange().getValues();
  let result = [];
  for(let i=1; i<data.length; i++){ if(data[i][0] == p.id_laporan) result.push({ field: data[i][1], old_value: data[i][2], new_value: data[i][3], admin_user: data[i][4], timestamp: data[i][5] }); }
  return { status: 'success', data: result.reverse() };
}

function handleGetUsers() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.USERS).getDataRange().getValues();
  let result = [];
  for(let i=1; i<data.length; i++) { result.push({ username: data[i][1], role: data[i][3], created_at: data[i][4] }); }
  return { status: 'success', data: result };
}

function handleCreateUser(p) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) { if(data[i][1] === p.username) throw new Error("Username sudah dipakai"); }
  sheet.appendRow([Date.now().toString(), p.username, p.password, p.role, new Date().toISOString()]);
  return { status: 'success' };
}

function handleDeleteUser(p) {
  if(p.username === 'admin') throw new Error("Admin utama tidak bisa dihapus!");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][1] === p.username) { sheet.deleteRow(i+1); return { status: 'success' }; }
  }
  throw new Error("User tidak ditemukan");
}

// ==========================================
// 7. MANAJEMEN DATA GANDA
// ==========================================
function handleGetDuplicatesData(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName(SHEET_NAMES.DATA_RUMAH).getDataRange().getValues();
  const bantuanData = ss.getSheetByName(SHEET_NAMES.DATA_BANTUAN).getDataRange().getValues();
  
  let mapBantuan = {};
  for(let i=1; i<bantuanData.length; i++) { 
    mapBantuan[bantuanData[i][0]] = { jenis: bantuanData[i][1], no_sk: bantuanData[i][2], ket: bantuanData[i][3] }; 
  }

  let nikCounts = {};
  let kkCounts = {};
  
  for (let i = 1; i < data.length; i++) {
    let nik = data[i][4] ? data[i][4].toString().trim() : '';
    let kk = data[i][5] ? data[i][5].toString().trim() : '';
    if (nik) nikCounts[nik] = (nikCounts[nik] || 0) + 1;
    if (kk) kkCounts[kk] = (kkCounts[kk] || 0) + 1;
  }
  
  let result = [];
  for (let i = 1; i < data.length; i++) {
    let nik = data[i][4] ? data[i][4].toString().trim() : '';
    let kk = data[i][5] ? data[i][5].toString().trim() : '';
    
    if ((nik && nikCounts[nik] > 1) || (kk && kkCounts[kk] > 1)) {
      // FIX: Menggunakan variabel b sebagai pengganti Optional Chaining (?.)
      let b = mapBantuan[data[i][0]];

      let item = {
        id: data[i][0], petugas_username: data[i][1], surveyor_username: data[i][2],
        nama_pemilik: data[i][3], nik: nik, no_kk: kk,
        dusun: data[i][6], desa: data[i][7], kecamatan: data[i][8],
        koordinat: data[i][9], tingkat_kerusakan: data[i][10], kebutuhan_mendesak: data[i][11],
        foto_1: data[i][12], foto_2: data[i][13], foto_3: data[i][14], foto_4: data[i][15], foto_5: data[i][16], foto_6: data[i][17],
        status: data[i][18], created_at: data[i][19], fase_bnba: data[i][20] || '-',
        catatan_tolak: data[i][21] || '',
        jenis_bantuan: b ? b.jenis : null,
        no_sk_bantuan: b ? b.no_sk : null,
        keterangan_bantuan: b ? b.ket : null,
        keterangan_ganda: (nikCounts[nik] > 1 ? '⚠️ NIK Ganda ' : '') + (kkCounts[kk] > 1 ? '⚠️ KK Ganda' : '')
      };
      result.push(item);
    }
  }
  
  result.sort((a, b) => (a.nik || '').localeCompare(b.nik || ''));
  return { status: 'success', data: result };
}

function handleSyncDuplicatesToSheet(p) {
   const duplicates = handleGetDuplicatesData(p).data;
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   let sheet = ss.getSheetByName(SHEET_NAMES.DATA_GANDA);
   
   sheet.getRange(2, 1, Math.max(sheet.getMaxRows(), 2), sheet.getMaxColumns()).clearContent();
   
   if(duplicates.length > 0) {
       let rows = duplicates.map(d => [
           d.id, d.nama_pemilik, `'${d.nik}`, `'${d.no_kk}`, d.desa, d.keterangan_ganda, d.status, d.fase_bnba, new Date().toISOString()
       ]);
       sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
   }
   return { status: 'success', count: duplicates.length };
}

function simpanFotoKeFolderNIK(fotos, folderName) {
  // ⚠️ GANTI TEKS DI BAWAH INI DENGAN ID FOLDER UTAMA GOOGLE DRIVE ANDA
  // const FOLDER_UTAMA_ID = "ISI_DENGAN_ID_FOLDER_UTAMA_ANDA"; 
  const FOLDER_UTAMA_ID = "1SfCZ-Z8cNxMwsmGUFJLhy-x_J2uxDjr9"; 
  
  let urls = {1:'', 2:'', 3:'', 4:'', 5:'', 6:''};
  if (!fotos || fotos.length === 0) return urls;

  const mainFolder = DriveApp.getFolderById(FOLDER_UTAMA_ID);
  let targetFolder;
  
  // Cek apakah folder dengan nama NIK ini sudah ada di dalam Folder Utama
  const folderIterator = mainFolder.getFoldersByName(folderName);
  if (folderIterator.hasNext()) {
    targetFolder = folderIterator.next(); // Jika ada, gunakan folder tersebut
  } else {
    targetFolder = mainFolder.createFolder(folderName); // Jika belum ada, BUAT folder baru
  }

  // Simpan semua foto ke dalam folder NIK tersebut
  for (let i = 0; i < fotos.length; i++) {
    let f = fotos[i];
    // Nama file: F1_3201234567890_17838...jpg
    let fileName = "F" + f.index + "_" + folderName + "_" + Date.now() + ".jpg";
    let blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.mime, fileName);
    let file = targetFolder.createFile(blob);
    
    // Buka akses file agar bisa dilihat di web
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    urls[f.index] = file.getUrl();
  }
  
  return urls;
}
