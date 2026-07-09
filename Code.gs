// ================== GLOBAL SPREADSHEET ==================
// Script ini terikat dengan spreadsheet, jadi gunakan getActiveSpreadsheet()
const ss = SpreadsheetApp.getActiveSpreadsheet();

// ================== DO GET ==================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('InventoryPro')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ================== SESSION MANAGEMENT ==================
function login(username, password) {
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][1] === password) {
      const user = { name: data[i][2] };
      const token = Utilities.getUuid();
      PropertiesService.getUserProperties().setProperty('session', JSON.stringify({ user, token }));
      return { success: true, user: user, token: token };
    }
  }
  return { success: false, message: 'Username atau password salah' };
}

function logout() {
  PropertiesService.getUserProperties().deleteProperty('session');
}

function getSession() {
  const sessionStr = PropertiesService.getUserProperties().getProperty('session');
  if (sessionStr) {
    try {
      return JSON.parse(sessionStr).user;
    } catch(e) {}
  }
  return null;
}

// ================== DATA RETRIEVAL ==================
function getAllData() {
  return {
    masterData: getMasterData_(),
    transaksiKeluar: getTransaksiKeluar_(),
    peminjaman: getPeminjaman_(),
    riwayat: getRiwayat_()
  };
}

// Helper functions (private)
function getMasterData_() {
  const sheet = ss.getSheetByName('MasterData');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => ({
    id_barang: row[0],
    nomor_ticket: row[1],
    kategori: row[2],
    nama_barang: row[3],
    serial_number: row[4],
    fa_number: row[5],
    qty: row[6],
    lokasi: row[7],
    keterangan: row[8]
  }));
}

function getTransaksiKeluar_() {
  const sheet = ss.getSheetByName('TransaksiKeluar');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => ({
    id: row[0],
    tanggal: row[1],
    id_barang: row[2],
    kategori: row[3],
    nama_barang: row[4],
    serial_number: row[5],
    fa_number: row[6],
    no_ticket: row[7],
    qty: row[8],
    keterangan: row[9]
  }));
}

function getPeminjaman_() {
  const sheet = ss.getSheetByName('Peminjaman');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => ({
    id_pinjam: row[0],
    nama_peminjam: row[1],
    nomor_ticket: row[2],
    tanggal_pinjam: row[3],
    items: JSON.parse(row[4] || '[]'),
    status: row[5],
    tanggal_kembali: row[6]
  }));
}

function getRiwayat_() {
  const sheet = ss.getSheetByName('Riwayat');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => ({
    id: row[0],
    waktu: row[1],
    tipe: row[2],
    detail: row[3],
    qty: row[4],
    user: row[5]
  }));
}

// ================== ID COUNTER ==================
function getNextId(key) {
  const sheet = ss.getSheetByName('IdCounter');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  return 1;
}

function updateIdCounter(key) {
  const sheet = ss.getSheetByName('IdCounter');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(data[i][1] + 1);
      break;
    }
  }
}

function generateId(prefix, key) {
  const next = getNextId(key);
  const numStr = ('000' + next).slice(-3);
  const newId = prefix + numStr;
  updateIdCounter(key);
  return newId;
}

// ================== MASTER DATA OPERATIONS ==================
function addMasterData(item) {
  const sheet = ss.getSheetByName('MasterData');
  sheet.appendRow([
    item.id_barang,
    item.nomor_ticket,
    item.kategori,
    item.nama_barang,
    item.serial_number,
    item.fa_number,
    item.qty,
    item.lokasi,
    item.keterangan
  ]);
  addRiwayat_('tambah_barang', `Tambah barang ${item.id_barang} - ${item.nama_barang}`, item.qty);
}

function updateMasterData(item) {
  const sheet = ss.getSheetByName('MasterData');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === item.id_barang) {
      sheet.getRange(i + 1, 1, 1, 9).setValues([[
        item.id_barang,
        item.nomor_ticket,
        item.kategori,
        item.nama_barang,
        item.serial_number,
        item.fa_number,
        item.qty,
        item.lokasi,
        item.keterangan
      ]]);
      addRiwayat_('edit_barang', `Edit barang ${item.id_barang} - ${item.nama_barang}`, item.qty);
      break;
    }
  }
}

function deleteMasterData(id_barang) {
  const sheet = ss.getSheetByName('MasterData');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id_barang) {
      const nama = data[i][3];
      const qty = data[i][6];
      sheet.deleteRow(i + 1);
      addRiwayat_('hapus_barang', `Hapus barang ${id_barang} - ${nama}`, qty);
      break;
    }
  }
}

// ================== STOK OPERATIONS ==================
function kurangiStok_(id_barang, qty) {
  const sheet = ss.getSheetByName('MasterData');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id_barang) {
      const current = data[i][6];
      sheet.getRange(i + 1, 7).setValue(current - qty);
      break;
    }
  }
}

function tambahStok_(id_barang, qty) {
  const sheet = ss.getSheetByName('MasterData');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id_barang) {
      const current = data[i][6];
      sheet.getRange(i + 1, 7).setValue(current + qty);
      break;
    }
  }
}

// ================== BARANG KELUAR ==================
function addBarangKeluar(transaksi) {
  const sheet = ss.getSheetByName('TransaksiKeluar');
  sheet.appendRow([
    transaksi.id,
    transaksi.tanggal,
    transaksi.id_barang,
    transaksi.kategori,
    transaksi.nama_barang,
    transaksi.serial_number,
    transaksi.fa_number,
    transaksi.no_ticket,
    transaksi.qty,
    transaksi.keterangan
  ]);
  kurangiStok_(transaksi.id_barang, transaksi.qty);
  addRiwayat_('barang_keluar', `Keluar: ${transaksi.id} - ${transaksi.nama_barang}`, transaksi.qty);
}

// ================== PEMINJAMAN ==================
function addPeminjaman(peminjaman) {
  const sheet = ss.getSheetByName('Peminjaman');
  peminjaman.items.forEach(item => kurangiStok_(item.id_barang, item.qty));
  sheet.appendRow([
    peminjaman.id_pinjam,
    peminjaman.nama_peminjam,
    peminjaman.nomor_ticket,
    peminjaman.tanggal_pinjam,
    JSON.stringify(peminjaman.items),
    peminjaman.status,
    ''
  ]);
  const totalQty = peminjaman.items.reduce((s, i) => s + i.qty, 0);
  addRiwayat_('peminjaman', `Pinjam: ${peminjaman.id_pinjam} oleh ${peminjaman.nama_peminjam}`, totalQty);
}

function kembalikanPeminjaman(id_pinjam) {
  const sheet = ss.getSheetByName('Peminjaman');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id_pinjam && data[i][5] === 'dipinjam') {
      const items = JSON.parse(data[i][4] || '[]');
      items.forEach(item => tambahStok_(item.id_barang, item.qty));
      sheet.getRange(i + 1, 6).setValue('dikembalikan');
      sheet.getRange(i + 1, 7).setValue(new Date().toISOString());
      const totalQty = items.reduce((s, item) => s + item.qty, 0);
      const namaPeminjam = data[i][1];
      addRiwayat_('pengembalian', `Pengembalian ${id_pinjam} oleh ${namaPeminjam}`, totalQty);
      break;
    }
  }
}

// ================== RIWAYAT ==================
function addRiwayat_(tipe, detail, qty) {
  const sheet = ss.getSheetByName('Riwayat');
  const id = generateId('RWY-', 'riwayat');
  const currentUser = getSession();
  const username = currentUser ? currentUser.name : 'System';
  sheet.appendRow([id, new Date().toISOString(), tipe, detail, qty, username]);
}