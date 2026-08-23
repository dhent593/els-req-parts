/**
 * BACKEND APP SCRIPT - ELS COMPUTER SERVICE
 * Simpan file ini sebagai Code.gs
 */

const DB_ID = '1m36knzb3UvofROd9oHXBS52q5yGSW0iIGHKexiUjoVY';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ELS Service Admin')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// UTILITY CACHE & DATABASE FUNCTIONS
// ==========================================

function _getServerCache(key) {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);
    return cached ? JSON.parse(cached) : null;
  } catch(e) { return null; }
}

function _setServerCache(key, data) {
  try { 
    CacheService.getScriptCache().put(key, JSON.stringify(data), 900); // 15 menit
  } catch(e) {}
}

function _invalidateCache(keysArray) {
  try {
    const cache = CacheService.getScriptCache();
    keysArray.forEach(k => cache.remove(k));
  } catch(e) {}
}

function getSheetInfo(sheetName) {
  const ss = SpreadsheetApp.openById(DB_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

function getRowsData(sheetName) {
  const sheet = getSheetInfo(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    obj._rowIndex = i + 1; 
    result.push(obj);
  }
  return result;
}

// ==========================================
// AUTH & DASHBOARD
// ==========================================

// Hash password dengan MD5 menggunakan Google Apps Script Utilities
function _hashPassword(plain) {
  try {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(plain), Utilities.Charset.UTF_8);
    return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
  } catch(e) { return String(plain); }
}

function doLogin(username, password) {
  try {
    const users = getRowsData('users');
    const inputHash = _hashPassword(String(password));
    const user = users.find(u => {
      const uname = String(u.username||u.Username||'').trim();
      const storedPass = String(u.password||u.Password||'');
      if (uname !== String(username).trim()) return false;
      // 1. Cek hash (standar baru)
      if (storedPass === inputHash) return true;
      // 2. Backward compat: cek plaintext lama, lalu auto-migrate ke hash
      if (storedPass === String(password)) {
        try {
          const sh = getSheetInfo('users');
          const hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
          const pIdx = hdrs.findIndex(h => String(h).toLowerCase() === 'password') + 1;
          if (pIdx > 0) sh.getRange(u._rowIndex, pIdx).setValue(inputHash);
          _invalidateCache(['users']);
        } catch(migErr) {}
        return true;
      }
      return false;
    });
    
    if (user) {
      return { success: true, user: { username: user.username||user.Username, role: user.role||user.Role, cabang: user.cabang||user.Cabang } };
    } else {
      return { success: false, message: 'Username atau Password salah!' };
    }
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getDashboardData(role, cabang, force = false) {
  try {
    const userRole = String(role).toUpperCase();
    const cacheKey = userRole === 'PUSAT' ? 'dashboard_v3_PUSAT' : `dashboard_v3_${userRole}_${cabang}`;
    
    if (!force) {
        const cached = _getServerCache(cacheKey);
        if (cached) return { success: true, data: cached };
    }

    const data = {};
    const users = getRowsData('users');
    const requests = getRowsData('request_cabang');
    
    if (userRole === 'PUSAT') {
      const snMasuk = getRowsData('sn_masuk');
      data.totalCabang = users.filter(u => String(u.role||u.Role).toUpperCase() === 'CABANG').length;
      data.totalRequest = requests.length;
      data.totalSNMasuk = snMasuk.length;
      
      const activeOrders = [];
      requests.forEach(r => {
        let rawStatus = r['Status'] || r['status'] || 'Menunggu';
        let rawItems = r['Detail Items'] || r['detail items'] || r['Data Item'] || r['data item'] || r['Data Items'] || '[]';
        if (String(rawStatus).trim().startsWith('[')) rawStatus = rawItems || 'Menunggu';
        
        if (!String(rawStatus).includes('Selesai')) {
          let itemsText = '-';
          try {
            const parsed = JSON.parse(String(r['Detail Items'] || r['detail items'] || r['Data Item'] || r['data item'] || r['Data Items'] || '[]'));
            if(Array.isArray(parsed) && parsed.length > 0) {
              itemsText = parsed.map(i => `${i.qty}x ${i.nama}`).join(', ');
            }
          } catch(e) {}
          activeOrders.push({
            id: r['ID Request'] || r['id request'] || r['No. Request'] || r['no. request'] || r['No Request'] || r['no request'] || '-',
            cabang: r['Cabang'] || r['cabang'],
            tanggal: r['Tanggal'] || r['tanggal'],
            status: rawStatus,
            items: itemsText
          });
        }
      });
      data.activeOrders = activeOrders.reverse();
      
      // Menghitung Part Paling Sering Diorder (Top 15) dari SN Masuk
      const topPartsCacheKey = 'dashboard_topParts';
      let topParts = null;
      if (!force) topParts = _getServerCache(topPartsCacheKey);
      
      const snMasukSheet = getSheetInfo('sn_masuk');
      data.totalSNMasuk = Math.max(0, snMasukSheet.getLastRow() - 1);
      
      if (!topParts) {
        // Gunakan snMasuk yang sudah difetch di atas — TIDAK fetch ulang!
        const freqMap = {};
        snMasuk.forEach(r => {
          const kode = r['Kode Barang'] || r['kode barang'] || r['Kode'] || r['kode'] || r['ID Barang'] || r['id barang'];
          const nama = r['Nama Barang'] || r['nama barang'] || r['Daftar Barang'] || r['daftar barang'] || r['Nama'] || r['nama'];
          if (kode && String(kode).trim() !== '') {
            const k = String(kode).trim();
            const nm = nama ? String(nama).trim() : 'Unknown';
            if (!freqMap[k]) freqMap[k] = { kode: k, nama: nm, count: 0, stok: 0 };
            freqMap[k].count++;
          }
        });
        
        const masterData = getRowsData('master');
        masterData.forEach(r => {
          const kode = r['Kode Barang'] || r['kode barang'] || r['Kode'] || r['kode'] || r['ID Barang'] || r['id barang'];
          if (kode) {
            const k = String(kode).trim();
            if (freqMap[k]) {
              freqMap[k].stok = parseInt(r['Stok'] || r['stok'] || 0);
            }
          }
        });
        
        topParts = Object.values(freqMap).sort((a, b) => b.count - a.count).slice(0, 15);
        _setServerCache(topPartsCacheKey, topParts); // cache for 15 mins
      }
      
      data.topParts = topParts;
      
    } else {
      const myReqs = requests.filter(r => String(r.Cabang||r.cabang).toUpperCase() === String(cabang).toUpperCase());
      
      let countProses = 0;
      let countSelesai = 0;
      
      myReqs.forEach(r => {
        let rawStatus = r['Status'] || r['status'] || 'Menunggu';
        let rawItems = r['Detail Items'] || r['detail items'] || r['Data Item'] || r['data item'] || r['Data Items'] || '[]';
        
        if (String(rawStatus).trim().startsWith('[')) {
          rawStatus = rawItems || 'Menunggu';
        }
        
        const statStr = String(rawStatus);
        if (statStr.includes('Selesai')) {
          countSelesai++;
        } else {
          countProses++;
        }
      });
      
      data.orderBulanIni = myReqs.length;
      data.proses = countProses;
      data.selesai = countSelesai;
    }
    
    const uniqueCabang = [...new Set(users.filter(u => String(u.role||u.Role).toUpperCase() === 'CABANG').map(u => u.cabang||u.Cabang))];
    data.listCabang = uniqueCabang;
    
    _setServerCache(cacheKey, data);
    return { success: true, data: data };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function createNewUser(username, password, role, cabang) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheetInfo('users');
    if (sheet.getLastRow() === 0) sheet.appendRow(['username', 'password', 'role', 'cabang']);
    
    const users = getRowsData('users');
    if (users.find(u => String(u.username||u.Username).toLowerCase() === String(username).toLowerCase())) {
      return { success: false, message: 'Username sudah digunakan!' };
    }
    
    sheet.appendRow([username, _hashPassword(password), role, cabang]);
    _invalidateCache(['users']);
    return { success: true, message: 'User baru berhasil ditambahkan!' };
  } catch (e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function getAllUsers(force = false) {
  try {
    let data;
    if (!force) {
        const cached = _getServerCache('users');
        if (cached) data = cached;
    }
    
    if (!data) {
      data = getRowsData('users');
      _setServerCache('users', data);
    }
    
    const safeData = data.map(u => ({
      username: u.username || u.Username,
      role: u.role || u.Role,
      cabang: u.cabang || u.Cabang,
      _rowIndex: u._rowIndex
    }));
    return { success: true, data: safeData };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteUserAccount(username) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheetInfo('users');
    const data = getRowsData('users');
    const user = data.find(u => String(u.username||u.Username).toLowerCase() === String(username).toLowerCase());
    
    if(!user) return { success: false, message: 'User tidak ditemukan' };
    
    sheet.deleteRow(user._rowIndex);
    _invalidateCache(['users']);
    return { success: true, message: 'User berhasil dihapus' };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function changeUserPassword(username, newPassword) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheetInfo('users');
    const data = getRowsData('users');
    const user = data.find(u => String(u.username||u.Username).toLowerCase() === String(username).toLowerCase());
    
    if(!user) return { success: false, message: 'User tidak ditemukan' };
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const passColIndex = headers.findIndex(h => String(h).toLowerCase() === 'password') + 1;
    
    if(passColIndex === 0) return { success: false, message: 'Kolom password tidak ditemukan di database' };
    
    sheet.getRange(user._rowIndex, passColIndex).setValue(_hashPassword(newPassword));
    _invalidateCache(['users']);
    return { success: true, message: 'Password berhasil diubah' };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

// ==========================================
// STOCK & MASTER DATA
// ==========================================
function getMasterData(force = false) {
  try {
    if (!force) {
        const cached = _getServerCache('master');
        if (cached) return { success: true, data: cached };
    }
    
    const data = getRowsData('master');
    _setServerCache('master', data);
    return { success: true, data: data };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function bulkUpdateMaster(importType, parsedData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const masterSheet = getSheetInfo('master');
    let masterData = masterSheet.getDataRange().getValues();
    
    if (masterData.length <= 1) {
       masterData = [['Kode Barang', 'Nama Barang', 'Harga Modal', 'Stok']];
    }
    
    const dataMap = new Map();
    masterData.forEach((row, idx) => { if(idx > 0) dataMap.set(String(row[0]).trim(), idx); });

    let updatedCount = 0;
    let newCount = 0;
    let duplicateSNs = [];

    if (importType === 'pembelian') {
      parsedData.forEach(item => {
        const kode = String(item['Kode Barang'] || item['kode barang'] || item['Kode #'] || item['kode #'] || '').trim(); 
        const nama = item['Nama Barang'] || item['nama barang']; 
        const harga = item['@Harga'] || item['@harga'] || item['Harga'] || item['harga'] || 0;
        if (kode && nama) {
          if (dataMap.has(kode)) {
            masterData[dataMap.get(kode)][2] = harga;
            updatedCount++;
          } else {
            masterData.push([kode, nama, harga, 0]);
            dataMap.set(kode, masterData.length - 1);
            newCount++;
          }
        }
      });
      masterSheet.getRange(1, 1, masterData.length, masterData[0].length).setValues(masterData);
      
    } else if (importType === 'stok') {
      const snSheet = getSheetInfo('sn_masuk');
      const existingSNData = getRowsData('sn_masuk');
      const existingSNSet = new Set(existingSNData.map(r => String(r['SN'] || r['sn'] || '').trim()));
      
      let rowsToAppend = [];
      const stockCounter = {};
      const dateInput = new Date();
      
      parsedData.forEach(item => {
        const kode = String(item['Kode Barang'] || item['kode barang'] || '').trim();
        const nama = item['Nama Barang'] || item['nama barang'] || '';
        const sn = String(item['SN'] || item['sn'] || item['Serial Number'] || '').trim();
        const nota = item['No Nota'] || item['no nota'] || item['Nota'] || 'IMPORT';
        
        if (kode && sn) {
          if (existingSNSet.has(sn)) {
            duplicateSNs.push(sn);
          } else {
            rowsToAppend.push([nota, kode, nama, sn, dateInput, '']);
            existingSNSet.add(sn);
            stockCounter[kode] = (stockCounter[kode] || 0) + 1;
            newCount++;
          }
        }
      });
      
      // Update stok di master
      for (const [kode, qty] of Object.entries(stockCounter)) {
        if (dataMap.has(kode)) {
          masterData[dataMap.get(kode)][3] = Number(masterData[dataMap.get(kode)][3] || 0) + qty;
          updatedCount++;
        } else {
          // Jika kode barang belum ada di master, tambahkan
          masterData.push([kode, 'Barang Import Baru', 0, qty]);
          dataMap.set(kode, masterData.length - 1);
          updatedCount++;
        }
      }
      
      // Simpan perubahan master
      masterSheet.getRange(1, 1, masterData.length, masterData[0].length).setValues(masterData);
      
      // Simpan SN baru
      if (rowsToAppend.length > 0) {
        if (snSheet.getLastRow() === 0) snSheet.appendRow(['No Nota', 'Kode Barang', 'Nama Barang', 'SN', 'Tanggal Input', 'Catatan']);
        snSheet.getRange(snSheet.getLastRow() + 1, 1, rowsToAppend.length, 6).setValues(rowsToAppend);
      }
      
      _invalidateCache(['sn_masuk', 'dashboard_v3_PUSAT']);
    }
    
    if(newCount === 0 && updatedCount === 0 && duplicateSNs.length === 0) {
      return { success: false, message: 'Data tidak sesuai atau kosong. Pastikan kolom Excel (Kode Barang, SN) benar.' };
    }
    
    _invalidateCache(['master']);
    
    return { success: true, newCount: newCount, updatedCount: updatedCount, duplicates: duplicateSNs, message: 'Berhasil' };
  } catch (e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

// ==========================================
// SN (SERIAL NUMBER) PUSAT - MULTI BARANG
// ==========================================

// ── SATU sumber kebenaran untuk prefix SN: dipakai oleh generate & save ──
// Urutan: KEY > ADP > BAT (spesifik dulu) → LED/LCD → fallback
// Ini mencegah "Keyboard Lenovo 14 W/LED" salah dikategorikan sebagai LED.
function _getSNPrefix(nama, kode) {
  const nm = String(nama || '').trim().toUpperCase();
  const kd = String(kode || '').trim().toUpperCase();

  if (nm.includes('KEYBOARD') || kd.includes('KEY')) {
    let mk = 'XX';
    const w = nm.split(' ');
    if (w.length > 1) { const s = w[1].replace(/[^A-Z]/g,''); if (s.length >= 2) mk = s.substring(0,2); }
    else if (kd.includes('-')) { const p = kd.split('-'); if (p.length > 2 && p[2].length >= 2) mk = p[2].substring(0,2); }
    return `KEY-${mk}`;
  }
  if (nm.includes('ADAPTOR') || nm.includes('ADAPTER') || kd.includes('ADP')) {
    let mk = 'XX';
    const w = nm.split(' ');
    if (w.length > 1) { const s = w[1].replace(/[^A-Z]/g,''); if (s.length >= 2) mk = s.substring(0,2); }
    return `ADP-${mk}`;
  }
  if (nm.includes('BATERAI') || nm.includes('BATTERY') || kd.includes('BAT')) {
    let mk = 'XX';
    const w = nm.split(' ');
    if (w.length > 1) { const s = w[1].replace(/[^A-Z]/g,''); if (s.length >= 2) mk = s.substring(0,2); }
    return `BAT-${mk}`;
  }
  if (nm.includes('LCD') || nm.includes('LED') || kd.includes('LCD') || kd.includes('LED')) {
    return 'LED';
  }
  // Fallback: 3 huruf nama + 2 huruf merk
  let ac = nm.replace(/[^A-Z]/g,'').substring(0,3);
  if (ac.length < 3) ac = 'PRT';
  let mk = 'XX';
  const w = nm.split(' ');
  if (w.length > 1) { const s = w[1].replace(/[^A-Z]/g,''); if (s.length >= 2) mk = s.substring(0,2); }
  return `${ac}-${mk}`;
}

function generateSNLogicMulti(nota, itemsArray) {
  try {
    const d = new Date();
    const y = d.getFullYear().toString().slice(-2);
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const dt = String(d.getDate()).padStart(2, '0');
    const dateStr = `${y}${m}${dt}`; 

    let allExistingSNs = [];
    const cachedSN = _getServerCache('sn_masuk');
    if (cachedSN) {
      for (const n in cachedSN) {
        cachedSN[n].forEach(r => {
          let val = r['SN'] || r['sn'] || r['Daftar SN'] || r['daftar sn'] || r['Serial Number'] || '';
          // Fix: split koma agar setiap SN individu ter-capture (SN disimpan comma-separated)
          if (val) String(val).split(',').forEach(s => { const c = s.trim().toUpperCase(); if (c) allExistingSNs.push(c); });
        });
      }
    } else {
      const snSheetData = getRowsData('sn_masuk');
      snSheetData.forEach(r => {
        let val = r['SN'] || r['sn'] || r['Daftar SN'] || r['daftar sn'] || r['Serial Number'] || '';
        if (val) String(val).split(',').forEach(s => { const c = s.trim().toUpperCase(); if (c) allExistingSNs.push(c); });
      });
    }

    let resultItems = [];

    itemsArray.forEach(item => {
      // Gunakan _getSNPrefix: fungsi terpusat agar preview & save selalu konsisten
      const prefixPart = _getSNPrefix(item.nama, item.kode);
      const prefix = `${prefixPart}-${dateStr}`;

      const existingPrefixSNs = allExistingSNs
        .filter(sn => sn.startsWith(prefix))
        .map(sn => {
          const numStr = sn.substring(prefix.length);
          return parseInt(numStr);
        })
        .filter(num => !isNaN(num));

      let generated = [];
      let currentNumber = 1;

      for (let i = 0; i < item.qty; i++) {
        while (existingPrefixSNs.includes(currentNumber)) {
          currentNumber++;
        }
        
        const numStr = currentNumber.toString().padStart(3, '0');
        const newSN = `${prefix}${numStr}`;
        
        generated.push(newSN);
        existingPrefixSNs.push(currentNumber);
        allExistingSNs.push(newSN);
      }

      resultItems.push({ kode: item.kode, nama: item.nama, qty: item.qty, sns: generated });
    });

    return { success: true, data: resultItems };
  } catch(e) { 
    return { success: false, message: e.toString() }; 
  }
}

// 1. Perbarui saveSNMasukMulti
function saveSNMasukMulti(nota, generatedItems) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheetInfo('sn_masuk');
    if (sheet.getLastRow() === 0) sheet.appendRow(['No Nota', 'Kode Barang', 'Nama Barang', 'SN', 'Tanggal Input', 'Catatan']);
    
    // RE-CHECK DUPLICATES INSIDE LOCK TO PREVENT COLLISIONS
    const currentData = sheet.getDataRange().getValues();
    const existingSNs = new Set();
    if (currentData.length > 1) {
      const headers = currentData[0];
      const snIdx = headers.findIndex(h => String(h).toLowerCase().trim() === 'sn' || String(h).toLowerCase().trim() === 'daftar sn');
      if (snIdx !== -1) {
        for (let i = 1; i < currentData.length; i++) {
          const val = currentData[i][snIdx];
          if (val) {
            String(val).split(',').forEach(s => {
               const clean = s.trim().toUpperCase();
               if(clean) existingSNs.add(clean);
            });
          }
        }
      }
    }
    
    const d = new Date();
    const y = d.getFullYear().toString().slice(-2);
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const dt = String(d.getDate()).padStart(2, '0');
    const dateStr = `${y}${m}${dt}`; 
    
    const dateInput = new Date();
    let rowsToAppend = [];
    
    generatedItems.forEach(item => {
      // Gunakan _getSNPrefix: fungsi terpusat agar konsisten dengan generateSNLogicMulti
      const prefixPart = _getSNPrefix(item.nama, item.kode);
      const prefix = `${prefixPart}-${dateStr}`;

      let updatedSns = [];
      item.sns.forEach(sn => {
        let finalSN = sn;
        
        // COLLISION CHECK
        if (existingSNs.has(finalSN)) {
           let maxNum = 0;
           existingSNs.forEach(esn => {
             if (esn.startsWith(prefix)) {
               const numStr = esn.substring(prefix.length);
               const num = parseInt(numStr);
               if (!isNaN(num) && num > maxNum) maxNum = num;
             }
           });
           const currentNumber = maxNum + 1;
           finalSN = `${prefix}${currentNumber.toString().padStart(3, '0')}`;
        }
        
        existingSNs.add(finalSN);
        updatedSns.push(finalSN);
      });
      item.sns = updatedSns; // Update the returned object with corrected SNs
      rowsToAppend.push([nota, item.kode, item.nama, updatedSns.join(', '), dateInput, '']);
    });
    
    if (rowsToAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 6).setValues(rowsToAppend);
      
      // MENGUPDATE STOK MASTER
      const masterSheet = getSheetInfo('master');
      const masterData = masterSheet.getDataRange().getValues();
      if (masterData.length > 1) {
        const mHeaders = masterData[0];
        const kdIdx = mHeaders.findIndex(h => String(h).toLowerCase().trim() === 'kode barang' || String(h).toLowerCase().trim() === 'kode');
        const stokIdx = mHeaders.findIndex(h => String(h).toLowerCase().trim() === 'stok' || String(h).toLowerCase().trim() === 'stock');
        
        if (kdIdx !== -1 && stokIdx !== -1) {
           let masterUpdated = false;
           generatedItems.forEach(item => {
              const qtyToAdd = item.sns.length;
              for (let i = 1; i < masterData.length; i++) {
                 if (String(masterData[i][kdIdx]).trim().toUpperCase() === String(item.kode).trim().toUpperCase()) {
                    let currentStok = parseInt(masterData[i][stokIdx]) || 0;
                    masterData[i][stokIdx] = currentStok + qtyToAdd;
                    masterUpdated = true;
                    break;
                 }
              }
           });
           
           if (masterUpdated) {
             masterSheet.getRange(1, 1, masterData.length, masterData[0].length).setValues(masterData);
           }
        }
      }
    }
    
    // Invalidate cache secara menyeluruh terkait SN Masuk & Dashboard & Master
    _invalidateCache(['sn_masuk', 'masterData', 'master', 'dashboard_v3_PUSAT', 'dashboard_topParts']);
    return { success: true, message: 'Semua SN berhasil disimpan ke database & stok ditambahkan!', finalItems: generatedItems };
  } catch (e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function getSNMasuk(force = false) {
  try {
    if (!force) {
        const cached = _getServerCache('sn_masuk');
        if (cached) return { success: true, data: cached };
    }

    const data = getRowsData('sn_masuk');
    
    // Urutkan data berdasarkan Tanggal Input descending (nota terbaru di atas)
    data.sort((a, b) => {
      const ta = new Date(a['Tanggal Input'] || a['tanggal input'] || a['Tanggal'] || a['tanggal'] || 0);
      const tb = new Date(b['Tanggal Input'] || b['tanggal input'] || b['Tanggal'] || b['tanggal'] || 0);
      return tb - ta;
    });
    
    const grouped = {};
    data.forEach(row => {
      const nota = row['No Nota'] || row['no nota'] || row['No. Nota'] || row['no. nota'] || row['NO NOTA'] || '';
      if(nota) {
        if (!grouped[nota]) grouped[nota] = [];
        grouped[nota].push(row);
      }
    });
    
    _setServerCache('sn_masuk', grouped);
    return { success: true, data: grouped };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function deleteSNNota(nota) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheetInfo('sn_masuk');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, message: 'Data kosong' };
    
    const headers = data[0];
    const notaColIdx = headers.findIndex(h => String(h).toLowerCase().replace(/[^a-z]/g, '') === 'nonota');
    const kodeColIdx = headers.findIndex(h => String(h).toLowerCase().trim() === 'kode barang' || String(h).toLowerCase().trim() === 'kode');
    const snColIdx = headers.findIndex(h => String(h).toLowerCase().trim() === 'sn' || String(h).toLowerCase().trim() === 'daftar sn');
    
    if (notaColIdx === -1) return { success: false, message: 'Kolom No Nota tidak ditemukan' };
    
    const masterSheet = getSheetInfo('master');
    const masterData = masterSheet.getDataRange().getValues();
    const mHeaders = masterData[0] || [];
    const mKdIdx = mHeaders.findIndex(h => String(h).toLowerCase().trim() === 'kode barang' || String(h).toLowerCase().trim() === 'kode');
    const mStokIdx = mHeaders.findIndex(h => String(h).toLowerCase().trim() === 'stok' || String(h).toLowerCase().trim() === 'stock');
    
    let masterUpdated = false;
    let deletedCount = 0;
    
    for (let i = data.length - 1; i > 0; i--) {
      if (String(data[i][notaColIdx]).trim() === String(nota).trim()) {
        // Kurangi stok di master
        if (kodeColIdx !== -1 && snColIdx !== -1 && mKdIdx !== -1 && mStokIdx !== -1) {
            const kodeToMinus = String(data[i][kodeColIdx]).trim().toUpperCase();
            const snStr = String(data[i][snColIdx] || '').trim();
            const qtyToMinus = snStr ? snStr.split(',').filter(x => x.trim() !== '').length : 0;
            
            for (let j = 1; j < masterData.length; j++) {
                if (String(masterData[j][mKdIdx]).trim().toUpperCase() === kodeToMinus) {
                    let currentStok = parseInt(masterData[j][mStokIdx]) || 0;
                    masterData[j][mStokIdx] = Math.max(0, currentStok - qtyToMinus);
                    masterUpdated = true;
                    break;
                }
            }
        }
        
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    
    if (masterUpdated) {
        masterSheet.getRange(1, 1, masterData.length, masterData[0].length).setValues(masterData);
    }
    
    _invalidateCache(['sn_masuk', 'masterData', 'master', 'dashboard_v3_PUSAT', 'dashboard_topParts']);
    return { success: true, message: `Berhasil menghapus ${deletedCount} item untuk Nota ${nota} & stok telah dikurangi` };
  } catch (e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function updateSNNota(nota, updatedItems) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheetInfo('sn_masuk');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, message: 'Data kosong' };
    
    const headers = data[0];
    const notaColIdx = headers.findIndex(h => String(h).toLowerCase().replace(/[^a-z]/g, '') === 'nonota');
    const kodeColIdx = headers.findIndex(h => String(h).toLowerCase().trim() === 'kode barang' || String(h).toLowerCase().trim() === 'kode');
    const snColIdx = headers.findIndex(h => String(h).toLowerCase().trim() === 'sn' || String(h).toLowerCase().trim() === 'daftar sn');
    
    if (notaColIdx === -1) return { success: false, message: 'Kolom No Nota tidak ditemukan' };
    
    const masterSheet = getSheetInfo('master');
    const masterData = masterSheet.getDataRange().getValues();
    const mHeaders = masterData[0] || [];
    const mKdIdx = mHeaders.findIndex(h => String(h).toLowerCase().trim() === 'kode barang' || String(h).toLowerCase().trim() === 'kode');
    const mStokIdx = mHeaders.findIndex(h => String(h).toLowerCase().trim() === 'stok' || String(h).toLowerCase().trim() === 'stock');
    
    const catatanColIdx = headers.findIndex(h => String(h).toLowerCase().trim() === 'catatan');
    let masterUpdated = false;
    let savedCatatan = ''; // Preservasi catatan sebelum delete untuk di-restore
    
    // 1. Delete lama & Kurangi stok
    for (let i = data.length - 1; i > 0; i--) {
      if (String(data[i][notaColIdx]).trim() === String(nota).trim()) {
        // Simpan catatan sebelum baris dihapus (ambil dari baris mana saja)
        if (!savedCatatan && catatanColIdx !== -1 && data[i][catatanColIdx]) {
          savedCatatan = String(data[i][catatanColIdx]);
        }
        // Kurangi stok lama
        if (kodeColIdx !== -1 && snColIdx !== -1 && mKdIdx !== -1 && mStokIdx !== -1) {
            const kodeToMinus = String(data[i][kodeColIdx]).trim().toUpperCase();
            const snStr = String(data[i][snColIdx] || '').trim();
            const qtyToMinus = snStr ? snStr.split(',').filter(x => x.trim() !== '').length : 0;
            
            for (let j = 1; j < masterData.length; j++) {
                if (String(masterData[j][mKdIdx]).trim().toUpperCase() === kodeToMinus) {
                    let currentStok = parseInt(masterData[j][mStokIdx]) || 0;
                    masterData[j][mStokIdx] = Math.max(0, currentStok - qtyToMinus);
                    masterUpdated = true;
                    break;
                }
            }
        }
        sheet.deleteRow(i + 1);
      }
    }
    
    // 2. Insert baru & Tambah stok
    let rowsToAppend = [];
    updatedItems.forEach(item => {
      const dt = item.tanggal ? new Date(item.tanggal) : new Date();
      // item.sn bisa berupa string koma jika diedit.
      const snStr = String(item.sn || '').trim();
      const qtyToAdd = snStr ? snStr.split(',').filter(x => x.trim() !== '').length : 0;
      
      if (mKdIdx !== -1 && mStokIdx !== -1) {
          for (let j = 1; j < masterData.length; j++) {
              if (String(masterData[j][mKdIdx]).trim().toUpperCase() === String(item.kode).trim().toUpperCase()) {
                  let currentStok = parseInt(masterData[j][mStokIdx]) || 0;
                  masterData[j][mStokIdx] = currentStok + qtyToAdd;
                  masterUpdated = true;
                  break;
              }
          }
      }
      
      rowsToAppend.push([nota, item.kode, item.nama, snStr, dt, savedCatatan]);
    });
    
    if (rowsToAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 6).setValues(rowsToAppend);
    }
    
    if (masterUpdated) {
        masterSheet.getRange(1, 1, masterData.length, masterData[0].length).setValues(masterData);
    }
    
    _invalidateCache(['sn_masuk', 'masterData', 'master', 'dashboard_v3_PUSAT', 'dashboard_topParts']);
    return { success: true, message: `Nota ${nota} berhasil diperbarui & stok disesuaikan!` };
  } catch (e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function updateCatatanNota(nota, catatan) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheetInfo('sn_masuk');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, message: 'Data kosong' };
    
    const headers = data[0];
    const notaColIdx = headers.findIndex(h => String(h).toLowerCase().replace(/[^a-z]/g,'') === 'nonota');
    let catatanColIdx = headers.findIndex(h => String(h).toLowerCase().trim() === 'catatan');
    
    if (notaColIdx === -1) return { success: false, message: 'Kolom No Nota tidak ditemukan' };
    
    // Jika kolom Catatan belum ada di sheet, buat otomatis
    if (catatanColIdx === -1) {
      const newColNum = headers.length + 1;
      sheet.getRange(1, newColNum).setValue('Catatan');
      catatanColIdx = headers.length;
    }
    
    // Update semua baris yang nota-nya cocok
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][notaColIdx]).trim() === String(nota).trim()) {
        sheet.getRange(i + 1, catatanColIdx + 1).setValue(catatan);
      }
    }
    
    _invalidateCache(['sn_masuk']);
    return { success: true, message: 'Catatan tersimpan' };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

// ==========================================
// REQUEST & ALOKASI
// ==========================================
function saveRequestCabang(cabang, cartData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheetInfo('request_cabang');
    if (sheet.getLastRow() === 0) sheet.appendRow(['ID Request', 'Tanggal', 'Cabang', 'Detail Items', 'Status']);
    
    const reqId = 'REQ-' + String(cabang).substring(0,3).toUpperCase() + '-' + Date.now().toString().slice(-6);
    sheet.appendRow([reqId, new Date(), cabang, JSON.stringify(cartData), 'Menunggu']);
    
    _invalidateCache(['all_request_cabang', 'request_cabang', 'dashboard_v3_PUSAT', `dashboard_v3_CABANG_${cabang}`]);
    return { success: true, reqId: reqId };
  } catch (e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function getRequestCabang(role, cabang, force = false) {
  try {
    const cacheKey = 'all_request_cabang';
    let data;
    if (!force) data = _getServerCache(cacheKey);
    
    if (!data) {
      data = getRowsData('request_cabang');
      _setServerCache(cacheKey, data);
    }

    if (String(role).toUpperCase() === 'CABANG') {
      data = data.filter(r => String(r.Cabang||r.cabang).toUpperCase() === String(cabang).toUpperCase());
    }
    
    const result = [...data].reverse();
    return { success: true, data: result }; 
  } catch (e) { return { success: false, message: e.toString() }; }
}

// 2. Perbarui prosesAlokasiSN agar cache riwayat & request ikut bersih
function prosesAlokasiSN(reqId, rowIndex, ioNumber, snList, keterangan, targetType, targetValue, cabangReq) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const reqSheet = getSheetInfo('request_cabang');
    const headers = reqSheet.getDataRange().getValues()[0] || [];
    const statusColIdx = headers.findIndex(h => String(h).toLowerCase().trim() === 'status') + 1;
    
    if (rowIndex && statusColIdx > 0) {
      const finalStatus = keterangan ? `Selesai - ${keterangan}` : 'Selesai';
      reqSheet.getRange(rowIndex, statusColIdx).setValue(finalStatus);
    }
    
    const riwSheet = getSheetInfo('riwayat_sn');
    if (riwSheet.getLastRow() === 0) riwSheet.appendRow(['ID Request', 'No IO', 'Tujuan', 'Tipe Tujuan', 'SN', 'Keterangan', 'Tanggal Alokasi']);
    
    const masterSheet = getSheetInfo('master');
    const masterFullData = masterSheet.getDataRange().getValues();
    
    const snData = getRowsData('sn_masuk');
    const snMap = new Map();
    // Fix: split comma-separated SNs & normalisasi UPPERCASE agar lookup per-SN berhasil
    snData.forEach(r => {
      const snStr = r['SN'] || r['sn'] || r['Daftar SN'] || r['daftar sn'];
      const kodeBrg = r['Kode Barang'] || r['kode barang'];
      if (snStr && kodeBrg) {
        String(snStr).split(',').forEach(s => {
          const cleanSN = s.trim().toUpperCase();
          if (cleanSN) snMap.set(cleanSN, String(kodeBrg).trim());
        });
      }
    });

    const dateNow = new Date();
    // Fix: normalisasi SN input ke UPPERCASE agar cocok dengan snMap
    const snArray = String(snList).split(',').map(s => s.trim().toUpperCase()).filter(s => s);
    
    let riwRowsToAppend = [];
    let errorLog = [];
    
    snArray.forEach(sn => {
      riwRowsToAppend.push([reqId || '-', ioNumber, targetValue, targetType, sn, keterangan, dateNow]);
      
      const kodeBrg = snMap.get(sn);
      if (kodeBrg) {
        // Fix: search dari index 1 (skip header) & strict string comparison
        const rowIdx = masterFullData.findIndex((r, idx) => idx > 0 && String(r[0]).trim().toUpperCase() === String(kodeBrg).trim().toUpperCase());
        if (rowIdx > 0) {
          masterFullData[rowIdx][3] = Math.max(0, parseInt(masterFullData[rowIdx][3] || 0) - 1);
        } else {
          errorLog.push(sn);
        }
      } else {
        errorLog.push(sn);
      }
    });
    
    if (riwRowsToAppend.length > 0) {
      riwSheet.getRange(riwSheet.getLastRow() + 1, 1, riwRowsToAppend.length, 7).setValues(riwRowsToAppend);
    }
    
    masterSheet.getRange(1, 1, masterFullData.length, masterFullData[0].length).setValues(masterFullData);
    
    // Invalidate seluruh cache terkait request, riwayat, dan master stok
    _invalidateCache(['all_request_cabang', 'request_cabang', 'masterData', 'master', 'dashboard_v3_PUSAT', `dashboard_v3_CABANG_${cabangReq}`]);
    
    let msg = 'Alokasi berhasil & Stok Terpotong.';
    if (errorLog.length > 0) msg += ` (Info: SN ${errorLog.join(', ')} tidak terdeteksi kode barangnya, stok tidak dipotong).`;
    
    return { success: true, message: msg };
  } catch (e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function getRiwayatSN(force = false) {
  try {
    if (!force) {
        const cached = _getServerCache('riwayat_sn');
        if (cached) return { success: true, data: cached };
    }

    const data = getRowsData('riwayat_sn').reverse();
    _setServerCache('riwayat_sn', data);
    return { success: true, data: data }; 
  } catch (e) { return { success: false, message: e.toString() }; }
}

// ==========================================
// UPDATE HARGA MODAL (HANYA PUSAT)
// ==========================================
function updateHargaModal(kodeBarang, hargaBaru) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheetInfo('master');
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) return { success: false, message: 'Data master kosong' };
    
    const headers = data[0];
    const kdIdx = headers.findIndex(h => String(h).toLowerCase().trim() === 'kode barang' || String(h).toLowerCase().trim() === 'kode');
    const hrgIdx = headers.findIndex(h => String(h).toLowerCase().trim() === 'harga modal' || String(h).toLowerCase().trim() === 'harga');
    
    if (kdIdx === -1 || hrgIdx === -1) return { success: false, message: 'Kolom Harga Modal tidak ditemukan' };
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][kdIdx]).trim().toUpperCase() === String(kodeBarang).trim().toUpperCase()) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, message: 'Barang tidak ditemukan' };
    
    sheet.getRange(rowIndex, hrgIdx + 1).setValue(hargaBaru);
    
    _invalidateCache(['masterData', 'master', 'dashboard_v3_PUSAT']);
    return { success: true, message: 'Harga berhasil diperbarui' };
  } catch (e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function deleteRequestCabang(reqId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheetInfo('request_cabang');
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, message: 'Data request kosong' };
    
    const headers = data[0];
    const idIdx = headers.findIndex(h => {
      const hStr = String(h).toLowerCase().trim();
      return hStr === 'id request' || hStr === 'id' || hStr === 'no. request' || hStr === 'no request';
    });
    
    if (idIdx === -1) return { success: false, message: 'Kolom ID Request tidak ditemukan' };
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]).trim().toUpperCase() === String(reqId).trim().toUpperCase()) {
        rowIndex = i + 1; // +1 for 1-based sheet row indexing
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, message: 'Request tidak ditemukan' };
    
    sheet.deleteRow(rowIndex);
    
    _invalidateCache(['all_request_cabang', 'request_cabang', 'dashboard_v3_PUSAT']);
    return { success: true, message: 'Request berhasil dihapus' };
  } catch (e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

// ==========================================
// CHAT SYSTEM
// ==========================================

function getChatUnreadCount(role, cabang) {
  try {
    const isPusat = String(role).toUpperCase() === 'PUSAT';
    const cacheKey = isPusat ? 'chat_unread_PUSAT' : `chat_unread_CABANG_${cabang}`;
    const cached = _getServerCache(cacheKey);
    if (cached !== null) return { success: true, count: cached };

    const data = getRowsData('chat_history');
    if (!data || data.length === 0) return { success: true, count: 0 };
    
    let count = 0;
    
    data.forEach(row => {
      if (isPusat) {
        if (String(row['Pengirim'] || '').toUpperCase() !== 'PUSAT' && String(row['Status_Read_Pusat'] || '').toUpperCase() !== 'TRUE') {
          count++;
        }
      } else {
        if (String(row['Thread_Cabang'] || '').toUpperCase() === String(cabang).toUpperCase() && 
            String(row['Pengirim'] || '').toUpperCase() === 'PUSAT' && 
            String(row['Status_Read_Cabang'] || '').toUpperCase() !== 'TRUE') {
          count++;
        }
      }
    });
    
    _setServerCache(cacheKey, count);
    return { success: true, count: count };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getChatList(role, cabang) {
  try {
    const data = getRowsData('chat_history');
    const isPusat = String(role).toUpperCase() === 'PUSAT';
    
    if (isPusat) {
      const branches = {};
      const usersData = getRowsData('users');
      usersData.forEach(u => {
        if (String(u.role || u.Role).toUpperCase() === 'CABANG') {
          const c = String(u.cabang || u.Cabang).toUpperCase();
          branches[c] = { cabang: c, unread: 0, lastMessage: '', lastTime: null };
        }
      });
      
      data.forEach(row => {
        const c = String(row['Thread_Cabang'] || '').toUpperCase();
        if (c && branches[c]) {
          if (String(row['Pengirim'] || '').toUpperCase() !== 'PUSAT' && String(row['Status_Read_Pusat'] || '').toUpperCase() !== 'TRUE') {
            branches[c].unread++;
          }
          const t = new Date(row['Waktu']);
          if (!isNaN(t.getTime()) && (!branches[c].lastTime || t > branches[c].lastTime)) {
            branches[c].lastTime = t;
            branches[c].lastMessage = row['Pesan'];
          }
        }
      });
      
      const list = Object.values(branches).sort((a, b) => {
        if (a.lastTime && b.lastTime) return b.lastTime - a.lastTime;
        if (a.lastTime) return -1;
        if (b.lastTime) return 1;
        return a.cabang.localeCompare(b.cabang);
      }).map(c => ({
        ...c,
        lastTime: c.lastTime ? c.lastTime.toISOString() : null
      }));
      return { success: true, data: list };
      
    } else {
      return getChatMessages(cabang, role);
    }
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getChatMessages(cabang, role) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheetInfo('chat_history');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['ID_Pesan', 'Waktu', 'Thread_Cabang', 'Pengirim', 'Pesan', 'Status_Read_Pusat', 'Status_Read_Cabang']);
      return { success: true, data: [] };
    }
    
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    const headers = data[0];
    
    const threadIdx = headers.findIndex(h => String(h).toLowerCase() === 'thread_cabang');
    const readPusatIdx = headers.findIndex(h => String(h).toLowerCase() === 'status_read_pusat');
    const readCabangIdx = headers.findIndex(h => String(h).toLowerCase() === 'status_read_cabang');
    
    const isPusat = String(role).toUpperCase() === 'PUSAT';
    let needsUpdate = false;
    const messages = [];
    
    for (let i = 1; i < data.length; i++) {
      const rowCabang = String(data[i][threadIdx] || '').toUpperCase();
      if (rowCabang === String(cabang).toUpperCase()) {
        const rawWaktu = data[i][1];
        const safeWaktu = rawWaktu instanceof Date ? rawWaktu.toISOString() : rawWaktu;
        
        messages.push({
          id: data[i][0],
          waktu: safeWaktu,
          cabang: data[i][2],
          pengirim: data[i][3],
          pesan: data[i][4]
        });
        
        if (isPusat && String(data[i][3]).toUpperCase() !== 'PUSAT' && String(data[i][readPusatIdx]).toUpperCase() !== 'TRUE') {
          data[i][readPusatIdx] = 'TRUE';
          needsUpdate = true;
        } else if (!isPusat && String(data[i][3]).toUpperCase() === 'PUSAT' && String(data[i][readCabangIdx]).toUpperCase() !== 'TRUE') {
          data[i][readCabangIdx] = 'TRUE';
          needsUpdate = true;
        }
      }
    }
    
    if (needsUpdate) {
      dataRange.setValues(data);
      if (isPusat) _invalidateCache(['chat_unread_PUSAT']);
      else _invalidateCache([`chat_unread_CABANG_${cabang}`]);
    }
    
    return { success: true, data: messages };
  } catch(e) {
    return { success: false, message: e.toString() };
  } finally {
    if(lock.hasLock()) lock.releaseLock();
  }
}

function saveChatMessage(threadCabang, role, username, message) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheetInfo('chat_history');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['ID_Pesan', 'Waktu', 'Thread_Cabang', 'Pengirim', 'Pesan', 'Status_Read_Pusat', 'Status_Read_Cabang']);
    }
    
    const isPusat = String(role).toUpperCase() === 'PUSAT';
    const id = Utilities.getUuid();
    const time = new Date();
    const pengirim = isPusat ? 'PUSAT' : username;
    
    const statusReadPusat = isPusat ? 'TRUE' : 'FALSE';
    const statusReadCabang = isPusat ? 'FALSE' : 'TRUE';
    
    sheet.appendRow([id, time, String(threadCabang).toUpperCase(), pengirim, message, statusReadPusat, statusReadCabang]);
    
    if (isPusat) _invalidateCache([`chat_unread_CABANG_${threadCabang}`]);
    else _invalidateCache(['chat_unread_PUSAT']);
    
    return { success: true, message: 'Terkirim', data: { id, waktu: time.toISOString(), cabang: threadCabang, pengirim, pesan: message } };
  } catch(e) {
    return { success: false, message: e.toString() };
  } finally {
    if(lock.hasLock()) lock.releaseLock();
  }
}