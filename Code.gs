// ============================================================
//  Excellent Mandarin Course – Google Apps Script Backend
//  PT Global Sinar Dunia
//
//  SETUP:
//  1. Paste this entire file into Google Apps Script editor
//  2. Deploy → Web App → Execute as Me, Anyone access
//  3. Copy the /exec URL → paste into api.js SCRIPT_URL
//  4. Open admin.html → browser console → ketik: API.setup()
//     Ini akan buat semua sheet otomatis.
// ============================================================

const SS   = SpreadsheetApp.getActiveSpreadsheet();
const FOLDER_NAME = 'EMC_BuktiBayar';

// ── Helpers ─────────────────────────────────────────────────
function uid() {
  return Utilities.getUuid().replace(/-/g,'').slice(0,16);
}
function getSheet(name) {
  return SS.getSheetByName(name) || SS.insertSheet(name);
}
function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}
function ok(data)  { return ContentService.createTextOutput(JSON.stringify({ status:'ok', data })).setMimeType(ContentService.MimeType.JSON); }
function err(msg)  { return ContentService.createTextOutput(JSON.stringify({ status:'error', message: msg })).setMimeType(ContentService.MimeType.JSON); }
function cors(out) { return out; }

// ── Entry Points ─────────────────────────────────────────────
function doGet(e) {
  try {
    const p = e.parameter || {};
    switch (p.action) {
      case 'getMuridByLink': return cors(ok(getMuridByLink(p.link_id)));
      default:               return cors(err('Unknown GET action: ' + p.action));
    }
  } catch(ex) { return cors(err(ex.message)); }
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    switch (p.action) {
      // SETUP
      case 'setup':               return cors(ok(setupSheets()));
      // MURID
      case 'getMurid':            return cors(ok(getMurid(p)));
      case 'addMurid':            return cors(ok(addMurid(p)));
      case 'updateMurid':         return cors(ok(updateMurid(p)));
      case 'deleteMurid':         return cors(ok(deleteMurid(p.id)));
      // GURU
      case 'getGuru':             return cors(ok(getGuru()));
      case 'addGuru':             return cors(ok(addGuru(p)));
      case 'updateGuru':          return cors(ok(updateGuru(p)));
      case 'deleteGuru':          return cors(ok(deleteGuru(p.id)));
      // LAPORAN
      case 'getLaporan':          return cors(ok(getLaporan(p)));
      case 'addLaporan':          return cors(ok(addLaporan(p)));
      case 'deleteLaporan':       return cors(ok(deleteLaporan(p.id)));
      // ABSENSI
      case 'getAbsensi':          return cors(ok(getAbsensi(p)));
      case 'saveAbsensi':         return cors(ok(saveAbsensi(p.tanggal, p.data)));
      // INVOICE
      case 'getInvoice':          return cors(ok(getInvoice(p)));
      case 'generateInvoice':     return cors(ok(generateInvoice(p.murid_id, p.bulan, p.catatan)));
      case 'updateInvoiceStatus': return cors(ok(updateInvoiceStatus(p.id, p.status, p.tgl_kirim)));
      case 'deleteInvoice':       return cors(ok(deleteInvoice(p.id)));
      case 'uploadBukti':         return cors(ok(uploadBukti(p)));
      default:                    return cors(err('Unknown action: ' + p.action));
    }
  } catch(ex) { return cors(err(ex.message)); }
}

// ══════════════════════════════════════════════════════════════
//  SETUP — buat semua sheet dengan header yang benar
// ══════════════════════════════════════════════════════════════
function setupSheets() {
  const schemas = {
    'Murid':   ['id','nama','tipe_program','mode','program','fee_per_lesson','wa_ortu','wa_laporan','tgl_masuk','aktif','link_id'],
    'Guru':    ['id','nama','id_karyawan','jabatan','status_ptkp','tgl_bergabung','honor_onsite','honor_online','wa','bagian'],
    'Laporan': ['id','murid_id','guru_id','tanggal','subject','analisis','catatan','timestamp'],
    'Absensi': ['id','murid_id','tanggal','status','catatan'],
    'Invoice': ['id','murid_id','bulan','sesi','nominal','status','tgl_kirim','bukti_url','catatan'],
  };
  const results = [];
  for (const [name, headers] of Object.entries(schemas)) {
    const sheet = getSheet(name);
    const existing = sheet.getLastRow() > 0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] : [];
    if (existing.length === 0 || existing[0] === '') {
      sheet.clearContents();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      // Style header row
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#C8281E').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
      results.push(name + ': created');
    } else {
      results.push(name + ': already exists');
    }
  }
  return results;
}

// ══════════════════════════════════════════════════════════════
//  MURID
// ══════════════════════════════════════════════════════════════
function getMurid(p) {
  const rows = sheetToObjects(getSheet('Murid'));
  if (p.id)     return rows.find(r => r.id === p.id) || null;
  if (p.aktif)  return rows.filter(r => r.aktif === p.aktif);
  return rows;
}

function getMuridByLink(link_id) {
  const rows = sheetToObjects(getSheet('Murid'));
  const murid = rows.find(r => r.link_id === link_id);
  if (!murid) throw new Error('Murid tidak ditemukan untuk link: ' + link_id);
  return murid;
}

function addMurid(p) {
  const sheet = getSheet('Murid');
  const id      = uid();
  const link_id = p.nama.toLowerCase().replace(/\s+/g,'-') + '-' + id.slice(0,6);
  sheet.appendRow([
    id,
    p.nama        || '',
    p.tipe_program|| '',
    p.mode        || '',
    p.program     || '',
    Number(p.fee_per_lesson) || 0,
    p.wa_ortu     || '',
    p.wa_laporan  || '',
    p.tgl_masuk   || '',
    p.aktif       || 'aktif',
    link_id,
  ]);
  return { id, link_id };
}

function updateMurid(p) {
  const sheet = getSheet('Murid');
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === p.id) {
      const fields = { nama:p.nama, tipe_program:p.tipe_program, mode:p.mode, program:p.program,
                       fee_per_lesson:Number(p.fee_per_lesson)||0, wa_ortu:p.wa_ortu,
                       wa_laporan:p.wa_laporan, aktif:p.aktif };
      for (const [key, val] of Object.entries(fields)) {
        const col = headers.indexOf(key);
        if (col >= 0 && val !== undefined) sheet.getRange(i+1, col+1).setValue(val);
      }
      return { updated: p.id };
    }
  }
  throw new Error('Murid tidak ditemukan: ' + p.id);
}

function deleteMurid(id) {
  const sheet = getSheet('Murid');
  const data  = sheet.getDataRange().getValues();
  const idIdx = data[0].indexOf('id');
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][idIdx] === id) { sheet.deleteRow(i+1); return { deleted: id }; }
  }
  throw new Error('Murid tidak ditemukan');
}

// ══════════════════════════════════════════════════════════════
//  GURU
// ══════════════════════════════════════════════════════════════
function getGuru() {
  return sheetToObjects(getSheet('Guru'));
}

function addGuru(p) {
  const sheet = getSheet('Guru');
  const id = uid();
  sheet.appendRow([
    id,
    p.nama         || '',
    p.id_karyawan  || '',
    p.jabatan      || 'Teacher',
    p.status_ptkp  || 'TK0',
    p.tgl_bergabung|| '',
    Number(p.honor_onsite) || 0,
    Number(p.honor_online) || 0,
    p.wa           || '',
    p.bagian       || '1',
  ]);
  return { id };
}

function updateGuru(p) {
  const sheet = getSheet('Guru');
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === p.id) {
      const fields = { nama:p.nama, id_karyawan:p.id_karyawan, jabatan:p.jabatan,
                       status_ptkp:p.status_ptkp, tgl_bergabung:p.tgl_bergabung,
                       honor_onsite:Number(p.honor_onsite)||0,
                       honor_online:Number(p.honor_online)||0, wa:p.wa };
      for (const [key, val] of Object.entries(fields)) {
        const col = headers.indexOf(key);
        if (col >= 0 && val !== undefined) sheet.getRange(i+1, col+1).setValue(val);
      }
      return { updated: p.id };
    }
  }
  throw new Error('Guru tidak ditemukan');
}

function deleteGuru(id) {
  const sheet = getSheet('Guru');
  const data  = sheet.getDataRange().getValues();
  const idIdx = data[0].indexOf('id');
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][idIdx] === id) { sheet.deleteRow(i+1); return { deleted: id }; }
  }
  throw new Error('Guru tidak ditemukan');
}

// ══════════════════════════════════════════════════════════════
//  LAPORAN
// ══════════════════════════════════════════════════════════════
function getLaporan(p) {
  let rows = sheetToObjects(getSheet('Laporan'));
  if (p.murid_id) rows = rows.filter(r => r.murid_id === p.murid_id);
  if (p.guru_id)  rows = rows.filter(r => r.guru_id  === p.guru_id);
  if (p.bulan)    rows = rows.filter(r => String(r.tanggal).startsWith(p.bulan));
  return rows.sort((a,b) => b.tanggal > a.tanggal ? 1 : -1);
}

function addLaporan(p) {
  const sheet = getSheet('Laporan');
  const id = uid();
  sheet.appendRow([
    id,
    p.murid_id || '',
    p.guru_id  || '',
    p.tanggal  || '',
    p.subject  || '',
    p.analisis || '',
    p.catatan  || '',
    new Date().toISOString(),
  ]);
  return { id };
}

function deleteLaporan(id) {
  const sheet = getSheet('Laporan');
  const data  = sheet.getDataRange().getValues();
  const idIdx = data[0].indexOf('id');
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][idIdx] === id) { sheet.deleteRow(i+1); return { deleted: id }; }
  }
  throw new Error('Laporan tidak ditemukan');
}

// ══════════════════════════════════════════════════════════════
//  ABSENSI
// ══════════════════════════════════════════════════════════════
function getAbsensi(p) {
  let rows = sheetToObjects(getSheet('Absensi'));
  if (p.tanggal)  rows = rows.filter(r => r.tanggal === p.tanggal);
  if (p.murid_id) rows = rows.filter(r => r.murid_id === p.murid_id);
  if (p.bulan)    rows = rows.filter(r => String(r.tanggal).startsWith(p.bulan));
  return rows;
}

function saveAbsensi(tanggal, data) {
  const sheet = getSheet('Absensi');
  const existing = sheet.getDataRange().getValues();
  const headers  = existing[0];
  const tglIdx   = headers.indexOf('tanggal');
  const midIdx   = headers.indexOf('murid_id');
  // Remove existing entries for this date
  for (let i = existing.length - 1; i >= 1; i--) {
    if (existing[i][tglIdx] === tanggal) sheet.deleteRow(i+1);
  }
  // Insert new entries
  data.forEach(d => {
    sheet.appendRow([ uid(), d.murid_id, tanggal, d.status, d.catatan || '' ]);
  });
  return { saved: data.length };
}

// ══════════════════════════════════════════════════════════════
//  INVOICE
// ══════════════════════════════════════════════════════════════
function getInvoice(p) {
  let rows = sheetToObjects(getSheet('Invoice'));
  if (p.murid_id) rows = rows.filter(r => r.murid_id === p.murid_id);
  if (p.bulan)    rows = rows.filter(r => r.bulan    === p.bulan);
  return rows.sort((a,b) => b.bulan > a.bulan ? 1 : -1);
}

function generateInvoice(murid_id, bulan, catatan) {
  // Get murid fee_per_lesson
  const murid = getMurid({ id: murid_id });
  if (!murid) throw new Error('Murid tidak ditemukan');
  const fee = Number(murid.fee_per_lesson || murid.fee) || 0;

  // Count sesi in this month
  const laporan = getLaporan({ murid_id, bulan });
  const sesi    = laporan.length;
  const nominal = fee * sesi;

  // Check if invoice already exists
  const existing = getInvoice({ murid_id, bulan });
  if (existing.length > 0) throw new Error('Invoice untuk ' + bulan + ' sudah ada');

  const sheet = getSheet('Invoice');
  const id    = uid();
  sheet.appendRow([ id, murid_id, bulan, sesi, nominal, 'belum_bayar', '', '', catatan || '' ]);
  return { id, sesi, nominal };
}

function updateInvoiceStatus(id, status, tgl_kirim) {
  const sheet   = getSheet('Invoice');
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx   = headers.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === id) {
      sheet.getRange(i+1, headers.indexOf('status')+1).setValue(status);
      if (tgl_kirim) sheet.getRange(i+1, headers.indexOf('tgl_kirim')+1).setValue(tgl_kirim);
      return { updated: id, status };
    }
  }
  throw new Error('Invoice tidak ditemukan');
}

function deleteInvoice(id) {
  const sheet = getSheet('Invoice');
  const data  = sheet.getDataRange().getValues();
  const idIdx = data[0].indexOf('id');
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][idIdx] === id) { sheet.deleteRow(i+1); return { deleted: id }; }
  }
  throw new Error('Invoice tidak ditemukan');
}

function uploadBukti(p) {
  // Save base64 image to Drive, update invoice
  let folder;
  try {
    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
  } catch(e) { folder = DriveApp.getRootFolder(); }

  const base64 = p.base64.split(',')[1];
  const mime   = p.base64.split(';')[0].split(':')[1];
  const blob   = Utilities.newBlob(Utilities.base64Decode(base64), mime, p.filename || 'bukti.jpg');
  const file   = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/uc?id=' + file.getId();

  // Update invoice: status → menunggu, bukti_url → url
  const sheet   = getSheet('Invoice');
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx   = headers.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === p.invoice_id) {
      sheet.getRange(i+1, headers.indexOf('status')+1).setValue('menunggu');
      sheet.getRange(i+1, headers.indexOf('bukti_url')+1).setValue(url);
      return { url };
    }
  }
  throw new Error('Invoice tidak ditemukan');
}

