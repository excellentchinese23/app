// ============================================================
//  Excellent Mandarin Course – Client API Helper
// ============================================================

const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwiTlmkdRnR5zNrODbBYXbv9qtqGTRPr3TEs06LPVhRx6OW4MsmKb509sjNtjZdRDU1Kw/exec';

const API = {
  async post(payload) {
    const res  = await fetch(SCRIPT_URL, { method:'POST', headers:{'Content-Type':'text/plain'}, body:JSON.stringify(payload) });
    const json = await res.json();
    if (json.status === 'error') throw new Error(json.message);
    return json.data;
  },
  async get(params) {
    const qs  = new URLSearchParams(params).toString();
    const res = await fetch(SCRIPT_URL + '?' + qs);
    const json = await res.json();
    if (json.status === 'error') throw new Error(json.message);
    return json.data;
  },

  // MURID  (tipe_program | mode | program | fee_per_lesson)
  getMurid:      (opts={}) => API.post({ action:'getMurid', ...opts }),
  getMuridByLink:(link_id) => API.post({ action:'getMuridByLink', link_id }),
  addMurid:      (d)       => API.post({ action:'addMurid', ...d }),
  updateMurid:   (d)       => API.post({ action:'updateMurid', ...d }),
  deleteMurid:   (id)      => API.post({ action:'deleteMurid', id }),

  // GURU  (id_karyawan | jabatan | honor_onsite | honor_online | wa | tgl_bergabung | status_ptkp)
  getGuru:    ()   => API.post({ action:'getGuru' }),
  addGuru:    (d)  => API.post({ action:'addGuru', ...d }),
  updateGuru: (d)  => API.post({ action:'updateGuru', ...d }),
  deleteGuru: (id) => API.post({ action:'deleteGuru', id }),

  // LAPORAN  (murid_id | guru_id | tanggal | subject | analisis | catatan)
  getLaporan:    (opts={}) => API.post({ action:'getLaporan', ...opts }),
  addLaporan:    (d)       => API.post({ action:'addLaporan', ...d }),
  deleteLaporan: (id)      => API.post({ action:'deleteLaporan', id }),

  // ABSENSI
  getAbsensi:  (opts={})       => API.post({ action:'getAbsensi', ...opts }),
  saveAbsensi: (tanggal, data) => API.post({ action:'saveAbsensi', tanggal, data }),

  // INVOICE
  getInvoice:          (opts={})                     => API.post({ action:'getInvoice', ...opts }),
  generateInvoice:     (murid_id, bulan, catatan='') => API.post({ action:'generateInvoice', murid_id, bulan, catatan }),
  updateInvoiceStatus: (id, status, tgl_kirim)       => API.post({ action:'updateInvoiceStatus', id, status, tgl_kirim }),
  deleteInvoice:       (id)                          => API.post({ action:'deleteInvoice', id }),
  uploadBukti: (invoice_id, murid_id, bulan, file) =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try { resolve(await API.post({ action:'uploadBukti', invoice_id, murid_id, bulan, base64:e.target.result, filename:file.name })); }
        catch(err) { reject(err); }
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    }),

  // SETUP — jalankan sekali untuk buat semua sheet
  setup: () => API.post({ action:'setup' }),
};
