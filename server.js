const express = require('express');
const bodyParser = require('body-parser');
const path = require('path');
const ExcelJS = require('exceljs');
const { createClient } = require('@supabase/supabase-js');

const app = express();
const PORT = process.env.PORT || 3000;

// Inisialisasi Supabase Client menggunakan Environment Variables
const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_ANON_KEY;
const supabase = createClient(supabaseUrl, supabaseKey);

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Helper: Mengambil seluruh data pasien dari Supabase
async function getPasienData() {
  const { data, error } = await supabase
    .from('pasien')
    .select('*')
    .order('created_at', { ascending: true });

  if (error) {
    console.error('Error mengambil data dari Supabase:', error.message);
    return [];
  }
  return data || [];
}

function hitungKehamilan(hphtStr) {
  if (!hphtStr || hphtStr === '-') return { hpl: '-', usia_kehamilan: '-' };
  const hpht = new Date(hphtStr);
  const today = new Date();

  const hplDate = new Date(hpht);
  hplDate.setDate(hplDate.getDate() + 280);
  const hpl = hplDate.toISOString().split('T')[0];

  const diffTime = Math.abs(today - hpht);
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  const minggu = Math.floor(diffDays / 7);
  const hari = diffDays % 7;

  return {
    hpl,
    usia_kehamilan: `${minggu} Minggu ${hari} Hari`
  };
}

// 0. Redirect Root & Halaman Login
app.get('/', (req, res) => {
  res.redirect('/login');
});

app.get('/login', (req, res) => {
  res.render('login', { error: null });
});

app.post('/login', (req, res) => {
  const { username, password } = req.body;

  if (username === 'loket' && password === '123') {
    return res.redirect('/pendaftaran');
  }

  if (username === 'bidan' && password === '123') {
    return res.redirect('/kebidanan');
  }

  res.render('login', { error: 'Username atau Password salah!' });
});

// 1. Akun Khusus Pendaftaran (Loket) - Prioritas Rawat Luka di atas
app.get('/pendaftaran', async (req, res) => {
  try {
    const masterPasien = await getPasienData();
    const search = req.query.search;
    const today = new Date().toLocaleDateString('en-CA');

    let totalPasienTerdaftar = masterPasien.length;
    let pasienBulanIniCount = 0;
    let pasienTahunIniCount = 0;

    const currentMonth = today.substring(0, 7);
    const currentYear = today.substring(0, 4);

    masterPasien.forEach(p => {
      if (p.kunjungan && Array.isArray(p.kunjungan)) {
        p.kunjungan.forEach(k => {
          if (k.tanggal && k.tanggal.startsWith(currentMonth)) pasienBulanIniCount++;
          if (k.tanggal && k.tanggal.startsWith(currentYear)) pasienTahunIniCount++;
        });
      }
    });

    let pasienTampil = masterPasien.filter(p => {
      const terdaftarHariIni = p.tanggal_daftar === today;
      const belumBatal = p.status_antrean !== 'Batal Layanan';
      return terdaftarHariIni && belumBatal;
    });

    if (search) {
      const q = search.toLowerCase();
      pasienTampil = pasienTampil.filter(p =>
        (p.nama && p.nama.toLowerCase().includes(q)) ||
        (p.nik && p.nik.toLowerCase().includes(q)) ||
        (p.id && p.id.toLowerCase().includes(q)) ||
        (p.suami && p.suami.toLowerCase().includes(q)) ||
        (p.alamat && p.alamat.toLowerCase().includes(q))
      );
    }

    // Urutkan: Rawat Luka selalu di posisi paling atas
    pasienTampil.sort((a, b) => {
      const isLukaA = (a.jenis_kunjungan || '').toLowerCase().includes('rawat luka');
      const isLukaB = (b.jenis_kunjungan || '').toLowerCase().includes('rawat luka');
      if (isLukaA && !isLukaB) return -1;
      if (!isLukaA && isLukaB) return 1;
      return 0;
    });

    res.render('index', {
      pasienList: pasienTampil,
      masterPasienList: masterPasien,
      search: search || '',
      statistik: {
        totalPasienTerdaftar,
        pasienHariIni: pasienTampil.length,
        pasienBulanIni: pasienBulanIniCount,
        pasienTahunIni: pasienTahunIniCount
      }
    });
  } catch (err) {
    console.error('Error /pendaftaran:', err);
    res.status(500).send('Terjadi kesalahan saat memuat data pendaftaran.');
  }
});

// 2. Akun Khusus Poli Kebidanan / Dokter (SOAP) - Prioritas Rawat Luka di atas
app.get('/kebidanan', async (req, res) => {
  try {
    const masterPasien = await getPasienData();
    const today = new Date().toLocaleDateString('en-CA');

    let pasienAntrean = masterPasien.filter(p => {
      const terdaftarHariIni = p.tanggal_daftar === today;
      const belumSelesai =
        (!p.status_antrean || p.status_antrean === 'Menunggu' || p.status_antrean === 'Sedang Dilayani') &&
        p.status_antrean !== 'Batal Layanan';
      return terdaftarHariIni && belumSelesai;
    });

    // Urutkan: Rawat Luka selalu di posisi paling atas
    pasienAntrean.sort((a, b) => {
      const isLukaA = (a.jenis_kunjungan || '').toLowerCase().includes('rawat luka');
      const isLukaB = (b.jenis_kunjungan || '').toLowerCase().includes('rawat luka');
      if (isLukaA && !isLukaB) return -1;
      if (!isLukaA && isLukaB) return 1;
      return 0;
    });

    res.render('kebidanan_view', { pasienList: pasienAntrean });
  } catch (err) {
    console.error('Error /kebidanan:', err);
    res.status(500).send('Terjadi kesalahan saat memuat antrean kebidanan.');
  }
});

// 3. Executive Analytics & Traffic Dashboard
app.get('/rekap-trafik', async (req, res) => {
  try {
    const pasienList = await getPasienData();

    let totalKunjungan = 0;
    let totalPendapatan = 0;
    let countObstetri = 0;
    let countGinekologi = 0;
    let countSelesai = 0;
    let countBatal = 0;

    let monthlyData = {};
    let alamatData = {};
    let diagnosaData = {};

    pasienList.forEach(p => {
      const wil = p.alamat ? p.alamat.trim() : 'Lainnya';
      alamatData[wil] = (alamatData[wil] || 0) + 1;

      if (p.status_antrean === 'Batal' || p.status_antrean === 'Batal Layanan') {
        countBatal++;
      } else if (p.status_antrean === 'Selesai' || (p.kunjungan && p.kunjungan.length > 0)) {
        countSelesai++;
      }

      if (p.kunjungan && Array.isArray(p.kunjungan)) {
        p.kunjungan.forEach(k => {
          totalKunjungan++;
          if (k.kasir && k.kasir.total) totalPendapatan += k.kasir.total;

          if (k.jenis_layanan === 'Ginekologi') countGinekologi++;
          else countObstetri++;

          const bulan = k.tanggal ? k.tanggal.substring(0, 7) : 'Lainnya';
          monthlyData[bulan] = (monthlyData[bulan] || 0) + 1;

          if (k.diagnosa) {
            const diag = k.diagnosa.trim();
            diagnosaData[diag] = (diagnosaData[diag] || 0) + 1;
          }
        });
      }
    });

    res.render('rekap_trafik', {
      totalPasien: pasienList.length,
      totalKunjungan,
      totalPendapatan,
      countObstetri,
      countGinekologi,
      countSelesai,
      countBatal,
      monthlyData,
      alamatData,
      diagnosaData
    });
  } catch (err) {
    console.error('Error /rekap-trafik:', err);
    res.status(500).send('Terjadi kesalahan saat memuat rekap trafik.');
  }
});

// 4. Monitor Antrean TV Ruang Tunggu
app.get('/antrean', async (req, res) => {
  try {
    const masterPasien = await getPasienData();
    const today = new Date().toLocaleDateString('en-CA');

    const pasienHariIni = masterPasien.filter(p => {
      return p.tanggal_daftar === today && p.status_antrean !== 'Batal Layanan';
    });
    res.render('antrean', { pasienList: pasienHariIni });
  } catch (err) {
    console.error('Error /antrean:', err);
    res.status(500).send('Terjadi kesalahan saat memuat antrean.');
  }
});

// 5. Tambah / Edit Pasien (Registrasi Loket) -> Menggunakan Supabase
app.post('/tambah-pasien', async (req, res) => {
  try {
    const editId = req.body.edit_id;
    const today = new Date().toLocaleDateString('en-CA');

    if (editId && editId.trim() !== '') {
      // Ambil data lama terlebih dahulu untuk fallback field kosong
      const { data: existingPasien } = await supabase
        .from('pasien')
        .select('*')
        .eq('id', editId.trim())
        .single();

      if (existingPasien) {
        const { error } = await supabase
          .from('pasien')
          .update({
            nik: req.body.nik || existingPasien.nik,
            nama: req.body.nama || existingPasien.nama,
            usia: req.body.usia || existingPasien.usia,
            suami: req.body.suami || existingPasien.suami,
            telepon: req.body.telepon || existingPasien.telepon,
            alamat: req.body.alamat || existingPasien.alamat,
            jenis_kunjungan: req.body.jenis_kunjungan || existingPasien.jenis_kunjungan,
            tanggal_daftar: today,
            status_antrean: 'Menunggu'
          })
          .eq('id', editId.trim());

        if (error) throw error;
      }
    } else {
      // Hitung jumlah pasien saat ini untuk format ID POG-XXX
      const { count } = await supabase
        .from('pasien')
        .select('*', { count: 'exact', head: true });

      const nextNum = (count || 0) + 1;
      const newId = 'POG-' + String(nextNum).padStart(3, '0');

      const newPasien = {
        id: newId,
        nik: req.body.nik || '-',
        nama: req.body.nama,
        usia: req.body.usia,
        suami: req.body.suami,
        telepon: req.body.telepon,
        alamat: req.body.alamat || 'Jember',
        jenis_kunjungan: req.body.jenis_kunjungan || 'USG',
        tanggal_daftar: today,
        status_antrean: 'Menunggu',
        kunjungan: []
      };

      const { error } = await supabase.from('pasien').insert([newPasien]);
      if (error) throw error;
    }

    res.redirect('/pendaftaran');
  } catch (err) {
    console.error('Error /tambah-pasien:', err.message);
    res.status(500).send('Gagal menyimpan data pasien ke Supabase: ' + err.message);
  }
});

// 6. Batal Layanan Pasien dengan Ketik Alasan Wajib
app.post('/pasien/batal/:id', async (req, res) => {
  try {
    const alasanBatal = req.body.alasan || 'Tidak ada alasan';

    const { data, error } = await supabase
      .from('pasien')
      .update({
        status_antrean: 'Batal Layanan',
        alasan_batal: alasanBatal
      })
      .eq('id', req.params.id)
      .select();

    if (error) throw error;

    if (data && data.length > 0) {
      return res.json({ success: true });
    }
    res.status(404).json({ success: false });
  } catch (err) {
    console.error('Error /pasien/batal:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// 7. Detail Rekam Medis (SOAP)
app.get('/pasien/:id', async (req, res) => {
  try {
    const { data: pasien, error } = await supabase
      .from('pasien')
      .select('*')
      .eq('id', req.params.id)
      .single();

    if (error || !pasien) return res.status(404).send('Pasien tidak ditemukan');
    res.render('detail', { pasien });
  } catch (err) {
    console.error('Error GET /pasien/:id:', err.message);
    res.status(500).send('Terjadi kesalahan saat memuat detail pasien.');
  }
});

// 8. Input Kunjungan SOAP & Lab Plano Tes
app.post('/pasien/:id/kunjungan', async (req, res) => {
  try {
    const { data: pasien, error: fetchError } = await supabase
      .from('pasien')
      .select('*')
      .eq('id', req.params.id)
      .single();

    if (fetchError || !pasien) return res.status(404).send('Pasien tidak ditemukan');

    const jenisLayanan = req.body.jenis_layanan || 'USG';
    const hpht = req.body.hpht;
    const hasilHitung = hitungKehamilan(hpht);

    let orderLab = req.body.order_lab || [];
    if (!Array.isArray(orderLab)) orderLab = [orderLab];

    const biayaKonsultasi = parseInt(req.body.biaya_konsultasi) || 0;
    const biayaUSG = parseInt(req.body.biaya_usg) || 0;
    const biayaObat = parseInt(req.body.biaya_obat) || 0;
    const biayaLab = parseInt(req.body.biaya_lab) || 0;
    const totalBiaya = biayaKonsultasi + biayaUSG + biayaObat + biayaLab;

    const newKunjungan = {
      id_kunjungan: 'K-' + Date.now(),
      tanggal: new Date().toLocaleDateString('en-CA'),
      jenis_layanan: jenisLayanan,
      hpht: hpht || '-',
      hpl: hasilHitung.hpl,
      usia_kehamilan: hasilHitung.usia_kehamilan,
      bb: req.body.bb || '-',
      td: req.body.td || '-',
      tfu: req.body.tfu || '-',
      djj: req.body.djj || '-',
      usg_efw: req.body.usg_efw || '-',
      order_lab: orderLab,
      diagnosa: req.body.diagnosa,
      terapi: req.body.terapi,
      kasir: {
        biaya_konsultasi: biayaKonsultasi,
        biaya_usg: biayaUSG,
        biaya_obat: biayaObat,
        biaya_lab: biayaLab,
        total: totalBiaya
      }
    };

    const updatedKunjungan = Array.isArray(pasien.kunjungan) ? [...pasien.kunjungan] : [];
    updatedKunjungan.unshift(newKunjungan);

    const { error: updateError } = await supabase
      .from('pasien')
      .update({
        status_antrean: 'Selesai',
        kunjungan: updatedKunjungan
      })
      .eq('id', req.params.id);

    if (updateError) throw updateError;

    res.redirect(`/pasien/${req.params.id}`);
  } catch (err) {
    console.error('Error POST /pasien/:id/kunjungan:', err.message);
    res.status(500).send('Gagal menyimpan data kunjungan: ' + err.message);
  }
});

// 9. Cetak Surat & Kwitansi
app.get('/pasien/:id/cetak-surat/:idKunjungan', async (req, res) => {
  try {
    const { data: pasien, error } = await supabase
      .from('pasien')
      .select('*')
      .eq('id', req.params.id)
      .single();

    if (error || !pasien) return res.status(404).send('Pasien tidak ditemukan');
    const kunjungan = (pasien.kunjungan || []).find(k => k.id_kunjungan === req.params.idKunjungan);
    res.render('cetak_surat', { pasien, kunjungan });
  } catch (err) {
    console.error('Error cetak-surat:', err.message);
    res.status(500).send('Terjadi kesalahan saat mencetak surat.');
  }
});

app.get('/pasien/:id/cetak-kwitansi/:idKunjungan', async (req, res) => {
  try {
    const { data: pasien, error } = await supabase
      .from('pasien')
      .select('*')
      .eq('id', req.params.id)
      .single();

    if (error || !pasien) return res.status(404).send('Pasien tidak ditemukan');
    const kunjungan = (pasien.kunjungan || []).find(k => k.id_kunjungan === req.params.idKunjungan);
    res.render('cetak_kwitansi', { pasien, kunjungan });
  } catch (err) {
    console.error('Error cetak-kwitansi:', err.message);
    res.status(500).send('Terjadi kesalahan saat mencetak kwitansi.');
  }
});

// 10. Ekspor Excel (Termasuk Kolom Alasan Batal)
app.get('/export-excel', async (req, res) => {
  try {
    const pasienList = await getPasienData();
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Rekap Pasien & Kunjungan');

    sheet.columns = [
      { header: 'ID Pasien', key: 'id_pasien', width: 12 },
      { header: 'NIK Pasien', key: 'nik', width: 18 },
      { header: 'Nama Pasien', key: 'nama', width: 22 },
      { header: 'Nama Suami', key: 'suami', width: 20 },
      { header: 'No Telepon', key: 'telepon', width: 16 },
      { header: 'Alamat', key: 'alamat', width: 20 },
      { header: 'Status Antrean', key: 'status_antrean', width: 16 },
      { header: 'Alasan Batal', key: 'alasan_batal', width: 25 },
      { header: 'Jenis Layanan', key: 'jenis_layanan', width: 15 },
      { header: 'Tgl Kunjungan', key: 'tgl_kunjungan', width: 15 },
      { header: 'HPHT', key: 'hpht', width: 13 },
      { header: 'HPL', key: 'hpl', width: 13 },
      { header: 'Usia Kehamilan', key: 'uk', width: 18 },
      { header: 'Order Lab', key: 'order_lab', width: 20 },
      { header: 'Diagnosa', key: 'diagnosa', width: 30 },
      { header: 'Total Biaya (Rp)', key: 'total_biaya', width: 18 }
    ];

    pasienList.forEach(p => {
      if (!p.kunjungan || p.kunjungan.length === 0) {
        sheet.addRow({
          id_pasien: p.id,
          nik: p.nik || '-',
          nama: p.nama,
          suami: p.suami,
          telepon: p.telepon,
          alamat: p.alamat || '-',
          status_antrean: p.status_antrean || 'Menunggu',
          alasan_batal: p.alasan_batal || '-',
          jenis_layanan: '-',
          tgl_kunjungan: '-',
          hpht: '-',
          hpl: '-',
          uk: '-',
          order_lab: '-',
          diagnosa: '-',
          total_biaya: 0
        });
      } else {
        p.kunjungan.forEach(k => {
          sheet.addRow({
            id_pasien: p.id,
            nik: p.nik || '-',
            nama: p.nama,
            suami: p.suami,
            telepon: p.telepon,
            alamat: p.alamat || '-',
            status_antrean: p.status_antrean || 'Selesai',
            alasan_batal: p.alasan_batal || '-',
            jenis_layanan: k.jenis_layanan || 'USG',
            tgl_kunjungan: k.tanggal,
            hpht: k.hpht || '-',
            hpl: k.hpl || '-',
            uk: k.usia_kehamilan || '-',
            order_lab: k.order_lab ? k.order_lab.join(', ') : '-',
            diagnosa: k.diagnosa || '-',
            total_biaya: k.kasir ? k.kasir.total : 0
          });
        });
      }
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=' + `Rekap_Praktek_Dr_ShodiqKukuh_${Date.now()}.xlsx`);

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('Error /export-excel:', err.message);
    res.status(500).send('Gagal mengekspor data Excel.');
  }
});

// Ekspor app untuk Vercel Serverless & jalankan listen untuk lokal
if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`Server SIM-OBGYN berjalan di http://localhost:${PORT}`);
  });
}

module.exports = app;
