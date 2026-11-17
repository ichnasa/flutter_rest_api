const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, LevelFormat, PageBreak, ExternalHyperlink,
  ShadingType,
} = require('docx');

const DOCS = process.argv[2];
const OUT = process.argv[3];
const SHOTS = path.join(DOCS, 'screenshots');
const FONT = 'Times New Roman';
const CONTENT_W = 11906 - 1701 * 2; // A4 minus 3 cm left/right (DXA)

// ---------- helpers ----------
function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}
// Parse "text with `code`" into runs (code in Courier New)
function runs(text, opts = {}) {
  return text.split(/(`[^`]+`)/).filter(Boolean).map(part =>
    part.startsWith('`')
      ? new TextRun({ text: part.slice(1, -1), font: 'Courier New', size: 21, ...opts })
      : new TextRun({ text: part, ...opts }));
}
const p = (text, o = {}) => new Paragraph({
  alignment: o.align ?? AlignmentType.JUSTIFIED,
  indent: o.noIndent ? undefined : { firstLine: 709 },
  spacing: { after: 120, line: 360 },
  children: runs(text, o.run),
  keepNext: o.keepNext,
});
const center = (text, o = {}) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: o.after ?? 0, line: 360, before: o.before ?? 0 },
  children: runs(text, { bold: o.bold, size: o.size }),
});
const h1 = (a, b) => [
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ style: 'Heading1', children: [new TextRun(a)] }),
  new Paragraph({ style: 'Heading1', spacing: { before: 0, after: 360 }, children: [new TextRun(b)] }),
];
const h2 = t => new Paragraph({ style: 'Heading2', children: [new TextRun(t)] });
const num = (text, ref) => new Paragraph({
  numbering: { reference: ref, level: 0 },
  alignment: AlignmentType.JUSTIFIED,
  spacing: { after: 60, line: 360 },
  children: runs(text),
});
const caption = t => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { before: 60, after: 240 },
  children: [new TextRun({ text: t, size: 22 })],
});
const tableCaption = t => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 }, keepNext: true,
  children: [new TextRun({ text: t, size: 22 })],
});

function image(file, maxWpx, maxHpx) {
  const { w, h } = pngSize(file);
  const s = Math.min(maxWpx / w, maxHpx / h, 1);
  return new ImageRun({
    type: 'png', data: fs.readFileSync(file),
    transformation: { width: Math.round(w * s), height: Math.round(h * s) },
  });
}
function figure(file, cap) {
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 120 },
      children: [image(file, 560, 400)] }),
    caption(cap),
  ];
}

const border = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const borders = { top: border, bottom: border, left: border, right: border };
function table(widths, rows) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((r, i) => new TableRow({
      tableHeader: i === 0,
      children: r.map((c, j) => new TableCell({
        borders,
        width: { size: widths[j], type: WidthType.DXA },
        shading: i === 0 ? { fill: 'E7E6E6', type: ShadingType.CLEAR, color: 'auto' } : undefined,
        margins: { top: 40, bottom: 40, left: 90, right: 90 },
        children: [new Paragraph({
          alignment: i === 0 ? AlignmentType.CENTER : AlignmentType.LEFT,
          children: runs(c, { size: 22, bold: i === 0 }),
        })],
      })),
    })),
  });
}

// pakai screenshot Network tab bila ada, selain itu screenshot test console
const apiShot = n => {
  const net = path.join(SHOTS, `network-${n}.png`);
  return fs.existsSync(net) ? net : path.join(SHOTS, `api-${n}.png`);
};

// ---------- content ----------
const cover = [
  center('LAPORAN PRAKTIKUM', { bold: true, size: 28 }),
  center('TOPIC 10: REST API DAN JSON PARSING', { bold: true, size: 28 }),
  center('Implementasi REST API CRUD UMKM dengan CodeIgniter 4', { bold: true, size: 28 }),
  center('dan JSON Parsing pada Flutter', { bold: true, size: 28, after: 1800 }),
  center('Dosen Pengampu: Rafi’e, M.Kom', { after: 1800 }),
  center('Disusun oleh:', { after: 120 }),
  center('Nama : ..........................................'),
  center('NIM   : ..........................................', { after: 2600 }),
  center('PROGRAM STUDI ILMU KOMPUTER', { bold: true }),
  center('FAKULTAS MATEMATIKA DAN ILMU PENGETAHUAN ALAM', { bold: true }),
  center('UNIVERSITAS LAMBUNG MANGKURAT', { bold: true }),
  center('BANJARBARU', { bold: true }),
  center('2026', { bold: true }),
];

const bab1 = [
  ...h1('BAB I', 'PENDAHULUAN'),
  h2('1.1 Latar Belakang'),
  p('REST API adalah antarmuka berbasis protokol HTTP yang memungkinkan aplikasi klien bertukar data dengan server. Data umumnya dikirim dalam format JSON karena ringan dan mudah diproses. Praktikum ini menerapkan materi tersebut dengan membangun REST API data UMKM menggunakan CodeIgniter 4, lalu mengonsumsinya dari aplikasi Flutter.'),
  h2('1.2 Tujuan'),
  num('Membuat REST API CRUD (GET, POST, PUT, DELETE) dengan CodeIgniter 4.', 'tujuan'),
  num('Melakukan JSON parsing pada Flutter menggunakan model class.', 'tujuan'),
  num('Menguji API dan menganalisis hasilnya.', 'tujuan'),
  h2('1.3 Alat yang Digunakan'),
  p('PHP 8.4, CodeIgniter 4.7.6, SQLite/MySQL, Flutter 3 dengan package `http`, Google Chrome (DevTools), dan Visual Studio Code.', { noIndent: true }),
];

const tests = [
  ['Ambil daftar UMKM', '`GET /api/umkm?page=1&per_page=3`', '200 OK', 'GET daftar UMKM'],
  ['Ambil detail UMKM', '`GET /api/umkm/1`', '200 OK', 'GET detail UMKM'],
  ['Tambah data valid', '`POST /api/umkm`', '201 Created', 'POST tambah UMKM'],
  ['Tambah data tidak valid', '`POST /api/umkm`', '400 Bad Request', 'POST dengan data tidak valid'],
  ['Perbarui data', '`PUT /api/umkm/6`', '200 OK', 'PUT perbarui UMKM'],
  ['Hapus data', '`DELETE /api/umkm/6`', '200 OK', 'DELETE hapus UMKM'],
  ['Ambil data yang sudah dihapus', '`GET /api/umkm/6`', '404 Not Found', 'GET data yang tidak ditemukan'],
  ['Pencarian dan filter kategori', '`GET /api/umkm?kategori=Kuliner&search=Banjar`', '200 OK', 'GET dengan pencarian dan filter kategori'],
];

const bab2 = [
  ...h1('BAB II', 'HASIL DAN PEMBAHASAN'),
  h2('2.1 Rancangan API'),
  p('API dibuat dengan `ResourceController` pada `app/Controllers/Api/UmkmController.php`. Validasi data ditulis pada `app/Models/UmkmModel.php`. Setiap response memakai format `{success, message, data, meta, errors}` sesuai materi. Daftar endpoint ditunjukkan pada Tabel 2.1.'),
  tableCaption('Tabel 2.1 Daftar endpoint REST API UMKM'),
  table([1200, 2600, CONTENT_W - 3800], [
    ['Method', 'Endpoint', 'Fungsi'],
    ['GET', '`/api/umkm`', 'Mengambil daftar UMKM (pagination, pencarian, filter)'],
    ['GET', '`/api/umkm/{id}`', 'Mengambil detail UMKM'],
    ['POST', '`/api/umkm`', 'Menambah data UMKM'],
    ['PUT', '`/api/umkm/{id}`', 'Memperbarui data UMKM'],
    ['DELETE', '`/api/umkm/{id}`', 'Menghapus data UMKM'],
  ]),
  h2('2.2 Hasil Pengujian API'),
  p('Pengujian dilakukan pada server lokal `http://localhost:8080`. Request dan response diamati melalui tab Network pada Chrome DevTools. Ringkasan hasil ditunjukkan pada Tabel 2.2.'),
  tableCaption('Tabel 2.2 Hasil pengujian API'),
  table([600, 2400, 3804, 1700], [
    ['No', 'Pengujian', 'Request', 'Status'],
    ...tests.map((t, i) => [String(i + 1), t[0], t[1], t[2]]),
  ]),
  new Paragraph({ spacing: { after: 120 } }),
  p('Seluruh pengujian memberikan status sesuai harapan. Pada data tidak valid (Gambar 2.4), API mengembalikan pesan kesalahan untuk setiap field. Pada data yang tidak ditemukan (Gambar 2.7), API mengembalikan status 404.'),
  ...tests.flatMap((t, i) => figure(apiShot(i + 1), `Gambar 2.${i + 1} ${t[3]}`)),
  h2('2.3 JSON Parsing pada Flutter'),
  p('Response JSON diubah menjadi objek Dart melalui model class `Umkm` dengan `factory Umkm.fromJson()` dan method `toJson()`. Request dikirim oleh `UmkmService` menggunakan `async/await` dan `try-catch`. Tampilan aplikasi ditunjukkan pada Gambar 2.9.'),
  new Paragraph({
    alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 120 },
    children: ['app-1-list.png', 'app-5-form.png', 'app-3-detail.png'].flatMap((f, i) => [
      ...(i ? [new TextRun('  ')] : []),
      image(path.join(SHOTS, f), 180, 380),
    ]),
  }),
  caption('Gambar 2.9 Halaman daftar, form tambah, dan detail UMKM'),
];

const bab3 = [
  ...h1('BAB III', 'REFLEKSI DAN PENUTUP'),
  h2('3.1 Refleksi'),
  p('Praktikum ini menunjukkan bahwa REST API berperan sebagai kontrak antara server dan klien. Format response dan status code yang konsisten memudahkan klien membedakan keberhasilan dan kegagalan tanpa membaca isi pesan. CodeIgniter 4 mempermudah pembuatan API karena `resource route` langsung memetakan method HTTP ke fungsi controller, dan validasi cukup ditulis sekali pada model.'),
  p('Pada sisi Flutter, parsing dengan model class lebih aman daripada mengakses `Map` secara langsung karena tipe data diperiksa saat kompilasi. Package `http` sudah cukup untuk kebutuhan CRUD sederhana, sedangkan `Dio` lebih cocok bila dibutuhkan interceptor atau upload file.'),
  p('Kendala yang dihadapi adalah sebagai berikut:'),
  num('Request dari Flutter Web diblokir CORS. Masalah ini diatasi dengan mengaktifkan filter `cors` dan menambahkan route `OPTIONS`.', 'kendala'),
  num('Emulator Android tidak dapat mengakses `localhost`, sehingga base URL diganti menjadi `10.0.2.2`.', 'kendala'),
  num('Nilai `id` dari database terkadang bertipe string, sehingga dikonversi ke integer pada backend dan saat parsing di Flutter.', 'kendala'),
  h2('3.2 Kesimpulan'),
  p('REST API CRUD UMKM berhasil dibuat dengan CodeIgniter 4. Seluruh endpoint telah diuji dan memberikan status serta response sesuai rancangan. Aplikasi Flutter berhasil menampilkan, menambah, mengubah, dan menghapus data melalui API dengan JSON parsing berbasis model class.'),
  h2('3.3 Link Source Code'),
  new Paragraph({
    spacing: { after: 120, line: 360 },
    children: [
      new TextRun('Repository GitHub: '),
      new ExternalHyperlink({
        link: 'https://github.com/ichnasa/flutter_rest_api',
        children: [new TextRun({ text: 'https://github.com/ichnasa/flutter_rest_api', style: 'Hyperlink' })],
      }),
    ],
  }),
  new Paragraph({ spacing: { after: 120, line: 360 }, alignment: AlignmentType.JUSTIFIED,
    children: runs('Folder `backend/` berisi REST API CodeIgniter 4, dan folder `frontend/` berisi aplikasi Flutter.') }),
];

const numCfg = ref => ({
  reference: ref,
  levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 709, hanging: 360 } } } }],
});

const doc = new Document({
  creator: 'Mahasiswa',
  title: 'Laporan Praktikum Topic 10 - REST API dan JSON Parsing',
  styles: {
    default: { document: { run: { font: FONT, size: 24 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 24, bold: true },
        paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0, line: 360 }, keepNext: true, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 24, bold: true },
        paragraph: { spacing: { before: 240, after: 120, line: 360 }, keepNext: true, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [numCfg('tujuan'), numCfg('kendala')] },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1701, bottom: 1417, left: 1701, right: 1701 },
      },
    },
    children: [...cover, ...bab1, ...bab2, ...bab3],
  }],
});

Packer.toBuffer(doc).then(buf => { fs.writeFileSync(OUT, buf); console.log('written', OUT); });
