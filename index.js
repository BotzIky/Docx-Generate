const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, PageBreak, LevelFormat, ExternalHyperlink,
  TabStopType, TabStopPosition, Header, Footer, UnderlineType
} = require('docx');
const fs = require('fs');

// Helper functions
const p = (text, opts = {}) => new Paragraph({
  alignment: opts.align || AlignmentType.JUSTIFIED,
  spacing: { line: 480, before: opts.before || 0, after: opts.after || 0 },
  indent: opts.indent ? { firstLine: 720 } : {},
  children: [new TextRun({
    text: text || '',
    bold: opts.bold || false,
    italics: opts.italics || false,
    size: 24,
    font: 'Times New Roman',
    underline: opts.underline ? { type: UnderlineType.SINGLE } : undefined
  })]
});

const pCenter = (text, opts = {}) => p(text, { ...opts, align: AlignmentType.CENTER });

const pBold = (text, opts = {}) => p(text, { ...opts, bold: true });

const pBoldCenter = (text, opts = {}) => p(text, { ...opts, bold: true, align: AlignmentType.CENTER });

const heading1 = (text) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { line: 480, before: 240, after: 240 },
  children: [new TextRun({ text, bold: true, size: 24, font: 'Times New Roman', allCaps: true })]
});

const heading2 = (text) => new Paragraph({
  alignment: AlignmentType.LEFT,
  spacing: { line: 480, before: 240, after: 120 },
  children: [new TextRun({ text, bold: true, size: 24, font: 'Times New Roman' })]
});

const heading3 = (text) => new Paragraph({
  alignment: AlignmentType.LEFT,
  spacing: { line: 480, before: 180, after: 80 },
  children: [new TextRun({ text, bold: true, size: 24, font: 'Times New Roman' })]
});

const heading4 = (text) => new Paragraph({
  alignment: AlignmentType.LEFT,
  spacing: { line: 480, before: 120, after: 60 },
  children: [new TextRun({ text, bold: true, italics: true, size: 24, font: 'Times New Roman' })]
});

const emptyLine = () => new Paragraph({
  spacing: { line: 480 },
  children: [new TextRun({ text: '', size: 24, font: 'Times New Roman' })]
});

const pageBreak = () => new Paragraph({
  children: [new TextRun({ break: 1 })]
});

const pMixed = (runs) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { line: 480, before: 0, after: 0 },
  indent: { firstLine: 720 },
  children: runs.map(r => new TextRun({
    text: r.text,
    bold: r.bold || false,
    italics: r.italics || false,
    size: 24,
    font: 'Times New Roman'
  }))
});

const numbered = (text, num, opts = {}) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { line: 480, before: 0, after: 0 },
  indent: { left: 720, hanging: 360 },
  children: [new TextRun({ text: `${num}. ${text}`, size: 24, font: 'Times New Roman', bold: opts.bold || false })]
});

const bulleted = (text) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { line: 480, before: 0, after: 0 },
  indent: { left: 720, hanging: 360 },
  children: [new TextRun({ text: `• ${text}`, size: 24, font: 'Times New Roman' })]
});

const tableCell = (text, opts = {}) => new TableCell({
  borders: {
    top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
    bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
    left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
    right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
  },
  width: { size: opts.width || 2340, type: WidthType.DXA },
  shading: opts.header ? { fill: 'D9D9D9', type: ShadingType.CLEAR } : undefined,
  margins: { top: 80, bottom: 80, left: 120, right: 120 },
  verticalAlign: VerticalAlign.CENTER,
  children: [new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    children: [new TextRun({ text, bold: opts.bold || opts.header || false, size: 20, font: 'Times New Roman' })]
  })]
});

// ================= DOCUMENT CONTENT =================

const children = [

  // ==================== COVER PAGE ====================
  emptyLine(), emptyLine(), emptyLine(),
  pBoldCenter('SKRIPSI'),
  emptyLine(),
  pBoldCenter('PENGARUH MANAJEMEN WAKTU, KERJA PARUH WAKTU, DAN LINGKUNGAN KAMPUS TERHADAP PRESTASI AKADEMIK MAHASISWA PRODI MANAJEMEN FAKULTAS EKONOMI DAN BISNIS UNIVERSITAS MALIKUSSALEH', { align: AlignmentType.CENTER }),
  emptyLine(), emptyLine(),
  pCenter('Diajukan Sebagai Salah Satu Syarat untuk Memperoleh Gelar Sarjana Manajemen pada Program Studi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh'),
  emptyLine(), emptyLine(),
  pCenter('Disusun Oleh:'),
  pBoldCenter('[NAMA MAHASISWA]'),
  pCenter('NIM: [NIM MAHASISWA]'),
  emptyLine(), emptyLine(), emptyLine(),
  pBoldCenter('PROGRAM STUDI MANAJEMEN'),
  pBoldCenter('FAKULTAS EKONOMI DAN BISNIS'),
  pBoldCenter('UNIVERSITAS MALIKUSSALEH'),
  pBoldCenter('LHOKSEUMAWE'),
  pBoldCenter('[BULAN, TAHUN]'),

  pageBreak(),

  // ==================== BAB I ====================
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 480, before: 0, after: 120 },
    children: [new TextRun({ text: 'BAB I', bold: true, size: 24, font: 'Times New Roman' })]
  }),
  heading1('PENDAHULUAN'),
  emptyLine(),

  // 1.1 Latar Belakang
  heading2('1.1 Latar Belakang Penelitian'),
  p('Pendidikan tinggi merupakan fondasi utama dalam pembentukan sumber daya manusia yang kompeten, berdaya saing, dan adaptif terhadap perubahan global. Di era persaingan yang semakin ketat, prestasi akademik mahasiswa menjadi salah satu indikator keberhasilan proses pembelajaran dan cermin kualitas suatu perguruan tinggi. Prestasi akademik yang diwujudkan dalam Indeks Prestasi Kumulatif (IPK) tidak hanya mencerminkan kemampuan kognitif mahasiswa, tetapi juga mencerminkan kesiapan mereka dalam memasuki dunia kerja yang kompetitif.', { indent: true }),
  emptyLine(),
  p('Berdasarkan data Pangkalan Data Pendidikan Tinggi (PDDikti) Kementerian Pendidikan Tinggi, Sains, dan Teknologi Tahun 2024, jumlah mahasiswa aktif di Indonesia mencapai lebih dari 9 juta orang yang tersebar di berbagai perguruan tinggi negeri maupun swasta. Besarnya jumlah mahasiswa ini membawa konsekuensi tersendiri, terutama terkait dengan beragamnya kondisi dan latar belakang sosial-ekonomi mahasiswa. Salah satu fenomena yang semakin umum dijumpai adalah maraknya mahasiswa yang memilih untuk bekerja paruh waktu sambil menempuh pendidikan formal di perguruan tinggi.', { indent: true }),
  emptyLine(),
  p('Fenomena mahasiswa yang bekerja paruh waktu (part-time work) bukan merupakan hal baru di Indonesia. Menurut data Badan Pusat Statistik (BPS) tahun 2024, terdapat peningkatan signifikan jumlah penduduk muda usia 15-24 tahun yang aktif bekerja sambil bersekolah atau berkuliah sebagai respons terhadap tekanan ekonomi dan tingginya biaya hidup. Di Provinsi Aceh khususnya, kondisi ekonomi masyarakat yang masih dalam tahap pemulihan pasca-konflik dan bencana mendorong banyak mahasiswa untuk mencari penghasilan tambahan di luar jam kuliah guna membiayai kebutuhan pendidikan maupun kehidupan sehari-hari mereka.', { indent: true }),
  emptyLine(),
  p('Di Universitas Malikussaleh (UNIMAL), yang merupakan perguruan tinggi negeri terbesar di Aceh Utara, fenomena ini terlihat nyata di Program Studi Manajemen Fakultas Ekonomi dan Bisnis. Berdasarkan data SNBP 2025 yang dirilis oleh UNIMAL, Prodi Manajemen merupakan salah satu prodi terbanyak menerima calon mahasiswa, yaitu sebanyak 200 orang. Dengan jumlah mahasiswa yang cukup besar, berbagai latar belakang sosial-ekonomi mahasiswa menjadi sangat beragam. Sebagian besar mahasiswa berasal dari keluarga menengah ke bawah, sehingga tidak sedikit dari mereka yang memilih untuk bekerja paruh waktu demi memenuhi kebutuhan finansial selama menempuh pendidikan.', { indent: true }),
  emptyLine(),
  p('Dari hasil observasi awal dan wawancara tidak terstruktur yang dilakukan peneliti terhadap beberapa mahasiswa aktif Prodi Manajemen FEB UNIMAL, terungkap bahwa terdapat variasi prestasi akademik yang cukup signifikan antara mahasiswa yang bekerja paruh waktu dengan yang tidak bekerja. Beberapa mahasiswa yang bekerja paruh waktu mengaku kesulitan membagi waktu antara pekerjaan dan kewajiban akademik, sementara yang lain mampu mempertahankan IPK yang memuaskan berkat kemampuan manajemen waktu yang baik. Fenomena ini mengindikasikan bahwa terdapat faktor-faktor yang memediasi atau memoderasi hubungan antara kerja paruh waktu dan prestasi akademik, salah satunya adalah manajemen waktu.', { indent: true }),
  emptyLine(),
  p('Manajemen waktu merupakan salah satu variabel kunci yang banyak diteliti dalam kaitannya dengan prestasi akademik. Inayah et al. (2023) dalam penelitiannya terhadap 89 mahasiswa pekerja di Makassar menemukan adanya pengaruh positif dan signifikan manajemen waktu terhadap prestasi akademik mahasiswa yang bekerja (Nagelkerke R\u00b2 = 0,154; p = 0,005). Sementara itu, Kamilatunnisa et al. (2025) dalam penelitian terhadap mahasiswa Fakultas Ekonomi Unpas menemukan bahwa manajemen waktu menjelaskan 50,5% variasi prestasi akademik (R\u00b2 = 0,505; p = 0,000). Temuan ini menegaskan bahwa kemampuan merencanakan, mengorganisasikan, dan mengontrol penggunaan waktu secara efektif merupakan prediktor kuat keberhasilan akademik mahasiswa, terutama bagi mereka yang memiliki aktivitas ganda seperti bekerja sambil kuliah.', { indent: true }),
  emptyLine(),
  p('Namun, hubungan antara kerja paruh waktu dan prestasi akademik masih menunjukkan hasil yang kontradiktif dalam berbagai penelitian. Huda et al. (2023) menemukan bahwa kerja paruh waktu berpengaruh negatif signifikan terhadap prestasi akademik mahasiswa UIN Sumatera Utara. Berbeda dengan itu, Agustina dan Mardalis (2024) dalam penelitian terhadap 250 mahasiswa menemukan pengaruh positif dan signifikan kerja paruh waktu terhadap prestasi akademik (original sample 0,378; t-statistik 4,637; p = 0,000). Demikian pula Linggasari dan Kurniawan (2019) yang menemukan hubungan positif signifikan antara kerja paruh waktu dan prestasi akademik mahasiswa dengan koefisien korelasi r = 0,330. Ketidakkonsistenan temuan ini mengindikasikan adanya faktor-faktor lain yang memoderasi atau memediasi hubungan tersebut, yang memerlukan kajian lebih mendalam.', { indent: true }),
  emptyLine(),
  p('Selain manajemen waktu dan kerja paruh waktu, lingkungan kampus juga menjadi faktor penting yang memengaruhi prestasi akademik mahasiswa. Lingkungan kampus mencakup kualitas fasilitas belajar, suasana akademik, kualitas pengajaran dosen, serta kondisi fisik dan sosial kampus secara keseluruhan. Robyansyah et al. (2022) dalam penelitiannya terhadap 107 mahasiswa Politeknik Negeri Bengkalis menemukan bahwa lingkungan belajar berpengaruh positif dan signifikan langsung terhadap prestasi belajar dengan koefisien 0,498 (sig. = 0,000), dan secara keseluruhan model menjelaskan 82,9% variasi prestasi. Dewi et al. (2025) juga menemukan pengaruh positif signifikan lingkungan kampus terhadap prestasi mahasiswa dengan koefisien 0,171 (t = 2,824; p = 0,005) dalam model yang menjelaskan 78% variasi prestasi.', { indent: true }),
  emptyLine(),
  p('Research gap dalam penelitian ini terletak pada beberapa hal. Pertama, penelitian terdahulu yang mengkaji ketiga variabel secara simultan, yaitu manajemen waktu, kerja paruh waktu, dan lingkungan kampus dalam satu model penelitian masih sangat terbatas. Sebagian besar penelitian hanya mengkaji dua variabel secara bersamaan. Kedua, penelitian tentang prestasi akademik mahasiswa dengan ketiga variabel tersebut di lingkungan Universitas Malikussaleh, khususnya Prodi Manajemen FEB, belum pernah dilakukan. Ketiga, konteks Aceh yang memiliki karakteristik sosial-ekonomi dan budaya yang unik memerlukan kajian tersendiri yang tidak dapat digeneralisasikan begitu saja dari penelitian yang dilakukan di daerah lain.', { indent: true }),
  emptyLine(),
  p('Berdasarkan uraian latar belakang di atas, peneliti tertarik untuk melakukan penelitian dengan judul "Pengaruh Manajemen Waktu, Kerja Paruh Waktu, dan Lingkungan Kampus terhadap Prestasi Akademik Mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh." Penelitian ini diharapkan dapat memberikan kontribusi teoritis maupun praktis dalam memahami faktor-faktor yang memengaruhi prestasi akademik mahasiswa, serta memberikan rekomendasi bagi pihak universitas dalam merancang kebijakan yang mendukung peningkatan prestasi mahasiswa secara menyeluruh.', { indent: true }),
  emptyLine(),

  // 1.2 Rumusan Masalah
  heading2('1.2 Rumusan Masalah Penelitian'),
  p('Berdasarkan latar belakang yang telah diuraikan di atas, maka rumusan masalah dalam penelitian ini adalah sebagai berikut:', { indent: true }),
  numbered('Apakah manajemen waktu berpengaruh terhadap prestasi akademik mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh?', '1'),
  numbered('Apakah kerja paruh waktu berpengaruh terhadap prestasi akademik mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh?', '2'),
  numbered('Apakah lingkungan kampus berpengaruh terhadap prestasi akademik mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh?', '3'),
  numbered('Apakah manajemen waktu, kerja paruh waktu, dan lingkungan kampus secara simultan berpengaruh terhadap prestasi akademik mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh?', '4'),
  emptyLine(),

  // 1.3 Tujuan Penelitian
  heading2('1.3 Tujuan Penelitian'),
  p('Berdasarkan rumusan masalah di atas, tujuan penelitian ini adalah:', { indent: true }),
  numbered('Untuk mengetahui dan menganalisis pengaruh manajemen waktu terhadap prestasi akademik mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh.', '1'),
  numbered('Untuk mengetahui dan menganalisis pengaruh kerja paruh waktu terhadap prestasi akademik mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh.', '2'),
  numbered('Untuk mengetahui dan menganalisis pengaruh lingkungan kampus terhadap prestasi akademik mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh.', '3'),
  numbered('Untuk mengetahui dan menganalisis pengaruh manajemen waktu, kerja paruh waktu, dan lingkungan kampus secara simultan terhadap prestasi akademik mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh.', '4'),
  emptyLine(),

  // 1.4 Manfaat Penelitian
  heading2('1.4 Manfaat Penelitian'),

  heading3('1.4.1 Manfaat Teoritis'),
  p('Secara teoritis, hasil penelitian ini diharapkan memberikan manfaat sebagai berikut:', { indent: true }),
  numbered('Menambah khasanah ilmu pengetahuan dan memperkaya literatur di bidang manajemen sumber daya manusia, khususnya yang berkaitan dengan faktor-faktor yang memengaruhi prestasi akademik mahasiswa.', '1'),
  numbered('Memberikan kontribusi teoritis dalam pengembangan Self-Determination Theory (SDT) yang dikemukakan oleh Deci dan Ryan, khususnya dalam konteks pengaruh faktor internal (manajemen waktu) dan eksternal (kerja paruh waktu dan lingkungan kampus) terhadap motivasi dan prestasi belajar mahasiswa di perguruan tinggi.', '2'),
  numbered('Dapat dijadikan sebagai referensi dan bahan pertimbangan bagi penelitian-penelitian selanjutnya yang mengkaji topik serupa, baik di lingkungan Universitas Malikussaleh maupun di perguruan tinggi lainnya.', '3'),
  emptyLine(),

  heading3('1.4.2 Manfaat Praktis'),
  p('Secara praktis, hasil penelitian ini diharapkan memberikan manfaat sebagai berikut:', { indent: true }),
  numbered('Bagi Mahasiswa: Memberikan pemahaman kepada mahasiswa tentang pentingnya manajemen waktu yang efektif dalam menjaga keseimbangan antara aktivitas akademik dan non-akademik, termasuk bagi mereka yang memilih untuk bekerja paruh waktu selama masa studi.', '1'),
  numbered('Bagi Program Studi Manajemen FEB UNIMAL: Memberikan masukan dan rekomendasi kebijakan terkait upaya peningkatan prestasi akademik mahasiswa melalui program-program pengembangan soft skills, termasuk pelatihan manajemen waktu dan optimalisasi lingkungan kampus yang kondusif bagi belajar.', '2'),
  numbered('Bagi Universitas Malikussaleh: Memberikan gambaran komprehensif tentang kondisi akademik mahasiswa, khususnya terkait dengan faktor-faktor yang memengaruhi prestasi akademik, sehingga dapat menjadi dasar pengambilan kebijakan strategis di tingkat universitas.', '3'),
  numbered('Bagi Peneliti Selanjutnya: Memberikan data empiris dan kerangka konseptual yang dapat digunakan sebagai pijakan untuk penelitian lanjutan yang lebih mendalam, baik dengan menambah variabel mediasi/moderasi maupun dengan memperluas cakupan populasi penelitian.', '4'),

  pageBreak(),

  // ==================== BAB II ====================
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 480, before: 0, after: 120 },
    children: [new TextRun({ text: 'BAB II', bold: true, size: 24, font: 'Times New Roman' })]
  }),
  heading1('TINJAUAN PUSTAKA'),
  emptyLine(),

  // 2.1 Landasan Teori
  heading2('2.1 Landasan Teori'),

  heading3('2.1.1 Self-Determination Theory (SDT)'),
  p('Self-Determination Theory (SDT) atau Teori Penentuan Diri merupakan grand theory dalam penelitian ini. Teori ini dikembangkan oleh Edward L. Deci dan Richard M. Ryan pada tahun 1985 dan terus disempurnakan hingga saat ini. SDT merupakan teori motivasi makro yang berfokus pada bagaimana faktor-faktor sosial dan lingkungan memfasilitasi atau menghalangi motivasi intrinsik, pengembangan diri, dan kesejahteraan psikologis individu (Ryan & Deci, 2000).', { indent: true }),
  emptyLine(),
  p('SDT berargumen bahwa manusia memiliki tiga kebutuhan psikologis dasar yang bersifat universal, yaitu: (1) kebutuhan akan kompetensi (competence), yaitu kebutuhan untuk merasa mampu dan efektif dalam berinteraksi dengan lingkungan; (2) kebutuhan akan otonomi (autonomy), yaitu kebutuhan untuk merasa memiliki kebebasan dan inisiatif dalam tindakan; dan (3) kebutuhan akan keterkaitan (relatedness), yaitu kebutuhan untuk merasa terhubung dan diterima oleh orang lain. Ketika ketiga kebutuhan ini terpenuhi, individu cenderung menunjukkan motivasi yang lebih tinggi, performa yang lebih baik, dan kesejahteraan psikologis yang lebih optimal (Deci & Ryan, 2000; Ryan & Deci, 2017).', { indent: true }),
  emptyLine(),
  p('Dalam konteks pendidikan, SDT telah banyak digunakan untuk menjelaskan variasi prestasi akademik di kalangan mahasiswa. Howard et al. (2021) dalam meta-analisis mereka yang mencakup lebih dari 100.000 partisipan menemukan bahwa motivasi otonom yang bersumber dari kepuasan kebutuhan psikologis dasar berhubungan positif dan signifikan dengan prestasi akademik. Selanjutnya, Bureau et al. (2022) dalam meta-analisis terhadap 344 studi menemukan bahwa dukungan otonomi dari lingkungan (termasuk lingkungan kampus) merupakan prediktor kuat motivasi otonom yang kemudian memengaruhi prestasi.', { indent: true }),
  emptyLine(),
  p('Keterkaitan SDT dengan penelitian ini terletak pada beberapa hal. Pertama, manajemen waktu sebagai variabel bebas pertama merupakan manifestasi dari kebutuhan kompetensi dan otonomi dalam SDT. Mahasiswa yang mampu mengelola waktu dengan efektif menunjukkan kompetensi diri yang tinggi dan otonomi dalam mengatur aktivitas belajar mereka, yang pada gilirannya mendukung prestasi akademik yang lebih baik. Kedua, kerja paruh waktu yang dilakukan mahasiswa merupakan respons terhadap dorongan eksternal (kebutuhan finansial) yang dapat memengaruhi kualitas dan kuantitas waktu belajar, sehingga berkaitan langsung dengan pemenuhan kebutuhan kompetensi akademik. Ketiga, lingkungan kampus yang kondusif merupakan faktor ekstrinsik yang dalam kerangka SDT berperan sebagai konteks yang memfasilitasi atau menghambat pemenuhan kebutuhan psikologis dasar mahasiswa, yang selanjutnya memengaruhi motivasi dan prestasi akademik mereka.', { indent: true }),
  emptyLine(),

  heading3('2.1.2 Prestasi Akademik'),

  heading4('2.1.2.1 Definisi Prestasi Akademik'),
  p('Prestasi akademik merupakan hasil yang dicapai oleh mahasiswa dalam proses pembelajaran yang mencerminkan tingkat penguasaan pengetahuan, keterampilan, dan sikap yang telah dipelajari selama menempuh pendidikan formal di perguruan tinggi. Prestasi akademik pada umumnya diukur melalui nilai atau skor yang diperoleh mahasiswa dari serangkaian evaluasi dan penilaian yang dilakukan selama proses perkuliahan.', { indent: true }),
  emptyLine(),
  p('Menurut Zimmerman (2002), prestasi akademik tidak sekadar mencerminkan kemampuan kognitif atau kecerdasan seseorang, melainkan merupakan produk dari interaksi kompleks antara faktor internal seperti motivasi, strategi belajar, dan pengaturan diri, serta faktor eksternal seperti lingkungan belajar dan dukungan sosial. Dalam konteks pendidikan tinggi di Indonesia, prestasi akademik mahasiswa umumnya direpresentasikan melalui Indeks Prestasi Semester (IPS) dan Indeks Prestasi Kumulatif (IPK) yang dihitung berdasarkan nilai akhir seluruh mata kuliah yang telah ditempuh.', { indent: true }),
  emptyLine(),
  p('Nana Sudjana (2014) mendefinisikan prestasi belajar sebagai kemampuan-kemampuan yang dimiliki siswa atau mahasiswa setelah ia menerima pengalaman belajarnya. Lebih lanjut, Hamalik (2016) menyatakan bahwa prestasi akademik merupakan ukuran keberhasilan suatu proses belajar yang mencerminkan perubahan pada aspek pengetahuan (kognitif), sikap (afektif), dan keterampilan (psikomotor) yang diperoleh mahasiswa sebagai hasil dari kegiatan belajar di perguruan tinggi.', { indent: true }),
  emptyLine(),
  p('Dalam penelitian ini, prestasi akademik mahasiswa dioperasionalisasikan melalui Indeks Prestasi Kumulatif (IPK) yang diperoleh mahasiswa hingga semester terakhir yang ditempuh. IPK merupakan indikator prestasi akademik yang paling umum digunakan dalam penelitian pendidikan tinggi di Indonesia karena bersifat objektif, terukur, dan dapat dibandingkan secara konsisten antar mahasiswa dan antar periode waktu.', { indent: true }),
  emptyLine(),

  heading4('2.1.2.2 Faktor-Faktor yang Memengaruhi Prestasi Akademik'),
  p('Prestasi akademik mahasiswa dipengaruhi oleh berbagai faktor, baik yang bersumber dari dalam diri mahasiswa (faktor internal) maupun dari luar diri mahasiswa (faktor eksternal). Pemahaman tentang faktor-faktor ini penting untuk merancang intervensi yang tepat guna meningkatkan prestasi akademik.', { indent: true }),
  emptyLine(),
  p('Faktor-faktor internal yang memengaruhi prestasi akademik meliputi:', { indent: true }),
  numbered('Kemampuan kognitif dan kecerdasan, yang merupakan faktor bawaan yang memengaruhi kemampuan mahasiswa dalam memahami dan memproses informasi.', 'a'),
  numbered('Motivasi belajar, yang mencakup motivasi intrinsik dan ekstrinsik yang mendorong mahasiswa untuk belajar dengan sungguh-sungguh.', 'b'),
  numbered('Manajemen waktu, yaitu kemampuan mahasiswa dalam merencanakan, mengorganisasikan, dan menggunakan waktu belajar secara efektif dan efisien.', 'c'),
  numbered('Efikasi diri (self-efficacy), yaitu keyakinan mahasiswa terhadap kemampuan mereka untuk berhasil dalam tugas-tugas akademik.', 'd'),
  numbered('Disiplin dan kebiasaan belajar, yang mencakup konsistensi dan regularitas dalam melakukan aktivitas belajar.', 'e'),
  emptyLine(),
  p('Adapun faktor-faktor eksternal yang memengaruhi prestasi akademik meliputi:', { indent: true }),
  numbered('Lingkungan kampus, mencakup kualitas fasilitas belajar, suasana akademik, kualitas pengajaran, dan kondisi fisik kampus.', 'a'),
  numbered('Dukungan keluarga, yang mencakup dukungan emosional, finansial, dan motivasional dari orang tua dan anggota keluarga lainnya.', 'b'),
  numbered('Status pekerjaan, yaitu apakah mahasiswa juga bekerja di luar jam kuliah (bekerja paruh waktu) yang dapat memengaruhi ketersediaan waktu belajar.', 'c'),
  numbered('Interaksi sosial, mencakup hubungan dengan dosen, teman sebaya, dan komunitas kampus yang dapat mendukung atau menghambat proses belajar.', 'd'),
  numbered('Kondisi sosial-ekonomi, yang memengaruhi akses mahasiswa terhadap sumber daya belajar dan tingkat stres finansial yang dihadapi.', 'e'),
  emptyLine(),

  heading4('2.1.2.3 Indikator Prestasi Akademik'),
  p('Pengukuran prestasi akademik dalam penelitian ini menggunakan beberapa indikator yang dikembangkan berdasarkan kajian literatur dari berbagai sumber. Mengacu pada Agustina dan Mardalis (2024), Dewi et al. (2025), serta Inayah et al. (2023), indikator prestasi akademik yang digunakan dalam penelitian ini adalah:', { indent: true }),
  numbered('Indeks Prestasi Kumulatif (IPK), yaitu rata-rata nilai tertimbang seluruh mata kuliah yang telah ditempuh mahasiswa hingga semester terakhir yang diselesaikan.', '1'),
  numbered('Ketepatan penyelesaian tugas, yaitu kemampuan mahasiswa dalam menyelesaikan dan mengumpulkan tugas-tugas akademik tepat waktu sesuai tenggat yang ditetapkan dosen.', '2'),
  numbered('Keaktifan dalam perkuliahan, yang mencakup partisipasi aktif mahasiswa dalam diskusi kelas, tanya-jawab, dan kegiatan pembelajaran lainnya.', '3'),
  numbered('Tingkat kehadiran, yaitu persentase kehadiran mahasiswa dalam setiap pertemuan perkuliahan selama satu semester.', '4'),
  emptyLine(),

  heading3('2.1.3 Manajemen Waktu'),

  heading4('2.1.3.1 Definisi Manajemen Waktu'),
  p('Manajemen waktu adalah proses pengendalian dan pemanfaatan waktu secara terencana dan sistematis guna mencapai tujuan yang telah ditetapkan secara efektif dan efisien. Dalam konteks akademik, manajemen waktu merujuk pada kemampuan mahasiswa untuk mengalokasikan, mengatur, dan memanfaatkan waktu yang tersedia secara optimal antara kegiatan belajar, aktivitas sosial, kegiatan ekstrakurikuler, dan apabila relevan, kegiatan bekerja.', { indent: true }),
  emptyLine(),
  p('Macan (1994) mendefinisikan manajemen waktu sebagai kemampuan untuk mengalokasikan waktu secara efektif di antara berbagai kegiatan dengan menetapkan tujuan, merencanakan kegiatan, dan memprioritaskan tugas. Definisi ini menekankan tiga komponen utama manajemen waktu, yaitu penetapan tujuan, perencanaan kegiatan, dan prioritisasi tugas. Sementara itu, Kamilatunnisa et al. (2025) mendefinisikan manajemen waktu sebagai kemampuan individu dalam menggunakan waktu secara sadar dan terarah untuk mencapai tujuan hidup, yang meliputi kemampuan merencanakan, mengorganisasikan, melaksanakan, dan mengevaluasi penggunaan waktu.', { indent: true }),
  emptyLine(),
  p('Dalam penelitian ini, manajemen waktu didefinisikan sebagai kemampuan mahasiswa dalam merencanakan, mengorganisasikan, mengarahkan, dan mengendalikan penggunaan waktu sehari-hari secara efektif guna memenuhi tuntutan akademik dan non-akademik secara seimbang, sehingga dapat mencapai prestasi belajar yang optimal.', { indent: true }),
  emptyLine(),

  heading4('2.1.3.2 Faktor-Faktor yang Memengaruhi Manajemen Waktu'),
  p('Kemampuan manajemen waktu seseorang dipengaruhi oleh berbagai faktor, baik internal maupun eksternal. Beberapa faktor yang memengaruhi manajemen waktu mahasiswa antara lain:', { indent: true }),
  numbered('Motivasi dan kesadaran diri: Mahasiswa dengan motivasi intrinsik yang tinggi cenderung lebih teratur dalam mengelola waktu karena memiliki kesadaran akan pentingnya efisiensi waktu bagi pencapaian tujuan akademik mereka.', '1'),
  numbered('Beban kerja dan aktivitas: Semakin banyak aktivitas yang dilakukan mahasiswa di luar perkuliahan, seperti bekerja paruh waktu atau berorganisasi, semakin kompleks tantangan manajemen waktu yang dihadapi.', '2'),
  numbered('Lingkungan belajar: Kondisi lingkungan yang mendukung, seperti tersedianya ruang belajar yang nyaman dan fasilitas yang memadai, memudahkan mahasiswa dalam mengalokasikan waktu untuk belajar secara efektif.', '3'),
  numbered('Teknologi dan media sosial: Penggunaan teknologi dan media sosial yang tidak terkontrol dapat menjadi \u201ctime trap\u201d yang mengganggu efektivitas manajemen waktu mahasiswa.', '4'),
  numbered('Dukungan sosial: Dukungan dari keluarga, dosen, dan teman sebaya dapat membantu mahasiswa dalam menetapkan prioritas dan mengelola waktu dengan lebih baik.', '5'),
  emptyLine(),

  heading4('2.1.3.3 Indikator Manajemen Waktu'),
  p('Pengukuran manajemen waktu dalam penelitian ini mengacu pada dimensi yang dikembangkan oleh Macan (1994) dan dimodifikasi oleh berbagai peneliti berikutnya. Berdasarkan Kamilatunnisa et al. (2025) dan Andis Tira et al. (2025), indikator manajemen waktu yang digunakan dalam penelitian ini adalah:', { indent: true }),
  numbered('Menetapkan tujuan dan prioritas (goal setting and prioritizing): kemampuan mahasiswa dalam menetapkan tujuan belajar yang jelas dan menentukan skala prioritas dalam penggunaan waktu.', '1'),
  numbered('Membuat daftar tugas (making to-do list): kebiasaan mahasiswa dalam membuat perencanaan tertulis tentang tugas-tugas yang harus diselesaikan beserta tenggat waktunya.', '2'),
  numbered('Penjadwalan (scheduling): kemampuan mahasiswa dalam membuat jadwal kegiatan harian, mingguan, dan semesteran yang mengalokasikan waktu secara proporsional untuk belajar dan aktivitas lainnya.', '3'),
  numbered('Menghindari penundaan (avoiding procrastination): kemampuan mahasiswa untuk tidak menunda-nunda penyelesaian tugas dan memulai pekerjaan tepat waktu.', '4'),
  numbered('Pengorganisasian diri (self-organization): kemampuan mahasiswa dalam mengorganisasikan tempat belajar, materi perkuliahan, dan sumber daya belajar agar dapat diakses dengan mudah dan efisien.', '5'),
  emptyLine(),

  heading3('2.1.4 Kerja Paruh Waktu'),

  heading4('2.1.4.1 Definisi Kerja Paruh Waktu'),
  p('Kerja paruh waktu (part-time work) adalah pekerjaan yang dilakukan dengan jam kerja lebih sedikit dibandingkan jam kerja penuh waktu (full-time), umumnya kurang dari 35 jam per minggu. Dalam konteks mahasiswa, kerja paruh waktu merujuk pada aktivitas bekerja yang dilakukan di samping kegiatan perkuliahan formal, baik yang bersifat tetap maupun tidak tetap, dengan tujuan utama untuk memperoleh penghasilan.', { indent: true }),
  emptyLine(),
  p('Huda et al. (2023) mendefinisikan kerja paruh waktu sebagai pekerjaan yang dilakukan oleh mahasiswa di luar jam kuliah, baik di sektor formal maupun informal, dengan waktu kerja yang terbatas sehingga tidak mengganggu jadwal perkuliahan secara keseluruhan. Menurut Agustina dan Mardalis (2024), kerja paruh waktu mahasiswa adalah keterlibatan mahasiswa dalam kegiatan ekonomi produktif yang memberikan imbalan finansial, yang dilakukan secara bersamaan dengan kegiatan belajar di perguruan tinggi.', { indent: true }),
  emptyLine(),
  p('Dalam penelitian ini, kerja paruh waktu didefinisikan sebagai segala bentuk aktivitas bekerja yang dilakukan mahasiswa Prodi Manajemen FEB UNIMAL di luar jam perkuliahan resmi, baik bersifat tetap maupun tidak tetap, dengan jam kerja tidak melebihi 35 jam per minggu, dengan tujuan memperoleh penghasilan untuk memenuhi kebutuhan finansial selama menempuh pendidikan.', { indent: true }),
  emptyLine(),

  heading4('2.1.4.2 Faktor-Faktor yang Mendorong Kerja Paruh Waktu'),
  p('Keputusan mahasiswa untuk bekerja paruh waktu dipengaruhi oleh berbagai faktor. Berdasarkan kajian Huda et al. (2023), Agustina dan Mardalis (2024), serta Linggasari dan Kurniawan (2019), faktor-faktor yang mendorong mahasiswa untuk bekerja paruh waktu antara lain:', { indent: true }),
  numbered('Faktor ekonomi: kebutuhan finansial untuk membiayai uang kuliah tunggal (UKT), biaya hidup, dan kebutuhan personal merupakan motivasi utama mahasiswa bekerja paruh waktu. Hal ini terutama relevan bagi mahasiswa yang berasal dari keluarga berpenghasilan menengah ke bawah.', '1'),
  numbered('Pengembangan pengalaman kerja: banyak mahasiswa yang bekerja paruh waktu untuk mendapatkan pengalaman praktis di bidang yang relevan dengan jurusan mereka, sebagai persiapan memasuki dunia kerja setelah lulus.', '2'),
  numbered('Pengembangan jaringan sosial (networking): bekerja memberikan kesempatan kepada mahasiswa untuk membangun relasi profesional yang dapat bermanfaat di masa mendatang.', '3'),
  numbered('Dorongan kemandirian: beberapa mahasiswa memilih bekerja sebagai manifestasi dari kemandirian dan keinginan untuk tidak bergantung sepenuhnya pada orang tua.', '4'),
  numbered('Kesempatan yang tersedia: adanya peluang kerja paruh waktu yang cocok dengan jadwal perkuliahan, seperti menjadi tentor, barista, atau pekerja lepas di bidang kreatif, mendorong mahasiswa untuk memanfaatkan kesempatan tersebut.', '5'),
  emptyLine(),

  heading4('2.1.4.3 Indikator Kerja Paruh Waktu'),
  p('Pengukuran kerja paruh waktu dalam penelitian ini mengacu pada dimensi yang dikembangkan oleh Huda et al. (2023) dan Agustina dan Mardalis (2024). Indikator kerja paruh waktu yang digunakan dalam penelitian ini adalah:', { indent: true }),
  numbered('Jumlah jam kerja per minggu: rata-rata waktu yang dihabiskan untuk bekerja dalam satu minggu, yang memengaruhi ketersediaan waktu untuk belajar dan beristirahat.', '1'),
  numbered('Jenis pekerjaan: karakteristik dan sifat pekerjaan yang dilakukan, apakah relevan dengan bidang studi atau tidak, serta tingkat fleksibilitas jadwal kerja yang tersedia.', '2'),
  numbered('Pengaruh pekerjaan terhadap jadwal kuliah: sejauh mana pekerjaan yang dilakukan memengaruhi kehadiran dan partisipasi mahasiswa dalam kegiatan perkuliahan.', '3'),
  numbered('Motivasi bekerja: alasan atau tujuan utama mahasiswa memilih untuk bekerja paruh waktu, apakah didorong oleh kebutuhan finansial, pengembangan diri, atau keduanya.', '4'),
  numbered('Manfaat finansial: kontribusi penghasilan dari pekerjaan terhadap pemenuhan kebutuhan finansial selama menempuh pendidikan.', '5'),
  emptyLine(),

  heading3('2.1.5 Lingkungan Kampus'),

  heading4('2.1.5.1 Definisi Lingkungan Kampus'),
  p('Lingkungan kampus merupakan keseluruhan kondisi fisik, sosial, dan akademik yang ada di dalam kampus yang secara langsung maupun tidak langsung memengaruhi proses belajar dan perkembangan mahasiswa. Lingkungan kampus mencakup aspek-aspek seperti kondisi fasilitas fisik, kualitas pengajaran, iklim akademik, hubungan interpersonal di dalam kampus, serta kebijakan dan program yang dijalankan oleh institusi pendidikan.', { indent: true }),
  emptyLine(),
  p('Robyansyah et al. (2022) mendefinisikan lingkungan belajar sebagai segala sesuatu yang ada di sekitar peserta didik, baik yang bersifat fisik maupun non-fisik, yang berpotensi memberikan pengaruh terhadap proses dan hasil belajar. Dalam konteks perguruan tinggi, Munira et al. (2024) mendeskripsikan lingkungan kampus sebagai kondisi keseluruhan yang mencakup fasilitas belajar, kualitas pengajaran, dan suasana akademik yang secara bersama-sama memengaruhi prestasi akademik mahasiswa. Dewi et al. (2025) mengoperasionalisasikan lingkungan kampus sebagai persepsi mahasiswa terhadap kondisi fisik kampus, dukungan akademik, dan iklim sosial yang dirasakan selama menempuh pendidikan.', { indent: true }),
  emptyLine(),
  p('Dalam penelitian ini, lingkungan kampus didefinisikan sebagai keseluruhan kondisi dan situasi di dalam kampus, yang mencakup kondisi fisik fasilitas kampus, kualitas pembelajaran yang diberikan oleh dosen, suasana akademik, serta hubungan antarmahasiswa dan antara mahasiswa dengan dosen, yang secara bersama-sama membentuk iklim belajar yang dapat mendukung atau menghambat pencapaian prestasi akademik mahasiswa.', { indent: true }),
  emptyLine(),

  heading4('2.1.5.2 Faktor-Faktor Lingkungan Kampus yang Memengaruhi Prestasi Akademik'),
  p('Lingkungan kampus yang memengaruhi prestasi akademik mahasiswa dapat dikelompokkan ke dalam beberapa dimensi utama. Berdasarkan kajian Robyansyah et al. (2022), Berhanu dan Sewagegn (2024), serta Dewi et al. (2025), faktor-faktor lingkungan kampus yang memengaruhi prestasi akademik antara lain:', { indent: true }),
  numbered('Kondisi fisik kampus: mencakup ketersediaan dan kualitas ruang kelas, perpustakaan, laboratorium, akses internet, serta fasilitas penunjang belajar lainnya yang memadai.', '1'),
  numbered('Kualitas pengajaran: mencakup kompetensi akademik dan pedagogis dosen, metode pengajaran yang digunakan, serta kualitas materi dan bahan ajar yang disediakan.', '2'),
  numbered('Iklim akademik: mencakup suasana belajar di kelas, budaya akademik yang berlaku di kampus, serta norma dan nilai yang mengedepankan prestasi dan integritas akademik.', '3'),
  numbered('Hubungan sosial: mencakup kualitas interaksi antara mahasiswa dengan dosen, antarsesama mahasiswa, dan antara mahasiswa dengan staf administrasi kampus.', '4'),
  numbered('Dukungan institusional: mencakup kebijakan, program, dan layanan yang disediakan oleh institusi untuk mendukung keberhasilan akademik mahasiswa, seperti program bimbingan akademik, beasiswa, dan konseling.', '5'),
  emptyLine(),

  heading4('2.1.5.3 Indikator Lingkungan Kampus'),
  p('Pengukuran lingkungan kampus dalam penelitian ini mengacu pada dimensi yang dikembangkan oleh Robyansyah et al. (2022), Munira et al. (2024), dan Dewi et al. (2025). Indikator lingkungan kampus yang digunakan dalam penelitian ini adalah:', { indent: true }),
  numbered('Kenyamanan dan kelengkapan fasilitas fisik: persepsi mahasiswa terhadap kondisi dan kelengkapan fasilitas fisik kampus seperti ruang kelas, perpustakaan, laboratorium komputer, dan akses internet yang tersedia.', '1'),
  numbered('Kualitas proses pembelajaran: persepsi mahasiswa terhadap kualitas pengajaran yang diberikan oleh dosen, mencakup kejelasan penjelasan materi, penggunaan metode belajar yang variatif, dan ketepatan waktu dalam mengajar.', '2'),
  numbered('Suasana dan iklim akademik: persepsi mahasiswa terhadap kondisi suasana belajar di lingkungan kampus, termasuk tingkat kompetisi sehat, motivasi kolektif antar mahasiswa, dan budaya belajar yang berlaku.', '3'),
  numbered('Interaksi sosial akademik: kualitas hubungan antara mahasiswa dengan dosen dan sesama mahasiswa dalam konteks akademik, termasuk keterbukaan dosen dalam berdiskusi dan kemudahan untuk berkonsultasi.', '4'),
  numbered('Dukungan administrasi dan layanan kampus: kualitas layanan administrasi akademik, kemudahan akses informasi akademik, dan responsivitas staf kampus dalam membantu keperluan mahasiswa.', '5'),
  emptyLine(),

  // 2.2 Penelitian Terdahulu
  heading2('2.2 Penelitian Terdahulu'),
  p('Penelitian terdahulu yang relevan dengan topik penelitian ini disajikan pada tabel berikut:', { indent: true }),
  emptyLine(),

  // Research table
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [540, 1620, 1440, 2160, 2160, 1440],
    rows: [
      new TableRow({
        children: [
          tableCell('No.', { width: 540, header: true, center: true }),
          tableCell('Peneliti & Tahun', { width: 1620, header: true }),
          tableCell('Judul', { width: 1440, header: true }),
          tableCell('Metode', { width: 2160, header: true }),
          tableCell('Hasil Penelitian', { width: 2160, header: true }),
          tableCell('Persamaan & Perbedaan', { width: 1440, header: true }),
        ]
      }),
      new TableRow({
        children: [
          tableCell('1', { width: 540, center: true }),
          tableCell('Kamilatunnisa et al. (2025)\nDOI: 10.62383/katalis.v2i3.2072', { width: 1620 }),
          tableCell('Analisis Pengaruh Perilaku Manajemen Waktu terhadap Prestasi Akademik Mahasiswa', { width: 1440 }),
          tableCell('Kuantitatif, regresi linear sederhana, n=50 mahasiswa FE UNPAS', { width: 2160 }),
          tableCell('Manajemen waktu berpengaruh positif signifikan terhadap prestasi akademik (R2=0,505; p=0,000); Y=10,742+0,844X', { width: 2160 }),
          tableCell('Persamaan: variabel manajemen waktu dan prestasi akademik. Perbedaan: tidak ada variabel kerja paruh waktu dan lingkungan kampus', { width: 1440 }),
        ]
      }),
      new TableRow({
        children: [
          tableCell('2', { width: 540, center: true }),
          tableCell('Andis Tira et al. (2025)\nDOI: 10.60036/jbm.689', { width: 1620 }),
          tableCell('Pengaruh Manajemen Waktu, Efikasi Diri, dan Lingkungan Kampus terhadap Prestasi Akademik', { width: 1440 }),
          tableCell('Kuantitatif, regresi berganda, n=90 mahasiswa Teknik Mesin UR', { width: 2160 }),
          tableCell('Manajemen waktu (p=0,003; t=3,114) dan lingkungan kampus berpengaruh signifikan terhadap prestasi akademik; R2adj=0,643', { width: 2160 }),
          tableCell('Persamaan: variabel manajemen waktu, lingkungan kampus, dan prestasi akademik. Perbedaan: menggunakan efikasi diri sebagai variabel bebas, tidak ada variabel kerja paruh waktu', { width: 1440 }),
        ]
      }),
      new TableRow({
        children: [
          tableCell('3', { width: 540, center: true }),
          tableCell('Agustina & Mardalis (2024)\nDOI: 10.46306/jbbe.v17i2.556', { width: 1620 }),
          tableCell('Pengaruh Kerja Paruh Waktu, Motivasi Belajar dan Time Management Terhadap Prestasi Akademik', { width: 1440 }),
          tableCell('Kuantitatif PLS-SEM, n=250 mahasiswa bekerja part-time', { width: 2160 }),
          tableCell('Kerja paruh waktu (original sample 0,378; t=4,637; p=0,000) dan manajemen waktu berpengaruh positif signifikan terhadap prestasi akademik', { width: 2160 }),
          tableCell('Persamaan: variabel kerja paruh waktu, manajemen waktu, dan prestasi akademik. Perbedaan: tidak ada variabel lingkungan kampus, menambah motivasi belajar', { width: 1440 }),
        ]
      }),
      new TableRow({
        children: [
          tableCell('4', { width: 540, center: true }),
          tableCell('Inayah et al. (2023)\nDOI: 10.56799/peshum.v2i2.1391', { width: 1620 }),
          tableCell('Pengaruh Manajemen Waktu Terhadap Prestasi Akademik Mahasiswa yang Bekerja di Kota Makassar', { width: 1440 }),
          tableCell('Kuantitatif, regresi ordinal, n=89 mahasiswa pekerja Makassar', { width: 2160 }),
          tableCell('Manajemen waktu berpengaruh positif signifikan terhadap prestasi akademik (Nagelkerke R2=0,154; p=0,005); 71,9% IPK mahasiswa bekerja tergolong baik', { width: 2160 }),
          tableCell('Persamaan: variabel manajemen waktu, konteks mahasiswa yang bekerja. Perbedaan: tidak ada variabel lingkungan kampus, fokus pada mahasiswa di satu kota', { width: 1440 }),
        ]
      }),
      new TableRow({
        children: [
          tableCell('5', { width: 540, center: true }),
          tableCell('Huda et al. (2023)\nDOI: 10.5281/zenodo.8127903', { width: 1620 }),
          tableCell('Pengaruh Kerja Paruh Waktu Terhadap Prestasi Akademik Mahasiswa UIN SU', { width: 1440 }),
          tableCell('Kuantitatif, kuesioner Likert, n=109 mahasiswa UINSU', { width: 2160 }),
          tableCell('Kerja paruh waktu berpengaruh negatif signifikan terhadap prestasi akademik; skor terendah pada kemampuan menjaga kualitas tugas akademik', { width: 2160 }),
          tableCell('Persamaan: variabel kerja paruh waktu dan prestasi akademik. Perbedaan: tidak ada variabel manajemen waktu dan lingkungan kampus', { width: 1440 }),
        ]
      }),
      new TableRow({
        children: [
          tableCell('6', { width: 540, center: true }),
          tableCell('Dewi et al. (2025)\nDOI: 10.33480/jasdim.v4i1.6666', { width: 1620 }),
          tableCell('Efikasi Diri sebagai Mediasi dalam Hubungan Lingkungan Kampus dan Manajemen Waktu pada Prestasi Mahasiswa', { width: 1440 }),
          tableCell('PLS-SEM SmartPLS 4.0, n=270 mahasiswa S1 Manajemen FEB Universitas Jember', { width: 2160 }),
          tableCell('Lingkungan kampus berpengaruh positif signifikan langsung (koef=0,171; t=2,824; p=0,005) dan tidak langsung via efikasi diri (indirect=0,079) terhadap prestasi; R2=0,780', { width: 2160 }),
          tableCell('Persamaan: variabel manajemen waktu, lingkungan kampus, prestasi mahasiswa, mahasiswa manajemen. Perbedaan: menggunakan efikasi diri sebagai mediator, tidak ada variabel kerja paruh waktu', { width: 1440 }),
        ]
      }),
      new TableRow({
        children: [
          tableCell('7', { width: 540, center: true }),
          tableCell('Robyansyah et al. (2022)\nJurnal Daya Saing Vol.8 No.3', { width: 1620 }),
          tableCell('Pengaruh Lingkungan Belajar dan Disiplin Terhadap Motivasi Belajar dan Prestasi Belajar', { width: 1440 }),
          tableCell('Path analysis SPSS, n=107 mahasiswa Politeknik Negeri Bengkalis', { width: 2160 }),
          tableCell('Lingkungan belajar berpengaruh positif signifikan terhadap prestasi belajar (koef=0,498; t=6,051; sig=0,000); R2=0,829', { width: 2160 }),
          tableCell('Persamaan: variabel lingkungan belajar dan prestasi akademik. Perbedaan: tidak ada variabel manajemen waktu dan kerja paruh waktu, menambah disiplin dan motivasi', { width: 1440 }),
        ]
      }),
      new TableRow({
        children: [
          tableCell('8', { width: 540, center: true }),
          tableCell('Linggasari & Kurniawan (2019)\nJUPE Vol.7 No.3', { width: 1620 }),
          tableCell('Hubungan Kerja Paruh Waktu dengan Prestasi Akademik Mahasiswa Pendidikan Ekonomi UNESA', { width: 1440 }),
          tableCell('Kuantitatif deskriptif, korelasi Karl Pearson, sensus n=56 mahasiswa', { width: 2160 }),
          tableCell('Terdapat hubungan positif signifikan antara kerja paruh waktu dan prestasi akademik (r=0,330; sig=0,013); mahasiswa bekerja <8 jam/hari dapat mempertahankan IPK memuaskan', { width: 2160 }),
          tableCell('Persamaan: variabel kerja paruh waktu dan prestasi akademik. Perbedaan: tidak ada variabel manajemen waktu dan lingkungan kampus', { width: 1440 }),
        ]
      }),
    ]
  }),
  new Paragraph({
    spacing: { line: 480, before: 80, after: 80 },
    children: [new TextRun({ text: 'Sumber: Diolah oleh peneliti dari berbagai sumber, 2025', italics: true, size: 20, font: 'Times New Roman' })]
  }),
  emptyLine(),

  // 2.3 Pengaruh Antar Variabel
  heading2('2.3 Pengaruh Antara Variabel'),

  heading3('2.3.1 Pengaruh Manajemen Waktu terhadap Prestasi Akademik'),
  p('Manajemen waktu merupakan salah satu faktor internal yang memiliki peran krusial dalam menentukan prestasi akademik mahasiswa. Mahasiswa yang mampu mengelola waktu dengan efektif cenderung memiliki lebih banyak waktu berkualitas untuk belajar, mengerjakan tugas, dan mempersiapkan diri menghadapi ujian, yang pada akhirnya berkontribusi pada perolehan IPK yang lebih tinggi.', { indent: true }),
  emptyLine(),
  p('Kamilatunnisa et al. (2025) dalam penelitiannya menemukan bahwa perilaku manajemen waktu berpengaruh positif dan signifikan terhadap prestasi akademik mahasiswa dengan nilai R\u00b2 sebesar 0,505, yang berarti manajemen waktu menjelaskan 50,5% variasi prestasi akademik. Hasil serupa ditemukan oleh Inayah et al. (2023) yang menemukan pengaruh positif signifikan manajemen waktu terhadap prestasi akademik mahasiswa yang bekerja (p = 0,005). Andis Tira et al. (2025) juga menemukan pengaruh parsial signifikan manajemen waktu terhadap prestasi akademik mahasiswa (p = 0,003; t = 3,114). Dalam kerangka SDT, mahasiswa dengan manajemen waktu yang baik menunjukkan kompetensi dan otonomi yang tinggi dalam mengelola aktivitas belajar mereka, yang mendukung motivasi intrinsik dan pada akhirnya meningkatkan prestasi akademik.', { indent: true }),
  emptyLine(),
  p('Berdasarkan uraian tersebut, dapat disimpulkan bahwa manajemen waktu berpengaruh positif terhadap prestasi akademik mahasiswa. Semakin baik kemampuan manajemen waktu yang dimiliki mahasiswa, semakin tinggi pula prestasi akademik yang dapat mereka capai.', { indent: true }),
  emptyLine(),

  heading3('2.3.2 Pengaruh Kerja Paruh Waktu terhadap Prestasi Akademik'),
  p('Pengaruh kerja paruh waktu terhadap prestasi akademik merupakan topik yang masih diperdebatkan dalam literatur akademik. Terdapat dua pandangan utama: pandangan yang menyatakan bahwa kerja paruh waktu berdampak negatif terhadap prestasi akademik karena mengurangi waktu dan energi yang tersedia untuk belajar; serta pandangan yang menyatakan bahwa kerja paruh waktu justru dapat berdampak positif melalui peningkatan motivasi, disiplin, dan rasa tanggung jawab.', { indent: true }),
  emptyLine(),
  p('Agustina dan Mardalis (2024) menemukan pengaruh positif dan signifikan kerja paruh waktu terhadap prestasi akademik mahasiswa (original sample 0,378; t-statistik 4,637; p = 0,000). Linggasari dan Kurniawan (2019) juga menemukan hubungan positif signifikan antara kerja paruh waktu dan prestasi akademik (r = 0,330; sig = 0,013), dengan catatan bahwa mahasiswa yang bekerja kurang dari 8 jam per hari masih dapat mempertahankan IPK yang memuaskan. Di sisi lain, Huda et al. (2023) menemukan bahwa kerja paruh waktu berpengaruh negatif signifikan terhadap prestasi akademik mahasiswa, terutama karena berkurangnya waktu untuk mengerjakan tugas dan mempersiapkan ujian.', { indent: true }),
  emptyLine(),
  p('Penelitian ini memandang bahwa pengaruh kerja paruh waktu terhadap prestasi akademik bersifat kondisional dan bergantung pada faktor-faktor lain seperti jumlah jam kerja, jenis pekerjaan, dan kemampuan manajemen waktu mahasiswa. Dalam konteks Prodi Manajemen FEB UNIMAL, di mana banyak mahasiswa berasal dari latar belakang sosial-ekonomi menengah ke bawah, kerja paruh waktu merupakan fenomena yang perlu dikaji secara komprehensif terhadap dampaknya pada prestasi akademik.', { indent: true }),
  emptyLine(),

  heading3('2.3.3 Pengaruh Lingkungan Kampus terhadap Prestasi Akademik'),
  p('Lingkungan kampus merupakan faktor eksternal yang memiliki pengaruh signifikan terhadap prestasi akademik mahasiswa. Kampus yang memiliki fasilitas lengkap, dosen berkualitas, iklim akademik yang kondusif, dan sistem dukungan yang baik menciptakan ekosistem belajar yang memungkinkan mahasiswa mencapai potensi akademik terbaik mereka.', { indent: true }),
  emptyLine(),
  p('Robyansyah et al. (2022) menemukan pengaruh positif dan signifikan lingkungan belajar terhadap prestasi belajar dengan koefisien jalur 0,498 (sig. = 0,000), sementara model secara keseluruhan menjelaskan 82,9% variasi prestasi. Dewi et al. (2025) juga menemukan pengaruh positif signifikan lingkungan kampus terhadap prestasi mahasiswa Prodi Manajemen (koef. = 0,171; t = 2,824; p = 0,005), dengan total model menjelaskan 78% variasi prestasi. Berhanu dan Sewagegn (2024) dari perspektif internasional menemukan bahwa persepsi iklim kampus berhubungan positif dengan prestasi akademik yang dimediasi oleh keterlibatan mahasiswa (student engagement). Munira et al. (2024) juga menemukan bahwa fasilitas belajar, kualitas pengajaran, dan suasana akademik berpengaruh signifikan terhadap prestasi akademik mahasiswa.', { indent: true }),
  emptyLine(),
  p('Dalam kerangka SDT, lingkungan kampus yang mendukung berperan dalam memenuhi kebutuhan psikologis mahasiswa, terutama kebutuhan kompetensi (melalui kualitas pengajaran), otonomi (melalui metode pembelajaran yang memberdayakan), dan keterkaitan (melalui interaksi sosial yang positif di kampus). Pemenuhan kebutuhan psikologis ini mendorong motivasi otonom yang pada gilirannya meningkatkan prestasi akademik.', { indent: true }),
  emptyLine(),

  // 2.4 Kerangka Konseptual
  heading2('2.4 Kerangka Konseptual'),
  p('Berdasarkan landasan teori dan penelitian terdahulu yang telah diuraikan, kerangka konseptual penelitian ini menggambarkan hubungan antara variabel bebas (manajemen waktu, kerja paruh waktu, dan lingkungan kampus) dengan variabel terikat (prestasi akademik mahasiswa). Penelitian ini menggunakan Self-Determination Theory sebagai grand theory yang menjadi payung teoritis dalam menjelaskan mekanisme pengaruh ketiga variabel bebas terhadap prestasi akademik.', { indent: true }),
  emptyLine(),
  p('Kerangka konseptual penelitian ini dapat digambarkan sebagai berikut:', { indent: true }),
  emptyLine(),
  pBoldCenter('[GAMBAR 2.1 KERANGKA KONSEPTUAL PENELITIAN]'),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 480 },
    children: [new TextRun({ text: 'Keterangan: Prodi Manajemen FEB Universitas Malikussaleh menunjukkan hubungan pengaruh dari variabel X1 (Manajemen Waktu), X2 (Kerja Paruh Waktu), dan X3 (Lingkungan Kampus) terhadap Y (Prestasi Akademik) baik secara parsial maupun simultan.', italics: true, size: 22, font: 'Times New Roman' })]
  }),
  emptyLine(),
  p('Dalam kerangka konseptual ini, manajemen waktu (X1), kerja paruh waktu (X2), dan lingkungan kampus (X3) secara bersama-sama maupun sendiri-sendiri diduga berpengaruh terhadap prestasi akademik (Y) mahasiswa Prodi Manajemen FEB UNIMAL. Manajemen waktu (X1) yang baik memungkinkan mahasiswa mengalokasikan waktu yang cukup dan berkualitas untuk belajar sehingga berdampak positif pada prestasi akademik. Kerja paruh waktu (X2) yang dikontrol dengan baik dapat memberikan manfaat finansial dan pengalaman yang mendukung motivasi belajar, namun kerja yang berlebihan dapat mengurangi waktu belajar sehingga berdampak negatif pada prestasi. Lingkungan kampus (X3) yang kondusif menyediakan fasilitas, pengajaran berkualitas, dan iklim akademik yang mendorong mahasiswa untuk berprestasi lebih baik.', { indent: true }),
  emptyLine(),

  // 2.5 Hipotesis Penelitian
  heading2('2.5 Hipotesis Penelitian'),
  p('Berdasarkan landasan teori, penelitian terdahulu, pengaruh antar variabel, dan kerangka konseptual yang telah diuraikan, hipotesis penelitian ini dirumuskan sebagai berikut:', { indent: true }),
  emptyLine(),
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 480 },
    children: [
      new TextRun({ text: 'H1: ', bold: true, size: 24, font: 'Times New Roman' }),
      new TextRun({ text: 'Manajemen waktu berpengaruh positif dan signifikan terhadap prestasi akademik mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh.', size: 24, font: 'Times New Roman' })
    ]
  }),
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 480 },
    children: [
      new TextRun({ text: 'H2: ', bold: true, size: 24, font: 'Times New Roman' }),
      new TextRun({ text: 'Kerja paruh waktu berpengaruh signifikan terhadap prestasi akademik mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh.', size: 24, font: 'Times New Roman' })
    ]
  }),
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 480 },
    children: [
      new TextRun({ text: 'H3: ', bold: true, size: 24, font: 'Times New Roman' }),
      new TextRun({ text: 'Lingkungan kampus berpengaruh positif dan signifikan terhadap prestasi akademik mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh.', size: 24, font: 'Times New Roman' })
    ]
  }),
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 480 },
    children: [
      new TextRun({ text: 'H4: ', bold: true, size: 24, font: 'Times New Roman' }),
      new TextRun({ text: 'Manajemen waktu, kerja paruh waktu, dan lingkungan kampus secara simultan berpengaruh signifikan terhadap prestasi akademik mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh.', size: 24, font: 'Times New Roman' })
    ]
  }),
  emptyLine(),

  pageBreak(),

  // ==================== BAB III ====================
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 480, before: 0, after: 120 },
    children: [new TextRun({ text: 'BAB III', bold: true, size: 24, font: 'Times New Roman' })]
  }),
  heading1('METODE PENELITIAN'),
  emptyLine(),

  heading2('3.1 Objek dan Lokasi Penelitian'),
  p('Penelitian ini dilakukan di Program Studi Manajemen, Fakultas Ekonomi dan Bisnis, Universitas Malikussaleh yang berlokasi di Kampus Bukit Indah, Blang Pulo, Kecamatan Muara Satu, Kota Lhokseumawe, Provinsi Aceh. Pemilihan lokasi ini didasarkan pada beberapa pertimbangan, yaitu: (1) Prodi Manajemen FEB UNIMAL merupakan salah satu program studi dengan jumlah mahasiswa terbanyak di Universitas Malikussaleh, sehingga memungkinkan diperolehnya sampel yang representatif; (2) Karakteristik sosial-ekonomi mahasiswa Prodi Manajemen FEB UNIMAL yang beragam, termasuk banyaknya mahasiswa yang bekerja paruh waktu, menjadikannya objek penelitian yang relevan; dan (3) Peneliti merupakan mahasiswa aktif di program studi yang sama, sehingga memiliki kemudahan akses terhadap subjek penelitian.', { indent: true }),
  emptyLine(),
  p('Objek penelitian dalam penelitian ini adalah pengaruh manajemen waktu (X1), kerja paruh waktu (X2), dan lingkungan kampus (X3) terhadap prestasi akademik (Y) mahasiswa aktif Program Studi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh.', { indent: true }),
  emptyLine(),

  heading2('3.2 Jenis Penelitian'),
  p('Penelitian ini menggunakan pendekatan penelitian kuantitatif, yaitu metode penelitian yang berlandaskan pada filsafat positivisme, digunakan untuk meneliti pada populasi atau sampel tertentu, pengumpulan data menggunakan instrumen penelitian, analisis data bersifat kuantitatif/statistik, dengan tujuan untuk menguji hipotesis yang telah ditetapkan (Sugiyono, 2023). Jenis penelitian ini adalah penelitian asosiatif kausal yang bertujuan untuk mengetahui pengaruh antara dua variabel atau lebih, yaitu untuk mengetahui pengaruh variabel bebas (manajemen waktu, kerja paruh waktu, dan lingkungan kampus) terhadap variabel terikat (prestasi akademik mahasiswa).', { indent: true }),
  emptyLine(),

  heading2('3.3 Populasi dan Sampel'),

  heading3('3.3.1 Populasi'),
  p('Populasi dalam penelitian ini adalah seluruh mahasiswa aktif Program Studi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh dari angkatan 2021 hingga angkatan 2023 yang masih aktif berkuliah dan telah menempuh minimal 4 semester. Batasan angkatan ini digunakan karena mahasiswa angkatan tersebut telah memiliki pengalaman perkuliahan yang cukup untuk memberikan penilaian yang akurat terhadap variabel lingkungan kampus dan prestasi akademik (IPK) yang sudah terbentuk. Berdasarkan data yang diperoleh dari program studi, total populasi mahasiswa aktif Prodi Manajemen FEB UNIMAL angkatan 2021 s.d. 2023 diperkirakan berjumlah sekitar 500 mahasiswa.', { indent: true }),
  emptyLine(),

  heading3('3.3.2 Sampel dan Teknik Pengambilan Sampel'),
  p('Penentuan ukuran sampel dalam penelitian ini menggunakan rumus Slovin, yaitu:', { indent: true }),
  emptyLine(),
  pBoldCenter('n = N / (1 + N \u00b7 e\u00b2)'),
  emptyLine(),
  p('Di mana:', { indent: true }),
  p('n  = ukuran sampel', { indent: true }),
  p('N  = ukuran populasi (500 mahasiswa)', { indent: true }),
  p('e  = tingkat kesalahan yang ditoleransi (10% atau 0,10)', { indent: true }),
  emptyLine(),
  p('Dengan menggunakan rumus di atas, maka ukuran sampel minimum yang diperlukan adalah:', { indent: true }),
  pBoldCenter('n = 500 / (1 + 500 \u00b7 (0,10)\u00b2) = 500 / (1 + 5) = 500 / 6 \u2248 84 responden'),
  emptyLine(),
  p('Untuk meningkatkan representativitas dan mengantisipasi data yang tidak lengkap atau tidak valid, peneliti menetapkan jumlah sampel sebesar 100 responden. Jumlah ini juga relevan dengan ketentuan analisis regresi berganda yang memerlukan minimal 5-10 responden per variabel bebas (Hair et al., 2014). Dengan tiga variabel bebas, maka diperlukan minimal 30 responden, dan penggunaan 100 responden memberikan tingkat representativitas yang lebih baik.', { indent: true }),
  emptyLine(),
  p('Teknik pengambilan sampel yang digunakan adalah purposive sampling, yaitu teknik penentuan sampel dengan pertimbangan tertentu (Sugiyono, 2023). Kriteria pemilihan sampel dalam penelitian ini adalah:', { indent: true }),
  numbered('Mahasiswa aktif Program Studi Manajemen FEB UNIMAL angkatan 2021, 2022, atau 2023.', '1'),
  numbered('Telah menempuh minimal 4 semester perkuliahan (agar telah memiliki IPK yang representatif dan pengalaman perkuliahan yang cukup).', '2'),
  numbered('Bersedia menjadi responden dan mengisi kuesioner penelitian secara lengkap dan jujur.', '3'),
  emptyLine(),

  heading2('3.4 Teknik Pengumpulan Data'),
  p('Data yang digunakan dalam penelitian ini terdiri dari data primer dan data sekunder:', { indent: true }),
  emptyLine(),
  heading3('3.4.1 Data Primer'),
  p('Data primer diperoleh langsung dari responden melalui teknik survei menggunakan kuesioner. Kuesioner dalam penelitian ini menggunakan skala Likert dengan lima tingkatan jawaban, yaitu: (1) Sangat Tidak Setuju (STS), (2) Tidak Setuju (TS), (3) Netral (N), (4) Setuju (S), dan (5) Sangat Setuju (SS). Kuesioner akan didistribusikan secara langsung (tatap muka) kepada responden yang memenuhi kriteria sampel yang telah ditetapkan, serta dapat didistribusikan secara daring melalui Google Form untuk memperluas jangkauan dan efisiensi pengumpulan data.', { indent: true }),
  emptyLine(),
  heading3('3.4.2 Data Sekunder'),
  p('Data sekunder diperoleh dari sumber-sumber yang sudah ada, antara lain dokumen akademik dari Program Studi Manajemen FEB UNIMAL, pangkalan data PDDikti, laporan dan publikasi ilmiah yang relevan, serta sumber-sumber lainnya yang dapat mendukung penelitian ini.', { indent: true }),
  emptyLine(),

  heading2('3.5 Definisi Operasional Variabel'),
  p('Definisi operasional variabel dalam penelitian ini disajikan pada tabel berikut:', { indent: true }),
  emptyLine(),

  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [540, 1080, 1440, 3000, 3300],
    rows: [
      new TableRow({
        children: [
          tableCell('No.', { width: 540, header: true, center: true }),
          tableCell('Variabel', { width: 1080, header: true }),
          tableCell('Definisi Operasional', { width: 1440, header: true }),
          tableCell('Indikator', { width: 3000, header: true }),
          tableCell('Skala Pengukuran', { width: 3300, header: true }),
        ]
      }),
      new TableRow({
        children: [
          tableCell('1', { width: 540, center: true }),
          tableCell('Manajemen Waktu (X1)', { width: 1080 }),
          tableCell('Kemampuan mahasiswa dalam merencanakan, mengorganisasikan, dan menggunakan waktu belajar secara efektif', { width: 1440 }),
          tableCell('1. Menetapkan tujuan dan prioritas\n2. Membuat daftar tugas\n3. Penjadwalan kegiatan belajar\n4. Menghindari penundaan\n5. Pengorganisasian diri\n(Kamilatunnisa et al., 2025; Andis Tira et al., 2025)', { width: 3000 }),
          tableCell('Skala Likert 1-5 (1=Sangat Tidak Setuju, 5=Sangat Setuju)', { width: 3300 }),
        ]
      }),
      new TableRow({
        children: [
          tableCell('2', { width: 540, center: true }),
          tableCell('Kerja Paruh Waktu (X2)', { width: 1080 }),
          tableCell('Aktivitas bekerja yang dilakukan mahasiswa di luar jam kuliah untuk memperoleh penghasilan tambahan', { width: 1440 }),
          tableCell('1. Jumlah jam kerja per minggu\n2. Jenis pekerjaan dan fleksibilitas\n3. Pengaruh pekerjaan terhadap jadwal kuliah\n4. Motivasi bekerja\n5. Manfaat finansial dari bekerja\n(Huda et al., 2023; Agustina & Mardalis, 2024)', { width: 3000 }),
          tableCell('Skala Likert 1-5 (1=Sangat Tidak Setuju, 5=Sangat Setuju)', { width: 3300 }),
        ]
      }),
      new TableRow({
        children: [
          tableCell('3', { width: 540, center: true }),
          tableCell('Lingkungan Kampus (X3)', { width: 1080 }),
          tableCell('Persepsi mahasiswa terhadap keseluruhan kondisi fisik, akademik, dan sosial kampus yang memengaruhi proses belajar', { width: 1440 }),
          tableCell('1. Kenyamanan dan kelengkapan fasilitas fisik\n2. Kualitas proses pembelajaran\n3. Suasana dan iklim akademik\n4. Interaksi sosial akademik\n5. Dukungan administrasi dan layanan kampus\n(Robyansyah et al., 2022; Dewi et al., 2025; Munira et al., 2024)', { width: 3000 }),
          tableCell('Skala Likert 1-5 (1=Sangat Tidak Setuju, 5=Sangat Setuju)', { width: 3300 }),
        ]
      }),
      new TableRow({
        children: [
          tableCell('4', { width: 540, center: true }),
          tableCell('Prestasi Akademik (Y)', { width: 1080 }),
          tableCell('Hasil belajar mahasiswa yang diwujudkan dalam IPK dan berbagai aspek keberhasilan akademik lainnya', { width: 1440 }),
          tableCell('1. Indeks Prestasi Kumulatif (IPK)\n2. Ketepatan penyelesaian tugas\n3. Keaktifan dalam perkuliahan\n4. Tingkat kehadiran\n(Agustina & Mardalis, 2024; Dewi et al., 2025; Inayah et al., 2023)', { width: 3000 }),
          tableCell('Skala Likert 1-5 untuk aspek non-IPK; IPK aktual untuk item IPK', { width: 3300 }),
        ]
      }),
    ]
  }),
  new Paragraph({
    spacing: { line: 480, before: 80, after: 80 },
    children: [new TextRun({ text: 'Sumber: Diolah oleh peneliti dari berbagai sumber, 2025', italics: true, size: 20, font: 'Times New Roman' })]
  }),
  emptyLine(),

  heading2('3.6 Uji Instrumen Penelitian'),

  heading3('3.6.1 Uji Validitas'),
  p('Uji validitas dilakukan untuk mengukur sah atau valid tidaknya suatu kuesioner (Ghozali, 2018). Suatu kuesioner dikatakan valid jika pernyataan pada kuesioner mampu mengungkapkan sesuatu yang akan diukur oleh kuesioner tersebut. Dalam penelitian ini, uji validitas menggunakan teknik korelasi Pearson Product Moment. Butir pernyataan dikatakan valid apabila nilai r-hitung lebih besar dari r-tabel pada tingkat signifikansi 5% (r-hitung > r-tabel). Pengujian validitas dilakukan dengan bantuan program SPSS.', { indent: true }),
  emptyLine(),

  heading3('3.6.2 Uji Reliabilitas'),
  p('Uji reliabilitas digunakan untuk mengukur konsistensi atau keandalan suatu instrumen pengukuran. Instrumen yang reliabel adalah instrumen yang bila digunakan beberapa kali untuk mengukur obyek yang sama akan menghasilkan data yang sama (Sugiyono, 2023). Dalam penelitian ini, uji reliabilitas menggunakan metode Cronbach\'s Alpha. Suatu instrumen dinyatakan reliabel apabila nilai Cronbach\'s Alpha \u2265 0,60 (Ghozali, 2018).', { indent: true }),
  emptyLine(),

  heading2('3.7 Uji Asumsi Klasik'),

  heading3('3.7.1 Uji Normalitas'),
  p('Uji normalitas bertujuan untuk menguji apakah dalam model regresi, variabel pengganggu atau residual memiliki distribusi normal (Ghozali, 2018). Model regresi yang baik adalah yang memiliki distribusi data normal atau mendekati normal. Dalam penelitian ini, uji normalitas dilakukan menggunakan uji Kolmogorov-Smirnov. Data dikatakan berdistribusi normal apabila nilai signifikansi (Asymp. Sig. 2-tailed) > 0,05.', { indent: true }),
  emptyLine(),

  heading3('3.7.2 Uji Multikolinearitas'),
  p('Uji multikolinearitas bertujuan untuk menguji apakah model regresi ditemukan adanya korelasi antar variabel bebas (independen) (Ghozali, 2018). Model regresi yang baik seharusnya tidak terjadi korelasi di antara variabel independen. Deteksi ada atau tidaknya multikolinearitas di dalam model regresi dapat dilihat dari nilai Tolerance dan Variance Inflation Factor (VIF). Jika nilai Tolerance > 0,10 atau nilai VIF < 10, maka tidak terjadi multikolinearitas.', { indent: true }),
  emptyLine(),

  heading3('3.7.3 Uji Heteroskedastisitas'),
  p('Uji heteroskedastisitas bertujuan menguji apakah dalam model regresi terjadi ketidaksamaan variance dari residual satu pengamatan ke pengamatan yang lain (Ghozali, 2018). Jika variance dari residual satu pengamatan ke pengamatan lain tetap, maka disebut homoskedastisitas dan jika berbeda disebut heteroskedastisitas. Model regresi yang baik adalah yang homoskedastisitas atau tidak terjadi heteroskedastisitas. Uji heteroskedastisitas dalam penelitian ini menggunakan uji Glejser. Apabila nilai signifikansi (Sig.) dari masing-masing variabel bebas > 0,05, maka tidak terjadi gejala heteroskedastisitas.', { indent: true }),
  emptyLine(),

  heading2('3.8 Metode Analisis Data'),

  heading3('3.8.1 Analisis Statistik Deskriptif'),
  p('Analisis statistik deskriptif digunakan untuk memberikan gambaran atau deskripsi mengenai data yang diperoleh dari responden, yang dilihat dari nilai rata-rata (mean), standar deviasi, nilai minimum, dan nilai maksimum dari setiap variabel penelitian. Analisis ini dilakukan dengan menggunakan program SPSS.', { indent: true }),
  emptyLine(),

  heading3('3.8.2 Analisis Regresi Linear Berganda'),
  p('Metode analisis utama yang digunakan dalam penelitian ini adalah analisis regresi linear berganda (multiple linear regression). Analisis regresi linear berganda digunakan untuk mengetahui pengaruh dua atau lebih variabel bebas terhadap satu variabel terikat (Sugiyono, 2023). Adapun model persamaan regresi linear berganda dalam penelitian ini adalah:', { indent: true }),
  emptyLine(),
  pBoldCenter('Y = a + b1X1 + b2X2 + b3X3 + e'),
  emptyLine(),
  p('Keterangan:', { indent: true }),
  p('Y  = Prestasi Akademik Mahasiswa', { indent: true }),
  p('a   = Konstanta', { indent: true }),
  p('b1, b2, b3  = Koefisien regresi masing-masing variabel', { indent: true }),
  p('X1 = Manajemen Waktu', { indent: true }),
  p('X2 = Kerja Paruh Waktu', { indent: true }),
  p('X3 = Lingkungan Kampus', { indent: true }),
  p('e   = Error term (variabel pengganggu)', { indent: true }),
  emptyLine(),

  heading2('3.9 Pengujian Hipotesis'),

  heading3('3.9.1 Uji Parsial (Uji t)'),
  p('Uji t atau uji parsial digunakan untuk mengetahui apakah masing-masing variabel bebas secara individual (parsial) mempunyai pengaruh yang signifikan terhadap variabel terikat (Ghozali, 2018). Kriteria pengujian:', { indent: true }),
  numbered('Jika nilai t-hitung > t-tabel atau nilai Sig. < 0,05, maka variabel bebas secara parsial berpengaruh signifikan terhadap variabel terikat, sehingga H\u2080 ditolak dan H\u2090 diterima.', 'a'),
  numbered('Jika nilai t-hitung \u2264 t-tabel atau nilai Sig. \u2265 0,05, maka variabel bebas secara parsial tidak berpengaruh signifikan terhadap variabel terikat, sehingga H\u2080 diterima dan H\u2090 ditolak.', 'b'),
  emptyLine(),

  heading3('3.9.2 Uji Simultan (Uji F)'),
  p('Uji F atau uji simultan digunakan untuk mengetahui apakah seluruh variabel bebas secara bersama-sama (simultan) mempunyai pengaruh yang signifikan terhadap variabel terikat (Ghozali, 2018). Kriteria pengujian:', { indent: true }),
  numbered('Jika nilai F-hitung > F-tabel atau nilai Sig. < 0,05, maka seluruh variabel bebas secara simultan berpengaruh signifikan terhadap variabel terikat, sehingga H\u2080 ditolak dan H\u2090 diterima.', 'a'),
  numbered('Jika nilai F-hitung \u2264 F-tabel atau nilai Sig. \u2265 0,05, maka seluruh variabel bebas secara simultan tidak berpengaruh signifikan terhadap variabel terikat, sehingga H\u2080 diterima dan H\u2090 ditolak.', 'b'),
  emptyLine(),

  heading3('3.9.3 Koefisien Determinasi (R\u00b2)'),
  p('Koefisien determinasi (R\u00b2) digunakan untuk mengukur seberapa jauh kemampuan model dalam menerangkan variasi variabel terikat (Ghozali, 2018). Nilai R\u00b2 berkisar antara 0 dan 1. Nilai yang kecil berarti kemampuan variabel-variabel bebas dalam menjelaskan variasi variabel terikat sangat terbatas. Nilai yang mendekati 1 berarti variabel-variabel bebas memberikan hampir semua informasi yang dibutuhkan untuk memprediksi variasi variabel terikat. Dalam penelitian ini digunakan nilai Adjusted R\u00b2 untuk menghindari bias pada kenaikan nilai R\u00b2 akibat penambahan variabel bebas yang tidak signifikan.', { indent: true }),
  emptyLine(),

  pageBreak(),

  // ==================== DAFTAR PUSTAKA ====================
  heading1('DAFTAR PUSTAKA'),
  emptyLine(),

  p('Agustina, A., & Mardalis, A. (2024). Pengaruh kerja paruh waktu, motivasi belajar dan time management terhadap prestasi akademik (studi kasus mahasiswa bekerja part-time). Jurnal Bisnis dan Ekonomi (JBBE), 17(2). https://doi.org/10.46306/jbbe.v17i2.556', { indent: false }),
  emptyLine(),
  p('Andis Tira, S., Eka, P., & Hermansyah, D. (2025). Pengaruh manajemen waktu, efikasi diri, dan lingkungan kampus terhadap prestasi akademik mahasiswa Teknik Mesin Universitas Riau. Jurnal Bisnis dan Manajemen. https://doi.org/10.60036/jbm.689', { indent: false }),
  emptyLine(),
  p('Berhanu, K. Z., & Sewagegn, A. A. (2024). The role of perceived campus climate in students\u2019 academic achievements as mediated by students\u2019 engagement in higher education institutions. Cogent Education, 11(1). https://doi.org/10.1080/2331186X.2024.2377839', { indent: false }),
  emptyLine(),
  p('Bureau, J. S., Howard, J. L., Chong, J. X. Y., & Guay, F. (2022). Pathways to student motivation: A meta-analysis of antecedents of autonomous and controlled motivations. Review of Educational Research, 92(1), 46\u201388. https://doi.org/10.3102/00346543211042426', { indent: false }),
  emptyLine(),
  p('Deci, E. L., & Ryan, R. M. (2000). The \u201cwhat\u201d and \u201cwhy\u201d of goal pursuits: Human needs and the self-determination of behavior. Psychological Inquiry, 11(4), 227\u2013268. https://doi.org/10.1207/S15327965PLI1104_01', { indent: false }),
  emptyLine(),
  p('Dewi, A. F. D., Kholidy, M. A. N., & Prajitiasari, E. D. (2025). Efikasi diri sebagai mediasi dalam hubungan lingkungan kampus dan manajemen waktu pada prestasi mahasiswa. Jurnal Pariwisata, Bisnis Digital dan Manajemen, 4(1). https://doi.org/10.33480/jasdim.v4i1.6666', { indent: false }),
  emptyLine(),
  p('Ghozali, I. (2018). Aplikasi analisis multivariate dengan program IBM SPSS 25 (edisi 9). Badan Penerbit Universitas Diponegoro.', { indent: false }),
  emptyLine(),
  p('Hamalik, O. (2016). Proses belajar mengajar. Bumi Aksara.', { indent: false }),
  emptyLine(),
  p('Hair, J. F., Black, W. C., Babin, B. J., & Anderson, R. E. (2014). Multivariate data analysis (7th ed.). Pearson Education.', { indent: false }),
  emptyLine(),
  p('Howard, J. L., Bureau, J., Guay, F., Chong, J. X. Y., & Ryan, R. M. (2021). Student motivation and associated outcomes: A meta-analysis from self-determination theory. Perspectives on Psychological Science, 16(6), 1300\u20131323. https://doi.org/10.1177/1745691620966789', { indent: false }),
  emptyLine(),
  p('Huda, M. A. A., Fani, M., Saragih, R. M., & Lestari, D. (2023). Pengaruh kerja paruh waktu terhadap prestasi akademik mahasiswa UIN Sumatera Utara. Madani Jurnal Ilmiah Multidisiplin, 1(6), 447\u2013456. https://doi.org/10.5281/zenodo.8127903', { indent: false }),
  emptyLine(),
  p('Inayah, D. N., Daud, M., & Nur, H. (2023). Pengaruh manajemen waktu terhadap prestasi akademik mahasiswa yang bekerja di Kota Makassar. Jurnal Penelitian, Pendidikan, dan Pengajaran (PESHUM), 2(2). https://doi.org/10.56799/peshum.v2i2.1391', { indent: false }),
  emptyLine(),
  p('Kamilatunnisa, K., Nurhayati, N., & Rahayu, S. (2025). Analisis pengaruh perilaku manajemen waktu terhadap prestasi akademik mahasiswa. Katalis: Jurnal Pendidikan Ekonomi dan Kewirausahaan, 2(3). https://doi.org/10.62383/katalis.v2i3.2072', { indent: false }),
  emptyLine(),
  p('Linggasari, L. Y., & Kurniawan, R. Y. (2019). Hubungan kerja paruh waktu dengan prestasi akademik mahasiswa Jurusan Pendidikan Ekonomi Universitas Negeri Surabaya angkatan 2015. JUPE: Jurnal Pendidikan Ekonomi, 7(3), 92\u201393.', { indent: false }),
  emptyLine(),
  p('Macan, T. H. (1994). Time management: Test of a process model. Journal of Applied Psychology, 79(3), 381\u2013391. https://doi.org/10.1037/0021-9010.79.3.381', { indent: false }),
  emptyLine(),
  p('Munira, R., Fonna, T., Nadia, S., & Marsitah, I. (2024). Pengaruh lingkungan belajar terhadap prestasi akademik mahasiswa di Universitas Almuslim. Jurnal Pendidikan Guru Sekolah Dasar, 1(4). https://doi.org/10.47134/pgsd.v1i4.770', { indent: false }),
  emptyLine(),
  p('Nurrahmaniah. (2019). Peningkatan prestasi akademik melalui manajemen waktu dan minat belajar. Andragogi: Jurnal Diklat Teknis Pendidikan dan Keagamaan, 1(1). https://doi.org/10.36671/andragogi.v1i1.52', { indent: false }),
  emptyLine(),
  p('Robyansyah, Indarti, S., & Widayatsari, A. (2022). Pengaruh lingkungan belajar dan disiplin terhadap motivasi belajar dan prestasi belajar taruna Politeknik Negeri Bengkalis Jurusan Kemaritiman. Jurnal Daya Saing, 8(3). https://jurnal.polbeng.ac.id/index.php/JDS', { indent: false }),
  emptyLine(),
  p('Ryan, R. M., & Deci, E. L. (2000). Self-determination theory and the facilitation of intrinsic motivation, social development, and well-being. American Psychologist, 55(1), 68\u201378. https://doi.org/10.1037/0003-066X.55.1.68', { indent: false }),
  emptyLine(),
  p('Ryan, R. M., & Deci, E. L. (2017). Self-determination theory: Basic psychological needs in motivation, development, and wellness. Guilford Press.', { indent: false }),
  emptyLine(),
  p('Sugiyono. (2023). Metode penelitian kuantitatif, kualitatif dan R&D (edisi terbaru). Alfabeta.', { indent: false }),
  emptyLine(),
  p('Sudjana, N. (2014). Penilaian hasil proses belajar mengajar. Remaja Rosdakarya.', { indent: false }),
  emptyLine(),
  p('Zimmerman, B. J. (2002). Becoming a self-regulated learner: An overview. Theory Into Practice, 41(2), 64\u201370. https://doi.org/10.1207/s15430421tip4102_2', { indent: false }),

];

// ====================== BUILD DOCUMENT ======================
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: 'Times New Roman', size: 24 }
      }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 2268, right: 1701, bottom: 1701, left: 2835 } // 4cm left, 3cm others
      }
    },
    children: children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/mnt/user-data/outputs/Skripsi_Manajemen_Waktu_Kerja_Paruh_Waktu_Lingkungan_Kampus.docx', buffer);
  console.log('Skripsi berhasil dibuat!');
}).catch(err => {
  console.error('Error:', err);
});
