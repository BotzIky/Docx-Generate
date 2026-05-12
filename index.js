const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, LevelFormat, PageNumber, PageBreak, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// ===================== HELPER FUNCTIONS =====================

function makeHeading1(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 400, after: 200 },
    children: [
      new TextRun({ text, bold: true, size: 24, font: "Times New Roman", allCaps: true })
    ]
  });
}

function makeHeading2(text) {
  return new Paragraph({
    spacing: { before: 360, after: 160 },
    children: [
      new TextRun({ text, bold: true, size: 24, font: "Times New Roman" })
    ]
  });
}

function makeHeading3(text) {
  return new Paragraph({
    spacing: { before: 280, after: 120 },
    children: [
      new TextRun({ text, bold: true, italic: true, size: 24, font: "Times New Roman" })
    ]
  });
}

function makePara(text, options = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: 200, line: 480, lineRule: "auto" },
    indent: { firstLine: 720 },
    children: [
      new TextRun({ text, size: 24, font: "Times New Roman", ...options })
    ]
  });
}

function makeParaNoIndent(text, options = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: 200, line: 480, lineRule: "auto" },
    children: [
      new TextRun({ text, size: 24, font: "Times New Roman", ...options })
    ]
  });
}

function makeParaRuns(runs) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: 200, line: 480, lineRule: "auto" },
    indent: { firstLine: 720 },
    children: runs.map(r => new TextRun({ size: 24, font: "Times New Roman", ...r }))
  });
}

function makeNumberedItem(num, text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: 100, line: 480, lineRule: "auto" },
    indent: { left: 720, hanging: 360 },
    children: [
      new TextRun({ text: `${num}. ${text}`, size: 24, font: "Times New Roman" })
    ]
  });
}

function makeAlphaItem(letter, text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: 100, line: 480, lineRule: "auto" },
    indent: { left: 720, hanging: 360 },
    children: [
      new TextRun({ text: `${letter}. ${text}`, size: 24, font: "Times New Roman" })
    ]
  });
}

function makeTableCaption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [
      new TextRun({ text, bold: true, size: 24, font: "Times New Roman" })
    ]
  });
}

function makeItalicCaption(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 100, after: 200 },
    children: [
      new TextRun({ text, italic: true, size: 20, font: "Times New Roman" })
    ]
  });
}

function makeDaftarPustakaItem(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: 160, line: 480, lineRule: "auto" },
    indent: { left: 720, hanging: 720 },
    children: [
      new TextRun({ text, size: 24, font: "Times New Roman" })
    ]
  });
}

// ===================== TABLE HELPERS =====================

const borderThin = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
const borderNone = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const allBorders = { top: borderThin, bottom: borderThin, left: borderThin, right: borderThin };

function makeHeaderCell(text, width) {
  return new TableCell({
    borders: allBorders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: "D9D9D9", type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, size: 18, font: "Times New Roman" })]
    })]
  });
}

function makeBodyCell(text, width) {
  return new TableCell({
    borders: allBorders,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: [new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      children: [new TextRun({ text, size: 18, font: "Times New Roman" })]
    })]
  });
}

// ===================== KERANGKA KONSEPTUAL =====================

function makeKerangkaBox() {
  // Simple table-based diagram for kerangka konseptual
  const w = 8640; // total width
  const col1 = 2200;
  const col2 = 800;
  const col3 = 2200;

  function varBox(label, kode) {
    return new TableCell({
      borders: allBorders,
      width: { size: col1, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 80, right: 80 },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: label, bold: true, size: 20, font: "Times New Roman" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: `(${kode})`, size: 20, font: "Times New Roman" })]
        })
      ]
    });
  }

  function arrowCell() {
    return new TableCell({
      borders: { top: borderNone, bottom: borderNone, left: borderNone, right: borderNone },
      width: { size: col2, type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "→", size: 28, font: "Times New Roman" })]
      })]
    });
  }

  function depBox(rowspan) {
    return new TableCell({
      borders: allBorders,
      width: { size: col3, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 80, right: 80 },
      rowSpan: rowspan,
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Prestasi Akademik", bold: true, size: 20, font: "Times New Roman" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "(Y)", size: 20, font: "Times New Roman" })]
        })
      ]
    });
  }

  return new Table({
    width: { size: w, type: WidthType.DXA },
    columnWidths: [col1, col2, col3],
    rows: [
      new TableRow({
        children: [
          varBox("Manajemen Waktu", "X1"),
          arrowCell(),
          depBox(3)
        ]
      }),
      new TableRow({
        children: [
          varBox("Kerja Paruh Waktu", "X2"),
          arrowCell()
        ]
      }),
      new TableRow({
        children: [
          varBox("Lingkungan Kampus", "X3"),
          arrowCell()
        ]
      })
    ]
  });
}

// ===================== PENELITIAN TERDAHULU TABLE =====================

function makePenelitianTable() {
  const totalW = 9026;
  const cols = [300, 900, 1200, 700, 800, 1300, 1500, 1000, 1326];
  // No, Peneliti, Judul, Variabel, Metode, Sampel, Hasil, Kelemahan, Perbedaan

  const headers = [
    "No.", "Peneliti & Tahun", "Judul Penelitian",
    "Variabel", "Metode", "Sampel",
    "Hasil Penelitian", "Kelemahan Metodologis", "Perbedaan dengan Penelitian Ini"
  ];

  function hCell(t, w) {
    return new TableCell({
      borders: allBorders,
      width: { size: w, type: WidthType.DXA },
      shading: { fill: "D9D9D9", type: ShadingType.CLEAR },
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: t, bold: true, size: 16, font: "Times New Roman" })]
      })]
    });
  }

  function bCell(t, w) {
    return new TableCell({
      borders: allBorders,
      width: { size: w, type: WidthType.DXA },
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
      children: [new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        children: [new TextRun({ text: t, size: 16, font: "Times New Roman" })]
      })]
    });
  }

  const data = [
    [
      "1",
      "Alvionita et al. (2022)\nJurnal Oportunitas Vol.3 No.2\nISSN: E-2721-5954",
      "Pengaruh Kerja Part-Time dan Aktivitas Belajar terhadap Prestasi Akademik Mahasiswa",
      "X1: Kerja Part-Time\nX2: Aktivitas Belajar\nY: Prestasi Akademik",
      "Kuantitatif, regresi linear sederhana dan berganda",
      "50 mahasiswa yang bekerja part-time (Prodi Pendidikan Ekonomi Univ. PGRI Ronggolawe Tuban)",
      "Kerja part-time berpengaruh positif signifikan (koef=0,491; t=7,235; sig=0,000; R²=0,522). Aktivitas belajar berpengaruh positif signifikan (koef=0,798; t=4,866; sig=0,000). Secara simultan F=35,323; R²=0,600.",
      "Sampel kecil (50 responden), hanya satu prodi, tidak ada DOI resmi.",
      "Penelitian ini menambahkan variabel manajemen waktu dan lingkungan kampus, objek di FEB UNIMAL, sampel lebih besar."
    ],
    [
      "2",
      "Robyansyah et al. (2022)\nJurnal Daya Saing Vol.8 No.3\nDOI: 10.35446/dayasaing.v8i3.972",
      "Pengaruh Lingkungan Belajar dan Disiplin terhadap Motivasi Belajar dan Prestasi Belajar",
      "X1: Lingkungan Belajar\nX2: Disiplin\nZ: Motivasi Belajar\nY: Prestasi Belajar",
      "Kuantitatif, path analysis (SPSS v.21)",
      "104 taruna Politeknik Negeri Bengkalis Jur. Kemaritiman",
      "Lingkungan belajar berpengaruh positif signifikan terhadap prestasi (koef=0,498; t=6,051; sig=0,000). Disiplin juga signifikan (koef=0,262; t=3,098; sig=0,003). Model menjelaskan 82,9% variasi prestasi.",
      "Hanya satu institusi, variabel terbatas, bergantung kejujuran responden.",
      "Penelitian ini tidak menggunakan variabel disiplin dan motivasi, menambahkan manajemen waktu dan kerja paruh waktu, objek mahasiswa manajemen FEB UNIMAL."
    ],
    [
      "3",
      "Huda et al. (2023)\nMadani: Jurnal Ilmiah Multidisiplin Vol.1 No.6\nDOI: 10.5281/zenodo.8127903",
      "Pengaruh Kerja Paruh Waktu terhadap Prestasi Akademik Mahasiswa UIN SU",
      "X: Kerja Paruh Waktu\nY: Prestasi Akademik",
      "Kuantitatif, kuesioner Likert",
      "109 mahasiswa UIN Sumatera Utara",
      "Kerja paruh waktu berpengaruh negatif signifikan terhadap prestasi akademik. Skor terendah pada kemampuan menjaga kualitas tugas akademik.",
      "Hanya dua variabel, tidak mengontrol faktor lain seperti manajemen waktu dan lingkungan kampus.",
      "Penelitian ini menambahkan variabel manajemen waktu dan lingkungan kampus sebagai prediktor sekaligus, objek di FEB UNIMAL Lhokseumawe."
    ],
    [
      "4",
      "Agustina & Mardalis (2024)\nJurnal Bisnis & Ekonomi (JBBE) Vol.17 No.2\nDOI: 10.46306/jbbe.v17i2.556",
      "Pengaruh Kerja Paruh Waktu, Motivasi Belajar dan Time Management terhadap Prestasi Akademik",
      "X1: Kerja Paruh Waktu\nX2: Motivasi Belajar\nX3: Time Management\nY: Prestasi Akademik",
      "Kuantitatif, PLS-SEM (SmartPLS)",
      "250 mahasiswa yang bekerja part-time",
      "Kerja paruh waktu berpengaruh positif signifikan (OS=0,378; t=4,637; p=0,000). Time management berpengaruh signifikan. Motivasi belajar signifikan.",
      "Tidak mengkaji lingkungan kampus, menggunakan PLS-SEM yang membutuhkan keahlian teknis tinggi.",
      "Penelitian ini mengganti variabel motivasi belajar dengan lingkungan kampus, menggunakan regresi berganda, objek di FEB UNIMAL."
    ],
    [
      "5",
      "Andis Tira et al. (2025)\nJurnal Bisnis Mahasiswa\nDOI: 10.60036/jbm.689",
      "Pengaruh Manajemen Waktu, Efikasi Diri, dan Lingkungan Kampus terhadap Prestasi Akademik Mahasiswa Teknik Mesin",
      "X1: Manajemen Waktu\nX2: Efikasi Diri\nX3: Lingkungan Kampus\nY: Prestasi Akademik",
      "Kuantitatif, regresi berganda (SPSS v.24)",
      "90 mahasiswa Teknik Mesin Universitas Riau",
      "Manajemen waktu berpengaruh positif signifikan (p=0,003; t=3,114). Lingkungan kampus berpengaruh signifikan. R²adj=0,643.",
      "Tidak mengkaji kerja paruh waktu, objek mahasiswa teknik tidak dapat digeneralisasi ke mahasiswa manajemen.",
      "Penelitian ini mengganti variabel efikasi diri dengan kerja paruh waktu, objek di Prodi Manajemen FEB UNIMAL."
    ],
    [
      "6",
      "Dewi et al. (2025)\nJurnal Pariwisata, Bisnis Digital dan Manajemen Vol.4 No.1\nDOI: 10.33480/jasdim.v4i1.6666",
      "Efikasi Diri sebagai Mediasi dalam Hubungan Lingkungan Kampus dan Manajemen Waktu pada Prestasi Mahasiswa",
      "X1: Lingkungan Kampus\nX2: Manajemen Waktu\nZ: Efikasi Diri (mediasi)\nY: Prestasi Mahasiswa",
      "PLS-SEM (SmartPLS 4.0)",
      "270 mahasiswa S1 Manajemen FEB Universitas Jember",
      "Lingkungan kampus berpengaruh positif signifikan (koef=0,171; t=2,824; p=0,005). Manajemen waktu signifikan. Efikasi diri sebagai mediator signifikan. R²=0,780.",
      "Tidak mengkaji kerja paruh waktu, menggunakan mediasi yang mempersulit interpretasi efek langsung.",
      "Penelitian ini menambahkan kerja paruh waktu, tidak menggunakan mediasi efikasi diri, menggunakan regresi berganda, objek di FEB UNIMAL."
    ],
    [
      "7",
      "Kamilatunnisa et al. (2025)\nKatalis: Jurnal Pendidikan Ekonomi Vol.2 No.3\nDOI: 10.62383/katalis.v2i3.2072",
      "Analisis Pengaruh Perilaku Manajemen Waktu terhadap Prestasi Akademik Mahasiswa",
      "X: Manajemen Waktu\nY: Prestasi Akademik",
      "Kuantitatif, regresi linear sederhana",
      "50 mahasiswa FE Universitas Pasundan",
      "Manajemen waktu berpengaruh positif signifikan terhadap prestasi akademik (R²=0,505; p=0,000; Y=10,742+0,844X).",
      "Hanya dua variabel, sampel kecil (50 responden), tidak mengkaji faktor eksternal.",
      "Penelitian ini menambahkan variabel kerja paruh waktu dan lingkungan kampus, menggunakan regresi berganda, objek di FEB UNIMAL."
    ],
    [
      "8",
      "Inayah et al. (2023)\nJurnal PESHUM Vol.2 No.2\nDOI: 10.56799/peshum.v2i2.1391",
      "Pengaruh Manajemen Waktu terhadap Prestasi Akademik Mahasiswa yang Bekerja di Kota Makassar",
      "X: Manajemen Waktu\nY: Prestasi Akademik",
      "Kuantitatif, regresi ordinal",
      "89 mahasiswa pekerja di Kota Makassar",
      "Manajemen waktu berpengaruh positif signifikan (Nagelkerke R²=0,154; p=0,005). 71,9% mahasiswa bekerja memiliki IPK tergolong baik.",
      "Hanya dua variabel, tidak mengkaji lingkungan kampus, cakupan multi-institusi tanpa kontrol institusional.",
      "Penelitian ini menambahkan kerja paruh waktu dan lingkungan kampus, fokus satu institusi (FEB UNIMAL) agar lebih terkontrol."
    ],
    [
      "9",
      "Linggasari & Kurniawan (2019)\nJUPE: Jurnal Pendidikan Ekonomi Vol.7 No.3",
      "Hubungan Kerja Paruh Waktu dengan Prestasi Akademik Mahasiswa Pendidikan Ekonomi UNESA",
      "X: Kerja Paruh Waktu\nY: Prestasi Akademik",
      "Kuantitatif deskriptif, korelasi Karl Pearson",
      "56 mahasiswa (sensus) Prodi Pend. Ekonomi UNESA angk.2015",
      "Terdapat hubungan positif signifikan (r=0,330; sig=0,013). Mahasiswa bekerja <8 jam/hari dapat mempertahankan IPK memuaskan.",
      "Hanya dua variabel, sampel kecil dari satu angkatan, tidak mengontrol manajemen waktu dan lingkungan kampus.",
      "Penelitian ini menambahkan manajemen waktu dan lingkungan kampus, menggunakan regresi berganda, objek di Prodi Manajemen FEB UNIMAL."
    ]
  ];

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => hCell(h, cols[i]))
  });

  const dataRows = data.map(row => new TableRow({
    children: row.map((cell, i) => bCell(cell, cols[i]))
  }));

  return new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: cols,
    rows: [headerRow, ...dataRows]
  });
}

// ===================== MAIN DOCUMENT =====================

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "Times New Roman", size: 24 }
      }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 2268, right: 1701, bottom: 1701, left: 2268 } // 4cm left, 3cm others (approx)
      }
    },
    children: [

      // ===================== BAB II HEADER =====================
      makeHeading1("BAB II"),
      makeHeading1("TINJAUAN PUSTAKA"),

      // ===================== 2.1 LANDASAN TEORI =====================
      makeHeading2("2.1 Landasan Teori"),

      // 2.1.1 Grand Theory: SDT
      makeHeading2("2.1.1 Self-Determination Theory (SDT) sebagai Grand Theory"),

      makePara("Self-Determination Theory (SDT) atau Teori Penentuan Diri merupakan grand theory yang digunakan dalam penelitian ini sebagai kerangka teoritis utama. Teori ini dikembangkan secara sistematis oleh Edward L. Deci dan Richard M. Ryan sejak tahun 1985 dan terus disempurnakan hingga saat ini. SDT merupakan teori motivasi makro yang berfokus pada bagaimana faktor-faktor sosial dan lingkungan memfasilitasi atau menghalangi motivasi intrinsik, pengembangan diri, dan kesejahteraan psikologis individu (Ryan & Deci, 2000)."),

      makePara("Deci dan Ryan membangun SDT atas dasar keprihatinan terhadap perspektif behavioristik yang terlalu menekankan peran reward dan punishment eksternal dalam membentuk perilaku manusia. Deci (1971) dalam penelitian awalnya menemukan bahwa pemberian reward ekstrinsik justru dapat mengurangi motivasi intrinsik seseorang, fenomena yang kemudian dikenal sebagai \"crowding out effect\". Temuan ini mendorong Deci dan Ryan untuk mengembangkan teori yang lebih komprehensif tentang motivasi manusia yang mempertimbangkan faktor internal maupun eksternal secara bersamaan."),

      makePara("SDT berargumen bahwa manusia memiliki tiga kebutuhan psikologis dasar yang bersifat universal, yaitu: (1) kebutuhan akan kompetensi (competence), yaitu kebutuhan untuk merasa mampu dan efektif dalam berinteraksi dengan lingkungan; (2) kebutuhan akan otonomi (autonomy), yaitu kebutuhan untuk merasa memiliki kebebasan dan inisiatif dalam tindakan dan keputusan; dan (3) kebutuhan akan keterkaitan (relatedness), yaitu kebutuhan untuk merasa terhubung, diterima, dan dipedulikan oleh orang lain. Ketika ketiga kebutuhan ini terpenuhi secara optimal, individu cenderung menunjukkan motivasi yang lebih tinggi, performa yang lebih baik, dan kesejahteraan psikologis yang lebih optimal (Deci & Ryan, 2000; Ryan & Deci, 2017)."),

      makePara("Dalam konteks pendidikan, SDT telah banyak digunakan untuk menjelaskan variasi prestasi akademik di kalangan mahasiswa. Howard et al. (2021) dalam meta-analisis yang mencakup lebih dari 100.000 partisipan menemukan bahwa motivasi otonom yang bersumber dari kepuasan kebutuhan psikologis dasar berhubungan positif dan signifikan dengan prestasi akademik. Bureau et al. (2022) dalam meta-analisis terhadap 344 studi menemukan bahwa dukungan otonomi dari lingkungan, termasuk lingkungan kampus, merupakan prediktor kuat motivasi otonom yang kemudian memengaruhi prestasi."),

      makePara("Relevansi SDT dengan penelitian ini terletak pada tiga aspek utama. Pertama, manajemen waktu sebagai variabel bebas pertama (X1) merupakan manifestasi dari kebutuhan kompetensi dan otonomi dalam SDT. Mahasiswa yang mampu mengelola waktu dengan efektif menunjukkan kompetensi diri yang tinggi dan otonomi dalam mengatur aktivitas belajar mereka, yang pada gilirannya mendukung motivasi intrinsik dan prestasi akademik yang lebih baik. Kedua, kerja paruh waktu sebagai variabel bebas kedua (X2) merupakan respons terhadap dorongan eksternal berupa kebutuhan finansial yang dapat memengaruhi kualitas dan kuantitas waktu belajar, berkaitan langsung dengan pemenuhan kebutuhan kompetensi akademik dalam kerangka SDT. Ketiga, lingkungan kampus sebagai variabel bebas ketiga (X3) berperan sebagai konteks ekstrinsik yang dalam kerangka SDT memfasilitasi atau menghambat pemenuhan kebutuhan psikologis dasar mahasiswa, yang selanjutnya memengaruhi motivasi dan prestasi akademik mereka."),

      // 2.1.2 Prestasi Akademik
      makeHeading2("2.1.2 Prestasi Akademik (Variabel Y)"),
      makeHeading3("2.1.2.1 Definisi Prestasi Akademik"),

      makePara("Prestasi akademik merupakan hasil yang dicapai oleh mahasiswa dalam proses pembelajaran yang mencerminkan tingkat penguasaan pengetahuan, keterampilan, dan sikap yang telah dipelajari selama menempuh pendidikan formal di perguruan tinggi. Prestasi akademik pada umumnya diukur melalui nilai atau skor yang diperoleh mahasiswa dari serangkaian evaluasi dan penilaian yang dilakukan selama proses perkuliahan."),

      makePara("Nana Sudjana (2014) mendefinisikan prestasi belajar sebagai kemampuan-kemampuan yang dimiliki siswa atau mahasiswa setelah ia menerima pengalaman belajarnya. Kemampuan tersebut mencakup aspek kognitif yang berkaitan dengan pemahaman konsep dan penerapan pengetahuan, aspek afektif yang berkaitan dengan sikap dan nilai-nilai, serta aspek psikomotor yang berkaitan dengan keterampilan praktis. Hamalik (2016) menyatakan bahwa prestasi akademik merupakan ukuran keberhasilan suatu proses belajar yang mencerminkan perubahan pada aspek pengetahuan (kognitif), sikap (afektif), dan keterampilan (psikomotor) yang diperoleh mahasiswa sebagai hasil dari kegiatan belajar di perguruan tinggi."),

      makePara("Zimmerman (2002) menekankan bahwa prestasi akademik tidak sekadar mencerminkan kemampuan kognitif atau kecerdasan seseorang, melainkan merupakan produk dari interaksi kompleks antara faktor internal seperti motivasi, strategi belajar, dan pengaturan diri (self-regulation), serta faktor eksternal seperti lingkungan belajar dan dukungan sosial. Pandangan ini relevan dengan penelitian ini yang mengkaji pengaruh manajemen waktu (faktor internal) dan kerja paruh waktu serta lingkungan kampus (faktor eksternal) terhadap prestasi akademik."),

      makePara("Dalam konteks pendidikan tinggi di Indonesia, prestasi akademik mahasiswa umumnya direpresentasikan melalui Indeks Prestasi Semester (IPS) dan Indeks Prestasi Kumulatif (IPK) yang dihitung berdasarkan nilai akhir seluruh mata kuliah yang telah ditempuh. Dalam penelitian ini, prestasi akademik mahasiswa dioperasionalisasikan melalui IPK yang diperoleh mahasiswa hingga semester terakhir yang ditempuh, ditambah dengan beberapa indikator perilaku akademik lainnya."),

      makeHeading3("2.1.2.2 Faktor-Faktor yang Memengaruhi Prestasi Akademik"),

      makePara("Prestasi akademik mahasiswa dipengaruhi oleh berbagai faktor, baik yang bersumber dari dalam diri mahasiswa (faktor internal) maupun dari luar diri mahasiswa (faktor eksternal). Faktor-faktor internal yang memengaruhi prestasi akademik meliputi:"),

      makeAlphaItem("a", "Kemampuan kognitif dan kecerdasan, yang merupakan faktor bawaan yang memengaruhi kemampuan mahasiswa dalam memahami dan memproses informasi."),
      makeAlphaItem("b", "Motivasi belajar, yang mencakup motivasi intrinsik dan ekstrinsik yang mendorong mahasiswa untuk belajar dengan sungguh-sungguh."),
      makeAlphaItem("c", "Manajemen waktu, yaitu kemampuan mahasiswa dalam merencanakan, mengorganisasikan, dan menggunakan waktu belajar secara efektif dan efisien."),
      makeAlphaItem("d", "Efikasi diri (self-efficacy), yaitu keyakinan mahasiswa terhadap kemampuan mereka untuk berhasil dalam tugas-tugas akademik."),
      makeAlphaItem("e", "Disiplin dan kebiasaan belajar, yang mencakup konsistensi dan regularitas dalam melakukan aktivitas belajar."),

      makePara("Adapun faktor-faktor eksternal yang memengaruhi prestasi akademik meliputi:"),

      makeAlphaItem("a", "Lingkungan kampus, mencakup kualitas fasilitas belajar, suasana akademik, kualitas pengajaran, dan kondisi fisik kampus."),
      makeAlphaItem("b", "Dukungan keluarga, yang mencakup dukungan emosional, finansial, dan motivasional dari orang tua dan anggota keluarga lainnya."),
      makeAlphaItem("c", "Status pekerjaan, yaitu apakah mahasiswa juga bekerja di luar jam kuliah (bekerja paruh waktu) yang dapat memengaruhi ketersediaan waktu belajar."),
      makeAlphaItem("d", "Interaksi sosial, mencakup hubungan dengan dosen, teman sebaya, dan komunitas kampus yang dapat mendukung atau menghambat proses belajar."),
      makeAlphaItem("e", "Kondisi sosial-ekonomi, yang memengaruhi akses mahasiswa terhadap sumber daya belajar dan tingkat stres finansial yang dihadapi."),

      makeHeading3("2.1.2.3 Indikator Prestasi Akademik"),

      makePara("Pengukuran prestasi akademik dalam penelitian ini menggunakan beberapa indikator yang dikembangkan berdasarkan kajian literatur. Mengacu pada Agustina dan Mardalis (2024), Dewi et al. (2025), serta Inayah et al. (2023), indikator prestasi akademik yang digunakan dalam penelitian ini adalah:"),

      makeNumberedItem("1", "Indeks Prestasi Kumulatif (IPK), yaitu rata-rata nilai tertimbang seluruh mata kuliah yang telah ditempuh mahasiswa hingga semester terakhir yang diselesaikan."),
      makeNumberedItem("2", "Ketepatan penyelesaian tugas, yaitu kemampuan mahasiswa dalam menyelesaikan dan mengumpulkan tugas-tugas akademik tepat waktu sesuai tenggat yang ditetapkan dosen."),
      makeNumberedItem("3", "Keaktifan dalam perkuliahan, yang mencakup partisipasi aktif mahasiswa dalam diskusi kelas, tanya-jawab, dan kegiatan pembelajaran lainnya."),
      makeNumberedItem("4", "Tingkat kehadiran, yaitu persentase kehadiran mahasiswa dalam setiap pertemuan perkuliahan selama satu semester."),

      // 2.1.3 Manajemen Waktu
      makeHeading2("2.1.3 Manajemen Waktu (Variabel X1)"),
      makeHeading3("2.1.3.1 Definisi Manajemen Waktu"),

      makePara("Manajemen waktu adalah proses pengendalian dan pemanfaatan waktu secara terencana dan sistematis guna mencapai tujuan yang telah ditetapkan secara efektif dan efisien. Dalam konteks akademik, manajemen waktu merujuk pada kemampuan mahasiswa untuk mengalokasikan, mengatur, dan memanfaatkan waktu yang tersedia secara optimal antara kegiatan belajar, aktivitas sosial, kegiatan ekstrakurikuler, dan apabila relevan, kegiatan bekerja."),

      makePara("Macan (1994) mendefinisikan manajemen waktu sebagai kemampuan untuk mengalokasikan waktu secara efektif di antara berbagai kegiatan dengan menetapkan tujuan, merencanakan kegiatan, dan memprioritaskan tugas. Definisi ini menekankan tiga komponen utama manajemen waktu, yaitu penetapan tujuan (goal setting), perencanaan kegiatan (planning), dan prioritisasi tugas (prioritizing). Model Macan (1994) menjadi salah satu model manajemen waktu yang paling banyak dikutip dalam penelitian akademik dan menjadi acuan dalam pengembangan instrumen penelitian ini."),

      makePara("Kamilatunnisa et al. (2025) mendefinisikan manajemen waktu sebagai kemampuan individu dalam menggunakan waktu secara sadar dan terarah untuk mencapai tujuan hidup, yang meliputi kemampuan merencanakan, mengorganisasikan, melaksanakan, dan mengevaluasi penggunaan waktu. Nurrahmaniah (2019) menegaskan bahwa manajemen waktu merupakan keterampilan yang dapat dipelajari dan ditingkatkan melalui latihan dan kesadaran diri, bukan sekadar bakat bawaan."),

      makePara("Dalam penelitian ini, manajemen waktu didefinisikan sebagai kemampuan mahasiswa dalam merencanakan, mengorganisasikan, mengarahkan, dan mengendalikan penggunaan waktu sehari-hari secara efektif guna memenuhi tuntutan akademik dan non-akademik secara seimbang, sehingga dapat mencapai prestasi belajar yang optimal."),

      makeHeading3("2.1.3.2 Faktor-Faktor yang Memengaruhi Manajemen Waktu"),

      makePara("Kemampuan manajemen waktu seseorang dipengaruhi oleh berbagai faktor, baik internal maupun eksternal. Beberapa faktor yang memengaruhi manajemen waktu mahasiswa antara lain:"),

      makeNumberedItem("1", "Motivasi dan kesadaran diri: Mahasiswa dengan motivasi intrinsik yang tinggi cenderung lebih teratur dalam mengelola waktu karena memiliki kesadaran akan pentingnya efisiensi waktu bagi pencapaian tujuan akademik mereka."),
      makeNumberedItem("2", "Beban kerja dan aktivitas: Semakin banyak aktivitas yang dilakukan mahasiswa di luar perkuliahan, seperti bekerja paruh waktu atau berorganisasi, semakin kompleks tantangan manajemen waktu yang dihadapi."),
      makeNumberedItem("3", "Lingkungan belajar: Kondisi lingkungan yang mendukung, seperti tersedianya ruang belajar yang nyaman dan fasilitas yang memadai, memudahkan mahasiswa dalam mengalokasikan waktu untuk belajar secara efektif."),
      makeNumberedItem("4", "Teknologi dan media sosial: Penggunaan teknologi dan media sosial yang tidak terkontrol dapat menjadi time trap yang mengganggu efektivitas manajemen waktu mahasiswa."),
      makeNumberedItem("5", "Dukungan sosial: Dukungan dari keluarga, dosen, dan teman sebaya dapat membantu mahasiswa dalam menetapkan prioritas dan mengelola waktu dengan lebih baik."),

      makeHeading3("2.1.3.3 Indikator Manajemen Waktu"),

      makePara("Pengukuran manajemen waktu dalam penelitian ini mengacu pada dimensi yang dikembangkan oleh Macan (1994) dan dimodifikasi oleh berbagai peneliti berikutnya. Berdasarkan Kamilatunnisa et al. (2025) dan Andis Tira et al. (2025), indikator manajemen waktu yang digunakan dalam penelitian ini adalah:"),

      makeNumberedItem("1", "Menetapkan tujuan dan prioritas (goal setting and prioritizing): kemampuan mahasiswa dalam menetapkan tujuan belajar yang jelas dan menentukan skala prioritas dalam penggunaan waktu."),
      makeNumberedItem("2", "Membuat daftar tugas (making to-do list): kebiasaan mahasiswa dalam membuat perencanaan tertulis tentang tugas-tugas yang harus diselesaikan beserta tenggat waktunya."),
      makeNumberedItem("3", "Penjadwalan (scheduling): kemampuan mahasiswa dalam membuat jadwal kegiatan harian, mingguan, dan semesteran yang mengalokasikan waktu secara proporsional untuk belajar dan aktivitas lainnya."),
      makeNumberedItem("4", "Menghindari penundaan (avoiding procrastination): kemampuan mahasiswa untuk tidak menunda-nunda penyelesaian tugas dan memulai pekerjaan tepat waktu."),
      makeNumberedItem("5", "Pengorganisasian diri (self-organization): kemampuan mahasiswa dalam mengorganisasikan tempat belajar, materi perkuliahan, dan sumber daya belajar agar dapat diakses dengan mudah dan efisien."),

      // 2.1.4 Kerja Paruh Waktu
      makeHeading2("2.1.4 Kerja Paruh Waktu (Variabel X2)"),
      makeHeading3("2.1.4.1 Definisi Kerja Paruh Waktu"),

      makePara("Kerja paruh waktu (part-time work) adalah pekerjaan yang dilakukan dengan jam kerja lebih sedikit dibandingkan jam kerja penuh waktu (full-time), umumnya kurang dari 35 jam per minggu. Dalam konteks mahasiswa, kerja paruh waktu merujuk pada aktivitas bekerja yang dilakukan di samping kegiatan perkuliahan formal, baik yang bersifat tetap maupun tidak tetap, dengan tujuan utama untuk memperoleh penghasilan."),

      makePara("Huda et al. (2023) mendefinisikan kerja paruh waktu sebagai pekerjaan yang dilakukan oleh mahasiswa di luar jam kuliah, baik di sektor formal maupun informal, dengan waktu kerja yang terbatas sehingga tidak mengganggu jadwal perkuliahan secara keseluruhan. Menurut Agustina dan Mardalis (2024), kerja paruh waktu mahasiswa adalah keterlibatan mahasiswa dalam kegiatan ekonomi produktif yang memberikan imbalan finansial, yang dilakukan secara bersamaan dengan kegiatan belajar di perguruan tinggi. Alvionita et al. (2022) menambahkan bahwa kerja part-time mahasiswa mencakup berbagai jenis pekerjaan, mulai dari tenaga pengajar, barista, staf administrasi, hingga pekerja lepas di bidang kreatif dan digital, yang umumnya memiliki fleksibilitas jadwal yang cukup tinggi."),

      makePara("Dalam penelitian ini, kerja paruh waktu didefinisikan sebagai segala bentuk aktivitas bekerja yang dilakukan mahasiswa Prodi Manajemen FEB UNIMAL di luar jam perkuliahan resmi, baik bersifat tetap maupun tidak tetap, dengan jam kerja tidak melebihi 35 jam per minggu, dengan tujuan memperoleh penghasilan untuk memenuhi kebutuhan finansial selama menempuh pendidikan."),

      makeHeading3("2.1.4.2 Karakteristik dan Faktor Pendorong Kerja Paruh Waktu"),

      makePara("Keputusan mahasiswa untuk bekerja paruh waktu dipengaruhi oleh berbagai faktor. Berdasarkan kajian Huda et al. (2023), Agustina dan Mardalis (2024), serta Alvionita et al. (2022), faktor-faktor yang mendorong mahasiswa untuk bekerja paruh waktu antara lain:"),

      makeNumberedItem("1", "Faktor ekonomi: kebutuhan finansial untuk membiayai uang kuliah tunggal (UKT), biaya hidup, dan kebutuhan personal merupakan motivasi utama mahasiswa bekerja paruh waktu. Hal ini terutama relevan bagi mahasiswa yang berasal dari keluarga berpenghasilan menengah ke bawah."),
      makeNumberedItem("2", "Pengembangan pengalaman kerja: banyak mahasiswa yang bekerja paruh waktu untuk mendapatkan pengalaman praktis di bidang yang relevan dengan jurusan mereka, sebagai persiapan memasuki dunia kerja setelah lulus."),
      makeNumberedItem("3", "Pengembangan jaringan sosial (networking): bekerja memberikan kesempatan kepada mahasiswa untuk membangun relasi profesional yang dapat bermanfaat di masa mendatang."),
      makeNumberedItem("4", "Dorongan kemandirian: beberapa mahasiswa memilih bekerja sebagai manifestasi dari kemandirian dan keinginan untuk tidak bergantung sepenuhnya pada orang tua."),
      makeNumberedItem("5", "Kesempatan yang tersedia: adanya peluang kerja paruh waktu yang cocok dengan jadwal perkuliahan mendorong mahasiswa untuk memanfaatkan kesempatan tersebut."),

      makeHeading3("2.1.4.3 Indikator Kerja Paruh Waktu"),

      makePara("Pengukuran kerja paruh waktu dalam penelitian ini mengacu pada dimensi yang dikembangkan oleh Huda et al. (2023), Agustina dan Mardalis (2024), serta Alvionita et al. (2022). Indikator kerja paruh waktu yang digunakan dalam penelitian ini adalah:"),

      makeNumberedItem("1", "Jumlah jam kerja per minggu: rata-rata waktu yang dihabiskan untuk bekerja dalam satu minggu, yang memengaruhi ketersediaan waktu untuk belajar dan beristirahat."),
      makeNumberedItem("2", "Jenis pekerjaan: karakteristik dan sifat pekerjaan yang dilakukan, apakah relevan dengan bidang studi atau tidak, serta tingkat fleksibilitas jadwal kerja yang tersedia."),
      makeNumberedItem("3", "Pengaruh pekerjaan terhadap jadwal kuliah: sejauh mana pekerjaan yang dilakukan memengaruhi kehadiran dan partisipasi mahasiswa dalam kegiatan perkuliahan."),
      makeNumberedItem("4", "Motivasi bekerja: alasan atau tujuan utama mahasiswa memilih untuk bekerja paruh waktu, apakah didorong oleh kebutuhan finansial, pengembangan diri, atau keduanya."),
      makeNumberedItem("5", "Manfaat finansial: kontribusi penghasilan dari pekerjaan terhadap pemenuhan kebutuhan finansial selama menempuh pendidikan."),

      // 2.1.5 Lingkungan Kampus
      makeHeading2("2.1.5 Lingkungan Kampus (Variabel X3)"),
      makeHeading3("2.1.5.1 Definisi Lingkungan Kampus"),

      makePara("Lingkungan kampus merupakan keseluruhan kondisi fisik, sosial, dan akademik yang ada di dalam kampus yang secara langsung maupun tidak langsung memengaruhi proses belajar dan perkembangan mahasiswa. Lingkungan kampus mencakup aspek-aspek seperti kondisi fasilitas fisik, kualitas pengajaran, iklim akademik, hubungan interpersonal di dalam kampus, serta kebijakan dan program yang dijalankan oleh institusi pendidikan."),

      makePara("Robyansyah et al. (2022) mendefinisikan lingkungan belajar sebagai segala sesuatu yang ada di sekitar peserta didik, baik yang bersifat fisik maupun non-fisik, yang berpotensi memberikan pengaruh terhadap proses dan hasil belajar. Kondisi fisik mencakup ruang kelas, laboratorium, perpustakaan, dan fasilitas penunjang lainnya, sedangkan kondisi non-fisik mencakup suasana belajar, interaksi sosial, dan iklim akademik yang berlaku. Berhanu dan Sewagegn (2024) dari perspektif internasional mendefinisikan iklim kampus sebagai persepsi kolektif mahasiswa terhadap nilai-nilai, norma, dan praktik institusional yang memengaruhi keterlibatan dan prestasi akademik mereka."),

      makePara("Dewi et al. (2025) mengoperasionalisasikan lingkungan kampus sebagai persepsi mahasiswa terhadap kondisi fisik kampus, dukungan akademik, dan iklim sosial yang dirasakan selama menempuh pendidikan. Munira et al. (2024) mendeskripsikan lingkungan kampus sebagai kondisi keseluruhan yang mencakup fasilitas belajar, kualitas pengajaran, dan suasana akademik yang secara bersama-sama memengaruhi prestasi akademik mahasiswa."),

      makePara("Dalam penelitian ini, lingkungan kampus didefinisikan sebagai keseluruhan kondisi dan situasi di dalam kampus, yang mencakup kondisi fisik fasilitas kampus, kualitas pembelajaran yang diberikan oleh dosen, suasana akademik, serta hubungan antarmahasiswa dan antara mahasiswa dengan dosen, yang secara bersama-sama membentuk iklim belajar yang dapat mendukung atau menghambat pencapaian prestasi akademik mahasiswa Prodi Manajemen FEB UNIMAL."),

      makeHeading3("2.1.5.2 Faktor-Faktor Lingkungan Kampus yang Memengaruhi Prestasi Akademik"),

      makePara("Lingkungan kampus yang memengaruhi prestasi akademik mahasiswa dapat dikelompokkan ke dalam beberapa dimensi utama. Berdasarkan kajian Robyansyah et al. (2022), Berhanu dan Sewagegn (2024), Dewi et al. (2025), serta Munira et al. (2024), faktor-faktor lingkungan kampus yang memengaruhi prestasi akademik antara lain:"),

      makeNumberedItem("1", "Kondisi fisik kampus: mencakup ketersediaan dan kualitas ruang kelas, perpustakaan, laboratorium, akses internet, serta fasilitas penunjang belajar lainnya yang memadai."),
      makeNumberedItem("2", "Kualitas pengajaran: mencakup kompetensi akademik dan pedagogis dosen, metode pengajaran yang digunakan, serta kualitas materi dan bahan ajar yang disediakan."),
      makeNumberedItem("3", "Iklim akademik: mencakup suasana belajar di kelas, budaya akademik yang berlaku di kampus, serta norma dan nilai yang mengedepankan prestasi dan integritas akademik."),
      makeNumberedItem("4", "Hubungan sosial: mencakup kualitas interaksi antara mahasiswa dengan dosen, antarsesama mahasiswa, dan antara mahasiswa dengan staf administrasi kampus."),
      makeNumberedItem("5", "Dukungan institusional: mencakup kebijakan, program, dan layanan yang disediakan oleh institusi untuk mendukung keberhasilan akademik mahasiswa, seperti program bimbingan akademik, beasiswa, dan konseling."),

      makeHeading3("2.1.5.3 Indikator Lingkungan Kampus"),

      makePara("Pengukuran lingkungan kampus dalam penelitian ini mengacu pada dimensi yang dikembangkan oleh Robyansyah et al. (2022), Munira et al. (2024), dan Dewi et al. (2025). Indikator lingkungan kampus yang digunakan dalam penelitian ini adalah:"),

      makeNumberedItem("1", "Kenyamanan dan kelengkapan fasilitas fisik: persepsi mahasiswa terhadap kondisi dan kelengkapan fasilitas fisik kampus seperti ruang kelas, perpustakaan, laboratorium komputer, dan akses internet yang tersedia."),
      makeNumberedItem("2", "Kualitas proses pembelajaran: persepsi mahasiswa terhadap kualitas pengajaran yang diberikan oleh dosen, mencakup kejelasan penjelasan materi, penggunaan metode belajar yang variatif, dan ketepatan waktu dalam mengajar."),
      makeNumberedItem("3", "Suasana dan iklim akademik: persepsi mahasiswa terhadap kondisi suasana belajar di lingkungan kampus, termasuk tingkat kompetisi sehat, motivasi kolektif antar mahasiswa, dan budaya belajar yang berlaku."),
      makeNumberedItem("4", "Interaksi sosial akademik: kualitas hubungan antara mahasiswa dengan dosen dan sesama mahasiswa dalam konteks akademik, termasuk keterbukaan dosen dalam berdiskusi dan kemudahan untuk berkonsultasi."),
      makeNumberedItem("5", "Dukungan administrasi dan layanan kampus: kualitas layanan administrasi akademik, kemudahan akses informasi akademik, dan responsivitas staf kampus dalam membantu keperluan mahasiswa."),

      // Tabel Indikator Variabel
      makeHeading3("2.1.5.4 Tabel Operasionalisasi Konsep Variabel"),

      makeTableCaption("Tabel 2.1"),
      makeTableCaption("Indikator Variabel Penelitian"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [400, 1500, 2300, 2626, 2200],
        rows: [
          new TableRow({
            tableHeader: true,
            children: [
              makeHeaderCell("No.", 400),
              makeHeaderCell("Variabel", 1500),
              makeHeaderCell("Definisi Operasional", 2300),
              makeHeaderCell("Indikator", 2626),
              makeHeaderCell("Sumber", 2200)
            ]
          }),
          new TableRow({ children: [
            makeBodyCell("1", 400),
            makeBodyCell("Manajemen Waktu (X1)", 1500),
            makeBodyCell("Kemampuan mahasiswa dalam merencanakan, mengorganisasikan, dan menggunakan waktu belajar secara efektif dan efisien", 2300),
            makeBodyCell("1. Menetapkan tujuan & prioritas\n2. Membuat daftar tugas\n3. Penjadwalan kegiatan\n4. Menghindari penundaan\n5. Pengorganisasian diri", 2626),
            makeBodyCell("Macan (1994); Kamilatunnisa et al. (2025); Andis Tira et al. (2025)", 2200)
          ]}),
          new TableRow({ children: [
            makeBodyCell("2", 400),
            makeBodyCell("Kerja Paruh Waktu (X2)", 1500),
            makeBodyCell("Aktivitas bekerja yang dilakukan mahasiswa di luar jam kuliah untuk memperoleh penghasilan tambahan", 2300),
            makeBodyCell("1. Jumlah jam kerja/minggu\n2. Jenis pekerjaan & fleksibilitas\n3. Pengaruh thd jadwal kuliah\n4. Motivasi bekerja\n5. Manfaat finansial", 2626),
            makeBodyCell("Huda et al. (2023); Agustina & Mardalis (2024); Alvionita et al. (2022)", 2200)
          ]}),
          new TableRow({ children: [
            makeBodyCell("3", 400),
            makeBodyCell("Lingkungan Kampus (X3)", 1500),
            makeBodyCell("Persepsi mahasiswa terhadap kondisi fisik, akademik, dan sosial kampus yang memengaruhi proses belajar", 2300),
            makeBodyCell("1. Fasilitas fisik kampus\n2. Kualitas proses pembelajaran\n3. Suasana & iklim akademik\n4. Interaksi sosial akademik\n5. Dukungan administrasi", 2626),
            makeBodyCell("Robyansyah et al. (2022); Dewi et al. (2025); Munira et al. (2024)", 2200)
          ]}),
          new TableRow({ children: [
            makeBodyCell("4", 400),
            makeBodyCell("Prestasi Akademik (Y)", 1500),
            makeBodyCell("Hasil belajar mahasiswa yang diwujudkan dalam IPK dan aspek keberhasilan akademik lainnya", 2300),
            makeBodyCell("1. Indeks Prestasi Kumulatif (IPK)\n2. Ketepatan penyelesaian tugas\n3. Keaktifan dalam perkuliahan\n4. Tingkat kehadiran", 2626),
            makeBodyCell("Agustina & Mardalis (2024); Dewi et al. (2025); Inayah et al. (2023)", 2200)
          ]})
        ]
      }),
      makeItalicCaption("Sumber: Diolah oleh peneliti dari berbagai sumber, 2025"),

      // ===================== 2.2 PENELITIAN TERDAHULU =====================
      makeHeading2("2.2 Penelitian Terdahulu"),

      makePara("Penelitian terdahulu yang relevan dengan topik penelitian ini disajikan sebagai landasan empiris dalam penyusunan hipotesis. Berikut ini adalah ringkasan hasil kajian terhadap penelitian-penelitian sebelumnya yang berkaitan dengan variabel manajemen waktu, kerja paruh waktu, lingkungan kampus, dan prestasi akademik mahasiswa:"),

      makeTableCaption("Tabel 2.2"),
      makeTableCaption("Penelitian Terdahulu"),

      makePenelitianTable(),
      makeItalicCaption("Sumber: Diolah oleh peneliti dari berbagai sumber, 2025"),

      makeHeading3("Analisis Kritis Penelitian Terdahulu"),

      makePara("Berdasarkan kajian terhadap sembilan penelitian terdahulu di atas, terdapat beberapa hal yang dapat disintesis. Pertama, dari segi persamaan, seluruh penelitian di atas menggunakan pendekatan kuantitatif dengan variabel dependen berupa prestasi akademik yang diukur melalui IPK atau nilai akademik. Hampir seluruh penelitian juga menggunakan instrumen kuesioner dengan skala Likert dan teknik analisis regresi atau SEM."),

      makePara("Kedua, dari segi perbedaan, mayoritas penelitian sebelumnya hanya mengkaji satu atau dua variabel independen secara bersamaan, tanpa mengintegrasikan ketiga variabel utama penelitian ini (manajemen waktu, kerja paruh waktu, dan lingkungan kampus) dalam satu model. Selain itu, tidak ada penelitian sebelumnya yang dilakukan di Prodi Manajemen FEB Universitas Malikussaleh, sehingga konteks lokal Aceh belum terwakili dalam literatur yang ada."),

      makePara("Ketiga, research gap yang teridentifikasi meliputi: (a) belum adanya penelitian yang mengintegrasikan ketiga variabel tersebut secara simultan dalam satu model regresi; (b) belum ada penelitian serupa di lingkungan FEB UNIMAL; dan (c) terdapat inkonsistensi hasil penelitian terkait arah pengaruh kerja paruh waktu terhadap prestasi akademik, di mana sebagian penelitian menemukan pengaruh positif (Agustina & Mardalis, 2024; Alvionita et al., 2022; Linggasari & Kurniawan, 2019) sementara penelitian lain menemukan pengaruh negatif (Huda et al., 2023). Inkonsistensi ini menegaskan perlunya penelitian yang mengontrol variabel-variabel lain seperti manajemen waktu dan lingkungan kampus secara bersamaan."),

      // ===================== 2.3 PENGARUH ANTAR VARIABEL =====================
      makeHeading2("2.3 Pengaruh Antar Variabel"),
      makeHeading2("2.3.1 Pengaruh Manajemen Waktu terhadap Prestasi Akademik"),

      makePara("Manajemen waktu merupakan salah satu faktor internal yang memiliki peran krusial dalam menentukan prestasi akademik mahasiswa. Mahasiswa yang mampu mengelola waktu dengan efektif cenderung memiliki lebih banyak waktu berkualitas untuk belajar, mengerjakan tugas, dan mempersiapkan diri menghadapi ujian, yang pada akhirnya berkontribusi pada perolehan IPK yang lebih tinggi."),

      makePara("Kamilatunnisa et al. (2025) dalam penelitiannya menemukan bahwa perilaku manajemen waktu berpengaruh positif dan signifikan terhadap prestasi akademik mahasiswa dengan nilai R2 sebesar 0,505, yang berarti manajemen waktu menjelaskan 50,5% variasi prestasi akademik mahasiswa Fakultas Ekonomi Universitas Pasundan. Hasil serupa ditemukan oleh Inayah et al. (2023) yang menemukan pengaruh positif signifikan manajemen waktu terhadap prestasi akademik mahasiswa yang bekerja di Kota Makassar (p = 0,005). Andis Tira et al. (2025) juga menemukan pengaruh parsial signifikan manajemen waktu terhadap prestasi akademik mahasiswa Teknik Mesin Universitas Riau (p = 0,003; t = 3,114), sementara Dewi et al. (2025) mengonfirmasi bahwa manajemen waktu berpengaruh positif signifikan terhadap prestasi mahasiswa Manajemen FEB Universitas Jember."),

      makePara("Dalam kerangka SDT, mahasiswa dengan manajemen waktu yang baik menunjukkan pemenuhan kebutuhan kompetensi dan otonomi yang tinggi. Kemampuan merencanakan dan mengalokasikan waktu secara sadar merupakan ekspresi dari otonomi mahasiswa dalam mengatur kehidupan akademik mereka, yang mendorong motivasi intrinsik dan pada akhirnya meningkatkan prestasi akademik. Nurrahmaniah (2019) menegaskan bahwa peningkatan manajemen waktu secara konsisten berkontribusi pada peningkatan prestasi akademik mahasiswa."),

      makePara("Berdasarkan uraian teoritis dan empiris tersebut, dapat disimpulkan bahwa manajemen waktu berpengaruh positif terhadap prestasi akademik mahasiswa. Semakin baik kemampuan manajemen waktu yang dimiliki mahasiswa, semakin tinggi pula prestasi akademik yang dapat mereka capai."),

      makeHeading2("2.3.2 Pengaruh Kerja Paruh Waktu terhadap Prestasi Akademik"),

      makePara("Pengaruh kerja paruh waktu terhadap prestasi akademik merupakan topik yang masih diperdebatkan dalam literatur akademik dan menunjukkan inkonsistensi hasil di berbagai konteks penelitian. Terdapat dua pandangan utama yang saling berhadapan: pertama, pandangan yang menyatakan bahwa kerja paruh waktu berdampak negatif terhadap prestasi akademik karena mengurangi waktu dan energi yang tersedia untuk belajar; kedua, pandangan yang menyatakan bahwa kerja paruh waktu justru dapat berdampak positif melalui peningkatan motivasi, disiplin, dan rasa tanggung jawab mahasiswa."),

      makePara("Mendukung pandangan positif, Agustina dan Mardalis (2024) menemukan pengaruh positif dan signifikan kerja paruh waktu terhadap prestasi akademik mahasiswa (original sample 0,378; t-statistik 4,637; p = 0,000). Linggasari dan Kurniawan (2019) juga menemukan hubungan positif signifikan antara kerja paruh waktu dan prestasi akademik (r = 0,330; sig = 0,013), dengan catatan penting bahwa mahasiswa yang bekerja kurang dari 8 jam per hari masih dapat mempertahankan IPK yang memuaskan. Alvionita et al. (2022) mengonfirmasi hasil ini dengan menemukan bahwa kerja part-time berpengaruh positif signifikan terhadap prestasi akademik mahasiswa Pendidikan Ekonomi Universitas PGRI Ronggolawe (koef = 0,491; t = 7,235; sig = 0,000)."),

      makePara("Sebaliknya, mendukung pandangan negatif, Huda et al. (2023) menemukan bahwa kerja paruh waktu berpengaruh negatif signifikan terhadap prestasi akademik mahasiswa UIN Sumatera Utara, terutama karena berkurangnya waktu untuk mengerjakan tugas dan mempersiapkan ujian. Temuan yang kontradiktif ini mengindikasikan bahwa pengaruh kerja paruh waktu terhadap prestasi akademik bersifat kondisional dan sangat bergantung pada faktor-faktor lain seperti jumlah jam kerja, jenis pekerjaan, kemampuan manajemen waktu, dan konteks lingkungan mahasiswa."),

      makePara("Dalam konteks Prodi Manajemen FEB UNIMAL, di mana banyak mahasiswa berasal dari latar belakang sosial-ekonomi menengah ke bawah dan kondisi ekonomi Aceh yang masih berkembang, fenomena kerja paruh waktu perlu dikaji secara komprehensif terhadap dampaknya pada prestasi akademik. Research gap inilah yang menjadi salah satu justifikasi penting dilakukannya penelitian ini."),

      makeHeading2("2.3.3 Pengaruh Lingkungan Kampus terhadap Prestasi Akademik"),

      makePara("Lingkungan kampus merupakan faktor eksternal yang memiliki pengaruh signifikan terhadap prestasi akademik mahasiswa. Kampus yang memiliki fasilitas lengkap, dosen berkualitas, iklim akademik yang kondusif, dan sistem dukungan yang baik menciptakan ekosistem belajar yang memungkinkan mahasiswa mencapai potensi akademik terbaik mereka."),

      makePara("Robyansyah et al. (2022) menemukan pengaruh positif dan signifikan lingkungan belajar terhadap prestasi belajar taruna Politeknik Negeri Bengkalis dengan koefisien jalur 0,498 (sig. = 0,000), sementara model secara keseluruhan menjelaskan 82,9% variasi prestasi. Dewi et al. (2025) juga menemukan pengaruh positif signifikan lingkungan kampus terhadap prestasi mahasiswa S1 Manajemen FEB Universitas Jember (koef = 0,171; t = 2,824; p = 0,005), dengan total model menjelaskan 78% variasi prestasi. Berhanu dan Sewagegn (2024) dari perspektif internasional menemukan bahwa persepsi iklim kampus berhubungan positif dengan prestasi akademik yang dimediasi oleh keterlibatan mahasiswa (student engagement). Munira et al. (2024) juga menemukan bahwa fasilitas belajar, kualitas pengajaran, dan suasana akademik berpengaruh signifikan terhadap prestasi akademik mahasiswa."),

      makePara("Dalam kerangka SDT, lingkungan kampus yang mendukung berperan dalam memenuhi kebutuhan psikologis dasar mahasiswa. Kualitas pengajaran yang tinggi mendukung pemenuhan kebutuhan kompetensi; metode pembelajaran yang memberdayakan mahasiswa mendukung pemenuhan kebutuhan otonomi; dan interaksi sosial yang positif di kampus mendukung pemenuhan kebutuhan keterkaitan. Pemenuhan ketiga kebutuhan psikologis ini secara simultan mendorong motivasi otonom yang pada gilirannya meningkatkan prestasi akademik mahasiswa."),

      makeHeading2("2.3.4 Pengaruh Simultan Manajemen Waktu, Kerja Paruh Waktu, dan Lingkungan Kampus terhadap Prestasi Akademik"),

      makePara("Secara simultan, manajemen waktu, kerja paruh waktu, dan lingkungan kampus diduga berpengaruh terhadap prestasi akademik mahasiswa. Ketiga variabel ini tidak bekerja secara terpisah-pisah, melainkan saling berinteraksi dalam membentuk hasil akademik mahasiswa. Mahasiswa yang memiliki kemampuan manajemen waktu yang baik, dapat mengelola pekerjaan paruh waktu secara bijaksana, dan berada dalam lingkungan kampus yang kondusif, diperkirakan akan mampu mencapai prestasi akademik yang lebih tinggi dibandingkan mahasiswa yang hanya memiliki salah satu keunggulan tersebut."),

      makePara("Andis Tira et al. (2025) menemukan bahwa manajemen waktu, efikasi diri, dan lingkungan kampus secara simultan berpengaruh signifikan terhadap prestasi akademik dengan nilai Radj sebesar 0,643. Agustina dan Mardalis (2024) mengonfirmasi bahwa kerja paruh waktu, motivasi belajar, dan time management secara simultan berpengaruh signifikan terhadap prestasi akademik. Hasil-hasil penelitian ini memberikan dukungan empiris yang kuat terhadap hipotesis bahwa kombinasi faktor internal (manajemen waktu) dan faktor eksternal (kerja paruh waktu dan lingkungan kampus) secara bersama-sama berkontribusi terhadap prestasi akademik mahasiswa."),

      // ===================== 2.4 KERANGKA KONSEPTUAL =====================
      makeHeading2("2.4 Kerangka Konseptual"),

      makePara("Berdasarkan landasan teori Self-Determination Theory, penelitian terdahulu yang relevan, dan analisis hubungan antar variabel yang telah diuraikan, kerangka konseptual penelitian ini menggambarkan hubungan antara variabel bebas (manajemen waktu, kerja paruh waktu, dan lingkungan kampus) dengan variabel terikat (prestasi akademik mahasiswa). Kerangka konseptual ini menjadi panduan logis dalam pengembangan hipotesis dan desain penelitian."),

      makeTableCaption("Gambar 2.1"),
      makeTableCaption("Kerangka Konseptual Penelitian"),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 200 },
        children: []
      }),

      makeKerangkaBox(),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 100 },
        children: [
          new TextRun({ text: "H1, H2, H3 (parsial) dan H4 (simultan)", italic: true, size: 20, font: "Times New Roman" })
        ]
      }),
      makeItalicCaption("Sumber: Dikembangkan oleh peneliti berdasarkan kajian teoritis dan empiris, 2025"),

      makePara("Dalam kerangka konseptual di atas, manajemen waktu (X1), kerja paruh waktu (X2), dan lingkungan kampus (X3) secara bersama-sama maupun sendiri-sendiri diduga berpengaruh terhadap prestasi akademik (Y) mahasiswa Prodi Manajemen FEB UNIMAL. Manajemen waktu (X1) yang baik memungkinkan mahasiswa mengalokasikan waktu yang cukup dan berkualitas untuk belajar sehingga berdampak positif pada prestasi akademik. Kerja paruh waktu (X2) yang terkelola dengan baik dapat memberikan manfaat finansial dan pengalaman yang mendukung motivasi belajar, namun kerja yang berlebihan dapat mengurangi waktu belajar sehingga berpotensi berdampak negatif pada prestasi. Lingkungan kampus (X3) yang kondusif menyediakan fasilitas, pengajaran berkualitas, dan iklim akademik yang mendorong mahasiswa untuk berprestasi lebih baik. Grand theory Self-Determination Theory menjadi payung teoritis yang menjelaskan mekanisme pengaruh ketiga variabel tersebut terhadap prestasi akademik melalui pemenuhan kebutuhan psikologis dasar mahasiswa (kompetensi, otonomi, dan keterkaitan)."),

      // ===================== 2.5 HIPOTESIS =====================
      makeHeading2("2.5 Hipotesis Penelitian"),

      makePara("Berdasarkan landasan teori, penelitian terdahulu, pengaruh antar variabel, dan kerangka konseptual yang telah diuraikan, hipotesis penelitian ini dirumuskan sebagai berikut:"),

      makePara("H1: Manajemen waktu berpengaruh positif dan signifikan terhadap prestasi akademik mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh.", { bold: true }),

      makePara("Hipotesis ini didasarkan pada temuan Kamilatunnisa et al. (2025) yang menemukan pengaruh positif signifikan manajemen waktu terhadap prestasi akademik (R2=0,505; p=0,000), serta Inayah et al. (2023) dan Andis Tira et al. (2025) yang mengonfirmasi hasil serupa. Secara teoritis, manajemen waktu yang baik mencerminkan pemenuhan kebutuhan kompetensi dan otonomi dalam kerangka SDT, yang mendorong motivasi intrinsik dan peningkatan prestasi akademik."),

      makePara("H2: Kerja paruh waktu berpengaruh signifikan terhadap prestasi akademik mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh.", { bold: true }),

      makePara("Hipotesis ini dirumuskan berdasarkan temuan empiris yang menunjukkan adanya pengaruh kerja paruh waktu terhadap prestasi akademik, meskipun dengan arah yang kontradiktif antar penelitian. Agustina dan Mardalis (2024), Alvionita et al. (2022), serta Linggasari dan Kurniawan (2019) menemukan pengaruh positif, sementara Huda et al. (2023) menemukan pengaruh negatif. Mengingat inkonsistensi temuan ini, hipotesis dirumuskan tanpa menetapkan arah pengaruh, untuk kemudian dibuktikan secara empiris dalam konteks FEB UNIMAL."),

      makePara("H3: Lingkungan kampus berpengaruh positif dan signifikan terhadap prestasi akademik mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh.", { bold: true }),

      makePara("Hipotesis ini didasarkan pada temuan konsisten dalam berbagai penelitian yang menunjukkan pengaruh positif lingkungan kampus terhadap prestasi akademik. Robyansyah et al. (2022) menemukan koefisien pengaruh sebesar 0,498 (sig.=0,000), dan Dewi et al. (2025) mengonfirmasi dengan koefisien 0,171 (t=2,824; p=0,005). Dalam kerangka SDT, lingkungan kampus yang kondusif memenuhi kebutuhan psikologis dasar mahasiswa yang mendorong motivasi otonom dan prestasi akademik."),

      makePara("H4: Manajemen waktu, kerja paruh waktu, dan lingkungan kampus secara simultan berpengaruh signifikan terhadap prestasi akademik mahasiswa Prodi Manajemen Fakultas Ekonomi dan Bisnis Universitas Malikussaleh.", { bold: true }),

      makePara("Hipotesis simultan ini didasarkan pada temuan Andis Tira et al. (2025) yang menemukan pengaruh simultan signifikan dari variabel-variabel serupa (Radj=0,643) dan Agustina dan Mardalis (2024) yang mengonfirmasi pengaruh simultan signifikan dari kombinasi variabel kerja paruh waktu dan manajemen waktu terhadap prestasi akademik. Secara teoritis, ketiga variabel independen dalam penelitian ini merepresentasikan aspek-aspek berbeda yang secara bersama-sama membentuk kondisi optimal bagi prestasi akademik mahasiswa dalam kerangka SDT."),

      // ===================== DAFTAR PUSTAKA =====================
      new Paragraph({
        children: [new PageBreak()]
      }),

      makeHeading1("DAFTAR PUSTAKA"),

      makeDaftarPustakaItem("Agustina, A., & Mardalis, A. (2024). Pengaruh kerja paruh waktu, motivasi belajar dan time management terhadap prestasi akademik (studi kasus mahasiswa bekerja part-time). Jurnal Bisnis dan Ekonomi (JBBE), 17(2). https://doi.org/10.46306/jbbe.v17i2.556"),

      makeDaftarPustakaItem("Alvionita, W. A., Windrayadi, Y. D. P., & Purwanto, H. (2022). Pengaruh kerja part-time dan aktivitas belajar terhadap prestasi akademik mahasiswa Pendidikan Ekonomi Universitas PGRI Ronggolawe Tuban. OPORTUNITAS: Jurnal Pendidikan Ekonomi, Manajemen, Kewirausahaan dan Koperasi, 3(02), 62–67. https://journal.unirow.ac.id/index.php/oportunitas/article/view/504"),

      makeDaftarPustakaItem("Andis Tira, S., Eka, P., & Hermansyah, D. (2025). Pengaruh manajemen waktu, efikasi diri, dan lingkungan kampus terhadap prestasi akademik mahasiswa Teknik Mesin Universitas Riau. Jurnal Bisnis Mahasiswa. https://doi.org/10.60036/jbm.689"),

      makeDaftarPustakaItem("Berhanu, K. Z., & Sewagegn, A. A. (2024). The role of perceived campus climate in students' academic achievements as mediated by students' engagement in higher education institutions. Cogent Education, 11(1). https://doi.org/10.1080/2331186X.2024.2377839"),

      makeDaftarPustakaItem("Bureau, J. S., Howard, J. L., Chong, J. X. Y., & Guay, F. (2022). Pathways to student motivation: A meta-analysis of antecedents of autonomous and controlled motivations. Review of Educational Research, 92(1), 46–88. https://doi.org/10.3102/00346543211042426"),

      makeDaftarPustakaItem("Deci, E. L., & Ryan, R. M. (2000). The \"what\" and \"why\" of goal pursuits: Human needs and the self-determination of behavior. Psychological Inquiry, 11(4), 227–268. https://doi.org/10.1207/S15327965PLI1104_01"),

      makeDaftarPustakaItem("Dewi, A. F. D., Kholidy, M. A. N., & Prajitiasari, E. D. (2025). Efikasi diri sebagai mediasi dalam hubungan lingkungan kampus dan manajemen waktu pada prestasi mahasiswa. Jurnal Pariwisata, Bisnis Digital dan Manajemen, 4(1). https://doi.org/10.33480/jasdim.v4i1.6666"),

      makeDaftarPustakaItem("Ghozali, I. (2018). Aplikasi analisis multivariate dengan program IBM SPSS 25 (edisi 9). Badan Penerbit Universitas Diponegoro."),

      makeDaftarPustakaItem("Hair, J. F., Black, W. C., Babin, B. J., & Anderson, R. E. (2014). Multivariate data analysis (7th ed.). Pearson Education."),

      makeDaftarPustakaItem("Hamalik, O. (2016). Proses belajar mengajar. Bumi Aksara."),

      makeDaftarPustakaItem("Howard, J. L., Bureau, J., Guay, F., Chong, J. X. Y., & Ryan, R. M. (2021). Student motivation and associated outcomes: A meta-analysis from self-determination theory. Perspectives on Psychological Science, 16(6), 1300–1323. https://doi.org/10.1177/1745691620966789"),

      makeDaftarPustakaItem("Huda, M. A. A., Fani, M., Saragih, R. M., & Lestari, D. (2023). Pengaruh kerja paruh waktu terhadap prestasi akademik mahasiswa UIN Sumatera Utara. Madani Jurnal Ilmiah Multidisiplin, 1(6), 447–456. https://doi.org/10.5281/zenodo.8127903"),

      makeDaftarPustakaItem("Inayah, D. N., Daud, M., & Nur, H. (2023). Pengaruh manajemen waktu terhadap prestasi akademik mahasiswa yang bekerja di Kota Makassar. Jurnal Penelitian, Pendidikan, dan Pengajaran (PESHUM), 2(2). https://doi.org/10.56799/peshum.v2i2.1391"),

      makeDaftarPustakaItem("Kamilatunnisa, K., Nurhayati, N., & Rahayu, S. (2025). Analisis pengaruh perilaku manajemen waktu terhadap prestasi akademik mahasiswa. Katalis: Jurnal Pendidikan Ekonomi dan Kewirausahaan, 2(3). https://doi.org/10.62383/katalis.v2i3.2072"),

      makeDaftarPustakaItem("Linggasari, L. Y., & Kurniawan, R. Y. (2019). Hubungan kerja paruh waktu dengan prestasi akademik mahasiswa Jurusan Pendidikan Ekonomi Universitas Negeri Surabaya angkatan 2015. JUPE: Jurnal Pendidikan Ekonomi, 7(3), 92–93."),

      makeDaftarPustakaItem("Macan, T. H. (1994). Time management: Test of a process model. Journal of Applied Psychology, 79(3), 381–391. https://doi.org/10.1037/0021-9010.79.3.381"),

      makeDaftarPustakaItem("Munira, R., Fonna, T., Nadia, S., & Marsitah, I. (2024). Pengaruh lingkungan belajar terhadap prestasi akademik mahasiswa di Universitas Almuslim. Jurnal Pendidikan Guru Sekolah Dasar, 1(4). https://doi.org/10.47134/pgsd.v1i4.770"),

      makeDaftarPustakaItem("Nurrahmaniah. (2019). Peningkatan prestasi akademik melalui manajemen waktu dan minat belajar. Andragogi: Jurnal Diklat Teknis Pendidikan dan Keagamaan, 1(1). https://doi.org/10.36671/andragogi.v1i1.52"),

      makeDaftarPustakaItem("Robyansyah, Indarti, S., & Widayatsari, A. (2022). Pengaruh lingkungan belajar dan disiplin terhadap motivasi belajar dan prestasi belajar taruna Politeknik Negeri Bengkalis Jurusan Kemaritiman. Jurnal Daya Saing, 8(3), 384–395. https://doi.org/10.35446/dayasaing.v8i3.972"),

      makeDaftarPustakaItem("Ryan, R. M., & Deci, E. L. (2000). Self-determination theory and the facilitation of intrinsic motivation, social development, and well-being. American Psychologist, 55(1), 68–78. https://doi.org/10.1037/0003-066X.55.1.68"),

      makeDaftarPustakaItem("Ryan, R. M., & Deci, E. L. (2017). Self-determination theory: Basic psychological needs in motivation, development, and wellness. Guilford Press."),

      makeDaftarPustakaItem("Sugiyono. (2023). Metode penelitian kuantitatif, kualitatif dan R&D (edisi terbaru). Alfabeta."),

      makeDaftarPustakaItem("Sudjana, N. (2014). Penilaian hasil proses belajar mengajar. Remaja Rosdakarya."),

      makeDaftarPustakaItem("Zimmerman, B. J. (2002). Becoming a self-regulated learner: An overview. Theory Into Practice, 41(2), 64–70. https://doi.org/10.1207/s15430421tip4102_2"),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('./hasil.docx', buffer);
  console.log('File berhasil dibuat!');
}).catch(err => {
  console.error('Error:', err);
});
