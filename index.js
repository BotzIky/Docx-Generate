const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  LevelFormat, PageBreak
} = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" };
const borders = { top: border, bottom: border, left: border, right: border };

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, size: 28, font: "Times New Roman" })],
    spacing: { before: 360, after: 180 }
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, bold: true, size: 26, font: "Times New Roman" })],
    spacing: { before: 240, after: 120 }
  });
}

function heading3(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 24, font: "Times New Roman" })],
    spacing: { before: 200, after: 100 }
  });
}

function para(text, options = {}) {
  return new Paragraph({
    children: [new TextRun({ text, size: 24, font: "Times New Roman", ...options })],
    spacing: { before: 80, after: 80, line: 360 },
    indent: { firstLine: 720 },
    alignment: AlignmentType.JUSTIFIED
  });
}

function paraBold(text) {
  return new Paragraph({
    children: [new TextRun({ text, size: 24, font: "Times New Roman", bold: true })],
    spacing: { before: 80, after: 80, line: 360 },
    alignment: AlignmentType.JUSTIFIED
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, size: 24, font: "Times New Roman" })],
    spacing: { before: 60, after: 60, line: 300 }
  });
}

function numbered(text, num) {
  return new Paragraph({
    children: [new TextRun({ text: `${num}. ${text}`, size: 24, font: "Times New Roman" })],
    spacing: { before: 60, after: 60, line: 300 },
    indent: { left: 360, hanging: 360 }
  });
}

function emptyLine() {
  return new Paragraph({ children: [new TextRun("")], spacing: { before: 60, after: 60 } });
}

function tableRow(label, value, isHeader = false) {
  const cellStyle = {
    borders,
    margins: { top: 100, bottom: 100, left: 150, right: 150 }
  };
  return new TableRow({
    children: [
      new TableCell({
        ...cellStyle,
        width: { size: 3000, type: WidthType.DXA },
        shading: isHeader ? { fill: "1F4E79", type: ShadingType.CLEAR } : { fill: "D6E4F0", type: ShadingType.CLEAR },
        children: [new Paragraph({
          children: [new TextRun({
            text: label,
            bold: true,
            size: 22,
            font: "Times New Roman",
            color: isHeader ? "FFFFFF" : "000000"
          })]
        })]
      }),
      new TableCell({
        ...cellStyle,
        width: { size: 6360, type: WidthType.DXA },
        shading: isHeader ? { fill: "1F4E79", type: ShadingType.CLEAR } : { fill: "FFFFFF", type: ShadingType.CLEAR },
        children: [new Paragraph({
          children: [new TextRun({
            text: value,
            size: 22,
            font: "Times New Roman",
            color: isHeader ? "FFFFFF" : "000000"
          })]
        })]
      })
    ]
  });
}

function statusRow(aspek, status, catatan) {
  const statusColor = status === "SESUAI" ? "1D6A39" : status === "KURANG" ? "7B3F00" : "8B0000";
  const bgStatus = status === "SESUAI" ? "C8F0D8" : status === "KURANG" ? "FDEBD0" : "FADBD8";
  return new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: 2800, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: aspek, size: 20, font: "Times New Roman" })] })]
      }),
      new TableCell({
        borders,
        width: { size: 1200, type: WidthType.DXA },
        shading: { fill: bgStatus, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: status, bold: true, size: 20, font: "Times New Roman", color: statusColor })]
        })]
      }),
      new TableCell({
        borders,
        width: { size: 5360, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: catatan, size: 20, font: "Times New Roman" })] })]
      })
    ]
  });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: {
      document: { run: { font: "Times New Roman", size: 24 } }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal",
        run: { size: 28, bold: true, font: "Times New Roman", color: "1F4E79" },
        paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal",
        run: { size: 26, bold: true, font: "Times New Roman", color: "2E75B6" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 }
      }
    },
    children: [
      // ============================================================
      // HALAMAN JUDUL
      // ============================================================
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 480, after: 120 },
        children: [new TextRun({ text: "DOKUMEN HASIL TELAAH DAN PERBANDINGAN", bold: true, size: 32, font: "Times New Roman", color: "1F4E79" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120 },
        children: [new TextRun({ text: "PROPOSAL SKRIPSI", bold: true, size: 28, font: "Times New Roman" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120 },
        children: [new TextRun({ text: "\"Pengaruh Product Description, Product Awareness, Enduring Involvement,", bold: true, size: 24, font: "Times New Roman", italics: true })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: "Situational Involvement, dan Perceived Trust Terhadap Purchase Intention", bold: true, size: 24, font: "Times New Roman", italics: true })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 240 },
        children: [new TextRun({ text: "Pada Marketplace Shopee di Kota Lhokseumawe\"", bold: true, size: 24, font: "Times New Roman", italics: true })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 60 },
        children: [new TextRun({ text: "Disusun oleh:", size: 24, font: "Times New Roman" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: "Putri Shalsabila | NIM. 190410088", bold: true, size: 24, font: "Times New Roman" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: "Program Studi Manajemen | Fakultas Ekonomi dan Bisnis", size: 24, font: "Times New Roman" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 240 },
        children: [new TextRun({ text: "Universitas Malikussaleh, Lhokseumawe", size: 24, font: "Times New Roman" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 60 },
        children: [new TextRun({ text: "Ditelaah oleh:", size: 24, font: "Times New Roman" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: "Khoirul Anwar", bold: true, size: 24, font: "Times New Roman" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 480 },
        children: [new TextRun({ text: "Mei 2026", size: 24, font: "Times New Roman" })]
      }),

      // PAGE BREAK
      new Paragraph({ children: [new PageBreak()] }),

      // ============================================================
      // BAGIAN 1: PENDAHULUAN DOKUMEN
      // ============================================================
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "A. PENDAHULUAN DOKUMEN TELAAH INI", bold: true, size: 28, font: "Times New Roman" })],
        spacing: { before: 240, after: 180 }
      }),

      para("Dokumen ini disusun sebagai hasil telaah menyeluruh terhadap proposal skripsi yang ditulis oleh Putri Shalsabila (NIM. 190410088) dari Program Studi Manajemen, Fakultas Ekonomi dan Bisnis, Universitas Malikussaleh. Telaah ini mencakup penilaian kesesuaian proposal dengan panduan penulisan skripsi yang berlaku, kelengkapan unsur-unsur akademik yang diharapkan, serta keterhubungan antar bab dari Bab 1 sampai Bab 3."),
      
      para("Perlu dicatat bahwa akses langsung ke file panduan penulisan skripsi FEB Unimal dari tautan yang diberikan tidak berhasil terbuka secara otomatis. Oleh karena itu, telaah ini didasarkan pada: (1) isi lengkap proposal yang sudah dilampirkan, (2) standar umum panduan penulisan skripsi FEB di perguruan tinggi Indonesia, dan (3) kaidah akademik yang lazim digunakan pada jenjang S1."),

      para("Bahasa yang digunakan dalam dokumen telaah ini sengaja dibuat lebih santai dan mudah dipahami — tujuannya supaya semua catatan dan masukan bisa langsung dimengerti tanpa perlu ditafsirkan dua kali."),

      emptyLine(),

      // ============================================================
      // BAGIAN 2: GAMBARAN UMUM
      // ============================================================
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "B. GAMBARAN UMUM PROPOSAL", bold: true, size: 28, font: "Times New Roman" })],
        spacing: { before: 240, after: 180 }
      }),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3000, 6360],
        rows: [
          tableRow("ASPEK", "KETERANGAN", true),
          tableRow("Judul", "Pengaruh Product Description, Product Awareness, Enduring Involvement, Situational Involvement, dan Perceived Trust Terhadap Purchase Intention pada Marketplace Shopee di Kota Lhokseumawe"),
          tableRow("Penulis", "Putri Shalsabila, NIM. 190410088"),
          tableRow("Pembimbing", "Dr. Mariyudi, S.E., M.M"),
          tableRow("Program Studi", "Manajemen - FEB Universitas Malikussaleh"),
          tableRow("Tahun Penulisan", "Mei 2026"),
          tableRow("Variabel Independen", "Product Description (X1), Product Awareness (X2), Enduring Involvement (X3), Situational Involvement (X4), Perceived Trust (X5)"),
          tableRow("Variabel Dependen", "Purchase Intention (Y)"),
          tableRow("Objek Penelitian", "Pengguna Marketplace Shopee di Kota Lhokseumawe"),
          tableRow("Metode Analisis", "Regresi Linier Berganda (SPSS v.26)"),
          tableRow("Jumlah Sampel", "96 responden (menggunakan Rumus Lemeshow)")
        ]
      }),

      emptyLine(),

      // ============================================================
      // BAGIAN 3: REKAPITULASI CEPAT
      // ============================================================
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "C. REKAPITULASI CEPAT: SUDAH ADA / BELUM?", bold: true, size: 28, font: "Times New Roman" })],
        spacing: { before: 240, after: 180 }
      }),

      para("Berikut ini ringkasan cepat mengenai unsur-unsur penting yang biasanya harus ada dalam proposal skripsi S1 bidang manajemen. Tabel ini membantu melihat secara sekilas mana yang sudah oke dan mana yang masih perlu perhatian."),
      emptyLine(),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2800, 1200, 5360],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders,
                width: { size: 2800, type: WidthType.DXA },
                shading: { fill: "1F4E79", type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "UNSUR", bold: true, size: 22, font: "Times New Roman", color: "FFFFFF" })] })]
              }),
              new TableCell({
                borders,
                width: { size: 1200, type: WidthType.DXA },
                shading: { fill: "1F4E79", type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "STATUS", bold: true, size: 22, font: "Times New Roman", color: "FFFFFF" })] })]
              }),
              new TableCell({
                borders,
                width: { size: 5360, type: WidthType.DXA },
                shading: { fill: "1F4E79", type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "CATATAN SINGKAT", bold: true, size: 22, font: "Times New Roman", color: "FFFFFF" })] })]
              })
            ]
          }),
          // BAB 1
          statusRow("Latar Belakang Masalah", "SESUAI", "Ada, cukup panjang dan menjelaskan konteks e-commerce di Indonesia"),
          statusRow("Fenomena / Fakta Empiris", "KURANG", "Data statistik dari Databoks ada, tapi fenomena masalah spesifik di Lhokseumawe sangat minim — hanya survei 17 orang"),
          statusRow("Urgensi Penelitian", "KURANG", "Tidak ada paragraf yang secara tegas menjelaskan 'mengapa penelitian ini penting dan mendesak untuk dilakukan sekarang'"),
          statusRow("Research Gap", "KURANG", "Gap disebutkan lewat hasil penelitian yang berbeda-beda, tapi tidak dirumuskan secara eksplisit sebagai gap"),
          statusRow("Rumusan Masalah", "SESUAI", "Ada, 5 pertanyaan sesuai 5 variabel independen"),
          statusRow("Tujuan Penelitian", "SESUAI", "Ada, 5 tujuan sesuai 5 rumusan masalah"),
          statusRow("Manfaat Penelitian", "SESUAI", "Ada teoritis dan praktis, terbagi ke beberapa pihak"),
          // BAB 2
          statusRow("Landasan Teori", "SESUAI", "Ada, cukup lengkap membahas semua variabel"),
          statusRow("Penelitian Terdahulu", "SESUAI", "Ada tabel penelitian terdahulu, 6 penelitian"),
          statusRow("Kerangka Konseptual", "SESUAI", "Ada gambar dan penjelasan hubungan antar variabel"),
          statusRow("Hipotesis", "SESUAI", "Ada, 5 hipotesis parsial sesuai variabel"),
          statusRow("Teori Dasar / Grand Theory", "KURANG", "Tidak ada grand theory atau teori utama yang menjadi payung penelitian (misal: Theory of Planned Behavior, TAM, dll)"),
          // BAB 3
          statusRow("Lokasi & Objek Penelitian", "SESUAI", "Ada, jelas menyebut Kota Lhokseumawe"),
          statusRow("Populasi & Sampel", "SESUAI", "Ada rumus Lemeshow, purposive sampling, kriteria sampel"),
          statusRow("Jenis & Sumber Data", "SESUAI", "Ada primer dan sekunder"),
          statusRow("Teknik Pengumpulan Data", "SESUAI", "Ada kuesioner dengan skala Likert 1-5"),
          statusRow("Definisi Operasional", "SESUAI", "Ada tabel variabel, definisi, indikator, skala"),
          statusRow("Metode Analisis", "SESUAI", "Regresi linier berganda, SPSS 26"),
          statusRow("Uji Instrumen", "SESUAI", "Validitas dan reliabilitas ada"),
          statusRow("Uji Asumsi Klasik", "SESUAI", "Normalitas, multikolinearitas, heteroskedastisitas"),
          statusRow("Uji Hipotesis", "SESUAI", "Uji t dan koefisien determinasi R2"),
          statusRow("Uji F / Simultan", "TIDAK ADA", "Tidak ada uji F (simultan) padahal ini lazim untuk regresi berganda dan biasanya diwajibkan panduan")
        ]
      }),

      emptyLine(),

      // ============================================================
      // BAGIAN 4: ANALISIS BAB 1
      // ============================================================
      new Paragraph({ children: [new PageBreak()] }),

      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "D. ANALISIS MENDALAM BAB 1 — PENDAHULUAN", bold: true, size: 28, font: "Times New Roman" })],
        spacing: { before: 240, after: 180 }
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: "D.1 Latar Belakang Masalah", bold: true, size: 26, font: "Times New Roman" })],
        spacing: { before: 200, after: 120 }
      }),

      para("Secara umum, latar belakang yang ditulis sudah menunjukkan alur pikir yang cukup baik. Penulisnya memulai dari konteks besar — yaitu perkembangan teknologi dan digitalisasi — lalu menyempit ke e-commerce Indonesia, kemudian ke Shopee sebagai platform utama, dan akhirnya ke variabel-variabel yang diteliti. Alur seperti ini (dari umum ke khusus) sudah benar dan sesuai kaidah penulisan skripsi."),

      para("Tapi ada beberapa hal yang perlu diperhatikan:"),

      para("Pertama, soal fenomena. Data dari Databoks yang menunjukkan Shopee sebagai e-commerce dengan pengunjung terbanyak (2,35 miliar) sudah bagus sebagai fakta pendukung. Namun, latar belakang ini kurang menampilkan fenomena masalah yang nyata — bukan sekedar fakta bahwa Shopee populer, tapi mengapa niat beli di Shopee perlu diteliti? Apa masalahnya? Apakah ada kasus di mana deskripsi produk yang buruk menyebabkan konsumen batal membeli? Apakah ada keluhan terkait kepercayaan di platform? Fenomena negatif atau gap antara kondisi ideal dan kondisi nyata ini yang perlu dimunculkan lebih kuat."),

      para("Kedua, soal urgensi. Tidak ada satu paragraf pun yang secara langsung mengatakan: 'inilah mengapa penelitian ini penting untuk dilakukan, terutama di konteks Lhokseumawe dan pada tahun ini'. Pembaca yang kritis bisa saja mempertanyakan: kenapa harus Lhokseumawe? Apa keunikan kota ini dibanding kota lain? Sebaiknya ada penjelasan singkat tentang kondisi masyarakat Lhokseumawe sebagai pengguna e-commerce."),

      para("Ketiga, soal research gap. Gap penelitian sudah disinggung lewat 'hasil berbeda' dari penelitian sebelumnya. Misalnya, ada peneliti yang bilang product description berpengaruh signifikan, tapi ada juga yang bilang tidak. Ini sebenarnya sudah merupakan bentuk gap, tapi sayangnya tidak dirangkum secara eksplisit. Idealnya ada satu paragraf atau kalimat yang langsung menyatakan: 'Masih terdapat inkonsistensi hasil penelitian tentang variabel X, sehingga penelitian ini hadir untuk menjawab gap tersebut, khususnya di konteks marketplace Shopee di Kota Lhokseumawe'."),

      para("Keempat, soal survei awal. Survei awal hanya dilakukan pada 17 mahasiswa. Jumlah ini terlalu kecil untuk dijadikan justifikasi kuat. Selain itu, kata 'mahasiswa' dalam survei ini berbeda dengan 'masyarakat Kota Lhokseumawe' yang menjadi populasi dalam Bab 3. Ini sedikit tidak konsisten."),

      emptyLine(),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: "D.2 Rumusan Masalah", bold: true, size: 26, font: "Times New Roman" })],
        spacing: { before: 200, after: 120 }
      }),

      para("Rumusan masalah sudah ditulis dengan format yang benar — menggunakan kalimat tanya dan disesuaikan satu per satu dengan masing-masing variabel independen. Ada 5 rumusan masalah untuk 5 variabel, ini sudah pas dan konsisten."),

      para("Yang perlu dicatat: tidak ada rumusan masalah yang menanyakan pengaruh simultan (secara bersama-sama) semua variabel terhadap purchase intention. Padahal kalau nanti ada uji F (uji simultan) dalam Bab 3, maka harusnya ada rumusan masalah yang sesuai. Sebaliknya, kalau memang tidak mau uji simultan, maka uji F di Bab 3 sebaiknya juga ditiadakan — tapi masalahnya Bab 3 tidak punya uji F sama sekali, jadi ini sebenarnya konsisten, meski tetap perlu dipertimbangkan apakah panduan mengharuskan uji F."),

      emptyLine(),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: "D.3 Tujuan dan Manfaat Penelitian", bold: true, size: 26, font: "Times New Roman" })],
        spacing: { before: 200, after: 120 }
      }),

      para("Tujuan penelitian sudah sesuai dengan rumusan masalah — ada 5 tujuan untuk 5 rumusan masalah. Ini sudah benar."),

      para("Manfaat penelitian dibagi dua: teoritis dan praktis. Untuk manfaat praktis sudah cukup detail (bagi perusahaan, bagi Shopee, bagi pelanggan, bagi peneliti selanjutnya). Untuk manfaat teoritis, kalimatnya terlalu umum dan lebih terlihat sebagai formalitas. Sebaiknya disebutkan secara spesifik: teori apa yang diperkuat atau dikembangkan dari penelitian ini?"),

      emptyLine(),

      // ============================================================
      // BAGIAN 5: ANALISIS BAB 2
      // ============================================================
      new Paragraph({ children: [new PageBreak()] }),

      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "E. ANALISIS MENDALAM BAB 2 — KAJIAN PUSTAKA", bold: true, size: 28, font: "Times New Roman" })],
        spacing: { before: 240, after: 180 }
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: "E.1 Landasan Teori", bold: true, size: 26, font: "Times New Roman" })],
        spacing: { before: 200, after: 120 }
      }),

      para("Secara keseluruhan, kajian teori di Bab 2 sudah cukup panjang dan cukup lengkap. Setiap variabel dibahas mulai dari pengertian, teori dari beberapa ahli, sampai indikator-indikatornya. Ini sudah sesuai dengan format yang umum diminta."),

      para("Namun ada satu hal yang cukup penting yang hilang: tidak ada grand theory atau teori utama yang menjadi payung dari keseluruhan penelitian ini. Untuk penelitian tentang perilaku konsumen dan niat beli di platform online, biasanya ada teori besar yang mendasari seperti:"),

      bullet("Theory of Planned Behavior (TPB) dari Ajzen — menjelaskan hubungan antara keyakinan, sikap, dan niat berperilaku"),
      bullet("Technology Acceptance Model (TAM) — relevan untuk konteks belanja online"),
      bullet("Consumer Involvement Theory — sangat relevan karena dua dari lima variabel adalah jenis involvement"),
      bullet("Information Processing Theory — relevan untuk variabel product description dan awareness"),

      para("Ketiadaan grand theory ini bisa menjadi celah yang ditanyakan penguji saat sidang."),

      emptyLine(),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: "E.2 Pembahasan Manajemen Pemasaran dan Pemasaran Digital", bold: true, size: 26, font: "Times New Roman" })],
        spacing: { before: 200, after: 120 }
      }),

      para("Bab 2 dibuka dengan pembahasan cukup panjang tentang manajemen pemasaran (Sub-bab 2.1) dan pemasaran digital (Sub-bab 2.2). Keduanya memang relevan sebagai konteks, tapi jujur saja — porsinya terlalu besar. Hampir 10 halaman dihabiskan untuk membahas konsep-konsep umum yang tidak langsung terhubung ke variabel penelitian."),

      para("Idealnya, manajemen pemasaran dan pemasaran digital cukup dibahas dalam 1-2 halaman sebagai pengantar saja, kemudian langsung masuk ke variabel-variabel penelitian yang spesifik."),

      emptyLine(),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: "E.3 Indikator Variabel: Sumber yang Dominan", bold: true, size: 26, font: "Times New Roman" })],
        spacing: { before: 200, after: 120 }
      }),

      para("Semua variabel dalam penelitian ini menggunakan indikator dari satu sumber yang sama, yaitu Zhu et al. (2019). Hal ini perlu diwaspadai karena:"),

      bullet("Bisa mengesankan bahwa peneliti terlalu bergantung pada satu referensi saja"),
      bullet("Dalam penelitian sebelumnya (tabel 2.1), Zhu et al. (2019) juga menjadi salah satu penelitian pembanding — ini berarti sumber yang sama digunakan baik sebagai teori maupun sebagai penelitian terdahulu, yang agak rancu"),
      bullet("Sebaiknya indikator dicampur dari beberapa sumber agar lebih kaya dan terdukung"),

      para("Walaupun dalam kajian teori setiap variabel memang dikutip dari beberapa ahli berbeda, tapi ketika menentukan indikator yang dipakai, semuanya berujung ke Zhu et al. (2019). Ini perlu dijelaskan secara eksplisit mengapa indikator dari Zhu et al. (2019) yang dipilih dibanding sumber lain."),

      emptyLine(),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: "E.4 Penelitian Terdahulu", bold: true, size: 26, font: "Times New Roman" })],
        spacing: { before: 200, after: 120 }
      }),

      para("Penelitian terdahulu sudah ada dalam format tabel (Tabel 2.1) dengan 6 penelitian. Format tabelnya sudah baik: ada nama peneliti, judul, perbedaan, dan hasil penelitian. Namun ada beberapa catatan:"),

      bullet("6 penelitian terdahulu terasa sedikit. Umumnya, skripsi FEB disarankan minimal 8-10 penelitian terdahulu"),
      bullet("Semua penelitian terdahulu menggunakan variabel yang sangat mirip dan berasal dari konteks e-commerce atau marketplace, tidak ada penelitian dari konteks yang berbeda yang kemudian diadaptasi"),
      bullet("Penelitian terdahulu yang menggunakan data dari Indonesia sangat sedikit — hampir semuanya adalah penelitian internasional. Padahal konteks perilaku konsumen Indonesia bisa sangat berbeda"),
      bullet("Zhu et al. (2019) muncul sebagai sumber teori DAN sebagai penelitian terdahulu — ini bisa menimbulkan pertanyaan dari penguji"),

      emptyLine(),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: "E.5 Kerangka Konseptual dan Hipotesis", bold: true, size: 26, font: "Times New Roman" })],
        spacing: { before: 200, after: 120 }
      }),

      para("Kerangka konseptual sudah ada dan bentuknya sudah benar — menggambarkan hubungan antara variabel X1 sampai X5 terhadap Y dengan tanda panah. Penjelasan di sub-bab 2.10 juga sudah menjelaskan alasan tiap hubungan didukung oleh teori dan penelitian terdahulu."),

      para("Hipotesis sudah ada 5 hipotesis parsial (H1 sampai H5) yang semuanya berbunyi 'berpengaruh signifikan'. Ini sudah konsisten dengan rumusan masalah. Namun yang perlu diperhatikan:"),

      bullet("Tidak ada hipotesis simultan (Ha: secara bersama-sama semua variabel berpengaruh terhadap Y)"),
      bullet("Hipotesis hanya menyebutkan 'berpengaruh signifikan' tanpa menyebut arah pengaruh (positif atau negatif) — padahal dalam kajian teori jelas disebutkan bahwa semuanya diharapkan berpengaruh positif. Akan lebih kuat jika hipotesis juga menyebutkan arah: 'berpengaruh positif dan signifikan'"),

      emptyLine(),

      // ============================================================
      // BAGIAN 6: ANALISIS BAB 3
      // ============================================================
      new Paragraph({ children: [new PageBreak()] }),

      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "F. ANALISIS MENDALAM BAB 3 — METODE PENELITIAN", bold: true, size: 28, font: "Times New Roman" })],
        spacing: { before: 240, after: 180 }
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: "F.1 Desain Penelitian", bold: true, size: 26, font: "Times New Roman" })],
        spacing: { before: 200, after: 120 }
      }),

      para("Bab 3 tidak secara eksplisit menyebutkan desain atau jenis penelitian. Apakah ini penelitian kuantitatif? Deskriptif? Asosiatif/kausal? Biasanya di awal Bab 3 ada satu sub-bab atau minimal satu paragraf yang menyebutkan jenis penelitian yang digunakan. Ini terlihat kurang dan sebaiknya ditambahkan."),

      para("Dari isi yang ada, penelitian ini jelas bersifat kuantitatif dengan pendekatan asosiatif/kausal (mencari pengaruh). Ini perlu dinyatakan secara eksplisit."),

      emptyLine(),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: "F.2 Populasi dan Sampel", bold: true, size: 26, font: "Times New Roman" })],
        spacing: { before: 200, after: 120 }
      }),

      para("Populasi didefinisikan sebagai 'seluruh masyarakat Kota Lhokseumawe yang menggunakan Aplikasi Shopee'. Ini tidak bermasalah, hanya saja populasinya tidak diketahui jumlahnya — maka rumus Lemeshow sudah tepat digunakan."),

      para("Sampel sebesar 96 responden sudah cukup masuk akal untuk penelitian kuantitatif dengan 5 variabel independen. Purposive sampling juga sudah tepat mengingat ada kriteria khusus (pengguna Shopee di Kota Lhokseumawe)."),

      para("Catatan: Kriteria sampel yang disebutkan hanya 2 (masyarakat Lhokseumawe + pengguna Shopee). Sebaiknya ada kriteria tambahan yang lebih spesifik, misalnya: minimal berapa kali pernah bertransaksi di Shopee, rentang usia, dll. Ini akan memperkuat validitas sampling."),

      para("Catatan kedua: Di latar belakang, survei awal dilakukan pada 'mahasiswa', tapi di Bab 3, populasinya adalah 'masyarakat' secara umum. Ini inkonsistensi kecil yang perlu diselaraskan."),

      emptyLine(),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: "F.3 Definisi Operasional Variabel", bold: true, size: 26, font: "Times New Roman" })],
        spacing: { before: 200, after: 120 }
      }),

      para("Tabel definisi operasional variabel sudah cukup lengkap — ada variabel, definisi, indikator, dan skala. Format ini sudah standar dan sudah benar."),

      para("Satu hal yang perlu diperhatikan: indikator untuk semua variabel berasal dari Zhu et al. (2019), tapi di kolom kuesioner (Lampiran 1), untuk variabel Situational Involvement, sumber yang tercantum adalah 'Croitoru et al. (2024)'. Ini berbeda dengan yang disebutkan di Bab 2 dan Bab 3 yang menggunakan Zhu et al. (2019). Ini adalah inkonsistensi yang perlu diperbaiki — pilih satu sumber yang konsisten di seluruh dokumen."),

      emptyLine(),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: "F.4 Metode Analisis Data", bold: true, size: 26, font: "Times New Roman" })],
        spacing: { before: 200, after: 120 }
      }),

      para("Regresi linier berganda sudah tepat untuk penelitian ini mengingat ada lebih dari satu variabel independen dan variabel dependennya bersifat kontinu (skala Likert yang dijumlahkan)."),

      para("Namun ada beberapa hal yang perlu dicek:"),

      bullet("Tidak ada uji F (uji simultan). Dalam regresi berganda, uji F hampir selalu diwajibkan karena menguji apakah model secara keseluruhan fit. Ketiadaan ini bisa jadi temuan penguji"),
      bullet("Uji R2 ada, tapi tidak disebutkan apakah menggunakan adjusted R2 atau tidak. Untuk regresi berganda, adjusted R2 lebih tepat"),
      bullet("Tidak ada penjelasan tentang bagaimana data kuesioner diolah sebelum dianalisis (apakah dijumlahkan, dirata-ratakan, dll)"),
      bullet("Tidak ada uji linearitas, padahal asumsi regresi linier berganda salah satunya adalah bahwa hubungan antar variabel bersifat linier"),

      emptyLine(),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: "F.5 Uji Asumsi Klasik", bold: true, size: 26, font: "Times New Roman" })],
        spacing: { before: 200, after: 120 }
      }),

      para("Tiga uji asumsi klasik yang disebutkan (normalitas, multikolinearitas, heteroskedastisitas) sudah cukup standar. Penjelasan cara dan dasar pengambilan keputusannya juga sudah ada. Secara umum bagian ini sudah oke."),

      para("Tambahan yang disarankan: uji autokorelasi (Durbin-Watson) biasanya juga termasuk dalam asumsi klasik untuk data regresi, meskipun umumnya lebih relevan untuk data time-series. Beberapa panduan FEB tetap mewajibkannya."),

      emptyLine(),

      // ============================================================
      // BAGIAN 7: KETERHUBUNGAN ANTAR BAB
      // ============================================================
      new Paragraph({ children: [new PageBreak()] }),

      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "G. KETERHUBUNGAN ANTAR BAB 1, 2, DAN 3", bold: true, size: 28, font: "Times New Roman" })],
        spacing: { before: 240, after: 180 }
      }),

      para("Ini adalah bagian yang sering diabaikan tapi sangat penting. Skripsi yang baik bukan hanya memiliki bab-bab yang isinya benar, tapi semua babnya harus 'ngobrol' satu sama lain — saling mendukung dan konsisten."),

      emptyLine(),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2000, 2000, 1200, 4160],
        rows: [
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: "1F4E79", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "DARI BAB", bold: true, size: 20, font: "Times New Roman", color: "FFFFFF" })] })] }),
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: "1F4E79", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "KE BAB", bold: true, size: 20, font: "Times New Roman", color: "FFFFFF" })] })] }),
              new TableCell({ borders, width: { size: 1200, type: WidthType.DXA }, shading: { fill: "1F4E79", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "STATUS", bold: true, size: 20, font: "Times New Roman", color: "FFFFFF" })] })] }),
              new TableCell({ borders, width: { size: 4160, type: WidthType.DXA }, shading: { fill: "1F4E79", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "PENJELASAN", bold: true, size: 20, font: "Times New Roman", color: "FFFFFF" })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Bab 1: Rumusan Masalah", size: 20, font: "Times New Roman" })] })] }),
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Bab 2: Hipotesis", size: 20, font: "Times New Roman" })] })] }),
              new TableCell({ borders, width: { size: 1200, type: WidthType.DXA }, shading: { fill: "C8F0D8", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "SESUAI", bold: true, size: 20, font: "Times New Roman", color: "1D6A39" })] })] }),
              new TableCell({ borders, width: { size: 4160, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "5 rumusan masalah selaras dengan 5 hipotesis yang dirumuskan", size: 20, font: "Times New Roman" })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Bab 1: Variabel penelitian", size: 20, font: "Times New Roman" })] })] }),
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Bab 2: Kajian teori", size: 20, font: "Times New Roman" })] })] }),
              new TableCell({ borders, width: { size: 1200, type: WidthType.DXA }, shading: { fill: "C8F0D8", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "SESUAI", bold: true, size: 20, font: "Times New Roman", color: "1D6A39" })] })] }),
              new TableCell({ borders, width: { size: 4160, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Semua variabel yang disebut di Bab 1 dibahas teorinya di Bab 2", size: 20, font: "Times New Roman" })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Bab 2: Indikator variabel", size: 20, font: "Times New Roman" })] })] }),
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Bab 3: Def. Operasional & Kuesioner", size: 20, font: "Times New Roman" })] })] }),
              new TableCell({ borders, width: { size: 1200, type: WidthType.DXA }, shading: { fill: "FDEBD0", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "KURANG", bold: true, size: 20, font: "Times New Roman", color: "7B3F00" })] })] }),
              new TableCell({ borders, width: { size: 4160, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Ada inkonsistensi sumber indikator Situational Involvement: Bab 2 & 3 pakai Zhu et al. (2019), tapi kuesioner tulis Croitoru et al. (2024)", size: 20, font: "Times New Roman" })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Bab 1: Survei pada mahasiswa", size: 20, font: "Times New Roman" })] })] }),
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Bab 3: Populasi = masyarakat", size: 20, font: "Times New Roman" })] })] }),
              new TableCell({ borders, width: { size: 1200, type: WidthType.DXA }, shading: { fill: "FDEBD0", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "KURANG", bold: true, size: 20, font: "Times New Roman", color: "7B3F00" })] })] }),
              new TableCell({ borders, width: { size: 4160, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Survei awal menyasar mahasiswa, tapi penelitian ditujukan untuk masyarakat umum. Perlu diselaraskan", size: 20, font: "Times New Roman" })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Bab 2: Hipotesis (hanya parsial)", size: 20, font: "Times New Roman" })] })] }),
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Bab 3: Uji hipotesis (hanya uji t)", size: 20, font: "Times New Roman" })] })] }),
              new TableCell({ borders, width: { size: 1200, type: WidthType.DXA }, shading: { fill: "C8F0D8", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "SESUAI", bold: true, size: 20, font: "Times New Roman", color: "1D6A39" })] })] }),
              new TableCell({ borders, width: { size: 4160, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Konsisten: tidak ada hipotesis simultan, tidak ada uji F — meski ini perlu dikonfirmasi ke panduan apakah uji F wajib", size: 20, font: "Times New Roman" })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Bab 1: Tujuan penelitian", size: 20, font: "Times New Roman" })] })] }),
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Bab 3: Metode yang dipilih", size: 20, font: "Times New Roman" })] })] }),
              new TableCell({ borders, width: { size: 1200, type: WidthType.DXA }, shading: { fill: "C8F0D8", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "SESUAI", bold: true, size: 20, font: "Times New Roman", color: "1D6A39" })] })] }),
              new TableCell({ borders, width: { size: 4160, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Tujuan menguji pengaruh variabel X terhadap Y sudah didukung metode regresi yang dipilih di Bab 3", size: 20, font: "Times New Roman" })] })] })
            ]
          })
        ]
      }),

      emptyLine(),

      // ============================================================
      // BAGIAN 8: DAFTAR TEMUAN DAN SARAN PERBAIKAN
      // ============================================================
      new Paragraph({ children: [new PageBreak()] }),

      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "H. DAFTAR LENGKAP TEMUAN DAN SARAN PERBAIKAN", bold: true, size: 28, font: "Times New Roman" })],
        spacing: { before: 240, after: 180 }
      }),

      para("Bagian ini merangkum semua temuan yang perlu diperbaiki, diurutkan dari yang paling krusial (bisa ditanya penguji) ke yang bersifat penyempurnaan. Setiap temuan dilengkapi saran konkret apa yang perlu dilakukan."),

      emptyLine(),

      paraBold("TEMUAN KRUSIAL (Harus Diperbaiki)"),
      emptyLine(),

      paraBold("1. Tidak ada grand theory / teori utama (Bab 2)"),
      para("Apa masalahnya: Penelitian ini tidak memiliki teori payung yang menjelaskan mengapa semua variabel ini (product description, awareness, involvement, trust) bisa berhubungan dengan purchase intention. Penguji bisa bertanya: 'Teori apa yang mendasari penelitian ini secara keseluruhan?'"),
      para("Saran perbaikan: Tambahkan sub-bab baru di awal Bab 2 yang membahas grand theory, misalnya Consumer Involvement Theory atau Theory of Planned Behavior. Jelaskan bagaimana teori tersebut menghubungkan semua variabel yang diteliti."),

      emptyLine(),
      paraBold("2. Fenomena masalah yang lemah (Bab 1)"),
      para("Apa masalahnya: Latar belakang lebih banyak membahas kepopuleran Shopee daripada fenomena masalah yang membutuhkan solusi melalui penelitian. Penguji bisa bertanya: 'Apa masalah nyatanya?'"),
      para("Saran perbaikan: Tambahkan fakta atau data tentang keluhan konsumen Shopee terkait deskripsi produk yang tidak jelas, atau rendahnya kepercayaan konsumen, atau masalah niat beli yang tidak terwujud. Bisa dari ulasan konsumen, berita, atau laporan lembaga."),

      emptyLine(),
      paraBold("3. Research gap tidak dirumuskan secara eksplisit (Bab 1)"),
      para("Apa masalahnya: Gap sudah tersirat dari perbedaan hasil penelitian, tapi tidak pernah dinyatakan secara langsung. Penguji bisa meminta peneliti menjelaskan apa yang membedakan penelitian ini dari yang sebelumnya."),
      para("Saran perbaikan: Tambahkan 1-2 paragraf di latar belakang yang secara eksplisit merumuskan gap: 'Meskipun banyak penelitian sudah membahas tentang X, Y, Z terhadap purchase intention, namun hasil penelitian masih tidak konsisten. Selain itu, belum ada penelitian yang meneliti kelima variabel ini secara bersama-sama dalam konteks marketplace Shopee di Kota Lhokseumawe. Oleh karena itu...'"),

      emptyLine(),
      paraBold("4. Tidak ada uji F / uji simultan (Bab 3)"),
      para("Apa masalahnya: Regresi berganda hampir selalu disertai uji simultan (uji F) untuk menguji kelayakan model. Tidak adanya uji F bisa menjadi pertanyaan penguji, terutama jika panduan FEB mewajibkannya."),
      para("Saran perbaikan: Tambahkan sub-bab 3.9.3 tentang Uji F (Simultan) dan tambahkan juga rumusan masalah serta hipotesis simultan yang sesuai di Bab 1 dan 2."),

      emptyLine(),
      paraBold("5. Inkonsistensi sumber indikator Situational Involvement"),
      para("Apa masalahnya: Di Bab 2 dan Bab 3, sumber indikator Situational Involvement adalah Zhu et al. (2019). Tapi di kuesioner (Lampiran 1), sumbernya tertulis 'Croitoru et al. (2024)'. Ini kontradiksi yang jelas dan akan langsung terlihat penguji."),
      para("Saran perbaikan: Pilih satu sumber yang konsisten dan perbaiki di semua tempat yang menyebutkan sumber indikator Situational Involvement."),

      emptyLine(),
      paraBold("TEMUAN MINOR (Disarankan untuk Diperbaiki)"),
      emptyLine(),

      paraBold("6. Tidak ada sub-bab jenis / desain penelitian (Bab 3)"),
      para("Saran: Tambahkan sub-bab 3.1 atau minimal satu paragraf di awal Bab 3 yang menyatakan jenis penelitian (kuantitatif, asosiatif/kausal)."),

      emptyLine(),
      paraBold("7. Kriteria sampel terlalu umum (Bab 3)"),
      para("Saran: Tambahkan kriteria yang lebih spesifik, misalnya: pernah bertransaksi di Shopee minimal 1 kali dalam 3 bulan terakhir, atau berdomisili di Kota Lhokseumawe."),

      emptyLine(),
      paraBold("8. Hipotesis tidak menyebut arah pengaruh (Bab 2)"),
      para("Saran: Ubah dari 'berpengaruh signifikan' menjadi 'berpengaruh positif dan signifikan' agar konsisten dengan kajian teori yang sudah menunjukkan arah positif."),

      emptyLine(),
      paraBold("9. Penelitian terdahulu kurang (Bab 2)"),
      para("Saran: Tambahkan minimal 3-4 penelitian terdahulu lagi, terutama yang berasal dari konteks Indonesia agar lebih relevan."),

      emptyLine(),
      paraBold("10. Manfaat teoritis terlalu umum (Bab 1)"),
      para("Saran: Sebutkan secara spesifik teori apa yang dikembangkan atau diperkuat oleh penelitian ini."),

      emptyLine(),
      paraBold("11. Porsi manajemen pemasaran dan digital marketing terlalu panjang (Bab 2)"),
      para("Saran: Pangkas kedua sub-bab ini menjadi lebih ringkas (sekitar 1 halaman masing-masing) dan alihkan ruang untuk memperkuat kajian teori variabel utama."),

      emptyLine(),
      paraBold("12. Tidak ada uji autokorelasi (Bab 3)"),
      para("Saran: Pertimbangkan untuk menambahkan uji Durbin-Watson, tergantung apakah panduan FEB mewajibkannya."),

      emptyLine(),

      // ============================================================
      // BAGIAN 9: PENILAIAN KESELURUHAN
      // ============================================================
      new Paragraph({ children: [new PageBreak()] }),

      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "I. PENILAIAN KESELURUHAN", bold: true, size: 28, font: "Times New Roman" })],
        spacing: { before: 240, after: 180 }
      }),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3000, 2000, 4360],
        rows: [
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, shading: { fill: "1F4E79", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "ASPEK PENILAIAN", bold: true, size: 22, font: "Times New Roman", color: "FFFFFF" })] })] }),
              new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: "1F4E79", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NILAI", bold: true, size: 22, font: "Times New Roman", color: "FFFFFF" })] })] }),
              new TableCell({ borders, width: { size: 4360, type: WidthType.DXA }, shading: { fill: "1F4E79", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "KOMENTAR", bold: true, size: 22, font: "Times New Roman", color: "FFFFFF" })] })] })
            ]
          }),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Kelengkapan Struktur Proposal", size: 22, font: "Times New Roman" })] })] }),
            new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: "C8F0D8", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "BAIK (78/100)", bold: true, size: 22, font: "Times New Roman", color: "1D6A39" })] })] }),
            new TableCell({ borders, width: { size: 4360, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Semua bagian utama sudah ada, tapi ada yang perlu dilengkapi", size: 22, font: "Times New Roman" })] })] })
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Kekuatan Latar Belakang", size: 22, font: "Times New Roman" })] })] }),
            new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: "FDEBD0", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "CUKUP (62/100)", bold: true, size: 22, font: "Times New Roman", color: "7B3F00" })] })] }),
            new TableCell({ borders, width: { size: 4360, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Fenomena dan gap perlu diperkuat lebih lanjut", size: 22, font: "Times New Roman" })] })] })
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Kajian Teori (Bab 2)", size: 22, font: "Times New Roman" })] })] }),
            new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: "C8F0D8", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "BAIK (72/100)", bold: true, size: 22, font: "Times New Roman", color: "1D6A39" })] })] }),
            new TableCell({ borders, width: { size: 4360, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Lengkap tapi perlu grand theory dan penelitian terdahulu lebih banyak", size: 22, font: "Times New Roman" })] })] })
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Metode Penelitian (Bab 3)", size: 22, font: "Times New Roman" })] })] }),
            new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: "C8F0D8", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "BAIK (74/100)", bold: true, size: 22, font: "Times New Roman", color: "1D6A39" })] })] }),
            new TableCell({ borders, width: { size: 4360, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Sudah lengkap tapi perlu uji F dan jenis penelitian", size: 22, font: "Times New Roman" })] })] })
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Konsistensi Antar Bab", size: 22, font: "Times New Roman" })] })] }),
            new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: "FDEBD0", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "CUKUP (65/100)", bold: true, size: 22, font: "Times New Roman", color: "7B3F00" })] })] }),
            new TableCell({ borders, width: { size: 4360, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Ada inkonsistensi pada sumber indikator dan target responden", size: 22, font: "Times New Roman" })] })] })
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, shading: { fill: "D6E4F0", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "NILAI RATA-RATA KESELURUHAN", bold: true, size: 22, font: "Times New Roman" })] })] }),
            new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: "D6E4F0", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "70/100", bold: true, size: 22, font: "Times New Roman" })] })] }),
            new TableCell({ borders, width: { size: 4360, type: WidthType.DXA }, shading: { fill: "D6E4F0", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Layak lanjut dengan perbaikan pada poin-poin krusial", bold: true, size: 22, font: "Times New Roman" })] })] })
          ]})
        ]
      }),

      emptyLine(),

      para("Secara keseluruhan, proposal skripsi Putri Shalsabila sudah menunjukkan kerja yang cukup serius. Topiknya relevan, variabelnya jelas, dan metodologinya sudah mengarah ke yang benar. Yang perlu dilakukan sekarang adalah memperkuat bagian-bagian yang masih lemah — terutama latar belakang, grand theory, dan konsistensi antar bagian."),

      para("Dengan perbaikan pada 12 poin yang sudah diidentifikasi di atas — terutama 5 poin krusial — proposal ini akan jauh lebih kuat dan lebih siap untuk dipertahankan di hadapan penguji."),

      emptyLine(),

      // ============================================================
      // PENUTUP
      // ============================================================
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "J. CATATAN PENUTUP", bold: true, size: 28, font: "Times New Roman" })],
        spacing: { before: 240, after: 180 }
      }),

      para("Dokumen telaah ini disusun dengan pendekatan yang jujur dan konstruktif. Tujuannya bukan untuk menghakimi proposal yang sudah ditulis, tapi untuk memberikan gambaran yang jelas tentang apa yang sudah baik dan apa yang masih perlu ditingkatkan sebelum proposal ini diajukan ke penguji."),

      para("Penting juga untuk diingat: telaah ini dilakukan tanpa bisa mengakses file panduan resmi FEB Unimal secara langsung (tautan Google Drive memerlukan login). Oleh karena itu, ada beberapa poin — khususnya tentang wajib tidaknya uji F dan jumlah minimum penelitian terdahulu — yang perlu dikonfirmasi langsung ke panduan resmi yang berlaku di Universitas Malikussaleh."),

      para("Semoga telaah ini bermanfaat dan bisa membantu proses penyempurnaan proposal skripsi sebelum sidang. Tetap semangat!"),

      emptyLine(),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 240, after: 60 },
        children: [new TextRun({ text: "Ditelaah oleh,", size: 24, font: "Times New Roman" })]
      }),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: "Khoirul Anwar", bold: true, size: 24, font: "Times New Roman" })]
      }),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: "Mei 2026", size: 24, font: "Times New Roman" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("./hasil.docx", buffer);
  console.log("DONE");
});
