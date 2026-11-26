const SPREADSHEET_ID = "1JvSGCzfrG7RaCGnFpkZzD7a9gw7UW2rWF9JLTs7YkfE"; 

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard & Pendaftaran Pelayanan Sore')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getNextNoUrutInternal(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 5) return 1; 

    const timeZone = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), timeZone, "dd/MM/yyyy");

    const lastRowValues = sheet.getRange(lastRow, 1, 1, 2).getDisplayValues()[0];
    const lastNo = parseInt(lastRowValues[0], 10) || 0;
    let lastTanggal = String(lastRowValues[1] || "").trim();

    if (lastTanggal.includes("-")) {
      const parts = lastTanggal.split("-");
      if (parts.length === 3 && parts[0].length === 4) {
        lastTanggal = `${parts[2].padStart(2, '0')}/${parts[1].padStart(2, '0')}/${parts[0]}`;
      }
    }

    return (lastTanggal === todayStr) ? (lastNo + 1) : 1;
  } catch (err) {
    Logger.log("Error getNextNoUrutInternal: " + err.toString());
    return 1;
  }
}

function getNextNoUrut() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("PELAYANAN SORE");
    if (!sheet) return 1;
    return getNextNoUrutInternal(sheet);
  } catch (err) {
    Logger.log("Error getNextNoUrut: " + err.toString());
    return 1;
  }
}

function cleanText(text) {
  if (!text) return "";
  let str = String(text).trim();

  if (str.startsWith('"') && str.endsWith('"')) {
    str = str.substring(1, str.length - 1);
  }

  return str
    .replace(/[\r\n]+/g, ", ") 
    .replace(/\s+/g, " ")    
    .replace(/,\s*,/g, ",")    
    .toUpperCase()
    .trim();
}

function simpanPendaftaran(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { status: "error", message: "Sistem sedang sibuk, gagal menyimpan data. Silakan coba lagi." };
  }

  try {
    if (!data || !data.namaPasien || !data.tanggal || !data.alamat) {
      return { status: "error", message: "Data tidak lengkap! Mohon isi semua kolom yang wajib." };
    }

    const targetSheetName = "PELAYANAN SORE";
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(targetSheetName);

    if (!sheet) {
      throw new Error(`Sheet '${targetSheetName}' tidak ditemukan!`);
    }

    const noUrutValid = getNextNoUrutInternal(sheet);
    const newRow = sheet.getLastRow() < 5 ? 5 : sheet.getLastRow() + 1;

    const rowData = [
      noUrutValid,                                         // A: NO
      "'" + String(data.tanggal).trim(),                   // B: TANGGAL
      cleanText(data.namaPasien),                          // C: NAMA PASIEN
      String(data.jenisKelamin || "").trim(),              // D: JENIS KELAMIN
      "'" + String(data.tanggalLahirFormatted || "").trim(), // E: TANGGAL LAHIR
      String(data.usia || "").trim(),                      // F: USIA
      String(data.kategoriUsia || "").trim(),              // G: KATEGORI USIA
      cleanText(data.alamat),                              // H: ALAMAT
      String(data.jenisLayanan || "").trim(),              // I: JENIS LAYANAN / POLI
      String(data.penjamin || "").trim(),                  // J: PENJAMIN
      "", "", "", "", "", "", "", "", "",                  // K s/d S (Diisi oleh Dokter/Bidan)
      cleanText(data.keterangan)                           // T: KETERANGAN / CKG
    ];

    sheet.getRange(newRow, 1, 1, 20).setValues([rowData]);

    const rangeOPQ = sheet.getRange(newRow, 15, 1, 3);
    if (!rangeOPQ.isPartOfMerge()) {
      rangeOPQ.merge();
    }

    return { 
      status: "success", 
      message: `Data pendaftaran Pelayanan Sore berhasil disimpan!`,
      noUrut: noUrutValid 
    };

  } catch (err) {
    Logger.log("Error pada simpanPendaftaran: " + err.toString());
    return { status: "error", message: "Gagal menyimpan data: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

function getPasienHariIni() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("PELAYANAN SORE");
    const timeZone = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), timeZone, "dd/MM/yyyy");

    if (!sheet || sheet.getLastRow() < 5) return [];

    const data = sheet.getRange(5, 1, sheet.getLastRow() - 4, 20).getDisplayValues();
    const daftarPasien = [];

    data.forEach((row, index) => {
      let tgl = String(row[1] || "").trim();
      if (tgl.includes("-")) {
        const p = tgl.split("-");
        if (p.length === 3 && p[0].length === 4) tgl = `${p[2].padStart(2,'0')}/${p[1].padStart(2,'0')}/${p[0]}`;
      }

      if (tgl === todayStr) {
        daftarPasien.push({
          rowIndex: index + 5,
          no: row[0],
          nama: cleanText(row[2]),
          jk: row[3],
          usia: row[5],
          jenisLayanan: row[8] || "UMUM", 
          tindakanKat: row[12] || "",     
          uraianTindakan: cleanText(row[13]), 
          diagnosa: cleanText(row[14]),      
          rujukan: row[17] || "",         
          namaPemeriksa: row[18] || ""    
        });
      }
    });

    return daftarPasien;
  } catch (err) {
    Logger.log("Error getPasienHariIni: " + err.toString());
    return [];
  }
}

function simpanPemeriksaanDokter(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("PELAYANAN SORE");

    if (!sheet) throw new Error("Sheet PELAYANAN SORE tidak ditemukan!");
    const rowIndex = parseInt(data.rowIndex, 10);

    const tindakanKat = cleanText(data.tindakanKat);
    const uraianTindakan = cleanText(data.uraianTindakan);
    const diagnosaVal = cleanText(data.diagnosa);
    const rujukanVal = cleanText(data.rujukan);
    const namaPemeriksaVal = String(data.namaPemeriksa || "").trim();

    const valuesBatch = [
      [
        tindakanKat,
        uraianTindakan,
        diagnosaVal,
        "", 
        "", 
        rujukanVal,
        namaPemeriksaVal
      ]
    ];

    sheet.getRange(rowIndex, 13, 1, 7).setValues(valuesBatch);

    return { status: "success", message: "Data pemeriksaan medis berhasil disimpan!" };
  } catch (err) {
    Logger.log("Error simpanPemeriksaanDokter: " + err.toString());
    return { status: "error", message: "Gagal menyimpan data pemeriksaan: " + err.message };
  }
}

function getDashboardStats() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const timeZone = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), timeZone, "dd/MM/yyyy");

    const sheetSore = ss.getSheetByName("PELAYANAN SORE");

    let totalSore = 0;
    let totalLaki = 0;
    let totalPerempuan = 0;

    const wilayah = {
      "DALAM_WILAYAH": 0,
      "DALAM_KOTA_BANDUNG": 0,
      "LUAR_KOTA_BANDUNG": 0
    };

    const penjamin = { BPJS: 0, UMUM: 0, GRATIS: 0 };

    const layananMap = {
      "KLASTER 3 : LANSIA": 0,
      "KLASTER 3 : DEWASA": 0,
      "ANAK & REMAJA": 0,
      "BALITA <5 TAHUN": 0,
      "BAYI (0 - 6 BULAN)": 0,
      "GIGI": 0,
      "ANC": 0,
      "PNC": 0,
      "CATIN": 0,
      "KB": 0,
      "KONSELING": 0,
      "KESEHATAN REPRODUKSI": 0
    };

    let totalCkg = 0;
    const mapKasus = {};
    const mapTindakan = {}; 

    const kelurahanDalamWilayah = ["KEBON WARU", "KEBON GEDANG", "CIBANGKONG"];

    if (sheetSore && sheetSore.getLastRow() >= 5) {
      const dataSore = sheetSore.getRange(5, 1, sheetSore.getLastRow() - 4, 20).getDisplayValues();

      dataSore.forEach(row => {
        let tgl = String(row[1] || "").trim();
        if (tgl.includes("-")) {
          const p = tgl.split("-");
          if (p.length === 3 && p[0].length === 4) tgl = `${p[2].padStart(2,'0')}/${p[1].padStart(2,'0')}/${p[0]}`;
        }

        if (tgl === todayStr) {
          totalSore++;
          const jk = String(row[3] || "").trim().toUpperCase(); 
          const alamat = cleanText(row[7]); 
          const jenisLay = String(row[8] || "").trim().toUpperCase(); 
          const penjaminVal = String(row[9] || "").trim().toUpperCase(); 
          const katTindakan = String(row[12] || "").trim().toUpperCase(); 
          const detailTindakan = cleanText(row[13]);                     
          const diagnosaVal = cleanText(row[14]);                        
          const ketVal = cleanText(row[19]);                             

          const isLaki = (jk === "L" || jk.includes("LAKI"));
          if (isLaki) totalLaki++; else totalPerempuan++;

          if (kelurahanDalamWilayah.includes(alamat)) {
            wilayah["DALAM_WILAYAH"]++;
          } else if (alamat === "LUAR KOTA BANDUNG") {
            wilayah["LUAR_KOTA_BANDUNG"]++;
          } else {
            wilayah["DALAM_KOTA_BANDUNG"]++;
          }

          if (penjaminVal.includes("BPJS")) penjamin.BPJS++;
          else if (penjaminVal.includes("UMUM")) penjamin.UMUM++;
          else if (penjaminVal.includes("GRATIS")) penjamin.GRATIS++;

          if (jenisLay !== "") {
            Object.keys(layananMap).forEach(key => {
              if (jenisLay.includes(key) || key.includes(jenisLay)) {
                layananMap[key]++;
              }
            });
          }

          if (ketVal.includes("CKG")) totalCkg++;

          const namaTindakan = detailTindakan || (katTindakan === "TINDAKAN" ? "TINDAKAN MEDIS GENERAL" : "");
          if (namaTindakan) {
            const listTindakan = namaTindakan.split(/[\r\n,]+/);
            listTindakan.forEach(tItem => {
              const cleanedT = tItem.trim();
              if (cleanedT && cleanedT !== "NON TINDAKAN") {
                mapTindakan[cleanedT] = (mapTindakan[cleanedT] || 0) + 1;
              }
            });
          }

          if (diagnosaVal) {
            const listKasus = diagnosaVal.split(/[\r\n,]+/);
            listKasus.forEach(kItem => {
              const cleanedK = kItem.trim();
              if (cleanedK) {
                mapKasus[cleanedK] = (mapKasus[cleanedK] || 0) + 1;
              }
            });
          }
        }
      });
    }

    const sortedKasus = Object.keys(mapKasus)
      .map(k => ({ kasus: k, jumlah: mapKasus[k] }))
      .sort((a, b) => b.jumlah - a.jumlah)
      .slice(0, 3);

    const sortedTindakan = Object.keys(mapTindakan)
      .map(t => ({ tindakan: t, jumlah: mapTindakan[t] }))
      .sort((a, b) => b.jumlah - a.jumlah);

    return {
      totalSore: totalSore,
      totalLaki: totalLaki,
      totalPerempuan: totalPerempuan,
      wilayah: wilayah,
      penjamin: penjamin,
      layanan: layananMap,
      totalCkg: totalCkg,
      top3Kasus: sortedKasus,
      semuaTindakan: sortedTindakan
    };
  } catch (err) {
    Logger.log("Error getDashboardStats: " + err.toString());
    return {
      totalSore: 0, totalLaki: 0, totalPerempuan: 0,
      wilayah: { "DALAM_WILAYAH": 0, "DALAM_KOTA_BANDUNG": 0, "LUAR_KOTA_BANDUNG": 0 },
      penjamin: { BPJS: 0, UMUM: 0, GRATIS: 0 },
      layanan: {}, totalCkg: 0, top3Kasus: [], semuaTindakan: []
    };
  }
}