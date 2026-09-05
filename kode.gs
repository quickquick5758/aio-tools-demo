const MASTER_FILE_ID = '1vuHgFUcSOUzBonCF_w_e3j-sBYlQQy8o6clL4GmcuXQ';
const PROD_FILE_ID = '1HsWvRkojtcvBz333BDBVXBLELIPMrAhf4UR10C3rML8';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('AiO-Tools Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function verifikasiLogin(inputUsername, inputPassword) {
  try {
    const ss = SpreadsheetApp.openById(MASTER_FILE_ID);
    const sheet = ss.getSheetByName("MASTER_ID_ACCOUNT"); 
    if (!sheet) return { success: false, message: "Sheet MASTER_ID_ACCOUNT tidak ditemukan!" };
    
    const data = sheet.getDataRange().getValues();
    const cleanUser = String(inputUsername || '').trim();
    const cleanPass = String(inputPassword || '').trim();

    // Mulai dari baris ke-3 (indeks 2)
    for (let i = 2; i < data.length; i++) {
      // Paksa konversi ke String agar jika di Sheet berupa angka (misal PIN/NIK), tetap cocok
      const sheetUser = String(data[i][2] !== undefined ? data[i][2] : '').trim();
      const sheetPass = String(data[i][3] !== undefined ? data[i][3] : '').trim();

      if (sheetUser === cleanUser && sheetPass === cleanPass) {
        return { 
          success: true, 
          name: data[i][1],      // Nama User (Kolom B)
          role: data[i][4],      // Level/Role (Kolom E)
          unit: data[i][5]       // Divisi/Unit (Kolom F)
        };
      }
    }
    return { success: false, message: "Username atau Password salah!" };
  } catch (err) {
    return { success: false, message: "Error Server: " + err.message };
  }
}

function getMasterBarangData() {
  const ss = SpreadsheetApp.openById(MASTER_FILE_ID);
  const sheet = ss.getSheetByName("MASTER_BARANG");
  if (!sheet) return { success: false, message: "Sheet tidak ditemukan" };
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return { success: true, data: [] };
  const data = sheet.getRange(4, 1, lastRow - 3, 18).getDisplayValues(); // Diperbarui ke 18 kolom (Ada Supplier)
  return { success: true, data: data };
}

function getMasterOperatorData() {
  const ss = SpreadsheetApp.openById(MASTER_FILE_ID);
  const sheet = ss.getSheetByName("MASTER_OPERATOR");
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  // Ambil Kolom C sampai F (4 Kolom: Nama, Pekerjaan, Karu, Status)
  const data = sheet.getRange(2, 3, lastRow - 1, 4).getValues();
  return data
    .filter(row => String(row[3] || '').trim().toLowerCase() === 'aktif') // Hanya status Aktif
    .map(row => ({ nama: row[0], pekerjaan: row[1], karu: row[2], status: row[3] }));
}

function getSaldoMutasiData() {
  try {
    const ss = SpreadsheetApp.openById(PROD_FILE_ID);
    const sheet = ss.getSheetByName("DATA_MUTASI");
    if (!sheet) return { success: false, data: [] };
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return { success: true, data: [] };
    // Ambil data mulai baris ke-3 dari Kolom A sampai M (13 Kolom)
    const data = sheet.getRange(3, 1, lastRow - 2, 13).getValues();
    return { success: true, data: data };
  } catch (e) {
    return { success: false, message: e.message, data: [] };
  }
}

function getDashboardStats() {
  try {
    const ss = SpreadsheetApp.openById(PROD_FILE_ID);
    let qtyMasuk = 0, qtyProd = 0, qtyKirim = 0;
    let rpMasuk = 0, rpProd = 0, rpKirim = 0;
    let chartDataRaw = {}; 
    let activeOperators = {};

    const sheetDetail = ss.getSheetByName("DATA_PRODUKSI_DETAIL");
      if (sheetDetail) {
        const lastRow = sheetDetail.getLastRow();
        if (lastRow >= 3) {
          const detailData = sheetDetail.getRange(3, 1, lastRow - 2, 18).getValues();
          
          detailData.forEach(row => {
            let unit = String(row[0] || '').trim();
            let rawTgl = row[2];
            let tipe = String(row[3] || '').trim().toUpperCase();
            let pekerjaan = String(row[6] || '').trim().toLowerCase();
            let opNames = String(row[7] || '').trim();
            let supplier = String(row[11] || '').trim().toUpperCase();
            let qty = parseFloat(row[13]) || 0;
            
            let insentifRaw = String(row[16] || '0').split(',')[0].replace(/[^\d]/g, '');
            let insentif = parseFloat(insentifRaw) || 0;

          if (tipe === 'PRD') {
            rpProd += insentif; // Tetap mengakumulasi total rupiah dari semua pekerjaan
            
            // Logika Filter: Packing All Supplier + Jahit Tutup khusus PT. TIRAI PELANGI NUSANTARA
            let isPacking = pekerjaan.includes('packing');
            let isJahitTutupPelangi = pekerjaan.includes('jahit tutup') && supplier.includes('TIRAI PELANGI NUSANTARA');
            
            if (isPacking || isJahitTutupPelangi) {
                qtyProd += qty;
                
                if (rawTgl) {
                  let tglStr = "";
                  if (rawTgl instanceof Date) {
                    tglStr = Utilities.formatDate(rawTgl, "Asia/Jakarta", "yyyy-MM-dd");
                  } else {
                    let d = new Date(rawTgl);
                    if (!isNaN(d.getTime())) tglStr = Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd");
                  }
                  if (tglStr) {
                    if (!chartDataRaw[tglStr]) chartDataRaw[tglStr] = 0;
                    chartDataRaw[tglStr] += qty;
                  }
                }
            }
          }

          if (unit && opNames && opNames !== "-") {
            let ops = opNames.split(',');
            ops.forEach(op => {
              let cleanOp = op.trim();
              if (cleanOp) {
                if (!activeOperators[unit]) activeOperators[unit] = new Set();
                activeOperators[unit].add(cleanOp);
              }
            });
          }
        });
      }
    }

    const sheetStat = ss.getSheetByName("DATA_STATISTIK");
    if (sheetStat) {
      let qMasukRaw = String(sheetStat.getRange("C3").getDisplayValue()).split(',')[0];
      let qKirimRaw = String(sheetStat.getRange("E3").getDisplayValue()).split(',')[0];
      qtyMasuk = parseFloat(qMasukRaw.replace(/[^\d]/g, '')) || 0;
      qtyKirim = parseFloat(qKirimRaw.replace(/[^\d]/g, '')) || 0;
      
      let rpMasukRaw = String(sheetStat.getRange("L3").getDisplayValue()).split(',')[0];
      let rpKirimRaw = String(sheetStat.getRange("N3").getDisplayValue()).split(',')[0];
      rpMasuk = parseFloat(rpMasukRaw.replace(/[^\d]/g, '')) || 0;
      rpKirim = parseFloat(rpKirimRaw.replace(/[^\d]/g, '')) || 0;
    }
    
    let operatorStats = { total: 0, units: {} };
    for (let unit in activeOperators) {
      let count = activeOperators[unit].size;
      operatorStats.units[unit] = count;
      operatorStats.total += count;
    }

    return {
      success: true,
      qty: { masuk: qtyMasuk, prod: qtyProd, kirim: qtyKirim },
      rp: { masuk: rpMasuk, prod: rpProd, kirim: rpKirim },
      operators: operatorStats,
      chartData: chartDataRaw
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getDataLaporanLatest() {
  const ss = SpreadsheetApp.openById(PROD_FILE_ID);
  let sheet = ss.getSheetByName("DATA_PRODUKSI_GLOBAL");
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const startRow = Math.max(2, lastRow - 20);
  const numRows = lastRow - startRow + 1;
  const rawData = sheet.getRange(startRow, 1, numRows, 10).getValues();
  return rawData.map(r => {
    let tgl = r[2];
    if (tgl instanceof Date) {
      tgl = Utilities.formatDate(tgl, "Asia/Jakarta", "yyyy-MM-dd");
    } else if (typeof tgl === "number") {
      tgl = Utilities.formatDate(new Date((tgl - 25569) * 86400 * 1000), "Asia/Jakarta", "yyyy-MM-dd");
    } else if (typeof tgl === "string" && tgl.trim() !== "") {
      let d = new Date(tgl);
      if (!isNaN(d.getTime())) tgl = Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd");
    }
    r[2] = tgl;
    return r;
  });
}

function getDataLaporan() {
  const ss = SpreadsheetApp.openById(PROD_FILE_ID);
  let sheet = ss.getSheetByName("DATA_PRODUKSI_GLOBAL");
  if (!sheet) {
    sheet = ss.insertSheet("DATA_PRODUKSI_GLOBAL");
    sheet.appendRow(["Unit", "ID Laporan", "Tanggal", "Tipe", "Karu", "Shift", "Pekerjaan", "Operator", "Total Qty", "Status"]);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rawData = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  return rawData.map(r => {
    let tgl = r[2];
    if (tgl instanceof Date) {
      tgl = Utilities.formatDate(tgl, "Asia/Jakarta", "yyyy-MM-dd");
    } else if (typeof tgl === "number") {
      tgl = Utilities.formatDate(new Date((tgl - 25569) * 86400 * 1000), "Asia/Jakarta", "yyyy-MM-dd");
    } else if (typeof tgl === "string" && tgl.trim() !== "") {
      let d = new Date(tgl);
      if (!isNaN(d.getTime())) tgl = Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd");
    }
    r[2] = tgl;
    return r;
  });
}

function getDataProduksi() {
  const ss = SpreadsheetApp.openById(PROD_FILE_ID);

  // 1. Ambil Data Produksi Detail
  let sheetProd = ss.getSheetByName("DATA_PRODUKSI_DETAIL");
  if (!sheetProd) {
    sheetProd = ss.insertSheet("DATA_PRODUKSI_DETAIL");
    sheetProd.appendRow(["Unit", "ID Laporan", "Tanggal", "Tipe", "Karu", "Shift", "Pekerjaan", "Operator", "Kode", "Nama Barang", "Kategori", "Supplier", "Motif", "Qty", "Keterangan", "Harga Jatuh Anak", "Insentif", "Status"]);
  }
  let rawProd = [];
  const prodLastRow = sheetProd.getLastRow();
  if (prodLastRow >= 3) {
    rawProd = sheetProd.getRange(3, 1, prodLastRow - 2, 18).getValues();
  }

  // 2. Ambil Data BBK Detail & Gabungkan
  let sheetBbk = ss.getSheetByName("DATA_PENGECEKAN_BBK_DETAIL");
  let rawBbk = [];
  if (sheetBbk) {
    const bbkLastRow = sheetBbk.getLastRow();
    if (bbkLastRow >= 3) {
      let bData = sheetBbk.getRange(3, 1, bbkLastRow - 2, 16).getValues();
      rawBbk = bData.map(b => {
        // Pemetaan Array View Data (18 Kolom)
        return [
          b[0],       // Unit
          b[3],       // Nomor Dokumen (LPS-BB-...) -> Masuk sebagai ID Laporan
          b[1],       // Tanggal
          "BBK",      // Tipe (Hardcode "BBK")
          "-",        // Karu
          b[5],       // Shift
          b[2],       // Pekerjaan
          b[4],       // Nama Operator -> Masuk sebagai Operator
          b[9],       // Kode
          b[10],      // Nama barang
          b[11],      // Kategori
          b[6],       // Supplier
          b[12],      // Motif
          b[13],      // Qty
          b[14],      // Keterangan
          0,          // Harga Jatuh Anak di-set ke 0 (karena ini Bahan Masuk, bukan Jasa Produksi)
          0,          // Insentif di-set ke 0
          b[15]       // Status
        ];
      });
    }
  }

  // 3. Proses format Tanggal untuk Gabungan Produksi + BBK
  let combinedData = rawProd.concat(rawBbk);

  return combinedData.map(r => {
    let tgl = r[2];
    if (tgl instanceof Date) {
      tgl = Utilities.formatDate(tgl, "Asia/Jakarta", "yyyy-MM-dd");
    } else if (typeof tgl === "number") {
      tgl = Utilities.formatDate(new Date((tgl - 25569) * 86400 * 1000), "Asia/Jakarta", "yyyy-MM-dd");
    } else if (typeof tgl === "string" && tgl.trim() !== "") {
      let d = new Date(tgl);
      if (!isNaN(d.getTime())) tgl = Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd");
    }
    r[2] = tgl;
    return r;
  });
}

function getBbkGlobalData() {
  const ss = SpreadsheetApp.openById(PROD_FILE_ID);
  let sheet = ss.getSheetByName("DATA_PENGECEKAN_BBK_GLOBAL");
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return []; // Header ada di baris 2, data di baris 3
  
  // Kolom A-K berarti 11 kolom
  const rawData = sheet.getRange(3, 1, lastRow - 2, 11).getValues();
  return rawData.map(r => {
    let tgl = r[1]; // Index 1 adalah Tanggal
    if (tgl instanceof Date) {
      tgl = Utilities.formatDate(tgl, "Asia/Jakarta", "yyyy-MM-dd");
    } else if (typeof tgl === "number") {
      tgl = Utilities.formatDate(new Date((tgl - 25569) * 86400 * 1000), "Asia/Jakarta", "yyyy-MM-dd");
    } else if (typeof tgl === "string" && tgl.trim() !== "") {
      let d = new Date(tgl);
      if (!isNaN(d.getTime())) tgl = Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd");
    }
    r[1] = tgl;
    return r;
  });
}

function getBbkDetailData() {
  const ss = SpreadsheetApp.openById(PROD_FILE_ID);
  let sheet = ss.getSheetByName("DATA_PENGECEKAN_BBK_DETAIL");
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return []; // Header di baris 2
  
  // Kolom A - P (16 Kolom)
  const rawData = sheet.getRange(3, 1, lastRow - 2, 16).getValues();
  return rawData.map(r => {
    let tgl = r[1]; // Index 1 adalah Tanggal
    if (tgl instanceof Date) {
      tgl = Utilities.formatDate(tgl, "Asia/Jakarta", "yyyy-MM-dd");
    } else if (typeof tgl === "number") {
      tgl = Utilities.formatDate(new Date((tgl - 25569) * 86400 * 1000), "Asia/Jakarta", "yyyy-MM-dd");
    } else if (typeof tgl === "string" && tgl.trim() !== "") {
      let d = new Date(tgl);
      if (!isNaN(d.getTime())) tgl = Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd");
    }
    r[1] = tgl;
    return r;
  });
}

function saveProduksiToSheet(data) {
  try {
    const ss = SpreadsheetApp.openById(PROD_FILE_ID);
    
    let detailSheet = ss.getSheetByName("DATA_PRODUKSI_DETAIL");
    if (!detailSheet) {
      detailSheet = ss.insertSheet("DATA_PRODUKSI_DETAIL");
      detailSheet.appendRow(["Unit", "ID Laporan", "Tanggal", "Tipe", "Karu", "Shift", "Pekerjaan", "Operator", "Kode", "Nama Barang", "Kategori", "Supplier", "Motif", "Qty", "Keterangan", "Harga Jatuh Anak", "Insentif", "Status"]);
    }
    
    let globalSheet = ss.getSheetByName("DATA_PRODUKSI_GLOBAL");
    if (!globalSheet) {
      globalSheet = ss.insertSheet("DATA_PRODUKSI_GLOBAL");
      globalSheet.appendRow(["Unit", "ID Laporan", "Tanggal", "Tipe", "Karu", "Shift", "Pekerjaan", "Operator", "Total Qty", "Status"]);
    }

    let idLaporan = data.idLaporan;
    let isEdit = !!idLaporan; // Cek apakah ID sudah ada (Mode Edit)
    
    if (!isEdit) {
      // Generate ID Laporan otomatis hanya untuk dokumen BARU
      const dateObj = new Date(data.tanggal);
      const yy = String(dateObj.getFullYear()).slice(-2);
      const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
      const prefix = `LPS-${data.tipe}-${yy}${mm}-`;
      
      const lastRow = detailSheet.getLastRow();
      let runningNum = 1;
      if (lastRow >= 2) {
        const ids = detailSheet.getRange(2, 2, lastRow - 1, 1).getValues();
        let maxNum = 0;
        ids.forEach(row => {
          if (row[0] && row[0].startsWith(prefix)) {
            const num = parseInt(row[0].replace(prefix, '')) || 0;
            if (num > maxNum) maxNum = num;
          }
        });
        runningNum = maxNum + 1;
      }
      idLaporan = prefix + String(runningNum).padStart(6, '0');
    }
    
    const unit = data.unit || "CV JHB 1";
    // Gunakan Status dari frontend agar jika Author edit dokumen Approved, statusnya tetap Approved.
    const docStatus = data.status || "Submitted"; 

    let masterBarangData = [];
    try {
      const masterSs = SpreadsheetApp.openById(MASTER_FILE_ID);
      const masterSheet = masterSs.getSheetByName("MASTER_BARANG");
      if (masterSheet) {
        const masterLastRow = masterSheet.getLastRow();
        if (masterLastRow >= 4) {
          masterBarangData = masterSheet.getRange(4, 1, masterLastRow - 3, 18).getValues();
        }
      }
    } catch(e) {
      console.error("Gagal get MASTER_BARANG: " + e);
    }

    const mapPekerjaan = {
      "tutup besi": 11, "jahit bulatan": 12, "pengisian": 13,
      "jahit tutup": 14, "penyarungan": 15, "packing/press": 16, "ikat/ball": 17
    };
    
    const pekerjaanKey = String(data.pekerjaan).toLowerCase().trim();
    const colIndexHarga = mapPekerjaan[pekerjaanKey] !== undefined ? mapPekerjaan[pekerjaanKey] : -1;
    
    let totalQtyGlobal = 0;
    let newDetailRows = [];

    data.items.forEach(item => {
        let kode = "";
        let namaBarang = "";
        let inputMotif = "";
        const rawBarang = String(item.barang || "").trim();
        
        if (rawBarang) {
          const parts = rawBarang.split(" - ");
          kode = parts[0] ? parts[0].trim() : "";
          inputMotif = parts.length > 2 ? parts[parts.length - 1].trim() : "";
        }

        let kategori = "-";
        let supplier = "-";
        let motif = "-";
        let hargaJatuhAnak = 0;
        let insentif = 0;
        const qtyNum = parseFloat(item.qty) || 0;

        // Cari berdasarkan kesesuaian Kode DAN Motif
        let found = masterBarangData.find(row => {
          const mKode = String(row[0] || '').trim().toUpperCase();
          const mMotif = String(row[3] || '').trim().toUpperCase();
          return mKode === kode.toUpperCase() && mMotif === inputMotif.toUpperCase();
        });

        // Fallback: Jika tidak ketemu kombinasi spesifik, cari berdasarkan Kode saja
        if (!found) {
          found = masterBarangData.find(row => String(row[0]).trim().toUpperCase() === kode.toUpperCase());
        }

        if (found) {
          kode = found[0] || kode;
          namaBarang = found[1] || namaBarang; // Menggunakan Nama Barang asli dari master
          kategori = found[2] || "-";
          motif = found[3] || "-";
          supplier = found[5] || "-";
          if (colIndexHarga !== -1) {
            hargaJatuhAnak = parseFloat(found[colIndexHarga]) || 0;
          }
        }

      totalQtyGlobal += qtyNum;
      insentif = hargaJatuhAnak * qtyNum;

      newDetailRows.push([
        unit, idLaporan, data.tanggal, data.tipe, data.karu, data.shift,
        data.pekerjaan, data.operators, kode, namaBarang, kategori, supplier, motif,
        qtyNum, item.keterangan, hargaJatuhAnak, insentif, docStatus
      ]);
    });

    if (isEdit) {
      // MODE UPDATE: Hapus row rincian lama, lalu timpa
      const detailData = detailSheet.getDataRange().getValues();
      for (let i = detailData.length - 1; i >= 1; i--) {
        if (detailData[i][1] === idLaporan) {
          detailSheet.deleteRow(i + 1);
        }
      }
      
      // Masukkan row rincian yang baru di update
      if (newDetailRows.length > 0) {
        detailSheet.getRange(detailSheet.getLastRow() + 1, 1, newDetailRows.length, newDetailRows[0].length).setValues(newDetailRows);
      }
      
      // Update data di sheet Global
      const globalData = globalSheet.getDataRange().getValues();
      let globalUpdated = false;
      for (let i = 1; i < globalData.length; i++) {
        if (globalData[i][1] === idLaporan) {
          globalSheet.getRange(i + 1, 1, 1, 10).setValues([[
            unit, idLaporan, data.tanggal, data.tipe, data.karu, data.shift,
            data.pekerjaan, data.operators, totalQtyGlobal, docStatus
          ]]);
          globalUpdated = true;
          break;
        }
      }
      if (!globalUpdated) {
        globalSheet.appendRow([unit, idLaporan, data.tanggal, data.tipe, data.karu, data.shift, data.pekerjaan, data.operators, totalQtyGlobal, docStatus]);
      }
      
    } else {
      // MODE INSERT: Masukkan dokumen baru sepenuhnya
      if (newDetailRows.length > 0) {
        detailSheet.getRange(detailSheet.getLastRow() + 1, 1, newDetailRows.length, newDetailRows[0].length).setValues(newDetailRows);
      }
      globalSheet.appendRow([unit, idLaporan, data.tanggal, data.tipe, data.karu, data.shift, data.pekerjaan, data.operators, totalQtyGlobal, docStatus]);
    }

    return { success: true, idLaporan: idLaporan };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getSisaOmData() {
  try {
    const ss = SpreadsheetApp.openById(PROD_FILE_ID);
    const sheet = ss.getSheetByName("RINCIAN_SISA_OM");
    
    // Jika sheet belum ada, kembalikan array kosong agar frontend tidak error
    if (!sheet) {
      return []; 
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    // Jika sheet kosong (tidak ada baris atau kolom), kembalikan array kosong
    if (lastRow < 1 || lastCol < 1) {
      return []; 
    }
    
    // Ambil seluruh data (termasuk header baris pertama) menggunakan getDisplayValues
    // agar format angka/tanggal terbaca sebagai teks sesuai dengan tampilan di Google Sheets
    const data = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    
    return data;
  } catch (error) {
    // Kembalikan array kosong jika terjadi kegagalan akses (misal: ID file salah/dihapus)
    return [];
  }
}

function updateStatusBatal(idLaporan) {
  try {
    const ss = SpreadsheetApp.openById(PROD_FILE_ID);
    const sheetGlobal = ss.getSheetByName("DATA_PRODUKSI_GLOBAL");
    const sheetDetail = ss.getSheetByName("DATA_PRODUKSI_DETAIL");
    
    if (!sheetGlobal || !sheetDetail) throw new Error("Sheet tidak ditemukan.");

    let updatedCount = 0;

    // 1. Update Global: Status -> "Batal", Total Qty -> 0
    const dataGlobal = sheetGlobal.getDataRange().getValues();
    for (let i = 0; i < dataGlobal.length; i++) {
      if (dataGlobal[i][1] === idLaporan) {
        sheetGlobal.getRange(i + 1, 9).setValue(0); // Kolom I: Total Qty
        sheetGlobal.getRange(i + 1, 10).setValue("Batal"); // Kolom J: Status
        updatedCount++;
      }
    }

    // 2. Update Detail: Status -> "Batal", Qty -> 0, Insentif -> 0
    const dataDetail = sheetDetail.getDataRange().getValues();
    for (let i = 0; i < dataDetail.length; i++) {
      if (dataDetail[i][1] === idLaporan) {
        sheetDetail.getRange(i + 1, 14).setValue(0); // Kolom N: Qty
        sheetDetail.getRange(i + 1, 17).setValue(0); // Kolom Q: Insentif
        sheetDetail.getRange(i + 1, 18).setValue("Batal"); // Kolom R: Status
      }
    }

    if (updatedCount === 0) return { success: false, message: "ID Laporan tidak ditemukan." };
    return { success: true, message: "Dokumen dibatalkan." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateStatusApprove(idLaporan) {
  try {
    const ss = SpreadsheetApp.openById(PROD_FILE_ID);
    
    // Harus update kedua Sheet: Global & Detail
    const sheetGlobal = ss.getSheetByName("DATA_PRODUKSI_GLOBAL"); 
    const sheetDetail = ss.getSheetByName("DATA_PRODUKSI_DETAIL"); 
    
    if (!sheetGlobal || !sheetDetail) {
      throw new Error("Sheet Database Produksi tidak ditemukan di server.");
    }

    let updatedCount = 0;

    // 1. Update Sheet Global (Status berada di Kolom J / Kolom ke-10)
    const dataGlobal = sheetGlobal.getDataRange().getValues();
    for (let i = 0; i < dataGlobal.length; i++) {
      if (dataGlobal[i][1] === idLaporan) {
        sheetGlobal.getRange(i + 1, 10).setValue("Approved");
        updatedCount++;
      }
    }

    // 2. Update Sheet Detail (Status berada di Kolom Q / Kolom ke-17)
    // 2. Update Sheet Detail (Status berada di Kolom R / Kolom ke-18)
    const dataDetail = sheetDetail.getDataRange().getValues();
    for (let i = 0; i < dataDetail.length; i++) {
      if (dataDetail[i][1] === idLaporan) {
        sheetDetail.getRange(i + 1, 18).setValue("Approved");
      }
    }

    if (updatedCount === 0) {
      return { success: false, message: "ID Laporan " + idLaporan + " tidak ditemukan." };
    }

    return { success: true, message: "Dokumen berhasil disahkan!" };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getTrackingSjData() {
  try {
    const ss = SpreadsheetApp.openById(PROD_FILE_ID);
    let sheet = ss.getSheetByName("TRACKING_SJ_DO");
    
    if (!sheet) {
      sheet = ss.insertSheet("TRACKING_SJ_DO");
      sheet.appendRow(["TRACKING PENGIRIMAN SURAT JALAN"]);
      sheet.appendRow([""]);
      sheet.appendRow(["Unit", "No. SJ", "Tgl SJ", "Tgl Kirim Finance", "Tgl Diterima", "Pelanggan", "Status", "SLA (hari)"]);
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) return [];
    
    const rawData = sheet.getRange(4, 1, lastRow - 3, 8).getValues();
    
    // Format ulang agar selalu 'yyyy-MM-dd' supaya filter kalender frontend tidak error/hilang
    return rawData.map(r => {
        let tglSj = r[2];
        let tglKirim = r[3];
        let tglTerima = r[4];
        
        // Konversi jika data terbaca sebagai Objek Tanggal
        if (tglSj instanceof Date) {
            tglSj = Utilities.formatDate(tglSj, "Asia/Jakarta", "yyyy-MM-dd");
        } else if (typeof tglSj === "number") {
            // Konversi jika terbaca sebagai serial number Excel
            tglSj = Utilities.formatDate(new Date((tglSj - 25569) * 86400 * 1000), "Asia/Jakarta", "yyyy-MM-dd");
        } else if (typeof tglSj === "string" && tglSj.trim() !== "") {
            // FIX BUG: Tangkap data yang terlanjur terekam sebagai Plain-Text
            let d = new Date(tglSj);
            if (!isNaN(d.getTime())) tglSj = Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd");
        }

        if (tglKirim instanceof Date) {
            tglKirim = Utilities.formatDate(tglKirim, "Asia/Jakarta", "yyyy-MM-dd");
        }
        if (tglTerima instanceof Date) {
            tglTerima = Utilities.formatDate(tglTerima, "Asia/Jakarta", "yyyy-MM-dd");
        }
        
        // Index kembalian: 0:Unit, 1:No SJ, 2:Tgl SJ, 3:Tgl Kirim, 4:Tgl Terima, 5:Pelanggan, 6:Status, 7:SLA
        return [r[0], r[1], tglSj, tglKirim, tglTerima, r[5], r[6], r[7]];
    });
  } catch(e) {
    return [];
  }
}

function updateStatusSj(ids, newStatus) {
  try {
    const ss = SpreadsheetApp.openById(PROD_FILE_ID);
    const sheet = ss.getSheetByName("TRACKING_SJ_DO");
    if(!sheet) return {success:false, message:"Sheet TRACKING_SJ_DO tidak ditemukan"};
    
    const data = sheet.getDataRange().getValues();
    let updated = 0;
    const today = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");

    for(let i=3; i<data.length; i++) { // Mulai dari baris ke-4 (indeks 3)
      if(ids.includes(data[i][1])) { 
        sheet.getRange(i+1, 7).setValue(newStatus); // Update Kolom G (7) - Status
        
        if(newStatus === "PROSES KIRIM FINANCE") {
          sheet.getRange(i+1, 4).setValue(today); // Tgl Kirim Finance (Kolom D/4)
        } 
        else if(newStatus === "DITERIMA FINANCE") {
          sheet.getRange(i+1, 5).setValue(today); // Tgl Diterima Finance (Kolom E/5)
        }
        else if(newStatus === "DIKEMBALIKAN KE UNIT" || newStatus === "DIKIRIM KE SUPPLIER") {
          sheet.getRange(i+1, 5).setValue(""); // Hapus tanggal terima jika dikembalikan / direset
        }
        
        // Kalkulasi SLA (Hari) Berjalan
        let tglSj = new Date(data[i][2]); // Kolom C (Tgl SJ)
        if (!isNaN(tglSj.getTime())) {
          let diffTime = Math.abs(new Date() - tglSj);
          let sla = Math.floor(diffTime / (1000 * 60 * 60 * 24));
          sheet.getRange(i+1, 8).setValue(sla); // Update Kolom H (8) - SLA
        }
        
        updated++;
      }
    }
    return {success:true, message: updated + " data diperbarui"};
  } catch(e) {
    return {success:false, message: e.message};
  }
}

function saveImportSJ(rows) {
  try {
    const ss = SpreadsheetApp.openById(PROD_FILE_ID);
    let sheet = ss.getSheetByName("TRACKING_SJ_DO");
    
    // 1. Setup Sheet jika belum ada
    if (!sheet) {
      sheet = ss.insertSheet("TRACKING_SJ_DO");
      sheet.appendRow(["TRACKING PENGIRIMAN SURAT JALAN"]);
      sheet.appendRow([""]);
      sheet.appendRow(["Unit", "No. SJ", "Tgl SJ", "Tgl Kirim Finance", "Tgl Diterima", "Pelanggan", "Status", "SLA (hari)"]);
    }
    
    // 2. Ambil data yang sudah ada di Sheet (Baris 4 ke bawah)
    const lastRow = sheet.getLastRow();
    let existingData = [];
    if (lastRow >= 4) {
      existingData = sheet.getRange(4, 1, lastRow - 3, 8).getValues();
    }
    
    // 3. Buat "Kamus/Map" berdasarkan No. SJ untuk pencarian cepat
    const existingMap = new Map();
    existingData.forEach((row, index) => {
      const noSj = String(row[1]).trim(); // Index 1 adalah No SJ
      existingMap.set(noSj, { data: row, index: index });
    });
    
    let newDataToAppend = [];
    const today = new Date();
    
    // 4. Proses data dari file Excel (Import)
    rows.forEach(r => {
       // FIX BUG: Pastikan format tanggal masuk ke server dengan format standar (yyyy-MM-dd)
       let rawDate = r.tglSj;
       if (typeof rawDate === "number") {
          rawDate = Utilities.formatDate(new Date((rawDate - 25569) * 86400 * 1000), "Asia/Jakarta", "yyyy-MM-dd");
       } else {
          let d = new Date(rawDate);
          if (!isNaN(d.getTime())) {
             rawDate = Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd");
          }
       }
       
       const noSjKey = String(r.noSj).trim();
       
       if (existingMap.has(noSjKey)) {
         // --- JIKA DATA SUDAH ADA (UPDATE INFO DASAR SAJA) ---
         let existingRow = existingMap.get(noSjKey).data;
         
         existingRow[0] = r.unit;        // Update Unit
         existingRow[2] = rawDate;       // Update Tgl SJ
         existingRow[5] = r.pelanggan;   // Update Pelanggan
         
         // Hitung ulang SLA hanya jika status BUKAN "DITERIMA FINANCE"
         let tglSjObj = new Date(rawDate);
         if (existingRow[6] !== "DITERIMA FINANCE" && !isNaN(tglSjObj.getTime())) {
            existingRow[7] = Math.floor(Math.abs(today - tglSjObj) / (1000 * 60 * 60 * 24));
         }
         
         // CATATAN PENTING: 
         // existingRow[3] (Tgl Kirim), existingRow[4] (Tgl Terima), dan existingRow[6] (Status) 
         // TIDAK KITA SENTUH agar historinya tetap aman.
         
       } else {
         // --- JIKA DATA BELUM ADA (TAMBAH SEBAGAI DATA BARU) ---
         let sla = 0;
         let tglSjObj = new Date(rawDate);
         if (!isNaN(tglSjObj.getTime())) {
            sla = Math.floor(Math.abs(today - tglSjObj) / (1000 * 60 * 60 * 24));
         }
         
         // Urutan 8 Kolom: Unit, No SJ, Tgl SJ, Tgl Kirim, Tgl Terima, Pelanggan, Status, SLA
         newDataToAppend.push([r.unit, r.noSj, rawDate, "", "", r.pelanggan, r.status, sla]);
       }
    });
    
    // 5. Simpan kembali data yang sudah di-update (menimpa data lama dengan data lama yang diperbarui)
    if (existingData.length > 0) {
       sheet.getRange(4, 1, existingData.length, 8).setValues(existingData);
    }
    
    // 6. Tambahkan data baru di baris paling bawah
    if (newDataToAppend.length > 0) {
       sheet.getRange(sheet.getLastRow() + 1, 1, newDataToAppend.length, 8).setValues(newDataToAppend);
    }
    
    return {success: true};
  } catch(e) {
    return {success: false, message: e.message};
  }
}

function saveBbkToSheet(data) {
  try {
    const ss = SpreadsheetApp.openById(PROD_FILE_ID);
    
    // 1. Setup Sheet Detail
    let sheet = ss.getSheetByName("DATA_PENGECEKAN_BBK_DETAIL");
    if (!sheet) {
      sheet = ss.insertSheet("DATA_PENGECEKAN_BBK_DETAIL");
      sheet.appendRow(["DATA HASIL PENGECEKAN BAHAN BAKU DETAIL"]);
      sheet.appendRow(["Unit", "Tanggal", "Pekerjaan", "Nomor Dokumen", "Nama Operator", "Shift", "Supplier", "Refferensi Dokumen SJ", "Refferensi Nomor OM", "Kode", "Nama barang", "Kategori", "Motif", "Qty", "Keterangan", "Status"]);
    }
    
    // 2. Setup Sheet Global
    let globalSheet = ss.getSheetByName("DATA_PENGECEKAN_BBK_GLOBAL");
    if (!globalSheet) {
      globalSheet = ss.insertSheet("DATA_PENGECEKAN_BBK_GLOBAL");
      globalSheet.appendRow(["DATA HASIL PENGECEKAN BAHAN BAKU GLOBAL"]);
      globalSheet.appendRow(["Unit", "Tanggal", "Proses Item", "Nomor Dokumen", "Nama Operator", "Shift", "Supplier", "Refferensi Dokumen SJ", "Refferensi Nomor OM", "Keterangan", "Status"]);
    }

    // 3. Generate Nomor Dokumen BBK (Contoh: LPS-BB-2608-000001)
    const dateObj = new Date(data.tanggal);
    const yy = String(dateObj.getFullYear()).slice(-2);
    const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
    const prefix = `LPS-BB-${yy}${mm}-`;
    
    const globalLastRow = globalSheet.getLastRow();
    let runningNum = 1;
    if (globalLastRow >= 3) {
      const ids = globalSheet.getRange(3, 4, globalLastRow - 2, 1).getValues(); // Kolom D (index 4) = Nomor Dokumen
      let maxNum = 0;
      ids.forEach(row => {
        if (row[0] && row[0].toString().startsWith(prefix)) {
          const num = parseInt(row[0].replace(prefix, '')) || 0;
          if (num > maxNum) maxNum = num;
        }
      });
      runningNum = maxNum + 1;
    }
    const noDokumenBBK = prefix + String(runningNum).padStart(6, '0');

    let masterBarangData = [];
    try {
      const masterSs = SpreadsheetApp.openById(MASTER_FILE_ID);
      const masterSheet = masterSs.getSheetByName("MASTER_BARANG");
      if (masterSheet && masterSheet.getLastRow() >= 4) {
        masterBarangData = masterSheet.getRange(4, 1, masterSheet.getLastRow() - 3, 18).getValues();
      }
    } catch(e) {}

    let rowsToAppend = [];
    let docStatus = "Sesuai"; // Menggantikan waktu server/kondisi

    data.items.forEach(item => {
      let rawBarang = String(item.barang || "").trim();
      let parts = rawBarang.split(" - ");
      let kode = parts[0] ? parts[0].trim() : "";
      let inputMotif = parts.length > 2 ? parts[parts.length - 1].trim() : "";
      let namaBarang = "";
      let kategori = "-";
      let motif = "-";

      let found = masterBarangData.find(row => {
        const mKode = String(row[0] || '').trim().toUpperCase();
        const mMotif = String(row[3] || '').trim().toUpperCase();
        return mKode === kode.toUpperCase() && mMotif === inputMotif.toUpperCase();
      });

      if (!found) {
        found = masterBarangData.find(row => String(row[0]).trim().toUpperCase() === kode.toUpperCase());
      }

      if (found) {
        kode = found[0];
        namaBarang = found[1];
        kategori = found[2] || "-";
        motif = found[3] || "-";
      }

      const qtyNum = parseFloat(item.qty) || 0;

      rowsToAppend.push([
        data.unit || "-",
        data.tanggal || "-",
        data.pekerjaan || "-",
        noDokumenBBK,
        data.operator || "-",
        data.shift || "-",
        data.supplier || "-",
        data.refDokumen || "-",
        data.refOm || "-",
        kode,
        namaBarang,
        kategori,
        motif,
        qtyNum,
        item.keterangan || "-",
        docStatus
      ]);
    });

    if (rowsToAppend.length > 0) {
      // Simpan di sheet detail
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
      
      // Hitung rangkuman untuk keperluan lain (opsional)
      let totalQty = 0;
      data.items.forEach(it => totalQty += (parseFloat(it.qty) || 0));
      
// Simpan di sheet global
      globalSheet.appendRow([
        data.unit || "-",
        data.tanggal || "-",
        data.pekerjaan || "-",
        noDokumenBBK,
        data.operator || "-",
        data.shift || "-",
        data.supplier || "-",
        data.refDokumen || "-",
        data.refOm || "-",
        data.ketHeader || "-",
        docStatus
      ]);
    }

    // Kembalikan ID Dokumen ke frontend
    return { success: true, idLaporan: noDokumenBBK };
  } catch(error) {
    return { success: false, message: error.toString() };
  }
}

// Tambahkan ini di file Kode.gs Anda
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Data kiriman kosong (e.postData is empty).");
    }

    let request = JSON.parse(e.postData.contents);
    let action = request.action;   
    let payload = request.payload || {}; 
    
    let result;

    if (action === 'verifikasiLogin') {
      result = verifikasiLogin(payload.username, payload.password);
    } 
    else if (action === 'getMasterBarangData') {
      result = getMasterBarangData();
    }
    else if (action === 'getMasterOperatorData') {
      result = getMasterOperatorData();
    }
    else if (action === 'getSaldoMutasiData') {
      result = getSaldoMutasiData();
    }
    else if (action === 'getDashboardStats') {
      result = getDashboardStats();
    }
    else if (action === 'getDataLaporanLatest') {
      result = getDataLaporanLatest();
    }
    else if (action === 'getDataLaporan') {
      result = getDataLaporan();
    }
    else if (action === 'getDataProduksi') {
      result = getDataProduksi();
    }
    else if (action === 'getBbkGlobalData') {
      result = getBbkGlobalData();
    }
    else if (action === 'getBbkDetailData') {
      result = getBbkDetailData();
    }
    else if (action === 'saveProduksiToSheet') {
      result = saveProduksiToSheet(payload); 
    }
    else if (action === 'getSisaOmData') {
      result = getSisaOmData();
    }
    else if (action === 'updateStatusBatal') {
      result = updateStatusBatal(payload.idLaporan);
    }
    else if (action === 'updateStatusApprove') {
      result = updateStatusApprove(payload.idLaporan);
    }
    else if (action === 'getTrackingSjData') {
      result = getTrackingSjData();
    }
    else if (action === 'updateStatusSj') {
      result = updateStatusSj(payload.ids, payload.newStatus);
    }
    else if (action === 'saveImportSJ') {
      result = saveImportSJ(payload.rows);
    }
    else if (action === 'saveBbkToSheet') {
      result = saveBbkToSheet(payload);
    }
    else {
      throw new Error("Action '" + action + "' tidak terdaftar di router.");
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}