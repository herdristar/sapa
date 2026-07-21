/**
 * ZettBOT 7.4 - code.gs
 * RESTORED: Backend Logic, Routing, Aggregation, Gemini AI & CRUD User Admin
 */

// =========================================================================
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
// =========================================================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('SAPA - Survei Adaptif ASN Demak')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function generateId(sheet) {
  const timeZone = 'Asia/Jakarta';
  const today = Utilities.formatDate(new Date(), timeZone, 'yyyyMMdd');
  const prefix = 'SRV-' + today + '-';
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return prefix + '0001';

  const lastId = sheet.getRange(lastRow, 1).getValue().toString();
  if (lastId.startsWith(prefix)) {
    const sequence = parseInt(lastId.split('-')[2], 10);
    let strSeq = (sequence + 1).toString();
    while (strSeq.length < 4) strSeq = '0' + strSeq;
    return prefix + strSeq;
  }
  return prefix + '0001';
}

function kalkulasiHasil(totalSkor) {
  const indeks = (totalSkor / 85) * 100;
  let kategori = ""; let rekomendasi = "";
  if (indeks >= 84.01) { kategori = "Sangat Baik"; rekomendasi = "ASN telah sepenuhnya menginternalisasi budaya kerja lokal dan memiliki mindset pembelajar. Siap diikutsertakan dalam Talent Pool."; }
  else if (indeks >= 68.01) { kategori = "Baik"; rekomendasi = "ASN mampu beradaptasi dengan baik. Cukup pertahankan melalui skema Knowledge Management atau Coaching."; }
  else if (indeks >= 52.01) { kategori = "Cukup"; rekomendasi = "ASN masih dalam fase transisi dan terkadang mengalami hambatan psikologis. Membutuhkan mentoring spesifik."; }
  else if (indeks >= 36.01) { kategori = "Kurang"; rekomendasi = "ASN kesulitan menyesuaikan diri dengan sistem kerja birokrasi, kaku, atau resisten terhadap perubahan."; }
  else { kategori = "Sangat Kurang"; rekomendasi = "Terdapat red flag ketidakcocokan nilai (culture mismatch). Evaluasi komprehensif diperlukan."; }
  return { indeks: indeks.toFixed(2), kategori, rekomendasi };
}

function simpanRespon(rawData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respon_Survei");
    const data = typeof rawData === 'string' ? JSON.parse(rawData) : rawData;

    const timeZone = 'Asia/Jakarta';
    const timestamp = Utilities.formatDate(new Date(), timeZone, 'dd/MM/yyyy HH:mm:ss');
    const newId = generateId(sheet);

    let totalSkor = 0; const qValues = [];
    for (let i = 1; i <= 17; i++) {
      let val = parseInt(data['Q' + i], 10) || 0;
      totalSkor += val; qValues.push(val);
    }

    const hasil = kalkulasiHasil(totalSkor);

    const newRow = [
      newId, timestamp, data.nip, data.nama, data.opd, data.jabatan,
      ...qValues, totalSkor, parseFloat(hasil.indeks), hasil.kategori, hasil.rekomendasi,
      data.Q1_FU || "-", data.Q4_FU || "-", data.Q8_FU || "-",
      data.Q11_FU || "-", data.Q13_FU || "-", data.Q17_FU || "-",
      "Belum", "-"
    ];

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();

    return { status: 'success', id: newId, nip: data.nip, nama: data.nama, opd: data.opd, skor: totalSkor, indeks: hasil.indeks, kategori: hasil.kategori, rekomendasi: hasil.rekomendasi };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}

// ================= RESTORED: CRUD USER ADMIN =================
function listUserAdmin() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("User_Admin");
    const data = sheet.getDataRange().getDisplayValues();
    const users = [];
    for (let i = 1; i < data.length; i++) {
      users.push({ index: i + 1, id: data[i][0], role: data[i][1], opd: data[i][2], pin: data[i][3] });
    }
    return { status: 'success', data: users };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}

function simpanUserAdmin(opd, pin) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("User_Admin");
    const lastRow = sheet.getLastRow();
    const nextId = "ADM-" + (lastRow).toString().padStart(3, '0');
    sheet.appendRow([nextId, "Pimpinan OPD", opd, pin.toString()]);
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}

function hapusUserAdmin(rowIndex) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("User_Admin");
    sheet.deleteRow(parseInt(rowIndex));
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}

function ubahPinSelf(oldPin, newPin) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("User_Admin");
    const data = sheet.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === oldPin.toString()) {
        sheet.getRange(i + 1, 4).setValue(newPin.toString());
        SpreadsheetApp.flush();
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'PIN lama tidak sesuai.' };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}

// ================= AI VALIDATION LOGIC =================
function callGeminiAPI(promptText) {
  // Hanya cek apakah variabel kosong, tidak perlu cek isinya lagi
  if (!GEMINI_API_KEY || GEMINI_API_KEY === "") {
    throw new Error("API Key Gemini tidak ditemukan atau kosong.");
  }

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key=' + GEMINI_API_KEY;
  const payload = { "contents": [{ "parts": [{ "text": promptText }] }] };
  const options = { method: "post", contentType: "application/json", muteHttpExceptions: true, payload: JSON.stringify(payload) };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (json.error) throw new Error(json.error.message);
  if (!json.candidates || json.candidates.length === 0) throw new Error("Tidak ada respon dari AI Gemini.");

  return json.candidates[0].content.parts[0].text;
}

function prosesKoreksiAI(rowId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respon_Survei");
    const data = sheet.getDataRange().getDisplayValues();
    let targetRowIdx = -1; let rowData = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === rowId) { targetRowIdx = i + 1; rowData = [...data[i]]; break; }
    }

    if (targetRowIdx === -1) return { status: 'error', message: "Data tidak ditemukan." };

    const targets = [
      { qIdx: 6, fuIdx: 27, text: "Ceritakan pengalaman saat diberi tugas mendadak/aplikasi baru." },
      { qIdx: 9, fuIdx: 28, text: "Bagaimana cara Anda menyelesaikan tugas di luar jadwal/tim baru?" },
      { qIdx: 13, fuIdx: 29, text: "Sebutkan ide inovasi atau kesalahan teknis dan perbaikannya." },
      { qIdx: 16, fuIdx: 30, text: "Evaluasi inovasi dan perbaikan." },
      { qIdx: 18, fuIdx: 31, text: "Contoh nyata inisiatif kerja di luar instruksi." },
      { qIdx: 22, fuIdx: 32, text: "Usulan penyederhanaan SOP." }
    ];

    let totalKoreksiLog = "";
    let isModified = false;

    for (let t of targets) {
      let skorAwal = parseInt(rowData[t.qIdx], 10);
      let jawabanFU = rowData[t.fuIdx];

      if ((skorAwal === 4 || skorAwal === 5) && jawabanFU !== "-") {
        let promptStr = "Anda adalah Asesor SDM Senior Pemkab Demak.\n";
        promptStr += "Konteks Pertanyaan: " + t.text + "\n";
        promptStr += "Skor Awal Responden: " + skorAwal + "\n";
        promptStr += "Jawaban Responden: " + jawabanFU + "\n\n";
        promptStr += "Evaluasi dengan metode STAR (Situation, Task, Action, Result).\n";
        promptStr += "Level 3: Valid (spesifik, nyata, aksi mandiri).\nLevel 2: Parsial (general, teoritis).\nLevel 1: Invalid (kosong, normatif).\n";
        promptStr += "Output JSON MURNI: {\"Kategori\": \"[Level 1/2/3]\", \"Analisis\": \"[Alasan]\", \"Skor_Akhir\": [Jika L3=tetap, L2=kurangi 1, L1=jadikan 2]}";

        try {
          let aiRawResult = callGeminiAPI(promptStr);
          let cleanJsonStr = aiRawResult.replace(/```json/g, '').replace(/```/g, '').trim();
          let aiResult = JSON.parse(cleanJsonStr);

          if (aiResult.Skor_Akhir !== skorAwal) {
            rowData[t.qIdx] = aiResult.Skor_Akhir;
            totalKoreksiLog += "[Q" + (t.qIdx - 5) + "] Terkoreksi: " + aiResult.Analisis + "\n";
            isModified = true;
          } else {
            totalKoreksiLog += "[Q" + (t.qIdx - 5) + "] Valid.\n";
          }
        } catch (e) { totalKoreksiLog += "[Q" + (t.qIdx - 5) + "] Gagal AI: " + e.message + "\n"; }
      }
    }

    if (isModified) {
      let newTotalSkor = 0;
      for (let i = 6; i <= 22; i++) newTotalSkor += parseInt(rowData[i], 10) || 0;
      let newHasil = kalkulasiHasil(newTotalSkor);
      rowData[23] = newTotalSkor; rowData[24] = parseFloat(newHasil.indeks);
      rowData[25] = newHasil.kategori; rowData[26] = newHasil.rekomendasi;
    }

    rowData[33] = isModified ? "Terkoreksi" : "Tervalidasi (Aman)";
    rowData[34] = totalKoreksiLog || "Tidak ada jawaban Level 4/5";

    sheet.getRange(targetRowIdx, 1, 1, rowData.length).setValues([rowData]);
    SpreadsheetApp.flush();
    return { status: 'success', message: "Koreksi AI Selesai." };

  } catch (error) { return { status: 'error', message: error.toString() }; }
}

function generateExecutiveSummaryAI() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respon_Survei");
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return { status: 'error', message: "Data kosong." };

    let kumpulanHambatan = [];
    for (let i = 1; i < data.length; i++) {
      const fuCols = [27, 28, 29, 30, 31, 32];
      const opd = data[i][4];
      fuCols.forEach(colIdx => {
        let val = data[i][colIdx];
        if (val !== "-" && val.length < 100) kumpulanHambatan.push("[" + opd + "] " + val);
      });
    }

    if (kumpulanHambatan.length === 0) return { status: 'error', message: "Belum ada indikasi hambatan adaptasi." };

    let promptStr = "Data hambatan adaptasi ASN:\n" + kumpulanHambatan.join("\n") + "\n\n";
    promptStr += "Buat Executive Summary 3 paragraf untuk Kepala BKPSDM (Format HTML MURNI: <b>, <p>, <ul>).\n";
    promptStr += "1. Identifikasi 2 hambatan terbesar (Pareto).\n2. Pola OPD.\n3. Dua rekomendasi intervensi HR.";

    let summaryHTML = callGeminiAPI(promptStr);
    summaryHTML = summaryHTML.replace(/```html/g, '').replace(/```/g, '').trim();
    return { status: 'success', report: summaryHTML };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}

// ================= ANALYTICS & DASHBOARD =================
function verifikasiPin(pin) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("User_Admin");
    const data = sheet.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === pin.toString()) return { status: 'success', role: data[i][1], opd: data[i][2] };
    }
    return { status: 'error', message: 'PIN salah atau tidak aktif!' };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}

function ambilDataAdmin(page = 1, limit = 10, search = "", filterOpd = "", currentAdminRole = "Admin Pemda", currentAdminOpd = "Semua") {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respon_Survei");
    const rawData = sheet.getDataRange().getDisplayValues();

    let targetFilterOpd = (currentAdminRole === "Pimpinan OPD") ? currentAdminOpd : filterOpd;

    if (rawData.length <= 1) {
      return {
        status: 'success', totalData: 0,
        globalStats: { count: 0, I_org: 0, D1: 0, D2: 0, D3: 0 },
        opdStats: { count: 0, I_org: 0, D1: 0, D2: 0, D3: 0 },
        data: [], totalPages: 0, top10Global: [], top5Opd: []
      };
    }

    const body = rawData.slice(1);
    let gCount = 0, gIndeks = 0, gD1 = 0, gD2 = 0, gD3 = 0;
    let oCount = 0, oIndeks = 0, oD1 = 0, oD2 = 0, oD3 = 0;
    let validRows = [];

    body.forEach(row => {
      const opdRow = row[4];
      const rawIndeks = row[24] ? row[24].toString().replace(/,/g, '.') : "0";
      const idxStr = parseFloat(rawIndeks);

      if (!isNaN(idxStr) && idxStr > 0) {
        let d1Sum = 0, d2Sum = 0, d3Sum = 0;
        for (let i = 6; i <= 12; i++) d1Sum += parseFloat(row[i] ? row[i].toString().replace(/,/g, '.') : "0") || 0;
        for (let i = 13; i <= 17; i++) d2Sum += parseFloat(row[i] ? row[i].toString().replace(/,/g, '.') : "0") || 0;
        for (let i = 18; i <= 22; i++) d3Sum += parseFloat(row[i] ? row[i].toString().replace(/,/g, '.') : "0") || 0;

        let d1Idx = (d1Sum / 35) * 100; let d2Idx = (d2Sum / 25) * 100; let d3Idx = (d3Sum / 25) * 100;

        gCount++; gIndeks += idxStr; gD1 += d1Idx; gD2 += d2Idx; gD3 += d3Idx;

        if (targetFilterOpd !== "" && opdRow === targetFilterOpd) {
          oCount++; oIndeks += idxStr; oD1 += d1Idx; oD2 += d2Idx; oD3 += d3Idx;
        }

        validRows.push({ nama: row[3], nip: row[2], opd: row[4], skor: row[23], indeks: row[24], parsedIndeks: idxStr, kategori: row[25], timestamp: row[1], id: row[0], status_ai: row[33] });
      }
    });

    validRows.sort((a, b) => b.parsedIndeks - a.parsedIndeks);

    // RESTORED: Top 10 & Top 5 Generation
    const top10Global = validRows.slice(0, 10).map(r => ({ nama: r.nama, nip: r.nip, opd: r.opd, skor: r.skor, indeks: r.indeks, kategori: r.kategori }));
    let top5Opd = [];
    if (targetFilterOpd !== "") top5Opd = validRows.filter(r => r.opd === targetFilterOpd).slice(0, 5).map(r => ({ nama: r.nama, nip: r.nip, skor: r.skor, indeks: r.indeks, kategori: r.kategori }));

    const globalStats = { count: gCount, I_org: gCount > 0 ? (gIndeks / gCount).toFixed(2) : 0, D1: gCount > 0 ? (gD1 / gCount).toFixed(2) : 0, D2: gCount > 0 ? (gD2 / gCount).toFixed(2) : 0, D3: gCount > 0 ? (gD3 / gCount).toFixed(2) : 0 };
    const opdStats = { count: oCount, I_org: oCount > 0 ? (oIndeks / oCount).toFixed(2) : 0, D1: oCount > 0 ? (oD1 / oCount).toFixed(2) : 0, D2: oCount > 0 ? (oD2 / oCount).toFixed(2) : 0, D3: oCount > 0 ? (oD3 / oCount).toFixed(2) : 0 };

    let tableData = [...validRows];
    if (currentAdminRole === "Pimpinan OPD") tableData = tableData.filter(row => row.opd === currentAdminOpd);
    if (filterOpd !== "" && currentAdminRole === "Admin Pemda") tableData = tableData.filter(row => row.opd === filterOpd);

    if (search.trim() !== "") {
      const q = search.toLowerCase();
      tableData = tableData.filter(row => row.nip.toLowerCase().includes(q) || row.nama.toLowerCase().includes(q) || row.opd.toLowerCase().includes(q));
    }

    const totalData = tableData.length;
    const totalPages = Math.ceil(totalData / limit);
    const currentPage = page > totalPages && totalPages > 0 ? totalPages : page;
    const startIndex = (currentPage - 1) * limit;

    const paginatedData = tableData.slice(startIndex, startIndex + limit).map(row => ({
      id: row.id, timestamp: row.timestamp, nip: row.nip, nama: row.nama, opd: row.opd, total_skor: row.skor, indeks: row.indeks, kategori: row.kategori, status_ai: row.status_ai || "Belum"
    }));

    return { status: 'success', totalData: totalData, globalStats: globalStats, opdStats: opdStats, totalPages: totalPages, currentPage: currentPage, data: paginatedData, top10Global: top10Global, top5Opd: top5Opd };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}

function downloadDataCSV() {
  try {
    const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respon_Survei").getDataRange().getDisplayValues();
    let csvString = "";
    data.forEach(row => {
      let cleanRow = row.map(cell => '"' + cell.toString().replace(/"/g, '""') + '"');
      csvString += cleanRow.join(",") + "\n";
    });
    return { status: 'success', csv: csvString };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}
