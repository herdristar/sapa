/**
 * ZettBOT 7.4 - setup.gs
 * Fungsi Inisialisasi Database Lengkap (Dengan AI Log & Follow-Up)
 */

function inisialisasiDatabase() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Setup Sheet Respon Survei
    let sheetRespon = ss.getSheetByName("Respon_Survei");
    if (!sheetRespon) {
      sheetRespon = ss.insertSheet("Respon_Survei");
    }
    
    // Header 35 Kolom
    const headersRespon = [
      "ID_Survei", "Timestamp", "NIP", "Nama", "OPD", "Jabatan",
      "Q1", "Q2", "Q3", "Q4", "Q5", "Q6", "Q7",
      "Q8", "Q9", "Q10", "Q11", "Q12", 
      "Q13", "Q14", "Q15", "Q16", "Q17",
      "Total_Skor", "Indeks_Individu", "Kategori", "Rekomendasi",
      "Q1_Lanjutan", "Q4_Lanjutan", "Q8_Lanjutan", "Q11_Lanjutan", "Q13_Lanjutan", "Q17_Lanjutan",
      "Status_AI", "Analisis_AI_Log" 
    ];
    
    sheetRespon.getRange(1, 1, 1, headersRespon.length).setValues([headersRespon]);
    const headerRange1 = sheetRespon.getRange(1, 1, 1, headersRespon.length);
    headerRange1.setFontWeight("bold").setBackground("#10b981").setFontColor("white");
    sheetRespon.setFrozenRows(1);
    sheetRespon.autoResizeColumns(1, headersRespon.length);

    // 2. Setup Sheet User Admin
    let sheetAdmin = ss.getSheetByName("User_Admin");
    if (!sheetAdmin) {
      sheetAdmin = ss.insertSheet("User_Admin");
      const headersAdmin = ["ID_User", "Role", "OPD", "PIN"];
      sheetAdmin.getRange(1, 1, 1, headersAdmin.length).setValues([headersAdmin]);
      sheetAdmin.getRange(1, 1, 1, headersAdmin.length).setFontWeight("bold").setBackground("#3b82f6").setFontColor("white");
      sheetAdmin.appendRow(["ADM-001", "Admin Pemda", "Semua", "123456"]);
      sheetAdmin.setFrozenRows(1);
      sheetAdmin.autoResizeColumns(1, headersAdmin.length);
    }
    
    SpreadsheetApp.getUi().alert("✅ Sukses: Inisialisasi Database SAPA selesai!");
  } catch (error) {
    SpreadsheetApp.getUi().alert("❌ Gagal Setup: " + error.toString());
  }
}