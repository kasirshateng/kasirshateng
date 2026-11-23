const SHEET_MENU = "Menu";
const SHEET_TRANSAKSI = "Transaksi";
const SHEET_USERS = "Users"; // <-- Tambahan tab Users untuk login

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const action = e.parameter.action || "all";

  if (action === "menu") {
    return jsonResponse(getSheetData(ss.getSheetByName(SHEET_MENU)));
  } else if (action === "transaksi") {
    return jsonResponse(getSheetData(ss.getSheetByName(SHEET_TRANSAKSI)));
  } else if (action === "users") {
    return jsonResponse(getSheetData(ss.getSheetByName(SHEET_USERS)));
  } else {
    return jsonResponse({
      menu: getSheetData(ss.getSheetByName(SHEET_MENU)),
      transaksi: getSheetData(ss.getSheetByName(SHEET_TRANSAKSI)),
      users: getSheetData(ss.getSheetByName(SHEET_USERS))
    });
  }
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = JSON.parse(e.postData.contents);
  const action = data.action;

  // 1. Fitur Cek Login Berdasarkan PIN di Tab Users
  if (action === "login") {
    const sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet) {
      return jsonResponse({ status: "error", message: "Tab 'Users' tidak ditemukan di Google Sheets!" });
    }
    const rows = sheet.getDataRange().getValues();
    
    // Lewati baris header (index 0)
    for (let i = 1; i < rows.length; i++) {
      // Kolom C adalah PIN (index 2), Kolom D adalah Role (index 3), Kolom B adalah Nama (index 1)
      if (String(rows[i][2]) === String(data.pin)) {
        return jsonResponse({
          status: "success",
          role: rows[i][3], // 'kasir' atau 'owner'
          nama: rows[i][1]  // Nama pengguna
        });
      }
    }
    return jsonResponse({ status: "error", message: "PIN salah atau tidak ditemukan!" });
  }

  // 2. Tambah Transaksi
  if (action === "add_transaksi") {
    const sheet = ss.getSheetByName(SHEET_TRANSAKSI);
    const now = new Date();
    const waktu = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
    const bulanTahun = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM");
    const idTrx = "TRX-" + now.getTime();

    const totalHarga = Number(data.total_harga) || 0;
    const totalModal = Number(data.total_modal) || 0;

    // Kolom: A ID | B Waktu | C Bulan_Tahun | D Detail_Item | E Total_Harga
    //        F Metode_Bayar | G Catatan | H Total_Modal | I Keuntungan
    sheet.appendRow([
      idTrx,
      waktu,
      bulanTahun,
      data.detail_item,
      totalHarga,
      data.metode_bayar,
      data.catatan || "-",
      totalModal,
      totalHarga - totalModal
    ]);

    return jsonResponse({ status: "success", id_transaksi: idTrx });
  }

  // 3. Tambah Menu
  if (action === "add_menu") {
    const sheet = ss.getSheetByName(SHEET_MENU);
    const idMenu = "M" + Math.floor(100 + Math.random() * 900);
    sheet.appendRow([
      idMenu,
      data.nama_menu,
      data.kategori,
      Number(data.harga),
      data.status || "Tersedia",
      Number(data.harga_modal) || 0,  // Kolom F: Harga_Modal
      data.jenis || "Kafe"            // Kolom G: Jenis (Kafe / Tobacco)
    ]);
    return jsonResponse({ status: "success", id_menu: idMenu });
  }

  // 4. Update Menu
  if (action === "update_menu") {
    const sheet = ss.getSheetByName(SHEET_MENU);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id_menu) {
        sheet.getRange(i + 1, 2).setValue(data.nama_menu);
        sheet.getRange(i + 1, 3).setValue(data.kategori);
        sheet.getRange(i + 1, 4).setValue(Number(data.harga));
        sheet.getRange(i + 1, 5).setValue(data.status);
        sheet.getRange(i + 1, 6).setValue(Number(data.harga_modal) || 0);
        sheet.getRange(i + 1, 7).setValue(data.jenis || "Kafe");
        return jsonResponse({ status: "success" });
      }
    }
    return jsonResponse({ status: "error", message: "Menu tidak ditemukan" });
  }

  // 5. Hapus Menu
  if (action === "delete_menu") {
    const sheet = ss.getSheetByName(SHEET_MENU);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id_menu) {
        sheet.deleteRow(i + 1);
        return jsonResponse({ status: "success" });
      }
    }
    return jsonResponse({ status: "error", message: "Menu tidak ditemukan" });
  }

  return jsonResponse({ status: "error", message: "Action tidak dikenali" });
}

function getSheetData(sheet) {
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  return rows.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, idx) => obj[h] = row[idx]);
    return obj;
  });
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}