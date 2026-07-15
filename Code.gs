/* ============================================================
   GANTI DENGAN ID GOOGLE SHEET ANDA
   ============================================================ */
const SHEET_ID = "1PxLr6PTtnkx4t3IjgAIwOEY38QvcnExrI5618XWH4Xg";
const SHEET_DATA = "DataNOP";
const SHEET_ADMIN = "Admin";

function doGet(e) {
  const action = e.parameter.action || "read";
  const token = e.parameter.token || "";
  if (action === "read") {
    if (!validateToken(token)) return jsonResponse({ success: false, message: "Token invalid" });
    return jsonResponse({ success: true, data: getAllData() });
  }
  return jsonResponse({ success: false, message: "Action tidak dikenal" });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { action, data, token } = body;
    let result = {};

    switch(action) {
      case "create":
        if (!validateToken(token)) throw new Error("Sesi berakhir");
        const user = getUserFromToken(token);
        if (user.role !== 'admin') throw new Error("Hanya admin yang bisa input data");
        result = { success: true, message: "✅ Data berhasil disimpan", id: createData(data) };
        break;
      case "update":
        if (!validateToken(token)) throw new Error("Sesi berakhir");
        const user2 = getUserFromToken(token);
        if (user2.role !== 'admin') throw new Error("Hanya admin yang bisa update");
        result = { success: true, message: updateData(data) };
        break;
      case "login":
        result = doLogin(data);
        break;
      default:
        throw new Error("Action tidak valid");
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

/* ====== AUTH ====== */
function doLogin(data) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_ADMIN);
  if (!sh) throw new Error("Sheet Admin belum dibuat");
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.username && rows[i][1] === data.password) {
      const payload = { username: data.username, role: rows[i][2] || 'operator', time: Date.now() };
      const token = Utilities.base64Encode(JSON.stringify(payload));
      return { success: true, token, username: data.username, role: rows[i][2] || 'operator' };
    }
  }
  throw new Error("Username atau password salah");
}

function getUserFromToken(token) {
  try {
    const decoded = JSON.parse(Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString());
    return decoded;
  } catch(e) {
    throw new Error("Token invalid");
  }
}

function validateToken(token) {
  if (!token || token.length < 10) return false;
  try {
    const user = getUserFromToken(token);
    return user.username && user.role;
  } catch(e) {
    return false;
  }
}

/* ====== DATA CRUD ====== */
function getSheetData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_DATA);
  if (!sh) {
    sh = ss.insertSheet(SHEET_DATA);
    sh.appendRow([
      "ID","NOP","Nama Pemohon","Jenis Layanan",
      "Sertifikat","NIB EL","Luas Tanah",
      "KoordinatLat","KoordinatLng",
      "PenyandingUtara","PenyandingTimur","PenyandingSelatan","PenyandingBarat",
      "Status","Catatan","Timestamp","UpdatedAt"
    ]);
  }
  return sh;
}

function getAllData() {
  const sh = getSheetData();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  return values.slice(1).map(r => ({
    id: r[0],
    nop: r[1],
    namaPemohon: r[2],
    jenisLayanan: r[3],
    sertifikat: r[4],
    nibEl: r[5],
    luasTanah: r[6],
    koordinatLat: r[7],
    koordinatLng: r[8],
    penyandingUtara: r[9],
    penyandingTimur: r[10],
    penyandingSelatan: r[11],
    penyandingBarat: r[12],
    status: r[13],
    catatan: r[14],
    timestamp: r[15] ? Utilities.formatDate(new Date(r[15]), "Asia/Jakarta", "dd/MM/yyyy HH:mm") : "",
    updatedAt: r[16] ? Utilities.formatDate(new Date(r[16]), "Asia/Jakarta", "dd/MM/yyyy HH:mm") : ""
  }));
}

function createData(d) {
  const sh = getSheetData();
  const id = "NOP-" + new Date().getTime();
  sh.appendRow([
    id, d.nop, d.namaPemohon, d.jenisLayanan,
    d.sertifikat, d.nibEl, d.luasTanah,
    d.koordinatLat, d.koordinatLng,
    d.penyandingUtara, d.penyandingTimur, d.penyandingSelatan, d.penyandingBarat,
    d.status || "Baru", d.catatan, new Date(), ""
  ]);
  return id;
}

function updateData(d) {
  const sh = getSheetData();
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] == d.id) {
      sh.getRange(i+1, 1, 1, 16).setValues([[
        d.id, d.nop, d.namaPemohon, d.jenisLayanan,
        d.sertifikat, d.nibEl, d.luasTanah,
        d.koordinatLat, d.koordinatLng,
        d.penyandingUtara, d.penyandingTimur, d.penyandingSelatan, d.penyandingBarat,
        d.status, d.catatan, values[i][15], new Date()
      ]]);
      return "✅ Data berhasil diperbarui";
    }
  }
  throw new Error("Data tidak ditemukan");
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}