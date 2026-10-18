// MOSTLY2STROKE (M2S) — AUTHENTICATED GOOGLE APPS SCRIPT BACKEND
// Copy this file into the Apps Script project only after reviewing the deployment checklist.
// Configure Script Properties: M2S_ADMIN_USERNAME and M2S_ADMIN_PASSWORD.
// Do not put credentials in index.html or in GitHub.

var SHEET_NAME = "KAMAR";
var TOKEN_TTL_SECONDS = 21600; // 6 hours; CacheService maximum is 6 hours.

function doGet(e) {
  // Deliberately do not expose rider data via unauthenticated GET requests.
  return jsonOutput_({ status: "error", message: "Endpoint tidak tersedia. Gunakan login." });
}

function doPost(e) {
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    var action = String(data.action || "");

    if (action === "login") return login_(data);
    if (action === "logout") return logout_(data);
    if (action === "getData") {
      if (!isValidToken_(data.token)) return unauthorized_();
      return getData_();
    }
    if (action === "checkIn") {
      if (!isValidToken_(data.token)) return unauthorized_();
      return processCheckIn_(data);
    }
    return jsonOutput_({ status: "error", message: "Aksi tidak dikenal." });
  } catch (err) {
    return jsonOutput_({ status: "error", message: "Permintaan tidak dapat diproses." });
  }
}

function login_(data) {
  var props = PropertiesService.getScriptProperties();
  var username = props.getProperty("M2S_ADMIN_USERNAME");
  var password = props.getProperty("M2S_ADMIN_PASSWORD");
  if (!username || !password) {
    return jsonOutput_({ status: "error", message: "Konfigurasi akun belum lengkap. Hubungi administrator." });
  }

  // Basic server-side throttling per username. This is not a substitute for a managed identity provider.
  var cache = CacheService.getScriptCache();
  var attemptKey = "login_attempts_" + digest_(String(data.username || "").toLowerCase());
  var attempts = Number(cache.get(attemptKey) || "0");
  if (attempts >= 8) {
    return jsonOutput_({ status: "error", message: "Terlalu banyak percobaan. Coba lagi setelah 15 menit." });
  }

  if (String(data.username || "") !== username || String(data.password || "") !== password) {
    cache.put(attemptKey, String(attempts + 1), 900);
    return jsonOutput_({ status: "error", message: "Username atau password salah." });
  }

  cache.remove(attemptKey);
  var token = Utilities.getUuid() + Utilities.getUuid();
  var tokenKey = "m2s_session_" + digest_(token);
  cache.put(tokenKey, JSON.stringify({ user: username, issuedAt: Date.now() }), TOKEN_TTL_SECONDS);
  return jsonOutput_({
    status: "success",
    token: token,
    expiresIn: TOKEN_TTL_SECONDS,
    message: "Login berhasil."
  });
}

function isValidToken_(token) {
  if (!token || typeof token !== "string" || token.length > 200) return false;
  var value = CacheService.getScriptCache().get("m2s_session_" + digest_(token));
  if (!value) return false;
  try {
    var session = JSON.parse(value);
    return !!session.user && (Date.now() - Number(session.issuedAt)) < TOKEN_TTL_SECONDS * 1000;
  } catch (err) {
    return false;
  }
}

function logout_(data) {
  var token = data.token;
  if (token && typeof token === "string" && token.length <= 200) {
    CacheService.getScriptCache().remove("m2s_session_" + digest_(token));
  }
  return jsonOutput_({ status: "success", message: "Sesi berakhir." });
}

function getData_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return jsonOutput_({ status: "error", message: "Sheet tidak ditemukan." });

  var rows = sheet.getDataRange().getValues();
  if (!rows.length) return jsonOutput_({ status: "success", data: [] });

  var headers = rows[0];
  var result = [];
  for (var i = 2; i < rows.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var value = rows[i][j];
      obj[headers[j]] = value instanceof Date
        ? Utilities.formatDate(value, "Asia/Jakarta", "yyyy-MM-dd'T'HH:mm:ss")
        : value;
    }
    obj.rowIndex = i + 1;
    result.push(obj);
  }
  return jsonOutput_({ status: "success", data: result });
}

function processCheckIn_(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return jsonOutput_({ status: "error", message: "Sheet tidak ditemukan." });

  var rows = sheet.getDataRange().getValues();
  if (!rows.length) return jsonOutput_({ status: "error", message: "Data peserta tidak tersedia." });
  var headers = rows[0];

  var colKodeReg = headers.indexOf("Kode Registrasi");
  var colKehadiran = headers.indexOf("Kehadiran");
  var colTimestamp = headers.indexOf("Timestamp");
  var colDiambilOleh = headers.indexOf("Diambil oleh");
  if (colKodeReg === -1 || colKehadiran === -1) {
    return jsonOutput_({ status: "error", message: "Kolom Kode Registrasi atau Kehadiran tidak ditemukan." });
  }

  var targetKode = String(data.kodeRegistrasi || "").trim().toLowerCase();
  if (!targetKode) return jsonOutput_({ status: "error", message: "Kode registrasi wajib diisi." });

  var targetRow = -1;
  for (var i = 2; i < rows.length; i++) {
    if (String(rows[i][colKodeReg]).trim().toLowerCase() === targetKode) {
      targetRow = i + 1;
      break;
    }
  }
  // Never trust client-supplied rowIndex as a fallback.
  if (targetRow === -1) return jsonOutput_({ status: "error", message: "Data peserta tidak ditemukan." });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var currentStatus = String(sheet.getRange(targetRow, colKehadiran + 1).getValue() || "").trim().toLowerCase();
    if (currentStatus === "hadir") {
      return jsonOutput_({ status: "error", message: "Peserta sudah check-in." });
    }

    var now = new Date();
    sheet.getRange(targetRow, colKehadiran + 1).setValue("Hadir");
    if (colTimestamp !== -1) sheet.getRange(targetRow, colTimestamp + 1).setValue(
      Utilities.formatDate(now, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss")
    );
    if (colDiambilOleh !== -1) {
      sheet.getRange(targetRow, colDiambilOleh + 1).setValue(String(data.diambilOleh || "").slice(0, 150));
    }
    return jsonOutput_({ status: "success", message: "Check-in berhasil." });
  } finally {
    lock.releaseLock();
  }
}

function digest_(value) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/g, "");
}

function unauthorized_() {
  return jsonOutput_({ status: "error", code: "UNAUTHORIZED", message: "Sesi tidak valid atau sudah kedaluwarsa. Silakan login kembali." });
}

function jsonOutput_(object) {
  return ContentService.createTextOutput(JSON.stringify(object))
    .setMimeType(ContentService.MimeType.JSON);
}

/*
SETUP CHECKLIST
1. In Apps Script > Project Settings > Script Properties, add:
   M2S_ADMIN_USERNAME = username pilihanmu
   M2S_ADMIN_PASSWORD = password unik yang panjang (never commit either value)
2. Replace the old backend code with this code (avoid duplicate function names).
3. Deploy > Manage deployments > Edit Web App > Execute as: Me.
   Restrict who can access the deployment as tightly as the app's use case permits.
   If anonymous/public access is required for GitHub Pages, the script still has a public URL;
   this code protects data actions with the token, but public Apps Script endpoints are not equivalent
   to a managed identity provider. Use HTTPS and never put secrets in frontend source.
4. Update the frontend to POST JSON actions: login {action,username,password};
   getData {action:"getData",token}; checkIn {action:"checkIn",token,...};
   logout {action:"logout",token}. Do not call getData with GET.
5. Test direct GET, getData without token, checkIn without token, expired token, wrong password,
   successful login, logout, and repeated check-in before announcing this is secured.
*/