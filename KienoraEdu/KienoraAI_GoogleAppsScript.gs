/* ==========================================================
   KIENORAAI - GOOGLE APPS SCRIPT BACKEND
   Dùng cho Web/PWA KienoraAI.
   Sheet chính: Sheet1  -> Q | A | FileName | FileUrl
   Tài khoản: Accounts -> Stt | Username | Password | Role |
              Email | Cá nhân phản hồi | Ngày cá nhân phản hồi | Status
   ========================================================== */

const SS = SpreadsheetApp.getActiveSpreadsheet();

function outputJSON(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doGet(e) {
  try {
    const action = String(e?.parameter?.action || "").trim();
    const username = String(e?.parameter?.username || "").trim();

    if (action === "get_data") {
      const shared = getKnowledgeData("Sheet1");
      const personal = username ? getKnowledgeData("user_" + username) : [];
      return outputJSON({
        status: "success",
        sheet1: shared,
        personal: personal,
        data: shared
      });
    }

    if (action === "search") {
      const keyword = String(e?.parameter?.keyword || "").trim();
      return outputJSON({ status: "success", total: searchKnowledge(keyword).length, data: searchKnowledge(keyword) });
    }

    return outputJSON({ status: "success", message: "KienoraAI API is running." });
  } catch (err) {
    return outputJSON({ status: "error", message: String(err) });
  }
}

function doPost(e) {
  try {
    const requestData = JSON.parse(e?.postData?.contents || "{}");
    const action = String(requestData.action || "").trim();

    switch (action) {
      case "register": return registerUser(requestData);
      case "login": return loginUser(requestData);
      case "teach":
      case "teach_with_file": return teachKnowledge(requestData);
      case "get_data": return getDataPost(requestData);
      case "search": return searchPost(requestData);
      case "ask": return askKnowledge(requestData);
      case "log_missing": return logMissing(requestData);
      case "sync_shared": return syncShared(requestData);
      case "delete_personal": return deletePersonal(requestData);
      case "admin_cmd": return adminCommand(requestData);
      case "get_math_knowledge": return getMathKnowledgeBase(requestData.filter || "all", requestData.keyword || "");
      default: return outputJSON({ status: "error", message: "Invalid action: " + action });
    }
  } catch (err) {
    return outputJSON({ status: "error", message: String(err) });
  }
}

/* ================= DATA ================= */

function getKnowledgeData(sheetName) {
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const lastColumn = Math.max(4, sheet.getLastColumn());
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastColumn).getValues();

  return rows.map(row => ({
    q: String(row[0] ?? "").trim(),
    a: String(row[1] ?? "").trim(),
    fileName: String(row[2] ?? "").trim(),
    fileUrl: String(row[3] ?? "").trim()
  })).filter(item => item.q || item.a);
}

function ensureKnowledgeSheet(sheetName) {
  let sheet = SS.getSheetByName(sheetName);
  if (!sheet) sheet = SS.insertSheet(sheetName);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([["Q", "A", "FileName", "FileUrl"]]);
  }
  return sheet;
}

function writeKnowledgeData(sheetName, data) {
  const sheet = ensureKnowledgeSheet(sheetName);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 4).setValues([["Q", "A", "FileName", "FileUrl"]]);

  if (!Array.isArray(data) || !data.length) return;

  const values = data.map(item => [
    String(item.q ?? item.Q ?? "").trim(),
    String(item.a ?? item.A ?? "").trim(),
    String(item.fileName ?? item.FileName ?? "").trim(),
    String(item.fileUrl ?? item.FileUrl ?? "").trim()
  ]);

  sheet.getRange(2, 1, values.length, 4).setValues(values);
}

function getDataPost(requestData) {
  const username = String(requestData.username || "").trim();
  const shared = getKnowledgeData("Sheet1");
  const personal = username ? getKnowledgeData("user_" + username) : [];
  return outputJSON({ status: "success", sheet1: shared, personal: personal, data: shared });
}

/* ================= ACCOUNT ================= */

function ensureAccountsSheet() {
  let sheet = SS.getSheetByName("Accounts");
  if (!sheet) {
    sheet = SS.insertSheet("Accounts");
    sheet.appendRow(["Stt", "Username", "Password", "Role", "Địa chỉ email", "Cá nhân phản hồi", "Ngày cá nhân phản hồi", "Status"]);
  }
  return sheet;
}

function registerUser(d) {
  const username = String(d.username || "").trim();
  const password = String(d.password || "").trim();
  const email = String(d.email || "").trim();

  if (!username || !password) return outputJSON({ status: "fail", message: "Thiếu tên đăng nhập hoặc mật khẩu." });

  const sheet = ensureAccountsSheet();
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1] || "").trim().toLowerCase() === username.toLowerCase()) {
      return outputJSON({ status: "fail", message: "Tên tài khoản đã tồn tại!" });
    }
  }

  sheet.appendRow([rows.length, username, password, "user", email, "", "", "Chờ duyệt"]);
  return outputJSON({ status: "success", message: "Đăng ký thành công! Vui lòng chờ quản trị viên duyệt." });
}

function loginUser(d) {
  const username = String(d.username || "").trim();
  const password = String(d.password || "").trim();
  const sheet = SS.getSheetByName("Accounts");

  if (!sheet) return outputJSON({ status: "fail", message: "Chưa có bảng Accounts." });

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const u = String(rows[i][1] || "").trim();
    const p = String(rows[i][2] || "").trim();
    if (u === username && p === password) {
      const role = String(rows[i][3] || "user").trim() || "user";
      const status = String(rows[i][7] || "").trim();
      if (role !== "admin" && status !== "Đã duyệt") {
        return outputJSON({ status: "fail", message: "Tài khoản đang chờ quản trị viên duyệt." });
      }
      return outputJSON({ status: "success", username: u, role: role, message: "Đăng nhập thành công!" });
    }
  }
  return outputJSON({ status: "fail", message: "Sai tên đăng nhập hoặc mật khẩu!" });
}

function isAdmin(username) {
  const name = String(username || "").trim();
  if (!name) return false;
  const sheet = SS.getSheetByName("Accounts");
  if (!sheet) return false;
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const u = String(rows[i][1] || "").trim();
    const role = String(rows[i][3] || "").trim().toLowerCase();
    const status = String(rows[i][7] || "").trim();
    if (u === name && role === "admin" && (!status || status === "Đã duyệt")) return true;
  }
  return false;
}

/* ================= TEACH ================= */

function teachKnowledge(d) {
  const username = String(d.username || "").trim();
  const q = String(d.q ?? d.question ?? "").trim();
  const a = String(d.a ?? d.answer ?? "").trim();
  const fileName = String(d.fileName || "").trim();
  const fileUrl = String(d.fileUrl || "").trim();

  if (!username) return outputJSON({ status: "fail", message: "Bạn cần đăng nhập để dạy KienoraAI." });
  if (!q || !a) return outputJSON({ status: "fail", message: "Câu hỏi và câu trả lời không được để trống." });
  if (fileUrl && !/^https?:\/\//i.test(fileUrl)) return outputJSON({ status: "fail", message: "Link tài liệu phải bắt đầu bằng http:// hoặc https://" });
  if (containsSensitiveWords(q + " " + a)) return outputJSON({ status: "fail", message: "Nội dung chứa từ khóa không được phép." });

  if (isAdmin(username)) {
    const data = getKnowledgeData("Sheet1");
    data.push({ q, a, fileName, fileUrl });
    writeKnowledgeData("Sheet1", data);
    return outputJSON({ status: "success", message: "Đã lưu trực tiếp vào kho tri thức chung KienoraAI." });
  }

  const personalName = "user_" + username;
  const personal = getKnowledgeData(personalName);
  personal.push({ q, a, fileName, fileUrl });
  writeKnowledgeData(personalName, personal);

  let pending = SS.getSheetByName("TriThucChoDuyet");
  if (!pending) {
    pending = SS.insertSheet("TriThucChoDuyet");
    pending.appendRow(["question", "answer", "fileName", "fileUrl", "Người dùng", "Duyệt/Chờ duyệt", "Thời gian"]);
  }
  pending.appendRow([q, a, fileName, fileUrl, username, "Chờ duyệt", new Date()]);

  return outputJSON({ status: "success", message: "Đã lưu vào kho cá nhân và gửi yêu cầu kiểm duyệt." });
}

/* ================= SEARCH / ASK ================= */

function normalizeSearchText(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function searchKnowledge(keyword) {
  const text = normalizeSearchText(keyword);
  const list = getKnowledgeData("Sheet1");
  if (!text) return list;

  const exact = [];
  const partial = [];
  const scored = [];
  const words = [...new Set(text.split(" ").filter(w => w.length >= 2))];

  list.forEach(item => {
    const q = normalizeSearchText(item.q);
    const a = normalizeSearchText(item.a);
    if (q === text) exact.push(item);
    else if (q.includes(text) || text.includes(q) || a.includes(text)) partial.push(item);
    else {
      let matched = 0;
      words.forEach(w => { if (q.includes(w) || a.includes(w)) matched++; });
      if (matched >= 2) scored.push({ item: item, score: matched / Math.max(words.length, 1) });
    }
  });

  scored.sort((x, y) => y.score - x.score);
  return exact.concat(partial, scored.map(x => x.item));
}

function askKnowledge(d) {
  const question = String(d.question || d.q || d.keyword || "").trim();
  const username = String(d.username || "").trim();
  if (!question) return outputJSON({ status: "fail", message: "Bạn chưa nhập câu hỏi." });

  if (username) {
    const personal = getKnowledgeData("user_" + username);
    const p = findBestKnowledge(personal, question);
    if (p) return outputJSON({ status: "success", source: "personal", data: p });
  }

  const shared = findBestKnowledge(getKnowledgeData("Sheet1"), question);
  if (shared) return outputJSON({ status: "success", source: "shared", data: shared });

  return outputJSON({ status: "not_found", message: "KienoraAI chưa tìm thấy tri thức phù hợp." });
}

function findBestKnowledge(list, question) {
  const text = normalizeSearchText(question);
  if (!text) return null;

  let best = null;
  let bestScore = 0;
  const words = [...new Set(text.split(" ").filter(w => w.length >= 2))];

  list.forEach(item => {
    const q = normalizeSearchText(item.q);
    if (!q) return;
    if (q === text) { best = item; bestScore = 1; return; }
    if (text.includes(q) || q.includes(text)) {
      const score = Math.min(q.length, text.length) / Math.max(q.length, text.length);
      if (score > bestScore) { best = item; bestScore = score; }
      return;
    }
    let matched = 0;
    words.forEach(w => { if (q.includes(w)) matched++; });
    const score = matched / Math.max(words.length, 1);
    if (matched >= 2 && score >= 0.4 && score > bestScore) { best = item; bestScore = score; }
  });
  return best;
}

function searchPost(d) {
  const data = searchKnowledge(d.keyword || d.question || "");
  return outputJSON({ status: "success", total: data.length, data: data });
}

/* ================= MISSING QUESTIONS ================= */

function logMissing(d) {
  const sheet = SS.getSheetByName("CauHoiCho") || SS.insertSheet("CauHoiCho");
  if (sheet.getLastRow() === 0) sheet.appendRow(["Thời gian", "Người hỏi", "Câu hỏi thiếu", "Trạng thái"]);
  sheet.appendRow([new Date(), String(d.username || "guest"), String(d.missingQuery || d.question || ""), "Chưa xử lý"]);
  return outputJSON({ status: "success", message: "Đã ghi nhận câu hỏi thiếu." });
}

/* ================= ADMIN SYNC ================= */

function syncShared(d) {
  const username = String(d.username || "").trim();
  if (!isAdmin(username)) return outputJSON({ status: "fail", message: "⛔ Chỉ Quản trị viên mới được chỉnh sửa kho tri thức chung." });
  if (!Array.isArray(d.data)) return outputJSON({ status: "fail", message: "Dữ liệu không hợp lệ." });
  writeKnowledgeData("Sheet1", d.data);
  return outputJSON({ status: "success", message: "Đã cập nhật kho tri thức chung. Tổng số: " + d.data.length + " mục." });
}

/* ================= DELETE PERSONAL ================= */

function deletePersonal(d) {
  const username = String(d.username || "").trim();
  const index = Number(d.index);
  if (!username || !Number.isInteger(index) || index < 0) return outputJSON({ status: "fail", message: "Thông tin xóa không hợp lệ." });

  const name = "user_" + username;
  const data = getKnowledgeData(name);
  if (index >= data.length) return outputJSON({ status: "fail", message: "Dữ liệu cần xóa không tồn tại." });
  data.splice(index, 1);
  writeKnowledgeData(name, data);
  return outputJSON({ status: "success", message: "Đã xóa dữ liệu khỏi kho cá nhân.", data: data });
}

/* ================= ADMIN COMMANDS ================= */

function adminCommand(d) {
  const username = String(d.username || "").trim();
  const command = String(d.command || "").trim();
  if (!isAdmin(username)) return outputJSON({ status: "fail", message: "Bạn không có quyền quản trị." });

  const parts = command.split(/\s+/);
  const cmd = (parts[0] || "").toLowerCase();
  const target = parts.slice(1).join(" ").trim();

  if (cmd === "/ds_tt" || cmd === "/duyet" || cmd === "/pending") return listPending();
  if (cmd === "/duyet_tt" || cmd === "/accept") return approvePending(Number(parts[1]));
  if (cmd === "/xoa_tt" || cmd === "/reject") return rejectPending(Number(parts[1]));
  if (cmd === "/ds_tk" || cmd === "/users") return listAccounts();
  if (cmd === "/duyet_tk") return approveAccount(parts[1]);
  if (cmd === "/khoa_tk" || cmd === "/xoa_tk") return lockAccount(parts[1]);
  if (cmd === "/backup") return outputJSON({ status: "success", message: "Hãy dùng nút Tải JSON trong kho tri thức để sao lưu." });
  if (cmd === "/logs") return outputJSON({ status: "success", message: "Nhật ký chi tiết có thể được mở rộng từ sheet Logs." });

  return outputJSON({ status: "fail", message: "Lệnh quản trị không hợp lệ." });
}

function listPending() {
  const sheet = SS.getSheetByName("TriThucChoDuyet");
  if (!sheet || sheet.getLastRow() < 2) return outputJSON({ status: "success", message: "✨ Không có tri thức đang chờ duyệt." });
  const rows = sheet.getDataRange().getValues();
  let out = "📋 TRI THỨC CHỜ DUYỆT:\n";
  let count = 0;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][5] || "Chờ duyệt") !== "Chờ duyệt") continue;
    const r = i + 1;
    out += `\n🔹 [Dòng ${r}] ${rows[i][4] || ""}\n   • Hỏi: ${rows[i][0] || ""}\n   • Đáp: ${rows[i][1] || ""}\n`;
    if (rows[i][2]) out += `   • Tệp: ${rows[i][2]}\n   • Link: ${rows[i][3] || ""}\n`;
    out += `   👉 /duyet_tt ${r}\n   👉 /xoa_tt ${r}\n`;
    count++;
  }
  return outputJSON({ status: "success", message: count ? out : "✨ Không có tri thức đang chờ duyệt." });
}

function approvePending(rowIndex) {
  const sheet = SS.getSheetByName("TriThucChoDuyet");
  if (!sheet || !Number.isInteger(rowIndex) || rowIndex < 2 || rowIndex > sheet.getLastRow()) return outputJSON({ status: "fail", message: "Dòng tri thức không hợp lệ." });
  const row = sheet.getRange(rowIndex, 1, 1, 6).getValues()[0];
  const data = getKnowledgeData("Sheet1");
  data.push({ q: row[0], a: row[1], fileName: row[2], fileUrl: row[3] });
  writeKnowledgeData("Sheet1", data);
  sheet.deleteRow(rowIndex);
  return outputJSON({ status: "success", message: "✅ Đã phê duyệt tri thức vào kho chung." });
}

function rejectPending(rowIndex) {
  const sheet = SS.getSheetByName("TriThucChoDuyet");
  if (!sheet || !Number.isInteger(rowIndex) || rowIndex < 2 || rowIndex > sheet.getLastRow()) return outputJSON({ status: "fail", message: "Dòng tri thức không hợp lệ." });
  sheet.deleteRow(rowIndex);
  return outputJSON({ status: "success", message: "🗑️ Đã xóa tri thức khỏi danh sách chờ duyệt." });
}

function listAccounts() {
  const sheet = SS.getSheetByName("Accounts");
  if (!sheet || sheet.getLastRow() < 2) return outputJSON({ status: "success", message: "Chưa có tài khoản." });
  const rows = sheet.getDataRange().getValues();
  let out = "👥 DANH SÁCH TÀI KHOẢN:\n";
  for (let i = 1; i < rows.length; i++) out += `• ${rows[i][1] || ""} | Role: ${rows[i][3] || "user"} | ${rows[i][7] || ""}\n`;
  return outputJSON({ status: "success", message: out });
}

function approveAccount(username) {
  username = String(username || "").trim();
  const sheet = SS.getSheetByName("Accounts");
  if (!sheet || !username) return outputJSON({ status: "fail", message: "Cú pháp: /duyet_tk ten_user" });
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1] || "").trim() === username) {
      sheet.getRange(i + 1, 8).setValue("Đã duyệt");
      ensureKnowledgeSheet("user_" + username);
      return outputJSON({ status: "success", message: "✅ Đã duyệt tài khoản " + username + "." });
    }
  }
  return outputJSON({ status: "fail", message: "Không tìm thấy tài khoản " + username + "." });
}

function lockAccount(username) {
  username = String(username || "").trim();
  const sheet = SS.getSheetByName("Accounts");
  if (!sheet || !username) return outputJSON({ status: "fail", message: "Cú pháp: /khoa_tk ten_user" });
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1] || "").trim() === username) {
      sheet.getRange(i + 1, 8).setValue("Đã khóa/Từ chối");
      return outputJSON({ status: "success", message: "🔒 Đã khóa tài khoản " + username + "." });
    }
  }
  return outputJSON({ status: "fail", message: "Không tìm thấy tài khoản " + username + "." });
}

/* ================= SAFETY / MATH ================= */

function containsSensitiveWords(text) {
  const sheet = SS.getSheetByName("TuKhoaCam");
  if (!sheet || sheet.getLastRow() < 2) return false;
  const lower = String(text || "").toLowerCase();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const bad = String(rows[i][0] || "").trim().toLowerCase();
    if (bad && lower.includes(bad)) return true;
  }
  return false;
}

function getMathKnowledgeBase(filterType, searchKeyword) {
  const sheet = SS.getSheetByName("ToanHoc_1000");
  if (!sheet) return outputJSON({ status: "fail", message: "Chưa tìm thấy sheet ToanHoc_1000." });
  const rows = sheet.getDataRange().getValues();
  const keyword = normalizeSearchText(searchKeyword);
  const results = [];
  for (let i = 1; i < rows.length; i++) {
    const q = String(rows[i][0] || "");
    const a = String(rows[i][1] || "");
    const fileName = String(rows[i][2] || "").trim();
    const fileUrl = String(rows[i][3] || "").trim();
    if (filterType === "has_file" && !fileName) continue;
    if (keyword && !normalizeSearchText(q).includes(keyword) && !normalizeSearchText(a).includes(keyword)) continue;
    results.push({ rowIndex: i + 1, q: q, a: a, fileName: fileName, fileUrl: fileUrl });
  }
  return outputJSON({ status: "success", total: results.length, data: results });
}

function onEdit(e) {
  try {
    const range = e.range;
    const sheet = range.getSheet();
    if (sheet.getName() !== "Accounts" || range.getColumn() !== 8 || e.value !== "Đã duyệt") return;
    const username = String(sheet.getRange(range.getRow(), 2).getValue() || "").trim();
    if (username) ensureKnowledgeSheet("user_" + username);
  } catch (err) {}
}
