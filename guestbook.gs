// 방명록 저장소. 배포: 웹 앱, 실행: 나, 액세스: 모든 사용자.
// 먼저 setupGuestbook 를 한 번 실행해 시트를 만든 뒤 배포합니다.

var PRIVATE_NAME = "모바일청첩장 방명록";
var PUBLIC_NAME = "모바일청첩장 방명록 공개";
var PRIVATE_HEADERS = ["id", "name", "date", "message", "passwordHash"];
var PUBLIC_HEADERS = ["id", "name", "date", "message"];
var MAX_NOTES = 500;

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function props_() {
  return PropertiesService.getScriptProperties();
}

function hash_(password) {
  var salt = props_().getProperty("SALT");
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + "\n" + String(password),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function (b) {
    var v = (b + 256) % 256;
    return ("0" + v.toString(16)).slice(-2);
  }).join("");
}

function textColumn_(sheet) {
  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 2), 1).setNumberFormat("@");
}

function ensureSheet_(ss, headers) {
  var sheet = ss.getSheets()[0];
  if (String(sheet.getRange(1, 1).getValue()) !== headers[0]) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  textColumn_(sheet);
  return sheet;
}

function openOrCreate_(key, name, headers) {
  var id = props_().getProperty(key);
  var ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create(name);
  if (!id) props_().setProperty(key, ss.getId());
  return ensureSheet_(ss, headers);
}

function privateSheet_() {
  return openOrCreate_("PRIVATE_SHEET_ID", PRIVATE_NAME, PRIVATE_HEADERS);
}

function publicSheet_() {
  var sheet = openOrCreate_("PUBLIC_SHEET_ID", PUBLIC_NAME, PUBLIC_HEADERS);
  var file = DriveApp.getFileById(props_().getProperty("PUBLIC_SHEET_ID"));
  if (file.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  return sheet;
}

function setupGuestbook() {
  if (!props_().getProperty("SALT")) props_().setProperty("SALT", Utilities.getUuid());
  privateSheet_();
  var sheet = publicSheet_();
  Logger.log(props_().getProperty("PUBLIC_SHEET_ID"));
  return sheet.getParent().getId();
}

function clip_(value, max) {
  return String(value || "").trim().slice(0, max);
}

function safe_(value) {
  var text = String(value || "");
  if (/^[=+\-@]/.test(text)) text = "'" + text;
  return text;
}

function list_() {
  var values = publicSheet_().getDataRange().getValues();
  var notes = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    notes.push({
      id: String(values[i][0]),
      name: String(values[i][1]),
      date: String(values[i][2]),
      message: String(values[i][3])
    });
  }
  return notes;
}

function findRow_(sheet, id) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1;
  }
  return 0;
}

function add_(data) {
  var name = clip_(data.name, 20);
  var message = clip_(data.message, 200);
  var password = String(data.password || "").trim().slice(0, 20);
  var date = clip_(data.date, 10);
  var id = clip_(data.id, 20);
  if (!name || !message || !password || !id) return { ok: false, error: "required" };
  if (data.website) return { ok: false, error: "spam" };
  var priv = privateSheet_();
  if (findRow_(priv, id)) return { ok: true, notes: list_() };
  if (priv.getLastRow() - 1 >= MAX_NOTES) return { ok: false, error: "full" };
  priv.appendRow([id, safe_(name), safe_(date), safe_(message), hash_(password)]);
  var pub = publicSheet_();
  pub.insertRowAfter(1);
  pub.getRange(2, 1, 1, 4).setValues([[id, safe_(name), safe_(date), safe_(message)]]);
  textColumn_(pub);
  textColumn_(priv);
  return { ok: true, notes: list_() };
}

function remove_(data) {
  var id = String(data.id || "");
  var password = String(data.password || "").trim();
  var priv = privateSheet_();
  var row = findRow_(priv, id);
  if (!row) return { ok: true, notes: list_() };
  var stored = String(priv.getRange(row, 5).getValue());
  if (hash_(password) !== stored) return { ok: false, error: "password" };
  priv.deleteRow(row);
  var pub = publicSheet_();
  var publicRow = findRow_(pub, id);
  if (publicRow) pub.deleteRow(publicRow);
  return { ok: true, notes: list_() };
}

function handle_(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return json_({ ok: false, error: "busy" });
  }
  try {
    if (!props_().getProperty("SALT")) props_().setProperty("SALT", Utilities.getUuid());
    var action = data.action || "list";
    if (action === "info") {
      publicSheet_();
      return json_({ ok: true, sheetId: props_().getProperty("PUBLIC_SHEET_ID") });
    }
    if (action === "list") return json_({ ok: true, notes: list_() });
    if (action === "add") return json_(add_(data));
    if (action === "delete") return json_(remove_(data));
    return json_({ ok: false, error: "unknown" });
  } catch (err) {
    return json_({ ok: false, error: "server" });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return handle_(e && e.parameter ? e.parameter : {});
}

function doPost(e) {
  var data = {};
  try {
    data = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : "{}");
  } catch (err) {
    data = e && e.parameter ? e.parameter : {};
  }
  return handle_(data);
}
