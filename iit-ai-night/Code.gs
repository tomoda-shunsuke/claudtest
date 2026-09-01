/**
 * IIT AIナイト（事例共有＋ピザ）参加受付 — 保存用バックエンド
 *
 * 使い方は README.md を参照してください。
 * Googleスプレッドシートに紐づけて「ウェブアプリ」としてデプロイすると、
 * 申込フォームからの登録と、管理画面からの一覧取得・更新を処理します。
 */

/* ===================== 設定 ===================== */
var ADMIN_PASSWORD = "ここに管理画面のパスワードを設定";  // 管理画面ログイン用（このファイル内だけに置きます）
var CAPACITY       = 25;      // 定員（名）
var PRE_BOOKED     = 0;       // このフォーム以外で既に受け付けている人数
var FEE            = 1000;    // お一人あたりの参加費（円）
var NOTIFY_TO      = "";      // 申込が入ったときの通知先。空なら通知しません
var SEND_CONFIRM   = true;    // 申込者へ受付確認メールを送るか
var SHEET_NAME     = "申込";

/* ===================== シート定義 ===================== */
var COLS = [
  ["id",        "受付番号"],
  ["createdAt", "申込日時"],
  ["status",    "ステータス"],
  ["company",   "会社名"],
  ["name",      "お名前"],
  ["kana",      "ふりがな"],
  ["title",     "部署・役職"],
  ["email",     "メールアドレス"],
  ["tel",       "電話番号"],
  ["kubun",     "区分"],
  ["share",     "事例共有"],
  ["caseText",  "事例の内容"],
  ["konshin",   "懇親"],
  ["drink",     "飲み物"],
  ["guests",    "お連れ様人数"],
  ["guestNames","お連れ様"],
  ["allergy",   "アレルギー等"],
  ["memo",      "備考"]
];

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(COLS.map(function (c) { return c[1]; }));
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLS.length).setFontWeight("bold");
  }
  return sh;
}

function readAll_() {
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, COLS.length).getValues();
  return values.map(function (row) {
    var o = {};
    COLS.forEach(function (c, i) { o[c[0]] = row[i]; });
    o.guests = Number(o.guests) || 0;
    if (o.createdAt instanceof Date) o.createdAt = o.createdAt.toISOString();
    return o;
  }).filter(function (o) { return o.id; });
}

function rowOf_(id) {
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function headcount_(r) { return 1 + (Number(r.guests) || 0); }

function seats_() {
  var people = readAll_()
    .filter(function (r) { return r.status !== "キャンセル"; })
    .reduce(function (n, r) { return n + headcount_(r); }, 0);
  return { people: people, left: CAPACITY - PRE_BOOKED - people, capacity: CAPACITY };
}

/* ===================== 入口 ===================== */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "count";
  if (action === "count") {
    var s = seats_();
    return json_({ ok: true, people: s.people, left: s.left, capacity: s.capacity });
  }
  return json_({ ok: false, reason: "unknown_action" });
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, reason: "bad_request" }); }

  var action = req.action;

  if (action === "add")    return json_(addEntry_(req.entry));
  if (action === "login")  return json_({ ok: req.pw === ADMIN_PASSWORD });

  /* ここから先は管理者のみ */
  if (req.pw !== ADMIN_PASSWORD) return json_({ ok: false, reason: "auth" });

  if (action === "list")   return json_({ ok: true, entries: readAll_(), seats: seats_() });
  if (action === "update") return json_(updateEntry_(req.id, req.patch));
  if (action === "delete") return json_(deleteEntry_(req.id));

  return json_({ ok: false, reason: "unknown_action" });
}

/* ===================== 操作 ===================== */
function addEntry_(entry) {
  if (!entry || !entry.name || !entry.email) return { ok: false, reason: "bad_request" };

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (err) { return { ok: false, reason: "busy" }; }

  try {
    var want = headcount_(entry);
    var s = seats_();
    if (want > s.left) return { ok: false, reason: "full", left: s.left };

    entry.id = entry.id || ("e" + Date.now().toString(36));
    entry.createdAt = entry.createdAt || new Date().toISOString();
    entry.status = "受付済";

    sheet_().appendRow(COLS.map(function (c) { return entry[c[0]] === undefined ? "" : entry[c[0]]; }));
    notify_(entry);
    return { ok: true, id: entry.id, left: s.left - want };
  } finally {
    lock.releaseLock();
  }
}

function updateEntry_(id, patch) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return { ok: false, reason: "busy" }; }
  try {
    var row = rowOf_(id);
    if (row < 0) return { ok: false, reason: "not_found" };
    var sh = sheet_();
    COLS.forEach(function (c, i) {
      if (patch[c[0]] !== undefined) sh.getRange(row, i + 1).setValue(patch[c[0]]);
    });
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function deleteEntry_(id) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return { ok: false, reason: "busy" }; }
  try {
    var row = rowOf_(id);
    if (row < 0) return { ok: false, reason: "not_found" };
    sheet_().deleteRow(row);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ===================== メール ===================== */
function notify_(entry) {
  var body = COLS.map(function (c) {
    var v = entry[c[0]];
    return v === "" || v === undefined ? null : c[1] + "：" + v;
  }).filter(String).join("\n");

  var total = headcount_(entry);
  var detail = body +
    "\nご参加人数（合計）：" + total + "名" +
    "\n参加費：" + (total * FEE).toLocaleString() + "円（当日会場でお支払い）";

  if (NOTIFY_TO) {
    try {
      MailApp.sendEmail(NOTIFY_TO,
        "【IIT AIナイト】新しい参加申込：" + entry.company + " " + entry.name,
        detail);
    } catch (err) { /* 通知に失敗しても申込自体は成立させる */ }
  }

  if (SEND_CONFIRM && entry.email) {
    try {
      MailApp.sendEmail(entry.email,
        "【IIT AIナイト】参加申込を受け付けました",
        entry.name + " 様\n\n" +
        "IIT AIナイト（事例共有＋ピザ）へのお申し込みありがとうございます。\n" +
        "下記の内容で受け付けました。\n\n" +
        "日時：9月16日（水）17:00〜20:00\n" +
        "場所：株式会社チームテック 事務所（会議室）\n" +
        "　　　〒153-0064 東京都目黒区下目黒1-8-1 ARCO TOWER 6階\n" +
        "参加費：" + (total * FEE).toLocaleString() + "円（当日会場でお支払いください）\n\n" +
        "----------------------------------------\n" + detail + "\n" +
        "----------------------------------------\n\n" +
        "当日お会いできるのを楽しみにしています。\n" +
        "東京都情報産業協会 技術部会");
    } catch (err) { /* 送信できなくても申込は成立させる */ }
  }
}
