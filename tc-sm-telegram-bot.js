// ============================================================
// TC Content Bot — Google Apps Script (ES5 compatible)
// Telegram digest + inline buttons + outstanding tracker
// ============================================================
// ── ONE-TIME SETUP ────────────────────────────────────────────
// Run this function ONCE in the Apps Script editor to store the
// bot token securely. Delete or comment it out afterwards.
function setupScriptProperties() {
  PropertiesService.getScriptProperties().setProperty(
    "BOT_TOKEN", "PASTE_YOUR_NEW_BOT_TOKEN_HERE"
  );
  Logger.log("BOT_TOKEN set successfully.");
}
// ── CONFIG ────────────────────────────────────────────────────
var BOT_TOKEN      = PropertiesService.getScriptProperties().getProperty("BOT_TOKEN");
var SHEET_ID       = "1a2aASYC8qUd0knij8GmuAo-yZzRaqlTW4-j5Uhfam-g";
var CHAT_ID        = "-5256560725";
var SHEET_NAME     = "Posts";
var TRACKER_SHEET  = "Bot_Tracker";
var CONTENT_CAL_ID = "c_52223acb4630f56df988a7199f5719f2e72a3c7388d7e0a99a10691e916cb2a9@group.calendar.google.com";
var MORNING_HOUR   = 9;
var EVENING_HOUR   = 18;
var TIMEZONE       = "Asia/Singapore";
var EVENT_TAG      = "TC_BOT";
// ← Paste your Web App URL here
var WEBAPP_URL = "https://script.google.com/macros/s/AKfycbyNxunmVq-FOVb2cv3w67YlF3I-1U6MkrE1VGCDty_bqhOGp1jkj2LqU1Gzq73lnleB/exec";
// ── OWNER → EMAIL ─────────────────────────────────────────────
var OWNER_EMAILS = {
  "mavia":    "mavia.kow@tcacoustic.com.sg",
  "yov":      "yovanna.lo@tcacoustic.com.sg",
  "yovanna":  "yovanna.lo@tcacoustic.com.sg",
  "bibs":     "bibiane.chua@tcacoustic.com.sg",
  "bibiane":  "bibiane.chua@tcacoustic.com.sg",
  "haripria": "haripria.gunalan@tcacoustic.com.my",
  "hari":     "haripria.gunalan@tcacoustic.com.my",
  "jc":       "jayciakhophet@tcacoustic.co.th",
  "jaycia":   "jayciakhophet@tcacoustic.co.th",
  "ryn":      "rynrynth04@gmail.com",
  "min":      "soo.min@tcacoustic.com.sg"
};
// ── BRAND → EMOJI / SHORT NAME ────────────────────────────────
var BRAND_TO_EMOJI = {
  "tc acoustic":      "⬜",
  "sonos":            "🔵",
  "bowers & wilkins": "⚫",
  "marshall":         "🔴"
};
// ── HANDLE → DISPLAY LABEL ────────────────────────────────────
var HANDLE_LABEL = {
  "tcacoustic":             "TC",
  "tc_acoustic":            "TC",
  "friendsoftc":            "TC",
  "sonosconceptstore":      "SCS",
  "sonos.hk":               "HK",
  "sonoshongkong":          "HK",
  "sonosthailandofficial":  "TH",
  "sonosthailand":          "TH",
  "bowerswilkinssingapore": "B&W-SG",
  "bowerswilkinsmalaysia":  "B&W-MY",
  "thatsoundspot":          "TSS"
};
// ── HELPERS ───────────────────────────────────────────────────
function getBrandEmoji(brand) {
  if (!brand) return "📌";
  return BRAND_TO_EMOJI[brand.toLowerCase().trim()] || "📌";
}
function getHandleLabel(handle) {
  if (!handle) return handle;
  return HANDLE_LABEL[handle.toLowerCase().trim()] || handle;
}
function getTodayString() {
  return Utilities.formatDate(new Date(), TIMEZONE, "dd MMM yyyy");
}
function getYesterdayString() {
  var d = new Date();
  d.setDate(d.getDate() - 1);
  return Utilities.formatDate(d, TIMEZONE, "dd MMM yyyy");
}
function normaliseDateCell(raw) {
  if (!raw) return "";
  if (raw instanceof Date) return Utilities.formatDate(raw, TIMEZONE, "dd MMM yyyy");
  var d = new Date(raw);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, TIMEZONE, "dd MMM yyyy");
  return String(raw).trim();
}
function stripTime(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
function formatYMD(d) {
  return Utilities.formatDate(d, TIMEZONE, "yyyy-MM-dd");
}
function getOwnerEmail(owner) {
  if (!owner) return null;
  return OWNER_EMAILS[owner.toLowerCase().trim()] || null;
}
function parsePostingTime(timeStr, dateStr) {
  if (!timeStr) return null;
  var hour = -1, minute = 0;
  var match12 = timeStr.match(/^(\d{1,2})(?::(\d{2}))?\s*(AM|PM)$/i);
  var match24 = timeStr.match(/^(\d{1,2}):(\d{2})$/);
  if (match12) {
    hour   = parseInt(match12[1]);
    minute = match12[2] ? parseInt(match12[2]) : 0;
    var period = match12[3].toUpperCase();
    if (period === "PM" && hour !== 12) hour += 12;
    if (period === "AM" && hour === 12) hour = 0;
  } else if (match24) {
    hour   = parseInt(match24[1]);
    minute = parseInt(match24[2]);
  }
  if (hour < 0 || hour > 23) return null;
  var end   = new Date(dateStr + "T00:00:00");
  end.setHours(hour, minute, 0, 0);
  var start = new Date(end.getTime() - 30 * 60 * 1000);
  return { start: start, end: end };
}
// ── PARSE POST ROW ────────────────────────────────────────────
// Accounts cell format: "IG/SG @tcacoustic, FB/SG @tcacoustic, IG/MY @sonosconceptstore"
// Each entry: PLATFORM/MARKET @handle (optional suffix e.g. "(Video)", "(Short)")
// Returns an ARRAY — one entry per unique handle, platforms collapsed per handle.
// e.g. "IG/SG @tcacoustic, FB/SG @tcacoustic" → [{label:"TC", platforms:"IG/FB", market:"SG"}]
// e.g. "IG/SG @bowerswilkinssingapore, IG/MY @bowerswilkinsmalaysia" → two entries
function parsePost(row, COL) {
  var accountsRaw = String(row[COL.account] || "").trim();
  var title       = COL.title   > -1 ? String(row[COL.title]   || "").trim() : "";
  var pillar      = COL.pillar  > -1 ? String(row[COL.pillar]  || "").trim() : "";
  var brand       = COL.brand   > -1 ? String(row[COL.brand]   || "").trim() : "";
  var owner       = COL.owner   > -1 ? String(row[COL.owner]   || "").trim() : "";
  var creator     = COL.creator > -1 ? String(row[COL.creator] || "").trim() : "";
  var rawTime     = COL.time > -1 ? row[COL.time] : "";
  var timeStr     = "";
  if (rawTime instanceof Date && !isNaN(rawTime.getTime())) {
    // Suppress midnight — empty time cells in Sheets come back as 00:00
    if (rawTime.getHours() !== 0 || rawTime.getMinutes() !== 0) {
      timeStr = Utilities.formatDate(rawTime, TIMEZONE, "h:mm a");
    }
  } else if (rawTime) {
    timeStr = String(rawTime).trim();
  }
  if (!accountsRaw || !title) return [];
  // Parse each comma-separated entry, grouping platforms by handle
  var entries    = accountsRaw.split(",");
  var handleMap  = {}; // handle → { label, platforms[], market }
  var handleOrder = [];
  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i].trim();
    var m = entry.match(/^([A-Za-z0-9]+)\/([A-Z]{2,3})\s+@(\S+)/i);
    if (!m) continue;
    var plat = m[1].trim().toUpperCase();
    var mkt  = m[2].trim().toUpperCase();
    var hndl = m[3].trim().toLowerCase().replace(/[()]/g, "");
    // Detect YT content type from suffix
    if (plat === "YT") {
      if (entry.toLowerCase().indexOf("short") !== -1) plat = "YT Short";
      else plat = "YT Video";
    }
    if (!handleMap[hndl]) {
      handleOrder.push(hndl);
      handleMap[hndl] = { label: getHandleLabel(hndl), platforms: [], market: mkt };
    }
    if (handleMap[hndl].platforms.indexOf(plat) === -1) handleMap[hndl].platforms.push(plat);
  }
  if (!handleOrder.length) return [];
  var results = [];
  for (var j = 0; j < handleOrder.length; j++) {
    var hndl2 = handleOrder[j];
    var h     = handleMap[hndl2];
    results.push({
      handle:      hndl2,
      label:       h.label,
      platform:    h.platforms.join("/"),
      postName:    title,
      postingTime: timeStr,
      creator:     creator,
      owner:       owner,
      pillar:      pillar,
      brand:       brand,
      market:      h.market,
      emoji:       getBrandEmoji(brand)
    });
  }
  return results;
}
function getColumnIndices(headers) {
  var COL = { posted: -1, date: -1, time: -1, account: -1, title: -1, pillar: -1, brand: -1, owner: -1, creator: -1 };
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (h === "posted")   COL.posted  = i;
    if (h === "date")     COL.date    = i;
    if (h === "time")     COL.time    = i;
    if (h === "accounts") COL.account = i;
    if (h === "title")    COL.title   = i;
    if (h === "pillar")   COL.pillar  = i;
    if (h === "brand")    COL.brand   = i;
    if (h === "owner")    COL.owner   = i;
    if (h === "creator")  COL.creator = i;
  }
  return COL;
}
// ── FETCH POSTS ───────────────────────────────────────────────
function getPostsForDate(dateStr) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');
  var data    = sheet.getDataRange().getValues();
  var headers = [];
  for (var h = 0; h < data[0].length; h++) headers.push(String(data[0][h]).trim().toLowerCase());
  var COL   = getColumnIndices(headers);
  var posts = [];
  for (var i = 1; i < data.length; i++) {
    if (normaliseDateCell(data[i][COL.date]) !== dateStr) continue;
    var parsed   = parsePost(data[i], COL);
    var sheetRow = i + 1;
    var isPosted = COL.posted > -1 ? (data[i][COL.posted] === true) : false;
    for (var p = 0; p < parsed.length; p++) {
      parsed[p].sheetRow = sheetRow;
      parsed[p].isPosted = isPosted;
      posts.push(parsed[p]);
    }
  }
  return posts;
}
function getPostsForDateRange(startDate, endDate) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');
  var data    = sheet.getDataRange().getValues();
  var headers = [];
  for (var h = 0; h < data[0].length; h++) headers.push(String(data[0][h]).trim().toLowerCase());
  var COL   = getColumnIndices(headers);
  var posts = [];
  var start = stripTime(startDate);
  var end   = stripTime(endDate);
  for (var i = 1; i < data.length; i++) {
    var d = data[i][COL.date];
    if (!(d instanceof Date)) continue;
    var rowDate = stripTime(d);
    if (rowDate < start || rowDate > end) continue;
    var parsed   = parsePost(data[i], COL);
    var sheetRow = i + 1;
    var isPosted = COL.posted > -1 ? (data[i][COL.posted] === true) : false;
    for (var p = 0; p < parsed.length; p++) {
      parsed[p].sheetRow = sheetRow;
      parsed[p].isPosted = isPosted;
      parsed[p].date     = rowDate;
      parsed[p].dateStr  = formatYMD(rowDate);
      posts.push(parsed[p]);
    }
  }
  return posts;
}
// ── COLLAPSE PLATFORMS ────────────────────────────────────────
function collapsePlatforms(posts) {
  var map   = {};
  var order = [];
  for (var i = 0; i < posts.length; i++) {
    var p   = posts[i];
    // Deduplicate by title+handle — if the same post appears twice in the sheet,
    // merge into one entry using the first occurrence's sheetRow as the stable key.
    // Collapse by label+market so different handles with the same display name
    // (e.g. @sonos.hk and @sonoshongkong both = "HK") merge into one line.
    var key = p.postName + "|" + p.label + "|" + p.market + "|" + p.owner;
    if (!map[key]) {
      order.push(key);
      map[key] = {
        handle:      p.handle,
        label:       p.label,
        platforms:   p.platform,
        postName:    p.postName,
        postingTime: p.postingTime || "",
        creator:     p.creator || "",
        owner:       p.owner,
        pillar:      p.pillar,
        brand:       p.brand,
        market:      p.market,
        emoji:       p.emoji,
        date:        p.date || null,
        dateStr:     p.dateStr || "",
        sheetRow:    p.sheetRow || null,
        isPosted:    p.isPosted || false
      };
    } else {
      if (map[key].platforms.indexOf(p.platform) === -1) map[key].platforms += "/" + p.platform;
      if (!map[key].market && p.market) map[key].market = p.market;
      if (!map[key].postingTime && p.postingTime) map[key].postingTime = p.postingTime;
    }
  }
  var result = [];
  for (var j = 0; j < order.length; j++) result.push(map[order[j]]);
  return result;
}
// ── GROUP BY OWNER ────────────────────────────────────────────
function groupByOwner(posts) {
  var map   = {};
  var order = [];
  for (var i = 0; i < posts.length; i++) {
    var owner = posts[i].owner || "\u2014";
    if (!map[owner]) { order.push(owner); map[owner] = []; }
    map[owner].push(posts[i]);
  }
  return { map: map, order: order };
}
// ── BUILD FLAT POST LIST ──────────────────────────────────────
// Stable key for a post — uses title slug + handle, so it survives row
// insertions/deletions in the Posts sheet (unlike sheetRow-based keys).
function postKey(p) {
  var titleSlug = (p.postName || "").toLowerCase()
    .replace(/[^a-z0-9]+/g, "-").slice(0, 40);
  return titleSlug + "|" + (p.handle || "");
}
// Returns all posts in display order.
function buildFlatPostList(grouped) {
  var flat = [];
  for (var i = 0; i < grouped.order.length; i++) {
    var owner = grouped.order[i];
    var posts = grouped.map[owner];
    for (var j = 0; j < posts.length; j++) flat.push(posts[j]);
  }
  return flat;
}
// ── TRACKER SHEET ─────────────────────────────────────────────
function ensureTrackerSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TRACKER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TRACKER_SHEET);
    sh.getRange(1, 1, 1, 6).setValues([["date", "message_id", "post_key", "status", "owner", "sheet_row"]]);
    sh.setFrozenRows(1);
  }
  return sh;
}
// flatPosts: ordered array from buildFlatPostList.
// initialStatuses: optional pre-seeded statuses keyed by postKey(p).
// Returns a rowIndices map { postKey: sheetRowNumber } for status cache use.
function saveTrackerEntry(dateStr, messageId, flatPosts, initialStatuses) {
  initialStatuses = initialStatuses || {};
  var sh       = ensureTrackerSheet();
  var startRow = sh.getLastRow() + 1;
  var rows     = [];
  var rowIndices = {};
  for (var i = 0; i < flatPosts.length; i++) {
    var p   = flatPosts[i];
    var key = postKey(p);
    var status = initialStatuses[key] || "pending";
    rows.push([dateStr, String(messageId), key, status, p.owner || "", p.sheetRow || ""]);
    rowIndices[key] = startRow + i;
  }
  if (rows.length) sh.getRange(startRow, 1, rows.length, 6).setValues(rows);
  return rowIndices;
}
// Returns Bot_Tracker statuses for a given date, keyed by sheet_row string.
// Used to determine outstanding status without touching the Posts sheet.
function getTrackerStatusesBySheetRow(dateStr) {
  var sh     = ensureTrackerSheet();
  var data   = sh.getDataRange().getValues();
  var result = {};
  for (var i = 1; i < data.length; i++) {
    if (normaliseDateCell(data[i][0]) !== dateStr) continue;
    var sr = String(data[i][5]);
    if (sr) result[sr] = String(data[i][3]);
  }
  return result;
}
// Batch-writes today's Bot_Tracker statuses back to the Posted column in Posts sheet.
// Run once at end of day — Bot_Tracker is the live source of truth, sheet is the record.
// Also repairs stale sheet_row values caused by sorts/insertions/deletions in Posts.
function syncPostedColumn() {
  try {
    var sh     = ensureTrackerSheet();
    var data   = sh.getDataRange().getValues();
    var today  = getTodayString();
    var ss     = SpreadsheetApp.openById(SHEET_ID);
    var sheet  = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return;

    // Read current Posts data and build column indices
    var postData    = sheet.getDataRange().getValues();
    var postHeaders = [];
    for (var h = 0; h < postData[0].length; h++) postHeaders.push(String(postData[0][h]).trim().toLowerCase());
    var COL  = getColumnIndices(postHeaders);
    var pCol = -1;
    for (var c = 0; c < postHeaders.length; c++) {
      if (postHeaders[c] === "posted") { pCol = c + 1; break; }
    }
    if (pCol === -1) return;

    // Build postKey → actual sheet row from current Posts layout
    // postKey format: titleSlug|handle  (matches postKey() helper)
    var rowLookup = {};
    if (COL.title > -1 && COL.account > -1) {
      for (var r = 1; r < postData.length; r++) {
        var title = String(postData[r][COL.title] || "").trim();
        if (!title) continue;
        var slug    = title.toLowerCase().replace(/[^a-z0-9]+/g, "-").slice(0, 40);
        var entries = String(postData[r][COL.account] || "").split(",");
        for (var ae = 0; ae < entries.length; ae++) {
          var m = entries[ae].trim().match(/@(\S+)/i);
          if (!m) continue;
          var handle = m[1].toLowerCase().replace(/[()]/g, "");
          rowLookup[slug + "|" + handle] = r + 1;
        }
      }
    }

    for (var i = 1; i < data.length; i++) {
      if (normaliseDateCell(data[i][0]) !== today) continue;
      var pkey      = String(data[i][2]);
      var lookedUp  = rowLookup[pkey];
      var actualRow = lookedUp || parseInt(data[i][5]);
      if (!actualRow) continue;

      // Repair stale sheet_row in tracker if the sort/edit shifted it
      if (lookedUp && lookedUp !== parseInt(data[i][5])) {
        sh.getRange(i + 1, 6).setValue(lookedUp);
      }

      sheet.getRange(actualRow, pCol).setValue(String(data[i][3]) === "posted");
    }
    Logger.log("syncPostedColumn done for " + today);
  } catch(err) { Logger.log("syncPostedColumn error: " + err.message); }
}
// ── BUILD POST LINE ───────────────────────────────────────────
// Format: • 9:00 PM: B&W-SG | IG/FB — Post Name (Creator)
function buildPostLine(p) {
  var line = "\u2022 ";
  if (p.postingTime) line += p.postingTime + ": ";
  line += p.label + " | " + p.platforms + " \u2014 " + p.postName;
  if (p.creator) line += " (" + p.creator + ")";
  return line + "\n";
}
// ── BUILD EVENING MESSAGE TEXT ────────────────────────────────
// statuses keyed by postKey(p) = "{sheetRow}|{handle}"
function buildEveningText(grouped, dateStr, statuses) {
  statuses = statuses || {};
  var msg = "";
  for (var i = 0; i < grouped.order.length; i++) {
    var owner      = grouped.order[i];
    var ownerPosts = grouped.map[owner];
    msg += "\ud83d\udc64 *" + owner + "*\n";
    for (var j = 0; j < ownerPosts.length; j++) {
      var status = statuses[postKey(ownerPosts[j])] || "pending";
      var prefix = status === "posted" ? "\u2705 " : "";
      msg += prefix + buildPostLine(ownerPosts[j]);
    }
    msg += "\n";
  }
  return msg.trim();
}
// ── BUILD INLINE KEYBOARD ─────────────────────────────────────
// flatPosts: ordered array from buildFlatPostList — one button per post
function buildKeyboard(flatPosts, statuses) {
  statuses = statuses || {};
  var buttons = [];
  for (var i = 0; i < flatPosts.length; i++) {
    var p      = flatPosts[i];
    var status = statuses[String(i)] || "pending";
    var label  = p.label + " | " + p.platforms + " \u2014 " + p.postName;
    if (status === "posted") label = "\u2705 " + label;
    buttons.push([{ text: label, callback_data: "toggle|" + i }]);
  }
  return { inline_keyboard: buttons };
}
// ── SEND / EDIT TELEGRAM ──────────────────────────────────────
function sendTelegramMessage(text) {
  var url     = "https://api.telegram.org/bot" + BOT_TOKEN + "/sendMessage";
  var payload = { chat_id: CHAT_ID, text: text, parse_mode: "Markdown" };
  var options = {
    method:             "post",
    contentType:        "application/json",
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var result = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
  if (!result.ok) throw new Error("Telegram error: " + result.description);
  return result;
}
function sendEveningMessage(text, keyboard) {
  var url     = "https://api.telegram.org/bot" + BOT_TOKEN + "/sendMessage";
  var payload = {
    chat_id:      CHAT_ID,
    text:         text,
    parse_mode:   "Markdown",
    reply_markup: keyboard
  };
  var options = {
    method:             "post",
    contentType:        "application/json",
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var result = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
  if (!result.ok) throw new Error("Telegram send error: " + result.description);
  return result;
}
function editTelegramMessage(messageId, text, keyboard) {
  var url     = "https://api.telegram.org/bot" + BOT_TOKEN + "/editMessageText";
  var payload = {
    chat_id:      CHAT_ID,
    message_id:   parseInt(messageId),
    text:         text,
    parse_mode:   "Markdown",
    reply_markup: keyboard
  };
  var options = {
    method:             "post",
    contentType:        "application/json",
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var result = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
  if (!result.ok) {
    // "message is not modified" is benign — happens on redundant double-taps
    if (result.description && result.description.indexOf("not modified") !== -1) return result;
    Logger.log("Edit error: " + result.description);
  }
  return result;
}
function getBotUsername() {
  var raw = PropertiesService.getScriptProperties().getProperty("BOT_USERNAME");
  if (raw) return raw;
  var result = JSON.parse(UrlFetchApp.fetch(
    "https://api.telegram.org/bot" + BOT_TOKEN + "/getMe"
  ).getContentText());
  var username = result.ok ? result.result.username : "";
  if (username) PropertiesService.getScriptProperties().setProperty("BOT_USERNAME", username);
  return username;
}
function answerCallbackQuery(callbackQueryId, text) {
  var url     = "https://api.telegram.org/bot" + BOT_TOKEN + "/answerCallbackQuery";
  var payload = { callback_query_id: callbackQueryId, text: text };
  var options = {
    method:             "post",
    contentType:        "application/json",
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };
  UrlFetchApp.fetch(url, options);
}
// ── BUILD FULL MESSAGE ────────────────────────────────────────
// Returns { text, keyboard }.
// pendingLines: pre-computed from buildYesterdayPendingLines — callers pass it in
// so button taps (doPost) can use a cached copy and skip sheet reads entirely.
function buildFullMessage(grouped, dateStr, statuses, pendingLines) {
  statuses     = statuses     || {};
  pendingLines = pendingLines || [];
  // ── Build text ────────────────────────────────────────────
  var lines = ["\ud83d\udcc5 *Today's Posts \u2014 " + dateStr + "*"];
  if (pendingLines.length > 0) {
    lines.push("");
    lines.push("\ud83d\udccb *Outstanding from yesterday*");
    for (var pl = 0; pl < pendingLines.length; pl++) lines.push(pendingLines[pl]);
    lines.push("\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500");
  }
  lines.push("");
  lines.push(buildEveningText(grouped, dateStr, statuses));
  // ── Build keyboard — today's posts only ──────────────────
  var buttons   = [];
  var flatPosts = buildFlatPostList(grouped);
  for (var tb = 0; tb < flatPosts.length; tb++) {
    var p      = flatPosts[tb];
    var status = statuses[postKey(p)] || "pending";
    var tlbl       = p.label + " | " + p.platforms + " \u2014 " + p.postName;
    if (status === "posted") tlbl = "\u2705 " + tlbl;
    // Use numeric index as callback_data — postKey can exceed Telegram's 64-byte limit.
    // doPost resolves the index back to postKey via the cached flat list.
    var nextAction = (status === "posted") ? "mark_pending" : "mark_posted";
    buttons.push([{ text: tlbl, callback_data: nextAction + "|" + tb }]);
  }
  return { text: lines.join("\n"), keyboard: { inline_keyboard: buttons } };
}
// ── BUILD PENDING REMINDER TEXT ───────────────────────────────
// Returns a simple count-only nudge. Returns null if everything is ticked or
// no morning message was sent yet.
function buildPendingReminderText(dateStr) {
  var sh     = ensureTrackerSheet();
  var dbData = sh.getDataRange().getValues();
  var total   = 0;
  var pending = 0;
  var hasEntry = false;
  for (var i = 1; i < dbData.length; i++) {
    if (normaliseDateCell(dbData[i][0]) !== dateStr) continue;
    hasEntry = true;
    total++;
    if (String(dbData[i][3]) !== "posted") pending++;
  }
  if (!hasEntry) return null; // morning message not sent yet
  if (pending === 0) return null; // all ticked — no reminder needed
  return "\ud83d\udd14 *Reminder \u2014 " + dateStr + "*\n" +
         pending + " post" + (pending === 1 ? "" : "s") + " still not ticked. Please post & tick!";
}
// ── POST CACHE ────────────────────────────────────────────────
// Today's grouped post structure is cached in Script Properties at morning send.
// doPost reads from cache — eliminates the Posts sheet read on every tap.
function cacheGrouped(dateStr, grouped) {
  var flat = buildFlatPostList(grouped);
  var data = { dateStr: dateStr, order: grouped.order, map: grouped.map, flat: flat };
  PropertiesService.getScriptProperties().setProperty("TODAY_POSTS", JSON.stringify(data));
}
function getCachedGrouped() {
  var raw = PropertiesService.getScriptProperties().getProperty("TODAY_POSTS");
  if (!raw) return null;
  try {
    var data = JSON.parse(raw);
    if (data.dateStr !== getTodayString()) return null; // stale — different day
    return { grouped: { order: data.order, map: data.map }, flat: data.flat, dateStr: data.dateStr };
  } catch(e) { return null; }
}
// ── PENDING LINES CACHE ───────────────────────────────────────
// Caches yesterday's outstanding post lines computed at morning send time.
// Button taps in doPost() use this so they never need to re-read sheets for
// yesterday's data — those lines don't change when today's buttons are tapped.
function getCachedPendingLines() {
  var raw = PropertiesService.getScriptProperties().getProperty("YESTERDAY_PENDING");
  if (!raw) return [];
  try {
    var d = JSON.parse(raw);
    // Cache is valid only for today's morning message
    return d.dateStr === getTodayString() ? (d.lines || []) : [];
  } catch(e) { return []; }
}
function setCachedPendingLines(lines) {
  PropertiesService.getScriptProperties().setProperty(
    "YESTERDAY_PENDING", JSON.stringify({ dateStr: getTodayString(), lines: lines })
  );
}
// ── STATUS CACHE ──────────────────────────────────────────────
// TODAY_STATUSES caches { dateStr, messageId, statuses, rowIndices } in Script Properties.
// Eliminates full tracker-sheet reads on every button tap in doPost().
function getCachedStatuses() {
  var raw = PropertiesService.getScriptProperties().getProperty("TODAY_STATUSES");
  if (!raw) return null;
  try {
    var d = JSON.parse(raw);
    return d.dateStr === getTodayString() ? d : null; // null if stale (different day)
  } catch(e) { return null; }
}
function setCachedStatuses(data) {
  PropertiesService.getScriptProperties()
    .setProperty("TODAY_STATUSES", JSON.stringify(data));
}
// Reads Bot_Tracker once to build yesterday's pending display lines.
// Cross-checks the Posts sheet so posts that were rescheduled or posted
// directly in the sheet don't appear as outstanding.
function buildYesterdayPendingLines(yStr) {
  var sh   = ensureTrackerSheet();
  var data = sh.getDataRange().getValues();

  // Read Posts sheet once for cross-checking rescheduled / directly-posted items
  var ss        = SpreadsheetApp.openById(SHEET_ID);
  var postSheet = ss.getSheetByName(SHEET_NAME);
  var postData  = postSheet ? postSheet.getDataRange().getValues() : [];
  var postHeaders = postData.length
    ? postData[0].map(function(h) { return String(h).trim().toLowerCase(); })
    : [];
  var COL = getColumnIndices(postHeaders);

  var lines = [];
  for (var i = 1; i < data.length; i++) {
    if (normaliseDateCell(data[i][0]) !== yStr) continue;
    if (String(data[i][3]) === "posted") continue;

    // Cross-check Posts sheet using the stored sheet row
    var sheetRow = parseInt(data[i][5]);
    if (sheetRow && postData.length >= sheetRow) {
      var postRow = postData[sheetRow - 1]; // sheetRow is 1-based
      // Skip if the post's date has been moved away from yesterday
      if (COL.date > -1 && normaliseDateCell(postRow[COL.date]) !== yStr) continue;
      // Skip if the post was marked posted directly in the sheet
      if (COL.posted > -1 && (postRow[COL.posted] === true || String(postRow[COL.posted]).toUpperCase() === "TRUE")) continue;
    }

    var pkParts = String(data[i][2]).split("|");
    var dispKey = pkParts.length > 1 ? pkParts.slice(1).join("|") : String(data[i][2]);
    lines.push("\u2022 " + dispKey + (data[i][4] ? " (" + data[i][4] + ")" : ""));
  }
  return lines;
}
// ── SEND MORNING DIGEST ───────────────────────────────────────
function sendMorningDigest() {
  try {
    var dateStr = getTodayString();
    var posts   = getPostsForDate(dateStr);
    if (!posts || posts.length === 0) {
      Logger.log("No posts today — morning skipped.");
      return;
    }
    var collapsed    = collapsePlatforms(posts);
    var grouped      = groupByOwner(collapsed);
    var flatPosts    = buildFlatPostList(grouped);
    // Compute yesterday's outstanding lines once here — cached for all button taps today
    var pendingLines = buildYesterdayPendingLines(getYesterdayString());
    setCachedPendingLines(pendingLines);
    var msg    = buildFullMessage(grouped, dateStr, {}, pendingLines);
    var result = sendEveningMessage(msg.text, msg.keyboard);
    var rowIndices = saveTrackerEntry(dateStr, result.result.message_id, flatPosts, {});
    cacheGrouped(dateStr, grouped); // cache so taps don't need to re-read Posts sheet
    // Cache today's statuses + rowIndices — eliminates tracker-sheet reads in doPost()
    var initStatuses = {};
    for (var k = 0; k < flatPosts.length; k++) initStatuses[postKey(flatPosts[k])] = "pending";
    setCachedStatuses({
      dateStr:    dateStr,
      messageId:  String(result.result.message_id),
      statuses:   initStatuses,
      rowIndices: rowIndices
    });
    Logger.log("Morning digest sent. Message ID: " + result.result.message_id);
  } catch(e) {
    Logger.log("Morning error: " + e.message);
    try { sendTelegramMessage("\u26a0\ufe0f Bot Error (morning)\n" + e.message); } catch(e2) {}
  }
}
// ── REFRESH TODAY DIGEST ─────────────────────────────────────
// Re-reads the Posts sheet and adds any new posts (added after the morning send)
// to the tracker and caches, then edits the existing Telegram message in place.
// Existing tick marks are preserved. Safe to call if nothing changed — no-op.
function refreshTodayDigest() {
  var dateStr    = getTodayString();
  var cachedData = getCachedStatuses();
  // No morning message was sent yet — nothing to refresh.
  if (!cachedData || !cachedData.messageId) {
    Logger.log("refreshTodayDigest: no morning message found, skipping.");
    return;
  }
  var statuses   = cachedData.statuses   || {};
  var rowIndices = cachedData.rowIndices || {};
  var messageId  = cachedData.messageId;

  // Re-read Posts sheet to find any posts added since morning.
  var freshPosts = getPostsForDate(dateStr);
  if (!freshPosts || freshPosts.length === 0) return;

  var sh       = ensureTrackerSheet();
  var newPosts = [];
  for (var i = 0; i < freshPosts.length; i++) {
    var key = postKey(freshPosts[i]);
    if (!statuses.hasOwnProperty(key)) {
      // New post — not in cache at morning send time.
      newPosts.push(freshPosts[i]);
      statuses[key] = "pending";
      // Append to Bot_Tracker so syncPostedColumn and pendingReminder pick it up.
      var newRow = sh.getLastRow() + 1;
      sh.appendRow([dateStr, String(messageId), key, "pending",
                    freshPosts[i].owner || "", freshPosts[i].sheetRow || ""]);
      rowIndices[key] = newRow;
    }
  }

  // Always update caches and rebuild — even if no new posts, this picks up
  // any changes to collapse/grouping logic and keeps the message in sync.
  cachedData.statuses   = statuses;
  cachedData.rowIndices = rowIndices;
  setCachedStatuses(cachedData);

  var collapsed    = collapsePlatforms(freshPosts);
  var grouped      = groupByOwner(collapsed);
  cacheGrouped(dateStr, grouped);

  var pendingLines = getCachedPendingLines();
  var rebuilt      = buildFullMessage(grouped, dateStr, statuses, pendingLines);
  editTelegramMessage(messageId, rebuilt.text, rebuilt.keyboard);
  Logger.log("refreshTodayDigest: " + (newPosts.length > 0 ? "added " + newPosts.length + " new post(s)" : "no new posts") + ", digest updated.");
}
// ── SEND EVENING / NIGHT REMINDER ────────────────────────────
// Sends a plain-text nudge for posts still pending. Skips silently if all ticked.
// At 6pm, also refreshes the digest to catch posts added after the morning send.
function sendPendingReminder(label) {
  try {
    var dateStr = getTodayString();
    if (label === "6pm") refreshTodayDigest();
    var msg = buildPendingReminderText(dateStr);
    if (msg) {
      sendTelegramMessage(msg);
      Logger.log(label + " reminder sent.");
    } else {
      Logger.log(label + " reminder skipped — all posts ticked.");
    }
  } catch(e) {
    Logger.log(label + " reminder error: " + e.message);
    try { sendTelegramMessage("\u26a0\ufe0f Bot Error (" + label + ")\n" + e.message); } catch(e2) {}
  }
}
function sendEveningReminder() { sendPendingReminder("6pm"); }
function sendNightReminder()   { sendPendingReminder("9pm"); }
// ── WEBHOOK: HANDLE BUTTON TAPS + TEXT COMMANDS ──────────────
function doPost(e) {
  var update = JSON.parse(e.postData.contents);

  // Handle text commands (e.g. /refresh sent in the group chat).
  if (update.message && update.message.text) {
    // Deduplicate — same guard as callback_query path.
    var msgProps    = PropertiesService.getScriptProperties();
    var msgLastSeen = parseInt(msgProps.getProperty("LAST_UPDATE_ID") || "0");
    if (update.update_id && update.update_id <= msgLastSeen) return ContentService.createTextOutput("OK");
    if (update.update_id) msgProps.setProperty("LAST_UPDATE_ID", String(update.update_id));

    var text     = String(update.message.text).trim();
    var botSuffix = "@" + getBotUsername();
    if (text === "/refresh" || text === "/refresh" + botSuffix) {
      try {
        refreshTodayDigest();
        sendTelegramMessage("\u2705 Digest refreshed — any new posts have been added.");
      } catch(err) {
        sendTelegramMessage("\u26a0\ufe0f Refresh failed: " + err.message);
      }
    } else if (text === "/syncposted" || text === "/syncposted" + botSuffix) {
      try {
        syncPostedColumn();
        sendTelegramMessage("\u2705 Posted column synced — sheet_row values repaired if needed.");
      } catch(err) {
        sendTelegramMessage("\u26a0\ufe0f Sync failed: " + err.message);
      }
    }
    return ContentService.createTextOutput("OK");
  }

  if (!update.callback_query) return ContentService.createTextOutput("OK");

  var cq        = update.callback_query;
  var messageId = cq.message.message_id;
  var parts     = cq.data.split("|");
  var action    = parts[0];
  // target may contain "|" (e.g. "5|tcacoustic") — rejoin everything after action
  var target    = parts.length > 1 ? parts.slice(1).join("|").trim() : "";
  // New buttons use a numeric flat-list index as target — resolve to postKey via cache.
  // Old postKey-format buttons (pre-fix) continue to work unchanged via direct lookup.
  if (/^\d+$/.test(target)) {
    var flatCache = getCachedGrouped();
    var fl = flatCache ? flatCache.flat : [];
    var fIdx = parseInt(target);
    if (fl[fIdx]) target = postKey(fl[fIdx]);
  }

  // Answer Telegram immediately — clears the spinner and gives the user instant feedback.
  // Meaningful toast text reduces urge to tap again while waiting for the message edit.
  var toastText = (action === "mark_posted")  ? "\u2705 Posted!" :
                  (action === "mark_pending") ? "\u23f3 Pending!" : "Updated!";
  answerCallbackQuery(cq.id, toastText);

  var lock = LockService.getScriptLock();
  try {
    // 6-second timeout keeps total response well within Telegram's retry window.
    // If another tap is already in progress, drop this one silently — the message
    // will be re-edited by the tap that holds the lock.
    if (!lock.tryLock(6000)) return ContentService.createTextOutput("OK");

    // Deduplicate — in case Telegram already retried before our answer landed.
    var props    = PropertiesService.getScriptProperties();
    var lastSeen = parseInt(props.getProperty("LAST_UPDATE_ID") || "0");
    if (update.update_id && update.update_id <= lastSeen) return ContentService.createTextOutput("OK");
    if (update.update_id) props.setProperty("LAST_UPDATE_ID", String(update.update_id));

    // Load statuses + rowIndices from PropertiesService cache (set at morning send).
    // Falls back to a full tracker-sheet read only on cold start (cache miss).
    var mid         = String(messageId);
    var cachedData  = getCachedStatuses();
    var statuses    = {};
    var rowIndices  = {};
    var dateStr     = getTodayString();
    var sh          = ensureTrackerSheet();
    if (cachedData) {
      statuses   = cachedData.statuses   || {};
      rowIndices = cachedData.rowIndices || {};
      dateStr    = cachedData.dateStr    || getTodayString();
    } else {
      // Cold start fallback: read tracker sheet once to rebuild cache
      var dbData = sh.getDataRange().getValues();
      for (var i = 1; i < dbData.length; i++) {
        if (String(dbData[i][1]) !== mid) continue;
        if (dateStr === getTodayString()) dateStr = normaliseDateCell(dbData[i][0]);
        var pk = String(dbData[i][2]).trim();
        statuses[pk]   = String(dbData[i][3]);
        rowIndices[pk] = i + 1;
      }
    }

    // Apply idempotent action — mark_posted/mark_pending cannot oscillate.
    // "toggle" is kept as a legacy fallback for old buttons still in chat history.
    if ((action === "mark_posted" || action === "mark_pending" || action === "toggle") && target !== "") {
      var newStatus;
      if (action === "toggle") {
        newStatus = statuses[target] === "posted" ? "pending" : "posted";
      } else {
        newStatus = (action === "mark_posted") ? "posted" : "pending";
      }
      if (statuses[target] !== newStatus) { // idempotent guard — skip if already at desired state
        statuses[target] = newStatus;
        if (rowIndices[target]) {
          sh.getRange(rowIndices[target], 4).setValue(newStatus);
        } else {
          sh.appendRow([dateStr, mid, target, newStatus, "", ""]);
        }
        // Write updated statuses back to cache
        if (cachedData) {
          cachedData.statuses = statuses;
          setCachedStatuses(cachedData);
        }
      }
    }

    // Use cached grouped structure + cached pending lines — zero sheet reads for button taps
    var cache        = getCachedGrouped();
    var grouped      = cache ? cache.grouped : groupByOwner(collapsePlatforms(getPostsForDate(dateStr)));
    var pendingLines = getCachedPendingLines();
    var rebuilt      = buildFullMessage(grouped, dateStr, statuses, pendingLines);
    editTelegramMessage(messageId, rebuilt.text, rebuilt.keyboard);
  } catch(err) {
    Logger.log("doPost error: " + err.message);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput("OK");
}
// ── DELETE ALL CALENDAR EVENTS ────────────────────────────────
// One-time cleanup: deletes all events tagged with EVENT_TAG across a 2-year window.
// Run once manually from the Apps Script editor after disabling the calendar sync.
function deleteAllCalendarEvents() {
  var cal = CalendarApp.getCalendarById(CONTENT_CAL_ID);
  if (!cal) throw new Error("Calendar not found: " + CONTENT_CAL_ID);
  var start = new Date();
  start.setFullYear(start.getFullYear() - 1);
  var end = new Date();
  end.setFullYear(end.getFullYear() + 1);
  var events  = cal.getEvents(start, end);
  var deleted = 0;
  for (var i = 0; i < events.length; i++) {
    var desc = events[i].getDescription() || "";
    if (desc.indexOf(EVENT_TAG) !== -1) {
      events[i].deleteEvent();
      deleted++;
      Utilities.sleep(100);
    }
  }
  Logger.log("Deleted " + deleted + " calendar event(s) tagged " + EVENT_TAG);
}
// ── SETUP ─────────────────────────────────────────────────────
function setupTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) ScriptApp.deleteTrigger(triggers[i]);
  ScriptApp.newTrigger("sendMorningDigest")
      .timeBased().atHour(MORNING_HOUR).everyDays(1).inTimezone(TIMEZONE).create();
  ScriptApp.newTrigger("sendEveningReminder")
      .timeBased().atHour(EVENING_HOUR).everyDays(1).inTimezone(TIMEZONE).create();
  ScriptApp.newTrigger("sendNightReminder")
      .timeBased().atHour(21).everyDays(1).inTimezone(TIMEZONE).create();
  ScriptApp.newTrigger("syncPostedColumn")
      .timeBased().atHour(23).everyDays(1).inTimezone(TIMEZONE).create();
  Logger.log("Triggers set: 9am digest, 6pm + 9pm reminders, 11pm sheet sync.");
}
function registerWebhook() {
  if (!WEBAPP_URL) {
    Logger.log("ERROR: Paste your Web App URL into WEBAPP_URL first.");
    return;
  }
  var url     = "https://api.telegram.org/bot" + BOT_TOKEN + "/setWebhook";
  var payload = { url: WEBAPP_URL };
  var options = {
    method:             "post",
    contentType:        "application/json",
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var result = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
  Logger.log(JSON.stringify(result));
}
// ── TEST ──────────────────────────────────────────────────────
function testMorning() {
  // Logs the morning message without sending. Change to sendMorningDigest() to send live.
  var posts = getPostsForDate(getTodayString());
  if (!posts || posts.length === 0) { Logger.log("(no posts today)"); return; }
  var collapsed = collapsePlatforms(posts);
  var grouped   = groupByOwner(collapsed);
  var flatPosts = buildFlatPostList(grouped);
  var dateStr   = getTodayString();
  var initialStatuses = {};
  for (var k = 0; k < flatPosts.length; k++) {
    if (flatPosts[k].isPosted) initialStatuses[String(k)] = "posted";
  }
  var msg = "\ud83d\udcc5 *Today's Posts \u2014 " + dateStr + "*\n\n" + buildEveningText(grouped, dateStr, initialStatuses);
  Logger.log(msg);
}
function testEditMessage() {
  var cachedData = getCachedStatuses();
  if (!cachedData) { sendTelegramMessage("testEdit: no cached status data"); return; }
  var url     = "https://api.telegram.org/bot" + BOT_TOKEN + "/editMessageText";
  var payload = {
    chat_id:    CHAT_ID,
    message_id: parseInt(cachedData.messageId),
    text:       "Test edit \u2014 " + new Date().toISOString(),
    parse_mode: "Markdown"
  };
  var result = JSON.parse(UrlFetchApp.fetch(url, {
    method: "post", contentType: "application/json",
    payload: JSON.stringify(payload), muteHttpExceptions: true
  }).getContentText());
  Logger.log("Edit test result: " + JSON.stringify(result));
  Logger.log("Cached message_id: " + cachedData.messageId);
}
function testRefreshEdit() {
  var dateStr    = getTodayString();
  var cachedData = getCachedStatuses();
  if (!cachedData) { Logger.log("No cache"); return; }
  var statuses     = cachedData.statuses || {};
  var cache        = getCachedGrouped();
  var grouped      = cache ? cache.grouped : groupByOwner(collapsePlatforms(getPostsForDate(dateStr)));
  var pendingLines = getCachedPendingLines();
  var rebuilt      = buildFullMessage(grouped, dateStr, statuses, pendingLines);
  Logger.log("Message text:\n" + rebuilt.text);
  var result = editTelegramMessage(cachedData.messageId, rebuilt.text, rebuilt.keyboard);
  Logger.log("Edit result: " + JSON.stringify(result));
}
function testEveningReminder() {
  sendEveningReminder();
}
function testNightReminder() {
  sendNightReminder();
}
function testCalendarSync() {
  syncCalendar();
  Logger.log("Calendar sync test complete.");
}
