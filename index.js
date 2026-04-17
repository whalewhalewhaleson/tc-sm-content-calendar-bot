'use strict';
require('dotenv').config();

const express    = require('express');
const axios      = require('axios');
const cron       = require('node-cron');
const { google } = require('googleapis');
const { DateTime } = require('luxon');

// ── CONFIG ────────────────────────────────────────────────────
const BOT_TOKEN     = process.env.BOT_TOKEN;
const SHEET_ID      = '1a2aASYC8qUd0knij8GmuAo-yZzRaqlTW4-j5Uhfam-g';
const CHAT_ID       = '-5256560725';
const SHEET_NAME    = 'Posts';
const TRACKER_SHEET = 'Bot_Tracker';
const TIMEZONE      = 'Asia/Singapore';
const PORT          = process.env.PORT || 3000;

// ── OWNER → EMAIL ─────────────────────────────────────────────
const OWNER_EMAILS = {
  mavia:    'mavia.kow@tcacoustic.com.sg',
  yov:      'yovanna.lo@tcacoustic.com.sg',
  yovanna:  'yovanna.lo@tcacoustic.com.sg',
  bibs:     'bibiane.chua@tcacoustic.com.sg',
  bibiane:  'bibiane.chua@tcacoustic.com.sg',
  haripria: 'haripria.gunalan@tcacoustic.com.my',
  hari:     'haripria.gunalan@tcacoustic.com.my',
  jc:       'jayciakhophet@tcacoustic.co.th',
  jaycia:   'jayciakhophet@tcacoustic.co.th',
  ryn:      'rynrynth04@gmail.com',
  min:      'soo.min@tcacoustic.com.sg'
};

const BRAND_TO_EMOJI = {
  'tc acoustic':      '⬜',
  'sonos':            '🔵',
  'bowers & wilkins': '⚫',
  'marshall':         '🔴'
};

const HANDLE_LABEL = {
  tcacoustic:             'TC',
  tc_acoustic:            'TC',
  friendsoftc:            'TC',
  sonosconceptstore:      'SCS',
  'sonos.hk':             'HK',
  sonoshongkong:          'HK',
  sonosthailandofficial:  'TH',
  sonosthailand:          'TH',
  bowerswilkinssingapore: 'B&W-SG',
  bowerswilkinsmalaysia:  'B&W-MY',
  thatsoundspot:          'TSS'
};

// ── IN-MEMORY CACHE (replaces PropertiesService) ──────────────
const _cache = {};
function setProp(key, value) {
  _cache[key] = (value !== null && value !== undefined && typeof value !== 'string')
    ? JSON.stringify(value)
    : value;
}
function getProp(key) {
  return _cache[key] !== undefined ? _cache[key] : null;
}

// ── ASYNC LOCK (replaces LockService) ─────────────────────────
let _lockChain = Promise.resolve();
function withLock(fn) {
  const next = _lockChain.then(() => fn()).catch(() => {});
  _lockChain = next.then(() => {}, () => {});
  return next;
}

// ── DEDUP ─────────────────────────────────────────────────────
let _lastUpdateId = 0;

// ── GOOGLE SHEETS ─────────────────────────────────────────────
let _sheets = null;
function getSheets() {
  if (_sheets) return _sheets;
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON || '{}');
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  _sheets = google.sheets({ version: 'v4', auth });
  return _sheets;
}

async function getSheetValues(sheetName) {
  const res = await getSheets().spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: sheetName,
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'SERIAL_NUMBER'
  });
  return res.data.values || [];
}

async function updateCell(sheetName, rowNum, colNum, value) {
  await getSheets().spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${sheetName}!${colToLetter(colNum)}${rowNum}`,
    valueInputOption: 'RAW',
    requestBody: { values: [[value]] }
  });
}

async function appendRows(sheetName, rows) {
  await getSheets().spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `${sheetName}!A1`,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: rows }
  });
}

async function batchUpdateCells(sheetName, updates) {
  if (!updates.length) return;
  await getSheets().spreadsheets.values.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      valueInputOption: 'RAW',
      data: updates.map(u => ({
        range: `${sheetName}!${colToLetter(u.col)}${u.row}`,
        values: [[u.value]]
      }))
    }
  });
}

function colToLetter(n) {
  let s = '';
  while (n > 0) {
    s = String.fromCharCode(64 + ((n - 1) % 26 + 1)) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// Google Sheets date serial → luxon DateTime in SGT
function serialToDateTime(serial) {
  return DateTime.fromObject({ year: 1899, month: 12, day: 30 }, { zone: TIMEZONE })
    .plus({ days: Math.floor(serial) });
}

// ── DATE / TIME HELPERS ───────────────────────────────────────
function getTodayString() {
  return DateTime.now().setZone(TIMEZONE).toFormat('dd MMM yyyy');
}

function getYesterdayString() {
  return DateTime.now().setZone(TIMEZONE).minus({ days: 1 }).toFormat('dd MMM yyyy');
}

function normaliseDateCell(raw) {
  if (raw === null || raw === undefined || raw === '') return '';
  if (typeof raw === 'number') return serialToDateTime(raw).toFormat('dd MMM yyyy');
  const str = String(raw).trim();
  if (!str) return '';
  let dt = DateTime.fromFormat(str, 'dd MMM yyyy', { zone: TIMEZONE });
  if (dt.isValid) return dt.toFormat('dd MMM yyyy');
  dt = DateTime.fromISO(str, { zone: TIMEZONE });
  if (dt.isValid) return dt.toFormat('dd MMM yyyy');
  return str;
}

function normaliseTimeCell(raw) {
  if (raw === null || raw === undefined || raw === '' || raw === 0) return '';
  if (typeof raw === 'number') {
    const totalMins = Math.round(raw * 24 * 60);
    const h = Math.floor(totalMins / 60) % 24;
    const m = totalMins % 60;
    if (h === 0 && m === 0) return '';
    return DateTime.fromObject({ hour: h, minute: m }).toFormat('h:mm a');
  }
  return String(raw).trim();
}

// ── TEXT HELPERS ──────────────────────────────────────────────
function escapeMd(text) {
  return String(text || '').replace(/([*_`\[])/g, '\\$1');
}

function getBrandEmoji(brand) {
  if (!brand) return '📌';
  return BRAND_TO_EMOJI[brand.toLowerCase().trim()] || '📌';
}

function getHandleLabel(handle) {
  if (!handle) return handle;
  return HANDLE_LABEL[handle.toLowerCase().trim()] || handle;
}

// ── PARSE POST ROW ────────────────────────────────────────────
function parsePost(row, COL) {
  const accountsRaw = String(row[COL.account] || '').trim();
  const title       = COL.title   > -1 ? String(row[COL.title]   || '').trim() : '';
  const pillar      = COL.pillar  > -1 ? String(row[COL.pillar]  || '').trim() : '';
  const brand       = COL.brand   > -1 ? String(row[COL.brand]   || '').trim() : '';
  const owner       = COL.owner   > -1 ? String(row[COL.owner]   || '').trim() : '';
  const creator     = COL.creator > -1 ? String(row[COL.creator] || '').trim() : '';
  const timeStr     = COL.time    > -1 ? normaliseTimeCell(row[COL.time]) : '';

  if (!accountsRaw || !title) return [];

  const entries     = accountsRaw.split(',');
  const handleMap   = {};
  const handleOrder = [];

  for (const entry of entries) {
    const m = entry.trim().match(/^([A-Za-z0-9]+)\/([A-Z]{2,3})\s+@(\S+)/i);
    if (!m) continue;
    let plat   = m[1].trim().toUpperCase();
    const mkt  = m[2].trim().toUpperCase();
    const hndl = m[3].trim().toLowerCase().replace(/[()]/g, '');
    if (plat === 'YT') plat = entry.toLowerCase().includes('short') ? 'YT Short' : 'YT Video';
    if (!handleMap[hndl]) {
      handleOrder.push(hndl);
      handleMap[hndl] = { label: getHandleLabel(hndl), platforms: [], market: mkt };
    }
    if (!handleMap[hndl].platforms.includes(plat)) handleMap[hndl].platforms.push(plat);
  }

  if (!handleOrder.length) return [];

  return handleOrder.map(hndl => ({
    handle:      hndl,
    label:       handleMap[hndl].label,
    platform:    handleMap[hndl].platforms.join('/'),
    postName:    title,
    postingTime: timeStr,
    creator,
    owner,
    pillar,
    brand,
    market:      handleMap[hndl].market,
    emoji:       getBrandEmoji(brand)
  }));
}

function getColumnIndices(headers) {
  const COL = { posted: -1, date: -1, time: -1, account: -1, title: -1, pillar: -1, brand: -1, owner: -1, creator: -1 };
  headers.forEach((h, i) => {
    if (h === 'posted')   COL.posted  = i;
    if (h === 'date')     COL.date    = i;
    if (h === 'time')     COL.time    = i;
    if (h === 'accounts') COL.account = i;
    if (h === 'title')    COL.title   = i;
    if (h === 'pillar')   COL.pillar  = i;
    if (h === 'brand')    COL.brand   = i;
    if (h === 'owner')    COL.owner   = i;
    if (h === 'creator')  COL.creator = i;
  });
  return COL;
}

// ── FETCH POSTS ───────────────────────────────────────────────
async function getPostsForDate(dateStr) {
  const data    = await getSheetValues(SHEET_NAME);
  if (!data.length) return [];
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const COL     = getColumnIndices(headers);
  const posts   = [];

  for (let i = 1; i < data.length; i++) {
    if (normaliseDateCell(data[i][COL.date]) !== dateStr) continue;
    const parsed   = parsePost(data[i], COL);
    const sheetRow = i + 1;
    const isPosted = COL.posted > -1 ? (data[i][COL.posted] === true) : false;
    for (const p of parsed) posts.push({ ...p, sheetRow, isPosted });
  }
  return posts;
}

// ── COLLAPSE PLATFORMS ────────────────────────────────────────
function collapsePlatforms(posts) {
  const map   = {};
  const order = [];
  for (const p of posts) {
    const key = `${p.postName}|${p.label}|${p.market}|${p.owner}`;
    if (!map[key]) {
      order.push(key);
      map[key] = {
        handle:      p.handle,
        label:       p.label,
        platforms:   p.platform,
        postName:    p.postName,
        postingTime: p.postingTime || '',
        creator:     p.creator || '',
        owner:       p.owner,
        pillar:      p.pillar,
        brand:       p.brand,
        market:      p.market,
        emoji:       p.emoji,
        sheetRow:    p.sheetRow || null,
        isPosted:    p.isPosted || false
      };
    } else {
      if (!map[key].platforms.includes(p.platform)) map[key].platforms += '/' + p.platform;
      if (!map[key].market && p.market) map[key].market = p.market;
      if (!map[key].postingTime && p.postingTime) map[key].postingTime = p.postingTime;
    }
  }
  return order.map(k => map[k]);
}

// ── GROUP BY OWNER ────────────────────────────────────────────
function groupByOwner(posts) {
  const map   = {};
  const order = [];
  for (const p of posts) {
    const owner = p.owner || '—';
    if (!map[owner]) { order.push(owner); map[owner] = []; }
    map[owner].push(p);
  }
  return { map, order };
}

// ── FLAT POST LIST + POST KEY ─────────────────────────────────
function postKey(p) {
  const titleSlug = (p.postName || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').slice(0, 40);
  return `${titleSlug}|${p.handle || ''}`;
}

function buildFlatPostList(grouped) {
  const flat = [];
  for (const owner of grouped.order) {
    for (const p of grouped.map[owner]) flat.push(p);
  }
  return flat;
}

// ── TRACKER SHEET ─────────────────────────────────────────────
async function saveTrackerEntry(dateStr, messageId, flatPosts, initialStatuses) {
  initialStatuses = initialStatuses || {};
  const existing  = await getSheetValues(TRACKER_SHEET);
  let nextRow     = existing.length + 1;
  const rows      = [];
  const rowIndices = {};

  for (const p of flatPosts) {
    const key    = postKey(p);
    const status = initialStatuses[key] || 'pending';
    rows.push([dateStr, String(messageId), key, status, p.owner || '', p.sheetRow || '']);
    rowIndices[key] = nextRow++;
  }
  if (rows.length) await appendRows(TRACKER_SHEET, rows);
  return rowIndices;
}

// ── BUILD POST LINE ───────────────────────────────────────────
function buildPostLine(p) {
  let line = '• ';
  if (p.postingTime) line += p.postingTime + ': ';
  line += `${p.label} | ${p.platforms} — ${escapeMd(p.postName)}`;
  if (p.creator) line += ` (${escapeMd(p.creator)})`;
  return line + '\n';
}

// ── BUILD EVENING TEXT ────────────────────────────────────────
function buildEveningText(grouped, statuses) {
  statuses = statuses || {};
  let msg = '';
  for (const owner of grouped.order) {
    msg += `👤 *${escapeMd(owner)}*\n`;
    for (const p of grouped.map[owner]) {
      const status = statuses[postKey(p)] || 'pending';
      msg += (status === 'posted' ? '✅ ' : '') + buildPostLine(p);
    }
    msg += '\n';
  }
  return msg.trim();
}

// ── BUILD FULL MESSAGE ────────────────────────────────────────
function buildFullMessage(grouped, dateStr, statuses, pendingLines) {
  statuses     = statuses     || {};
  pendingLines = pendingLines || [];

  const lines = [`📅 *Today's Posts — ${dateStr}*`];
  if (pendingLines.length > 0) {
    lines.push('', '📋 *Outstanding from yesterday*', ...pendingLines, '─────────────────');
  }
  lines.push('', buildEveningText(grouped, statuses));

  const flatPosts = buildFlatPostList(grouped);
  const buttons   = flatPosts.map((p, idx) => {
    const status     = statuses[postKey(p)] || 'pending';
    const nextAction = status === 'posted' ? 'mark_pending' : 'mark_posted';
    let tlbl = `${p.label} | ${p.platforms} — ${p.postName}`;
    if (status === 'posted') tlbl = '✅ ' + tlbl;
    return [{ text: tlbl, callback_data: `${nextAction}|${idx}` }];
  });

  return { text: lines.join('\n'), keyboard: { inline_keyboard: buttons } };
}

// ── PENDING REMINDER TEXT ─────────────────────────────────────
async function buildPendingReminderText(dateStr) {
  const data   = await getSheetValues(TRACKER_SHEET);
  let pending  = 0;
  let hasEntry = false;

  for (let i = 1; i < data.length; i++) {
    if (normaliseDateCell(data[i][0]) !== dateStr) continue;
    hasEntry = true;
    if (String(data[i][3]) !== 'posted') pending++;
  }
  if (!hasEntry || pending === 0) return null;
  return `🔔 *Reminder — ${dateStr}*\n${pending} post${pending === 1 ? '' : 's'} still not ticked. Please post & tick!`;
}

// ── CACHES ────────────────────────────────────────────────────
function cacheGrouped(dateStr, grouped) {
  setProp('TODAY_POSTS', { dateStr, order: grouped.order, map: grouped.map, flat: buildFlatPostList(grouped) });
}

function getCachedGrouped() {
  const raw = getProp('TODAY_POSTS');
  if (!raw) return null;
  try {
    const data = typeof raw === 'string' ? JSON.parse(raw) : raw;
    if (data.dateStr !== getTodayString()) return null;
    return { grouped: { order: data.order, map: data.map }, flat: data.flat };
  } catch (e) { return null; }
}

function getCachedPendingLines() {
  const raw = getProp('YESTERDAY_PENDING');
  if (!raw) return [];
  try {
    const d = typeof raw === 'string' ? JSON.parse(raw) : raw;
    return d.dateStr === getTodayString() ? (d.lines || []) : [];
  } catch (e) { return []; }
}

function setCachedPendingLines(lines) {
  setProp('YESTERDAY_PENDING', { dateStr: getTodayString(), lines });
}

function getCachedStatuses() {
  const raw = getProp('TODAY_STATUSES');
  if (!raw) return null;
  try {
    const d = typeof raw === 'string' ? JSON.parse(raw) : raw;
    return d.dateStr === getTodayString() ? d : null;
  } catch (e) { return null; }
}

function setCachedStatuses(data) {
  setProp('TODAY_STATUSES', data);
}

// ── YESTERDAY PENDING LINES ───────────────────────────────────
async function buildYesterdayPendingLines(yStr) {
  const [trackerData, postData] = await Promise.all([
    getSheetValues(TRACKER_SHEET),
    getSheetValues(SHEET_NAME)
  ]);
  const postHeaders = postData.length
    ? postData[0].map(h => String(h).trim().toLowerCase())
    : [];
  const COL   = getColumnIndices(postHeaders);
  const lines = [];

  for (let i = 1; i < trackerData.length; i++) {
    if (normaliseDateCell(trackerData[i][0]) !== yStr) continue;
    if (String(trackerData[i][3]) === 'posted') continue;

    const sheetRow = parseInt(trackerData[i][5]);
    if (sheetRow && postData.length >= sheetRow) {
      const postRow = postData[sheetRow - 1];
      if (COL.date > -1 && normaliseDateCell(postRow[COL.date]) !== yStr) continue;
      if (COL.posted > -1 && (postRow[COL.posted] === true || String(postRow[COL.posted]).toUpperCase() === 'TRUE')) continue;
    }

    const pkParts = String(trackerData[i][2]).split('|');
    const dispKey = pkParts.length > 1 ? pkParts.slice(1).join('|') : String(trackerData[i][2]);
    lines.push(`• ${escapeMd(dispKey)}${trackerData[i][4] ? ` (${escapeMd(trackerData[i][4])})` : ''}`);
  }
  return lines;
}

// ── TELEGRAM HELPERS ──────────────────────────────────────────
async function telegramRequest(method, payload) {
  const res = await axios.post(`https://api.telegram.org/bot${BOT_TOKEN}/${method}`, payload);
  return res.data;
}

async function sendTelegramMessage(text) {
  const res = await telegramRequest('sendMessage', { chat_id: CHAT_ID, text, parse_mode: 'Markdown' });
  if (!res.ok) throw new Error('Telegram error: ' + res.description);
  return res;
}

async function sendWithKeyboard(text, keyboard) {
  const res = await telegramRequest('sendMessage', {
    chat_id: CHAT_ID, text, parse_mode: 'Markdown', reply_markup: keyboard
  });
  if (!res.ok) throw new Error('Telegram send error: ' + res.description);
  return res;
}

async function editTelegramMessage(messageId, text, keyboard) {
  const res = await telegramRequest('editMessageText', {
    chat_id: CHAT_ID, message_id: parseInt(messageId),
    text, parse_mode: 'Markdown', reply_markup: keyboard
  });
  if (!res.ok) {
    if (res.description && res.description.includes('not modified')) return res;
    console.error('Edit error:', res.description);
  }
  return res;
}

async function answerCallbackQuery(callbackQueryId, text) {
  await telegramRequest('answerCallbackQuery', { callback_query_id: callbackQueryId, text });
}

let _botUsername = null;
async function getBotUsername() {
  if (_botUsername) return _botUsername;
  const res = await telegramRequest('getMe', {});
  _botUsername = res.ok ? res.result.username : '';
  return _botUsername;
}

// ── SEND MORNING DIGEST ───────────────────────────────────────
async function sendMorningDigest() {
  try {
    const dateStr = getTodayString();
    const posts   = await getPostsForDate(dateStr);
    if (!posts.length) { console.log('No posts today — morning skipped.'); return; }

    const collapsed    = collapsePlatforms(posts);
    const grouped      = groupByOwner(collapsed);
    const flatPosts    = buildFlatPostList(grouped);
    const pendingLines = await buildYesterdayPendingLines(getYesterdayString());
    setCachedPendingLines(pendingLines);

    const msg        = buildFullMessage(grouped, dateStr, {}, pendingLines);
    const result     = await sendWithKeyboard(msg.text, msg.keyboard);
    const rowIndices = await saveTrackerEntry(dateStr, result.result.message_id, flatPosts, {});
    cacheGrouped(dateStr, grouped);

    const initStatuses = {};
    for (const p of flatPosts) initStatuses[postKey(p)] = 'pending';
    setCachedStatuses({ dateStr, messageId: String(result.result.message_id), statuses: initStatuses, rowIndices });
    console.log('Morning digest sent. Message ID:', result.result.message_id);
  } catch (e) {
    console.error('Morning error:', e.message);
    try { await sendTelegramMessage('⚠️ Bot Error (morning)\n' + e.message); } catch (_) {}
  }
}

// ── REFRESH TODAY DIGEST ──────────────────────────────────────
async function refreshTodayDigest() {
  const dateStr    = getTodayString();
  const cachedData = getCachedStatuses();
  if (!cachedData || !cachedData.messageId) {
    console.log('refreshTodayDigest: no morning message found, skipping.');
    return;
  }
  const statuses   = cachedData.statuses   || {};
  const rowIndices = cachedData.rowIndices || {};
  const messageId  = cachedData.messageId;

  const freshPosts = await getPostsForDate(dateStr);
  if (!freshPosts.length) return;

  const existing    = await getSheetValues(TRACKER_SHEET);
  let nextRow       = existing.length + 1;
  const newRows     = [];
  const newPosts    = [];

  for (const p of freshPosts) {
    const key = postKey(p);
    if (!Object.prototype.hasOwnProperty.call(statuses, key)) {
      newPosts.push(p);
      statuses[key] = 'pending';
      newRows.push([dateStr, String(messageId), key, 'pending', p.owner || '', p.sheetRow || '']);
      rowIndices[key] = nextRow++;
    }
  }
  if (newRows.length) await appendRows(TRACKER_SHEET, newRows);

  cachedData.statuses   = statuses;
  cachedData.rowIndices = rowIndices;
  setCachedStatuses(cachedData);

  const collapsed    = collapsePlatforms(freshPosts);
  const grouped      = groupByOwner(collapsed);
  cacheGrouped(dateStr, grouped);

  const pendingLines = getCachedPendingLines();
  const rebuilt      = buildFullMessage(grouped, dateStr, statuses, pendingLines);
  await editTelegramMessage(messageId, rebuilt.text, rebuilt.keyboard);
  console.log(`refreshTodayDigest: ${newPosts.length ? `added ${newPosts.length} new post(s)` : 'no new posts'}, digest updated.`);
}

// ── PENDING REMINDER ──────────────────────────────────────────
async function sendPendingReminder(label) {
  try {
    const dateStr = getTodayString();
    if (label === '6pm') await refreshTodayDigest();
    const msg = await buildPendingReminderText(dateStr);
    if (msg) {
      await sendTelegramMessage(msg);
      console.log(`${label} reminder sent.`);
    } else {
      console.log(`${label} reminder skipped — all posts ticked.`);
    }
  } catch (e) {
    console.error(`${label} reminder error:`, e.message);
    try { await sendTelegramMessage(`⚠️ Bot Error (${label})\n${e.message}`); } catch (_) {}
  }
}

// ── SYNC POSTED COLUMN ────────────────────────────────────────
async function syncPostedColumn() {
  try {
    const today       = getTodayString();
    const [trackerData, postData] = await Promise.all([
      getSheetValues(TRACKER_SHEET),
      getSheetValues(SHEET_NAME)
    ]);
    const postHeaders = postData.length ? postData[0].map(h => String(h).trim().toLowerCase()) : [];
    const COL  = getColumnIndices(postHeaders);
    const pCol = postHeaders.indexOf('posted') + 1;
    if (pCol === 0) return;

    const rowLookup = {};
    if (COL.title > -1 && COL.account > -1) {
      for (let r = 1; r < postData.length; r++) {
        const title = String(postData[r][COL.title] || '').trim();
        if (!title) continue;
        const slug = title.toLowerCase().replace(/[^a-z0-9]+/g, '-').slice(0, 40);
        for (const ae of String(postData[r][COL.account] || '').split(',')) {
          const m = ae.trim().match(/@(\S+)/i);
          if (!m) continue;
          rowLookup[`${slug}|${m[1].toLowerCase().replace(/[()]/g, '')}`] = r + 1;
        }
      }
    }

    const postUpdates   = [];
    const trackerRepair = [];
    for (let i = 1; i < trackerData.length; i++) {
      if (normaliseDateCell(trackerData[i][0]) !== today) continue;
      const pkey      = String(trackerData[i][2]);
      const lookedUp  = rowLookup[pkey];
      const actualRow = lookedUp || parseInt(trackerData[i][5]);
      if (!actualRow) continue;
      if (lookedUp && lookedUp !== parseInt(trackerData[i][5])) {
        trackerRepair.push({ row: i + 1, col: 6, value: lookedUp });
      }
      postUpdates.push({ row: actualRow, col: pCol, value: String(trackerData[i][3]) === 'posted' });
    }
    await Promise.all([
      batchUpdateCells(SHEET_NAME, postUpdates),
      batchUpdateCells(TRACKER_SHEET, trackerRepair)
    ]);
    console.log('syncPostedColumn done for', today);
  } catch (e) { console.error('syncPostedColumn error:', e.message); }
}

// ── WEBHOOK HANDLER ───────────────────────────────────────────
async function handleUpdate(update) {
  if (update.message && update.message.text) {
    if (update.update_id && update.update_id <= _lastUpdateId) return;
    if (update.update_id) _lastUpdateId = update.update_id;

    const text = String(update.message.text).trim();
    try {
      const botName   = await getBotUsername();
      const isRefresh = text === '/refresh' || text === `/refresh@${botName}`;
      const isSync    = text === '/syncposted' || text === `/syncposted@${botName}`;
      if (isRefresh) {
        await refreshTodayDigest();
        await sendTelegramMessage('✅ Digest refreshed — any new posts have been added.');
      } else if (isSync) {
        await syncPostedColumn();
        await sendTelegramMessage('✅ Posted column synced — sheet_row values repaired if needed.');
      }
    } catch (err) {
      await sendTelegramMessage('⚠️ Command failed: ' + err.message);
    }
    return;
  }

  if (!update.callback_query) return;

  const cq        = update.callback_query;
  const messageId = cq.message.message_id;
  const parts     = cq.data.split('|');
  const action    = parts[0];
  let target      = parts.length > 1 ? parts.slice(1).join('|').trim() : '';

  if (/^\d+$/.test(target)) {
    const flatCache = getCachedGrouped();
    const fl        = flatCache ? flatCache.flat : [];
    const fIdx      = parseInt(target);
    if (fl[fIdx]) target = postKey(fl[fIdx]);
  }

  const toastText = action === 'mark_posted'  ? '✅ Posted!' :
                    action === 'mark_pending' ? '⏳ Pending!' : 'Updated!';
  await answerCallbackQuery(cq.id, toastText);

  await withLock(async () => {
    if (update.update_id && update.update_id <= _lastUpdateId) return;
    if (update.update_id) _lastUpdateId = update.update_id;

    const mid       = String(messageId);
    let cachedData  = getCachedStatuses();
    let statuses    = {};
    let rowIndices  = {};
    let dateStr     = getTodayString();

    if (cachedData) {
      statuses   = cachedData.statuses   || {};
      rowIndices = cachedData.rowIndices || {};
      dateStr    = cachedData.dateStr    || getTodayString();
    } else {
      const dbData = await getSheetValues(TRACKER_SHEET);
      for (let i = 1; i < dbData.length; i++) {
        if (String(dbData[i][1]) !== mid) continue;
        dateStr = normaliseDateCell(dbData[i][0]);
        const pk = String(dbData[i][2]).trim();
        statuses[pk]   = String(dbData[i][3]);
        rowIndices[pk] = i + 1;
      }
    }

    if ((action === 'mark_posted' || action === 'mark_pending' || action === 'toggle') && target !== '') {
      const newStatus = action === 'toggle'
        ? (statuses[target] === 'posted' ? 'pending' : 'posted')
        : (action === 'mark_posted' ? 'posted' : 'pending');

      if (statuses[target] !== newStatus) {
        statuses[target] = newStatus;
        if (cachedData) { cachedData.statuses = statuses; setCachedStatuses(cachedData); }
        // Write to Sheets in background — don't block the message edit
        if (rowIndices[target]) {
          updateCell(TRACKER_SHEET, rowIndices[target], 4, newStatus).catch(e => console.error('Tracker write error:', e));
        } else {
          appendRows(TRACKER_SHEET, [[dateStr, mid, target, newStatus, '', '']]).catch(e => console.error('Tracker append error:', e));
        }
      }
    }

    const cache        = getCachedGrouped();
    const grouped      = cache
      ? cache.grouped
      : groupByOwner(collapsePlatforms(await getPostsForDate(dateStr)));
    const pendingLines = getCachedPendingLines();
    const rebuilt      = buildFullMessage(grouped, dateStr, statuses, pendingLines);
    await editTelegramMessage(messageId, rebuilt.text, rebuilt.keyboard);
  });
}

// ── EXPRESS APP ───────────────────────────────────────────────
const app = express();
app.use(express.json());

app.post('/webhook', async (req, res) => {
  res.sendStatus(200);
  try { await handleUpdate(req.body); }
  catch (e) { console.error('Webhook error:', e.message); }
});

app.get('/health', (_req, res) => res.json({ status: 'ok', time: new Date().toISOString() }));

// ── REGISTER WEBHOOK ──────────────────────────────────────────
async function registerWebhook() {
  const webhookUrl = process.env.WEBHOOK_URL;
  if (!webhookUrl) { console.warn('WEBHOOK_URL not set — skipping webhook registration.'); return; }
  const res = await telegramRequest('setWebhook', { url: `${webhookUrl}/webhook` });
  console.log('Webhook registered:', JSON.stringify(res));
}

// ── CRON JOBS ─────────────────────────────────────────────────
cron.schedule('0 9 * * *',  () => sendMorningDigest(),        { timezone: TIMEZONE });
cron.schedule('0 18 * * *', () => sendPendingReminder('6pm'), { timezone: TIMEZONE });
cron.schedule('0 21 * * *', () => sendPendingReminder('9pm'), { timezone: TIMEZONE });
cron.schedule('0 23 * * *', () => syncPostedColumn(),         { timezone: TIMEZONE });

// ── START ─────────────────────────────────────────────────────
app.listen(PORT, async () => {
  console.log(`Bot running on port ${PORT}`);
  await registerWebhook();
});
