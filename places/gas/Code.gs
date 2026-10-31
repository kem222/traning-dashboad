/**
 * ふたりの行きたいところ — バックエンド (Google Apps Script)
 *
 * Googleスプレッドシートを置き場所にして、
 *   - Webページ (places/index.html) からの 読み込み / 追加 / 編集 / 削除
 *   - LINE Bot (Messaging API) の Webhook からの 追加 / 一覧 / 行った登録
 * をまとめて受け付ける。
 *
 * スクリプトプロパティ（プロジェクトの設定 → スクリプト プロパティ）:
 *   APP_TOKEN          Webページ用の合言葉（自分で決めたランダムな文字列）
 *   LINE_WEBHOOK_KEY   LINEのWebhook URL に付ける合言葉（?src=line&key=...）
 *   LINE_ACCESS_TOKEN  LINE チャネルアクセストークン（長期）
 *   APP_URL            Webページの URL（LINEの返信に載せる。任意）
 *
 * セットアップ手順は places/README.md を参照。
 */

const SHEET_NAME = 'places';
const HEADERS = [
  'id', 'title', 'size', 'area', 'url', 'lat', 'lng', 'memo',
  'addedBy', 'source', 'createdAt', 'status', 'visitedAt', 'visitedMemo',
];
const SIZES = {
  trip: { label: '旅行', emoji: '🧳' },
  outing: { label: 'おでかけ', emoji: '🚗' },
  date: { label: 'デート', emoji: '☕' },
};

/* ───────────────────────── HTTP entry points ───────────────────────── */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!checkToken_(p.token)) return json_({ ok: false, error: 'unauthorized' });
  if (p.action === 'geocode') return json_({ ok: true, result: geocode_(p.q || '') });
  return json_({ ok: true, items: readAll_() });
}

function doPost(e) {
  const p = (e && e.parameter) || {};

  // LINE Webhook
  if (p.src === 'line') {
    if (!p.key || p.key !== prop_('LINE_WEBHOOK_KEY')) return text_('NG');
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    (body.events || []).forEach(function (ev) {
      try { handleLineEvent_(ev); } catch (err) { console.error(err && err.stack || err); }
    });
    return text_('OK');
  }

  // Webページからの API
  let req = {};
  try { req = JSON.parse((e.postData && e.postData.contents) || '{}'); } catch (_) {}
  if (!checkToken_(req.token)) return json_({ ok: false, error: 'unauthorized' });

  try {
    switch (req.action) {
      case 'add': return json_({ ok: true, item: addItem_(req.item || {}) });
      case 'update': return json_({ ok: true, item: updateItem_(req.id, req.patch || {}) });
      case 'delete': deleteItem_(req.id); return json_({ ok: true });
      case 'geocode': return json_({ ok: true, result: geocode_(req.q || '') });
      default: return json_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ───────────────────────── Sheet storage ───────────────────────── */

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // 日付やIDが勝手に変換されないように全部テキスト扱い（lat/lng だけ数値）
    sh.getRange('A:E').setNumberFormat('@');
    sh.getRange('H:N').setNumberFormat('@');
  }
  return sh;
}

function readAll_() {
  const values = sheet_().getDataRange().getValues();
  const head = values.shift() || [];
  return values
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      const o = {};
      head.forEach(function (k, i) {
        const v = r[i];
        o[k] = v instanceof Date ? v.toISOString() : v;
      });
      o.lat = o.lat === '' ? null : Number(o.lat);
      o.lng = o.lng === '' ? null : Number(o.lng);
      return o;
    });
}

function rowOf_(item) {
  return HEADERS.map(function (k) {
    const v = item[k];
    return v === undefined || v === null ? '' : v;
  });
}

function findRow_(id) {
  const ids = sheet_().getRange(1, 1, sheet_().getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  return -1;
}

function addItem_(input) {
  const item = normalize_(input);
  enrich_(item);
  if (!item.title) item.title = item.url || '（名前なし）';
  if (!item.size) item.size = guessSize_(item.title + ' ' + (item.memo || ''));
  return withLock_(function () {
    sheet_().appendRow(rowOf_(item));
    return item;
  });
}

function updateItem_(id, patch) {
  return withLock_(function () {
    const row = findRow_(id);
    if (row < 0) throw new Error('not found');
    const sh = sheet_();
    const cur = {};
    const vals = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
    HEADERS.forEach(function (k, i) { cur[k] = vals[i]; });
    HEADERS.forEach(function (k) {
      if (k !== 'id' && Object.prototype.hasOwnProperty.call(patch, k)) cur[k] = patch[k];
    });
    sh.getRange(row, 1, 1, HEADERS.length).setValues([rowOf_(cur)]);
    return cur;
  });
}

function deleteItem_(id) {
  withLock_(function () {
    const row = findRow_(id);
    if (row > 0) sheet_().deleteRow(row);
  });
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function normalize_(input) {
  return {
    id: Utilities.getUuid().slice(0, 8),
    title: String(input.title || '').trim().slice(0, 120),
    size: SIZES[input.size] ? input.size : '',
    area: String(input.area || '').trim(),
    url: String(input.url || '').trim(),
    lat: isFinite(parseFloat(input.lat)) ? parseFloat(input.lat) : null,
    lng: isFinite(parseFloat(input.lng)) ? parseFloat(input.lng) : null,
    memo: String(input.memo || '').trim(),
    addedBy: String(input.addedBy || '').trim(),
    source: input.source || 'web',
    createdAt: new Date().toISOString(),
    status: 'want',
    visitedAt: '',
    visitedMemo: '',
  };
}

/* ───────────────────────── Enrichment ───────────────────────── */

/** URL からタイトル・座標を補完し、それでも座標がなければ名前でジオコーディング */
function enrich_(item) {
  try {
    if (item.url && isMapsUrl_(item.url)) {
      const m = parseGoogleMaps_(item.url);
      if (m) {
        if (!item.title && m.name) item.title = m.name;
        if (item.lat == null && m.lat != null) { item.lat = m.lat; item.lng = m.lng; }
      }
    } else if (item.url && !item.title) {
      item.title = fetchTitle_(item.url);
    }
  } catch (err) { console.warn('url enrich failed', err); }

  try {
    if (item.lat == null && item.title) {
      const g = geocode_(item.title + (item.area ? ' ' + item.area : ''));
      if (g) { item.lat = g.lat; item.lng = g.lng; if (!item.area) item.area = g.area; }
    } else if (item.lat != null && !item.area) {
      item.area = reverseArea_(item.lat, item.lng);
    }
  } catch (err) { console.warn('geocode failed', err); }
}

function isMapsUrl_(url) {
  return /^https?:\/\/(maps\.app\.goo\.gl|goo\.gl\/maps|maps\.google\.|(www\.)?google\.[a-z.]+\/maps)/i.test(url);
}

function parseGoogleMaps_(url) {
  let u = url;
  // 短縮URLはリダイレクト先を辿る
  for (let i = 0; i < 4 && /goo\.gl/.test(u); i++) {
    const res = UrlFetchApp.fetch(u, { followRedirects: false, muteHttpExceptions: true });
    const h = res.getHeaders();
    const loc = h.Location || h.location;
    if (!loc) break;
    u = loc;
  }
  const out = { name: '', lat: null, lng: null };
  const place = u.match(/\/place\/([^/@?]+)/);
  if (place) out.name = decodeURIComponent(place[1].replace(/\+/g, ' '));
  const q = u.match(/[?&]q=([^&]+)/);
  if (!out.name && q && !/^-?\d+\.\d+,-?\d+\.\d+$/.test(decodeURIComponent(q[1]))) {
    out.name = decodeURIComponent(q[1].replace(/\+/g, ' '));
  }
  const pin = u.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/);
  const at = u.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  const qll = u.match(/[?&](?:q|ll|query)=(-?\d+\.\d+)(?:,|%2C)(-?\d+\.\d+)/);
  const c = pin || qll || at;
  if (c) { out.lat = parseFloat(c[1]); out.lng = parseFloat(c[2]); }
  return out;
}

function fetchTitle_(url) {
  const res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; PlacesBot/1.0)', 'Accept-Language': 'ja' },
  });
  const html = res.getContentText().slice(0, 200000);
  const og = html.match(/<meta[^>]+property=["']og:title["'][^>]+content=["']([^"']+)["']/i)
    || html.match(/<meta[^>]+content=["']([^"']+)["'][^>]+property=["']og:title["']/i);
  const t = og ? og[1] : ((html.match(/<title[^>]*>([^<]+)<\/title>/i) || [])[1] || '');
  return decodeEntities_(t).replace(/\s+/g, ' ').trim().slice(0, 80);
}

function decodeEntities_(s) {
  return String(s)
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#0?39;/g, "'")
    .replace(/&#(\d+);/g, function (_, n) { return String.fromCharCode(Number(n)); });
}

function geocode_(q) {
  q = String(q || '').trim();
  if (!q) return null;
  const r = Maps.newGeocoder().setLanguage('ja').setRegion('jp').geocode(q);
  if (!r || r.status !== 'OK' || !r.results.length) return null;
  const top = r.results[0];
  return { lat: top.geometry.location.lat, lng: top.geometry.location.lng, area: areaOf_(top.address_components) };
}

function reverseArea_(lat, lng) {
  const r = Maps.newGeocoder().setLanguage('ja').reverseGeocode(lat, lng);
  return r && r.status === 'OK' && r.results.length ? areaOf_(r.results[0].address_components) : '';
}

function areaOf_(comps) {
  let pref = '', city = '';
  (comps || []).forEach(function (c) {
    if (c.types.indexOf('administrative_area_level_1') >= 0) pref = c.long_name;
    if (!city && (c.types.indexOf('locality') >= 0)) city = c.long_name;
  });
  return (pref + city) || ((comps || []).filter(function (c) { return c.types.indexOf('country') >= 0; })[0] || {}).long_name || '';
}

function areaFromAddress_(addr) {
  const m = String(addr || '').replace(/^日本、?\s*(〒\d{3}-?\d{4})?\s*/, '').match(/^(.+?[都道府県])(.+?[市区町村])?/);
  return m ? m[1] + (m[2] || '') : '';
}

/** 大きさの自動判定（Web側 guessSize と同じルール） */
function guessSize_(text) {
  const t = String(text || '');
  if (/旅行|泊|ホテル|旅館|民宿|お宿|の宿|宿$|温泉|リゾート|グランピング|コテージ|ヴィラ|海外|離島/.test(t)) return 'trip';
  if (/カフェ|喫茶|cafe|coffee|コーヒー|珈琲|ランチ|ディナー|モーニング|レストラン|居酒屋|バー|ご飯|ごはん|パン|ベーカリー|スイーツ|ケーキ|パフェ|寿司|焼肉|ラーメン|ビストロ|食堂|酒場|鮨|そば|うどん|カレー|ピザ/i.test(t)) return 'date';
  if (/水族館|動物園|遊園地|テーマパーク|美術館|博物館|公園|ドライブ|展望|花火|祭|紅葉|花見|ライブ|フェス|植物園|牧場|アウトレット|海水浴|ビーチ|登山|ハイキング|滝|展示|イルミネーション/.test(t)) return 'outing';
  return 'date';
}

/* ───────────────────────── LINE ───────────────────────── */

const ADD_TRIGGER = /^(?:#?行きたい|#?いきたい|[+＋]|📌)[!！\s:：、]*/;
const SIZE_PREFIX = /^(旅行|泊まり|おでかけ|お出かけ|日帰り|デート|カフェ|ごはん|ご飯)[\s:：、]+/;
const SIZE_WORD = {
  '旅行': 'trip', '泊まり': 'trip',
  'おでかけ': 'outing', 'お出かけ': 'outing', '日帰り': 'outing',
  'デート': 'date', 'カフェ': 'date', 'ごはん': 'date', 'ご飯': 'date',
};

function handleLineEvent_(ev) {
  if (ev.type === 'join' || ev.type === 'follow') {
    reply_(ev.replyToken, helpText_());
    return;
  }
  if (ev.type !== 'message') return;

  const isGroup = ev.source.type !== 'user';
  const msg = ev.message;

  if (msg.type === 'location') {
    const item = addItem_({
      title: msg.title || msg.address || '地図のピン',
      area: areaFromAddress_(msg.address),
      lat: msg.latitude, lng: msg.longitude,
      memo: msg.title ? (msg.address || '') : '',
      addedBy: displayName_(ev.source), source: 'line',
    });
    reply_(ev.replyToken, addedText_(item));
    return;
  }
  if (msg.type !== 'text') return;

  let text = String(msg.text || '').trim();

  if (/^(リスト|一覧|list)$/i.test(text)) { reply_(ev.replyToken, listText_()); return; }
  if (/^(ヘルプ|help|使い方)$/i.test(text)) { reply_(ev.replyToken, helpText_()); return; }

  const done = text.match(/^(?:行った|いった|行ってきた)[!！\s:：、]+(.+)$/);
  if (done) { reply_(ev.replyToken, markDoneText_(done[1].trim())); return; }

  const hasMapsUrl = /https?:\/\/\S*(maps\.app\.goo\.gl|goo\.gl\/maps|google\.[a-z.]+\/maps)/i.test(text);
  const triggered = ADD_TRIGGER.test(text);
  // グループでは普段の会話を拾わないよう、合図があるときだけ追加
  if (isGroup && !triggered && !hasMapsUrl) return;

  text = text.replace(ADD_TRIGGER, '');
  const parsed = parseText_(text);
  if (!parsed.title && !parsed.url) {
    reply_(ev.replyToken, '何を追加する？\n例）行きたい カフェ 〇〇珈琲');
    return;
  }
  parsed.addedBy = displayName_(ev.source);
  parsed.source = 'line';
  reply_(ev.replyToken, addedText_(addItem_(parsed)));
}

function parseText_(text) {
  const out = { title: '', url: '', memo: '', size: '' };
  let t = String(text || '').trim();
  const sp = t.match(SIZE_PREFIX);
  if (sp) { out.size = SIZE_WORD[sp[1]]; t = t.slice(sp[0].length); }
  const urlM = t.match(/https?:\/\/[^\s　]+/);
  if (urlM) { out.url = urlM[0]; t = t.replace(urlM[0], ' '); }
  const lines = t.split(/\n/).map(function (s) { return s.trim(); }).filter(Boolean);
  out.title = (lines.shift() || '').slice(0, 120);
  out.memo = lines.join('\n');
  return out;
}

function addedText_(item) {
  const s = SIZES[item.size] || SIZES.date;
  const lines = [
    '📌 追加したよ',
    s.emoji + ' ' + item.title + (item.area ? '（' + item.area + '）' : ''),
    item.lat != null ? '🗺 地図にのせました' : '🗺 場所が分からなかったので、あとでWebで指定してね',
  ];
  const app = prop_('APP_URL');
  if (app) lines.push('', app);
  return lines.join('\n');
}

function listText_() {
  const items = readAll_().filter(function (i) { return i.status !== 'done'; });
  if (!items.length) return 'まだ何もないよ。\n「行きたい 〇〇」で追加できます';
  const order = ['trip', 'outing', 'date'];
  const lines = ['行きたいところ（' + items.length + '件）'];
  order.forEach(function (sz) {
    const list = items.filter(function (i) { return (i.size || 'date') === sz; })
      .sort(function (a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); });
    if (!list.length) return;
    lines.push('', SIZES[sz].emoji + ' ' + SIZES[sz].label);
    list.slice(0, 5).forEach(function (i) { lines.push('・' + i.title + (i.area ? '（' + i.area + '）' : '')); });
    if (list.length > 5) lines.push('　ほか' + (list.length - 5) + '件');
  });
  const app = prop_('APP_URL');
  if (app) lines.push('', '全部見る → ' + app);
  return lines.join('\n');
}

function markDoneText_(q) {
  const items = readAll_().filter(function (i) { return i.status !== 'done'; });
  const hit = items.filter(function (i) { return String(i.title).indexOf(q) >= 0; })[0]
    || items.filter(function (i) { return q.indexOf(String(i.title)) >= 0; })[0];
  if (!hit) return '「' + q + '」は見つからなかった…';
  updateItem_(hit.id, { status: 'done', visitedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd') });
  return '🎉 「' + hit.title + '」を行ったリストに移しました';
}

function helpText_() {
  return [
    '📍 ふたりの行きたいところ',
    '',
    '【追加】',
    '行きたい 〇〇カフェ',
    '行きたい 旅行 箱根の〇〇旅館',
    '＋ おでかけ 〇〇水族館',
    '・GoogleマップのリンクやLINEの位置情報を送ってもOK',
    '・2行目以降はメモになります',
    '・旅行 / おでかけ / カフェ を先頭につけると大きさを指定できます',
    '',
    '【その他】',
    'リスト … 一覧を見る',
    '行った 〇〇 … 行ったリストへ',
  ].join('\n') + (prop_('APP_URL') ? '\n\n' + prop_('APP_URL') : '');
}

function displayName_(source) {
  if (!source || !source.userId) return '';
  const cache = CacheService.getScriptCache();
  const key = 'name_' + source.userId;
  const hit = cache.get(key);
  if (hit) return hit;
  let path = '/v2/bot/profile/' + source.userId;
  if (source.type === 'group') path = '/v2/bot/group/' + source.groupId + '/member/' + source.userId;
  if (source.type === 'room') path = '/v2/bot/room/' + source.roomId + '/member/' + source.userId;
  try {
    const res = UrlFetchApp.fetch('https://api.line.me' + path, {
      headers: { Authorization: 'Bearer ' + prop_('LINE_ACCESS_TOKEN') },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) return '';
    const name = JSON.parse(res.getContentText()).displayName || '';
    cache.put(key, name, 21600);
    return name;
  } catch (_) { return ''; }
}

function reply_(replyToken, text) {
  if (!replyToken) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + prop_('LINE_ACCESS_TOKEN') },
    payload: JSON.stringify({ replyToken: replyToken, messages: [{ type: 'text', text: String(text).slice(0, 4900) }] }),
    muteHttpExceptions: true,
  });
}

/* ───────────────────────── utils ───────────────────────── */

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }

function checkToken_(t) {
  const expected = prop_('APP_TOKEN');
  return !!expected && t === expected;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function text_(s) { return ContentService.createTextOutput(s); }

/** エディタから一度だけ実行: シート作成と権限の承認 */
function setup() {
  sheet_();
  geocode_('東京駅');
  UrlFetchApp.fetch('https://example.com', { muteHttpExceptions: true });
  console.log('OK: シート作成と権限確認が完了しました');
}
