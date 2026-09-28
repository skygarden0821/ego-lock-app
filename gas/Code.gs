/**
 * EGO LOCK — Sheet Bridge (Google Apps Script Web App)
 * ------------------------------------------------------------------
 * 非公開スプレッドシートの必要な数値だけをJSONで返す中継エンドポイント。
 * Yoshiの実シート構成に合わせて設定済み。基本そのまま使える。
 *
 *  取得する数値:
 *   - ライン追加 / AD / 契約  … 「日次進捗」タブ A〜D列（日付/LINE追加数/AD数/契約数）
 *   - お月謝（月商）          … 各月タブ 202601〜（回収列＝支払完了の合計）
 *
 *  カレンダー連携（doPost / KEY必須）:
 *   - read  … 空き時間計算用の予定と、EGO LOCKタスクの予定の現状を返す
 *   - ops   … タスクの予定を作成・更新・削除（EGO LOCKが作った予定だけ）
 *   - probe … 書き込み先カレンダーに書けるかを確認
 *
 *  デプロイ手順は gas/README.md 参照（種類:ウェブアプリ / 実行:自分 / アクセス:全員）。
 */

var CONFIG = {
  KEY: '', // 合言葉はスクリプトプロパティ EGOLOCK_KEY に設定する（_key() 参照）。ここは空のままでよい。
  DAYS_BACK: 400, // 日次を何日分返すか

  // ① 日次: ライン追加・AD・契約（日次進捗タブ）
  DAILY: {
    sheetId: '112RxU7evib3RqCxt-IJDClvIZ9JeQ93RSm4IeqLbRt4',
    tab: '日次進捗',       // タブ名。空なら先頭シート
    dateHeader: '日付',    // 日付列の見出し（部分一致）
    cols: { lineAdds: 'LINE追加', ad: 'AD', contracts: '契約' } // 見出し部分一致
  },

  // ② お月謝（月商）: 202601〜 の月次タブ、回収列の合計
  OTSUKI: {
    sheetId: '1PVbJmO3oG9-M2fCgem4ugO79BpWHSKXlpvz082Cy-kM',
    tabPattern: '^20\\d{4}$',   // 202601, 202602 … のタブだけ対象
    revenueHeader: '回収',       // 回収列（優先）
    feeHeader: '費用',           // フォールバック用
    statusHeader: '状況',        // フォールバック用
    paidStatuses: ['支払完了']   // 回収列が空の月はこの状況の費用を合算
  },

  // ③ カレンダー連携（EGO LOCK のタスク ⇄ Googleカレンダー）
  CAL: {
    writeId: 'xector1.kunoike@gmail.com',   // タスクを書き込むカレンダー（X1_九之池）
    readIds: ['xector1.kunoike@gmail.com'], // 空き時間の計算に読む予定（複数可）
    tagKey: 'egolock',                      // EGO LOCKが作った予定の目印。これが無い予定には一切触らない
    maxRangeDays: 40
  }
};

/** 合言葉。スクリプトプロパティ EGOLOCK_KEY を優先し、無ければ CONFIG.KEY にフォールバック。 */
function _key() {
  try {
    var v = PropertiesService.getScriptProperties().getProperty('EGOLOCK_KEY');
    if (v) return String(v);
  } catch (x) {}
  return CONFIG.KEY || '';
}

function doGet(e) {
  try {
    var key = _key();
    if (!key) return _json({ ok: false, error: '合言葉が未設定です（スクリプトプロパティ EGOLOCK_KEY を設定してください）' });
    var got = e && e.parameter ? e.parameter.key : '';
    if (got !== key) return _json({ ok: false, error: 'bad key' });
    var tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - CONFIG.DAYS_BACK);

    var days = {};        // 'yyyy-MM-dd' -> {lineAdds, ad, contracts}
    var monthsRevenue = {}; // 'yyyy-MM' -> revenue

    _readDaily(CONFIG.DAILY, days, tz, cutoff);
    _readOtsuki(CONFIG.OTSUKI, monthsRevenue);

    // 当月サマリ
    var now = new Date();
    var ym = Utilities.formatDate(now, tz, 'yyyy-MM');
    var month = { lineAdds: 0, ad: 0, contracts: 0, revenue: monthsRevenue[ym] || 0 };
    Object.keys(days).forEach(function (dk) {
      if (dk.indexOf(ym) === 0) {
        var v = days[dk];
        month.lineAdds += v.lineAdds || 0;
        month.ad += v.ad || 0;
        month.contracts += v.contracts || 0;
      }
    });

    return _json({
      ok: true,
      updated: Utilities.formatDate(now, tz, "yyyy-MM-dd'T'HH:mm:ssXXX"),
      tz: tz,
      count: Object.keys(days).length,
      month: month,
      days: days,
      monthsRevenue: monthsRevenue
    });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

function _readDaily(cfg, days, tz, cutoff) {
  var ss = SpreadsheetApp.openById(cfg.sheetId);
  var sh = cfg.tab ? ss.getSheetByName(cfg.tab) : ss.getSheets()[0];
  if (!sh) return;
  var vals = sh.getDataRange().getValues();
  var hr = -1, di = -1, li = -1, ai = -1, ci = -1;
  for (var r = 0; r < Math.min(vals.length, 15); r++) {
    var row = vals[r].map(function (x) { return String(x).trim(); });
    var d = _findCol(row, cfg.dateHeader), l = _findCol(row, cfg.cols.lineAdds);
    if (d >= 0 && l >= 0) {
      hr = r; di = d; li = l;
      ai = _findCol(row, cfg.cols.ad);
      ci = _findCol(row, cfg.cols.contracts);
      break;
    }
  }
  if (hr < 0) return;
  for (var r = hr + 1; r < vals.length; r++) {
    var dk = _dateKey(vals[r][di], tz);
    if (!dk) continue;
    if (new Date(dk) < cutoff) continue;
    if (!days[dk]) days[dk] = {};
    if (li >= 0) days[dk].lineAdds = (days[dk].lineAdds || 0) + _num(vals[r][li]);
    if (ai >= 0) days[dk].ad = (days[dk].ad || 0) + _num(vals[r][ai]);
    if (ci >= 0) days[dk].contracts = (days[dk].contracts || 0) + _num(vals[r][ci]);
  }
}

function _readOtsuki(cfg, monthsRevenue) {
  var ss = SpreadsheetApp.openById(cfg.sheetId);
  var re = new RegExp(cfg.tabPattern);
  ss.getSheets().forEach(function (sh) {
    var name = String(sh.getName()).trim();
    if (!re.test(name)) return;
    var ym = name.slice(0, 4) + '-' + name.slice(4, 6);
    var vals = sh.getDataRange().getValues();
    if (!vals.length) return;
    // 見出し行を探す（回収 or 費用 を含む最初の行、通常0行目）
    var hr = 0;
    for (var r = 0; r < Math.min(vals.length, 6); r++) {
      var row = vals[r].map(function (x) { return String(x).trim(); });
      if (_findCol(row, cfg.revenueHeader) >= 0 || _findCol(row, cfg.feeHeader) >= 0) { hr = r; break; }
    }
    var header = vals[hr].map(function (x) { return String(x).trim(); });
    var jr = _findCol(header, cfg.revenueHeader);
    var fe = _findCol(header, cfg.feeHeader);
    var st = _findCol(header, cfg.statusHeader);
    var rev = 0;
    if (jr >= 0) {
      for (var r = hr + 1; r < vals.length; r++) rev += _num(vals[r][jr]);
    }
    if (rev === 0 && fe >= 0 && st >= 0) {
      for (var r = hr + 1; r < vals.length; r++) {
        var s = String(vals[r][st]).trim();
        if (cfg.paidStatuses.indexOf(s) >= 0) rev += _num(vals[r][fe]);
      }
    }
    monthsRevenue[ym] = Math.round(rev);
  });
}

/* ================================================================
 *  カレンダー連携（doPost）
 *  アプリから text/plain で JSON を POST する（CORSのプリフライト回避）。
 *  body = { key, action: 'read' | 'ops' | 'probe', ... }
 * ================================================================ */
function doPost(e) {
  try {
    var body = {};
    try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
    catch (x) { return _json({ ok: false, error: 'bad json' }); }
    var key = _key();
    if (!key) return _json({ ok: false, error: '合言葉が未設定です（スクリプトプロパティ EGOLOCK_KEY を設定してください。カレンダー連携には必須）' });
    if (body.key !== key) return _json({ ok: false, error: 'bad key' });
    var a = body.action;
    if (a === 'probe') return _json(_calProbe());
    if (a === 'read') return _json(_calRead(body));
    if (a === 'ops') {
      var lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try { return _json(_calOps(body.ops || [])); }
      finally { lock.releaseLock(); }
    }
    return _json({ ok: false, error: 'unknown action' });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

function _calGet(id) {
  if (!id) return null;
  try { return CalendarApp.getCalendarById(String(id)); } catch (x) { return null; }
}

function _tagOf(ev) {
  try { return ev.getTag(CONFIG.CAL.tagKey) || ''; } catch (x) { return ''; }
}

function _mine(ev, id) { return !!ev && _tagOf(ev) === String(id); }

function _declined(ev) {
  try { return ev.getMyStatus() === CalendarApp.GuestStatus.NO; } catch (x) { return false; }
}

function _evJson(ev, calName, own) {
  return {
    id: ev.getId(), t: ev.getTitle(),
    s: ev.getStartTime().getTime(), e: ev.getEndTime().getTime(),
    ad: ev.isAllDayEvent(), tag: _tagOf(ev), cal: calName, own: !!own
  };
}

function _uniq(a) { var o = [], m = {}; (a || []).forEach(function (x) { if (x && !m[x]) { m[x] = 1; o.push(x); } }); return o; }

/** 予定の一覧（空き時間の計算用）＋ リンク済みタスク予定の現状 */
function _calRead(b) {
  var DAY = 86400000, now = Date.now();
  var from = Number(b.from) || (now - DAY), to = Number(b.to) || (now + 14 * DAY);
  if (to <= from) to = from + DAY;
  if (to - from > CONFIG.CAL.maxRangeDays * DAY) to = from + CONFIG.CAL.maxRangeDays * DAY;
  var wid = CONFIG.CAL.writeId, wc = _calGet(wid);
  var ids = _uniq([wid].concat(CONFIG.CAL.readIds || []));
  var out = [], seen = {}, have = {}, errors = [];
  ids.forEach(function (id) {
    var cal = (id === wid) ? wc : _calGet(id);
    if (!cal) { errors.push('カレンダーが見つからない: ' + id); return; }
    var name = cal.getName(), own = (id === wid);
    cal.getEvents(new Date(from), new Date(to)).forEach(function (ev) {
      var k = ev.getId() + '|' + ev.getStartTime().getTime(); // 繰り返し予定はIDを共有するので開始時刻で区別
      if (seen[k] || _declined(ev)) return;
      seen[k] = 1;
      var j = _evJson(ev, name, own);
      if (own) have[j.id] = 1;
      out.push(j);
    });
  });
  // 範囲外に動かされたタスク予定も追えるように、IDで直接確認する
  var linked = {};
  (b.ids || []).slice(0, 200).forEach(function (id) {
    id = String(id);
    if (have[id] || !wc) return;
    var ev = null;
    try { ev = wc.getEventById(id); } catch (x) {}
    if (!ev) { linked[id] = null; return; }
    var st = ev.getStartTime().getTime();
    if (st >= from && st < to) { linked[id] = null; return; } // 範囲内なのに一覧に無い＝削除済み
    linked[id] = _evJson(ev, wc.getName(), true);
  });
  return { ok: true, from: from, to: to, cal: { id: wid, name: wc ? wc.getName() : '' },
           events: out, linked: linked, errors: errors };
}

/** タスク予定の作成・更新・削除（最大50件／回） */
function _calOps(ops) {
  var cal = _calGet(CONFIG.CAL.writeId);
  if (!cal) return { ok: false, error: '書き込み先カレンダーが見つかりません: ' + CONFIG.CAL.writeId };
  var res = [];
  (ops || []).slice(0, 50).forEach(function (op) {
    try { res.push(op && op.op === 'delete' ? _calDel(cal, op) : _calUpsert(cal, op)); }
    catch (err) { res.push({ id: op && op.id, ok: false, error: String(err) }); }
  });
  return { ok: true, results: res };
}

function _calUpsert(cal, op) {
  var id = String(op.id || '');
  if (!/^[\w-]{1,64}$/.test(id)) throw 'bad id';
  var s = Number(op.start), e = Number(op.end);
  if (!(s > 0 && e > s && e - s <= 12 * 3600000)) throw 'bad time';
  var title = String(op.title || '').replace(/[\r\n]+/g, ' ').slice(0, 200) || '(無題)';
  var ev = null;
  if (op.calId) {
    try { ev = cal.getEventById(String(op.calId)); } catch (x) {}
    if (ev && !_mine(ev, id)) ev = null; // 目印の違う予定は絶対に触らない
  }
  if (!ev) { // 応答が途切れて予定IDを受け取れなかった場合の二重作成防止
    var DAY = 86400000;
    var near = cal.getEvents(new Date(s - DAY), new Date(e + DAY));
    for (var i = 0; i < near.length; i++) { if (_mine(near[i], id)) { ev = near[i]; break; } }
  }
  if (!ev) {
    ev = cal.createEvent(title, new Date(s), new Date(e), {
      description: 'EGO LOCK のタスク\nアプリで完了にするか、タイトルの先頭に ✅ を付けると完了扱いになります。'
    });
    ev.setTag(CONFIG.CAL.tagKey, id);
  } else {
    if (ev.getTitle() !== title) ev.setTitle(title);
    if (ev.getStartTime().getTime() !== s || ev.getEndTime().getTime() !== e) ev.setTime(new Date(s), new Date(e));
  }
  try {
    var C = CalendarApp.EventColor;
    ev.setColor(op.done ? C.GRAY : (op.lane === 'j' ? C.MAUVE : C.CYAN));
  } catch (x) {}
  return { id: id, ok: true, calId: ev.getId(), s: s, e: e };
}

function _calDel(cal, op) {
  var id = String(op.id || ''), cid = String(op.calId || ''), ev = null;
  if (cid) { try { ev = cal.getEventById(cid); } catch (x) {} }
  if (ev && _mine(ev, id)) ev.deleteEvent();
  return { id: id, ok: true, deleted: true, calId: cid };
}

/** 接続テスト：書き込み先に予定を作って即削除できるか */
function _calProbe() {
  var out = { ok: true, write: { id: CONFIG.CAL.writeId, found: false }, read: [] };
  var wc = _calGet(CONFIG.CAL.writeId);
  if (wc) {
    out.write.found = true;
    out.write.name = wc.getName();
    try { out.write.owned = wc.isOwnedByMe(); } catch (x) {}
    try {
      var t0 = new Date(Date.now() + 400 * 86400000);
      var tev = wc.createEvent('EGO LOCK 接続テスト（自動で削除）', t0, new Date(t0.getTime() + 60000));
      tev.deleteEvent();
      out.write.writable = true;
    } catch (err) { out.write.writable = false; out.write.error = String(err); }
  }
  _uniq(CONFIG.CAL.readIds || []).forEach(function (id) {
    var c = _calGet(id);
    out.read.push({ id: id, found: !!c, name: c ? c.getName() : '' });
  });
  return out;
}

function _findCol(header, needle) {
  needle = String(needle).trim();
  var i = header.indexOf(needle);
  if (i >= 0) return i;
  for (var j = 0; j < header.length; j++) {
    if (header[j] && header[j].indexOf(needle) >= 0) return j;
  }
  return -1;
}

function _dateKey(v, tz) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var s = String(v).trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  var d = new Date(s);
  if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  return null;
}

function _num(v) {
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** デプロイ前の動作確認用（エディタで実行 → 実行ログにJSON） */
function _test() {
  var key = _key();
  var out = doGet({ parameter: { key: key } });
  Logger.log(out.getContent());
  // カレンダー連携の確認（初回はカレンダーの権限許可が出る）
  var p = doPost({ postData: { contents: JSON.stringify({ key: key, action: 'probe' }) } });
  Logger.log(p.getContent());
}
