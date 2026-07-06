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
 *  デプロイ手順は gas/README.md 参照（種類:ウェブアプリ / 実行:自分 / アクセス:全員）。
 */

var CONFIG = {
  KEY: '', // 任意の合言葉（設定推奨）。アプリの「連携キー」と一致させる。
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
  }
};

function doGet(e) {
  try {
    if (CONFIG.KEY) {
      var got = e && e.parameter ? e.parameter.key : '';
      if (got !== CONFIG.KEY) return _json({ ok: false, error: 'bad key' });
    }
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
  var out = doGet({ parameter: { key: CONFIG.KEY } });
  Logger.log(out.getContent());
}
