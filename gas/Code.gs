/**
 * EGO LOCK — Sheet Bridge (Google Apps Script Web App)
 * ---------------------------------------------------------------
 * 非公開スプレッドシートの必要な数値だけをJSONで返す中継エンドポイント。
 * これを「ウェブアプリ」としてデプロイし、その /exec URL をアプリの
 * 設定 > シート連携 に貼るだけで連携できる。
 *
 * ■ セットアップ手順
 *  1) script.google.com で新規プロジェクトを作成、このコードを貼る
 *  2) 下の CONFIG を自分のシートに合わせて編集
 *     - sheetId  … スプレッドシートURLの /d/ と /edit の間の文字列
 *     - tab      … シート(タブ)名。空文字なら先頭シート
 *     - headerRow… 見出しがある行番号（通常 1）
 *     - dateCol  … 日付列の見出し（部分一致でOK）
 *     - cols     … 取得したい値の {出力キー: 見出し(部分一致)}
 *  3) （任意）KEY にランダムな文字列を設定するとURLを知られても叩けなくなる
 *  4) 右上「デプロイ > 新しいデプロイ > 種類:ウェブアプリ」
 *     - 実行するユーザー: 自分
 *     - アクセスできるユーザー: 全員（← fetchで読むために必須。KEYで保護推奨）
 *  5) 生成された /exec URL をアプリに貼る（KEY を設定したら ?key=... は
 *     アプリ側の「連携キー」欄に入れる）
 *
 * ※ シートは公開せず非公開のまま。GASがあなたの権限で読むだけ。
 */

var CONFIG = {
  KEY: '', // 任意の共有シークレット（空なら誰でも読める。設定推奨）
  DAYS_BACK: 180, // 何日分の日次データを返すか
  SOURCES: [
    {
      // ① 日次: ライン追加数・AD数・契約数
      sheetId: '112RxU7evib3RqCxt-IJDClvIZ9JeQ93RSm4IeqLbRt4',
      tab: '',          // 例: '日次' 空なら先頭シート
      headerRow: 1,
      dateCol: '日付',   // 日付列の見出し（部分一致）
      cols: {
        lineAdds: 'ライン', // 「ライン追加数」等に部分一致
        ad:       'AD',
        contracts:'契約'
      }
    },
    {
      // ② お月謝回収金額（売上メイン）
      sheetId: '1PVbJmO3oG9-M2fCgem4ugO79BpWHSKXlpvz082Cy-kM',
      tab: '',          // 例: '入金管理' 空なら先頭シート
      headerRow: 1,
      dateCol: '日付',   // 入金日 or 回収日の列見出し（部分一致）
      cols: {
        revenue: 'お月謝' // 「お月謝回収金額」等に部分一致。なければ '回収' '入金' 等に変更
      }
    }
  ]
};

function doGet(e) {
  try {
    if (CONFIG.KEY) {
      var got = e && e.parameter ? e.parameter.key : '';
      if (got !== CONFIG.KEY) return _json({ ok: false, error: 'bad key' });
    }
    var tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
    var days = {}; // 'yyyy-MM-dd' -> { lineAdds, ad, contracts, revenue }
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - CONFIG.DAYS_BACK);

    CONFIG.SOURCES.forEach(function (src) {
      var ss = SpreadsheetApp.openById(src.sheetId);
      var sh = src.tab ? ss.getSheetByName(src.tab) : ss.getSheets()[0];
      if (!sh) return;
      var values = sh.getDataRange().getValues();
      var hr = (src.headerRow || 1) - 1;
      if (values.length <= hr) return;
      var header = values[hr].map(function (h) { return String(h).trim(); });

      var dateIdx = _findCol(header, src.dateCol);
      if (dateIdx < 0) return;
      var map = {}; // outKey -> colIndex
      Object.keys(src.cols).forEach(function (k) {
        map[k] = _findCol(header, src.cols[k]);
      });

      for (var r = hr + 1; r < values.length; r++) {
        var row = values[r];
        var dk = _dateKey(row[dateIdx], tz);
        if (!dk) continue;
        var d = new Date(dk);
        if (d < cutoff) continue;
        if (!days[dk]) days[dk] = {};
        Object.keys(map).forEach(function (k) {
          var ci = map[k];
          if (ci < 0) return;
          var n = _num(row[ci]);
          days[dk][k] = (days[dk][k] || 0) + n; // 同日複数行は合算
        });
      }
    });

    // 当月サマリ（利便性のため事前集計）
    var now = new Date();
    var ym = Utilities.formatDate(now, tz, 'yyyy-MM');
    var month = { lineAdds: 0, ad: 0, contracts: 0, revenue: 0 };
    Object.keys(days).forEach(function (dk) {
      if (dk.indexOf(ym) === 0) {
        var v = days[dk];
        month.lineAdds += v.lineAdds || 0;
        month.ad += v.ad || 0;
        month.contracts += v.contracts || 0;
        month.revenue += v.revenue || 0;
      }
    });

    return _json({
      ok: true,
      updated: Utilities.formatDate(now, tz, "yyyy-MM-dd'T'HH:mm:ssXXX"),
      tz: tz,
      count: Object.keys(days).length,
      month: month,
      days: days
    });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

function _findCol(header, needle) {
  needle = String(needle).trim();
  // 完全一致優先 → 部分一致
  var i = header.indexOf(needle);
  if (i >= 0) return i;
  for (var j = 0; j < header.length; j++) {
    if (header[j] && header[j].indexOf(needle) >= 0) return j;
  }
  return -1;
}

function _dateKey(v, tz) {
  if (v instanceof Date && !isNaN(v)) {
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  if (!s) return null;
  // 2026/7/3, 2026-07-03, 2026.07.03 などを許容
  var m = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
  if (m) {
    return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  }
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
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** デプロイ前の動作確認用（エディタで実行してログを見る） */
function _test() {
  var out = doGet({ parameter: { key: CONFIG.KEY } });
  Logger.log(out.getContent());
}
