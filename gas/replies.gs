/**
 * replies.gs v0.2(2026-11-10 本部) Threadsの返信の自動化
 * 流れ: 1時間ごとに自分の投稿への返信を集める → 事業B(ハーピィ)が下書きを書き、J列に「自動」か空欄(承認待ち)を入れる
 *       → J列が「自動」か「OK」の行を、次の1時間ごとの実行で送る。J列に「NG」と書けば送らない。
 * 置き方: 投稿キューのApps Scriptに新しいファイル「replies」を作って貼る(poster_v2.gs は触らない)。
 * 事前準備: Metaのアプリに threads_read_replies と threads_manage_replies を追加 → poster_v2 のメニュー⑤⑥で
 *           4アカウント(main / kango / hoiku / daini)のThreadsを連携し直す(新しい権限は連携し直したトークンにだけ付く)。
 * 初回: setupReplies() を1回実行 → 「返信」タブと1時間ごとのトリガーができる。checkReplySetup() で確認。
 * 公式: GET /{media-id}/replies、POST /me/threads(reply_to_id)→約30秒→POST /me/threads_publish。返信は24時間で1,000件まで。
 */
var REPLY_CFG = {
  ACCOUNTS: ['main', 'kango', 'hoiku', 'daini'],   // main = フリーランス準備室(タブ queue)
  TAB: { main: 'queue', kango: 'kango', hoiku: 'hoiku', daini: 'daini' },
  SHEET: '返信',
  LOOKBACK_DAYS: 14,
  MAX_SEND_PER_RUN: 8,
  SEND_HOURS: [8, 23],   // 日本時間でこの時間帯だけ送る
  BASE: 'https://graph.threads.net/v1.0'
};
var REPLY_HEADER = ['取得日時', 'アカウント', '元投稿の行ID', '元投稿ID', '返信ID', '返信者', '返信本文', '返信日時',
  '返信の下書き(事業B)', '送信(自動/OK/NG)', '状態', '送信日時', '送信した返信ID', 'エラー', 'メモ(承認待ちの理由など)'];

function setupReplies() {
  replySheet_();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'runReplies') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('runReplies').timeBased().everyHours(1).create();
  Logger.log('返信タブと1時間ごとのトリガーを作りました。checkReplySetup() で確認できます。');
}

function runReplies() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try { collectReplies(); sendReplies(); } finally { lock.releaseLock(); }
}

/** 権限とトークンの確認(送信はしない。トークンは表示しない) */
function checkReplySetup() {
  REPLY_CFG.ACCOUNTS.forEach(function (acc) {
    var token = replyToken_(acc);
    if (!token) { Logger.log(acc + ': トークンなし(⑤⑥で連携してください)'); return; }
    var me = replyFetch_('get', '/me?fields=id,username', token);
    Logger.log(acc + ': /me ' + (me.ok ? 'OK @' + me.body.username : 'NG ' + me.text.slice(0, 200)));
    if (me.ok) {
      var lim = replyFetch_('get', '/me/threads_publishing_limit?fields=reply_quota_usage,reply_config', token);
      Logger.log(acc + ': 返信の権限 ' + (lim.ok ? 'OK ' + JSON.stringify(lim.body) : 'NG(連携し直しが必要) ' + lim.text.slice(0, 200)));
    }
  });
}

function collectReplies() {
  var sh = replySheet_(), known = {}, last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 5, last - 1, 1).getValues().forEach(function (r) { if (r[0]) known[String(r[0])] = true; });
  var ss = SpreadsheetApp.getActive(), since = new Date(Date.now() - REPLY_CFG.LOOKBACK_DAYS * 86400000), out = [];
  REPLY_CFG.ACCOUNTS.forEach(function (acc) {
    var token = replyToken_(acc), tab = ss.getSheetByName(REPLY_CFG.TAB[acc]);
    if (!token || !tab) return;
    var rows = tab.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (r[2] !== 'Threads' || r[6] !== '投稿済み' || !r[7]) continue;            // C媒体 G状態 H投稿ID
      if (REPLY_CFG.TAB[acc] === 'queue' && String(r[19] || '').trim()) continue; // queueのT列に別アカウント名がある行は除く
      var posted = r[8] instanceof Date ? r[8] : new Date(String(r[8]).replace(/\//g, '-'));
      if (!isNaN(posted) && posted < since) continue;
      var postId = String(r[7]).replace(/^'/, '');
      var res = replyFetch_('get', '/' + postId + '/replies?fields=id,text,username,timestamp,is_reply_owned_by_me,hide_status', token);
      if (!res.ok) { Logger.log(acc + ' ' + r[0] + ': 取得エラー ' + res.text.slice(0, 200)); continue; }
      (res.body.data || []).forEach(function (rep) {
        if (known[rep.id] || rep.is_reply_owned_by_me || rep.hide_status === 'HIDDEN') return;
        known[rep.id] = true;
        out.push([new Date(), acc, r[0], postId, String(rep.id), rep.username || '', rep.text || '', rep.timestamp || '',
          '', '', '新着', '', '', '', '']);
      });
      Utilities.sleep(300);
    }
  });
  if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, REPLY_HEADER.length).setValues(out);
}

function sendReplies() {
  var h = Number(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'H'));
  if (h < REPLY_CFG.SEND_HOURS[0] || h >= REPLY_CFG.SEND_HOURS[1]) return;
  var sh = replySheet_(), last = sh.getLastRow();
  if (last < 2) return;
  var vals = sh.getRange(2, 1, last - 1, REPLY_HEADER.length).getValues(), sent = 0, started = Date.now();
  for (var i = 0; i < vals.length && sent < REPLY_CFG.MAX_SEND_PER_RUN && Date.now() - started < 4.5 * 60000; i++) {
    var v = vals[i], row = i + 2, flag = String(v[9]).trim().toUpperCase();
    if (v[10] !== '下書き済み' || (flag !== '自動' && flag !== 'OK')) continue;
    var text = String(v[8]).trim(), err = replyCheckText_(text);
    if (err) { sh.getRange(row, 11).setValue('エラー'); sh.getRange(row, 14).setValue(err); continue; }
    var token = replyToken_(v[1]);
    if (!token) { sh.getRange(row, 14).setValue('トークンなし'); continue; }
    // 手で返信済みなら送らない
    var mine = replyFetch_('get', '/' + v[4] + '/replies?fields=id,is_reply_owned_by_me', token);
    if (mine.ok && (mine.body.data || []).some(function (x) { return x.is_reply_owned_by_me; })) {
      sh.getRange(row, 11).setValue('手で返信済み'); continue;
    }
    var c = replyFetch_('post', '/me/threads', token, { media_type: 'TEXT', text: text, reply_to_id: String(v[4]) });
    if (!c.ok) { sh.getRange(row, 11).setValue('エラー'); sh.getRange(row, 14).setValue('作成 ' + c.text.slice(0, 300)); continue; }
    Utilities.sleep(30000);
    var p = replyFetch_('post', '/me/threads_publish', token, { creation_id: c.body.id });
    if (!p.ok) { sh.getRange(row, 11).setValue('エラー'); sh.getRange(row, 14).setValue('公開 ' + p.text.slice(0, 300)); continue; }
    sh.getRange(row, 11, 1, 3).setValues([['送信済み', new Date(), String(p.body.id)]]);
    sh.getRange(row, 14).setValue('');
    sent++;
  }
}

/** 送る前の機械チェック(最低限) */
function replyCheckText_(text) {
  if (!text) return '下書きが空';
  if (text.length > 500) return '500字を超えている';
  if (/https?:\/\//.test(text)) return 'リンクを含む(自動では送らない)';
  if (/レバテック|ギークス|紹介コード|紹介URL|DM/.test(text)) return '紹介・DMの文言を含む(自動では送らない)';
  return '';
}

function replySheet_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(REPLY_CFG.SHEET);
  if (!sh) {
    sh = ss.insertSheet(REPLY_CFG.SHEET);
    sh.getRange(1, 1, 1, REPLY_HEADER.length).setValues([REPLY_HEADER]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** poster_v2 と同じ保存名: main は THREADS_TOKEN、それ以外は THREADS_TOKEN__アカウント名 */
function replyToken_(acc) {
  var key = acc === 'main' ? 'THREADS_TOKEN' : 'THREADS_TOKEN__' + acc;
  return PropertiesService.getScriptProperties().getProperty(key);
}

function replyFetch_(method, path, token, payload) {
  var url = REPLY_CFG.BASE + path + (path.indexOf('?') >= 0 ? '&' : '?') + 'access_token=' + encodeURIComponent(token);
  var opt = { method: method, muteHttpExceptions: true };
  if (payload) opt.payload = payload;
  var res = UrlFetchApp.fetch(url, opt), text = res.getContentText(), body = {};
  try { body = JSON.parse(text); } catch (e) {}
  return { ok: res.getResponseCode() === 200 && !body.error, body: body, text: text.split(token).join('***') };
}

/** Meta審査の前に必要な「検索APIの成功した呼び出し」を1回行う(審査前は自分の投稿だけが返る。送信はしない) */
function testKeywordSearch() {
  var token = replyToken_('main');
  if (!token) { Logger.log('main のトークンなし'); return; }
  var res = replyFetch_('get', '/keyword_search?q=' + encodeURIComponent('フリーランス') + '&search_type=RECENT&fields=id,text,username,timestamp', token);
  Logger.log('keyword_search: ' + (res.ok ? 'OK ' + (res.body.data || []).length + '件' : 'NG ' + res.text.slice(0, 300)));
}
