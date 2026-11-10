/**
 * シュハ社 SNS自動投稿スクリプト v2(複数アカウント対応)
 * 今の poster.gs を「全部消して」このファイルの中身に置き換える。
 *
 * タブ: queue=今のフリーランス準備室(T列でアカウント指定も可)。それ以外の半角英字のタブ(kango / hoiku / daini など)は
 *       タブ名=アカウント名。列はどのタブも同じ。
 * 列(v1と同じ + T列「アカウント」を追加):
 *  A:ID B:予定日時 C:媒体(X/Threads) D:型 E:テーマ F:本文 G:状態(下書き/予約/投稿済み/エラー/スキップ)
 *  H:投稿ID I:投稿日時 J:エラー K:表示 L:いいね M:返信 N:リポスト O:リンククリック P:数値取得日時
 *  Q:Day R:時刻枠 S:検証メモ(事業Bが使う)
 *  T:アカウント(空欄=今のフリーランス準備室。新しいアカウントは kango / hoiku / daini など半角英字)
 *  U:画像URL(Threadsのみ。公開URLの画像があれば画像つきで投稿。Xは本文だけ投稿)
 *
 * スクリプトプロパティ(秘密情報はここだけに入る。シートやチャットには書かない):
 *  ENABLED                      … "true" のときだけ投稿する(緊急停止は "false")
 *  X_API_KEY / X_API_SECRET     … Xのアプリの鍵(全アカウント共通。今のまま)
 *  X_ACCESS_TOKEN / X_ACCESS_TOKEN_SECRET / THREADS_TOKEN / THREADS_USER_ID … 今のアカウント用(今のまま)
 *  X_ACCESS_TOKEN__kango など    … 新アカウント用。メニュー「SNS連携」から自動で入る(手で入れない)
 *  X_DAILY_CAP / THREADS_DAILY_CAP … 1アカウント1日の上限(既定 6 / 6)
 *
 * 使い方: スプレッドシートを開き直すと上部に「SNS連携」メニューが出る。
 */

const SHEET_NAME = 'queue';
const HEADERS = ['ID', '予定日時', '媒体', '型', 'テーマ', '本文', '状態', '投稿ID', '投稿日時', 'エラー',
  '表示', 'いいね', '返信', 'リポスト', 'リンククリック', '数値取得日時', 'Day', '時刻枠', '検証メモ', 'アカウント', '画像URL'];
const C = { id: 0, when: 1, platform: 2, type: 3, theme: 4, text: 5, status: 6, postId: 7, postedAt: 8, error: 9,
  views: 10, likes: 11, replies: 12, reposts: 13, clicks: 14, metricsAt: 15, account: 19, image: 20 };
const STALE_HOURS = 3;
const MAIN = 'main';

function P_() { return PropertiesService.getScriptProperties(); }
function prop_(k, d) { const v = P_().getProperty(k); return (v === null || v === '') ? d : v; }
function acc_(v) { const s = String(v || '').trim(); return s === '' ? MAIN : s; }
function key_(base, account) { return account === MAIN ? base : base + '__' + account; }
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  const head = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (head[C.account] !== 'アカウント') sh.getRange(1, C.account + 1).setValue('アカウント');
  return sh;
}
/** 投稿用のタブ一覧: queue と、名前が半角英字のタブ */
function tabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet_();
  return ss.getSheets().filter(sh => sh.getName() === SHEET_NAME || /^[a-z][a-z0-9]{1,15}$/.test(sh.getName()))
    .filter(sh => sh.getName() !== MAIN);
}
/** その行のアカウント名: queue はT列(空欄=main)、それ以外はタブ名 */
function rowAcc_(sh, r) { return sh.getName() === SHEET_NAME ? acc_(r[C.account]) : sh.getName(); }
/** タブを用意(なければ見出しつきで作る) */
function ensureTab_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  return sh;
}
function threadsHost_() { return prop_('THREADS_HOST', 'https://graph.threads.net'); }

// ---------- メニュー ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('SNS連携')
    .addItem('① Xを連携する(新アカウント)', 'menuXStart')
    .addItem('② XのPINを入力する', 'menuXFinish')
    .addItem('③ Threadsのトークンを登録する', 'menuThreads')
    .addItem('④ 下書きを予約にする(今から時刻を割り振る)', 'menuSchedule')
    .addSeparator()
    .addItem('⑤ Threadsをログインで連携する(URLを出す)', 'menuThLoginStart')
    .addItem('⑥ 戻ってきたURLを貼る', 'menuThLoginFinish')
    .addItem('⑦ 画像URLをテストする(投稿はしない)', 'menuImageTest')
    .addSeparator()
    .addItem('接続を確認する(全アカウント)', 'checkSetup')
    .addItem('トリガーを作り直す', 'setup')
    .addToUi();
}
function askAccount_(ui, title, allowMain) {
  const r = ui.prompt(title, 'アカウント名を半角英字で入力(例: kango' + (allowMain ? '。フリーランス準備室は main' : '') + ')', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return '';
  const a = r.getResponseText().trim();
  if (!/^[a-z][a-z0-9]{1,15}$/.test(a) || (a === MAIN && !allowMain)) { ui.alert('半角英字(小文字)で' + (allowMain ? '' : '、main以外に') + 'してください'); return ''; }
  return a;
}

/** ① Xの許可画面のURLを出す(PIN方式。コールバック設定は不要) */
function menuXStart() {
  const ui = SpreadsheetApp.getUi();
  const a = askAccount_(ui, 'X連携 ①');
  if (!a) return;
  const url = 'https://api.x.com/oauth/request_token';
  const res = UrlFetchApp.fetch(url, { method: 'post', muteHttpExceptions: true,
    headers: { Authorization: xAuthHeader_('POST', url, {}, { token: '', tokenSecret: '' }, { oauth_callback: 'oob' }) } });
  if (res.getResponseCode() !== 200) { ui.alert('Xの許可URLを作れませんでした: ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200)); return; }
  const q = parseQs_(res.getContentText());
  P_().setProperties({ X_PENDING_ACCOUNT: a, X_PENDING_TOKEN: q.oauth_token, X_PENDING_SECRET: q.oauth_token_secret });
  const link = 'https://api.x.com/oauth/authorize?oauth_token=' + q.oauth_token;
  const html = HtmlService.createHtmlOutput(
    '<p>1. ブラウザで<b>新しいアカウント(' + a + ')</b>でXにログインした状態にする</p>' +
    '<p>2. <a href="' + link + '" target="_blank">ここを開いて「連携アプリを認証」を押す</a></p>' +
    '<p>3. 表示された7桁のPINを控え、メニュー「② XのPINを入力する」に入力</p>').setWidth(460).setHeight(220);
  ui.showModalDialog(html, 'X連携 ①');
}

/** ② PINを入れると合鍵を保存する */
function menuXFinish() {
  const ui = SpreadsheetApp.getUi();
  const a = prop_('X_PENDING_ACCOUNT', '');
  if (!a) { ui.alert('先に「① Xを連携する」を実行してください'); return; }
  const r = ui.prompt('X連携 ②(' + a + ')', '表示されたPINを入力', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const url = 'https://api.x.com/oauth/access_token';
  const creds = { token: prop_('X_PENDING_TOKEN', ''), tokenSecret: prop_('X_PENDING_SECRET', '') };
  const res = UrlFetchApp.fetch(url, { method: 'post', muteHttpExceptions: true,
    headers: { Authorization: xAuthHeader_('POST', url, {}, creds, { oauth_verifier: r.getResponseText().trim() }) } });
  if (res.getResponseCode() !== 200) { ui.alert('連携に失敗しました: ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200)); return; }
  const q = parseQs_(res.getContentText());
  const set = {};
  set[key_('X_ACCESS_TOKEN', a)] = q.oauth_token;
  set[key_('X_ACCESS_TOKEN_SECRET', a)] = q.oauth_token_secret;
  P_().setProperties(set);
  ['X_PENDING_ACCOUNT', 'X_PENDING_TOKEN', 'X_PENDING_SECRET'].forEach(k => P_().deleteProperty(k));
  ui.alert('Xを連携しました: @' + (q.screen_name || '?') + '(アカウント名 ' + a + ')');
}

/** ③ Threadsのトークンを登録(Metaの画面で発行したもの)。長期化・ユーザーIDの取得は自動 */
function menuThreads() {
  const ui = SpreadsheetApp.getUi();
  const a = askAccount_(ui, 'Threads連携');
  if (!a) return;
  const r = ui.prompt('Threads連携(' + a + ')', 'Metaの画面で発行したアクセストークンを貼り付け', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  let token = r.getResponseText().trim();
  const secret = prop_('THREADS_APP_SECRET', '');
  if (secret) { // 短期トークンなら長期トークンに交換(アプリシークレットがある場合だけ)
    const ex = UrlFetchApp.fetch(threadsHost_() + '/access_token?grant_type=th_exchange_token&client_secret=' +
      encodeURIComponent(secret) + '&access_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
    if (ex.getResponseCode() === 200) token = JSON.parse(ex.getContentText()).access_token;
  }
  const me = UrlFetchApp.fetch(threadsHost_() + '/v1.0/me?fields=id,username&access_token=' + encodeURIComponent(token),
    { muteHttpExceptions: true });
  if (me.getResponseCode() !== 200) { ui.alert('トークンが使えませんでした: ' + me.getContentText().slice(0, 200)); return; }
  const j = JSON.parse(me.getContentText());
  const set = {};
  set[key_('THREADS_TOKEN', a)] = token;
  set[key_('THREADS_USER_ID', a)] = j.id;
  P_().setProperties(set);
  refreshThreadsToken_(a); // 長期トークンなら期限を延ばしておく
  ui.alert('Threadsを連携しました: @' + j.username + '(アカウント名 ' + a + ')');
}

/** ⑤ Threadsの許可画面のURLを出す(トークン生成ツールを使わない方法)
 *  必要: アプリ設定の「コールバックURLをリダイレクト」に https://www.google.com/ を登録
 *        スクリプトプロパティ THREADS_APP_SECRET(アプリの app secret) */
const TH_APP_ID_DEFAULT = '1615645530298763';
function thRedirect_() { return prop_('THREADS_REDIRECT', 'https://www.google.com/'); }
function menuThLoginStart() {
  const ui = SpreadsheetApp.getUi();
  if (!prop_('THREADS_APP_SECRET', '')) { ui.alert('スクリプトプロパティに THREADS_APP_SECRET がありません。先に登録してください。'); return; }
  const a = askAccount_(ui, 'Threads連携 ⑤', true);
  if (!a) return;
  P_().setProperty('TH_PENDING_ACCOUNT', a);
  const link = 'https://threads.net/oauth/authorize?client_id=' + prop_('THREADS_APP_ID', TH_APP_ID_DEFAULT) +
    '&redirect_uri=' + encodeURIComponent(thRedirect_()) +
    '&scope=threads_basic,threads_content_publish,threads_manage_insights,threads_read_replies,threads_manage_replies,threads_keyword_search&response_type=code&state=' + a;
  const html = HtmlService.createHtmlOutput(
    '<p>1. 下のURLをコピーし、<b>シークレットウィンドウ</b>で開く</p>' +
    '<textarea style="width:100%;height:70px" onclick="this.select()">' + link + '</textarea>' +
    '<p>2. <b>' + a + ' のアカウント</b>でログインして「許可」を押す</p>' +
    '<p>3. Googleのページに移ったら、<b>アドレスバーのURLを全部コピー</b>し、メニュー「⑥ 戻ってきたURLを貼る」に貼る(1時間以内)</p>')
    .setWidth(480).setHeight(300);
  ui.showModalDialog(html, 'Threads連携 ⑤(' + a + ')');
}
/** ⑥ 戻ってきたURLから合鍵を作って保存(長期トークンに交換) */
function menuThLoginFinish() {
  const ui = SpreadsheetApp.getUi();
  const a = prop_('TH_PENDING_ACCOUNT', '');
  if (!a) { ui.alert('先に「⑤ Threadsをログインで連携する」を実行してください'); return; }
  const r = ui.prompt('Threads連携 ⑥(' + a + ')', 'Googleのページに移ったときのURLを全部貼り付け', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const m = r.getResponseText().trim().match(/[?&]code=([^&#]+)/);
  if (!m) { ui.alert('URLの中に code= が見つかりません。もう一度⑤からやり直してください'); return; }
  const res = UrlFetchApp.fetch(threadsHost_() + '/oauth/access_token', { method: 'post', muteHttpExceptions: true,
    payload: { client_id: prop_('THREADS_APP_ID', TH_APP_ID_DEFAULT), client_secret: prop_('THREADS_APP_SECRET', ''),
      grant_type: 'authorization_code', redirect_uri: thRedirect_(), code: decodeURIComponent(m[1]) } });
  if (res.getResponseCode() !== 200) { ui.alert('連携に失敗しました: ' + res.getContentText().slice(0, 200)); return; }
  let token = JSON.parse(res.getContentText()).access_token;
  const ex = UrlFetchApp.fetch(threadsHost_() + '/access_token?grant_type=th_exchange_token&client_secret=' +
    encodeURIComponent(prop_('THREADS_APP_SECRET', '')) + '&access_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
  if (ex.getResponseCode() === 200) token = JSON.parse(ex.getContentText()).access_token;
  const me = UrlFetchApp.fetch(threadsHost_() + '/v1.0/me?fields=id,username&access_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
  if (me.getResponseCode() !== 200) { ui.alert('トークンが使えませんでした: ' + me.getContentText().slice(0, 200)); return; }
  const j = JSON.parse(me.getContentText());
  const set = {};
  set[key_('THREADS_TOKEN', a)] = token;
  set[key_('THREADS_USER_ID', a)] = j.id;
  P_().setProperties(set);
  P_().deleteProperty('TH_PENDING_ACCOUNT');
  ui.alert('Threadsを連携しました: @' + j.username + '(アカウント名 ' + a + ')');
}

/** ⑦ 画像URLをThreadsが読めるか確かめる(コンテナを作るだけで公開しない) */
function menuImageTest() {
  const ui = SpreadsheetApp.getUi();
  const a = askAccount_(ui, '画像URLのテスト');
  if (!a) return;
  const r = ui.prompt('画像URLのテスト(' + a + ')', '画像のURLを貼り付け', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const token = prop_(key_('THREADS_TOKEN', a), '');
  const r1 = UrlFetchApp.fetch(threadsHost_() + '/v1.0/' + threadsUserId_(a) + '/threads', { method: 'post', muteHttpExceptions: true,
    payload: { media_type: 'IMAGE', image_url: r.getResponseText().trim(), text: 'test', access_token: token } });
  if (r1.getResponseCode() !== 200) { ui.alert('NG(作成): ' + r1.getContentText().slice(0, 200)); return; }
  const id = JSON.parse(r1.getContentText()).id;
  Utilities.sleep(20000);
  const st = UrlFetchApp.fetch(threadsHost_() + '/v1.0/' + id + '?fields=status,error_message&access_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
  ui.alert('結果: ' + st.getContentText().slice(0, 200) + '\n(FINISHED ならOK。公開はしていません)');
}

/** ④ 指定アカウントの「下書き」を、今から順に時刻を割り振って「予約」にする
 *  X は約2.5〜3.5時間おき、Threads は約4〜6時間おき。7:00〜23:30 の外は翌朝に回す */
function menuSchedule() {
  const ui = SpreadsheetApp.getUi();
  const a = askAccount_(ui, '下書きを予約にする');
  if (!a) return;
  const res = scheduleAccount_(a);
  const left = Object.keys(res.skipped).map(k => k + ' ' + res.skipped[k] + '件').join('、');
  ui.alert(a + ' の下書き ' + res.n + ' 件を予約にしました。' + (left ? '\n未連携のため下書きのまま: ' + left : ''));
}
/** 毎朝: main以外の全アカウントの下書きを自動で予約にする(連携済みの媒体だけ) */
function autoSchedule() {
  if (prop_('ENABLED', 'false') !== 'true') return;
  accounts_().filter(a => a !== MAIN).forEach(a => {
    const res = scheduleAccount_(a);
    if (res.n) Logger.log(a + ': ' + res.n + '件を予約');
  });
}
/** 下書きを予約にする本体。すでにある予約の最後の時刻の後ろに続けて割り振る */
function scheduleAccount_(a) {
  const gap = { X: [150, 210], Threads: [120, 180] }; // Threadsは2〜3時間おき(1日6本前後)
  const now = Date.now();
  const next = { X: new Date(now + 10 * 60000), Threads: new Date(now + 40 * 60000) };
  const tabs = tabs_().map(sh => ({ sh: sh, rows: sh.getDataRange().getValues() }));
  tabs.forEach(tb => tb.rows.slice(1).forEach(r => { // 既存の予約の後ろにつなげる
    if (r[C.status] !== '予約' || rowAcc_(tb.sh, r) !== a || !next[r[C.platform]]) return;
    const w = parseWhen_(r[C.when]), pf = r[C.platform];
    if (!isNaN(w) && w.getTime() + gap[pf][0] * 60000 > next[pf].getTime()) next[pf] = new Date(w.getTime() + gap[pf][0] * 60000);
  }));
  let n = 0;
  const skipped = {};
  tabs.forEach(tb => {
    const sh = tb.sh, rows = tb.rows;
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (r[C.status] !== '下書き' || rowAcc_(sh, r) !== a) continue;
      const pf = r[C.platform];
      if (!next[pf]) continue;
      if (!ready_(a, pf)) { skipped[pf] = (skipped[pf] || 0) + 1; continue; } // 未連携の媒体は下書きのまま残す
      if (!String(r[C.text]).trim()) continue;
      const t = fitWindow_(next[pf]);
      sh.getRange(i + 1, C.when + 1).setValue(Utilities.formatDate(t, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'));
      sh.getRange(i + 1, C.status + 1).setValue('予約');
      const g = gap[pf][0] + Math.floor(Math.random() * (gap[pf][1] - gap[pf][0]));
      next[pf] = new Date(t.getTime() + g * 60000);
      n++;
    }
  });
  return { n: n, skipped: skipped };
}
function fitWindow_(d) {
  const hm = Number(Utilities.formatDate(d, 'Asia/Tokyo', 'H')) * 60 + Number(Utilities.formatDate(d, 'Asia/Tokyo', 'm'));
  if (hm >= 7 * 60 && hm <= 23 * 60 + 30) return d;
  const day = Utilities.formatDate(new Date(d.getTime() + (hm > 23 * 60 + 30 ? 86400000 : 0)), 'Asia/Tokyo', 'yyyy-MM-dd');
  const mm = 30 + Math.floor(Math.random() * 40); // 7:30〜8:09
  return new Date(day + 'T0' + (7 + Math.floor(mm / 60)) + ':' + ('0' + (mm % 60)).slice(-2) + ':00+09:00');
}

/** トリガーを作る(初回とメニューから) */
function setup() {
  sheet_();
  accounts_().filter(a => a !== MAIN).forEach(a => ensureTab_(a));
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('postDue').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('fetchMetrics').timeBased().everyDays(1).atHour(5).create();
  ScriptApp.newTrigger('autoSchedule').timeBased().everyDays(1).atHour(6).create();
  ScriptApp.newTrigger('refreshThreadsToken').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(4).create();
  if (prop_('ENABLED', '') === '') P_().setProperty('ENABLED', 'false');
  Logger.log('トリガーを作成しました。');
}

/** 接続確認(投稿はしない)。全アカウントの結果をログとダイアログに出す */
function checkSetup() {
  const out = [];
  accounts_().forEach(a => {
    if (prop_(key_('X_ACCESS_TOKEN', a), '')) {
      const x = xRequest_('GET', 'https://api.x.com/2/users/me', {}, null, a);
      out.push(a + ' X: ' + x.getResponseCode() + ' ' + x.getContentText().slice(0, 120));
    }
    const tk = prop_(key_('THREADS_TOKEN', a), '');
    if (tk) {
      const t = UrlFetchApp.fetch(threadsHost_() + '/v1.0/me?fields=id,username&access_token=' + encodeURIComponent(tk), { muteHttpExceptions: true });
      out.push(a + ' Threads: ' + t.getResponseCode() + ' ' + t.getContentText().slice(0, 120));
      if (t.getResponseCode() === 200) P_().setProperty(key_('THREADS_USER_ID', a), JSON.parse(t.getContentText()).id);
    }
  });
  Logger.log(out.join('\n'));
  try { SpreadsheetApp.getUi().alert(out.join('\n') || '連携済みのアカウントがありません'); } catch (e) {}
}

/** 連携済みのアカウント一覧 */
function accounts_() {
  const set = new Set([MAIN]);
  Object.keys(P_().getProperties()).forEach(k => { const m = k.match(/^(?:X_ACCESS_TOKEN|THREADS_TOKEN)__([a-z0-9]+)$/); if (m) set.add(m[1]); });
  return [...set];
}
function ready_(a, pf) {
  return pf === 'X' ? !!prop_(key_('X_ACCESS_TOKEN', a), '') : !!prop_(key_('THREADS_TOKEN', a), '');
}

/** 10分ごと: 予定時刻を過ぎた予約を、アカウント×媒体ごとに1件ずつ投稿する */
function postDue() {
  if (prop_('ENABLED', 'false') !== 'true') return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const tabs = tabs_().map(sh => ({ sh: sh, rows: sh.getDataRange().getValues() }));
    const now = new Date();
    const caps = { X: Number(prop_('X_DAILY_CAP', 6)), Threads: Number(prop_('THREADS_DAILY_CAP', 6)) };
    const today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
    const postedToday = {}, done = {}, recent = new Set();
    tabs.forEach(tb => tb.rows.slice(1).forEach(r => {
      if (r[C.status] !== '投稿済み') return;
      const k = rowAcc_(tb.sh, r) + '|' + r[C.platform];
      if (r[C.postedAt] && Utilities.formatDate(new Date(r[C.postedAt]), 'Asia/Tokyo', 'yyyy-MM-dd') === today) postedToday[k] = (postedToday[k] || 0) + 1;
      recent.add(k + '|' + String(r[C.text]).trim());
    }));
    const started = Date.now();
    for (const tb of tabs) {
    const sh = tb.sh, rows = tb.rows;
    for (let i = 1; i < rows.length; i++) {
      if (Date.now() - started > 4.5 * 60 * 1000) break; // 実行時間の上限(6分)に余裕を残す
      const r = rows[i];
      if (r[C.status] !== '予約') continue;
      const pf = r[C.platform], a = rowAcc_(sh, r), k = a + '|' + pf;
      if (pf !== 'X' && pf !== 'Threads') continue;
      const when = parseWhen_(r[C.when]);
      if (isNaN(when) || when > now) continue;
      const rowNo = i + 1;
      if (now - when > STALE_HOURS * 3600 * 1000) { setRow_(sh, rowNo, { status: 'スキップ', error: '予定時刻から' + STALE_HOURS + '時間以上経過' }); continue; }
      if (!ready_(a, pf)) continue; // まだ連携していないアカウントは待つ
      if (done[k] || (postedToday[k] || 0) >= caps[pf]) continue;
      const text = String(r[C.text]).trim();
      if (!text) { setRow_(sh, rowNo, { status: 'エラー', error: '本文が空' }); continue; }
      if (recent.has(k + '|' + text)) { setRow_(sh, rowNo, { status: 'エラー', error: '同じ本文を投稿済み(重複防止)' }); continue; }
      if (pf === 'X' && countChars_(text) > 280) { setRow_(sh, rowNo, { status: 'エラー', error: 'Xの文字数上限超過' }); continue; }
      if (pf === 'Threads' && text.length > 500) { setRow_(sh, rowNo, { status: 'エラー', error: 'Threadsの500字超過' }); continue; }
      try {
        const id = pf === 'X' ? postX_(text, a) : postThreads_(text, a, String(r[C.image] || '').trim());
        setRow_(sh, rowNo, { status: '投稿済み', postId: id, postedAt: new Date(), error: '' });
        postedToday[k] = (postedToday[k] || 0) + 1;
        recent.add(k + '|' + text);
      } catch (e) {
        setRow_(sh, rowNo, { status: 'エラー', error: String(e).slice(0, 300) });
      }
      done[k] = true;
      SpreadsheetApp.flush();
    }
    }
  } finally {
    lock.releaseLock();
  }
}

function parseWhen_(v) {
  if (v instanceof Date) return v;
  const m = String(v).trim().match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})[ T](\d{1,2}):(\d{2})/);
  if (!m) return new Date(NaN);
  const p = n => ('0' + n).slice(-2);
  return new Date(m[1] + '-' + p(m[2]) + '-' + p(m[3]) + 'T' + p(m[4]) + ':' + m[5] + ':00+09:00');
}
function setRow_(sh, rowNo, v) {
  if (v.status !== undefined) sh.getRange(rowNo, C.status + 1).setValue(v.status);
  if (v.postId !== undefined) sh.getRange(rowNo, C.postId + 1).setValue("'" + v.postId);
  if (v.postedAt !== undefined) sh.getRange(rowNo, C.postedAt + 1).setValue(v.postedAt);
  if (v.error !== undefined) sh.getRange(rowNo, C.error + 1).setValue(v.error);
}
function countChars_(s) {
  const noUrl = s.replace(/https?:\/\/\S+/g, '');
  const urls = (s.match(/https?:\/\/\S+/g) || []).length;
  let n = urls * 23;
  for (const ch of noUrl) n += (ch.charCodeAt(0) <= 0x7f) ? 1 : 2;
  return n;
}
function parseQs_(s) {
  const o = {};
  s.split('&').forEach(p => { const i = p.indexOf('='); if (i > 0) o[decodeURIComponent(p.slice(0, i))] = decodeURIComponent(p.slice(i + 1)); });
  return o;
}

// ---------- X(OAuth 1.0a) ----------
function enc_(s) {
  return encodeURIComponent(String(s)).replace(/[!'()*]/g, c => '%' + c.charCodeAt(0).toString(16).toUpperCase());
}
function xCreds_(a) {
  return { token: prop_(key_('X_ACCESS_TOKEN', a), ''), tokenSecret: prop_(key_('X_ACCESS_TOKEN_SECRET', a), '') };
}
function xAuthHeader_(method, url, query, creds, extra) {
  const oauth = {
    oauth_consumer_key: prop_('X_API_KEY', ''),
    oauth_nonce: Utilities.getUuid().replace(/-/g, ''),
    oauth_signature_method: 'HMAC-SHA1',
    oauth_timestamp: String(Math.floor(Date.now() / 1000)),
    oauth_version: '1.0'
  };
  if (creds.token) oauth.oauth_token = creds.token;
  Object.assign(oauth, extra || {});
  const all = Object.assign({}, query, oauth);
  const paramStr = Object.keys(all).sort().map(k => enc_(k) + '=' + enc_(all[k])).join('&');
  const base = [method.toUpperCase(), enc_(url), enc_(paramStr)].join('&');
  const key = enc_(prop_('X_API_SECRET', '')) + '&' + enc_(creds.tokenSecret || '');
  oauth.oauth_signature = Utilities.base64Encode(
    Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_1, base, key));
  return 'OAuth ' + Object.keys(oauth).sort().map(k => enc_(k) + '="' + enc_(oauth[k]) + '"').join(', ');
}
function xRequest_(method, url, query, body, a) {
  const qs = Object.keys(query || {}).map(k => enc_(k) + '=' + enc_(query[k])).join('&');
  const opt = { method: method.toLowerCase(), headers: { Authorization: xAuthHeader_(method, url, query || {}, xCreds_(a || MAIN)) },
    muteHttpExceptions: true };
  if (body) { opt.contentType = 'application/json'; opt.payload = JSON.stringify(body); }
  return UrlFetchApp.fetch(url + (qs ? '?' + qs : ''), opt);
}
function postX_(text, a) {
  const res = xRequest_('POST', 'https://api.x.com/2/tweets', {}, { text: text }, a);
  if (res.getResponseCode() !== 201) throw new Error('X ' + res.getResponseCode() + ' ' + res.getContentText());
  return JSON.parse(res.getContentText()).data.id;
}

// ---------- Threads ----------
function threadsUserId_(a) {
  let id = prop_(key_('THREADS_USER_ID', a), '');
  if (!id) {
    const t = UrlFetchApp.fetch(threadsHost_() + '/v1.0/me?fields=id&access_token=' + encodeURIComponent(prop_(key_('THREADS_TOKEN', a), '')), { muteHttpExceptions: true });
    if (t.getResponseCode() === 200) { id = JSON.parse(t.getContentText()).id; P_().setProperty(key_('THREADS_USER_ID', a), id); }
  }
  if (!id) throw new Error('THREADS_USER_IDを取得できません(' + a + ')');
  return id;
}
function postThreads_(text, a, imageUrl) {
  const uid = threadsUserId_(a);
  const token = prop_(key_('THREADS_TOKEN', a), '');
  const payload = imageUrl ? { media_type: 'IMAGE', image_url: imageUrl, text: text, access_token: token }
                           : { media_type: 'TEXT', text: text, access_token: token };
  const r1 = UrlFetchApp.fetch(threadsHost_() + '/v1.0/' + uid + '/threads', { method: 'post', payload: payload, muteHttpExceptions: true });
  if (r1.getResponseCode() !== 200) throw new Error('Threads作成 ' + r1.getResponseCode() + ' ' + r1.getContentText());
  const creationId = JSON.parse(r1.getContentText()).id;
  Utilities.sleep(30000);
  if (imageUrl) { // 画像の取り込みが終わったか確認
    const st = UrlFetchApp.fetch(threadsHost_() + '/v1.0/' + creationId + '?fields=status,error_message&access_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
    const sj = st.getResponseCode() === 200 ? JSON.parse(st.getContentText()) : {};
    if (sj.status === 'ERROR' || sj.status === 'EXPIRED') throw new Error('Threads画像 ' + st.getContentText().slice(0, 200));
  }
  const r2 = UrlFetchApp.fetch(threadsHost_() + '/v1.0/' + uid + '/threads_publish', {
    method: 'post', payload: { creation_id: creationId, access_token: token }, muteHttpExceptions: true });
  if (r2.getResponseCode() !== 200) throw new Error('Threads公開 ' + r2.getResponseCode() + ' ' + r2.getContentText());
  return JSON.parse(r2.getContentText()).id;
}
/** 週1回: 全アカウントのThreadsトークンを延長する(60日で切れるのを防ぐ) */
function refreshThreadsToken() { accounts_().forEach(a => refreshThreadsToken_(a)); }
function refreshThreadsToken_(a) {
  const tk = prop_(key_('THREADS_TOKEN', a), '');
  if (!tk) return;
  const res = UrlFetchApp.fetch(threadsHost_() + '/refresh_access_token?grant_type=th_refresh_token&access_token=' +
    encodeURIComponent(tk), { muteHttpExceptions: true });
  if (res.getResponseCode() === 200) P_().setProperty(key_('THREADS_TOKEN', a), JSON.parse(res.getContentText()).access_token);
  else Logger.log('Threadsトークン更新失敗(' + a + '): ' + res.getContentText().slice(0, 200));
}

// ---------- 数値の取得(1日1回、投稿から7日以内) ----------
function fetchMetrics() {
  tabs_().forEach(sh => fetchMetricsTab_(sh));
}
function fetchMetricsTab_(sh) {
  const rows = sh.getDataRange().getValues();
  const now = new Date();
  const xByAcc = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r[C.status] !== '投稿済み' || !r[C.postId] || !r[C.postedAt]) continue;
    if (now - new Date(r[C.postedAt]) > 7 * 86400 * 1000) continue;
    const id = String(r[C.postId]).replace(/^'/, ''), a = rowAcc_(sh, r);
    if (r[C.platform] === 'X') { (xByAcc[a] = xByAcc[a] || {})[id] = i + 1; }
    if (r[C.platform] === 'Threads') {
      try {
        const res = UrlFetchApp.fetch(threadsHost_() + '/v1.0/' + id + '/insights?metric=views,likes,replies,reposts&access_token=' +
          encodeURIComponent(prop_(key_('THREADS_TOKEN', a), '')), { muteHttpExceptions: true });
        if (res.getResponseCode() === 200) {
          const m = {};
          JSON.parse(res.getContentText()).data.forEach(d => { m[d.name] = d.values ? d.values[0].value : (d.total_value || {}).value; });
          sh.getRange(i + 1, C.views + 1, 1, 6).setValues([[m.views || 0, m.likes || 0, m.replies || 0, m.reposts || 0, '', new Date()]]);
        }
      } catch (e) { Logger.log(e); }
    }
  }
  Object.keys(xByAcc).forEach(a => {
    const map = xByAcc[a], ids = Object.keys(map);
    for (let k = 0; k < ids.length; k += 100) {
      const res = xRequest_('GET', 'https://api.x.com/2/tweets',
        { ids: ids.slice(k, k + 100).join(','), 'tweet.fields': 'public_metrics,non_public_metrics' }, null, a);
      if (res.getResponseCode() !== 200) { Logger.log('X数値取得失敗(' + a + '): ' + res.getContentText().slice(0, 200)); continue; }
      (JSON.parse(res.getContentText()).data || []).forEach(t => {
        const pm = t.public_metrics || {}, npm = t.non_public_metrics || {};
        sh.getRange(map[t.id], C.views + 1, 1, 6).setValues([[pm.impression_count || npm.impression_count || 0,
          pm.like_count || 0, pm.reply_count || 0, (pm.retweet_count || 0) + (pm.quote_count || 0), npm.url_link_clicks || '', new Date()]]);
      });
    }
  });
}
