/**
 * マイナビ請求書 自動処理 v4 追加分
 * コード.gs の末尾にそのまま貼り付ける。既存コードは一切変更しない。
 *
 * 役割: エクセル/ の xlsx を変換サービスへ送り、返ってきたPDFを 送信前PDF/ に置く。
 *       その直後に既存の下書き生成を呼ぶので、着信から2分以内に下書きができる。
 *
 * 必要なスクリプトプロパティ
 *   CONVERTER_URL      https://mynavi-converter.clane.co.jp/convert
 *   CONVERTER_SECRET   合言葉（サーバー側 SHARED_SECRET と同じ値）
 * どちらか未設定の間は何もしないので、先に貼っても害はない。
 */

var MVC_YEAR_PARENT_ID = '1qp-9GsSDOiZ5JwPLgiAsNjTNrq8M2PvR'; // 「26年」フォルダ
var MVC_STAMP_FILE_ID  = '1sN2hec7sHxo1VjWD4yPBHjFXvXh_fPQu'; // 角印画像
var MVC_DRAFT_FN       = 'createMynaviDraftIfNeeded';          // 既存の下書き生成関数

// ---------------------------------------------------------------- entry

function convertAndStage_() {
  var props = PropertiesService.getScriptProperties();
  var url = props.getProperty('CONVERTER_URL');
  var secret = props.getProperty('CONVERTER_SECRET');
  if (!url || !secret) return; // 未設定なら沈黙して終了

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  try {
    var now = new Date();
    var month = now.getMonth() + 1;

    var monthFolder = mvcFindChild_(DriveApp.getFolderById(MVC_YEAR_PARENT_ID), month + '月');
    if (!monthFolder) return;

    var xlsxDir = mvcFindChild_(monthFolder, 'エクセル');
    var stageDir = mvcFindChild_(monthFolder, '送信前PDF');
    var sentDir = mvcFindChild_(monthFolder, 'PDF');
    if (!xlsxDir || !stageDir) return;

    if (sentDir && mvcCountPdf_(sentDir) >= 2) return; // 送信済み
    if (mvcCountPdf_(stageDir) >= 2) return;           // 生成済み

    var targets = mvcPickTargets_(xlsxDir);
    if (targets.length < 2) return; // テンプレが揃うまで待つ

    var stampB64 = null;
    try {
      stampB64 = Utilities.base64Encode(DriveApp.getFileById(MVC_STAMP_FILE_ID).getBlob().getBytes());
    } catch (e) {
      Logger.log('角印の読み込みに失敗（角印なしで続行）: ' + e);
    }

    var made = [];
    for (var i = 0; i < targets.length; i++) {
      var file = targets[i];
      var pdfName = file.getName().replace(/\.xlsx$/i, '') + '.pdf';
      if (mvcHasFile_(stageDir, pdfName)) { made.push(pdfName); continue; }

      var res = mvcCallConverter_(url, secret, file, stampB64);

      if (res.code === 409) {
        mvcNotifyOnce_('template-' + month,
          '🟠 マイナビ請求書 自動処理を停止しました（' + month + '月分）\n\n' +
          'テンプレートの様式が前月と変わっている可能性があります。\n' +
          res.detail + '\n\n' +
          'やること\n' +
          '1. Drive の ' + month + '月/エクセル/ を開いて様式を確認する\n' +
          '2. Claude に「マイナビのテンプレが変わった」と伝えて対象セルの修正を依頼する\n' +
          '3. 急ぎなら従来どおり手作業で作成・送信する\n\n' +
          '誤った請求書を送らないため、下書きは作られていません。');
        return;
      }
      if (res.code !== 200 || !res.pdfB64) {
        mvcNotifyOnce_('error-' + month + '-' + res.code,
          '⚠️ マイナビ請求書 自動処理が失敗しました（' + month + '月分・送信は行われていません）\n\n' +
          '工程 PDF変換\n' +
          '対象 ' + file.getName() + '\n' +
          '内容 ' + res.detail + '\n\n' +
          'やること\n' +
          '1. この通知を Claude に貼って原因調査を依頼する\n' +
          '2. 期日が近い場合は従来の手作業で作成・送信する\n\n' +
          'Excel原本はDriveに残っています。下書きは作られていないので誤送信の心配はありません。');
        return;
      }

      stageDir.createFile(
        Utilities.newBlob(Utilities.base64Decode(res.pdfB64), 'application/pdf', pdfName)
      );
      made.push(pdfName);
    }

    if (made.length >= 2) {
      mvcClearNotice_('template-' + month);
      Logger.log('送信前PDF生成: ' + made.join(' , '));
      try {
        if (typeof this[MVC_DRAFT_FN] === 'function') {
          this[MVC_DRAFT_FN]();
        }
      } catch (e) {
        Logger.log('下書き生成の呼び出しに失敗（10分トリガー側で再試行される）: ' + e);
      }
    }
  } catch (e) {
    Logger.log('convertAndStage_ 例外: ' + e);
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- helpers

function mvcCallConverter_(url, secret, file, stampB64) {
  var payload = {
    filename: file.getName(),
    xlsx_b64: Utilities.base64Encode(file.getBlob().getBytes()),
    kind: 'auto'
  };
  if (stampB64) payload.stamp_b64 = stampB64;

  var resp;
  try {
    resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'X-Shared-Secret': secret },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      followRedirects: false
    });
  } catch (e) {
    return { code: 0, detail: '変換サービスに接続できません: ' + e };
  }

  var code = resp.getResponseCode();
  var text = resp.getContentText();
  var body = {};
  try { body = JSON.parse(text); } catch (e) {}

  return {
    code: code,
    pdfB64: body.pdf_b64 || null,
    detail: body.detail || body.error || text.substring(0, 300)
  };
}

function mvcPickTargets_(dir) {
  // 「請求書」と「納品書」の2件だけを対象にする（納品明細は送付対象外）
  var wanted = [];
  var it = dir.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    var n = f.getName();
    if (!/\.xlsx$/i.test(n)) continue;
    if (n.indexOf('納品明細') >= 0) continue;
    if (n.indexOf('請求書') >= 0 || n.indexOf('納品書') >= 0) wanted.push(f);
  }
  return wanted;
}

function mvcFindChild_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : null;
}

function mvcCountPdf_(dir) {
  var n = 0;
  var it = dir.getFiles();
  while (it.hasNext()) { if (/\.pdf$/i.test(it.next().getName())) n++; }
  return n;
}

function mvcHasFile_(dir, name) {
  return dir.getFilesByName(name).hasNext();
}

/** 同じ内容の通知を24時間に1回に制限して Slack へ送る */
function mvcNotifyOnce_(key, text) {
  var props = PropertiesService.getScriptProperties();
  var pk = 'MVC_NOTICE_' + key;
  var last = Number(props.getProperty(pk) || 0);
  if (Date.now() - last < 24 * 60 * 60 * 1000) return;
  props.setProperty(pk, String(Date.now()));

  var token = props.getProperty('SLACK_BOT_TOKEN');
  var channel = props.getProperty('SLACK_CHANNEL_ID');
  if (!token || !channel) { Logger.log(text); return; }

  try {
    UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ channel: channel, text: text }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('Slack通知に失敗: ' + e);
  }
}

function mvcClearNotice_(key) {
  PropertiesService.getScriptProperties().deleteProperty('MVC_NOTICE_' + key);
}

// ---------------------------------------------------------------- setup

/**
 * convertAndStage_ の1分トリガーだけを設置する。
 * 既存トリガー（checkAndProcess / createMynaviDraftIfNeeded / collectSentPdfs）
 * には触らないので、既存の動作は変わらない。何度実行しても重複しない。
 */
function installConverterTrigger() {
  var all = ScriptApp.getProjectTriggers();
  for (var i = 0; i < all.length; i++) {
    if (all[i].getHandlerFunction() === 'convertAndStage_') {
      ScriptApp.deleteTrigger(all[i]);
    }
  }
  ScriptApp.newTrigger('convertAndStage_').timeBased().everyMinutes(1).create();

  var names = [];
  var after = ScriptApp.getProjectTriggers();
  for (var j = 0; j < after.length; j++) names.push(after[j].getHandlerFunction());
  Logger.log('現在のトリガー: ' + names.join(' , '));
}

/** 手動確認用。今すぐ1回だけ変換を試す。 */
function runConverterNow() {
  convertAndStage_();
  Logger.log('convertAndStage_ を1回実行しました。Driveの送信前PDF/を確認してください。');
}

/** 接続確認用。変換サービスの /health を叩く。 */
function pingConverter() {
  var props = PropertiesService.getScriptProperties();
  var url = props.getProperty('CONVERTER_URL');
  if (!url) { Logger.log('CONVERTER_URL が未設定です'); return; }
  var health = url.replace(/\/convert$/, '/health');
  var resp = UrlFetchApp.fetch(health, { muteHttpExceptions: true });
  Logger.log(resp.getResponseCode() + ' ' + resp.getContentText());
}
