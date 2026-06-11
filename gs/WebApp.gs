// ─── Web App エントリーポイント ────────────────────────────────────────────────

function doGet() {
  // ─── 必須シート存在チェック（Info + 管理IDリスト） ──────────────────────────
  // ※ メール残数チェックは「申込期間内」のときだけ行うため、クライアント側
  //   （onConfigLoaded）で getConfig の mailQuotaOk を見て判定する。
  const missing = [];

  const infoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INFO_SHEET_NAME);
  if (!infoSheet) missing.push('Info シート（メインスプレッドシート）');

  let kanriSheet = null;
  try {
    kanriSheet = SpreadsheetApp.openById(KANRI_SS_ID).getSheetByName(KANRI_SHEET);
  } catch (e) { /* アクセス不可 = なし扱い */ }
  if (!kanriSheet) missing.push(`管理IDリスト シート（ファイルID: ${KANRI_SS_ID}）`);

  if (missing.length > 0) {
    // 警告メール送信（ファイルオーナー宛・HTMLメールで折り返し防止）
    // ★ 重複送信抑制: LockService で直列化 + Properties で再送防止
    const lock = LockService.getScriptLock();
    try {
      // ロック取得（競合する実行を直列化。取れなければメール送信せず終了）
      if (!lock.tryLock(10000)) throw { _skip: true };

      const props    = PropertiesService.getScriptProperties();
      const propKey   = 'MISSING_ALERT_' + missing.join('|');
      const last      = Number(props.getProperty(propKey) || 0);
      const nowMs     = Date.now();
      // 30分（1800000ms）以内に送信済みならスキップ
      if (nowMs - last < 1800000) throw { _skip: true };
      props.setProperty(propKey, String(nowMs));

      let owner = '';
      try { owner = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail(); } catch (_) {}
      if (!owner) owner = Session.getEffectiveUser().getEmail();

      const missingHtml = missing
        .map(m => `<li style="white-space:nowrap;">${m}</li>`).join('');

      // 各種リンク取得
      let webUrl = '';
      try { webUrl = ScriptApp.getService().getUrl(); } catch (_) {}
      let mainUrl = '';
      try { mainUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl(); } catch (_) {}
      const kanriUrl = `https://docs.google.com/spreadsheets/d/${KANRI_SS_ID}/edit`;

      const linksHtml = [
        webUrl  ? `<li><a href="${webUrl}">協賛申込みフォーム（Web App）</a></li>` : '',
        mainUrl ? `<li><a href="${mainUrl}">メインスプレッドシート</a></li>`        : '',
        `<li><a href="${kanriUrl}">管理IDリスト スプレッドシート</a></li>`,
      ].join('');

      const htmlBody = `
        <div style="font-family:sans-serif;font-size:14px;color:#333;line-height:1.7;">
          <p>協賛申込みフォームにアクセスがありましたが、以下の必須シートが存在しないため
          エラーページを表示しました。</p>
          <p><b>■ 不足しているシート:</b></p>
          <ul style="margin:4px 0;">${missingHtml}</ul>
          <p>　アクセス日時：${nowStr()}</p>
          <hr>
          <p><b>■ 関連リンク:</b></p>
          <ul style="margin:4px 0;">${linksHtml}</ul>
          <hr>
          <p><b>【対応をお願いします】</b></p>
          <p>① シートが誰かに削除された可能性があります。内容をご確認ください。<br>
          ② 初めてお使いの場合は初期設定が必要です。<br>
          　 カスタムメニュー「⚙️ 初期設定」から以下を順番に実行してください:</p>
          <ol style="margin:4px 0;">
            <li>Info シート作成 → 必要な値を手動記入</li>
            <li>管理ID番号シート生成</li>
            <li>プロジェクト初期化</li>
          </ol>
          <hr>
        </div>`;

      MailApp.sendEmail({
        to:       owner,
        subject:  '【⚠️ 緊急】必須シートが見つかりません（フォームアクセス時）',
        body:     '※ HTMLメール対応の環境でご覧ください。\n\n' +
                  '不足シート:\n' + missing.map(m => `・${m}`).join('\n') + '\n\n' +
                  '関連リンク:\n' +
                  (webUrl  ? `・フォーム: ${webUrl}\n`  : '') +
                  (mainUrl ? `・メイン: ${mainUrl}\n`   : '') +
                  `・管理IDリスト: ${kanriUrl}`,
        htmlBody: htmlBody,
      });
    } catch (e) {
      if (!e._skip) console.error('警告メール送信失敗:', e.message);
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }

    // エラーページを表示
    return HtmlService.createHtmlOutput(`
      <html><body style="font-family:sans-serif;text-align:center;padding:60px;color:#c0392b;">
        <h2>⚠️ システムエラー</h2>
        <p>現在フォームをご利用いただけません。<br>管理者にお問い合わせください。</p>
      </body></html>
    `).setTitle('エラー').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(`【${getEventName()}】協賛申込み`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─── メール送信残数 関連ヘルパー ──────────────────────────────────────────────

/** メール残数を安全に取得（取得失敗時は 0 = 不足扱い） */
function _getMailQuotaSafe() {
  try {
    return MailApp.getRemainingDailyQuota();
  } catch (e) {
    console.error('メール残数取得失敗:', e.message);
    return 0;
  }
}

/**
 * メール残数不足をオーナーへ通知する（30分に1通だけ・LockServiceで直列化）
 * ※ 残数が枯渇していると通知メール自体も送れない可能性があるため try で保護
 */
function _notifyLowQuota(quota) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(5000)) return;

    const props  = PropertiesService.getScriptProperties();
    const propKey = 'LOW_QUOTA_ALERT';
    const last    = Number(props.getProperty(propKey) || 0);
    const nowMs   = Date.now();
    if (nowMs - last < 1800000) return; // 30分以内は再送しない
    props.setProperty(propKey, String(nowMs));

    let owner = '';
    try { owner = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail(); } catch (_) {}
    if (!owner) owner = Session.getEffectiveUser().getEmail();

    let webUrl = '';
    try { webUrl = ScriptApp.getService().getUrl(); } catch (_) {}

    MailApp.sendEmail({
      to:      owner,
      subject: '【⚠️ 緊急】メール送信残数が不足しています（フォーム停止中）',
      body: [
        'メール送信の1日あたり残数が最低ラインを下回ったため、',
        '協賛申込みフォームを一時停止しました。',
        '',
        `　現在の残数　：${quota} 通`,
        `　最低ライン　：${getMinMailQuota()} 通（Info の MIN_MAIL_QUOTA）`,
        `　検知日時　　：${nowStr()}`,
        '',
        '━━━━━━━━━━━━━━━━━━━━━━━━',
        '【ご確認ください】',
        '・残数は太平洋時間の深夜に自動リセットされます（約1日）。',
        '・リセット後は自動的にフォーム受付が再開されます。',
        '・上限を増やすには Google Workspace アカウントが必要です。',
        webUrl ? `\n■ フォームURL: ${webUrl}` : '',
        '━━━━━━━━━━━━━━━━━━━━━━━━',
      ].join('\n'),
    });
  } catch (e) {
    console.error('残数不足通知の送信失敗:', e.message);
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/**
 * クライアント(google.script.run)から呼ばれる設定取得
 */
function getConfig() {
  // メール残数チェック（不足時はオーナーへ通知）
  const quota      = _getMailQuotaSafe();
  const mailQuotaOk = quota >= getMinMailQuota();
  if (!mailQuotaOk) _notifyLowQuota(quota);

  return {
    startDate:  String(getConfigVal('START_DATE',  '2025-01-01T00:00:00')),
    endDate:    String(getConfigVal('END_DATE',     '2026-10-01T23:59:59')),
    sheetName1: DEFAULT_SHEET_NAME,
    sheetName2: DEFAULT_SHEET_NAME2,
    eventName:        getEventName(),
    receiptNoPrefix:  getReceiptNoPrefix(),
    mailQuotaOk:      mailQuotaOk,
  };
}

/**
 * クライアント(google.script.run)から呼ばれるフォーム送信
 */
function submitForm(data) {
  // ─── ① メール送信残数チェック（最優先・データ記録より前） ────────────────────
  // 残数不足のままデータを記録すると「申込みは登録されたがメール未送信」になるため、
  // 記録前に拒否する。
  const _quota = _getMailQuotaSafe();
  if (_quota < getMinMailQuota()) {
    _notifyLowQuota(_quota);
    throw new Error('ただいまお申し込みを受け付けできません。時間をおいて再度お試しください。');
  }

  // ─── Info シート 存在確認（必須） ────────────────────────────────────────────
  const _infoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INFO_SHEET_NAME);
  if (!_infoSheet) {
    try {
      const _owner = Session.getEffectiveUser().getEmail();
      MailApp.sendEmail({
        to:      _owner,
        subject: '【⚠️ 緊急】Info シートが見つかりません',
        body: [
          'Info シートが存在しないため、以下のユーザーの申込みが失敗しました。',
          '',
          `　会社名　：${data.company_name || '（不明）'}`,
          `　メール　：${data.email || '（不明）'}`,
          `　申込日時：${nowStr()}`,
          '',
          '━━━━━━━━━━━━━━━━━━━━━━━━',
          '【対応をお願いします】',
          '① Info シートが削除された可能性があります。',
          '   → メインスプレッドシートを開き「Info」シートを確認してください。',
          '',
          '② 初めてお使いの場合は初期設定が必要です。',
          '   → カスタムメニュー「⚙️ 初期設定 → Info シート作成」を実行してください。',
          '━━━━━━━━━━━━━━━━━━━━━━━━',
        ].join('\n'),
      });
    } catch (mailErr) {
      console.error('Info警告メール送信失敗:', mailErr.message);
    }
    throw new Error('システム設定が見つかりません。管理者に通知しました。');
  }

  // ─── サーバー側バリデーション ────────────────────────────────────────────────
  const zipcode = String(data.zipcode || '').trim().replace(/\D/g, '');
  const phone   = String(data.phone   || '').trim().replace(/\D/g, '');
  if (!/^\d{7}$/.test(zipcode)) {
    throw new Error('郵便番号は7桁の数字で入力してください。（入力値: ' + (data.zipcode || '') + '）');
  }
  if (!/^\d{9,11}$/.test(phone)) {
    throw new Error('電話番号は9〜11桁の数字で入力してください。（入力値: ' + (data.phone || '') + '）');
  }
  // バリデーション済みの値で上書き
  data.zipcode = zipcode;
  data.phone   = phone;

  const kubun    = (data.category || '').trim().toUpperCase();
  const autoSend = AUTO_SEND_KUBUN.includes(kubun);

  // ─── 管理IDリスト 存在確認（必須） ──────────────────────────────────────────
  if (!getKanriSheet()) {
    // 事務局へ警告メール送信
    try {
      MailApp.sendEmail({
        to:      getOfficeEmail(),
        subject: `【⚠️ 緊急】管理IDリストシートが見つかりません`,
        body: [
          '管理IDリストシートにアクセスできないため、以下のユーザーの申込みが失敗しました。',
          '',
          `　会社名　：${data.company_name || '（不明）'}`,
          `　メール　：${data.email || '（不明）'}`,
          `　申込日時：${nowStr()}`,
          '',
          '━━━━━━━━━━━━━━━━━━━━━━━━',
          '【対応をお願いします】',
          `　管理IDリスト ファイルID：${KANRI_SS_ID}`,
          '',
          '① シートが誰かに削除された可能性があります。',
          '   → 上記ファイルIDのスプレッドシートを開き「管理IDリスト」シートを確認してください。',
          '',
          '② 初めてお使いの場合は管理IDリストの初期化が必要です。',
          '   → GASエディタから「initKanriSheet」を実行してシートを作成してください。',
          '━━━━━━━━━━━━━━━━━━━━━━━━',
        ].join('\n'),
      });
    } catch (mailErr) {
      console.error('警告メール送信失敗:', mailErr.message);
    }
    throw new Error('管理IDリストシートにアクセスできません。管理者に通知しました。');
  }

  // ─── 管理ID処理 ─────────────────────────────────────────────────────────────
  let kanriId = String(data.kanri_id || '').trim();

  let isNewKanri = false;
  if (kanriId) {
    // 既存IDの場合 → 変更チェック → update_kanri フラグがあれば更新
    const doUpdate = data.update_kanri === 'yes';
    saveKanriId(kanriId, data, doUpdate);
  } else {
    // 新規ユーザー → 新ID発行 → 保存
    kanriId    = generateKanriId();
    saveKanriId(kanriId, data, true);
    isNewKanri = true;
  }
  data.kanri_id = kanriId;

  // ─── スプレッドシート登録 ─────────────────────────────────────────────────
  const receptNo = appendRow(data, DEFAULT_SHEET_NAME);
  appendToTesagyouSheet(receptNo, DEFAULT_SHEET_NAME2, data);

  if (data.email) {
    // 新規ユーザーには管理ID発行通知メールを先に送信
    if (isNewKanri) sendKanriIdIssuedEmail(data, kanriId);

    if (autoSend) {
      const pdf = generateInvoicePdf(data, receptNo);
      sendConfirmationEmail(data, receptNo, pdf, kanriId);
    } else {
      sendReceiptOnlyEmail(data, receptNo, kanriId);
    }
    _notifyOffice(data, receptNo, autoSend, kanriId);
  }

  return { result: 'success', receipt_no: receptNo, kanri_id: kanriId };
}

/** カスタムメニュー登録 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 協賛管理')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⚙️ 初期設定')
      .addItem('プロジェクト初期化',      'initProject')
      .addItem('Info シート作成',         'setupInfoSheet')
      .addItem('管理ID番号シート生成',     'initKanriSheet')
      .addItem('トリガー再登録',          'registerTriggers')
    )
    .addSeparator()
    .addItem('メールテンプレ保存', 'setupAllMailTemplates')
    .addToUi();
}

// ─── 内部ユーティリティ ────────────────────────────────────────────────────────

function _notifyOffice(data, receptNo, autoSent, kanriId) {
  const subject = `【協賛申込】${data.company_name || ''} (${data.category || ''}) 受付番号:${receptNo}`;
  const body = [
    '新規協賛申込みがありました。',
    '',
    `会社名　：${data.company_name || ''}`,
    `区分　　：${data.category || ''} プラン`,
    `担当者　：${data.staff_name || ''}`,
    `メール　：${data.email || ''}`,
    `受付番号：${receptNo}`,
    `管理ID　：${kanriId || ''}`,
    `請求書　：${autoSent ? '自動送信済み' : '手動送信待ち（S/Aプラン）'}`,
  ].join('\n');
  try {
    MailApp.sendEmail({ to: getOfficeEmail(), subject, body });
  } catch (e) {
    console.warn('事務局通知メール送信失敗:', e.message);
  }
}
