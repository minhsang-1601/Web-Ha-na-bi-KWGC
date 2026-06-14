// ─── Web App エントリーポイント ────────────────────────────────────────────────

function doGet() {
  // ─── 必須シート存在チェック（Info） ──────────────────────────
  const infoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INFO_SHEET_NAME);
  if (!infoSheet) {
    const lock = LockService.getScriptLock();
    try {
      if (!lock.tryLock(10000)) throw { _skip: true };
      const props  = PropertiesService.getScriptProperties();
      const propKey = 'MISSING_ALERT_INFO';
      const last   = Number(props.getProperty(propKey) || 0);
      const nowMs  = Date.now();
      if (nowMs - last < 1800000) throw { _skip: true };
      props.setProperty(propKey, String(nowMs));

      let owner = '';
      try { owner = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail(); } catch (_) {}
      if (!owner) owner = Session.getEffectiveUser().getEmail();

      let webUrl = '';
      try { webUrl = ScriptApp.getService().getUrl(); } catch (_) {}
      let mainUrl = '';
      try { mainUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl(); } catch (_) {}

      MailApp.sendEmail({
        to:      owner,
        subject: '【⚠️ 緊急】Info シートが見つかりません（フォームアクセス時）',
        body: [
          '協賛申込みフォームにアクセスがありましたが、Info シートが存在しないためエラーページを表示しました。',
          '',
          `アクセス日時：${nowStr()}`,
          webUrl  ? `フォーム: ${webUrl}`  : '',
          mainUrl ? `メイン: ${mainUrl}` : '',
        ].filter(Boolean).join('\n'),
      });
    } catch (e) {
      if (!e._skip) console.error('警告メール送信失敗:', e.message);
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }

    return HtmlService.createHtmlOutput(`
      <html><body style="font-family:sans-serif;text-align:center;padding:60px;color:#c0392b;">
        <h2>⚠️ システムエラー</h2>
        <p>現在フォームをご利用いただけません。<br>管理者にお問い合わせください。</p>
      </body></html>
    `).setTitle('エラー').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ─── 必須値の未入力チェック ───────────────────────────────────────────────────
  const missingKeys = _checkRequiredInfoKeys();
  if (missingKeys.length > 0) {
    _sendInfoIncompleteAlert(missingKeys);
    return HtmlService.createHtmlOutput(`
      <html><body style="font-family:sans-serif;text-align:center;padding:60px;color:#c0392b;">
        <h2>⚠️ システムエラー</h2>
        <p>設定が完了していないため、フォームを表示できません。<br>管理者にお問い合わせください。</p>
      </body></html>
    `).setTitle('エラー').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(`【${getEventName()}】協賛申込み`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 必須 Info キーが空でないか確認し、未入力のキー一覧を返す */
function _checkRequiredInfoKeys() {
  const REQUIRED = [
    'EVENT_NAME',
    'OFFICE_EMAIL',
    'START_DATE',
    'END_DATE',
    'PAYMENT_DUE',
    'ORG_NAME',
    'ORG_REP',
    'BANK_NAME',
    'BANK_NO',
    'BANK_HOLDER',
  ];
  const missing = [];
  REQUIRED.forEach(key => {
    const v = String(getConfigVal(key, '') || '').trim();
    if (!v) missing.push(key);
  });
  return missing;
}

/** Info 未入力をオーナーへ通知（30分に1通） */
function _sendInfoIncompleteAlert(missingKeys) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(5000)) return;
    const props  = PropertiesService.getScriptProperties();
    const propKey = 'INFO_INCOMPLETE_ALERT';
    const last   = Number(props.getProperty(propKey) || 0);
    const nowMs  = Date.now();
    if (nowMs - last < 1800000) return;
    props.setProperty(propKey, String(nowMs));

    let owner = '';
    try { owner = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail(); } catch (_) {}
    if (!owner) owner = Session.getEffectiveUser().getEmail();

    let webUrl = '';
    try { webUrl = ScriptApp.getService().getUrl(); } catch (_) {}
    let mainUrl = '';
    try { mainUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl(); } catch (_) {}

    MailApp.sendEmail({
      to:      owner,
      subject: '【⚠️ 設定未完了】Info シートの必須項目が未入力です（フォームアクセス時）',
      body: [
        '協賛申込みフォームにアクセスがありましたが、Info シートの必須項目が未入力のためエラーページを表示しました。',
        '',
        '■ 未入力の項目:',
        ...missingKeys.map(k => `　・${k}`),
        '',
        `アクセス日時：${nowStr()}`,
        webUrl  ? `フォームURL：${webUrl}`  : '',
        mainUrl ? `スプレッドシート：${mainUrl}` : '',
        '',
        '【対応】Info シートを開き、上記の項目に値を入力してください。',
      ].filter(Boolean).join('\n'),
    });
  } catch (e) {
    console.error('Info未入力通知メール送信失敗:', e.message);
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
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
    kubunSaStart:   String(getKubunSaStart()),
    kubunSaEnd:     String(getKubunSaEnd()),
    kubunBcdeStart: String(getKubunBcdeStart()),
    kubunBcdeEnd:   String(getKubunBcdeEnd()),
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

  // ─── スプレッドシート登録 ─────────────────────────────────────────────────
  const receptNo = appendRow(data, DEFAULT_SHEET_NAME);
  appendToTesagyouSheet(receptNo, DEFAULT_SHEET_NAME2, data);

  if (data.email) {
    if (autoSend) {
      const pdf = generateInvoicePdf(data, receptNo);
      sendConfirmationEmail(data, receptNo, pdf);
    } else {
      sendReceiptOnlyEmail(data, receptNo);
    }
    // _notifyOffice(data, receptNo, autoSend);  // 事務局通知メールを無効化
  }

  return { result: 'success', receipt_no: receptNo };
}

/** カスタムメニュー登録 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 協賛管理')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⚙️ 初期設定')
      .addItem('プロジェクト初期化',      'initProject')
      .addItem('Info シート作成',         'setupInfoSheet')
      .addItem('トリガー再登録',          'registerTriggers')
    )
    .addSeparator()
    .addItem('メールテンプレ保存', 'setupAllMailTemplates')
    .addItem('列名を最新に更新',   'renameHeaders')
    .addToUi();
}

// ─── 内部ユーティリティ ────────────────────────────────────────────────────────

function _notifyOffice(data, receptNo, autoSent) {
  const subject = `【協賛申込】${data.company_name || ''} (${data.category || ''}) 受付番号:${receptNo}`;
  const body = [
    '新規協賛申込みがありました。',
    '',
    `会社名　：${data.company_name || ''}`,
    `区分　　：${data.category || ''}`,
    `担当者　：${data.staff_name || ''}`,
    `メール　：${data.email || ''}`,
    `受付番号：${receptNo}`,
    `請求書　：${autoSent ? '自動送信済み' : '手動送信待ち（S/A）'}`,
  ].join('\n');
  try {
    MailApp.sendEmail({ to: getOfficeEmail(), subject, body });
  } catch (e) {
    console.warn('事務局通知メール送信失敗:', e.message);
  }
}
