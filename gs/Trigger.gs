// ─── チェックボックストリガー（手作業シート） ──────────────────────────────────
// ⚠️ GAS エディタ「トリガー」→ onEditInstallable → スプレッドシート → 編集時 で登録
//
// 手作業 列構成（I〜L）:
//   I (9)  受付完了         checkbox  手動  → 請求書送信確認ダイアログ
//   J (10) 請求書送信日時   timestamp 自動  ← _doSendInvoice でセット
//   K (11) 入金完了         checkbox  手動  → お礼状送信確認 ※J が必須
//   L (12) お礼状送信日時   timestamp 自動  ← _doSendNyukin でセット

function onEditInstallable(e) {
  // すべての編集を操作ログに記録（誰がどのセルを変更/削除したか）
  _logAudit(e);

  const sheet = e.range.getSheet();
  if (sheet.getName() !== DEFAULT_SHEET_NAME2) return;
  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row <= 2) return; // ヘッダー行・サブヘッダー行をスキップ

  if (col === COL_UKETSUKE)  handleUketsuke(e, sheet, row);
  if (col === COL_NYUKIN)    handleNyukin(e, sheet, row);
}

// ─── I列: 受付完了 → 請求書送信 ────────────────────────────────────────────────

function handleUketsuke(e, sheet, row) {
  const invDateCell = sheet.getRange(row, COL_INV_DATE);
  const invDate     = invDateCell.getValue();

  if (e.range.getValue() === false) {
    // ── クリア順序チェック: 下流（K→L）を先に解除しないと外せない ──
    const oreijouDate = sheet.getRange(row, COL_OREIJOU_DATE).getValue();
    const nyukin      = sheet.getRange(row, COL_NYUKIN).getValue();

    // I を外す → 全下流（K・L）を自動クリア
    const hasDownstream = nyukin || oreijouDate || invDate;
    if (hasDownstream) {
      const detail = [
        oreijouDate ? `　お礼状送信日時（L）：${formatTs(oreijouDate)}` : null,
        nyukin      ? `　入金完了（K）：チェック済み`                    : null,
        invDate     ? `　請求書送信日時（J）：${formatTs(invDate)}`      : null,
      ].filter(Boolean).join('\n');
      const res = SpreadsheetApp.getUi().alert(
        '⚠️ チェックを外しますか？',
        `以下の値が自動でクリアされます。\n\n${detail}\n\nよろしいですか？`,
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      if (res === SpreadsheetApp.getUi().Button.YES) {
        const nyukinCell = sheet.getRange(row, COL_NYUKIN);
        nyukinCell.setValue(false);
        sheet.getRange(row, COL_OREIJOU_DATE).clearContent();
        invDateCell.clearContent();
        const user = _auditUser(e);
        _logAction(e.source, user, sheet.getName(), e.range.getA1Notation(), '受付完了 解除', '操作成功');
        if (nyukin) {
          _logAction(e.source, user, sheet.getName(), nyukinCell.getA1Notation(), '入金完了 解除（連動）', '操作成功');
        }
      } else {
        e.range.setValue(true);
      }
    } else {
      _logAction(e.source, _auditUser(e), sheet.getName(), e.range.getA1Notation(), '受付完了 解除', '操作成功');
    }
    return;
  }

  if (invDate) {
    _logAction(e.source, _auditUser(e), sheet.getName(), e.range.getA1Notation(), '請求書送信（送信済み）', '操作失敗');
    SpreadsheetApp.getUi().alert(`⚠️ 請求書はすでに送信済みです。\n送信日時：${formatTs(invDate)}`);
    e.range.setValue(false);
    return;
  }

  // メール残数チェック
  if (_blockSendIfLowQuota(e.range)) {
    _logAction(e.source, _auditUser(e), sheet.getName(), e.range.getA1Notation(), '請求書送信（残数不足）', '操作失敗');
    return;
  }

  const receptNo  = sheet.getRange(row, COL_RECEPT_NO).getValue();
  const mainSheet = e.source.getSheetByName(DEFAULT_SHEET_NAME);
  if (!mainSheet) {
    _logAction(e.source, _auditUser(e), sheet.getName(), e.range.getA1Notation(), '請求書送信（申込みシート無）', '操作失敗');
    SpreadsheetApp.getUi().alert('申込みシートが見つかりません。');
    e.range.setValue(false);
    return;
  }

  const data = findRowByReceptNo(mainSheet, receptNo);
  if (!data || !data.email) {
    _logAction(e.source, _auditUser(e), sheet.getName(), e.range.getA1Notation(), '請求書送信（メール無）', '操作失敗');
    SpreadsheetApp.getUi().alert('メールアドレスが見つかりません。');
    e.range.setValue(false);
    return;
  }

  // ── 確認 → 送信（トリガー＝デプロイ者権限で実行されるため FROM はデプロイ者） ──
  const ui  = SpreadsheetApp.getUi();
  const res = ui.alert(
    '請求書送信の確認',
    `以下の宛先に「申込受理書兼請求書」を送信します。よろしいですか？\n\n` +
    `　受付番号：${receptNo}\n　会社名　：${data.company_name || ''}\n　送信先　：${data.email}`,
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) { e.range.setValue(false); return; }

  const a1 = e.range.getA1Notation();
  const r = _doSendInvoice(e.source, row, receptNo);
  if (!r.ok) {
    _logAction(e.source, _auditUser(e), sheet.getName(), a1, '請求書送信', '操作失敗');
    ui.alert('❌ 送信に失敗しました。\n' + (r.error || ''));
    e.range.setValue(false);
    return;
  }
  _logAction(e.source, _auditUser(e), sheet.getName(), a1, '請求書送信', '操作成功');
  ui.alert(`✅ 請求書を送信しました。\n　送信先：${data.email}`);
}

// ─── K列: 入金完了 → お礼状送信 ──────────────────────────────────────────────────

function handleNyukin(e, sheet, row) {
  const oreijouDateCell = sheet.getRange(row, COL_OREIJOU_DATE);
  const oreijouDate     = oreijouDateCell.getValue();

  if (e.range.getValue() === false) {
    // K を外す → L を自動クリア（値がある場合は確認）
    if (oreijouDate) {
      const detail = `　お礼状送信日時（L）：${formatTs(oreijouDate)}`;
      const res = SpreadsheetApp.getUi().alert(
        '⚠️ チェックを外しますか？',
        `以下の値が自動でクリアされます。\n\n${detail}\n\nよろしいですか？`,
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      if (res === SpreadsheetApp.getUi().Button.YES) {
        oreijouDateCell.clearContent();
        _logAction(e.source, _auditUser(e), sheet.getName(), e.range.getA1Notation(), '入金完了 解除', '操作成功');
      } else {
        e.range.setValue(true);
      }
    } else {
      _logAction(e.source, _auditUser(e), sheet.getName(), e.range.getA1Notation(), '入金完了 解除', '操作成功');
    }
    return;
  }

  // 前提条件: 請求書送信済み（J に日時あり）
  const invDate = sheet.getRange(row, COL_INV_DATE).getValue();
  if (!invDate) {
    _logAction(e.source, _auditUser(e), sheet.getName(), e.range.getA1Notation(), 'お礼状送信（請求書未送信）', '操作失敗');
    SpreadsheetApp.getUi().alert(
      '⚠️ 請求書がまだ送信されていません。\n先に「受付完了」をチェックして請求書を送信してください。'
    );
    e.range.setValue(false);
    return;
  }

  if (oreijouDate) {
    _logAction(e.source, _auditUser(e), sheet.getName(), e.range.getA1Notation(), 'お礼状送信（送信済み）', '操作失敗');
    SpreadsheetApp.getUi().alert(`⚠️ お礼状はすでに送信済みです。\n送信日時：${formatTs(oreijouDate)}`);
    e.range.setValue(false);
    return;
  }

  // メール残数チェック
  if (_blockSendIfLowQuota(e.range)) {
    _logAction(e.source, _auditUser(e), sheet.getName(), e.range.getA1Notation(), 'お礼状送信（残数不足）', '操作失敗');
    return;
  }

  const receptNo  = sheet.getRange(row, COL_RECEPT_NO).getValue();
  const mainSheet = e.source.getSheetByName(DEFAULT_SHEET_NAME);
  const data      = findRowByReceptNo(mainSheet, receptNo);
  if (!data || !data.email) {
    _logAction(e.source, _auditUser(e), sheet.getName(), e.range.getA1Notation(), 'お礼状送信（メール無）', '操作失敗');
    SpreadsheetApp.getUi().alert('メールアドレスが見つかりません。');
    e.range.setValue(false);
    return;
  }

  // ── 確認 → 送信（トリガー＝デプロイ者権限で実行されるため FROM はデプロイ者） ──
  const ui  = SpreadsheetApp.getUi();
  const res = ui.alert(
    '入金確認・お礼状送信',
    `以下の宛先に「お礼状」を送信します。よろしいですか？\n\n` +
    `　受付番号：${receptNo}\n　会社名　：${data.company_name || ''}\n　送信先　：${data.email}`,
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) { e.range.setValue(false); return; }

  const a1 = e.range.getA1Notation();
  const r = _doSendNyukin(e.source, row, receptNo);
  if (!r.ok) {
    _logAction(e.source, _auditUser(e), sheet.getName(), a1, 'お礼状送信', '操作失敗');
    ui.alert('❌ 送信に失敗しました。\n' + (r.error || ''));
    e.range.setValue(false);
    return;
  }
  _logAction(e.source, _auditUser(e), sheet.getName(), a1, 'お礼状送信', '操作成功');
  ui.alert(`✅ お礼状を送信しました。\n　送信先：${data.email}`);
}



// ─── 操作ログ（監査ログ） ───────────────────────────────────────────────────────
// 誰が・いつ・どのシートの・どのセルを・どう変更/削除したかを記録する。
// onEditInstallable（セル編集）と onChangeInstallable（行列の挿入/削除）から呼ぶ。

/** ログ用シートを取得（なければ作成してヘッダーを付与） */
function _getAuditSheet(ss) {
  let sheet = ss.getSheetByName(AUDIT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(AUDIT_SHEET);
    sheet.getRange(1, 1, 1, AUDIT_HEADERS.length).setValues([AUDIT_HEADERS]);
    sheet.getRange(1, 1, 1, AUDIT_HEADERS.length).setFontWeight('bold').setBackground('#e8f4f8');
    [150, 220, 100, 70, 90, 90].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** カスタム操作（送信など）を1行記録する。status = '操作成功' / '操作失敗' */
function _logAction(ss, user, sheetName, a1, op, status) {
  try {
    const log = _getAuditSheet(ss);
    log.appendRow([
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
      user, sheetName, a1, op, status || '',
    ]);
  } catch (err) {
    console.error('操作ログ記録エラー(action):', err.message);
  }
}

/** セル編集を記録（onEdit イベント） */
function _logAudit(e) {
  try {
    if (!e || !e.range) return;
    const editedSheet = e.range.getSheet();
    if (editedSheet.getName() === AUDIT_SHEET) return; // ログ自身は記録しない

    // 手作業の I列(受付完了)・K列(入金完了) は「送信操作」のため、ここでは記録しない。
    // 送信が成功したときだけ handleUketsuke/handleNyukin から _logAction で記録する
    // （失敗した試行はログに残さない）。
    const col = e.range.getColumn();
    const single = e.range.getNumRows() * e.range.getNumColumns() === 1;
    if (single && editedSheet.getName() === DEFAULT_SHEET_NAME2 &&
        (col === COL_UKETSUKE || col === COL_NYUKIN)) {
      return;
    }

    const ss   = e.source || SpreadsheetApp.getActiveSpreadsheet();
    const log  = _getAuditSheet(ss);
    const user = _auditUser(e);
    const cells = e.range.getNumRows() * e.range.getNumColumns();

    let before, after, op;
    if (cells === 1) {
      before = (e.oldValue !== undefined && e.oldValue !== null) ? e.oldValue : '';
      after  = (e.value    !== undefined && e.value    !== null) ? e.value    : '';
      op = (after === '' && before !== '') ? '削除' : (before === '' ? '入力' : '変更');
    } else {
      before = '（複数セル）';
      after  = '（複数セル）';
      op = '一括変更';
    }

    log.appendRow([
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
      user,
      editedSheet.getName(),
      e.range.getA1Notation(),
      op,
      '操作成功',
    ]);
  } catch (err) {
    console.error('操作ログ記録エラー(edit):', err.message);
  }
}

/** 行・列の挿入/削除などの構造変更を記録（onChange イベント） */
function onChangeInstallable(e) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const log = _getAuditSheet(ss);

    // 構造変更（行・列・シートの挿入/削除）のみ記録する。
    // EDIT/FORMAT/OTHER 等はノイズになるため記録しない（セル編集は onEdit 側で記録）。
    const TYPE_LABEL = {
      INSERT_ROW:    '行を挿入', REMOVE_ROW:    '行を削除',
      INSERT_COLUMN: '列を挿入', REMOVE_COLUMN: '列を削除',
      INSERT_GRID:   'シート追加', REMOVE_GRID:  'シート削除',
    };
    const changeType = (e && e.changeType) || '';
    if (!TYPE_LABEL[changeType]) return;

    let sheetName = '', a1 = '';
    try {
      const rng = ss.getActiveRange();
      if (rng) { sheetName = rng.getSheet().getName(); a1 = rng.getA1Notation(); }
    } catch (_) {}
    if (sheetName === AUDIT_SHEET) return;

    log.appendRow([
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
      _auditUser(e),
      sheetName,
      a1,
      TYPE_LABEL[changeType] || changeType,
      '操作成功',
    ]);
  } catch (err) {
    console.error('操作ログ記録エラー(change):', err.message);
  }
}

/** 操作者のメールアドレスを取得（取得できなければ unknown） */
function _auditUser(e) {
  try {
    if (e && e.user && e.user.getEmail && e.user.getEmail()) return e.user.getEmail();
  } catch (_) {}
  try {
    const u = Session.getActiveUser().getEmail();
    if (u) return u;
  } catch (_) {}
  return '(unknown)';
}

/** データ用スプレッドシートにも同じメニューを表示（onOpen trigger として登録） */
function onOpenEventSheet() {
  SpreadsheetApp.getUi()
    .createMenu('📋 協賛管理')
    .addToUi();
}

// ─── 送信実処理（installable トリガー＝デプロイ者権限で実行される） ───────────────
// handleUketsuke / handleNyukin（onEditInstallable 経由）から呼ばれる。
// トリガーはトリガー作成者（デプロイ者）の権限で動くため、誰が操作しても
// MailApp の FROM はデプロイ者に固定される。UrlFetch 等の追加権限は不要。

/** I列チェック → 請求書（B〜E）/ 抽選確定請求書（S・A）を送信 */
function _doSendInvoice(ss, row, receptNo) {
  const dataSs     = ss || getDataSpreadsheet();
  const tetsuSheet = dataSs.getSheetByName(DEFAULT_SHEET_NAME2);
  const mainSheet  = dataSs.getSheetByName(DEFAULT_SHEET_NAME);

  // 引数が失われた場合: rowで手作業シートから直接読む
  let rowNum = Number(row) || 0;
  if ((!receptNo || receptNo === 'undefined') && tetsuSheet && rowNum > 2) {
    receptNo = String(tetsuSheet.getRange(rowNum, COL_RECEPT_NO).getValue()).trim();
  }
  // それでも取得できない場合: I列チェック済み・J列空のrowをスキャン
  if (!receptNo && tetsuSheet) {
    const vals = tetsuSheet.getDataRange().getValues();
    for (let i = 2; i < vals.length; i++) {
      if (vals[i][COL_UKETSUKE - 1] === true && !vals[i][COL_INV_DATE - 1]) {
        receptNo = String(vals[i][COL_RECEPT_NO - 1]).trim();
        rowNum   = i + 1;
        break;
      }
    }
  }

  const data = findRowByReceptNo(mainSheet, receptNo);
  if (!data) return { ok: false, error: '受付番号が見つかりません: ' + receptNo };

  const pdf = generateInvoicePdf(data, receptNo);

  // 区分を取得して、S/A と B~E で異なるテンプレートで送信
  const kubun = String(tetsuSheet.getRange(rowNum, 2).getValue()).trim().toUpperCase();
  if (['S', 'A'].includes(kubun)) {
    sendSaInvoiceEmail(data, receptNo, pdf);
  } else {
    sendConfirmationEmail(data, receptNo, pdf);
  }

  tetsuSheet.getRange(rowNum, COL_INV_DATE).setValue(nowStr());
  return { ok: true };
}

/** K列チェック → お礼状を送信 */
function _doSendNyukin(ss, row, receptNo) {
  const dataSs     = ss || getDataSpreadsheet();
  const tetsuSheet = dataSs.getSheetByName(DEFAULT_SHEET_NAME2);
  const mainSheet  = dataSs.getSheetByName(DEFAULT_SHEET_NAME);
  const data = findRowByReceptNo(mainSheet, receptNo);
  if (!data) return { ok: false, error: '受付番号が見つかりません: ' + receptNo };

  try {
    sendOreijouEmail(data, receptNo);
  } catch (e) {
    return { ok: false, error: e.message };
  }

  tetsuSheet.getRange(Number(row), COL_OREIJOU_DATE).setValue(nowStr());
  return { ok: true };
}

// ─── メール残数チェック（手作業シートの送信操作用） ──────────────────────────────

/**
 * メール残数が最低ラインを下回っていれば警告して送信をブロックする
 * @param {Object} revertCell - 残数不足時に false に戻すチェックボックスセル（任意）
 * @returns {boolean} true = ブロックした（送信中止）, false = 続行可
 */
function _blockSendIfLowQuota(revertCell) {
  let quota = 0;
  try { quota = MailApp.getRemainingDailyQuota(); } catch (_) { quota = 0; }

  let minQuota = 5;
  try { minQuota = getMinMailQuota(); } catch (_) {}

  if (quota < minQuota) {
    try { _notifyLowQuota(quota); } catch (_) {}
    SpreadsheetApp.getUi().alert(
      '⚠️ 送信できません',
      `本日のメール送信残数が不足しています。\n\n` +
      `　現在の残数：${quota} 通\n` +
      `　最低ライン：${minQuota} 通\n\n` +
      `残数は約1日で自動リセットされます。\n時間をおいて再度お試しください。`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    if (revertCell) {
      try { revertCell.setValue(false); } catch (_) {}
    }
    return true;
  }
  return false;
}

// ─── ユーティリティ ─────────────────────────────────────────────────────────────

function findRowByReceptNo(sheet, receptNo) {
  // 協賛申込み一覧: A=受付番号 B=受付日時 C=会社名 D=会社名ふりがな
  //                E=代表者役職 F=役職ふりがな G=代表者名 H=代表者名ふりがな
  //                I=担当者名 J=担当者ふりがな K=郵便番号 L=住所 M=電話番号 N=メール O=区分 P=会社HP URL
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(receptNo).trim()) { // A列: 受付番号
      return {
        company_name:          rows[i][2],   // C
        company_furigana:      rows[i][3],   // D
        rep_position:          rows[i][4],   // E
        rep_position_furigana: rows[i][5],   // F
        rep_name:              rows[i][6],   // G
        rep_furigana:          rows[i][7],   // H
        staff_name:            rows[i][8],   // I
        staff_furigana:        rows[i][9],   // J
        zipcode:               rows[i][10],  // K
        address:               rows[i][11],  // L
        phone:                 rows[i][12],  // M
        email:                 rows[i][13],  // N
        category:              rows[i][14],  // O
        website_url:           rows[i][15],  // P
      };
    }
  }
  return null;
}
