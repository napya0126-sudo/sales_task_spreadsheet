// ──────────────────────────────────────
// 定数
// ──────────────────────────────────────
// 担当者の表示名（ドロップダウン・シート名すべてこれで統一）
var PERSONS    = ['Shoko', 'Momoka', 'Shunta', 'Shintaro', 'Naoya'];
var TAB_COLORS = ['#e91e63', '#9c27b0', '#2196f3', '#4caf50', '#ff9800'];

// 対象GoogleドキュメントID
var TARGET_DOC_ID = '1nepU8H6LHowXbU5FZTuH1VZNvHReu7aNK8Bijb_H2OU';

// 列インデックス（1始まり）
var COL_TASK_NAME = 2; // B列
var COL_DOC_URL   = 8; // H列

// メールプレフィックス → 表示名
var EMAIL_TO_NAME = {
  's.anada':    'Shoko',
  'm.tsuji':    'Momoka',
  's.nagai':    'Shunta',
  's.sakamoto': 'Shintaro',
  'n.hasegawa': 'Naoya'
};

// 列構成（A〜K、11列）
// A=No. B=タスク名 C=ステータス D=依頼者 E=依頼先 F=タグ G=優先度 H=Doc URL I=期限 J=病院 K=メモ
var HEADERS    = ['No.', 'タスク名', 'ステータス', '依頼者', '依頼先', 'タグ', '優先度', 'Doc URL', '期限', '病院', 'メモ'];
var COL_WIDTHS = [50, 320, 100, 110, 160, 120, 75, 200, 100, 180, 220];

var HEADER_ROW = 2;
var DATA_START = 3;
var ROWS = 500;
var COLS = HEADERS.length; // 11

// 担当者シートで編集→同期するカラム（C=ステータス, K=メモ）
var SYNC_COLS = [3, 11];

// ──────────────────────────────────────
// メニュー
// ──────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('タスク管理')
    .addItem('🔗 Docに見出しを作成', 'createHeadingAndLink')
    .addItem('📦 完了タスクをアーカイブへ移動', 'archiveCompleted')
    .addSeparator()
    .addItem('📋 書式を再適用（データ保持）', 'buildTaskTemplate')
    .addItem('🔄 担当者シートを最新に同期', 'syncPersonSheets')
    .addItem('👤 担当者シートを作成 / 更新', 'buildPersonSheets')
    .addSeparator()
    .addItem('⚙️ 初期設定（初回のみ・各自実行）', 'setup')
    .addItem('📖 使い方シートを作成', 'buildUsageSheet')
    .addToUi();
}

// ──────────────────────────────────────
// 初期設定
// ──────────────────────────────────────
function setup() {
  var email = Session.getActiveUser().getEmail();
  if (!email) {
    SpreadsheetApp.getUi().alert('⚠️ メールアドレスを取得できませんでした。');
    return;
  }

  var prefix = email.split('@')[0];
  var displayName = EMAIL_TO_NAME[prefix] || prefix;
  PropertiesService.getUserProperties().setProperty('displayName', displayName);

  // インストール可能トリガーを作成（重複防止）
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'handleEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('handleEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ 初期設定完了！\n\n'
    + 'あなたの表示名：' + displayName + '\n\n'
    + 'タスク名を入力すると以下が自動入力されます：\n'
    + '・依頼者：' + displayName + '\n'
    + '・ステータス：未着手\n'
    + '・期限：7日後'
  );
}

// ──────────────────────────────────────
// 編集トリガー（インストール可能トリガー）
// ──────────────────────────────────────
function handleEdit(e) {
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  var row = e.range.getRow();
  var col = e.range.getColumn();

  // ── 全てシート ──
  if (sheetName === '全て') {
    if (row < DATA_START) return;

    // タスク名入力時の自動入力（B列）
    if (col === 2 && e.value) {
      var requestorCell = sheet.getRange(row, 4);
      if (!requestorCell.getValue()) {
        var displayName = PropertiesService.getUserProperties().getProperty('displayName');
        if (displayName) requestorCell.setValue(displayName);
      }
      
      var statusCell = sheet.getRange(row, 3);
      if (!statusCell.getValue()) statusCell.setValue('未着手');
      
      var deadlineCell = sheet.getRange(row, 9);
      if (!deadlineCell.getValue()) {
        var d = new Date();
        d.setDate(d.getDate() + 7);
        deadlineCell.setValue(d);
      }
    }
    
    // ステータスまたはメモが更新されたら担当者シートへ同期
    if (SYNC_COLS.indexOf(col) !== -1) {
      updatePersonSheetsFromMain(row, col);
    }
  }
  // ── 担当者シート ──
  else if (PERSONS.indexOf(sheetName) !== -1) {
    if (row < DATA_START) return;
    if (SYNC_COLS.indexOf(col) !== -1) {
      updateMainFromPersonSheet(sheet, row, col);
    }
  }
}

// ──────────────────────────────────────
// タスク管理テンプレートの作成
// ──────────────────────────────────────
function buildTaskTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('全て');
  if (!sheet) {
    sheet = ss.insertSheet('全て', 0);
  }

  // A=No. B=タスク名 C=ステータス D=依頼者 E=依頼先 F=タグ G=優先度 H=Doc URL I=期限 J=病院 K=メモ
  sheet.getRange(HEADER_ROW, 1, 1, COLS).setValues([HEADERS]);
  
  // 列幅
  for (var i = 0; i < COLS; i++) {
    sheet.setColumnWidth(i + 1, COL_WIDTHS[i]);
  }

  // ヘッダー書式
  var headerRange = sheet.getRange(HEADER_ROW, 1, 1, COLS);
  headerRange.setBackground('#444444')
             .setFontColor('#ffffff')
             .setFontWeight('bold')
             .setHorizontalAlignment('center');

  // データ行の基本書式とドロップダウン
  var dataRange = sheet.getRange(DATA_START, 1, ROWS, COLS);
  dataRange.setVerticalAlignment('middle');

  // ステータス
  var statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['未着手', '進行中', '確認待ち', '完了', 'アーカイブ']).build();
  sheet.getRange(DATA_START, 3, ROWS, 1).setDataValidation(statusRule);

  // 依頼者・依頼先
  var personRule = SpreadsheetApp.newDataValidation().requireValueInList(PERSONS).build();
  sheet.getRange(DATA_START, 4, ROWS, 2).setDataValidation(personRule);

  // タグ
  var tagRule = SpreadsheetApp.newDataValidation().requireValueInList(['不具合再現', 'システム構築', 'マスタ登録', '動画制作', '帳票関係制作', 'その他']).build();
  sheet.getRange(DATA_START, 6, ROWS, 1).setDataValidation(tagRule);

  // 優先度
  var priorityRule = SpreadsheetApp.newDataValidation().requireValueInList(['高', '中', '低']).build();
  sheet.getRange(DATA_START, 7, ROWS, 1).setDataValidation(priorityRule);

  // 条件付き書式（ステータスによる行の色付け）
  sheet.clearConditionalFormatRules();
  var rules = [];

  // 完了 = 緑
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C3="完了"')
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange(DATA_START, 1, ROWS, COLS)])
    .build());

  // 進行中 = 青
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C3="進行中"')
    .setBackground('#c9daf8')
    .setRanges([sheet.getRange(DATA_START, 1, ROWS, COLS)])
    .build());

  // 未着手 = 赤（薄い）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C3="未着手"')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange(DATA_START, 1, ROWS, COLS)])
    .build());

  // アーカイブ = グレー
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C3="アーカイブ"')
    .setBackground('#efefef')
    .setFontColor('#999999')
    .setRanges([sheet.getRange(DATA_START, 1, ROWS, COLS)])
    .build());

  // 期限超過（未完了のもの）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($I3<>"", $I3<TODAY(), $C3<>"完了", $C3<>"アーカイブ")')
    .setBackground('#fce5cd')
    .setRanges([sheet.getRange(DATA_START, 9, ROWS, 1)])
    .build());

  sheet.setConditionalFormatRules(rules);
  
  // ウィンドウ枠の固定
  sheet.setFrozenRows(HEADER_ROW);

  SpreadsheetApp.getUi().alert('✅「全て」シートの書式を再適用しました。');
}

// ──────────────────────────────────────
// 担当者シートの作成 / 更新
// ──────────────────────────────────────
function buildPersonSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName('全て');
  
  for (var i = 0; i < PERSONS.length; i++) {
    var name = PERSONS[i];
    var color = TAB_COLORS[i];
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    sheet.setTabColor(color);
    sheet.clear();

    // ヘッダーコピー
    sheet.getRange(HEADER_ROW, 1, 1, COLS).setValues([HEADERS]);
    sheet.getRange(HEADER_ROW, 1, 1, COLS).setBackground(color).setFontColor('#ffffff').setFontWeight('bold');
    
    for (var j = 0; j < COLS; j++) {
      sheet.setColumnWidth(j + 1, COL_WIDTHS[j]);
    }
    sheet.setFrozenRows(HEADER_ROW);

    // データ抽出（依頼先 = name）
    var data = mainSheet.getRange(DATA_START, 1, mainSheet.getLastRow(), COLS).getValues();
    var filtered = [];
    for (var k = 0; k < data.length; k++) {
      if (data[k][4] === name && data[k][2] !== 'アーカイブ') {
        filtered.push(data[k]);
      }
    }

    if (filtered.length > 0) {
      sheet.getRange(DATA_START, 1, filtered.length, COLS).setValues(filtered);
      // ドロップダウン設定（ステータスとメモのみ編集を想定するが、書式として全適用）
      applyPersonSheetValidations(sheet, filtered.length);
    }
  }
  SpreadsheetApp.getUi().alert('✅ 担当者別のシートを更新しました。');
}

function applyPersonSheetValidations(sheet, rowCount) {
  var statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['未着手', '進行中', '確認待ち', '完了']).build();
  sheet.getRange(DATA_START, 3, rowCount, 1).setDataValidation(statusRule);
}

// ──────────────────────────────────────
// 同期ロジック
// ──────────────────────────────────────

// メイン → 担当者
function updatePersonSheetsFromMain(mainRow, col) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName('全て');
  var no = mainSheet.getRange(mainRow, 1).getValue();
  var personName = mainSheet.getRange(mainRow, 5).getValue();
  var value = mainSheet.getRange(mainRow, col).getValue();

  if (!personName) return;
  var pSheet = ss.getSheetByName(personName);
  if (!pSheet) return;

  var pData = pSheet.getRange(DATA_START, 1, pSheet.getLastRow(), 1).getValues();
  for (var i = 0; i < pData.length; i++) {
    if (pData[i][0] === no) {
      pSheet.getRange(i + DATA_START, col).setValue(value);
      break;
    }
  }
}

// 担当者 → メイン
function updateMainFromPersonSheet(pSheet, pRow, col) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName('全て');
  var no = pSheet.getRange(pRow, 1).getValue();
  var value = pSheet.getRange(pRow, col).getValue();

  var mainData = mainSheet.getRange(DATA_START, 1, mainSheet.getLastRow(), 1).getValues();
  for (var i = 0; i < mainData.length; i++) {
    if (mainData[i][0] === no) {
      mainSheet.getRange(i + DATA_START, col).setValue(value);
      break;
    }
  }
}

// 一括同期ボタン用
function syncPersonSheets() {
  buildPersonSheets();
}

// ──────────────────────────────────────
// アーカイブ移動
// ──────────────────────────────────────
function archiveCompleted() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName('全て');
  var archiveSheet = ss.getSheetByName('アーカイブ');
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet('アーカイブ');
    archiveSheet.getRange(1, 1, 1, COLS).setValues([HEADERS]);
    archiveSheet.setTabColor('#999999');
  }

  var data = mainSheet.getRange(DATA_START, 1, mainSheet.getLastRow(), COLS).getValues();
  var toKeep = [];
  var toArchive = [];

  for (var i = 0; i < data.length; i++) {
    if (data[i][1] === "") continue; // 空行スキップ
    if (data[i][2] === '完了' || data[i][2] === 'アーカイブ') {
      data[i][2] = 'アーカイブ'; // ステータスをアーカイブに統一
      toArchive.push(data[i]);
    } else {
      toKeep.push(data[i]);
    }
  }

  if (toArchive.length > 0) {
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, toArchive.length, COLS).setValues(toArchive);
  }

  // 「全て」シートをクリアして残ったものだけ再配置
  mainSheet.getRange(DATA_START, 1, ROWS, COLS).clearContent();
  if (toKeep.length > 0) {
    mainSheet.getRange(DATA_START, 1, toKeep.length, COLS).setValues(toKeep);
  }

  SpreadsheetApp.getUi().alert('✅ 完了したタスクをアーカイブへ移動しました（' + toArchive.length + '件）。');
}

// ──────────────────────────────────────
// 使い方シート
// ──────────────────────────────────────
function buildUsageSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('使い方');
  if (!sheet) {
    sheet = ss.insertSheet('使い方');
  }
  sheet.clear();
  sheet.setTabColor('#666666');

  var content = [
    ['項目', '説明'],
    ['■ 基本ルール', ''],
    ['', '1.「全て」シートにタスクを追記します。'],
    ['', '2. B列（タスク名）を入力すると、依頼者・ステータス・期限が自動補完されます。'],
    ['', '3. 依頼先（E列）を選ぶと、その担当者のシートに同期されます。'],
    ['', ''],
    ['■ ドロップダウン項目', ''],
    ['', '  ステータス：未着手 / 進行中 / 確認待ち / 完了 / アーカイブ'],
    ['', '  タグ：不具合再現 / システム構築 / マスタ登録 / 動画制作 / 帳票関係制作 / その他'],
    ['', '  優先度：高 / 中 / 低'],
    ['', '  依頼者・依頼先：Shoko / Momoka / Shunta / Shintaro / Naoya'],
    ['', ''],
    ['■ メニュー「タスク管理」の各機能', ''],
    ['', '  🔗 Docに見出しを作成 … 選択行のタスク名をGoogleドキュメント末尾に追加し、リンクを取得します。'],
    ['', '  📋 書式を再適用（データ保持）… ヘッダー・色・ドロップダウンを再適用します。'],
    ['', '  👤 担当者シートを作成 / 更新 … 担当者シートを新規作成または再構築します。'],
    ['', '  🔄 担当者シートを最新に同期 … 「全て」シートのデータを担当者シートに一括反映します。'],
    ['', '  📦 完了タスクをアーカイブへ移動 … ステータスが「完了」のタスクをアーカイブシートに移動します。'],
    ['', ''],
    ['■ 不具合再現依頼について', ''],
    ['', '  不具合再現依頼があったときはメニューからDoc連携を行い、'],
    ['', '  H列に入ったリンクを使って内容をドキュメントに記載してください。'],
  ];

  sheet.getRange(1, 1, content.length, 2).setValues(content);

  // タイトル行の書式
  sheet.getRange(1, 2).setFontSize(16).setFontWeight('bold').setFontColor('#1a73e8');
  sheet.setRowHeight(1, 40);

  // セクション見出しの書式（■で始まる行）
  for (var i = 0; i < content.length; i++) {
    var text = content[i][1];
    if (text.indexOf('■') === 0) {
      sheet.getRange(i + 1, 2).setFontSize(12).setFontWeight('bold').setFontColor('#333333');
      sheet.getRange(i + 1, 1, 1, 2).setBackground('#e8f0fe');
    }
  }

  // 全体の書式
  sheet.getRange(1, 1, content.length, 2).setVerticalAlignment('middle').setWrap(true);
  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('✅「使い方」シートを作成しました！');
}

// ──────────────────────────────────────
// 見出し作成とリンク取得
// ──────────────────────────────────────

/**
 * B列にタスク名が入っている「最新の行（最終行）」を自動判別し、
 * その内容をGoogleドキュメントの末尾に見出しとして追加、リンクを取得してH列に書き込む。
 */
function createHeadingAndLink() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('全て') || ss.getActiveSheet(); // 「全て」シートを優先
  
  // B列の最終行を取得
  const lastRow = sheet.getLastRow();
  let targetRow = 0;
  
  // 下から上にスキャンして、B列に値がある最初の行を探す
  for (let i = lastRow; i >= DATA_START; i--) {
    if (sheet.getRange(i, COL_TASK_NAME).getValue()) {
      targetRow = i;
      break;
    }
  }

  // 1. 基本チェック
  if (targetRow < DATA_START) {
    SpreadsheetApp.getUi().alert('⚠️ 3行目以降にタスク名が見つかりませんでした。');
    return;
  }

  const taskName = sheet.getRange(targetRow, COL_TASK_NAME).getValue();
  
  // 2. 重複（上書き）確認
  const existingUrl = sheet.getRange(targetRow, COL_DOC_URL).getValue();
  if (existingUrl) {
    const res = SpreadsheetApp.getUi().confirm('⚠️ 最新行（' + targetRow + '行目）のH列には既にリンクがあります。重複して作成しますか？');
    if (res !== SpreadsheetApp.getUi().Button.YES) return;
  }

  try {
    // 3. ドキュメントに見出しを追加
    const docId = TARGET_DOC_ID.trim();
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    
    // 末尾に改行を入れてから見出し3を追加
    body.appendParagraph('');
    const heading = body.appendParagraph(taskName);
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    doc.saveAndClose(); // 保存して確定

    // 4. Docs APIを使用して headingId を取得（高速化のため取得フィールドを限定）
    const fields = 'body/content(paragraph(paragraphStyle/headingId,elements/textRun/content))';
    const docData = Docs.Documents.get(TARGET_DOC_ID, { fields: fields });
    const content = docData.body.content;
    
    let headingId = '';
    // 最後から数件分をチェック（末尾に追加しているので、後ろから探すのが最速）
    for (let i = content.length - 1; i >= 0; i--) {
      const element = content[i];
      if (element.paragraph && element.paragraph.paragraphStyle && element.paragraph.paragraphStyle.headingId) {
        const textArr = element.paragraph.elements.map(function(e) {
          return e.textRun ? e.textRun.content : '';
        });
        const text = textArr.join('').trim();
        // 直前に書き込んだタスク名と一致する最初の見出しを採用
        if (text === taskName) {
          headingId = element.paragraph.paragraphStyle.headingId;
          break;
        }
      }
    }

    if (!headingId) {
      throw new Error('見出しの内部IDを取得できませんでした。');
    }

    // 5. URLを生成してH列に書き込む
    const docUrl = 'https://docs.google.com/document/d/' + TARGET_DOC_ID + '/edit#heading=h.' + headingId;
    sheet.getRange(targetRow, COL_DOC_URL).setValue(docUrl);

    SpreadsheetApp.getUi().alert('✅ 完了！（対象：' + targetRow + '行目）\nドキュメントの末尾に見出しを作成し、リンクをH列に記載しました。');

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ エラーが発生しました:\n' + e.message);
  }
}
