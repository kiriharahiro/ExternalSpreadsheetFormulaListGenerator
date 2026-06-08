/**
 * 外部スプレッドシート 数式一覧化ツール
 * 
 * 「設定表」シートに指定されたスプレッドシートIDと対象シート名の設定を読み込み、
 * 該当する外部スプレッドシート内の数式を一覧化して、このスプレッドシートに書き出します。
 */

/**
 * スプレッドシートが開かれたときに実行される関数。
 * カスタムメニューを追加します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 数式一覧化ツール')
    .addItem('設定表に基づいて数式を一覧化する', 'exportFormulasFromSetting')
    .addToUi();
}

/**
 * 設定表に基づいて数式を抽出し、一覧化するメイン関数
 */
function exportFormulasFromSetting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. 「設定表」シートの取得
  const settingSheet = ss.getSheetByName('設定表');
  if (!settingSheet) {
    ui.alert('エラー', '「設定表」という名前 of シートが見つかりません。シート名が正しいかご確認ください。', ui.ButtonSet.OK);
    return;
  }

  // 2. 設定値の読み込み
  // B4が「spreadsheetId」で、C4がその値
  const targetSpreadsheetId = String(settingSheet.getRange('C4').getValue()).trim();
  // B6が「sheetName」で、C6がその値（レイアウト変更に対応：C5はappScriptIdになったためC6を読み込む）
  const targetSheetConfig = String(settingSheet.getRange('C6').getValue()).trim();

  if (!targetSpreadsheetId) {
    ui.alert('エラー', '「設定表」シートの C4 セルにスプレッドシートID（spreadsheetId）を入力してください。', ui.ButtonSet.OK);
    return;
  }

  // 3. 対象シート名の収集
  // C6が「*ALL」の場合はすべてのシート、そうでない場合は C6以下のセルに入力されているシート名を収集
  const targetSheetNames = [];
  let isAllSheets = false;

  if (targetSheetConfig === '*ALL' || targetSheetConfig === '') {
    isAllSheets = true;
  } else {
    // C6 から C16 までを取得して配列にする（レイアウト変更に対応）
    const configValues = settingSheet.getRange('C6:C16').getValues();
    configValues.forEach(row => {
      const name = String(row[0]).trim();
      if (name && name !== '*ALL') {
        targetSheetNames.push(name);
      }
    });
  }

  // 4. 外部スプレッドシートを開く
  let targetSS;
  try {
    targetSS = SpreadsheetApp.openById(targetSpreadsheetId);
  } catch (e) {
    ui.alert(
      '接続エラー',
      '指定されたスプレッドシートIDを開くことができませんでした。\n\n【原因として考えられること】\n・スプレッドシートIDが間違っている\n・対象ファイルへのアクセス権限がない（共有されていない）\n・GASの初回実行時の承認で「許可」をしていない\n\n設定されているIDと共有権限をご確認ください。',
      ui.ButtonSet.OK
    );
    return;
  }

  // 5. 実行確認
  const confirmMsg = isAllSheets 
    ? `外部スプレッドシート「${targetSS.getName()}」の【すべてのシート】から数式を抽出します。よろしいですか？`
    : `外部スプレッドシート「${targetSS.getName()}」の指定されたシート（${targetSheetNames.join(', ')}）から数式を抽出します。よろしいですか？`;

  const confirm = ui.alert('数式一覧の作成開始', confirmMsg, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) {
    ui.alert('処理をキャンセルしました。');
    return;
  }

  // 6. 出力先「数式一覧」シートの準備
  const outputSheetName = '数式一覧';
  let outputSheet = ss.getSheetByName(outputSheetName);

  if (outputSheet) {
    outputSheet.clear();
  } else {
    outputSheet = ss.insertSheet(outputSheetName);
  }

  // オーナーメールアドレスの取得
  let ownerEmail = '';
  try {
    const owner = targetSS.getOwner();
    if (owner) {
      ownerEmail = owner.getEmail();
    }
  } catch (e) {
    ownerEmail = '取得権限なし';
  }

  // 1〜3行目の管理情報のセット
  outputSheet.getRange('A1').setValue('ファイル名');
  outputSheet.getRange('B1').setValue(targetSS.getName());
  outputSheet.getRange('A2').setValue('SpreadsheetId');
  outputSheet.getRange('B2').setValue(targetSS.getId());
  outputSheet.getRange('A3').setValue('ownerAccount');
  outputSheet.getRange('B3').setValue(ownerEmail);

  // 4行目のヘッダーのセット
  const outputHeader = ['対象シート', 'セル位置', '数式', '関数の役割（メモ用）'];
  outputSheet.getRange('A4:D4').setValues([outputHeader]);

  const outputRows = [];

  // 7. 外部スプレッドシートの各シートをスキャン
  const sheets = targetSS.getSheets();
  let scannedCount = 0;

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    
    // スキャン対象か判定
    if (!isAllSheets && !targetSheetNames.includes(sheetName)) {
      return; // 対象外のシートはスキップ
    }

    scannedCount++;
    const range = sheet.getDataRange();
    const formulas = range.getFormulas(); // 数式を取得

    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) {
        const formula = formulas[r][c];
        if (formula && formula.startsWith('=')) {
          const cellNotation = sheet.getRange(r + 1, c + 1).getA1Notation();
          outputRows.push([
            sheetName,      // シート名
            cellNotation,   // セル位置
            "'" + formula,  // アポストロフィを付加してテキストとして出力
            ''              // メモ用
          ]);
        }
      }
    }
  });

  if (scannedCount === 0) {
    ui.alert('完了', '指定された名前のシートが見つかりませんでした。「設定表」のシート名を確認してください。', ui.ButtonSet.OK);
    return;
  }

  if (outputRows.length === 0) {
    ui.alert('完了', `スキャンしたシート（${scannedCount}枚）の中に数式（＝から始まるセル）が見つかりませんでした。`, ui.ButtonSet.OK);
    return;
  }

  // 8. 「数式一覧」シートへの書き込み (5行目からデータ開始)
  outputSheet.getRange(5, 1, outputRows.length, outputHeader.length).setValues(outputRows);

  // 9. デザイン・フォーマット調整
  // A1:A3 のデザイン（深緑ヘッダー色）
  const infoLabelRange = outputSheet.getRange('A1:A3');
  infoLabelRange.setBackground('#2E7D32'); 
  infoLabelRange.setFontColor('#FFFFFF');
  infoLabelRange.setFontWeight('bold');
  infoLabelRange.setHorizontalAlignment('center');

  // B1:B3 のデザイン（薄緑背景）
  const infoValueRange = outputSheet.getRange('B1:B3');
  infoValueRange.setBackground('#E8F5E9');
  infoValueRange.setFontColor('#000000');
  infoValueRange.setHorizontalAlignment('left');

  // A4:D4 (メインヘッダー) のデザイン
  const headerRange = outputSheet.getRange('A4:D4');
  headerRange.setBackground('#2E7D32'); 
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // 全体データ範囲の罫線設定
  const lastRow = 4 + outputRows.length;
  const totalRange = outputSheet.getRange(1, 1, lastRow, outputHeader.length);
  totalRange.setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  
  // 各列の配置調整
  outputSheet.getRange(5, 1, outputRows.length, 2).setHorizontalAlignment('center'); // シート名, セル位置
  outputSheet.getRange(5, 3, outputRows.length, 2).setHorizontalAlignment('left');   // 数式, メモ欄

  // 列幅を自動調整
  for (let col = 1; col <= outputHeader.length; col++) {
    outputSheet.autoResizeColumn(col);
  }

  SpreadsheetApp.flush();

  ui.alert(
    '完了',
    `外部スプレッドシート「${targetSS.getName()}」から数式の抽出が完了しました！\n\n・スキャンしたシート数: ${scannedCount} 枚\n・検出された数式の数: ${outputRows.length} 個\n\n「${outputSheetName}」シートをご確認ください。`,
    ui.ButtonSet.OK
  );
}
