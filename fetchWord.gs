function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const considerTrueFalse = getConsiderTrueFalse();

  // 最初のメニューを作成
  ui.createMenu('Menue')  
    .addItem('Fetch Word', 'fetchWord')   
    .addItem('Add New Sheet', 'addNewSheet') 
    .addToUi();

  // Sortメニューを作成し、状態に応じて項目を更新
  const menu = ui.createMenu("📊 Sort")
    .addItem("ascending sort", "sortLargeRowsAsc")
    .addItem("descending sort", "sortLargeRowsDesc")
    .addItem("shuffle", "shuffleLargeRows")
    .addItem(`Exclude ✅s: ${considerTrueFalse ? 'ON' : 'OFF'}`, "toggleConsiderTrueFalse")
    .addToUi();

}

function toggleConsiderTrueFalse() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("setting");
  const currentValue = sheet.getRange("B3").getValue();
  const newValue = !currentValue;
  sheet.getRange("B3").setValue(newValue);
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(`✅単語の除外： ${newValue ? 'ON' : 'OFF'} `);

  // メニューを更新
  onOpen();
}

function fetchWord() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getActiveSheet();  
  const settingSheet = spreadsheet.getSheetByName('setting');

  if (!mainSheet || !settingSheet) {
    SpreadsheetApp.getUi().alert('指定されたシートが見つかりません。シート名を確認してください。');
    return;
  }

  if (mainSheet.getName() === 'setting') {
    SpreadsheetApp.getUi().alert('このシートでは実行できません。別のシートを開いてから実行してください。');
    return;
  }

  const language = settingSheet.getRange('B1').getValue();
  const apiKey = settingSheet.getRange('B2').getValue();
  if (!apiKey) {
    Logger.log("APIキーが設定されていません。");
    return;
  }

  let lastRow = mainSheet.getRange("B:B").getLastRow()-5;

  if (lastRow === 1) {
  SpreadsheetApp.getUi().alert('データが存在しません。');
  return;
  }

  const word = mainSheet.getRange(lastRow, 2).getValue();

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${apiKey}`;
  const headers = { 'Content-Type': 'application/json' };

  const prompt = `${language} の「${word}」という単語またはフレーズを辞書的に解説してください。\n\n` +
    "以下の9つの情報を、指定された順序で、改行区切りで返してください。\n\n" +
    "1. 単語 (word)\n" +
    "2. 発音 (pronunciation)（分かる場合は IPA で、分からない場合は「不明」と記載）\n" +
    "3. 意味 (meaning)（必ず日本語で書く）\n" +
    "4. 例文1 (example1)\n" +
    "5. 日本語訳1 (example1_translation)\n" +
    "6. 例文2 (example2)\n" +
    "7. 日本語訳2 (example2_translation)\n" +
    "8. 例文3 (example3)\n" +
    "9. 日本語訳3 (example3_translation)\n\n" +
    "### 追加ルール\n" +
    "- すべての項目を改行で区切ること。\n" +
    "- 例文は自然な文脈で使用されるものにすること。\n" +
    "- 情報が不足している場合は、「情報が不足しています」と記載し、9行で出力すること。\n" +
    "- 結果が見つからない場合は「結果が見つかりませんでした。」と出力すること。\n\n" +
    "### 出力例（インドネシア語「hati」の場合）\n\n" +
    "```\n" +
    "hati\n" +
    "ˈhati\n" +
    "心\n" +
    "Hati saya senang\n" +
    "私の心は喜んでいる\n" +
    "Dia adalah orang yang baik hati\n" +
    "彼/彼女は心優しい人です\n" +
    "Hati-hati!\n" +
    "気を付けて!\n" +
    "```";

  const payload = {
    'contents': [{ 'parts': [{ 'text': prompt }] }]
  };

  const objOptions = {
    'method': 'post',
    'headers': headers,
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, objOptions);
    const json = JSON.parse(response.getContentText());
    Logger.log(json);

    if (json.candidates && json.candidates.length > 0) {
      const text = json.candidates[0].content.parts[0].text;
      Logger.log(text);

      const lines = text.split('\n').filter(line => line.trim() !== '');  

      if (lines.length >= 9) {
        mainSheet.getRange(lastRow, 3).setValue(lines[1]);
        mainSheet.getRange(lastRow, 4).setValue(lines[2]);
        mainSheet.getRange(lastRow, 5).setValue(lines[3]);
        mainSheet.getRange(lastRow+1, 5).setValue(lines[4]);
        mainSheet.getRange(lastRow+2, 5).setValue(lines[5]);
        mainSheet.getRange(lastRow+3, 5).setValue(lines[6]);
        mainSheet.getRange(lastRow+4, 5).setValue(lines[7]);
        mainSheet.getRange(lastRow+5, 5).setValue(lines[8]);
        mainSheet.getRange(lastRow, 3, 6, 3).setFontColor("#0000FF");
        SpreadsheetApp.flush();
        
        const ui = SpreadsheetApp.getUi();
        const userResponse = ui.alert('内容を確定しますか？', ui.ButtonSet.YES_NO);
        
        if (userResponse == ui.Button.YES) {
          mainSheet.getRange(lastRow, 3, 6, 3).setFontColor("black");
          addNewBlock();
        } else {
          mainSheet.getRange(lastRow, 3, 6, 3).clearContent();
        }

      } else {
        Logger.log("APIのレスポンスが期待したフォーマットではありません。");
      }

    } else {
      Logger.log("APIから適切なレスポンスが得られませんでした。");
    }
  } catch (error) {
    Logger.log("エラー: " + error.toString());
  }
}

function addNewBlock() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const lastRow = sheet.getRange("B:B").getLastRow() - 5;

  const sourceRange = sheet.getRange(lastRow, 1, 6, sheet.getLastColumn() - 1);  // B列から最終列まで
  const destinationStartRow = lastRow + 6;

  // 必要な行数を確認し、シートの行数が足りなければ追加
  const sheetMaxRows = sheet.getMaxRows();
  const requiredRows = destinationStartRow + 5;
  if (sheetMaxRows < requiredRows) {
    sheet.insertRowsAfter(sheetMaxRows, requiredRows - sheetMaxRows);
  }

  // 書式と値をコピー
  const destinationRange = sheet.getRange(destinationStartRow, 1, 6, sheet.getLastColumn() - 1);
  sourceRange.copyTo(destinationRange, { contentsOnly: false });

  /////機能を追加/////

  // 1. B~F列の値をクリア（複製部分）
  sheet.getRange(destinationStartRow, 2, 6, 5).clearContent();

  // 2. A列の一番上のセルに (全行数 - 1) / 6 +1 の値をセット
  const totalRows = sheet.getLastRow();
  const blockNumber = Math.floor((totalRows - 1) / 6 + 1 );
  sheet.getRange(destinationStartRow, 1).setValue(blockNumber);

  // 3. G列の一番上に =A◯ という式をセット（◯はA列の一番上+3のセルの行番号）
  const targetRow = destinationStartRow + 3;
  sheet.getRange(destinationStartRow, 7).setFormula(`=A${targetRow}`);
}




function addNewSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  
  const sourceSheet = sheets.find(sheet => sheet.getName() !== 'setting');
  
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('コピー元となるシートが見つかりません。');
    return;
  }

  const newSheet = sourceSheet.copyTo(spreadsheet);
  newSheet.setName('sheet ' + sheets.length); 

  // B列～F列（2〜7行目）の内容を削除
  newSheet.getRange(2, 2, 6, 5).clearContent(); // (開始行, 開始列, 行数, 列数)

  // セル A2 に「1」をセット
  newSheet.getRange("A2").setValue(1);
  newSheet.getRange("G2").setValue("=A5");

  // 最新の最終行を取得してから8行目以降を削除
  SpreadsheetApp.flush(); // 内容変更を反映
  const updatedLastRow = newSheet.getLastRow();
  if (updatedLastRow > 7) {
    newSheet.deleteRows(8, updatedLastRow - 7); // 8行目から最後の行までを削除
    newSheet.deleteRow(8);
    newSheet.deleteRow(8);
  }

  const settingSheet = spreadsheet.getSheetByName('setting');
  if (settingSheet) {
    spreadsheet.setActiveSheet(settingSheet);
    spreadsheet.moveActiveSheet(sheets.length + 1); 
  }

  spreadsheet.setActiveSheet(newSheet);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('新しいシートが追加されました。');
}