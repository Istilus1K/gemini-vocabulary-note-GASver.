function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menue')  
    .addItem('Fetch Word', 'fetchWord')   
    .addItem('Add New Sheet', 'addNewSheet') 
    .addToUi();
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

  const lastRow = mainSheet.getLastRow();

  if (lastRow === 1) {
  SpreadsheetApp.getUi().alert('データが存在しません。');
  return;
  }

  const word = mainSheet.getRange(lastRow, 1).getValue();

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
        mainSheet.getRange(lastRow, 2).setValue(lines[1]);
        mainSheet.getRange(lastRow, 3).setValue(lines[2]);
        mainSheet.getRange(lastRow, 4).setValue(lines[3]);
        mainSheet.getRange(lastRow+1, 4).setValue(lines[4]);
        mainSheet.getRange(lastRow+2, 4).setValue(lines[5]);
        mainSheet.getRange(lastRow+3, 4).setValue(lines[6]);
        mainSheet.getRange(lastRow+4, 4).setValue(lines[7]);
        mainSheet.getRange(lastRow+5, 4).setValue(lines[8]);
        mainSheet.getRange(lastRow, 2, 6, 3).setFontColor("#0000FF");
        SpreadsheetApp.flush();
        
        const ui = SpreadsheetApp.getUi();
        const userResponse = ui.alert('内容を確定しますか？', ui.ButtonSet.YES_NO);
        
        if (userResponse == ui.Button.YES) {
          mainSheet.getRange(lastRow, 2, 6, 3).setFontColor("black");
        } else {
          mainSheet.getRange(lastRow, 2, 6, 3).clearContent();
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

  const lastRow = newSheet.getLastRow();
  if (lastRow > 1) {
    newSheet.getRange(2, 1, lastRow - 1, newSheet.getLastColumn()).clearContent();
  }

  const settingSheet = spreadsheet.getSheetByName('setting');
  if (settingSheet) {
    spreadsheet.setActiveSheet(settingSheet);
    spreadsheet.moveActiveSheet(sheets.length+1); 
  }
  spreadsheet.setActiveSheet(newSheet);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('新しいシートが追加されました。');
}