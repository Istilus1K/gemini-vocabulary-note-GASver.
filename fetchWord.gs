function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menue')  
    .addItem('Fetch Word', 'fetchWord')   
    .addToUi();
}

function fetchWord() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName('Main');
  const settingSheet = spreadsheet.getSheetByName('setting');

  if (!mainSheet || !settingSheet) {
    SpreadsheetApp.getUi().alert('指定されたシートが見つかりません。シート名を確認してください。');
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

  const prompt = `${language}の${word}という単語またはフレーズを辞書的に解説してください。\n
    + 以下の9つ情報を、それぞれ改行して返してください:\n
    + 1. 単語 (word),\n
    + 2. 発音 (pronunciation),\n
    + 3. 意味 (meaning),\n
    + 4. 例文1 (example1),\n
    + 5. 日本語訳1 (example1_translation),\n
    + 6. 例文2 (example2),\n
    + 7. 日本語訳2 (example2_translation),\n
    + 8. 例文3 (example3),\n
    + 9. 日本語訳3 (example3_translation)。\n\n
    + もし情報が見つからない場合や不完全な場合は、次のように回答してください：\n
    + - 結果が見つかりませんでした。\n
    + - 情報が不足しています（9行で返してください）。\n\n
    + 例 (インドネシア語のhatiという単語):\n
    + hati\n
    + /ˈɑːti/\n
    + 心\n
    + Hati saya senang\n
    + 私の心は喜んでいる\n
    + Dia adalah orang yang baik hati\n
    + 彼/彼女は心優しい人です\n
    + Hati-hati!\n
    + 気を付けて!\n`;

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