function sortLargeRowsAsc() {
  const considerTrueFalse = getConsiderTrueFalse();
  sortLargeRowsByAColumn(true, considerTrueFalse);
}

function sortLargeRowsDesc() {
  const considerTrueFalse = getConsiderTrueFalse();
  sortLargeRowsByAColumn(false, considerTrueFalse);
}

function toggleConsiderTrueFalse() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("setting");
  const currentValue = sheet.getRange("B3").getValue();
  const newValue = !currentValue;
  sheet.getRange("B3").setValue(newValue);
  onOpen();
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(`✅単語の除外： ${newValue ? 'ON' : 'OFF'} `);
}

function getConsiderTrueFalse() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("setting");
  return sheet.getRange("B3").getValue() === true;
}

function sortLargeRowsByAColumn(isAscending, considerTrueFalse) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sheet 1");
  const blockSize = 6;
  const startRow = 2;
  const numCols = 7;
  const lastRow = sheet.getLastRow();
  const totalRows = lastRow - startRow + 1;
  const numBlocks = Math.floor(totalRows / blockSize);

  const trueBlocks = [];
  const falseBlocks = [];
  const otherBlocks = [];

  for (let i = 0; i < numBlocks; i++) {
    const row = startRow + i * blockSize;
    if (row + blockSize - 1 > lastRow) break;

    const range = sheet.getRange(row, 1, blockSize, numCols);
    const values = range.getValues();
    const key = values[0][0];
    const flag = values[0][6]; // G列（index 6）

    const block = { key, values };
    if (flag === true) {
      trueBlocks.push(block);
    } else if (flag === false) {
      falseBlocks.push(block);
    } else {
      otherBlocks.push(block);
    }
  }

  if (considerTrueFalse) {
    // TRUEを下、FALSEを上に
    falseBlocks.sort((a, b) => {
      return isAscending ? (a.key > b.key ? 1 : -1) : (a.key < b.key ? 1 : -1);
    });

    trueBlocks.sort((a, b) => {
      return isAscending ? (a.key > b.key ? 1 : -1) : (a.key < b.key ? 1 : -1);
    });

    const sortedBlocks = falseBlocks.concat(trueBlocks).concat(otherBlocks);
    for (let i = 0; i < sortedBlocks.length; i++) {
      const row = startRow + i * blockSize;
      const range = sheet.getRange(row, 1, blockSize, numCols);
      range.setValues(sortedBlocks[i].values);
      
      // G列に数式を設定
      const formulaCell = sheet.getRange(row, 7); // G列 (2+m*6行目)
      formulaCell.setFormula(`=A${5 + 6 * i}`);
    }
  } else {
    // TRUE/FALSE考慮しない
    const allBlocks = trueBlocks.concat(falseBlocks).concat(otherBlocks);
    allBlocks.sort((a, b) => {
      return isAscending ? (a.key > b.key ? 1 : -1) : (a.key < b.key ? 1 : -1);
    });

    for (let i = 0; i < allBlocks.length; i++) {
      const row = startRow + i * blockSize;
      const range = sheet.getRange(row, 1, blockSize, numCols);
      range.setValues(allBlocks[i].values);

      // G列に数式を設定
      const formulaCell = sheet.getRange(row, 7); // G列 (2+m*6行目)
      formulaCell.setFormula(`=A${5 + 6 * i}`);
    }
  }
}

function shuffleLargeRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sheet 1");
  const blockSize = 6;
  const startRow = 2;
  const numCols = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const totalRows = lastRow - startRow + 1;
  const numBlocks = Math.floor(totalRows / blockSize);

  const trueBlocks = [];
  const falseBlocks = [];
  const otherBlocks = [];

  for (let i = 0; i < numBlocks; i++) {
    const row = startRow + i * blockSize;
    if (row + blockSize - 1 > lastRow) break;

    const range = sheet.getRange(row, 1, blockSize, numCols);
    const values = range.getValues();
    const flag = values[0][6]; // G列

    if (flag === true) {
      trueBlocks.push(values);
    } else if (flag === false) {
      falseBlocks.push(values);
    } else {
      otherBlocks.push(values);
    }
  }

  // Fisher-Yates Shuffle
  function shuffleArray(arr) {
    for (let i = arr.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [arr[i], arr[j]] = [arr[j], arr[i]];
    }
  }

  shuffleArray(falseBlocks);
  shuffleArray(trueBlocks);

  const allBlocks = falseBlocks.concat(trueBlocks).concat(otherBlocks);

  for (let i = 0; i < allBlocks.length; i++) {
    const row = startRow + i * blockSize;
    const range = sheet.getRange(row, 1, blockSize, numCols);
    range.setValues(allBlocks[i]);

    // G列に数式を設定
    const formulaCell = sheet.getRange(row, 7); // G列 (2+m*6行目)
    formulaCell.setFormula(`=A${5 + 6 * i}`);
  }
}