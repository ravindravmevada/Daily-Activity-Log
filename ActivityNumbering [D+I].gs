function resequenceDColumnAndSubNumbers(sheet) {
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  }

  const incrementColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INCREMENT_NUMBER);
  const subIncrementColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_INCREMENT_COL);
  const dataRowIndicatorColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INSERT_ROW_ABOVE);
  const dataStartRow = 3;
  const lastRow = sheet.getLastRow();

  if (lastRow < dataStartRow) return;

  const indicatorRange = sheet.getRange(dataStartRow, dataRowIndicatorColIndex, lastRow - dataStartRow + 1, 1);
  const indicatorValidations = indicatorRange.getDataValidations();

  const isDataRow = {};
  for (let i = 0; i < indicatorValidations.length; i++) {
    const actualRow = dataStartRow + i;
    const validation = indicatorValidations[i][0];
    if (validation && validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      isDataRow[actualRow] = true;
    }
  }

  const dRange = sheet.getRange(dataStartRow, incrementColIndex, lastRow - dataStartRow + 1, 1);
  const dMerges = dRange.getMergedRanges();

  const mergeTopRows = {};
  const mergeInfo = {};
  for (let i = 0; i < dMerges.length; i++) {
    const m = dMerges[i];
    const topRow = m.getRow();
    const numRows = m.getNumRows();
    mergeTopRows[topRow] = true;
    mergeInfo[topRow] = numRows;
    for (let r = topRow + 1; r < topRow + numRows; r++) {
      mergeTopRows[r] = false;
    }
  }

  let counter = 1;
  for (let r = lastRow; r >= dataStartRow; r--) {
    if (!isDataRow[r]) continue;
    if (mergeTopRows[r] === false) continue;

    sheet.getRange(r, incrementColIndex).setValue(counter);

    if (mergeInfo[r]) {
      const groupNumRows = mergeInfo[r];
      for (let sr = 0; sr < groupNumRows; sr++) {
        const positionFromBottom = groupNumRows - sr;
        const subNumber = parseFloat(counter + "." + positionFromBottom);
        sheet.getRange(r + sr, subIncrementColIndex).setValue(subNumber);
      }
    }

    counter++;
  }
}

function manualResequenceDColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  resequenceDColumnAndSubNumbers(sheet);
}
