function updateActivityId(sheet, row) {
  const idCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_ACTIVITY_ID.ID_COL);
  const prefix = DAILY_ACTIVITY_LOG_ACTIVITY_ID.PREFIX;
  const idCell = sheet.getRange(row, idCol);
  const current = idCell.getValue();

  if (typeof current === "string" && current.indexOf(prefix + "-") === 0) return;

  if (activityQualifiesForId(sheet, row)) {
    const datePart = getSheetDateCompact(sheet);
    const nextNum = getNextActivityIdNumber(sheet, idCol, prefix, datePart);
    writeActivityId(idCell, prefix + "-" + datePart + "-" + ("00" + nextNum).slice(-3));
  } else {
    const cat = computeRowDisplay(sheet, row).category;
    writeActivityEmoji(idCell, cat === "RED" ? "🔴" : cat === "GREEN" ? "🟢" : "🟡");
  }
}

function resequenceActivityIds(sheet) {
  const idCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_ACTIVITY_ID.ID_COL);
  const prefix = DAILY_ACTIVITY_LOG_ACTIVITY_ID.PREFIX;
  const datePart = getSheetDateCompact(sheet);
  const dataStartRow = 3;
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return;

  const numRows = lastRow - dataStartRow + 1;
  const g = DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS;
  const startHasCol = columnLetterToIndex(g[0].hasTimestampCol);
  const startTsCol  = columnLetterToIndex(g[0].timestampCol);
  const endHasCol   = columnLetterToIndex(g[1].hasTimestampCol);
  const endTsCol    = columnLetterToIndex(g[1].timestampCol);

  const startHasVals = sheet.getRange(dataStartRow, startHasCol, numRows, 1).getValues();
  const startTsVals  = sheet.getRange(dataStartRow, startTsCol,  numRows, 1).getValues();
  const endHasVals   = sheet.getRange(dataStartRow, endHasCol,   numRows, 1).getValues();
  const endTsVals    = sheet.getRange(dataStartRow, endTsCol,    numRows, 1).getValues();

  const yesLike = v => v === "Yes" || v === "Yes (Approx.)" || v === "Yes (Approx)";
  const noLike  = v => v === "No" || v === "Not Applicable";
  const filled  = v => parseTimestampString(v) !== null;

  const values = new Array(numRows);
  const fontColors = new Array(numRows);
  const fontWeights = new Array(numRows);

  let counter = 0;
  for (let i = numRows - 1; i >= 0; i--) {
    const sDone = yesLike(startHasVals[i][0]) && filled(startTsVals[i][0]);
    const eDone = yesLike(endHasVals[i][0]) && filled(endTsVals[i][0]);

    if (sDone || eDone) {
      counter++;
      values[i] = [prefix + "-" + datePart + "-" + ("00" + counter).slice(-3)];
      fontColors[i] = ["#f7ef00"];
      fontWeights[i] = ["bold"];
    } else {
      const red = noLike(startHasVals[i][0]) || noLike(endHasVals[i][0]);
      values[i] = [red ? "🔴" : "🟡"];
      fontColors[i] = ["#ffffff"];
      fontWeights[i] = ["normal"];
    }
  }

  const idRange = sheet.getRange(dataStartRow, idCol, numRows, 1);
  idRange.setValues(values);
  idRange.setFontColors(fontColors);
  idRange.setFontWeights(fontWeights);
  idRange.setHorizontalAlignment("center");
}

function activityQualifiesForId(sheet, row) {
  const g = DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS;
  const startHas = sheet.getRange(row, columnLetterToIndex(g[0].hasTimestampCol)).getValue();
  const startTs  = sheet.getRange(row, columnLetterToIndex(g[0].timestampCol)).getValue();
  const endHas   = sheet.getRange(row, columnLetterToIndex(g[1].hasTimestampCol)).getValue();
  const endTs    = sheet.getRange(row, columnLetterToIndex(g[1].timestampCol)).getValue();

  const yesLike = v => v === "Yes" || v === "Yes (Approx.)" || v === "Yes (Approx)";
  const filled  = v => parseTimestampString(v) !== null;

  const startDone = yesLike(startHas) && filled(startTs);
  const endDone   = yesLike(endHas)   && filled(endTs);

  return startDone || endDone;
}

function writeActivityId(cell, text) {
  cell.setValue(text);
  cell.setFontWeight("bold");
  cell.setFontColor("#f7ef00");
  cell.setHorizontalAlignment("center");
}

function writeActivityEmoji(cell, emoji) {
  cell.setValue(emoji);
  cell.setFontWeight("normal");
  cell.setFontColor("#ffffff");
  cell.setHorizontalAlignment("center");
}

function getSheetDateCompact(sheet) {
  const parts = sheet.getName().trim().split("-");
  if (parts.length === 3) return parts[2] + parts[1] + parts[0];
  return sheet.getName().replace(/[^0-9]/g, "");
}

function getNextActivityIdNumber(sheet, idCol, prefix, datePart) {
  const dataStartRow = 3;
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return 1;
  const values = sheet.getRange(dataStartRow, idCol, lastRow - dataStartRow + 1, 1).getValues();
  const needle = prefix + "-" + datePart + "-";
  let max = 0;
  for (let i = 0; i < values.length; i++) {
    const v = values[i][0];
    if (typeof v === "string" && v.indexOf(needle) === 0) {
      const n = parseInt(v.substring(needle.length), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  }
  return max + 1;
}
