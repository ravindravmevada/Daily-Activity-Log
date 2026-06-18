function isRowGroupUnmergerEnabled() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableFeatureRange = ss.getRangeByName(ROW_GROUP_UNMERGER.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange) return false;
  return enableFeatureRange.getValue() === true;
}

function setRguOperationResult(message, fontColor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ROW_GROUP_UNMERGER.OPERATION_RESULT_CELL);
  if (!resultRange) return;
  resultRange.setValue(message)
    .setBackground("#434343")
    .setFontFamily("Lexend")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontStyle("normal")
    .setFontColor(fontColor)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
}

function handleRguEnableFeature(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableFeatureRange = ss.getRangeByName(ROW_GROUP_UNMERGER.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange) return;
  if (row !== enableFeatureRange.getRow() || col !== enableFeatureRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true && e.value !== "FALSE" && e.value !== false)) return;

  const isEnabled = (e.value === "TRUE" || e.value === true);
  setFeatureStatusCell(ROW_GROUP_UNMERGER.FEATURE_STATUS_CELL, isEnabled);

  const targetSheetRange = ss.getRangeByName(ROW_GROUP_UNMERGER.TARGET_SHEET_CELL);
  const startRowRange = ss.getRangeByName(ROW_GROUP_UNMERGER.MERGE_START_ROW_CELL);
  const endRowRange = ss.getRangeByName(ROW_GROUP_UNMERGER.MERGE_END_ROW_CELL);

  if (isEnabled) {
    if (targetSheetRange) targetSheetRange.setValue("").setFontColor("#ffffff");
    if (startRowRange) startRowRange.setValue("").setFontColor("#ffffff");
    if (endRowRange) endRowRange.setValue("").setFontColor("#ffffff");
  } else {
    if (targetSheetRange) {
      targetSheetRange.setValue("DD-MM-YYYY")
        .setFontFamily("Lexend").setFontSize(11)
        .setFontWeight("normal").setFontStyle("normal")
        .setFontColor("#b7b7b7");
    }
    if (startRowRange) {
      startRowRange.setValue("Start Row")
        .setFontFamily("Lexend").setFontSize(11)
        .setFontWeight("normal").setFontStyle("normal")
        .setFontColor("#b7b7b7");
    }
    if (endRowRange) {
      endRowRange.setValue("End Row")
        .setFontFamily("Lexend").setFontSize(11)
        .setFontWeight("normal").setFontStyle("normal")
        .setFontColor("#b7b7b7");
    }
  }
}

function handleRguUnmergeNow(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const unmergeNowRange = ss.getRangeByName(ROW_GROUP_UNMERGER.UNMERGE_NOW_TICKBOX);
  if (!unmergeNowRange) return;
  if (row !== unmergeNowRange.getRow() || col !== unmergeNowRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true)) return;

  if (!isRowGroupUnmergerEnabled()) {
    setFeatureNotEnabledWarning(ROW_GROUP_UNMERGER.FEATURE_STATUS_CELL);
    unmergeNowRange.setValue(false);
    return;
  }

  const targetSheetRange = ss.getRangeByName(ROW_GROUP_UNMERGER.TARGET_SHEET_CELL);
  const startRowRange = ss.getRangeByName(ROW_GROUP_UNMERGER.MERGE_START_ROW_CELL);
  const endRowRange = ss.getRangeByName(ROW_GROUP_UNMERGER.MERGE_END_ROW_CELL);

  if (!targetSheetRange || !startRowRange || !endRowRange) {
    unmergeNowRange.setValue(false);
    return;
  }

  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") {
    setRguOperationResult("⚠️ No target sheet specified", "#faab17");
    unmergeNowRange.setValue(false);
    return;
  }

  const startRow = startRowRange.getValue();
  const endRow = endRowRange.getValue();

  if (typeof startRow !== "number" || typeof endRow !== "number") {
    setRguOperationResult("⚠️ Enter valid start and end row numbers", "#faab17");
    unmergeNowRange.setValue(false);
    return;
  }

  if (startRow > endRow) {
    setRguOperationResult("⚠️ Start row must be less than or equal to end row", "#faab17");
    unmergeNowRange.setValue(false);
    return;
  }

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) {
    setRguOperationResult("❌ Sheet not found: " + tabName, "#ff0000");
    unmergeNowRange.setValue(false);
    return;
  }

  executeRowGroupUnmerge(targetSheet, startRow, endRow);
  setRguOperationResult("✅ Unmerged rows " + startRow + " to " + endRow + " on " + tabName, "#00ff00");
  unmergeNowRange.setValue(false);
}

function executeRowGroupUnmerge(sheet, startRow, endRow) {
  const lastCol = sheet.getLastColumn();
  const numRows = endRow - startRow + 1;

  sheet.getRange(startRow, 1, numRows, lastCol).breakApart();

  const categoryCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_START_COL);
  const subcategoryCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_END_COL);
  const startHasCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[0].hasTimestampCol);
  const endHasCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[1].hasTimestampCol);
  const startTsCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[0].timestampCol);
  const endTsCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[1].timestampCol);
  const activityTypeCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.ACTIVITY_TYPE_COL);

  for (let r = startRow; r <= endRow; r++) {
    sheet.getRange(r, categoryCol).setValue("No Status");
    sheet.getRange(r, subcategoryCol).setValue("No Status");
    sheet.getRange(r, activityTypeCol).setValue("No Status");
    sheet.getRange(r, startHasCol).setValue("No Status");
    sheet.getRange(r, endHasCol).setValue("No Status");
    sheet.getRange(r, startTsCol).setValue("");
    sheet.getRange(r, endTsCol).setValue("");
  }

  resequenceDColumnAndSubNumbers(sheet);

  for (let r = startRow; r <= endRow; r++) {
    updateRowDuration(sheet, r);
  }
}

function handleRguResetToDefault(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resetRange = ss.getRangeByName(ROW_GROUP_UNMERGER.RESET_TO_DEFAULT_TICKBOX);
  if (!resetRange) return;
  if (row !== resetRange.getRow() || col !== resetRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true)) return;

  initRowGroupUnmerger();
  resetRange.setValue(false);
}

function initRowGroupUnmerger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const resultRange = ss.getRangeByName(ROW_GROUP_UNMERGER.OPERATION_RESULT_CELL);
  if (resultRange) {
    resultRange.setValue("✨ Unmerge Operation Result ✨")
      .setBackground("#434343")
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setFontWeight("normal")
      .setFontStyle("normal")
      .setFontColor("#f3f3f3")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);
  }

  const targetSheetRange = ss.getRangeByName(ROW_GROUP_UNMERGER.TARGET_SHEET_CELL);
  if (targetSheetRange) {
    targetSheetRange.setValue("DD-MM-YYYY")
      .setFontFamily("Lexend").setFontSize(11)
      .setFontWeight("normal").setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const startRowRange = ss.getRangeByName(ROW_GROUP_UNMERGER.MERGE_START_ROW_CELL);
  if (startRowRange) {
    startRowRange.setValue("Start Row")
      .setFontFamily("Lexend").setFontSize(11)
      .setFontWeight("normal").setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const endRowRange = ss.getRangeByName(ROW_GROUP_UNMERGER.MERGE_END_ROW_CELL);
  if (endRowRange) {
    endRowRange.setValue("End Row")
      .setFontFamily("Lexend").setFontSize(11)
      .setFontWeight("normal").setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const enableFeatureRange = ss.getRangeByName(ROW_GROUP_UNMERGER.ENABLE_FEATURE_TICKBOX);
  if (enableFeatureRange) enableFeatureRange.setValue(false);

  const unmergeNowRange = ss.getRangeByName(ROW_GROUP_UNMERGER.UNMERGE_NOW_TICKBOX);
  if (unmergeNowRange) unmergeNowRange.setValue(false);

  const resetRange = ss.getRangeByName(ROW_GROUP_UNMERGER.RESET_TO_DEFAULT_TICKBOX);
  if (resetRange) resetRange.setValue(false);

  setFeatureStatusCell(ROW_GROUP_UNMERGER.FEATURE_STATUS_CELL, false);
}
