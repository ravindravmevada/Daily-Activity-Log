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

function setRguStatusPlaceholder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ROW_GROUP_UNMERGER.OPERATION_RESULT_CELL);
  if (!resultRange) return;
  resultRange.setValue("✨ Row Group Unmerger Status ✨")
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

function setRguTargetSheetPlaceholder(range) {
  range.setValue("DD-MM-YYYY")
    .setFontFamily("Lexend").setFontSize(11)
    .setFontWeight("normal").setFontStyle("normal")
    .setFontColor("#b7b7b7");
}

function setRguRowPlaceholder(range) {
  range.setValue("—")
    .setFontFamily("Lexend").setFontSize(11)
    .setFontWeight("normal").setFontStyle("normal")
    .setFontColor("#b7b7b7");
}

function setRguRichHint(foundTabName, hintMessage) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ROW_GROUP_UNMERGER.OPERATION_RESULT_CELL);
  if (!resultRange) return;

  const line1 = "✅ Target sheet found: " + foundTabName;
  const line2 = "ℹ️ " + hintMessage;
  const fullText = line1 + "\n" + line2;

  const richText = SpreadsheetApp.newRichTextValue()
    .setText(fullText)
    .setTextStyle(0, line1.length, SpreadsheetApp.newTextStyle()
      .setFontFamily("Lexend").setFontSize(11).setBold(true).setItalic(false)
      .setForegroundColor("#00ff00").build())
    .setTextStyle(line1.length + 1, fullText.length, SpreadsheetApp.newTextStyle()
      .setFontFamily("Lexend").setFontSize(11).setBold(true).setItalic(false)
      .setForegroundColor("#faab17").build())
    .build();

  resultRange.setRichTextValue(richText);
  resultRange.setBackground("#434343")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
}

function updateRguHintFromState(ss, tabName) {
  const startVal = ss.getRangeByName(ROW_GROUP_UNMERGER.MERGE_START_ROW_CELL).getValue();
  const endVal = ss.getRangeByName(ROW_GROUP_UNMERGER.MERGE_END_ROW_CELL).getValue();

  const startRow = parseInt(startVal);
  const endRow = parseInt(endVal);

  const startMissing = !startVal || startVal === "—" || isNaN(startRow);
  const endMissing = !endVal || endVal === "—" || isNaN(endRow);

  if (startMissing && endMissing) {
    setRguRichHint(tabName, "Enter \"Merge Start Row\" and \"Merge End Row\" to continue");
    return;
  }
  if (startMissing) {
    setRguRichHint(tabName, "Enter \"Merge Start Row\" to continue");
    return;
  }
  if (endMissing) {
    setRguRichHint(tabName, "Enter \"Merge End Row\" to continue");
    return;
  }
  if (startRow > endRow) {
    setRguRichHint(tabName, "\"Merge Start Row\" must be less than or equal to \"Merge End Row\"");
    return;
  }

  setRguRichHint(tabName, "Tick \"Unmerge Now\" to unmerge rows " + startRow + "–" + endRow);
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

  if (!isEnabled) {
    setRguStatusPlaceholder();
    if (targetSheetRange) setRguTargetSheetPlaceholder(targetSheetRange);
    if (startRowRange) setRguRowPlaceholder(startRowRange);
    if (endRowRange) setRguRowPlaceholder(endRowRange);
    return;
  }

  if (!targetSheetRange) return;
  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") return;

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) {
    setRguOperationResult("❌ Target sheet not found: " + tabName, "#ff0000");
    return;
  }

  updateRguHintFromState(ss, tabName);
}

function handleRguTargetSheet(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheetRange = ss.getRangeByName(ROW_GROUP_UNMERGER.TARGET_SHEET_CELL);
  if (!targetSheetRange) return;
  if (row !== targetSheetRange.getRow() || col !== targetSheetRange.getColumn()) return;

  if (!isRowGroupUnmergerEnabled()) {
    setFeatureNotEnabledWarning(ROW_GROUP_UNMERGER.FEATURE_STATUS_CELL);
    setRguTargetSheetPlaceholder(targetSheetRange);
    return;
  }

  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") {
    setRguStatusPlaceholder();
    return;
  }

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) {
    setRguOperationResult("❌ Target sheet not found: " + tabName, "#ff0000");
    return;
  }

  updateRguHintFromState(ss, tabName);
}

function handleRguStartRow(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startRowRange = ss.getRangeByName(ROW_GROUP_UNMERGER.MERGE_START_ROW_CELL);
  if (!startRowRange) return;
  if (row !== startRowRange.getRow() || col !== startRowRange.getColumn()) return;

  if (!isRowGroupUnmergerEnabled()) {
    setFeatureNotEnabledWarning(ROW_GROUP_UNMERGER.FEATURE_STATUS_CELL);
    setRguRowPlaceholder(startRowRange);
    return;
  }

  const targetSheetRange = ss.getRangeByName(ROW_GROUP_UNMERGER.TARGET_SHEET_CELL);
  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") return;

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) return;

  updateRguHintFromState(ss, tabName);
}

function handleRguEndRow(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const endRowRange = ss.getRangeByName(ROW_GROUP_UNMERGER.MERGE_END_ROW_CELL);
  if (!endRowRange) return;
  if (row !== endRowRange.getRow() || col !== endRowRange.getColumn()) return;

  if (!isRowGroupUnmergerEnabled()) {
    setFeatureNotEnabledWarning(ROW_GROUP_UNMERGER.FEATURE_STATUS_CELL);
    setRguRowPlaceholder(endRowRange);
    return;
  }

  const targetSheetRange = ss.getRangeByName(ROW_GROUP_UNMERGER.TARGET_SHEET_CELL);
  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") return;

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) return;

  updateRguHintFromState(ss, tabName);
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
    setRguOperationResult("⚠️ Enter valid \"Merge Start Row\" and \"Merge End Row\"", "#faab17");
    unmergeNowRange.setValue(false);
    return;
  }

  if (startRow > endRow) {
    setRguOperationResult("⚠️ \"Merge Start Row\" must be less than or equal to \"Merge End Row\"", "#faab17");
    unmergeNowRange.setValue(false);
    return;
  }

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) {
    setRguOperationResult("❌ Target sheet not found: " + tabName, "#ff0000");
    unmergeNowRange.setValue(false);
    return;
  }

  executeRowGroupUnmerge(targetSheet, startRow, endRow);
  setRguOperationResult("✅ Unmerged rows " + startRow + "–" + endRow + " in sheet: " + tabName, "#00ff00");
  unmergeNowRange.setValue(false);
}

function executeRowGroupUnmerge(sheet, startRow, endRow) {
  const numRows = endRow - startRow + 1;
  const colAToL = columnLetterToIndex(ROW_GROUP_UNMERGER.UNMERGED_ROW_LAST_COL);

  const categoryCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_START_COL);
  const subcatStartCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.UNMERGED_ROW_DROPDOWN_START_COL);
  const subcatEndCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.UNMERGED_ROW_DROPDOWN_END_COL);
  const subcatWidth = subcatEndCol - subcatStartCol + 1;
  const tickboxCols = DAILY_ACTIVITY_LOG_COLS.ROW_HIGHLIGHT_COLS.map(columnLetterToIndex);
  const perRowDurationCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.PER_ROW_DURATION_COL);
  const groupTotalDurationCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.GROUP_TOTAL_DURATION_COL);
  const numDurationCols = groupTotalDurationCol - perRowDurationCol + 1;

  sheet.getRange(startRow, 1, numRows, colAToL).breakApart();
  sheet.getRange(startRow, perRowDurationCol, numRows, numDurationCols).breakApart();

  const templateRowNum = endRow + 1 <= sheet.getLastRow() ? endRow + 1 : startRow - 1;

  for (let r = startRow; r <= endRow; r++) {
    tickboxCols.forEach(col => {
      if (col <= colAToL && (col < subcatStartCol || col > subcatEndCol)) {
        sheet.getRange(r, col).setValue(false);
      }
    });
    sheet.getRange(r, categoryCol).setValue("No Status");
  }

  if (templateRowNum >= 1) {
    const templateMerges = sheet.getRange(templateRowNum, 1, 1, colAToL).getMergedRanges();
    const templateSubRange = sheet.getRange(templateRowNum, subcatStartCol, 1, subcatWidth);
    for (let r = startRow; r <= endRow; r++) {
      templateMerges.forEach(mergedRange => {
        const mergeStartCol = mergedRange.getColumn();
        const numMergeCols = mergedRange.getNumColumns();
        if (numMergeCols > 1 && mergeStartCol < subcatStartCol) {
          sheet.getRange(r, mergeStartCol, 1, numMergeCols).mergeAcross();
        }
      });
      const targetSubRange = sheet.getRange(r, subcatStartCol, 1, subcatWidth);
      targetSubRange.breakApart();
      templateSubRange.copyTo(targetSubRange);
      sheet.getRange(r, subcatStartCol).setValue("No Status");
    }
  }

  for (let r = startRow; r <= endRow; r++) {
    sheet.getRange(r, perRowDurationCol, 1, numDurationCols).mergeAcross();
  }

  sheet.getRange(startRow, 1, numRows, colAToL)
    .setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(startRow, perRowDurationCol, numRows, numDurationCols)
    .setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);

  resequenceActivityNumbers(sheet);

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

  const enableFeatureRange = ss.getRangeByName(ROW_GROUP_UNMERGER.ENABLE_FEATURE_TICKBOX);
  if (enableFeatureRange) enableFeatureRange.setValue(false);

  setFeatureStatusCell(ROW_GROUP_UNMERGER.FEATURE_STATUS_CELL, false);

  const targetSheetRange = ss.getRangeByName(ROW_GROUP_UNMERGER.TARGET_SHEET_CELL);
  if (targetSheetRange) setRguTargetSheetPlaceholder(targetSheetRange);

  const startRowRange = ss.getRangeByName(ROW_GROUP_UNMERGER.MERGE_START_ROW_CELL);
  if (startRowRange) setRguRowPlaceholder(startRowRange);

  const endRowRange = ss.getRangeByName(ROW_GROUP_UNMERGER.MERGE_END_ROW_CELL);
  if (endRowRange) setRguRowPlaceholder(endRowRange);

  const unmergeNowRange = ss.getRangeByName(ROW_GROUP_UNMERGER.UNMERGE_NOW_TICKBOX);
  if (unmergeNowRange) unmergeNowRange.setValue(false);

  setRguStatusPlaceholder();

  const resetRange = ss.getRangeByName(ROW_GROUP_UNMERGER.RESET_TO_DEFAULT_TICKBOX);
  if (resetRange) resetRange.setValue(false);
}
