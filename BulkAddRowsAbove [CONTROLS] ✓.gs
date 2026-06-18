function handleBulkAddEnableFeature(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableFeatureRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange) return;
  if (row !== enableFeatureRange.getRow() || col !== enableFeatureRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true && e.value !== "FALSE" && e.value !== false)) return;

  const isEnabled = (e.value === "TRUE" || e.value === true);
  setFeatureStatusCell(BULK_ADD_ROWS_ABOVE.FEATURE_STATUS_CELL, isEnabled);

  const targetSheetRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.TARGET_SHEET_CELL);
  const addRowsUpToRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.ADD_ROWS_UP_TO_CELL);
  const currentMaxRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.CURRENT_MAX_NUMBER_CELL);

  if (isEnabled) {
    if (targetSheetRange) targetSheetRange.setValue("").setFontColor("#ffffff");
    if (addRowsUpToRange) addRowsUpToRange.setValue("").setFontColor("#ffffff");
    if (currentMaxRange) currentMaxRange.setValue("").setFontColor("#ffffff");
  } else {
    if (targetSheetRange) {
      targetSheetRange.setValue("DD-MM-YYYY")
        .setFontFamily("Lexend").setFontSize(11)
        .setFontWeight("normal").setFontStyle("normal")
        .setFontColor("#b7b7b7");
    }
    if (currentMaxRange) {
      currentMaxRange.setValue("—")
        .setFontFamily("Lexend").setFontSize(11)
        .setFontWeight("normal").setFontStyle("normal")
        .setFontColor("#b7b7b7");
    }
    if (addRowsUpToRange) {
      addRowsUpToRange.setValue("Target Number")
        .setFontFamily("Lexend").setFontSize(11)
        .setFontWeight("normal").setFontStyle("normal")
        .setFontColor("#b7b7b7");
    }
  }
}

function handleBulkAddRowsAbove(sheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const enableFeatureRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange || enableFeatureRange.getValue() !== true) {
    setFeatureNotEnabledWarning(BULK_ADD_ROWS_ABOVE.FEATURE_STATUS_CELL);
    return;
  }

  const targetSheetRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.TARGET_SHEET_CELL);
  const runNowRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.RUN_NOW_TICKBOX);

  if (targetSheetRange && row === targetSheetRange.getRow() && col === targetSheetRange.getColumn()) {
    fetchBulkAddCurrentMaxNumber(sheet);
    return;
  }

  if (runNowRange && row === runNowRange.getRow() && col === runNowRange.getColumn()) {
    if (!e || (e.value !== "TRUE" && e.value !== true)) return;
    handleBulkAddRowsAboveActivate(sheet, row, col, e);
    return;
  }
}

function fetchBulkAddCurrentMaxNumber(controlsSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const enableFeatureRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange || enableFeatureRange.getValue() !== true) {
    setFeatureNotEnabledWarning(BULK_ADD_ROWS_ABOVE.FEATURE_STATUS_CELL);
    return;
  }

  const targetSheetRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.TARGET_SHEET_CELL);
  const currentMaxRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.CURRENT_MAX_NUMBER_CELL);
  if (!targetSheetRange || !currentMaxRange) return;

  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") {
    currentMaxRange.setValue("—");
    setBulkAddOperationResult("✨ Bulk Add Rows Above Operation Result ✨", "#f3f3f3");
    return;
  }

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) {
    currentMaxRange.setValue("—");
    setBulkAddOperationResult("❌ Sheet not found: " + tabName, "#ff0000");
    return;
  }

  const topDataRow = BULK_ADD_ROWS_ABOVE.TARGET_TAB_TOP_DATA_ROW;
  const incrementColIndex = columnLetterToIndex(BULK_ADD_ROWS_ABOVE.TARGET_TAB_INCREMENT_COL);
  const maxNumber = targetSheet.getRange(topDataRow, incrementColIndex).getValue();
  currentMaxRange.setValue(maxNumber);
  setBulkAddOperationResult("✅ Sheet found: " + tabName + " | Max: " + maxNumber, "#00ff00");
}

function handleBulkAddRowsAboveActivate(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (e && e.value !== "TRUE" && e.value !== true) return;

  const enableFeatureRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange || enableFeatureRange.getValue() !== true) {
    setFeatureNotEnabledWarning(BULK_ADD_ROWS_ABOVE.FEATURE_STATUS_CELL);
    const runNowRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.RUN_NOW_TICKBOX);
    if (runNowRange) runNowRange.setValue(false);
    return;
  }

  setFeatureStatusCell(BULK_ADD_ROWS_ABOVE.FEATURE_STATUS_CELL, true);

  const targetSheetRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.TARGET_SHEET_CELL);
  const currentMaxRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.CURRENT_MAX_NUMBER_CELL);
  const addRowsUpToRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.ADD_ROWS_UP_TO_CELL);
  const runNowRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.RUN_NOW_TICKBOX);

  if (!targetSheetRange || !currentMaxRange || !addRowsUpToRange || !runNowRange) {
    if (runNowRange) runNowRange.setValue(false);
    return;
  }

  const tabName = normalizeTabName(targetSheetRange.getValue());
  const currentMax = currentMaxRange.getValue();
  const addRowsUpTo = addRowsUpToRange.getValue();

  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") {
    setBulkAddOperationResult("⚠️ No target sheet specified", "#faab17");
    runNowRange.setValue(false);
    return;
  }

  if (typeof currentMax !== "number") {
    setBulkAddOperationResult("⚠️ Current max number not loaded", "#faab17");
    runNowRange.setValue(false);
    return;
  }

  if (typeof addRowsUpTo !== "number") {
    setBulkAddOperationResult("⚠️ Add Rows Up To value not set", "#faab17");
    runNowRange.setValue(false);
    return;
  }

  if (addRowsUpTo <= currentMax) {
    setBulkAddOperationResult("⚠️ Add Rows Up To must be greater than " + currentMax, "#faab17");
    runNowRange.setValue(false);
    return;
  }

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) {
    setBulkAddOperationResult("❌ Sheet not found: " + tabName, "#ff0000");
    runNowRange.setValue(false);
    return;
  }

  executeBulkAddRowsAbove(controlsSheet, targetSheet, currentMax, addRowsUpTo);

  currentMaxRange.setValue(addRowsUpTo);
  runNowRange.setValue(false);
}

function executeBulkAddRowsAbove(controlsSheet, targetSheet, currentMax, targetNumber) {
  const rowsToAdd = targetNumber - currentMax;
  const topDataRow = BULK_ADD_ROWS_ABOVE.TARGET_TAB_TOP_DATA_ROW;
  const incrementColIndex = columnLetterToIndex(BULK_ADD_ROWS_ABOVE.TARGET_TAB_INCREMENT_COL);
  const lastCol = targetSheet.getLastColumn();

  clearRowStatusHighlight(targetSheet);

  targetSheet.insertRowsBefore(topDataRow, rowsToAdd);

  const templateRow = topDataRow + rowsToAdd;
  const templateRange = targetSheet.getRange(templateRow, 1, 1, lastCol);
  const blockRange = targetSheet.getRange(topDataRow, 1, rowsToAdd, lastCol);

  templateRange.copyTo(blockRange, { contentsOnly: false });

  blockRange.breakApart();
  const templateMerges = templateRange.getMergedRanges();
  for (let j = 0; j < templateMerges.length; j++) {
    const m = templateMerges[j];
    targetSheet.getRange(topDataRow, m.getColumn(), rowsToAdd, m.getNumColumns()).mergeAcross();
  }

  const numbers = [];
  for (let i = 0; i < rowsToAdd; i++) {
    numbers.push([targetNumber - i]);
  }
  targetSheet.getRange(topDataRow, incrementColIndex, rowsToAdd, 1).setValues(numbers);

  handleRowStatusHighlight(targetSheet);

  setBulkAddOperationResult("✅ Added " + rowsToAdd + (rowsToAdd === 1 ? " row" : " rows") + " to " + targetSheet.getName(), "#00ff00");
}

function setBulkAddOperationResult(message, fontColor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.OPERATION_RESULT_CELL);
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

function initBulkAddOperationResult() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const resultRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.OPERATION_RESULT_CELL);
  if (resultRange) {
    resultRange.setValue("✨ Bulk Add Rows Above Operation Result ✨")
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

  const targetSheetRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.TARGET_SHEET_CELL);
  if (targetSheetRange) {
    targetSheetRange.setValue("DD-MM-YYYY")
      .setFontFamily("Lexend").setFontSize(11)
      .setFontWeight("normal").setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const currentMaxRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.CURRENT_MAX_NUMBER_CELL);
  if (currentMaxRange) {
    currentMaxRange.setValue("—")
      .setFontFamily("Lexend").setFontSize(11)
      .setFontWeight("normal").setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const addRowsUpToRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.ADD_ROWS_UP_TO_CELL);
  if (addRowsUpToRange) {
    addRowsUpToRange.setValue("Target Number")
      .setFontFamily("Lexend").setFontSize(11)
      .setFontWeight("normal").setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const enableFeatureRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.ENABLE_FEATURE_TICKBOX);
  if (enableFeatureRange) enableFeatureRange.setValue(false);

  const runNowRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.RUN_NOW_TICKBOX);
  if (runNowRange) runNowRange.setValue(false);

  const resetToDefaultRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.RESET_TO_DEFAULT_TICKBOX);
  if (resetToDefaultRange) resetToDefaultRange.setValue(false);

  setFeatureStatusCell(BULK_ADD_ROWS_ABOVE.FEATURE_STATUS_CELL, false);
}

function handleBulkAddResetToDefault(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resetRange = ss.getRangeByName(BULK_ADD_ROWS_ABOVE.RESET_TO_DEFAULT_TICKBOX);
  if (!resetRange) return;
  if (row !== resetRange.getRow() || col !== resetRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true)) return;

  initBulkAddOperationResult();
  resetRange.setValue(false);
}
