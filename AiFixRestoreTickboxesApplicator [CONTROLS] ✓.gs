function isAfrtaEnabled() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableFeatureRange = ss.getRangeByName(AI_FIX_RESTORE_TICKBOXES_APPLICATOR.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange) return false;
  return enableFeatureRange.getValue() === true;
}

function setAfrtaOperationResult(message, fontColor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(AI_FIX_RESTORE_TICKBOXES_APPLICATOR.OPERATION_RESULT_CELL);
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

function handleAfrtaEnableFeature(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableFeatureRange = ss.getRangeByName(AI_FIX_RESTORE_TICKBOXES_APPLICATOR.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange) return;
  if (row !== enableFeatureRange.getRow() || col !== enableFeatureRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true && e.value !== "FALSE" && e.value !== false)) return;

  const isEnabled = (e.value === "TRUE" || e.value === true);
  setFeatureStatusCell(AI_FIX_RESTORE_TICKBOXES_APPLICATOR.FEATURE_STATUS_CELL, isEnabled);

  const targetSheetRange = ss.getRangeByName(AI_FIX_RESTORE_TICKBOXES_APPLICATOR.TARGET_SHEET_CELL);
  if (isEnabled) {
    if (targetSheetRange) targetSheetRange.setValue("").setFontColor("#ffffff");
  } else {
    if (targetSheetRange) {
      targetSheetRange.setValue("DD-MM-YYYY")
        .setFontFamily("Lexend").setFontSize(11)
        .setFontWeight("normal").setFontStyle("normal")
        .setFontColor("#b7b7b7");
    }
  }
}

function handleAfrtaApplyNow(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const applyNowRange = ss.getRangeByName(AI_FIX_RESTORE_TICKBOXES_APPLICATOR.APPLY_NOW_TICKBOX);
  if (!applyNowRange) return;
  if (row !== applyNowRange.getRow() || col !== applyNowRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true)) return;

  if (!isAfrtaEnabled()) {
    setFeatureNotEnabledWarning(AI_FIX_RESTORE_TICKBOXES_APPLICATOR.FEATURE_STATUS_CELL);
    applyNowRange.setValue(false);
    return;
  }

  const targetSheetRange = ss.getRangeByName(AI_FIX_RESTORE_TICKBOXES_APPLICATOR.TARGET_SHEET_CELL);
  if (!targetSheetRange) { applyNowRange.setValue(false); return; }

  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") {
    setAfrtaOperationResult("⚠️ No target sheet specified", "#faab17");
    applyNowRange.setValue(false);
    return;
  }

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) {
    setAfrtaOperationResult("❌ Sheet not found: " + tabName, "#ff0000");
    applyNowRange.setValue(false);
    return;
  }

  refreshAiTickboxesForSheet(targetSheet);
  setAfrtaOperationResult("✅ Tickboxes applied: " + tabName, "#00ff00");
  applyNowRange.setValue(false);
}

function handleAfrtaResetToDefault(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resetRange = ss.getRangeByName(AI_FIX_RESTORE_TICKBOXES_APPLICATOR.RESET_TO_DEFAULT_TICKBOX);
  if (!resetRange) return;
  if (row !== resetRange.getRow() || col !== resetRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true)) return;

  initAfrta();
  resetRange.setValue(false);
}

function initAfrta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const resultRange = ss.getRangeByName(AI_FIX_RESTORE_TICKBOXES_APPLICATOR.OPERATION_RESULT_CELL);
  if (resultRange) {
    resultRange.setValue("✨ Apply Operation Result ✨")
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

  const targetSheetRange = ss.getRangeByName(AI_FIX_RESTORE_TICKBOXES_APPLICATOR.TARGET_SHEET_CELL);
  if (targetSheetRange) {
    targetSheetRange.setValue("DD-MM-YYYY")
      .setFontFamily("Lexend").setFontSize(11)
      .setFontWeight("normal").setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const enableFeatureRange = ss.getRangeByName(AI_FIX_RESTORE_TICKBOXES_APPLICATOR.ENABLE_FEATURE_TICKBOX);
  if (enableFeatureRange) enableFeatureRange.setValue(false);

  const applyNowRange = ss.getRangeByName(AI_FIX_RESTORE_TICKBOXES_APPLICATOR.APPLY_NOW_TICKBOX);
  if (applyNowRange) applyNowRange.setValue(false);

  const resetRange = ss.getRangeByName(AI_FIX_RESTORE_TICKBOXES_APPLICATOR.RESET_TO_DEFAULT_TICKBOX);
  if (resetRange) resetRange.setValue(false);

  setFeatureStatusCell(AI_FIX_RESTORE_TICKBOXES_APPLICATOR.FEATURE_STATUS_CELL, false);
}

function refreshAiTickboxesForSheet(sheet) {
  const hasNotesCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_ACTIVITY_NOTES.HAS_NOTES_COL);
  const fixColLetter = ACTIVITY_NOTES_AI_FIX_AND_RESTORE.TICKBOX_COL;
  const undoColLetter = ACTIVITY_NOTES_AI_FIX_AND_RESTORE.UNDO_TICKBOX_COL;
  const dataStartRow = 3;
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return;

  const numRows = lastRow - dataStartRow + 1;
  const hasNotesVals = sheet.getRange(dataStartRow, hasNotesCol, numRows, 1).getValues();

  const enabledFixA1 = [];
  const enabledUndoA1 = [];
  const disabledFixA1 = [];
  const disabledUndoA1 = [];

  for (let i = 0; i < numRows; i++) {
    const r = dataStartRow + i;
    const v = hasNotesVals[i][0];
    if (v === "Yes") {
      enabledFixA1.push(fixColLetter + r);
      enabledUndoA1.push(undoColLetter + r);
    } else if (v === "No" || v === "No Status") {
      disabledFixA1.push(fixColLetter + r);
      disabledUndoA1.push(undoColLetter + r);
    }
  }

  if (enabledFixA1.length > 0) {
    const list = sheet.getRangeList(enabledFixA1);
    list.insertCheckboxes();
    list.setFontColor("#34a853");
    list.setHorizontalAlignment("center");
  }

  if (enabledUndoA1.length > 0) {
    const list = sheet.getRangeList(enabledUndoA1);
    list.insertCheckboxes();
    list.setFontColor("#ea4335");
    list.setHorizontalAlignment("center");
  }

  if (disabledFixA1.length > 0) {
    const list = sheet.getRangeList(disabledFixA1);
    list.clearDataValidations();
    list.setValue("❌");
    list.setFontColor("#ff0000");
    list.setHorizontalAlignment("center");
  }

  if (disabledUndoA1.length > 0) {
    const list = sheet.getRangeList(disabledUndoA1);
    list.clearDataValidations();
    list.setValue("❌");
    list.setFontColor("#ff0000");
    list.setHorizontalAlignment("center");
  }
}
