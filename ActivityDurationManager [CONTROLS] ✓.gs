function isActivityDurationManagerEnabled() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableFeatureRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange) return false;
  return enableFeatureRange.getValue() === true;
}

function setAdmRefreshResult(message, fontColor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_OPERATION_RESULT_CELL);
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

function setAdmRebuildResult(message, fontColor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_OPERATION_RESULT_CELL);
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

function refreshDurationValuesForSheet(sheet) {
  const perRowCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.PER_ROW_DURATION_COL);
  const groupTotalCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.GROUP_TOTAL_DURATION_COL);
  const activityTypeCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.ACTIVITY_TYPE_COL);
  const parallelValue = DAILY_ACTIVITY_LOG_DURATION_COLS.PARALLEL_VALUE;
  const subRowFontColor = DAILY_ACTIVITY_LOG_DURATION_COLS.SUB_ROW_FONT_COLOR;
  const mergedResultFontColor = DAILY_ACTIVITY_LOG_DURATION_COLS.MERGED_RESULT_FONT_COLOR;
  const parallelFontColor = DAILY_ACTIVITY_LOG_DURATION_COLS.PARALLEL_FONT_COLOR;

  const dataStartRow = ACTIVITY_DURATION_MANAGER.DATA_START_ROW;
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return;

  const processedGroupStarts = {};

  for (let row = dataStartRow; row <= lastRow; row++) {
    const groupRange = findGroupRange(sheet, row);

    if (groupRange === null) {
      const display = computeRowDisplay(sheet, row);
      const isParallel = sheet.getRange(row, activityTypeCol).getValue() === parallelValue;
      const fontColor = isParallel ? parallelFontColor : mergedResultFontColor;

      const perRowCell = sheet.getRange(row, perRowCol);
      const mergedRanges = perRowCell.getMergedRanges();
      if (mergedRanges && mergedRanges.length > 0) {
        writeCellWithStyle(sheet.getRange(row, perRowCol, 1, 2), display.text, fontColor);
      } else {
        writeCellWithStyle(perRowCell, "🟡", mergedResultFontColor);
        writeCellWithStyle(sheet.getRange(row, groupTotalCol), "🟡", mergedResultFontColor);
      }
    } else {
      const startRow = groupRange.startRow;
      if (processedGroupStarts[startRow]) continue;
      processedGroupStarts[startRow] = true;

      const numRows = groupRange.numRows;
      let totalMs = 0;
      let hasAnyDuration = false;
      let hasRed = false;
      let allGreen = true;

      for (let r = 0; r < numRows; r++) {
        const subRow = startRow + r;
        const display = computeRowDisplay(sheet, subRow);
        const isParallel = sheet.getRange(subRow, activityTypeCol).getValue() === parallelValue;
        const rowFontColor = isParallel ? parallelFontColor : subRowFontColor;

        writeCellWithStyle(sheet.getRange(subRow, perRowCol), display.text, rowFontColor);

        if (display.category === "RED") hasRed = true;
        if (display.category !== "GREEN") allGreen = false;
        if (display.durMs !== null && !isParallel) {
          totalMs += display.durMs;
          hasAnyDuration = true;
        }
      }

      let afValue;
      if (hasAnyDuration) {
        afValue = formatDurationMs(totalMs);
      } else if (hasRed) {
        afValue = "🔴";
      } else if (allGreen) {
        afValue = "🟢";
      } else {
        afValue = "🟡";
      }

      writeCellWithStyle(sheet.getRange(startRow, groupTotalCol, numRows, 1), afValue, mergedResultFontColor);
      row = startRow + numRows - 1;
    }
  }
}

function handleAdmEnableFeature(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableFeatureRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange) return;
  if (row !== enableFeatureRange.getRow() || col !== enableFeatureRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true && e.value !== "FALSE" && e.value !== false)) return;

  const isEnabled = (e.value === "TRUE" || e.value === true);
  setFeatureStatusCell(ACTIVITY_DURATION_MANAGER.FEATURE_STATUS_CELL, isEnabled);

  const targetSheetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.TARGET_SHEET_CELL);
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

function handleAdmRefreshNow(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const refreshNowRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_NOW_TICKBOX);
  if (!refreshNowRange) return;
  if (row !== refreshNowRange.getRow() || col !== refreshNowRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true)) return;

  if (!isActivityDurationManagerEnabled()) {
    setFeatureNotEnabledWarning(ACTIVITY_DURATION_MANAGER.FEATURE_STATUS_CELL);
    refreshNowRange.setValue(false);
    return;
  }

  const targetSheetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.TARGET_SHEET_CELL);
  if (!targetSheetRange) { refreshNowRange.setValue(false); return; }

  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") {
    setAdmRefreshResult("⚠️ No target sheet specified", "#faab17");
    refreshNowRange.setValue(false);
    return;
  }

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) {
    setAdmRefreshResult("❌ Sheet not found: " + tabName, "#ff0000");
    refreshNowRange.setValue(false);
    return;
  }

  refreshDurationValuesForSheet(targetSheet);
  setAdmRefreshResult("✅ Duration values refreshed: " + tabName, "#00ff00");
  refreshNowRange.setValue(false);
}

function handleAdmRebuildNow(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rebuildNowRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_NOW_TICKBOX);
  if (!rebuildNowRange) return;
  if (row !== rebuildNowRange.getRow() || col !== rebuildNowRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true)) return;

  if (!isActivityDurationManagerEnabled()) {
    setFeatureNotEnabledWarning(ACTIVITY_DURATION_MANAGER.FEATURE_STATUS_CELL);
    rebuildNowRange.setValue(false);
    return;
  }

  const targetSheetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.TARGET_SHEET_CELL);
  if (!targetSheetRange) { rebuildNowRange.setValue(false); return; }

  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") {
    setAdmRebuildResult("⚠️ No target sheet specified", "#faab17");
    rebuildNowRange.setValue(false);
    return;
  }

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) {
    setAdmRebuildResult("❌ Sheet not found: " + tabName, "#ff0000");
    rebuildNowRange.setValue(false);
    return;
  }

  const dataStartRow = ACTIVITY_DURATION_MANAGER.DATA_START_ROW;
  const lastRow = targetSheet.getLastRow();
  if (lastRow >= dataStartRow) {
    const processedGroupStartRows = {};
    for (let r = dataStartRow; r <= lastRow; r++) {
      const groupRange = findGroupRange(targetSheet, r);
      if (groupRange !== null) {
        const key = groupRange.startRow;
        if (processedGroupStartRows[key]) continue;
        processedGroupStartRows[key] = true;
      }
      updateRowDuration(targetSheet, r);
    }
  }

  setAdmRebuildResult("✅ Duration columns rebuilt: " + tabName, "#00ff00");
  rebuildNowRange.setValue(false);
}

function handleAdmResetToDefault(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.RESET_TO_DEFAULT_TICKBOX);
  if (!resetRange) return;
  if (row !== resetRange.getRow() || col !== resetRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true)) return;

  initActivityDurationManager();
  resetRange.setValue(false);
}

function initActivityDurationManager() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const refreshResult = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_OPERATION_RESULT_CELL);
  if (refreshResult) {
    refreshResult.setValue("✨ Refresh Operation Result ✨")
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

  const rebuildResult = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_OPERATION_RESULT_CELL);
  if (rebuildResult) {
    rebuildResult.setValue("✨ Rebuild Operation Result ✨")
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

  const targetSheetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.TARGET_SHEET_CELL);
  if (targetSheetRange) {
    targetSheetRange.setValue("DD-MM-YYYY")
      .setFontFamily("Lexend").setFontSize(11)
      .setFontWeight("normal").setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const enableFeatureRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.ENABLE_FEATURE_TICKBOX);
  if (enableFeatureRange) enableFeatureRange.setValue(false);

  const refreshNowRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_NOW_TICKBOX);
  if (refreshNowRange) refreshNowRange.setValue(false);

  const rebuildNowRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_NOW_TICKBOX);
  if (rebuildNowRange) rebuildNowRange.setValue(false);

  const resetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.RESET_TO_DEFAULT_TICKBOX);
  if (resetRange) resetRange.setValue(false);

  setFeatureStatusCell(ACTIVITY_DURATION_MANAGER.FEATURE_STATUS_CELL, false);
}
