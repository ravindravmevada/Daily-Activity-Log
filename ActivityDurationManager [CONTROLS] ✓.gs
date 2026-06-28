function isActivityDurationManagerEnabled() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableFeatureRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange) return false;
  return enableFeatureRange.getValue() === true;
}

function setAdmRefreshStatus(message, fontColor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_STATUS_CELL);
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

function setAdmRebuildStatus(message, fontColor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_STATUS_CELL);
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

function setAdmStatusPlaceholder(namedRange, text) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getRangeByName(namedRange);
  if (!range) return;
  range.setValue(text)
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

function setAdmRichHint(statusCellKey, foundTabName, hintMessage) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(statusCellKey);
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

function checkMergeConflictsInRange(sheet, startRow, endRow) {
  const conflictRows = [];
  const checked = {};
  const groupCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.CREATE_GROUP_TRIGGER_COL);
  const durationGroupCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.GROUP_TOTAL_DURATION_COL);

  for (let row = startRow; row <= endRow; row++) {
    if (checked[row]) continue;

    const cell = sheet.getRange(row, groupCol);
    const mergedRanges = cell.getMergedRanges();

    if (!mergedRanges || mergedRanges.length === 0) {
      checked[row] = true;
      continue;
    }

    const m = mergedRanges[0];
    const grpStart = m.getRow();
    const grpEnd = grpStart + m.getNumRows() - 1;

    for (let r = grpStart; r <= grpEnd; r++) checked[r] = true;

    if (m.getNumRows() < 2) continue;

    const startsBeforeRange = grpStart < startRow;
    const endsAfterRange = grpEnd > endRow;

    if (startsBeforeRange || endsAfterRange) {
      conflictRows.push("rows " + grpStart + "–" + grpEnd);
      continue;
    }

    const durCell = sheet.getRange(grpStart, durationGroupCol);
    const durMerged = durCell.getMergedRanges();
    if (!durMerged || durMerged.length === 0) {
      conflictRows.push("rows " + grpStart + "–" + grpEnd);
      continue;
    }

    const dm = durMerged[0];
    const durStart = dm.getRow();
    const durEnd = durStart + dm.getNumRows() - 1;

    if (durStart !== grpStart || durEnd !== grpEnd) {
      conflictRows.push("rows " + grpStart + "–" + grpEnd);
    }
  }

  return conflictRows;
}

function updateAdmRefreshHintFromState(ss, tabName, targetSheet) {
  const scope = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_APPLY_TO_CELL).getValue().toString().trim();

  if (scope === "No Status") {
    setAdmRichHint(ACTIVITY_DURATION_MANAGER.REFRESH_STATUS_CELL, tabName, "Select \"Apply To\" scope to continue");
    return;
  }

  if (scope === "Whole Sheet") {
    setAdmRichHint(ACTIVITY_DURATION_MANAGER.REFRESH_STATUS_CELL, tabName, "Tick \"Refresh Now\" to run on whole sheet");
    return;
  }

  if (scope === "Custom Range") {
    const startVal = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_START_ROW_CELL).getValue();
    const endVal = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_END_ROW_CELL).getValue();

    const startRow = parseInt(startVal);
    const endRow = parseInt(endVal);

    const startMissing = !startVal || startVal === "—" || isNaN(startRow);
    const endMissing = !endVal || endVal === "—" || isNaN(endRow);

    if (startMissing && endMissing) {
      setAdmRichHint(ACTIVITY_DURATION_MANAGER.REFRESH_STATUS_CELL, tabName, "Enter \"Refresh Start Row\" and \"Refresh End Row\" to continue");
      return;
    }
    if (startMissing) {
      setAdmRichHint(ACTIVITY_DURATION_MANAGER.REFRESH_STATUS_CELL, tabName, "Enter \"Refresh Start Row\" to continue");
      return;
    }
    if (endMissing) {
      setAdmRichHint(ACTIVITY_DURATION_MANAGER.REFRESH_STATUS_CELL, tabName, "Enter \"Refresh End Row\" to continue");
      return;
    }

    const dataStartRow = ACTIVITY_DURATION_MANAGER.DATA_START_ROW;
    const lastRow = targetSheet.getLastRow();
    if (startRow < dataStartRow || endRow > lastRow || startRow > endRow) {
      setAdmRichHint(ACTIVITY_DURATION_MANAGER.REFRESH_STATUS_CELL, tabName, "Row range out of bounds: " + startRow + "–" + endRow);
      return;
    }

    const conflicts = checkMergeConflictsInRange(targetSheet, startRow, endRow);
    if (conflicts.length > 0) {
      setAdmRichHint(ACTIVITY_DURATION_MANAGER.REFRESH_STATUS_CELL, tabName, "Merge conflict detected: " + conflicts.join(", ") + " — run \"Full Duration Columns Rebuild\" instead");
      return;
    }

    setAdmRichHint(ACTIVITY_DURATION_MANAGER.REFRESH_STATUS_CELL, tabName, "Tick \"Refresh Now\" to run on rows " + startRow + "–" + endRow);
  }
}

function updateAdmRebuildHintFromState(ss, tabName, targetSheet) {
  const scope = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_APPLY_TO_CELL).getValue().toString().trim();

  if (scope === "No Status") {
    setAdmRichHint(ACTIVITY_DURATION_MANAGER.REBUILD_STATUS_CELL, tabName, "Select \"Apply To\" scope to continue");
    return;
  }

  if (scope === "Whole Sheet") {
    setAdmRichHint(ACTIVITY_DURATION_MANAGER.REBUILD_STATUS_CELL, tabName, "Tick \"Rebuild Now\" to run on whole sheet");
    return;
  }

  if (scope === "Custom Range") {
    const startVal = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_START_ROW_CELL).getValue();
    const endVal = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_END_ROW_CELL).getValue();

    const startRow = parseInt(startVal);
    const endRow = parseInt(endVal);

    const startMissing = !startVal || startVal === "—" || isNaN(startRow);
    const endMissing = !endVal || endVal === "—" || isNaN(endRow);

    if (startMissing && endMissing) {
      setAdmRichHint(ACTIVITY_DURATION_MANAGER.REBUILD_STATUS_CELL, tabName, "Enter \"Rebuild Start Row\" and \"Rebuild End Row\" to continue");
      return;
    }
    if (startMissing) {
      setAdmRichHint(ACTIVITY_DURATION_MANAGER.REBUILD_STATUS_CELL, tabName, "Enter \"Rebuild Start Row\" to continue");
      return;
    }
    if (endMissing) {
      setAdmRichHint(ACTIVITY_DURATION_MANAGER.REBUILD_STATUS_CELL, tabName, "Enter \"Rebuild End Row\" to continue");
      return;
    }

    const dataStartRow = ACTIVITY_DURATION_MANAGER.DATA_START_ROW;
    const lastRow = targetSheet.getLastRow();
    if (startRow < dataStartRow || endRow > lastRow || startRow > endRow) {
      setAdmRichHint(ACTIVITY_DURATION_MANAGER.REBUILD_STATUS_CELL, tabName, "Row range out of bounds: " + startRow + "–" + endRow);
      return;
    }

    setAdmRichHint(ACTIVITY_DURATION_MANAGER.REBUILD_STATUS_CELL, tabName, "Tick \"Rebuild Now\" to run on rows " + startRow + "–" + endRow);
  }
}

function handleAdmRefreshTargetSheet(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_TARGET_SHEET_CELL);
  if (!targetSheetRange) return;
  if (row !== targetSheetRange.getRow() || col !== targetSheetRange.getColumn()) return;

  if (!isActivityDurationManagerEnabled()) {
    setFeatureNotEnabledWarning(ACTIVITY_DURATION_MANAGER.FEATURE_STATUS_CELL);
    targetSheetRange.setValue("DD-MM-YYYY")
      .setFontFamily("Lexend").setFontSize(11)
      .setFontWeight("normal").setFontStyle("normal")
      .setFontColor("#b7b7b7");
    return;
  }

  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") {
    setAdmStatusPlaceholder(ACTIVITY_DURATION_MANAGER.REFRESH_STATUS_CELL, "✨ Duration Refresh Status ✨");
    return;
  }

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) {
    setAdmRefreshStatus("❌ Target sheet not found: " + tabName, "#ff0000");
    return;
  }

  updateAdmRefreshHintFromState(ss, tabName, targetSheet);
}

function handleAdmRefreshApplyTo(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const applyToRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_APPLY_TO_CELL);
  if (!applyToRange) return;
  if (row !== applyToRange.getRow() || col !== applyToRange.getColumn()) return;

  if (!isActivityDurationManagerEnabled()) {
    setFeatureNotEnabledWarning(ACTIVITY_DURATION_MANAGER.FEATURE_STATUS_CELL);
    applyToRange.setValue("No Status");
    return;
  }

  const targetSheetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_TARGET_SHEET_CELL);
  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") return;

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) return;

  updateAdmRefreshHintFromState(ss, tabName, targetSheet);
}

function handleAdmRefreshStartRow(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startRowRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_START_ROW_CELL);
  if (!startRowRange) return;
  if (row !== startRowRange.getRow() || col !== startRowRange.getColumn()) return;

  if (!isActivityDurationManagerEnabled()) {
    setFeatureNotEnabledWarning(ACTIVITY_DURATION_MANAGER.FEATURE_STATUS_CELL);
    startRowRange.setValue("—");
    return;
  }

  const applyTo = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_APPLY_TO_CELL).getValue().toString().trim();
  if (applyTo === "Whole Sheet" || applyTo === "No Status") {
    startRowRange.setValue("—");
    return;
  }

  const targetSheetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_TARGET_SHEET_CELL);
  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") return;

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) return;

  updateAdmRefreshHintFromState(ss, tabName, targetSheet);
}

function handleAdmRefreshEndRow(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const endRowRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_END_ROW_CELL);
  if (!endRowRange) return;
  if (row !== endRowRange.getRow() || col !== endRowRange.getColumn()) return;

  if (!isActivityDurationManagerEnabled()) {
    setFeatureNotEnabledWarning(ACTIVITY_DURATION_MANAGER.FEATURE_STATUS_CELL);
    endRowRange.setValue("—");
    return;
  }

  const applyTo = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_APPLY_TO_CELL).getValue().toString().trim();
  if (applyTo === "Whole Sheet" || applyTo === "No Status") {
    endRowRange.setValue("—");
    return;
  }

  const targetSheetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_TARGET_SHEET_CELL);
  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") return;

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) return;

  updateAdmRefreshHintFromState(ss, tabName, targetSheet);
}

function handleAdmRebuildTargetSheet(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_TARGET_SHEET_CELL);
  if (!targetSheetRange) return;
  if (row !== targetSheetRange.getRow() || col !== targetSheetRange.getColumn()) return;

  if (!isActivityDurationManagerEnabled()) {
    setFeatureNotEnabledWarning(ACTIVITY_DURATION_MANAGER.FEATURE_STATUS_CELL);
    targetSheetRange.setValue("DD-MM-YYYY")
      .setFontFamily("Lexend").setFontSize(11)
      .setFontWeight("normal").setFontStyle("normal")
      .setFontColor("#b7b7b7");
    return;
  }

  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") {
    setAdmStatusPlaceholder(ACTIVITY_DURATION_MANAGER.REBUILD_STATUS_CELL, "✨ Duration Rebuild Status ✨");
    return;
  }

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) {
    setAdmRebuildStatus("❌ Target sheet not found: " + tabName, "#ff0000");
    return;
  }

  updateAdmRebuildHintFromState(ss, tabName, targetSheet);
}

function handleAdmRebuildApplyTo(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const applyToRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_APPLY_TO_CELL);
  if (!applyToRange) return;
  if (row !== applyToRange.getRow() || col !== applyToRange.getColumn()) return;

  if (!isActivityDurationManagerEnabled()) {
    setFeatureNotEnabledWarning(ACTIVITY_DURATION_MANAGER.FEATURE_STATUS_CELL);
    applyToRange.setValue("No Status");
    return;
  }

  const targetSheetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_TARGET_SHEET_CELL);
  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") return;

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) return;

  updateAdmRebuildHintFromState(ss, tabName, targetSheet);
}

function handleAdmRebuildStartRow(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startRowRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_START_ROW_CELL);
  if (!startRowRange) return;
  if (row !== startRowRange.getRow() || col !== startRowRange.getColumn()) return;

  if (!isActivityDurationManagerEnabled()) {
    setFeatureNotEnabledWarning(ACTIVITY_DURATION_MANAGER.FEATURE_STATUS_CELL);
    startRowRange.setValue("—");
    return;
  }

  const applyTo = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_APPLY_TO_CELL).getValue().toString().trim();
  if (applyTo === "Whole Sheet" || applyTo === "No Status") {
    startRowRange.setValue("—");
    return;
  }

  const targetSheetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_TARGET_SHEET_CELL);
  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") return;

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) return;

  updateAdmRebuildHintFromState(ss, tabName, targetSheet);
}

function handleAdmRebuildEndRow(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const endRowRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_END_ROW_CELL);
  if (!endRowRange) return;
  if (row !== endRowRange.getRow() || col !== endRowRange.getColumn()) return;

  if (!isActivityDurationManagerEnabled()) {
    setFeatureNotEnabledWarning(ACTIVITY_DURATION_MANAGER.FEATURE_STATUS_CELL);
    endRowRange.setValue("—");
    return;
  }

  const applyTo = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_APPLY_TO_CELL).getValue().toString().trim();
  if (applyTo === "Whole Sheet" || applyTo === "No Status") {
    endRowRange.setValue("—");
    return;
  }

  const targetSheetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_TARGET_SHEET_CELL);
  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") return;

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) return;

  updateAdmRebuildHintFromState(ss, tabName, targetSheet);
}

function refreshDurationValuesForSheet(sheet, startRow, endRow) {
  const perRowCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.PER_ROW_DURATION_COL);
  const groupTotalCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.GROUP_TOTAL_DURATION_COL);
  const activityTypeCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.ACTIVITY_TYPE_COL);
  const parallelValue = DAILY_ACTIVITY_LOG_DURATION_COLS.PARALLEL_VALUE;
  const subRowFontColor = DAILY_ACTIVITY_LOG_DURATION_COLS.SUB_ROW_FONT_COLOR;
  const mergedResultFontColor = DAILY_ACTIVITY_LOG_DURATION_COLS.MERGED_RESULT_FONT_COLOR;
  const parallelFontColor = DAILY_ACTIVITY_LOG_DURATION_COLS.PARALLEL_FONT_COLOR;

  const processedGroupStarts = {};

  for (let row = startRow; row <= endRow; row++) {
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
      const grpStart = groupRange.startRow;
      if (processedGroupStarts[grpStart]) continue;
      processedGroupStarts[grpStart] = true;

      const numRows = groupRange.numRows;
      let totalMs = 0;
      let hasAnyDuration = false;
      let hasRed = false;
      let allGreen = true;

      for (let r = 0; r < numRows; r++) {
        const subRow = grpStart + r;
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

      writeCellWithStyle(sheet.getRange(grpStart, groupTotalCol, numRows, 1), afValue, mergedResultFontColor);
      row = grpStart + numRows - 1;
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

  if (!isEnabled) {
    setAdmStatusPlaceholder(ACTIVITY_DURATION_MANAGER.REFRESH_STATUS_CELL, "✨ Duration Refresh Status ✨");
    setAdmStatusPlaceholder(ACTIVITY_DURATION_MANAGER.REBUILD_STATUS_CELL, "✨ Duration Rebuild Status ✨");
    return;
  }

  const refreshTabName = normalizeTabName(ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_TARGET_SHEET_CELL).getValue());
  if (refreshTabName && refreshTabName !== "" && refreshTabName !== "DD-MM-YYYY") {
    const refreshSheet = ss.getSheetByName(refreshTabName);
    if (!refreshSheet) {
      setAdmRefreshStatus("❌ Target sheet not found: " + refreshTabName, "#ff0000");
    } else {
      updateAdmRefreshHintFromState(ss, refreshTabName, refreshSheet);
    }
  }

  const rebuildTabName = normalizeTabName(ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_TARGET_SHEET_CELL).getValue());
  if (rebuildTabName && rebuildTabName !== "" && rebuildTabName !== "DD-MM-YYYY") {
    const rebuildSheet = ss.getSheetByName(rebuildTabName);
    if (!rebuildSheet) {
      setAdmRebuildStatus("❌ Target sheet not found: " + rebuildTabName, "#ff0000");
    } else {
      updateAdmRebuildHintFromState(ss, rebuildTabName, rebuildSheet);
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

  const targetSheetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_TARGET_SHEET_CELL);
  if (!targetSheetRange) { refreshNowRange.setValue(false); return; }

  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") {
    setAdmRefreshStatus("⚠️ No target sheet specified", "#faab17");
    refreshNowRange.setValue(false);
    return;
  }

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) {
    setAdmRefreshStatus("❌ Target sheet not found: " + tabName, "#ff0000");
    refreshNowRange.setValue(false);
    return;
  }

  const scope = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_APPLY_TO_CELL).getValue().toString().trim();
  if (!scope || scope === "No Status") {
    setAdmRefreshStatus("⚠️ Select \"Apply To\" scope first", "#faab17");
    refreshNowRange.setValue(false);
    return;
  }

  let startRow, endRow;
  const dataStartRow = ACTIVITY_DURATION_MANAGER.DATA_START_ROW;
  const lastRow = targetSheet.getLastRow();

  if (scope === "Whole Sheet") {
    startRow = dataStartRow;
    endRow = lastRow;
  } else {
    const startVal = parseInt(ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_START_ROW_CELL).getValue());
    const endVal = parseInt(ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_END_ROW_CELL).getValue());

    if (isNaN(startVal) || isNaN(endVal)) {
      setAdmRefreshStatus("⚠️ Invalid \"Refresh Start Row\" or \"Refresh End Row\"", "#faab17");
      refreshNowRange.setValue(false);
      return;
    }
    if (startVal < dataStartRow || endVal > lastRow || startVal > endVal) {
      setAdmRefreshStatus("⚠️ Row range out of bounds: " + startVal + "–" + endVal, "#faab17");
      refreshNowRange.setValue(false);
      return;
    }

    const conflicts = checkMergeConflictsInRange(targetSheet, startVal, endVal);
    if (conflicts.length > 0) {
      setAdmRefreshStatus("⚠️ Merge conflict detected: " + conflicts.join(", ") + " — run \"Full Duration Columns Rebuild\" instead", "#faab17");
      refreshNowRange.setValue(false);
      return;
    }

    startRow = startVal;
    endRow = endVal;
  }

  refreshDurationValuesForSheet(targetSheet, startRow, endRow);

  if (scope === "Custom Range") {
    setAdmRefreshStatus("✅ Duration values refreshed for rows " + startRow + "–" + endRow + " in sheet: " + tabName, "#00ff00");
  } else {
    setAdmRefreshStatus("✅ Duration values refreshed for whole sheet: " + tabName, "#00ff00");
  }

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

  const targetSheetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_TARGET_SHEET_CELL);
  if (!targetSheetRange) { rebuildNowRange.setValue(false); return; }

  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") {
    setAdmRebuildStatus("⚠️ No target sheet specified", "#faab17");
    rebuildNowRange.setValue(false);
    return;
  }

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) {
    setAdmRebuildStatus("❌ Target sheet not found: " + tabName, "#ff0000");
    rebuildNowRange.setValue(false);
    return;
  }

  const scope = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_APPLY_TO_CELL).getValue().toString().trim();
  if (!scope || scope === "No Status") {
    setAdmRebuildStatus("⚠️ Select \"Apply To\" scope first", "#faab17");
    rebuildNowRange.setValue(false);
    return;
  }

  let startRow, endRow;
  const dataStartRow = ACTIVITY_DURATION_MANAGER.DATA_START_ROW;
  const lastRow = targetSheet.getLastRow();

  if (scope === "Whole Sheet") {
    startRow = dataStartRow;
    endRow = lastRow;
  } else {
    const startVal = parseInt(ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_START_ROW_CELL).getValue());
    const endVal = parseInt(ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_END_ROW_CELL).getValue());

    if (isNaN(startVal) || isNaN(endVal)) {
      setAdmRebuildStatus("⚠️ Invalid \"Rebuild Start Row\" or \"Rebuild End Row\"", "#faab17");
      rebuildNowRange.setValue(false);
      return;
    }
    if (startVal < dataStartRow || endVal > lastRow || startVal > endVal) {
      setAdmRebuildStatus("⚠️ Row range out of bounds: " + startVal + "–" + endVal, "#faab17");
      rebuildNowRange.setValue(false);
      return;
    }

    startRow = startVal;
    endRow = endVal;
  }

  const processedGroupStartRows = {};
  for (let r = startRow; r <= endRow; r++) {
    const groupRange = findGroupRange(targetSheet, r);
    if (groupRange !== null) {
      const key = groupRange.startRow;
      if (processedGroupStartRows[key]) continue;
      processedGroupStartRows[key] = true;
    }
    updateRowDuration(targetSheet, r);
  }

  if (scope === "Custom Range") {
    setAdmRebuildStatus("✅ Duration columns rebuilt for rows " + startRow + "–" + endRow + " in sheet: " + tabName, "#00ff00");
  } else {
    setAdmRebuildStatus("✅ Duration columns rebuilt for whole sheet: " + tabName, "#00ff00");
  }

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

  const enableFeatureRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.ENABLE_FEATURE_TICKBOX);
  if (enableFeatureRange) enableFeatureRange.setValue(false);

  setFeatureStatusCell(ACTIVITY_DURATION_MANAGER.FEATURE_STATUS_CELL, false);

  const refreshTargetSheet = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_TARGET_SHEET_CELL);
  if (refreshTargetSheet) {
    refreshTargetSheet.setValue("DD-MM-YYYY")
      .setFontFamily("Lexend").setFontSize(11)
      .setFontWeight("normal").setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const refreshApplyTo = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_APPLY_TO_CELL);
  if (refreshApplyTo) refreshApplyTo.setValue("No Status");

  const refreshStartRow = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_START_ROW_CELL);
  if (refreshStartRow) refreshStartRow.setValue("—");

  const refreshEndRow = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_END_ROW_CELL);
  if (refreshEndRow) refreshEndRow.setValue("—");

  const refreshNowRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REFRESH_NOW_TICKBOX);
  if (refreshNowRange) refreshNowRange.setValue(false);

  setAdmStatusPlaceholder(ACTIVITY_DURATION_MANAGER.REFRESH_STATUS_CELL, "✨ Duration Refresh Status ✨");

  const rebuildTargetSheet = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_TARGET_SHEET_CELL);
  if (rebuildTargetSheet) {
    rebuildTargetSheet.setValue("DD-MM-YYYY")
      .setFontFamily("Lexend").setFontSize(11)
      .setFontWeight("normal").setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const rebuildApplyTo = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_APPLY_TO_CELL);
  if (rebuildApplyTo) rebuildApplyTo.setValue("No Status");

  const rebuildStartRow = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_START_ROW_CELL);
  if (rebuildStartRow) rebuildStartRow.setValue("—");

  const rebuildEndRow = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_END_ROW_CELL);
  if (rebuildEndRow) rebuildEndRow.setValue("—");

  const rebuildNowRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.REBUILD_NOW_TICKBOX);
  if (rebuildNowRange) rebuildNowRange.setValue(false);

  setAdmStatusPlaceholder(ACTIVITY_DURATION_MANAGER.REBUILD_STATUS_CELL, "✨ Duration Rebuild Status ✨");

  const resetRange = ss.getRangeByName(ACTIVITY_DURATION_MANAGER.RESET_TO_DEFAULT_TICKBOX);
  if (resetRange) resetRange.setValue(false);
}

function checkRangeHasGroups(sheet, startRow, endRow) {
  const groupCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.CREATE_GROUP_TRIGGER_COL);
  const checked = {};

  for (let row = startRow; row <= endRow; row++) {
    if (checked[row]) continue;
    const cell = sheet.getRange(row, groupCol);
    const mergedRanges = cell.getMergedRanges();
    if (!mergedRanges || mergedRanges.length === 0) {
      checked[row] = true;
      continue;
    }
    const m = mergedRanges[0];
    if (m.getNumRows() >= 2) return true;
    checked[row] = true;
  }
  return false;
}
