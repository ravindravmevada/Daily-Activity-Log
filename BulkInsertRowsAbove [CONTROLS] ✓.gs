function handleBulkInsertEnableFeature(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableFeatureRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange) return;
  if (row !== enableFeatureRange.getRow() || col !== enableFeatureRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true && e.value !== "FALSE" && e.value !== false)) return;

  const isEnabled = (e.value === "TRUE" || e.value === true);
  setFeatureStatusCell(BULK_INSERT_ROWS_ABOVE.FEATURE_STATUS_CELL, isEnabled);

  if (!isEnabled) {
    const currentMaxRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.CURRENT_MAX_NUMBER_CELL);
    const insertRowsUpToRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.INSERT_ROWS_UP_TO_CELL);
    const statusRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.STATUS_CELL);

    if (currentMaxRange) {
      currentMaxRange.setValue("—")
        .setFontFamily("Lexend").setFontSize(11)
        .setFontWeight("normal").setFontStyle("normal")
        .setFontColor("#b7b7b7");
    }
    if (insertRowsUpToRange) {
      insertRowsUpToRange.setValue("Target Number")
        .setFontFamily("Lexend").setFontSize(11)
        .setFontWeight("normal").setFontStyle("normal")
        .setFontColor("#b7b7b7");
    }
    if (statusRange) {
      statusRange.setValue("✨ Bulk Insert Rows Above Status ✨")
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
  } else {
    fetchBulkInsertCurrentMaxNumber(controlsSheet);
  }
}

function handleBulkInsertRowsAbove(sheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const targetSheetRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.TARGET_SHEET_CELL);
  const insertNowRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.INSERT_NOW_TICKBOX);

  const isTargetSheetEdit = targetSheetRange && row === targetSheetRange.getRow() && col === targetSheetRange.getColumn();
  const isInsertNowEdit = insertNowRange && row === insertNowRange.getRow() && col === insertNowRange.getColumn();

  if (!isTargetSheetEdit && !isInsertNowEdit) return;

  const enableFeatureRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange || enableFeatureRange.getValue() !== true) {
    if (isInsertNowEdit && e && (e.value === "TRUE" || e.value === true)) {
      setFeatureNotEnabledWarning(BULK_INSERT_ROWS_ABOVE.FEATURE_STATUS_CELL);
      insertNowRange.setValue(false);
    }
    return;
  }

  if (isTargetSheetEdit) {
    fetchBulkInsertCurrentMaxNumber(sheet);
    return;
  }

  if (isInsertNowEdit) {
    if (!e || (e.value !== "TRUE" && e.value !== true)) return;
    handleBulkInsertNow(sheet, row, col, e);
    return;
  }
}

function fetchBulkInsertCurrentMaxNumber(controlsSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const enableFeatureRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange || enableFeatureRange.getValue() !== true) {
    setFeatureNotEnabledWarning(BULK_INSERT_ROWS_ABOVE.FEATURE_STATUS_CELL);
    return;
  }

  const targetSheetRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.TARGET_SHEET_CELL);
  const currentMaxRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.CURRENT_MAX_NUMBER_CELL);
  if (!targetSheetRange || !currentMaxRange) return;

  const tabName = normalizeTabName(targetSheetRange.getValue());
  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") {
    currentMaxRange.setValue("—");
    setBulkInsertStatus("✨ Bulk Insert Rows Above Status ✨", "#f3f3f3", true);
    return;
  }

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) {
    currentMaxRange.setValue("—");
    setBulkInsertSheetNotFoundStatus(tabName);
    return;
  }

  const topDataRow = BULK_INSERT_ROWS_ABOVE.TARGET_TAB_TOP_DATA_ROW;
  const incrementColIndex = columnLetterToIndex(BULK_INSERT_ROWS_ABOVE.TARGET_TAB_INCREMENT_COL);
  const maxNumber = targetSheet.getRange(topDataRow, incrementColIndex).getValue();
  currentMaxRange.setValue(maxNumber);
  setBulkInsertSheetFoundStatus(tabName, maxNumber, null);
}

function handleBulkInsertNow(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (e && e.value !== "TRUE" && e.value !== true) return;

  const enableFeatureRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange || enableFeatureRange.getValue() !== true) {
    setFeatureNotEnabledWarning(BULK_INSERT_ROWS_ABOVE.FEATURE_STATUS_CELL);
    const insertNowRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.INSERT_NOW_TICKBOX);
    if (insertNowRange) insertNowRange.setValue(false);
    return;
  }

  setFeatureStatusCell(BULK_INSERT_ROWS_ABOVE.FEATURE_STATUS_CELL, true);

  const targetSheetRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.TARGET_SHEET_CELL);
  const currentMaxRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.CURRENT_MAX_NUMBER_CELL);
  const insertRowsUpToRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.INSERT_ROWS_UP_TO_CELL);
  const insertNowRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.INSERT_NOW_TICKBOX);

  if (!targetSheetRange || !currentMaxRange || !insertRowsUpToRange || !insertNowRange) {
    if (insertNowRange) insertNowRange.setValue(false);
    return;
  }

  const tabName = normalizeTabName(targetSheetRange.getValue());
  const currentMax = currentMaxRange.getValue();
  const insertRowsUpTo = insertRowsUpToRange.getValue();

  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") {
    setBulkInsertStatus("⚠️ No target sheet specified", "#faab17");
    insertNowRange.setValue(false);
    return;
  }

  if (typeof currentMax !== "number") {
    setBulkInsertStatus("⚠️ Current max number not loaded", "#faab17");
    insertNowRange.setValue(false);
    return;
  }

  if (typeof insertRowsUpTo !== "number") {
    setBulkInsertStatus("⚠️ Insert Rows Up To value not set", "#faab17");
    insertNowRange.setValue(false);
    return;
  }

  if (insertRowsUpTo <= currentMax) {
    setBulkInsertStatus("⚠️ Insert Rows Up To must be greater than " + currentMax, "#faab17");
    insertNowRange.setValue(false);
    return;
  }

  const targetSheet = ss.getSheetByName(tabName);
  if (!targetSheet) {
    setBulkInsertSheetNotFoundStatus(tabName);
    insertNowRange.setValue(false);
    return;
  }

  executeBulkInsertRowsAbove(controlsSheet, targetSheet, currentMax, insertRowsUpTo);

  currentMaxRange.setValue(insertRowsUpTo);
  insertNowRange.setValue(false);
}

function executeBulkInsertRowsAbove(controlsSheet, targetSheet, currentMax, targetNumber) {
  const rowsToInsert = targetNumber - currentMax;
  const topDataRow = BULK_INSERT_ROWS_ABOVE.TARGET_TAB_TOP_DATA_ROW;
  const lastCol = targetSheet.getLastColumn();

  clearRowStatusHighlight(targetSheet);

  targetSheet.insertRowsBefore(topDataRow, rowsToInsert);

  const templateRow = topDataRow + rowsToInsert;
  const templateRange = targetSheet.getRange(templateRow, 1, 1, lastCol);
  const blockRange = targetSheet.getRange(topDataRow, 1, rowsToInsert, lastCol);

  templateRange.copyTo(blockRange, { contentsOnly: false });

  blockRange.breakApart();
  const templateMerges = templateRange.getMergedRanges();
  for (let j = 0; j < templateMerges.length; j++) {
    const m = templateMerges[j];
    targetSheet.getRange(topDataRow, m.getColumn(), rowsToInsert, m.getNumColumns()).mergeAcross();
  }

  resequenceActivityNumbers(targetSheet);

  handleRowStatusHighlight(targetSheet);

  setBulkInsertStatus("✅ Inserted " + rowsToInsert + (rowsToInsert === 1 ? " row" : " rows") + " above in " + targetSheet.getName(), "#00ff00");
}

function setBulkInsertStatus(message, fontColor, isPlaceholder) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statusRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.STATUS_CELL);
  if (!statusRange) return;

  if (isPlaceholder) {
    statusRange.setValue(message)
      .setBackground("#434343")
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setFontWeight("normal")
      .setFontStyle("normal")
      .setFontColor(fontColor)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);
  } else {
    statusRange.setValue(message)
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
}

function initBulkInsertRowsAboveDefaults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const enableFeatureRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.ENABLE_FEATURE_TICKBOX);
  if (enableFeatureRange) enableFeatureRange.setValue(false);

  setFeatureStatusCell(BULK_INSERT_ROWS_ABOVE.FEATURE_STATUS_CELL, false);

  const targetSheetRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.TARGET_SHEET_CELL);
  if (targetSheetRange) {
    targetSheetRange.setValue("DD-MM-YYYY")
      .setFontFamily("Lexend").setFontSize(11)
      .setFontWeight("normal").setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const currentMaxRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.CURRENT_MAX_NUMBER_CELL);
  if (currentMaxRange) {
    currentMaxRange.setValue("—")
      .setFontFamily("Lexend").setFontSize(11)
      .setFontWeight("normal").setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const insertRowsUpToRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.INSERT_ROWS_UP_TO_CELL);
  if (insertRowsUpToRange) {
    insertRowsUpToRange.setValue("Target Number")
      .setFontFamily("Lexend").setFontSize(11)
      .setFontWeight("normal").setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const insertNowRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.INSERT_NOW_TICKBOX);
  if (insertNowRange) insertNowRange.setValue(false);

  setBulkInsertStatus("✨ Bulk Insert Rows Above Status ✨", "#f3f3f3", true);

  const resetToDefaultRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.RESET_TO_DEFAULT_TICKBOX);
  if (resetToDefaultRange) resetToDefaultRange.setValue(false);

  setFeatureStatusCell(BULK_INSERT_ROWS_ABOVE.FEATURE_STATUS_CELL, false);
}

function handleBulkInsertResetToDefault(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resetRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.RESET_TO_DEFAULT_TICKBOX);
  if (!resetRange) return;
  if (row !== resetRange.getRow() || col !== resetRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true)) return;

  initBulkInsertRowsAboveDefaults();
  resetRange.setValue(false);
}

function writeBulkInsertDetailedStatus(statusRange, lines) {
  if (!statusRange) return;

  let fullText = "";
  const segments = [];

  for (let i = 0; i < lines.length; i++) {
    const prefix = i === 0 ? "" : "\n";
    const start = fullText.length + prefix.length;
    fullText += prefix + lines[i].text;
    const end = fullText.length;
    segments.push({ start: start, end: end, color: lines[i].color });
  }

  let builder = SpreadsheetApp.newRichTextValue().setText(fullText);

  segments.forEach(seg => {
    builder = builder.setTextStyle(seg.start, seg.end, SpreadsheetApp.newTextStyle()
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setBold(true)
      .setItalic(false)
      .setForegroundColor(seg.color)
      .build());
  });

  const richText = builder.build();

  statusRange.setRichTextValue(richText);
  statusRange.setBackground("#434343")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
}

function setBulkInsertSheetFoundStatus(tabName, maxNumber, insertRowsUpTo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statusRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.STATUS_CELL);

  const lines = [
    { text: "✅ Target sheet found: " + tabName, color: "#00ff00" },
    { text: "🔢 Current max number: " + maxNumber, color: "#18ffff" }
  ];

  if (typeof insertRowsUpTo === "number" && insertRowsUpTo > maxNumber) {
    const rowsToInsert = insertRowsUpTo - maxNumber;
    lines.push({ text: "➕ " + rowsToInsert + (rowsToInsert === 1 ? " row" : " rows") + " will be inserted (up to " + insertRowsUpTo + ")", color: "#18ffff" });
    lines.push({ text: "ℹ️ Tick \"Insert Now\" to insert these rows", color: "#faab17" });
  } else if (typeof insertRowsUpTo === "number" && insertRowsUpTo <= maxNumber) {
    lines.push({ text: "⚠️ \"Insert Rows Up To\" must be greater than " + maxNumber, color: "#faab17" });
  } else {
    lines.push({ text: "ℹ️ Enter a number in \"Insert Rows Up To\" to continue", color: "#faab17" });
  }

  writeBulkInsertDetailedStatus(statusRange, lines);
}

function handleBulkInsertUpToEdit(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const insertRowsUpToRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.INSERT_ROWS_UP_TO_CELL);
  if (!insertRowsUpToRange) return;
  if (row !== insertRowsUpToRange.getRow() || col !== insertRowsUpToRange.getColumn()) return;

  const enableFeatureRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange || enableFeatureRange.getValue() !== true) return;

  const targetSheetRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.TARGET_SHEET_CELL);
  const currentMaxRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.CURRENT_MAX_NUMBER_CELL);
  if (!targetSheetRange || !currentMaxRange) return;

  const tabName = normalizeTabName(targetSheetRange.getValue());
  const currentMax = currentMaxRange.getValue();

  if (!tabName || tabName === "" || tabName === "DD-MM-YYYY") return;
  if (typeof currentMax !== "number") return;

  const insertRowsUpToValue = insertRowsUpToRange.getValue();
  const insertRowsUpTo = typeof insertRowsUpToValue === "number" ? insertRowsUpToValue : null;
  setBulkInsertSheetFoundStatus(tabName, currentMax, insertRowsUpTo);
}

function setBulkInsertSheetNotFoundStatus(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statusRange = ss.getRangeByName(BULK_INSERT_ROWS_ABOVE.STATUS_CELL);

  const lines = [
    { text: "❌ Target sheet not found: " + tabName, color: "#ff0000" },
    { text: "ℹ️ Check the sheet name and try again", color: "#faab17" }
  ];

  writeBulkInsertDetailedStatus(statusRange, lines);
}
