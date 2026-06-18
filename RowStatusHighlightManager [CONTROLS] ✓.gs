function sheetCountText(count) {
  return count + (count === 1 ? " sheet" : " sheets");
}

function isRowStatusFeatureEnabled() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableFeatureRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange) return false;
  return enableFeatureRange.getValue() === true;
}

function isRowStatusHighlightEnabled() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!isRowStatusFeatureEnabled()) return false;
  const enableHighlightingRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_TICKBOX);
  if (!enableHighlightingRange) return false;
  return enableHighlightingRange.getValue() === true;
}

function setHighlightOperationResult(message, fontColor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.HIGHLIGHT_OPERATION_RESULT_CELL);
  if (!resultRange) return;

  const lines = message.split("\n");
  const mainText = lines[0];
  const countText = lines.length > 1 ? "\n" + lines[1] : "";
  const fullText = mainText + countText;

  const richText = SpreadsheetApp.newRichTextValue()
    .setText(fullText)
    .setTextStyle(0, mainText.length, SpreadsheetApp.newTextStyle()
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setBold(true)
      .setItalic(false)
      .setForegroundColor(fontColor)
      .build())
    .setTextStyle(mainText.length, fullText.length, SpreadsheetApp.newTextStyle()
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setBold(true)
      .setItalic(false)
      .setForegroundColor("#faab17")
      .build())
    .build();

  resultRange.setRichTextValue(richText);
  resultRange.setBackground("#434343")
    .setFontStyle("normal")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
}

function setClearOperationResult(message, fontColor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.CLEAR_OPERATION_RESULT_CELL);
  if (!resultRange) return;

  const lines = message.split("\n");
  const mainText = lines[0];
  const countText = lines.length > 1 ? "\n" + lines[1] : "";
  const fullText = mainText + countText;

  const richText = SpreadsheetApp.newRichTextValue()
    .setText(fullText)
    .setTextStyle(0, mainText.length, SpreadsheetApp.newTextStyle()
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setBold(true)
      .setItalic(false)
      .setForegroundColor(fontColor)
      .build())
    .setTextStyle(mainText.length, fullText.length, SpreadsheetApp.newTextStyle()
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setBold(true)
      .setItalic(false)
      .setForegroundColor("#faab17")
      .build())
    .build();

  resultRange.setRichTextValue(richText);
  resultRange.setBackground("#434343")
    .setFontStyle("normal")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
}

function getHighlightSheetNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const highlightSheetsRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.HIGHLIGHT_ON_SHEETS_CELL);
  if (!highlightSheetsRange) return [];
  const rawValue = highlightSheetsRange.getValue();
  if (rawValue === "" || rawValue === null) return [];
  const normalizedValue = normalizeTabName(rawValue);
  if (normalizedValue === "" || normalizedValue === "DD-MM-YYYY") return [];
  if (normalizedValue.toUpperCase() === "ALL") return ["ALL"];
  if (normalizedValue.indexOf("\n") === -1) return [normalizedValue].filter(s => s !== "" && s !== "DD-MM-YYYY");
  return normalizedValue.split("\n").map(s => normalizeTabName(s.trim())).filter(s => s !== "" && s !== "DD-MM-YYYY");
}

function getClearSheetNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clearSheetsRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.CLEAR_FROM_SHEETS_CELL);
  if (!clearSheetsRange) return [];
  const rawValue = clearSheetsRange.getValue();
  if (rawValue === "" || rawValue === null) return [];
  const normalizedValue = normalizeTabName(rawValue);
  if (normalizedValue === "" || normalizedValue === "DD-MM-YYYY") return [];
  if (normalizedValue.toUpperCase() === "ALL") return ["ALL"];
  if (normalizedValue.indexOf("\n") === -1) return [normalizedValue].filter(s => s !== "" && s !== "DD-MM-YYYY");
  return normalizedValue.split("\n").map(s => normalizeTabName(s.trim())).filter(s => s !== "" && s !== "DD-MM-YYYY");
}

function isDateFormattedSheetName(sheetName) {
  return /^\d{2}-\d{2}-\d{4}$/.test(sheetName);
}

function isSheetInHighlightList(sheetName) {
  const highlightSheetNames = getHighlightSheetNames();
  if (highlightSheetNames.length === 0) return false;
  if (highlightSheetNames[0] === "ALL") return isDateFormattedSheetName(sheetName);
  return highlightSheetNames.some(name => name === sheetName);
}

function handleRowStatusHighlight(sheet) {
  if (!isRowStatusFeatureEnabled()) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableHighlightingRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_TICKBOX);
  if (!enableHighlightingRange || enableHighlightingRange.getValue() !== true) return;
  if (!isSheetInHighlightList(sheet.getName())) return;
  applyRowStatusHighlight(sheet);
}

function applyRowStatusHighlight(sheet) {
  const startTimestampCol = columnLetterToIndex(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ACTIVITY_START_TIMESTAMP_COL);
  const endTimestampCol = columnLetterToIndex(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ACTIVITY_END_TIMESTAMP_COL);
  const dataStartRow = ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.DATA_START_ROW;
  const normalColor = ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.NORMAL_BG_COLOR;
  const bothUntickedColor = ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.BOTH_UNTICKED_COLOR;
  const oneUntickedColor = ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ONE_UNTICKED_COLOR;
  const placeholderColor = ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.PLACEHOLDER_COLOR;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < dataStartRow) return;

  const numDataRows = lastRow - dataStartRow + 1;
  const startTimestampValues = sheet.getRange(dataStartRow, startTimestampCol, numDataRows, 1).getValues();
  const endTimestampValues = sheet.getRange(dataStartRow, endTimestampCol, numDataRows, 1).getValues();
  const allRowBackgrounds = sheet.getRange(dataStartRow, 1, numDataRows, lastCol).getBackgrounds();

  let firstTimestampedIndex = -1;
  for (let i = 0; i < numDataRows; i++) {
    const startState = isTimestampDone(startTimestampValues[i][0]);
    const endState = isTimestampDone(endTimestampValues[i][0]);
    if (startState === "DONE" || endState === "DONE") {
      firstTimestampedIndex = i;
      break;
    }
  }

  const blueRowIndex = (firstTimestampedIndex > 0) ? firstTimestampedIndex - 1 : -1;

  for (let i = numDataRows - 1; i >= 0; i--) {
    const actualRow = dataStartRow + i;
    const startState = isTimestampDone(startTimestampValues[i][0]);
    const endState = isTimestampDone(endTimestampValues[i][0]);

    const startUnticked = startState === "UNTICKED";
    const endUnticked = endState === "UNTICKED";
    const startDone = startState === "DONE";
    const endDone = endState === "DONE";

    let targetColor;

    if (startUnticked && endUnticked) {
      if (i === blueRowIndex) {
        targetColor = bothUntickedColor;
      } else if (firstTimestampedIndex !== -1 && i > firstTimestampedIndex) {
        targetColor = placeholderColor;
      } else {
        targetColor = null;
      }
    } else if ((startUnticked && endDone) || (startDone && endUnticked)) {
      targetColor = oneUntickedColor;
    } else {
      targetColor = null;
    }

    let currentStatusColor = "";
    for (let c = 0; c < lastCol; c++) {
      const cellBackground = (allRowBackgrounds[i][c] || "").toLowerCase();
      if (cellBackground === bothUntickedColor.toLowerCase() ||
          cellBackground === oneUntickedColor.toLowerCase() ||
          cellBackground === placeholderColor.toLowerCase()) {
        currentStatusColor = cellBackground;
        break;
      }
    }
    const rowHasStatusColor = (currentStatusColor !== "");

    if (targetColor !== null) {
      if (currentStatusColor !== targetColor.toLowerCase()) {
        sheet.getRange(actualRow, 1, 1, lastCol).setBackground(targetColor);
      }
    } else if (rowHasStatusColor) {
      sheet.getRange(actualRow, 1, 1, lastCol).setBackground(normalColor);
    }
  }
}

function clearRowStatusHighlight(sheet) {
  const dataStartRow = ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.DATA_START_ROW;
  const normalColor = ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.NORMAL_BG_COLOR;
  const bothUntickedColor = ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.BOTH_UNTICKED_COLOR.toLowerCase();
  const oneUntickedColor = ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ONE_UNTICKED_COLOR.toLowerCase();
  const placeholderColor = ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.PLACEHOLDER_COLOR.toLowerCase();

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < dataStartRow) return;

  const numDataRows = lastRow - dataStartRow + 1;
  const allRowBackgrounds = sheet.getRange(dataStartRow, 1, numDataRows, lastCol).getBackgrounds();

  for (let i = 0; i < numDataRows; i++) {
    let rowHasStatusColor = false;
    for (let c = 0; c < lastCol; c++) {
      const cellBackground = (allRowBackgrounds[i][c] || "").toLowerCase();
      if (cellBackground === bothUntickedColor ||
          cellBackground === oneUntickedColor ||
          cellBackground === placeholderColor) {
        rowHasStatusColor = true;
        break;
      }
    }
    if (rowHasStatusColor) {
      sheet.getRange(dataStartRow + i, 1, 1, lastCol).setBackground(normalColor);
    }
  }
}

function handleRowStatusEnableFeature(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableFeatureRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange) return;
  if (row !== enableFeatureRange.getRow() || col !== enableFeatureRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true && e.value !== "FALSE" && e.value !== false)) return;

  const isEnabled = (e.value === "TRUE" || e.value === true);
  setFeatureStatusCell(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.FEATURE_STATUS_CELL, isEnabled);

  if (!isEnabled) {
    const enableHighlightingRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_TICKBOX);
    if (enableHighlightingRange) enableHighlightingRange.setValue(false);
    setHighlightOperationResult("✨ Highlight Operation Result ✨", "#f3f3f3");
    setClearOperationResult("✨ Clear Operation Result ✨", "#f3f3f3");
  }
}

function handleEnableHighlighting(controlsSheet, row, col, e) {
  if (!isRowStatusFeatureEnabled()) {
    setFeatureNotEnabledWarning(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.FEATURE_STATUS_CELL);
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableHighlightingRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_TICKBOX);
  if (!enableHighlightingRange) return;
  if (row !== enableHighlightingRange.getRow() || col !== enableHighlightingRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true && e.value !== "FALSE" && e.value !== false)) return;

  const isEnabled = (e.value === "TRUE" || e.value === true);
  const highlightSheetNames = getHighlightSheetNames();

  if (highlightSheetNames.length === 0) {
    if (isEnabled) setHighlightOperationResult("⚠️ No sheets listed", "#faab17");
    return;
  }

  const notFoundSheets = [];
  const appliedSheets = [];

  if (highlightSheetNames[0] === "ALL") {
    ss.getSheets().filter(s => isDateFormattedSheetName(s.getName())).forEach(targetSheet => {
      if (isEnabled) applyRowStatusHighlight(targetSheet);
      else clearRowStatusHighlight(targetSheet);
      appliedSheets.push(targetSheet.getName());
    });
  } else {
    highlightSheetNames.forEach(sheetName => {
      const targetSheet = ss.getSheetByName(sheetName);
      if (!targetSheet) { notFoundSheets.push(sheetName); return; }
      if (isEnabled) applyRowStatusHighlight(targetSheet);
      else clearRowStatusHighlight(targetSheet);
      appliedSheets.push(sheetName);
    });
  }

  if (notFoundSheets.length > 0 && appliedSheets.length === 0) {
    setHighlightOperationResult("❌ Not found: " + notFoundSheets.join(", ") + "\n" + sheetCountText(notFoundSheets.length), "#ff0000");
  } else if (notFoundSheets.length > 0) {
    setHighlightOperationResult("⚠️ Applied: " + appliedSheets.join(", ") + "\n" + sheetCountText(appliedSheets.length) + " applied | Not found: " + notFoundSheets.join(", "), "#faab17");
  } else if (isEnabled) {
    setHighlightOperationResult("✅ Applied: " + appliedSheets.join(", ") + "\n" + sheetCountText(appliedSheets.length), "#00ff00");
  } else {
    setHighlightOperationResult("✅ Cleared: " + appliedSheets.join(", ") + "\n" + sheetCountText(appliedSheets.length), "#00ff00");
  }
}

function handleClearNow(controlsSheet, row, col, e) {
  if (!isRowStatusFeatureEnabled()) {
    setFeatureNotEnabledWarning(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.FEATURE_STATUS_CELL);
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clearNowRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.CLEAR_NOW_TICKBOX);
  if (!clearNowRange) return;
  if (row !== clearNowRange.getRow() || col !== clearNowRange.getColumn()) return;
  if (e && e.value !== "TRUE" && e.value !== true) return;

  const clearSheetNames = getClearSheetNames();
  if (clearSheetNames.length === 0) {
    setClearOperationResult("⚠️ No sheets listed", "#faab17");
    clearNowRange.setValue(false);
    return;
  }

  const notFoundSheets = [];
  const clearedSheets = [];

  if (clearSheetNames[0] === "ALL") {
    ss.getSheets().filter(s => isDateFormattedSheetName(s.getName())).forEach(targetSheet => {
      clearRowStatusHighlight(targetSheet);
      clearedSheets.push(targetSheet.getName());
    });
  } else {
    clearSheetNames.forEach(sheetName => {
      const targetSheet = ss.getSheetByName(sheetName);
      if (!targetSheet) { notFoundSheets.push(sheetName); return; }
      clearRowStatusHighlight(targetSheet);
      clearedSheets.push(sheetName);
    });
  }

  const highlightSheetsRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.HIGHLIGHT_ON_SHEETS_CELL);
  if (highlightSheetsRange) {
    if (clearSheetNames[0] === "ALL") {
      highlightSheetsRange.setValue("");
      const enableHighlightingRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_TICKBOX);
      if (enableHighlightingRange) enableHighlightingRange.setValue(false);
      setHighlightOperationResult("✨ Highlight Operation Result ✨", "#f3f3f3");
    } else {
      const currentHighlightNames = getHighlightSheetNames();
      const remainingHighlightNames = currentHighlightNames.filter(name => !clearSheetNames.includes(name));
      highlightSheetsRange.setValue(remainingHighlightNames.join("\n"));
      if (remainingHighlightNames.length === 0) {
        const enableHighlightingRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_TICKBOX);
        if (enableHighlightingRange) enableHighlightingRange.setValue(false);
        setHighlightOperationResult("✨ Highlight Operation Result ✨", "#f3f3f3");
      } else {
        setHighlightOperationResult("✅ Active: " + remainingHighlightNames.join(", ") + "\n" + sheetCountText(remainingHighlightNames.length), "#00ff00");
      }
    }
  }

  if (notFoundSheets.length > 0 && clearedSheets.length === 0) {
    setClearOperationResult("❌ Not found: " + notFoundSheets.join(", ") + "\n" + sheetCountText(notFoundSheets.length), "#ff0000");
  } else if (notFoundSheets.length > 0) {
    setClearOperationResult("⚠️ Cleared: " + clearedSheets.join(", ") + "\n" + sheetCountText(clearedSheets.length) + " cleared | Not found: " + notFoundSheets.join(", "), "#faab17");
  } else {
    setClearOperationResult("✅ Cleared: " + clearedSheets.join(", ") + "\n" + sheetCountText(clearedSheets.length), "#00ff00");
  }

  clearNowRange.setValue(false);
}

function handleHighlightOnSheetsEdit(controlsSheet, row, col, e) {
  if (!isRowStatusFeatureEnabled()) {
    setFeatureNotEnabledWarning(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.FEATURE_STATUS_CELL);
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const highlightSheetsRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.HIGHLIGHT_ON_SHEETS_CELL);
  if (!highlightSheetsRange) return;
  if (row !== highlightSheetsRange.getRow() || col !== highlightSheetsRange.getColumn()) return;

  const enableHighlightingRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_TICKBOX);
  if (!enableHighlightingRange || enableHighlightingRange.getValue() !== true) return;

  const currentSheetNames = getHighlightSheetNames();
  const notFoundSheets = [];
  const appliedSheets = [];
  const newlyAddedSheets = [];

  if (e && e.oldValue) {
    const oldNames = e.oldValue.toString().split("\n").map(s => normalizeTabName(s.trim())).filter(s => s !== "" && s !== "DD-MM-YYYY");
    const removedNames = oldNames.filter(name => !currentSheetNames.includes(name));
    removedNames.forEach(sheetName => {
      const targetSheet = ss.getSheetByName(sheetName);
      if (targetSheet) clearRowStatusHighlight(targetSheet);
    });
    currentSheetNames.forEach(sheetName => {
      if (!oldNames.includes(sheetName)) newlyAddedSheets.push(sheetName);
    });
  }

  if (currentSheetNames.length === 0) {
    setHighlightOperationResult("⚠️ No sheets listed", "#faab17");
    return;
  }

  const sheetsToProcess = newlyAddedSheets.length > 0 ? newlyAddedSheets : currentSheetNames;

  if (currentSheetNames[0] === "ALL") {
    ss.getSheets().filter(s => isDateFormattedSheetName(s.getName())).forEach(targetSheet => {
      applyRowStatusHighlight(targetSheet);
      appliedSheets.push(targetSheet.getName());
    });
  } else {
    sheetsToProcess.forEach(sheetName => {
      const targetSheet = ss.getSheetByName(sheetName);
      if (!targetSheet) { notFoundSheets.push(sheetName); return; }
      applyRowStatusHighlight(targetSheet);
      appliedSheets.push(sheetName);
    });
  }

  if (notFoundSheets.length > 0 && appliedSheets.length === 0) {
    setHighlightOperationResult("❌ Not found: " + notFoundSheets.join(", ") + "\n" + sheetCountText(notFoundSheets.length), "#ff0000");
  } else if (notFoundSheets.length > 0) {
    setHighlightOperationResult("⚠️ Applied: " + appliedSheets.join(", ") + "\n" + sheetCountText(appliedSheets.length) + " applied | Not found: " + notFoundSheets.join(", "), "#faab17");
  } else {
    setHighlightOperationResult("✅ Applied: " + appliedSheets.join(", ") + "\n" + sheetCountText(appliedSheets.length), "#00ff00");
  }
}

function handleClearFromSheetsEdit(controlsSheet, row, col, e) {
  if (!isRowStatusFeatureEnabled()) {
    setFeatureNotEnabledWarning(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.FEATURE_STATUS_CELL);
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clearSheetsRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.CLEAR_FROM_SHEETS_CELL);
  if (!clearSheetsRange) return;
  if (row !== clearSheetsRange.getRow() || col !== clearSheetsRange.getColumn()) return;

  const clearNowRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.CLEAR_NOW_TICKBOX);
  if (!clearNowRange || clearNowRange.getValue() !== true) return;

  const currentClearNames = getClearSheetNames();
  if (currentClearNames.length === 0) {
    setClearOperationResult("⚠️ No sheets listed", "#faab17");
    return;
  }

  const notFoundSheets = [];
  const clearedSheets = [];

  if (currentClearNames[0] === "ALL") {
    ss.getSheets().filter(s => isDateFormattedSheetName(s.getName())).forEach(targetSheet => {
      clearRowStatusHighlight(targetSheet);
      clearedSheets.push(targetSheet.getName());
    });
  } else {
    currentClearNames.forEach(sheetName => {
      const targetSheet = ss.getSheetByName(sheetName);
      if (!targetSheet) { notFoundSheets.push(sheetName); return; }
      clearRowStatusHighlight(targetSheet);
      clearedSheets.push(sheetName);
    });
  }

  if (notFoundSheets.length > 0 && clearedSheets.length === 0) {
    setClearOperationResult("❌ Not found: " + notFoundSheets.join(", ") + "\n" + sheetCountText(notFoundSheets.length), "#ff0000");
  } else if (notFoundSheets.length > 0) {
    setClearOperationResult("⚠️ Cleared: " + clearedSheets.join(", ") + "\n" + sheetCountText(clearedSheets.length) + " cleared | Not found: " + notFoundSheets.join(", "), "#faab17");
  } else {
    setClearOperationResult("✅ Cleared: " + clearedSheets.join(", ") + "\n" + sheetCountText(clearedSheets.length), "#00ff00");
  }
}

function handleResetToDefault(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resetRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.RESET_TO_DEFAULT_TICKBOX);
  if (!resetRange) return;
  if (row !== resetRange.getRow() || col !== resetRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true)) return;

  initOperationResultCells();
  resetRange.setValue(false);
}

function initOperationResultCells() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const highlightResult = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.HIGHLIGHT_OPERATION_RESULT_CELL);
  if (highlightResult) {
    highlightResult.setValue("✨ Highlight Operation Result ✨")
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

  const clearResult = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.CLEAR_OPERATION_RESULT_CELL);
  if (clearResult) {
    clearResult.setValue("✨ Clear Operation Result ✨")
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

  const highlightSheetsRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.HIGHLIGHT_ON_SHEETS_CELL);
  if (highlightSheetsRange) {
    highlightSheetsRange.setValue("DD-MM-YYYY")
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setFontWeight("normal")
      .setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const clearSheetsRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.CLEAR_FROM_SHEETS_CELL);
  if (clearSheetsRange) {
    clearSheetsRange.setValue("DD-MM-YYYY")
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setFontWeight("normal")
      .setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const enableFeatureRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_FEATURE_TICKBOX);
  if (enableFeatureRange) enableFeatureRange.setValue(false);

  const enableHighlightingRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_TICKBOX);
  if (enableHighlightingRange) enableHighlightingRange.setValue(false);

  const clearNowRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.CLEAR_NOW_TICKBOX);
  if (clearNowRange) clearNowRange.setValue(false);

  const resetToDefaultRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.RESET_TO_DEFAULT_TICKBOX);
  if (resetToDefaultRange) resetToDefaultRange.setValue(false);

  setFeatureStatusCell(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.FEATURE_STATUS_CELL, false);

  ss.getSheets().filter(s => isDateFormattedSheetName(s.getName())).forEach(targetSheet => {
    clearRowStatusHighlight(targetSheet);
  });
}

function isTimestampDone(value) {
  if (value === false) return "UNTICKED";
  if (value === "" || value === null) return "UNTICKED";
  if (value.toString().indexOf("Not Applicable") !== -1) return "NA";
  return "DONE";
}
