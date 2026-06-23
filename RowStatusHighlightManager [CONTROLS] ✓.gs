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

function writeSimpleResultCell(resultRange, message, fontColor, isPlaceholder) {
  if (!resultRange) return;

  const lines = message.split("\n");
  const mainText = lines[0];
  const countText = lines.length > 1 ? "\n" + lines[1] : "";
  const fullText = mainText + countText;

  let builder = SpreadsheetApp.newRichTextValue()
    .setText(fullText)
    .setTextStyle(0, mainText.length, SpreadsheetApp.newTextStyle()
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setBold(!isPlaceholder)
      .setItalic(false)
      .setForegroundColor(fontColor)
      .build());

  if (fullText.length > mainText.length) {
    builder = builder.setTextStyle(mainText.length, fullText.length, SpreadsheetApp.newTextStyle()
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setBold(!isPlaceholder)
      .setItalic(false)
      .setForegroundColor("#faab17")
      .build());
  }

  const richText = builder.build();

  resultRange.setRichTextValue(richText);
  resultRange.setBackground("#434343")
    .setFontStyle("normal")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
}

function setEnableHighlightingStatus(message, fontColor, isPlaceholder) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_STATUS_CELL);
  writeSimpleResultCell(resultRange, message, fontColor, isPlaceholder === true);
}

function setDisableHighlightingStatus(message, fontColor, isPlaceholder) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.DISABLE_HIGHLIGHTING_STATUS_CELL);
  writeSimpleResultCell(resultRange, message, fontColor, isPlaceholder === true);
}

function getEnableHighlightingOnSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_ON_SHEETS_CELL);
  if (!range) return [];
  const rawValue = range.getValue();
  if (rawValue === "" || rawValue === null) return [];
  const normalizedValue = normalizeTabName(rawValue);
  if (normalizedValue === "" || normalizedValue === "DD-MM-YYYY / ALL") return [];
  if (normalizedValue.toUpperCase() === "ALL") return ["ALL"];
  if (normalizedValue.indexOf("\n") === -1) return [normalizedValue].filter(s => s !== "" && s !== "DD-MM-YYYY / ALL");
  return normalizedValue.split("\n").map(s => normalizeTabName(s.trim())).filter(s => s !== "" && s !== "DD-MM-YYYY / ALL");
}

function getDisableHighlightingOnSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.DISABLE_HIGHLIGHTING_ON_SHEETS_CELL);
  if (!range) return [];
  const rawValue = range.getValue();
  if (rawValue === "" || rawValue === null) return [];
  const normalizedValue = normalizeTabName(rawValue);
  if (normalizedValue === "" || normalizedValue === "DD-MM-YYYY / ALL") return [];
  if (normalizedValue.toUpperCase() === "ALL") return ["ALL"];
  if (normalizedValue.indexOf("\n") === -1) return [normalizedValue].filter(s => s !== "" && s !== "DD-MM-YYYY / ALL");
  return normalizedValue.split("\n").map(s => normalizeTabName(s.trim())).filter(s => s !== "" && s !== "DD-MM-YYYY / ALL");
}

function isDateFormattedSheetName(sheetName) {
  return /^\d{2}-\d{2}-\d{4}$/.test(sheetName);
}

function getAllDateFormattedSheetNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().filter(s => isDateFormattedSheetName(s.getName())).map(s => s.getName());
}

function isSheetInEnableHighlightingList(sheetName) {
  const sheetNames = getEnableHighlightingOnSheets();
  if (sheetNames.length === 0) return false;
  if (sheetNames[0] === "ALL") return isDateFormattedSheetName(sheetName);
  return sheetNames.some(name => name === sheetName);
}

function isEnableHighlightingSetToAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_ON_SHEETS_CELL);
  if (!range) return false;
  const rawValue = range.getValue();
  const normalizedValue = normalizeTabName(rawValue);
  return normalizedValue.toUpperCase() === "ALL";
}

function handleRowStatusHighlight(sheet) {
  if (!isRowStatusFeatureEnabled()) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableHighlightingRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_TICKBOX);
  if (!enableHighlightingRange || enableHighlightingRange.getValue() !== true) return;
  if (!isSheetInEnableHighlightingList(sheet.getName())) return;
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

function sheetHasAnyRowStatusHighlight(sheet) {
  const dataStartRow = ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.DATA_START_ROW;
  const bothUntickedColor = ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.BOTH_UNTICKED_COLOR.toLowerCase();
  const oneUntickedColor = ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ONE_UNTICKED_COLOR.toLowerCase();
  const placeholderColor = ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.PLACEHOLDER_COLOR.toLowerCase();

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < dataStartRow) return false;

  const numDataRows = lastRow - dataStartRow + 1;
  const allRowBackgrounds = sheet.getRange(dataStartRow, 1, numDataRows, lastCol).getBackgrounds();

  for (let i = 0; i < numDataRows; i++) {
    for (let c = 0; c < lastCol; c++) {
      const cellBackground = (allRowBackgrounds[i][c] || "").toLowerCase();
      if (cellBackground === bothUntickedColor ||
          cellBackground === oneUntickedColor ||
          cellBackground === placeholderColor) {
        return true;
      }
    }
  }
  return false;
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
    const wasHighlighting = enableHighlightingRange ? enableHighlightingRange.getValue() === true : false;

    if (wasHighlighting) {
      const sheetNames = getEnableHighlightingOnSheets();
      if (sheetNames.length > 0) {
        if (sheetNames[0] === "ALL") {
          ss.getSheets().filter(s => isDateFormattedSheetName(s.getName())).forEach(targetSheet => {
            clearRowStatusHighlight(targetSheet);
          });
        } else {
          sheetNames.forEach(sheetName => {
            const targetSheet = ss.getSheetByName(sheetName);
            if (targetSheet) clearRowStatusHighlight(targetSheet);
          });
        }
      }
    }

    if (enableHighlightingRange) enableHighlightingRange.setValue(false);
    setEnableHighlightingStatus("✨ Enable Highlighting Status ✨", "#f3f3f3", true);
    setDisableHighlightingStatus("✨ Disable Highlighting Status ✨", "#f3f3f3", true);

  } else {

    const enableSheetNames = getEnableHighlightingOnSheets();
    if (enableSheetNames.length > 0) {
      if (enableSheetNames[0] === "ALL") {
        setEnableHighlightingSheetFoundPreview(getAllDateFormattedSheetNames());
      } else {
        setEnableHighlightingSheetFoundPreview(enableSheetNames);
      }
    }

    const disableSheetNames = getDisableHighlightingOnSheets();
    if (disableSheetNames.length > 0) {
      if (disableSheetNames[0] === "ALL") {
        setDisableHighlightingSheetFoundPreview(getAllDateFormattedSheetNames());
      } else {
        setDisableHighlightingSheetFoundPreview(disableSheetNames);
      }
    }
  }
}

function handleEnableHighlighting(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableHighlightingRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_TICKBOX);
  if (!enableHighlightingRange) return;
  if (row !== enableHighlightingRange.getRow() || col !== enableHighlightingRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true && e.value !== "FALSE" && e.value !== false)) return;

  const isTicking = (e.value === "TRUE" || e.value === true);

  if (!isRowStatusFeatureEnabled()) {
    if (isTicking) {
      setFeatureNotEnabledWarning(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.FEATURE_STATUS_CELL);
      enableHighlightingRange.setValue(false);
    } else {
      setFeatureStatusCell(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.FEATURE_STATUS_CELL, false);
    }
    return;
  }

  const sheetNames = getEnableHighlightingOnSheets();

  if (!isTicking) {
    if (sheetNames.length > 0) {
      if (sheetNames[0] === "ALL") {
        ss.getSheets().filter(s => isDateFormattedSheetName(s.getName())).forEach(targetSheet => {
          clearRowStatusHighlight(targetSheet);
        });
        setEnableHighlightingSheetFoundPreview(getAllDateFormattedSheetNames());
      } else {
        sheetNames.forEach(sheetName => {
          const targetSheet = ss.getSheetByName(sheetName);
          if (targetSheet) clearRowStatusHighlight(targetSheet);
        });
        setEnableHighlightingSheetFoundPreview(sheetNames);
      }
    } else {
      setEnableHighlightingStatus("✨ Enable Highlighting Status ✨", "#f3f3f3", true);
    }
    return;
  }

  if (sheetNames.length === 0) {
    setEnableHighlightingStatus("⚠️ No sheets listed", "#faab17");
    enableHighlightingRange.setValue(false);
    return;
  }

  if (sheetNames[0] === "ALL") {
    const allSheets = getAllDateFormattedSheetNames();
    allSheets.forEach(sheetName => {
      const targetSheet = ss.getSheetByName(sheetName);
      if (targetSheet) applyRowStatusHighlight(targetSheet);
    });
    setEnableHighlightingAppliedResult(allSheets);
    setDisableHighlightingStatus("✨ Disable Highlighting Status ✨", "#f3f3f3", true);
    return;
  }

  const allInvalid = sheetNames.every(name => !ss.getSheetByName(name));
  if (allInvalid) {
    setEnableHighlightingNotFoundOnly(sheetNames);
    enableHighlightingRange.setValue(false);
    return;
  }

  const notFoundSheets = [];
  const appliedSheets = [];

  sheetNames.forEach(sheetName => {
    const targetSheet = ss.getSheetByName(sheetName);
    if (!targetSheet) { notFoundSheets.push(sheetName); return; }
    applyRowStatusHighlight(targetSheet);
    appliedSheets.push(sheetName);
  });

  if (notFoundSheets.length > 0) {
    setEnableHighlightingAppliedMixedResult(appliedSheets, notFoundSheets);
  } else {
    setEnableHighlightingAppliedResult(appliedSheets);
  }

  setDisableHighlightingStatus("✨ Disable Highlighting Status ✨", "#f3f3f3", true);
}

function handleEnableHighlightingOnSheetsEdit(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableHighlightingOnSheetsRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_ON_SHEETS_CELL);
  if (!enableHighlightingOnSheetsRange) return;
  if (row !== enableHighlightingOnSheetsRange.getRow() || col !== enableHighlightingOnSheetsRange.getColumn()) return;

  if (!isRowStatusFeatureEnabled()) {
    setFeatureNotEnabledWarning(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.FEATURE_STATUS_CELL);
    return;
  }

  const currentSheetNames = getEnableHighlightingOnSheets();
  const enableHighlightingRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_TICKBOX);
  const isHighlightingOn = enableHighlightingRange ? enableHighlightingRange.getValue() === true : false;

  if (!isHighlightingOn) {
    if (currentSheetNames.length === 0) {
      setEnableHighlightingStatus("✨ Enable Highlighting Status ✨", "#f3f3f3", true);
      setDisableHighlightingStatus("✨ Disable Highlighting Status ✨", "#f3f3f3", true);
      return;
    }
    if (currentSheetNames[0] === "ALL") {
      setEnableHighlightingSheetFoundPreview(getAllDateFormattedSheetNames());
      setDisableHighlightingStatus("✨ Disable Highlighting Status ✨", "#f3f3f3", true);
      return;
    }
    setEnableHighlightingSheetFoundPreview(currentSheetNames);
    setDisableHighlightingStatus("✨ Disable Highlighting Status ✨", "#f3f3f3", true);
    return;
  }

  const newlyAddedSheets = [];

  if (e && e.oldValue) {
    const oldNames = e.oldValue.toString().split("\n").map(s => normalizeTabName(s.trim())).filter(s => s !== "" && s !== "DD-MM-YYYY / ALL");
    const removedNames = oldNames.filter(name => !currentSheetNames.includes(name));
    removedNames.forEach(sheetName => {
      const targetSheet = ss.getSheetByName(sheetName);
      if (targetSheet) clearRowStatusHighlight(targetSheet);
    });
    currentSheetNames.forEach(sheetName => {
      if (!oldNames.includes(sheetName)) newlyAddedSheets.push(sheetName);
    });
  } else {
    newlyAddedSheets.push(...currentSheetNames);
  }

  if (currentSheetNames.length === 0) {
    setEnableHighlightingStatus("⚠️ No sheets listed", "#faab17");
    return;
  }

  if (currentSheetNames[0] === "ALL") {
    const allSheets = getAllDateFormattedSheetNames();
    allSheets.forEach(sheetName => {
      const targetSheet = ss.getSheetByName(sheetName);
      if (targetSheet) applyRowStatusHighlight(targetSheet);
    });
    setEnableHighlightingAppliedResult(allSheets);
    setDisableHighlightingStatus("✨ Disable Highlighting Status ✨", "#f3f3f3", true);
    return;
  }

  const sheetsToProcess = newlyAddedSheets.length > 0 ? newlyAddedSheets : currentSheetNames;
  sheetsToProcess.forEach(sheetName => {
    const targetSheet = ss.getSheetByName(sheetName);
    if (targetSheet) applyRowStatusHighlight(targetSheet);
  });

  const allCurrentApplied = currentSheetNames.filter(name => ss.getSheetByName(name));
  const allCurrentNotFound = currentSheetNames.filter(name => !ss.getSheetByName(name));

  if (allCurrentNotFound.length > 0 && allCurrentApplied.length > 0) {
    setEnableHighlightingAppliedMixedResult(allCurrentApplied, allCurrentNotFound);
  } else if (allCurrentNotFound.length > 0) {
    setEnableHighlightingNotFoundOnly(allCurrentNotFound);
    if (enableHighlightingRange) enableHighlightingRange.setValue(false);
    return;
  } else {
    setEnableHighlightingAppliedResult(allCurrentApplied);
  }

  setDisableHighlightingStatus("✨ Disable Highlighting Status ✨", "#f3f3f3", true);
}

function handleDisableHighlightingOnSheetsEdit(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const disableHighlightingOnSheetsRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.DISABLE_HIGHLIGHTING_ON_SHEETS_CELL);
  if (!disableHighlightingOnSheetsRange) return;
  if (row !== disableHighlightingOnSheetsRange.getRow() || col !== disableHighlightingOnSheetsRange.getColumn()) return;

  if (!isRowStatusFeatureEnabled()) {
    setFeatureNotEnabledWarning(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.FEATURE_STATUS_CELL);
    return;
  }

  const currentDisableNames = getDisableHighlightingOnSheets();

  if (isEnableHighlightingSetToAll() && currentDisableNames.length > 0 && currentDisableNames[0] !== "ALL") {
    setDisableAllModeRequiredWarning();
    return;
  }

  const disableNowRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.DISABLE_NOW_TICKBOX);
  const isDisableNowOn = disableNowRange ? disableNowRange.getValue() === true : false;

  if (!isDisableNowOn) {
    if (currentDisableNames.length === 0) {
      setDisableHighlightingStatus("✨ Disable Highlighting Status ✨", "#f3f3f3", true);
      return;
    }
    if (currentDisableNames[0] === "ALL") {
      setDisableHighlightingSheetFoundPreview(getAllDateFormattedSheetNames());
      return;
    }
    setDisableHighlightingSheetFoundPreview(currentDisableNames);
    return;
  }

  if (currentDisableNames.length === 0) {
    setDisableHighlightingStatus("⚠️ No sheets listed", "#faab17");
    return;
  }

  const notFoundSheets = [];
  const disabledSheets = [];
  const neverHighlightedSheets = [];

  if (currentDisableNames[0] === "ALL") {
    ss.getSheets().filter(s => isDateFormattedSheetName(s.getName())).forEach(targetSheet => {
      clearRowStatusHighlight(targetSheet);
      disabledSheets.push(targetSheet.getName());
    });
  } else {
    currentDisableNames.forEach(sheetName => {
      const targetSheet = ss.getSheetByName(sheetName);
      if (!targetSheet) { notFoundSheets.push(sheetName); return; }
      const hadHighlight = sheetHasAnyRowStatusHighlight(targetSheet);
      clearRowStatusHighlight(targetSheet);
      if (hadHighlight) {
        disabledSheets.push(sheetName);
      } else {
        neverHighlightedSheets.push(sheetName);
      }
    });
  }

  if (notFoundSheets.length > 0 || neverHighlightedSheets.length > 0) {
    setDisableHighlightingAppliedMixedResult(disabledSheets, notFoundSheets, neverHighlightedSheets);
  } else {
    setDisableHighlightingAppliedResult(disabledSheets);
  }
}

function handleDisableNow(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const disableNowRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.DISABLE_NOW_TICKBOX);
  if (!disableNowRange) return;
  if (row !== disableNowRange.getRow() || col !== disableNowRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true && e.value !== "FALSE" && e.value !== false)) return;

  const isTicking = (e.value === "TRUE" || e.value === true);

  if (!isRowStatusFeatureEnabled()) {
    if (isTicking) {
      setFeatureNotEnabledWarning(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.FEATURE_STATUS_CELL);
      disableNowRange.setValue(false);
    } else {
      setFeatureStatusCell(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.FEATURE_STATUS_CELL, false);
    }
    return;
  }

  if (!isTicking) return;

  const disableSheetNames = getDisableHighlightingOnSheets();
  if (disableSheetNames.length === 0) {
    setDisableHighlightingStatus("⚠️ No sheets listed", "#faab17");
    disableNowRange.setValue(false);
    return;
  }

  if (isEnableHighlightingSetToAll() && disableSheetNames[0] !== "ALL") {
    setDisableAllModeRequiredWarning();
    disableNowRange.setValue(false);
    return;
  }

  const allInvalid = disableSheetNames[0] !== "ALL" && disableSheetNames.every(name => !ss.getSheetByName(name));
  if (allInvalid) {
    setDisableNotFoundOnly(disableSheetNames);
    disableNowRange.setValue(false);
    return;
  }

  const notFoundSheets = [];
  const disabledSheets = [];
  const neverHighlightedSheets = [];

  if (disableSheetNames[0] === "ALL") {
    ss.getSheets().filter(s => isDateFormattedSheetName(s.getName())).forEach(targetSheet => {
      clearRowStatusHighlight(targetSheet);
      disabledSheets.push(targetSheet.getName());
    });
  } else {
    disableSheetNames.forEach(sheetName => {
      const targetSheet = ss.getSheetByName(sheetName);
      if (!targetSheet) { notFoundSheets.push(sheetName); return; }
      const hadHighlight = sheetHasAnyRowStatusHighlight(targetSheet);
      clearRowStatusHighlight(targetSheet);
      if (hadHighlight) {
        disabledSheets.push(sheetName);
      } else {
        neverHighlightedSheets.push(sheetName);
      }
    });
  }

  const enableHighlightingOnSheetsRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_ON_SHEETS_CELL);
  if (enableHighlightingOnSheetsRange) {
    if (disableSheetNames[0] === "ALL") {
      enableHighlightingOnSheetsRange.setValue("");
      const enableHighlightingRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_TICKBOX);
      if (enableHighlightingRange) enableHighlightingRange.setValue(false);
      setEnableHighlightingStatus("✨ Enable Highlighting Status ✨", "#f3f3f3", true);
    } else {
      const currentEnableNames = getEnableHighlightingOnSheets();
      const validEnableNames = currentEnableNames[0] === "ALL"
        ? getAllDateFormattedSheetNames()
        : currentEnableNames.filter(name => ss.getSheetByName(name));
      const remainingEnableNames = validEnableNames.filter(name => !disableSheetNames.includes(name));
      enableHighlightingOnSheetsRange.setValue(remainingEnableNames.join("\n"));
      if (remainingEnableNames.length === 0) {
        const enableHighlightingRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_TICKBOX);
        if (enableHighlightingRange) enableHighlightingRange.setValue(false);
        setEnableHighlightingStatus("✨ Enable Highlighting Status ✨", "#f3f3f3", true);
      } else {
        setEnableHighlightingAppliedResult(remainingEnableNames);
      }
    }
  }

  if (notFoundSheets.length > 0 || neverHighlightedSheets.length > 0) {
    setDisableHighlightingAppliedMixedResult(disabledSheets, notFoundSheets, neverHighlightedSheets);
  } else {
    setDisableHighlightingAppliedResult(disabledSheets);
  }

  disableNowRange.setValue(false);
}

function handleResetToDefault(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resetRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.RESET_TO_DEFAULT_TICKBOX);
  if (!resetRange) return;
  if (row !== resetRange.getRow() || col !== resetRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true)) return;

  initRowStatusHighlightManagerDefaults();
  resetRange.setValue(false);
}

function initRowStatusHighlightManagerDefaults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const enableFeatureRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_FEATURE_TICKBOX);
  if (enableFeatureRange) enableFeatureRange.setValue(false);

  setFeatureStatusCell(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.FEATURE_STATUS_CELL, false);

  const enableHighlightingOnSheetsRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_ON_SHEETS_CELL);
  if (enableHighlightingOnSheetsRange) {
    enableHighlightingOnSheetsRange.setValue("DD-MM-YYYY / ALL")
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setFontWeight("normal")
      .setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const enableHighlightingRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_TICKBOX);
  if (enableHighlightingRange) enableHighlightingRange.setValue(false);

  setEnableHighlightingStatus("✨ Enable Highlighting Status ✨", "#f3f3f3", true);

  const disableHighlightingOnSheetsRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.DISABLE_HIGHLIGHTING_ON_SHEETS_CELL);
  if (disableHighlightingOnSheetsRange) {
    disableHighlightingOnSheetsRange.setValue("DD-MM-YYYY / ALL")
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setFontWeight("normal")
      .setFontStyle("normal")
      .setFontColor("#b7b7b7");
  }

  const disableNowRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.DISABLE_NOW_TICKBOX);
  if (disableNowRange) disableNowRange.setValue(false);

  setDisableHighlightingStatus("✨ Disable Highlighting Status ✨", "#f3f3f3", true);

  const resetToDefaultRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.RESET_TO_DEFAULT_TICKBOX);
  if (resetToDefaultRange) resetToDefaultRange.setValue(false);

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

function writeRowStatusDetailedResult(resultRange, lineObjects) {
  if (!resultRange) return;

  let fullText = "";
  const segments = [];

  for (let i = 0; i < lineObjects.length; i++) {
    const prefix = i === 0 ? "" : "\n";
    const start = fullText.length + prefix.length;
    fullText += prefix + lineObjects[i].text;
    const end = fullText.length;
    segments.push({ start: start, end: end, color: lineObjects[i].color });
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

  resultRange.setRichTextValue(richText);
  resultRange.setBackground("#434343")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
}

function setEnableHighlightingSheetFoundPreview(sheetNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_STATUS_CELL);

  const foundSheets = [];
  const notFoundSheets = [];
  sheetNames.forEach(name => {
    if (ss.getSheetByName(name)) {
      foundSheets.push(name);
    } else {
      notFoundSheets.push(name);
    }
  });

  const lines = [];

  if (foundSheets.length > 0 && notFoundSheets.length === 0) {
    lines.push({ text: "✅ " + sheetCountText(foundSheets.length) + " found: " + foundSheets.join(", "), color: "#00ff00" });
    lines.push({ text: "ℹ️ Tick \"Enable Highlighting\" to apply", color: "#faab17" });
  } else if (foundSheets.length > 0 && notFoundSheets.length > 0) {
    lines.push({ text: "✅ " + sheetCountText(foundSheets.length) + " found: " + foundSheets.join(", "), color: "#00ff00" });
    lines.push({ text: "❌ " + sheetCountText(notFoundSheets.length) + " not found: " + notFoundSheets.join(", "), color: "#ff0000" });
    lines.push({ text: "ℹ️ Tick \"Enable Highlighting\" to apply", color: "#faab17" });
  } else {
    lines.push({ text: "❌ " + sheetCountText(notFoundSheets.length) + " not found: " + notFoundSheets.join(", "), color: "#ff0000" });
    lines.push({ text: "ℹ️ Check the sheet name and try again", color: "#faab17" });
  }

  writeRowStatusDetailedResult(resultRange, lines);
}

function setEnableHighlightingNotFoundOnly(notFoundSheets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_STATUS_CELL);
  const lines = [
    { text: "❌ " + sheetCountText(notFoundSheets.length) + " not found: " + notFoundSheets.join(", "), color: "#ff0000" },
    { text: "ℹ️ Check the sheet name and try again", color: "#faab17" }
  ];
  writeRowStatusDetailedResult(resultRange, lines);
}

function setEnableHighlightingAppliedResult(sheetNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_STATUS_CELL);
  const lines = [
    { text: "✅ Highlighting enabled on " + sheetCountText(sheetNames.length) + ": " + sheetNames.join(", "), color: "#00ff00" },
    { text: "ℹ️ Add more sheets for highlighting in a new line using Alt+Enter", color: "#faab17" }
  ];
  writeRowStatusDetailedResult(resultRange, lines);
}

function setEnableHighlightingAppliedMixedResult(appliedSheets, notFoundSheets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.ENABLE_HIGHLIGHTING_STATUS_CELL);
  const notFoundLabel = notFoundSheets.length === 1 ? "Sheet not found" : "Sheets not found";
  const lines = [
    { text: "✅ Highlighting enabled on " + sheetCountText(appliedSheets.length) + ": " + appliedSheets.join(", "), color: "#00ff00" },
    { text: "❌ Highlighting not enabled on " + sheetCountText(notFoundSheets.length) + ": " + notFoundSheets.join(", "), color: "#ff0000" },
    { text: "(" + notFoundLabel + ")", color: "#ff0000" },
    { text: "ℹ️ Add more sheets for highlighting in a new line using Alt+Enter", color: "#faab17" }
  ];
  writeRowStatusDetailedResult(resultRange, lines);
}

function setDisableHighlightingSheetFoundPreview(sheetNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.DISABLE_HIGHLIGHTING_STATUS_CELL);

  const foundSheets = [];
  const notFoundSheets = [];
  sheetNames.forEach(name => {
    if (ss.getSheetByName(name)) {
      foundSheets.push(name);
    } else {
      notFoundSheets.push(name);
    }
  });

  const lines = [];

  if (foundSheets.length > 0 && notFoundSheets.length === 0) {
    lines.push({ text: "✅ " + sheetCountText(foundSheets.length) + " found: " + foundSheets.join(", "), color: "#00ff00" });
    lines.push({ text: "ℹ️ Tick \"Disable Now\" to apply", color: "#faab17" });
    lines.push({ text: "ℹ️ Add more sheets to disable in a new line using Alt+Enter", color: "#faab17" });
  } else if (foundSheets.length > 0 && notFoundSheets.length > 0) {
    lines.push({ text: "✅ " + sheetCountText(foundSheets.length) + " found: " + foundSheets.join(", "), color: "#00ff00" });
    lines.push({ text: "❌ " + sheetCountText(notFoundSheets.length) + " not found: " + notFoundSheets.join(", "), color: "#ff0000" });
    lines.push({ text: "ℹ️ Tick \"Disable Now\" to apply", color: "#faab17" });
    lines.push({ text: "ℹ️ Add more sheets to disable in a new line using Alt+Enter", color: "#faab17" });
  } else {
    lines.push({ text: "❌ " + sheetCountText(notFoundSheets.length) + " not found: " + notFoundSheets.join(", "), color: "#ff0000" });
    lines.push({ text: "ℹ️ Check the sheet name and try again", color: "#faab17" });
  }

  writeRowStatusDetailedResult(resultRange, lines);
}

function setDisableNotFoundOnly(notFoundSheets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.DISABLE_HIGHLIGHTING_STATUS_CELL);
  const lines = [
    { text: "❌ " + sheetCountText(notFoundSheets.length) + " not found: " + notFoundSheets.join(", "), color: "#ff0000" },
    { text: "ℹ️ Check the sheet name and try again", color: "#faab17" }
  ];
  writeRowStatusDetailedResult(resultRange, lines);
}

function setDisableHighlightingAppliedResult(sheetNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.DISABLE_HIGHLIGHTING_STATUS_CELL);
  const lines = [
    { text: "✅ Highlighting disabled on " + sheetCountText(sheetNames.length) + ": " + sheetNames.join(", "), color: "#00ff00" }
  ];
  writeRowStatusDetailedResult(resultRange, lines);
}

function setDisableHighlightingAppliedMixedResult(disabledSheets, notFoundSheets, neverHighlightedSheets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.DISABLE_HIGHLIGHTING_STATUS_CELL);
  const lines = [];

  if (disabledSheets.length > 0) {
    lines.push({ text: "✅ Highlighting disabled on " + sheetCountText(disabledSheets.length) + ": " + disabledSheets.join(", "), color: "#00ff00" });
  }

  if (notFoundSheets && notFoundSheets.length > 0) {
    const notFoundLabel = notFoundSheets.length === 1 ? "Sheet not found" : "Sheets not found";
    lines.push({ text: "❌ Highlighting not disabled on " + sheetCountText(notFoundSheets.length) + ": " + notFoundSheets.join(", "), color: "#ff0000" });
    lines.push({ text: "(" + notFoundLabel + ")", color: "#ff0000" });
  }

  if (neverHighlightedSheets && neverHighlightedSheets.length > 0) {
    const label = neverHighlightedSheets.length === 1 ? "sheet had" : "sheets had";
    const reasonLabel = neverHighlightedSheets.length === 1 ? "Sheet was never highlighted" : "Sheets were never highlighted";
    lines.push({ text: "💭 " + neverHighlightedSheets.length + " " + label + " no highlighting to disable: " + neverHighlightedSheets.join(", "), color: "#18ffff" });
    lines.push({ text: "(" + reasonLabel + ")", color: "#18ffff" });
  }

  writeRowStatusDetailedResult(resultRange, lines);
}

function setDisableAllModeRequiredWarning() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultRange = ss.getRangeByName(ROW_STATUS_HIGHLIGHT_MANAGER_FEATURE.DISABLE_HIGHLIGHTING_STATUS_CELL);
  const lines = [
    { text: "⚠️ Highlighting has been applied to \"ALL\" sheets in \"Enable Highlighting On Sheets\". To use the disable operation, \"Disable Highlighting On Sheets\" must also be set to \"ALL\". Type \"ALL\" in \"Disable Highlighting On Sheets\".", color: "#faab17" }
  ];
  writeRowStatusDetailedResult(resultRange, lines);
}
