function isSheetVisibilityFeatureEnabled() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableFeatureRange = ss.getRangeByName(SHEET_VISIBILITY_MANAGER.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange) return false;
  return enableFeatureRange.getValue() === true;
}

function getSheetVisibilitySheetRows(controlsSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetListRange = ss.getRangeByName(SHEET_VISIBILITY_MANAGER.SHEET_LIST_CELL);
  const resetRange = ss.getRangeByName(SHEET_VISIBILITY_MANAGER.RESET_TO_DEFAULT_TICKBOX);
  if (!sheetListRange || !resetRange) return [];

  const startRow = sheetListRange.getRow() + 1;
  const endRow = resetRange.getRow() - 1;
  const entries = [];

  for (let r = startRow; r <= endRow; r += 2) {
    const sheetName = controlsSheet.getRange(r, 2).getValue();
    if (!sheetName || sheetName === "") continue;
    entries.push({
      sheetName: sheetName.toString().trim(),
      tickboxRow: r,
      resultRow: r + 1
    });
  }
  return entries;
}

function setSheetVisibilityRichResult(controlsSheet, resultRow, mainText, hintText, mainColor) {
  const resultRange = controlsSheet.getRange(resultRow, 2, 1, 4);
  const fullText = mainText + "\n" + hintText;
  const richText = SpreadsheetApp.newRichTextValue()
    .setText(fullText)
    .setTextStyle(0, mainText.length, SpreadsheetApp.newTextStyle()
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setBold(true)
      .setItalic(false)
      .setForegroundColor(mainColor)
      .build())
    .setTextStyle(mainText.length, fullText.length, SpreadsheetApp.newTextStyle()
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setBold(true)
      .setItalic(true)
      .setForegroundColor("#faab17")
      .build())
    .build();

  resultRange.setRichTextValue(richText);
  resultRange.setBackground("#434343")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
}

function setSheetVisibilityStatusPlaceholder(controlsSheet, resultRow) {
  const resultRange = controlsSheet.getRange(resultRow, 2, 1, 4);
  resultRange.setValue("✨ Sheet Visibility Status ✨")
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

function handleSheetVisibility(controlsSheet, row, col, e) {
  if (!e) return;
  const isVisible = (e.value === "TRUE" || e.value === true);
  const isHidden = (e.value === "FALSE" || e.value === false || e.value === "");
  if (!isVisible && !isHidden) return;

  if (!isSheetVisibilityFeatureEnabled()) {
    setFeatureNotEnabledWarning(SHEET_VISIBILITY_MANAGER.FEATURE_STATUS_CELL);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const controlsSheetObj = ss.getSheetByName(CONTROLS_SHEET_NAME);
  if (!controlsSheetObj) return;

  const entries = getSheetVisibilitySheetRows(controlsSheetObj);
  const matchedEntry = entries.find(entry => entry.tickboxRow === row && col === 5);
  if (!matchedEntry) return;

  const targetSheet = ss.getSheetByName(matchedEntry.sheetName);

  if (!targetSheet) {
    setSheetVisibilityRichResult(controlsSheetObj, matchedEntry.resultRow, "❌ Sheet not found: " + matchedEntry.sheetName, "", "#ff0000");
    return;
  }

  const linkCell = controlsSheetObj.getRange(matchedEntry.tickboxRow, 6);

  if (isVisible) {
    targetSheet.showSheet();
    const gid = targetSheet.getSheetId();
    const fullUrl = ss.getUrl() + "#gid=" + gid;
    const richText = SpreadsheetApp.newRichTextValue()
      .setText("Click Me")
      .setLinkUrl(fullUrl)
      .setTextStyle(SpreadsheetApp.newTextStyle()
        .setForegroundColor("#6d9eeb")
        .setBold(true)
        .setUnderline(false)
        .build())
      .build();
    linkCell.setRichTextValue(richText)
      .setBackground("#000000")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    setSheetVisibilityRichResult(controlsSheetObj, matchedEntry.resultRow, "✅ Sheet is now visible", "Untick the checkbox to hide this sheet again", "#00ff00");
  } else {
    targetSheet.hideSheet();
    linkCell.setValue("HIDDEN")
      .setFontColor("#ff0000")
      .setFontWeight("bold")
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setBackground("#000000")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    setSheetVisibilityRichResult(controlsSheetObj, matchedEntry.resultRow, "🚫 Sheet is now hidden", "Tick the checkbox to make this sheet visible again", "#ff0000");
  }
}

function handleSheetVisibilityEnableFeature(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableFeatureRange = ss.getRangeByName(SHEET_VISIBILITY_MANAGER.ENABLE_FEATURE_TICKBOX);
  if (!enableFeatureRange) return;
  if (row !== enableFeatureRange.getRow() || col !== enableFeatureRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true && e.value !== "FALSE" && e.value !== false)) return;

  const isEnabled = (e.value === "TRUE" || e.value === true);
  setFeatureStatusCell(SHEET_VISIBILITY_MANAGER.FEATURE_STATUS_CELL, isEnabled);
}

function handleSheetVisibilityResetToDefault(controlsSheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resetRange = ss.getRangeByName(SHEET_VISIBILITY_MANAGER.RESET_TO_DEFAULT_TICKBOX);
  if (!resetRange) return;
  if (row !== resetRange.getRow() || col !== resetRange.getColumn()) return;
  if (!e || (e.value !== "TRUE" && e.value !== true)) return;

  initSheetVisibilityDefaults();
  resetRange.setValue(false);
}

function initSheetVisibilityDefaults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const controlsSheetObj = ss.getSheetByName(CONTROLS_SHEET_NAME);
  if (!controlsSheetObj) return;

  const entries = getSheetVisibilitySheetRows(controlsSheetObj);

  entries.forEach(entry => {
    const targetSheet = ss.getSheetByName(entry.sheetName);
    if (targetSheet) targetSheet.hideSheet();

    controlsSheetObj.getRange(entry.tickboxRow, 5).setValue(false);

    const linkCell = controlsSheetObj.getRange(entry.tickboxRow, 6);
    linkCell.setValue("HIDDEN")
      .setFontColor("#ff0000")
      .setFontWeight("bold")
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setBackground("#000000")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    setSheetVisibilityStatusPlaceholder(controlsSheetObj, entry.resultRow);
  });

  const enableFeatureRange = ss.getRangeByName(SHEET_VISIBILITY_MANAGER.ENABLE_FEATURE_TICKBOX);
  if (enableFeatureRange) enableFeatureRange.setValue(false);

  const resetRange = ss.getRangeByName(SHEET_VISIBILITY_MANAGER.RESET_TO_DEFAULT_TICKBOX);
  if (resetRange) resetRange.setValue(false);

  setFeatureStatusCell(SHEET_VISIBILITY_MANAGER.FEATURE_STATUS_CELL, false);
}
