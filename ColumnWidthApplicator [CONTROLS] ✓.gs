function handleCwaEnableFeature(sheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableRange = ss.getRangeByName(COLUMN_WIDTH_APPLICATOR.ENABLE_FEATURE_TICKBOX);
  if (!enableRange || enableRange.getRow() !== row || enableRange.getColumn() !== col) return;
  const isEnabled = e.value === "TRUE";
  setFeatureStatusCell(COLUMN_WIDTH_APPLICATOR.FEATURE_STATUS_CELL, isEnabled);
}

function handleCwaApplyNow(sheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const applyRange = ss.getRangeByName(COLUMN_WIDTH_APPLICATOR.APPLY_NOW_TICKBOX);
  const enableRange = ss.getRangeByName(COLUMN_WIDTH_APPLICATOR.ENABLE_FEATURE_TICKBOX);
  if (!applyRange || applyRange.getRow() !== row || applyRange.getColumn() !== col) return;
  if (e.value !== "TRUE") return;
  const isEnabled = enableRange ? enableRange.getValue() : false;
  if (!isEnabled) {
    applyRange.setValue(false);
    setFeatureNotEnabledWarning(COLUMN_WIDTH_APPLICATOR.FEATURE_STATUS_CELL);
    return;
  }
  const widthRow = COLUMN_WIDTH_APPLICATOR.WIDTH_ROW;
  const synced = [];
  const skipped = [];
  for (let i = 0; i < COLUMN_WIDTH_APPLICATOR.TARGET_SHEETS.length; i++) {
    const sheetName = COLUMN_WIDTH_APPLICATOR.TARGET_SHEETS[i];
    const targetSheet = ss.getSheetByName(sheetName);
    if (!targetSheet) { skipped.push(sheetName); continue; }
    const lastCol = targetSheet.getLastColumn();
    if (lastCol < 1) { skipped.push(sheetName); continue; }
    const widthValues = targetSheet.getRange(widthRow, 1, 1, lastCol).getValues()[0];
    for (let c = 0; c < widthValues.length; c++) {
      const w = widthValues[c];
      if (typeof w === "number" && w > 0) targetSheet.setColumnWidth(c + 1, w);
    }
    synced.push(sheetName);
  }
  applyRange.setValue(false);
  const msg = synced.length > 0
    ? `Applied widths to ${synced.length} sheet(s)` + (skipped.length > 0 ? `, skipped ${skipped.length}` : "")
    : "No sheets found to apply";
  setFeatureStatusCell(COLUMN_WIDTH_APPLICATOR.OPERATION_RESULT_CELL, true);
  ss.getRangeByName(COLUMN_WIDTH_APPLICATOR.OPERATION_RESULT_CELL).setValue(msg);
}

function handleCwaResetToDefault(sheet, row, col, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resetRange = ss.getRangeByName(COLUMN_WIDTH_APPLICATOR.RESET_TO_DEFAULT_TICKBOX);
  const enableRange = ss.getRangeByName(COLUMN_WIDTH_APPLICATOR.ENABLE_FEATURE_TICKBOX);
  if (!resetRange || resetRange.getRow() !== row || resetRange.getColumn() !== col) return;
  if (e.value !== "TRUE") return;
  enableRange.setValue(false);
  resetRange.setValue(false);
  setFeatureStatusCell(COLUMN_WIDTH_APPLICATOR.FEATURE_STATUS_CELL, false);
  setFeatureStatusCell(COLUMN_WIDTH_APPLICATOR.OPERATION_RESULT_CELL, false);
}
