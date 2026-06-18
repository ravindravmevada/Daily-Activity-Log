function handleActivityIdBulk(controlsSheet, row, col) {
  const runCell = cellRefToRowCol(ACTIVITY_ID_BULK.RUN_TICKBOX);
  if (row !== runCell.row || col !== runCell.col) return;
  if (controlsSheet.getRange(runCell.row, runCell.col).getValue() !== true) return;

  const enableCell = cellRefToRowCol(ACTIVITY_ID_BULK.ENABLE_TICKBOX);
  if (controlsSheet.getRange(enableCell.row, enableCell.col).getValue() !== true) {
    controlsSheet.getRange(runCell.row, runCell.col).setValue(false);
    return;
  }

  const targetCell = cellRefToRowCol(ACTIVITY_ID_BULK.TARGET_SHEET_CELL);
  const targetSheetName = normalizeTabName(controlsSheet.getRange(targetCell.row, targetCell.col).getValue());
  const targetSheet = targetSheetName ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName) : null;
  if (!targetSheet) {
    controlsSheet.getRange(runCell.row, runCell.col).setValue(false);
    return;
  }

  const dataStartRow = ACTIVITY_ID_BULK.DATA_START_ROW;
  const lastRow = targetSheet.getLastRow();
  resequenceActivityIds(targetSheet);

  controlsSheet.getRange(runCell.row, runCell.col).setValue(false);
}
