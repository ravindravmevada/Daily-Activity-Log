function handleColumnWidthSync(controlsSheet, row, col) {
  const resizeCell = cellRefToRowCol(COLUMN_WIDTH_SYNC.RESIZE_NOW_TICKBOX);
  if (row !== resizeCell.row || col !== resizeCell.col) return;

  const tickValue = controlsSheet.getRange(resizeCell.row, resizeCell.col).getValue();
  if (tickValue !== true) return;

  const enableCell = cellRefToRowCol(COLUMN_WIDTH_SYNC.ENABLE_TICKBOX);
  const isEnabled = controlsSheet.getRange(enableCell.row, enableCell.col).getValue();
  if (isEnabled !== true) {
    controlsSheet.getRange(resizeCell.row, resizeCell.col).setValue(false);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const widthRow = COLUMN_WIDTH_SYNC.WIDTH_ROW;

  for (let i = 0; i < COLUMN_WIDTH_SYNC.TARGET_SHEETS.length; i++) {
    const sheetName = COLUMN_WIDTH_SYNC.TARGET_SHEETS[i];
    const targetSheet = ss.getSheetByName(sheetName);
    if (!targetSheet) continue;

    const lastCol = targetSheet.getLastColumn();
    if (lastCol < 1) continue;

    const widthValues = targetSheet.getRange(widthRow, 1, 1, lastCol).getValues()[0];

    for (let c = 0; c < widthValues.length; c++) {
      const w = widthValues[c];
      if (typeof w === "number" && w > 0) {
        targetSheet.setColumnWidth(c + 1, w);
      }
    }
  }

  controlsSheet.getRange(resizeCell.row, resizeCell.col).setValue(false);
}
