function handleAutoRowHighlight(sheet, row, col) {
  const editedCell = sheet.getRange(row, col);
  const value = editedCell.getValue();
  const lastCol = sheet.getLastColumn();

  let startRow, endRow;

  const mergedRanges = editedCell.getMergedRanges();
  if (mergedRanges && mergedRanges.length > 0) {
    const merged = mergedRanges[0];
    startRow = merged.getRow();
    endRow = merged.getRow() + merged.getNumRows() - 1;
  } else {
    startRow = row;
    endRow = row;
  }

  const numRows = endRow - startRow + 1;
  const startCol = col + 1;
  const numCols = lastCol - startCol + 1;
  const targetRange = sheet.getRange(startRow, startCol, numRows, numCols);

  const properties = PropertiesService.getDocumentProperties();
  const storageKey = "bgColors_" + sheet.getName() + "_" + startRow + "_" + startCol + "_" + numRows + "_" + numCols;

  if (value === true) {
    const originalColors = targetRange.getBackgrounds();
    properties.setProperty(storageKey, JSON.stringify(originalColors));
    targetRange.setBackground(HIGHLIGHT_BG_COLOR);
  } else {
    const savedColors = properties.getProperty(storageKey);
    if (savedColors) {
      targetRange.setBackgrounds(JSON.parse(savedColors));
      properties.deleteProperty(storageKey);
    }
  }
}
