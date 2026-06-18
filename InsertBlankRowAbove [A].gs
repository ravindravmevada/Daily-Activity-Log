function handleInsertBlankRowAbove(sheet, row, col) {
  const editedCell = sheet.getRange(row, col);
  const value = editedCell.getValue();
  if (value !== true) return;

  editedCell.setValue(false);

  const lastCol = sheet.getLastColumn();
  const incrementColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INCREMENT_NUMBER);

  sheet.insertRowsBefore(row, 1);

  const sourceRange = sheet.getRange(row + 1, 1, 1, lastCol);
  const newRange = sheet.getRange(row, 1, 1, lastCol);

  sourceRange.copyTo(newRange, { contentsOnly: false });

  const sourceMerges = sourceRange.getMergedRanges();
  for (let i = 0; i < sourceMerges.length; i++) {
    const m = sourceMerges[i];
    const newMergeRange = sheet.getRange(row, m.getColumn(), 1, m.getNumColumns());
    newMergeRange.merge();
  }

  const belowNumber = sheet.getRange(row + 1, incrementColIndex).getValue();
  if (typeof belowNumber === "number") {
    sheet.getRange(row, incrementColIndex).setValue(belowNumber + 1);
  }
}
