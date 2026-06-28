function handleInsertMergedGroupRowAbove(sheet, row, col) {
  const editedCell = sheet.getRange(row, col);
  if (editedCell.getValue() !== true) return;
  editedCell.setValue(false);

  const lastCol = sheet.getLastColumn();
  const labelCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.MERGED_GROUP_LABEL_COL);
  const incCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.MERGED_GROUP_INCREMENT_NUMBER);
  const sharedCols = DAILY_ACTIVITY_LOG_COLS.GROUP_VERTICAL_MERGE_COLS.map(columnLetterToIndex);
  const groupTotalCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.GROUP_TOTAL_DURATION_COL);

  const labelMerges = sheet.getRange(row, labelCol).getMergedRanges();
  let groupStartRow, groupEndRow, labelStartCol, labelNumCols;
  if (labelMerges && labelMerges.length > 0) {
    const m = labelMerges[0];
    groupStartRow = m.getRow();
    groupEndRow = m.getRow() + m.getNumRows() - 1;
    labelStartCol = m.getColumn();
    labelNumCols = m.getNumColumns();
  } else {
    groupStartRow = row; groupEndRow = row; labelStartCol = labelCol; labelNumCols = 1;
  }

  const labelValue = sheet.getRange(groupStartRow, labelStartCol).getValue();

  let maxSub = 0;
  for (let r = groupStartRow; r <= groupEndRow; r++) {
    const v = sheet.getRange(r, incCol).getValue();
    if (typeof v === "number" && v > maxSub) maxSub = v;
  }
  const newSub = Math.round((maxSub + 0.1) * 10) / 10;

  const groupRows = groupEndRow - groupStartRow + 1;

  // break ALL of the group's vertical merges across the full group range, BEFORE inserting
  sheet.getRange(groupStartRow, labelStartCol, groupRows, labelNumCols).breakApart();
  for (let i = 0; i < sharedCols.length; i++) {
    sheet.getRange(groupStartRow, sharedCols[i], groupRows, 1).breakApart();
  }
  sheet.getRange(groupStartRow, groupTotalCol, groupRows, 1).breakApart(); // <-- AH total

  sheet.insertRowsBefore(row, 1);

  const source = sheet.getRange(row + 1, 1, 1, lastCol);
  const dest = sheet.getRange(row, 1, 1, lastCol);
  source.copyTo(dest, { contentsOnly: false });

  const groupEndColForSkip = labelStartCol + labelNumCols - 1;
  const srcMerges = source.getMergedRanges();
  for (let i = 0; i < srcMerges.length; i++) {
    const m = srcMerges[i];
    if (m.getNumRows() !== 1) continue;
    const sCol = m.getColumn();
    const eCol = sCol + m.getNumColumns() - 1;
    if (!(eCol < labelStartCol || sCol > groupEndColForSkip)) continue;
    if (sharedCols.some(idx => idx >= sCol && idx <= eCol)) continue;
    sheet.getRange(row, sCol, 1, m.getNumColumns()).merge();
  }

  const newRows = groupRows + 1;
  sheet.getRange(groupStartRow, labelStartCol, newRows, labelNumCols).breakApart();
  sheet.getRange(groupStartRow, labelStartCol, newRows, labelNumCols).merge();
  sheet.getRange(groupStartRow, labelStartCol).setValue(labelValue);

  for (let i = 0; i < sharedCols.length; i++) {
    sheet.getRange(groupStartRow, sharedCols[i], newRows, 1).breakApart();
    sheet.getRange(groupStartRow, sharedCols[i], newRows, 1).merge();
  }

  sheet.getRange(row, incCol).setValue(newSub);
}
