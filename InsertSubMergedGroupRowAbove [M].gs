function handleInsertSubMergedGroupRowAbove(sheet, row, col) {
  const editedCell = sheet.getRange(row, col);
  if (editedCell.getValue() !== true) return;
  editedCell.setValue(false);

  const lastCol         = sheet.getLastColumn();
  const subLabelStartCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.SUB_GROUP_LABEL_MERGE_START_COL);
  const subLabelEndCol   = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.SUB_GROUP_LABEL_MERGE_END_COL);
  const subSubNumCol     = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_SUB_INCREMENT_COL);
  const subVertCols      = DAILY_ACTIVITY_LOG_COLS.SUB_GROUP_VERTICAL_MERGE_COLS.map(columnLetterToIndex);
  const mainTotalCol     = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.GROUP_TOTAL_DURATION_COL);
  const subTotalCol      = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.SUB_GROUP_TOTAL_DURATION_COL);

  // find sub-group bounds via K-L merge
  const subLabelMerges = sheet.getRange(row, subLabelStartCol).getMergedRanges();
  let subGroupStartRow, subGroupEndRow, subLabelStartColActual, subLabelNumCols;

  if (subLabelMerges && subLabelMerges.length > 0) {
    const m              = subLabelMerges[0];
    subGroupStartRow     = m.getRow();
    subGroupEndRow       = m.getRow() + m.getNumRows() - 1;
    subLabelStartColActual = m.getColumn();
    subLabelNumCols      = m.getNumColumns();
  } else {
    subGroupStartRow       = row;
    subGroupEndRow         = row;
    subLabelStartColActual = subLabelStartCol;
    subLabelNumCols        = subLabelEndCol - subLabelStartCol + 1;
  }

  const subLabelValue  = sheet.getRange(subGroupStartRow, subLabelStartColActual).getValue();

  // find max sub-sub number in this sub-group
  let maxSubSub = 0;
  for (let r = subGroupStartRow; r <= subGroupEndRow; r++) {
    const v = sheet.getRange(r, subSubNumCol).getValue();
    if (typeof v === "string" && v !== "") {
      const parts = v.toString().split(".");
      if (parts.length === 3) {
        const subSubPart = parseFloat(parts[2]);
        if (subSubPart > maxSubSub) maxSubSub = subSubPart;
      }
    }
  }
  const newSubSubSuffix = maxSubSub + 1;

  const subGroupRows = subGroupEndRow - subGroupStartRow + 1;

  // break sub-group vertical merges before insert
  for (let i = 0; i < subVertCols.length; i++) {
    sheet.getRange(subGroupStartRow, subVertCols[i], subGroupRows, 1).breakApart();
  }
  sheet.getRange(subGroupStartRow, subLabelStartColActual, subGroupRows, subLabelNumCols).breakApart();
  sheet.getRange(subGroupStartRow, subTotalCol, subGroupRows, 1).breakApart();
  sheet.getRange(subGroupStartRow, mainTotalCol, subGroupRows, 1).breakApart();

  // insert row
  sheet.insertRowsBefore(row, 1);

  // copy format from row below
  const source = sheet.getRange(row + 1, 1, 1, lastCol);
  const dest   = sheet.getRange(row,     1, 1, lastCol);
  source.copyTo(dest, { contentsOnly: false });

  // restore single-row merges from source (skip sub-group span merges)
  const srcMerges = source.getMergedRanges();
  for (let i = 0; i < srcMerges.length; i++) {
    const m = srcMerges[i];
    if (m.getNumRows() !== 1) continue;
    sheet.getRange(row, m.getColumn(), 1, m.getNumColumns()).merge();
  }

  const newSubGroupRows = subGroupRows + 1;

  // re-merge sub-group vertical cols G H I J
  for (let i = 0; i < subVertCols.length; i++) {
    const r = sheet.getRange(subGroupStartRow, subVertCols[i], newSubGroupRows, 1);
    r.breakApart();
    r.merge();
    r.setVerticalAlignment("middle");
  }

  // re-merge K-L sub-group label
  const subLabelRange = sheet.getRange(subGroupStartRow, subLabelStartColActual, newSubGroupRows, subLabelNumCols);
  subLabelRange.breakApart();
  subLabelRange.merge();
  subLabelRange.setValue(subLabelValue);
  subLabelRange.setVerticalAlignment("middle");

  // set new sub-sub number in col O
  const baseSubNum = sheet.getRange(subGroupStartRow, columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_INCREMENT_COL)).getValue();
  sheet.getRange(row, subSubNumCol).setValue(baseSubNum + "." + newSubSubSuffix);

  updateRowDuration(sheet, subGroupStartRow);
}
