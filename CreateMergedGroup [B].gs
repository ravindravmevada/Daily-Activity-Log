function handleCreateMergedGroupFromRange(sheet, row, col) {
  const editedCell = sheet.getRange(row, col);
  if (editedCell.getValue() !== true) return;

  const lastRow = sheet.getLastRow();
  const triggerValues = sheet.getRange(1, col, lastRow, 1).getValues();

  const tickedRows = [];
  for (let i = 0; i < triggerValues.length; i++) {
    if (triggerValues[i][0] === true) tickedRows.push(i + 1);
  }
  if (tickedRows.length < 2) return;

  let groupStartRow = Math.min.apply(null, tickedRows);
  let groupEndRow   = Math.max.apply(null, tickedRows);

  const labelColStart = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_START_COL);
  for (let i = 0; i < tickedRows.length; i++) {
    const mr = sheet.getRange(tickedRows[i], labelColStart).getMergedRanges();
    if (mr && mr.length > 0) {
      const m = mr[0];
      if (m.getRow() < groupStartRow)                     groupStartRow = m.getRow();
      if (m.getRow() + m.getNumRows() - 1 > groupEndRow) groupEndRow   = m.getRow() + m.getNumRows() - 1;
    }
  }

  const groupNumRows = groupEndRow - groupStartRow + 1;

  const mainNumColIdx     = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INCREMENT_NUMBER);
  const subNumColIdx      = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_INCREMENT_COL);
  const vertMergeColIdxs  = DAILY_ACTIVITY_LOG_COLS.GROUP_VERTICAL_MERGE_COLS.map(columnLetterToIndex);
  const labelStartColIdx  = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_START_COL);
  const labelEndColIdx    = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_END_COL);
  const labelNumCols      = labelEndColIdx - labelStartColIdx + 1;
  const indTickboxColIdxs = DAILY_ACTIVITY_LOG_COLS.GROUP_INDIVIDUAL_TICKBOX_COLS.map(columnLetterToIndex);
  const kqStartColIdx     = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.UNMERGED_ROW_DROPDOWN_START_COL);
  const kqEndColIdx       = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.UNMERGED_ROW_DROPDOWN_END_COL);
  const kqNumCols         = kqEndColIdx - kqStartColIdx + 1;
  const perRowColIdx      = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.PER_ROW_DURATION_COL);

  const rowBelowD  = sheet.getRange(groupEndRow + 1, mainNumColIdx).getValue();
  const newGroupNum = typeof rowBelowD === "number" ? rowBelowD + 1 : null;

  // safely break all merges in group range cols A to kqEnd
  const groupFullRange = sheet.getRange(groupStartRow, 1, groupNumRows, kqEndColIdx);
  const existingMerges = groupFullRange.getMergedRanges();
  for (let i = 0; i < existingMerges.length; i++) {
    const m = existingMerges[i];
    sheet.getRange(m.getRow(), m.getColumn(), m.getNumRows(), m.getNumColumns()).breakApart();
  }

  // uncheck all trigger tickboxes
  for (let i = 0; i < tickedRows.length; i++) {
    sheet.getRange(tickedRows[i], col).setValue(false);
  }

  // vertical merge A B C D
  for (let i = 0; i < vertMergeColIdxs.length; i++) {
    const r = sheet.getRange(groupStartRow, vertMergeColIdxs[i], groupNumRows, 1);
    r.merge();
    r.setVerticalAlignment("middle");
  }

  // merge E-F category label
  const labelRange = sheet.getRange(groupStartRow, labelStartColIdx, groupNumRows, labelNumCols);
  labelRange.merge();
  labelRange.setVerticalAlignment("middle");

  // set main group number
  if (newGroupNum !== null) sheet.getRange(groupStartRow, mainNumColIdx).setValue(newGroupNum);

  // set G H I tickboxes with correct colors
  const tickboxValidation = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
  const tickboxColors     = { "G": "#ff6d01", "H": "#34a853", "I": "#f7ef00" };
  for (let i = 0; i < indTickboxColIdxs.length; i++) {
    const colIdx    = indTickboxColIdxs[i];
    const colLetter = DAILY_ACTIVITY_LOG_COLS.GROUP_INDIVIDUAL_TICKBOX_COLS[i];
    const r         = sheet.getRange(groupStartRow, colIdx, groupNumRows, 1);
    r.setDataValidation(tickboxValidation);
    r.setValue(false);
    r.setFontColor(tickboxColors[colLetter] || "#ffffff");
    r.setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
  }

  // J sub numbers will be set by resequenceActivityNumbers after this function
  // just clear J and apply styling for now
  const groupNum    = newGroupNum !== null ? newGroupNum : sheet.getRange(groupStartRow, mainNumColIdx).getValue();
  const subNumRange = sheet.getRange(groupStartRow, subNumColIdx, groupNumRows, 1);
  subNumRange.clearDataValidations();
  subNumRange.clearContent();
  subNumRange.setNumberFormat('@');
  subNumRange.setFontWeight("bold").setFontColor("#f7ef00");
  subNumRange.setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);

  // K-Q = ONE horizontal merge per row (Activity dropdown)
  for (let r = 0; r < groupNumRows; r++) {
    const kqRange = sheet.getRange(groupStartRow + r, kqStartColIdx, 1, kqNumCols);
    kqRange.merge();
    kqRange.clearDataValidations();
    kqRange.clearContent();
    kqRange.setFontWeight('normal').setFontColor('#ffffff').setFontSize(11);
    kqRange.setHorizontalAlignment('center').setVerticalAlignment('middle');
    kqRange.setBorder(true, true, true, true, false, false, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
  }

  // apply activity dropdown from category label
  const labelValue = sheet.getRange(groupStartRow, labelStartColIdx).getValue();
  if (labelValue && labelValue !== "") {
    applyActivityDropdownToGroupRows(sheet, groupStartRow, groupNumRows, labelValue);
  }

  // merge AP across full main group range
  updateMainGroupDuration(sheet, groupStartRow, groupNumRows);
}
