function handleCreateSubMergedGroup(sheet, row, col) {
  const editedCell = sheet.getRange(row, col);
  if (editedCell.getValue() !== true) return;

  const lastRow = sheet.getLastRow();
  const triggerValues = sheet.getRange(1, col, lastRow, 1).getValues();

  const tickedRows = [];
  for (let i = 0; i < triggerValues.length; i++) {
    if (triggerValues[i][0] === true) tickedRows.push(i + 1);
  }
  if (tickedRows.length < 2) return;

  let subGroupStartRow = Math.min.apply(null, tickedRows);
  let subGroupEndRow   = Math.max.apply(null, tickedRows);

  const subLabelStartColIdx = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.SUB_GROUP_LABEL_MERGE_START_COL);
  const subLabelEndColIdx   = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.SUB_GROUP_LABEL_MERGE_END_COL);

  // expand range to include any existing sub-group merges
  for (let i = 0; i < tickedRows.length; i++) {
    const mr = sheet.getRange(tickedRows[i], subLabelStartColIdx).getMergedRanges();
    if (mr && mr.length > 0) {
      const m = mr[0];
      if (m.getRow() < subGroupStartRow)                         subGroupStartRow = m.getRow();
      if (m.getRow() + m.getNumRows() - 1 > subGroupEndRow)     subGroupEndRow   = m.getRow() + m.getNumRows() - 1;
    }
  }

  const subGroupNumRows = subGroupEndRow - subGroupStartRow + 1;

  const subNumColIdx        = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_INCREMENT_COL);
  const subSubNumColIdx     = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_SUB_INCREMENT_COL);
  const subVertMergeColIdxs = DAILY_ACTIVITY_LOG_COLS.SUB_GROUP_VERTICAL_MERGE_COLS.map(columnLetterToIndex);
  const subLabelNumCols     = subLabelEndColIdx - subLabelStartColIdx + 1;
  const subSubTickboxColIdxs = DAILY_ACTIVITY_LOG_COLS.SUB_SUB_GROUP_INDIVIDUAL_TICKBOX_COLS.map(columnLetterToIndex);
  const phaseStartColIdx    = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.ACTIVITY_PHASE_DROPDOWN_START_COL);
  const phaseEndColIdx      = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.ACTIVITY_PHASE_DROPDOWN_END_COL);
  const phaseNumCols        = phaseEndColIdx - phaseStartColIdx + 1;
  const perRowColIdx        = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.PER_ROW_DURATION_COL);
  const subTotalColIdx      = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.SUB_GROUP_TOTAL_DURATION_COL);
  const mainTotalColIdx     = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.GROUP_TOTAL_DURATION_COL);

  // break existing merges in sub-group range
  const subGroupColEnd = phaseEndColIdx;
  const subGroupRange  = sheet.getRange(subGroupStartRow, subLabelStartColIdx, subGroupNumRows, subGroupColEnd - subLabelStartColIdx + 1);
  const existingMerges = subGroupRange.getMergedRanges();
  const mergesToRestore = [];

  for (let i = 0; i < existingMerges.length; i++) {
    const m = existingMerges[i];
    if (m.getColumn() + m.getNumColumns() - 1 > subGroupColEnd) {
      mergesToRestore.push({ startRow: m.getRow(), startCol: m.getColumn(), numRows: m.getNumRows(), numCols: m.getNumColumns() });
    }
    m.breakApart();
  }

  // uncheck all H tickboxes
  for (let i = 0; i < tickedRows.length; i++) {
    sheet.getRange(tickedRows[i], col).setValue(false);
  }

  // vertical merge G H I J for sub-group
  for (let i = 0; i < subVertMergeColIdxs.length; i++) {
    const r = sheet.getRange(subGroupStartRow, subVertMergeColIdxs[i], subGroupNumRows, 1);
    r.merge();
    r.setVerticalAlignment("middle");
  }

  // merge K-L sub-group activity label
  const subLabelRange = sheet.getRange(subGroupStartRow, subLabelStartColIdx, subGroupNumRows, subLabelNumCols);
  subLabelRange.merge();
  subLabelRange.setVerticalAlignment("middle");

  // set sub-sub tickboxes M N with colors
  const tickboxValidation   = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
  const subSubTickboxColors = { "M": "#ff6d01", "N": "#f7ef00" };
  for (let i = 0; i < subSubTickboxColIdxs.length; i++) {
    const colIdx    = subSubTickboxColIdxs[i];
    const colLetter = DAILY_ACTIVITY_LOG_COLS.SUB_SUB_GROUP_INDIVIDUAL_TICKBOX_COLS[i];
    const r         = sheet.getRange(subGroupStartRow, colIdx, subGroupNumRows, 1);
    r.setDataValidation(tickboxValidation);
    r.setValue(false);
    r.setFontColor(subSubTickboxColors[colLetter] || "#ffffff");
    r.setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
  }

  // sub-sub numbers in col O will be set by resequenceActivityNumbers after flush
  // just clear O for now
  const subSubNumRange = sheet.getRange(subGroupStartRow, subSubNumColIdx, subGroupNumRows, 1);
  subSubNumRange.clearDataValidations();
  subSubNumRange.clearContent();
  subSubNumRange.setFontWeight("bold").setFontColor("#f7ef00");
  subSubNumRange.setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);

  // set phase dropdown P-Q per row
  for (let r = 0; r < subGroupNumRows; r++) {
    const phaseRange = sheet.getRange(subGroupStartRow + r, phaseStartColIdx, 1, phaseNumCols);
    phaseRange.clearDataValidations();
    phaseRange.clearContent();
    phaseRange.merge();
    phaseRange.setBorder(true, true, true, true, false, false, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
  }

  // apply phase dropdown cascade from sub-group activity label
  const subLabelValue = sheet.getRange(subGroupStartRow, subLabelStartColIdx).getValue();
  if (subLabelValue && subLabelValue !== "") {
    applyActivityPhaseDropdownToSubGroupRows(sheet, subGroupStartRow, subGroupNumRows, subLabelValue);
  }

  // restore merges beyond sub-group range
  for (let i = 0; i < mergesToRestore.length; i++) {
    const m = mergesToRestore[i];
    for (let r = 0; r < m.numRows; r++) {
      sheet.getRange(m.startRow + r, m.startCol, 1, m.numCols).merge();
    }
  }

  // merge AO across sub-group range only — do not touch AN or AP
  updateSubGroupDuration(sheet, subGroupStartRow, subGroupNumRows);
}
