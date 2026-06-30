function handleCreateMergedGroupFromRange(sheet, row, col) {
  const editedCell = sheet.getRange(row, col);
  if (editedCell.getValue() !== true) return;

  // ── Find all ticked rows in col B ────────────────────────────────────────
  const lastRow       = sheet.getLastRow();
  const triggerValues = sheet.getRange(1, col, lastRow, 1).getValues();
  const tickedRows    = [];
  for (let i = 0; i < triggerValues.length; i++) {
    if (triggerValues[i][0] === true) tickedRows.push(i + 1);
  }
  if (tickedRows.length < 2) return;

  const groupStartRow = Math.min.apply(null, tickedRows);
  const groupEndRow   = Math.max.apply(null, tickedRows);
  const groupNumRows  = groupEndRow - groupStartRow + 1;

  // uncheck all B tickboxes
  for (let i = 0; i < tickedRows.length; i++) {
    sheet.getRange(tickedRows[i], col).setValue(false);
  }

  // ── Column indexes from MainConfig ────────────────────────────────────────
  const mainNumColIdx  = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INCREMENT_NUMBER);
  const labelStartCol  = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_START_COL);
  const labelEndCol    = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_END_COL);
  const labelNumCols   = labelEndCol - labelStartCol + 1;
  const vertMergeCols  = DAILY_ACTIVITY_LOG_COLS.GROUP_VERTICAL_MERGE_COLS.map(columnLetterToIndex);
  const apCol          = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.GROUP_TOTAL_DURATION_COL);
  const gCol           = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_INDIVIDUAL_TICKBOX_COLS[0]);
  const hCol           = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_INDIVIDUAL_TICKBOX_COLS[1]);
  const iCol           = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_INDIVIDUAL_TICKBOX_COLS[2]);
  const jCol           = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_INCREMENT_COL);
  const defaultActCol  = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.UNMERGED_ROW_DROPDOWN_START_COL);
  const kqStartCol     = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_ROW_DROPDOWN_NARROWED_START_COL);
  const kqEndCol       = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.UNMERGED_ROW_DROPDOWN_END_COL);
  const kqNumCols      = kqEndCol - kqStartCol + 1;
  const gqNumCols      = kqEndCol - defaultActCol + 1;
  const WHITE          = "#ffffff";
  const SOLID          = SpreadsheetApp.BorderStyle.SOLID;

  // ── Derive group number from row below ────────────────────────────────────
  const rowBelowNum = sheet.getRange(groupEndRow + 1, mainNumColIdx).getValue();
  const groupNum    = typeof rowBelowNum === "number" ? rowBelowNum + 1 : 1;

  // ── A-D: vertical merge + group number ───────────────────────────────────
  for (let i = 0; i < vertMergeCols.length; i++) {
    sheet.getRange(groupStartRow, vertMergeCols[i], groupNumRows, 1)
      .merge().setVerticalAlignment("middle");
  }
  sheet.getRange(groupStartRow, mainNumColIdx).setValue(groupNum);

  // ── E-F: vertical merge category label ───────────────────────────────────
  sheet.getRange(groupStartRow, labelStartCol, groupNumRows, labelNumCols)
    .merge().setVerticalAlignment("middle");

  // ── Break old G-Q merge + clear everything ────────────────────────────────
  const gqFullRange    = sheet.getRange(groupStartRow, defaultActCol, groupNumRows, gqNumCols);
  const existingMerges = gqFullRange.getMergedRanges();
  for (let i = 0; i < existingMerges.length; i++) {
    const m = existingMerges[i];
    sheet.getRange(m.getRow(), m.getColumn(), m.getNumRows(), m.getNumColumns()).breakApart();
  }
  gqFullRange.clearDataValidations();
  gqFullRange.clearContent();

  // ── G H I tickboxes ───────────────────────────────────────────────────────
  const tickboxValidation = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
  sheet.getRange(groupStartRow, gCol, groupNumRows, 1)
    .setDataValidation(tickboxValidation).setValue(false)
    .setFontColor("#ff6d01").setBorder(true, true, true, true, true, true, WHITE, SOLID);
  sheet.getRange(groupStartRow, hCol, groupNumRows, 1)
    .setDataValidation(tickboxValidation).setValue(false)
    .setFontColor("#34a853").setBorder(true, true, true, true, true, true, WHITE, SOLID);
  sheet.getRange(groupStartRow, iCol, groupNumRows, 1)
    .setDataValidation(tickboxValidation).setValue(false)
    .setFontColor("#f7ef00").setBorder(true, true, true, true, true, true, WHITE, SOLID);

  // ── J: sub-numbers bottom=.1 top=.highest in one batch ───────────────────
  const jValues = [];
  for (let r = 0; r < groupNumRows; r++) {
    jValues.push([groupNum + '.' + (groupNumRows - r)]);
  }
  sheet.getRange(groupStartRow, jCol, groupNumRows, 1)
    .setNumberFormat('@').setValues(jValues)
    .setFontWeight("bold").setFontColor("#f7ef00")
    .setBorder(true, true, true, true, true, true, WHITE, SOLID);

  // ── K-Q: merge per row + activity dropdown in one batch ──────────────────
  for (let r = 0; r < groupNumRows; r++) {
    sheet.getRange(groupStartRow + r, kqStartCol, 1, kqNumCols).merge();
  }
  sheet.getRange(groupStartRow, kqStartCol, groupNumRows, kqNumCols)
    .setFontWeight('normal').setFontColor(WHITE).setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, WHITE, SOLID);

  const categoryValue = sheet.getRange(groupStartRow, labelStartCol).getValue();
  if (categoryValue && categoryValue !== "") {
    applyActivityDropdownToGroupRows(sheet, groupStartRow, groupNumRows, categoryValue);
  }

  // ── AP: vertical merge last ───────────────────────────────────────────────
  sheet.getRange(groupStartRow, apCol, groupNumRows, 1)
    .merge().setVerticalAlignment("middle");
}
