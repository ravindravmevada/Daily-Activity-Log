function handleCreateSubMergedGroup(sheet, row, col) {
  const editedCell = sheet.getRange(row, col);
  if (editedCell.getValue() !== true) return null;

  // ── Find all ticked rows in col H ────────────────────────────────────────
  const lastRow       = sheet.getLastRow();
  const triggerValues = sheet.getRange(1, col, lastRow, 1).getValues();
  const tickedRows    = [];
  for (let i = 0; i < triggerValues.length; i++) {
    if (triggerValues[i][0] === true) tickedRows.push(i + 1);
  }
  if (tickedRows.length < 2) return null;

  const subGroupStartRow = Math.min.apply(null, tickedRows);
  const subGroupEndRow   = Math.max.apply(null, tickedRows);
  const subGroupNumRows  = subGroupEndRow - subGroupStartRow + 1;

  // uncheck all H tickboxes
  for (let i = 0; i < tickedRows.length; i++) {
    sheet.getRange(tickedRows[i], col).setValue(false);
  }

  // ── Column indexes from MainConfig ────────────────────────────────────────
  const subVertMergeCols = DAILY_ACTIVITY_LOG_COLS.SUB_GROUP_VERTICAL_MERGE_COLS.map(columnLetterToIndex); // G H I J
  const klStartCol       = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.SUB_GROUP_LABEL_MERGE_START_COL);
  const klEndCol         = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.SUB_GROUP_LABEL_MERGE_END_COL);
  const klNumCols        = klEndCol - klStartCol + 1;
  const mCol             = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.SUB_SUB_GROUP_INDIVIDUAL_TICKBOX_COLS[0]);
  const nCol             = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.SUB_SUB_GROUP_INDIVIDUAL_TICKBOX_COLS[1]);
  const oCol             = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_SUB_INCREMENT_COL);
  const phaseStartCol    = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.ACTIVITY_PHASE_DROPDOWN_START_COL);
  const phaseEndCol      = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.ACTIVITY_PHASE_DROPDOWN_END_COL);
  const phaseNumCols     = phaseEndCol - phaseStartCol + 1;
  const kqEndCol         = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.UNMERGED_ROW_DROPDOWN_END_COL);
  const kqNumCols        = kqEndCol - klStartCol + 1;
  const labelStartCol    = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_START_COL);
  const aoCol            = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.SUB_GROUP_TOTAL_DURATION_COL);
  const WHITE            = "#ffffff";
  const SOLID            = SpreadsheetApp.BorderStyle.SOLID;

  // ── BATCH 1: G H I J vertical merge ──────────────────────────────────────
  for (let i = 0; i < subVertMergeCols.length; i++) {
    sheet.getRange(subGroupStartRow, subVertMergeCols[i], subGroupNumRows, 1)
      .merge().setVerticalAlignment("middle");
  }

  // ── BATCH 2: K-L vertical merge + activity dropdown ──────────────────────
  // read existing activity value BEFORE breaking K-Q
  const existingActivityValue = sheet.getRange(subGroupStartRow, klStartCol).getValue();

  // break existing K-Q horizontal merges
  const kqRange        = sheet.getRange(subGroupStartRow, klStartCol, subGroupNumRows, kqNumCols);
  const existingMerges = kqRange.getMergedRanges();
  for (let i = 0; i < existingMerges.length; i++) {
    const m = existingMerges[i];
    sheet.getRange(m.getRow(), m.getColumn(), m.getNumRows(), m.getNumColumns()).breakApart();
  }
  kqRange.clearDataValidations();
  kqRange.clearContent();

  // vertically merge K-L
  const klMergeRange = sheet.getRange(subGroupStartRow, klStartCol, subGroupNumRows, klNumCols);
  klMergeRange.merge().setVerticalAlignment("middle");

  // apply activity dropdown directly to K-L only (sub-group range only, not whole main group)
  const categoryValue  = sheet.getRange(subGroupStartRow, labelStartCol).getValue();
  const actSourceRange = getActivitySourceRange(categoryValue);
  if (actSourceRange) {
    const actValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(actSourceRange, true)
      .setAllowInvalid(false)
      .build();
    klMergeRange.setDataValidation(actValidation);
  }
  // restore existing activity value or set first activity from source
  if (existingActivityValue && existingActivityValue !== "") {
    sheet.getRange(subGroupStartRow, klStartCol).setValue(existingActivityValue);
  } else {
    const activities = getActivitiesForCategory(categoryValue);
    if (activities.length > 0) {
      sheet.getRange(subGroupStartRow, klStartCol).setValue(activities[0]);
    }
  }

  // ── BATCH 3: M N O P-Q in one fast sequence ───────────────────────────────
  const tickboxValidation = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();

  // M - orange
  sheet.getRange(subGroupStartRow, mCol, subGroupNumRows, 1)
    .setDataValidation(tickboxValidation).setValue(false)
    .setFontColor("#ff6d01").setBorder(true, true, true, true, true, true, WHITE, SOLID);

  // N - yellow
  sheet.getRange(subGroupStartRow, nCol, subGroupNumRows, 1)
    .setDataValidation(tickboxValidation).setValue(false)
    .setFontColor("#f7ef00").setBorder(true, true, true, true, true, true, WHITE, SOLID);

  // O - style only, values written by writeONumbersForSubGroup after resequenceJNumbers
  sheet.getRange(subGroupStartRow, oCol, subGroupNumRows, 1)
    .clearDataValidations().clearContent()
    .setNumberFormat('@').setFontWeight("bold").setFontColor("#f7ef00")
    .setBorder(true, true, true, true, true, true, WHITE, SOLID);

  // P-Q - merge per row then batch dropdown
  for (let r = 0; r < subGroupNumRows; r++) {
    sheet.getRange(subGroupStartRow + r, phaseStartCol, 1, phaseNumCols).merge();
  }
  sheet.getRange(subGroupStartRow, phaseStartCol, subGroupNumRows, phaseNumCols)
    .setFontWeight('normal').setFontColor(WHITE).setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, WHITE, SOLID);

  const activityValue = existingActivityValue || sheet.getRange(subGroupStartRow, klStartCol).getValue();
  if (activityValue && activityValue !== "") {
    applyActivityPhaseDropdownToSubGroupRows(sheet, subGroupStartRow, subGroupNumRows, activityValue);
  }

  // ── AO: vertical merge last ───────────────────────────────────────────────
  sheet.getRange(subGroupStartRow, aoCol, subGroupNumRows, 1)
    .merge().setVerticalAlignment("middle");

  return { startRow: subGroupStartRow, numRows: subGroupNumRows };
}

function writeONumbersForSubGroup(sheet, subGroupStartRow, subGroupNumRows) {
  const subNumColIdx    = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_INCREMENT_COL);
  const subSubNumColIdx = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_SUB_INCREMENT_COL);
  const WHITE           = "#ffffff";
  const SOLID           = SpreadsheetApp.BorderStyle.SOLID;

  const jValue  = sheet.getRange(subGroupStartRow, subNumColIdx).getValue();
  const oValues = [];
  for (let r = 0; r < subGroupNumRows; r++) {
    oValues.push([jValue + '.' + (subGroupNumRows - r)]);
  }
  sheet.getRange(subGroupStartRow, subSubNumColIdx, subGroupNumRows, 1)
    .setNumberFormat('@').setValues(oValues)
    .setFontWeight("bold").setFontColor("#f7ef00")
    .setBorder(true, true, true, true, true, true, WHITE, SOLID);
}
