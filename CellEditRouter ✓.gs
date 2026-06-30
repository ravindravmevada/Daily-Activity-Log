function routeDailyLogTickbox(sheet, row, col) {
  const insertRowAboveCol        = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INSERT_ROW_ABOVE);
  const createGroupCol           = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.CREATE_GROUP_TRIGGER_COL);
  const insertMergedGroupRowCol  = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INSERT_MERGED_GROUP_ROW_ABOVE);
  const createSubGroupCol        = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.CREATE_SUB_GROUP_TRIGGER_COL);
  const insertSubMergedRowCol    = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INSERT_SUB_MERGED_GROUP_ROW_ABOVE);
  const rowHighlightCols         = DAILY_ACTIVITY_LOG_COLS.ROW_HIGHLIGHT_COLS.map(columnLetterToIndex);
  const mainLabelStartCol        = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_START_COL);
  const mainLabelEndCol          = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_END_COL);
  const subLabelStartCol         = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.SUB_GROUP_LABEL_MERGE_START_COL);
  const subLabelEndCol           = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.SUB_GROUP_LABEL_MERGE_END_COL);
  const phaseStartCol            = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.ACTIVITY_PHASE_DROPDOWN_START_COL);
  const phaseEndCol              = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.ACTIVITY_PHASE_DROPDOWN_END_COL);
  const activityTypeCol          = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.ACTIVITY_TYPE_COL);
  const hasNotesCol              = columnLetterToIndex(DAILY_ACTIVITY_LOG_ACTIVITY_NOTES.HAS_NOTES_COL);
  const loggedFromCol            = columnLetterToIndex(DEVICE_VIA_CASCADE.LOGGED_FROM_COL);

  // ── Level 1 ──────────────────────────────────────────────────────────────
  if (col === insertRowAboveCol) {
    handleInsertBlankRowAbove(sheet, row, col);
    resequenceActivityNumbers(sheet);
    return;
  }

  if (col === createGroupCol) {
    handleCreateMergedGroupFromRange(sheet, row, col);
    SpreadsheetApp.flush();
    resequenceDNumbers(sheet);
    return;
  }

  // ── Level 2 ──────────────────────────────────────────────────────────────
  if (col === insertMergedGroupRowCol) {
    handleInsertMergedGroupRowAbove(sheet, row, col);
    resequenceActivityNumbers(sheet);
    updateRowDuration(sheet, row);
    return;
  }

  if (col === createSubGroupCol) {
    const subGroup = handleCreateSubMergedGroup(sheet, row, col);
    SpreadsheetApp.flush();
    resequenceJNumbers(sheet);
    if (subGroup) writeONumbersForSubGroup(sheet, subGroup.startRow, subGroup.numRows);
    return;
  }

  // ── Level 3 ──────────────────────────────────────────────────────────────
  if (col === insertSubMergedRowCol) {
    handleInsertSubMergedGroupRowAbove(sheet, row, col);
    resequenceActivityNumbers(sheet);
    updateRowDuration(sheet, row);
    return;
  }

  // ── Row highlight (C, I, N) ───────────────────────────────────────────────
  if (rowHighlightCols.includes(col)) {
    handleAutoRowHighlight(sheet, row, col);
    return;
  }

  // ── Category label changed (E-F) ─────────────────────────────────────────
  if (col >= mainLabelStartCol && col <= mainLabelEndCol) {
    handleActivityCascade(sheet, row, col);
    return;
  }

  // ── Activity label changed (K-L) → cascade to Phase ──────────────────────
  if (col >= subLabelStartCol && col <= subLabelEndCol) {
    handleActivityCascade(sheet, row, col);
    return;
  }

  // ── Activity Type col AG ──────────────────────────────────────────────────
  if (col === activityTypeCol) {
    updateRowDuration(sheet, row);
    return;
  }

  // ── Has Notes col AK ──────────────────────────────────────────────────────
  if (col === hasNotesCol) {
    handleActivityNotesState(sheet, row);
    return;
  }

  // ── Logged From col AT ────────────────────────────────────────────────────
  if (col === loggedFromCol) {
    handleDeviceViaCascade(sheet, row);
    return;
  }

  // ── Timestamp groups ──────────────────────────────────────────────────────
  for (let i = 0; i < DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS.length; i++) {
    const group            = DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[i];
    const hasTimestampCol  = columnLetterToIndex(group.hasTimestampCol);
    const timestampCol     = columnLetterToIndex(group.timestampCol);
    const timestampStatCol = columnLetterToIndex(group.timestampStatusCol);

    if (col === hasTimestampCol) {
      handleHasTimestampState(sheet, row, hasTimestampCol, timestampCol, group.label);
      handleHasTimestampStatus(sheet, row, hasTimestampCol, timestampStatCol);
      updateRowDuration(sheet, row);
      if (isRowStatusHighlightEnabled()) handleRowStatusHighlight(sheet);
      resequenceActivityIds(sheet);
      return;
    }

    if (col === timestampCol) {
      handleTickboxTimestampToggle(sheet, row, timestampCol);
      updateRowDuration(sheet, row);
      if (isRowStatusHighlightEnabled()) handleRowStatusHighlight(sheet);
      resequenceActivityIds(sheet);
      return;
    }
  }
}
