function routeDailyLogTickbox(sheet, row, col) {
  const loggedFromColIndex = columnLetterToIndex(DEVICE_VIA_CASCADE.LOGGED_FROM_COL);
  const activityTypeColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.ACTIVITY_TYPE_COL);
  const insertRowAboveCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INSERT_ROW_ABOVE);
  const rowHighlightCols = DAILY_ACTIVITY_LOG_COLS.ROW_HIGHLIGHT_COLS.map(columnLetterToIndex);
  const insertMergedGroupRowAboveCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INSERT_MERGED_GROUP_ROW_ABOVE);
  const createGroupTriggerCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.CREATE_GROUP_TRIGGER_COL);
  const labelMergeStartColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_START_COL);
  const labelMergeEndColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_END_COL);
  const incrementColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INCREMENT_NUMBER);
  const hasNotesColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_ACTIVITY_NOTES.HAS_NOTES_COL);

  if (col === insertRowAboveCol) {
    handleInsertBlankRowAbove(sheet, row, col);
    resequenceDColumnAndSubNumbers(sheet);
    return;
  }

  if (col === createGroupTriggerCol) {
    handleCreateMergedGroupFromRange(sheet, row, col);
    resequenceDColumnAndSubNumbers(sheet);
    return;
  }

  if (rowHighlightCols.includes(col)) {
    handleAutoRowHighlight(sheet, row, col);
    return;
  }

  if (col === insertMergedGroupRowAboveCol) {
    handleInsertMergedGroupRowAbove(sheet, row, col);
    resequenceDColumnAndSubNumbers(sheet);
    updateRowDuration(sheet, row);
    return;
  }

  if (col === incrementColIndex) {
    resequenceDColumnAndSubNumbers(sheet);
    return;
  }

  if (col >= labelMergeStartColIndex && col <= labelMergeEndColIndex) {
    handleCategoryChangeCascade(sheet, row, col);
    return;
  }

  if (col === activityTypeColIndex) {
    updateRowDuration(sheet, row);
    return;
  }

  if (col === hasNotesColIndex) {
    handleActivityNotesState(sheet, row);
    return;
  }

  if (col === loggedFromColIndex) {
    handleDeviceViaCascade(sheet, row);
    return;
  }

  for (let i = 0; i < DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS.length; i++) {
    const group = DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[i];
    const hasTimestampCol = columnLetterToIndex(group.hasTimestampCol);
    const timestampCol = columnLetterToIndex(group.timestampCol);
    const timestampStatusCol = columnLetterToIndex(group.timestampStatusCol);

    if (col === hasTimestampCol) {
      handleHasTimestampState(sheet, row, hasTimestampCol, timestampCol, group.label);
      handleHasTimestampStatus(sheet, row, hasTimestampCol, timestampStatusCol);
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
