function onEdit(e) {
  const sheet     = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const row       = e.range.getRow();
  const col       = e.range.getColumn();
  const value     = e.value;

  if (sheetName === CONTROLS_SHEET_NAME) return;
  if (!isDailyLogSheet(sheetName)) return;

  // ignore untick events on trigger columns to prevent double-fire
  if ((col === columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.CREATE_GROUP_TRIGGER_COL) ||
       col === columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.CREATE_SUB_GROUP_TRIGGER_COL)) &&
      (value === 'FALSE' || value === false || value === '')) return;

  routeDailyLogTickbox(sheet, row, col);
}
