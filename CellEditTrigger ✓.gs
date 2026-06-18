function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const editedCell = e.range;
  const row = editedCell.getRow();
  const col = editedCell.getColumn();

  if (sheetName === CONTROLS_SHEET_NAME) {
    handleColumnWidthSync(sheet, row, col);
    return;
  }

  if (SHEET_NAMES.DAILY_ACTIVITY_LOG.some(name => sheetName.includes(name))) {
    routeDailyLogTickbox(sheet, row, col);
  }
}
