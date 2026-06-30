function onEdit(e) {
  const sheet     = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const row       = e.range.getRow();
  const col       = e.range.getColumn();

  if (sheetName === CONTROLS_SHEET_NAME) return;
  if (!isDailyLogSheet(sheetName)) return;

  routeDailyLogTickbox(sheet, row, col);
}
