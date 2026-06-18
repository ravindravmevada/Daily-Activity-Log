function onChange(e) {
  if (!e || (e.changeType !== "REMOVE_ROW" && e.changeType !== "OTHER" && e.changeType !== "INSERT_ROW")) return;

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  if (SHEET_NAMES.DAILY_ACTIVITY_LOG.some(name => sheetName.includes(name))) {
    resequenceDColumnAndSubNumbers(sheet);
    resequenceActivityIds(sheet);
    if (isRowStatusHighlightEnabled()) handleRowStatusHighlight(sheet);
  }
}
