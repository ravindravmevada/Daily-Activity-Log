function onControlsEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (sheetName !== CONTROLS_SHEET_NAME) return;

  handleAdmEnableFeature(sheet, row, col, e);
  handleAdmRefreshNow(sheet, row, col, e);
  handleAdmRebuildNow(sheet, row, col, e);
  handleAdmResetToDefault(sheet, row, col, e);
  handleActivityIdBulk(sheet, row, col);
  handleBulkAddEnableFeature(sheet, row, col, e);
  handleBulkAddRowsAbove(sheet, row, col, e);
  handleBulkAddRowsAboveActivate(sheet, row, col, e);
  handleBulkAddResetToDefault(sheet, row, col, e);
  handleAiTickboxRefresh(sheet, row, col);
  handleRowStatusEnableFeature(sheet, row, col, e);
  handleResetToDefault(sheet, row, col, e);
  handleEnableHighlighting(sheet, row, col, e);
  handleClearNow(sheet, row, col, e);
  handleHighlightOnSheetsEdit(sheet, row, col, e);
  handleClearFromSheetsEdit(sheet, row, col, e);
  handleTabVisibilityEnableFeature(sheet, row, col, e);
  handleTabVisibility(sheet, row, col, e);
  handleTabVisibilityResetToDefault(sheet, row, col, e);
  handleRguEnableFeature(sheet, row, col, e);
  handleRguResetToDefault(sheet, row, col, e);
  handleRguUnmergeNow(sheet, row, col, e);
}
