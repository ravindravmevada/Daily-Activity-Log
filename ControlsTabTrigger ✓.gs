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

  handleBulkInsertEnableFeature(sheet, row, col, e);
  handleBulkInsertRowsAbove(sheet, row, col, e);
  handleBulkInsertUpToEdit(sheet, row, col, e);
  handleBulkInsertResetToDefault(sheet, row, col, e);

  handleAfrtaEnableFeature(sheet, row, col, e);
  handleAfrtaApplyNow(sheet, row, col, e);
  handleAfrtaResetToDefault(sheet, row, col, e);

  handleRowStatusEnableFeature(sheet, row, col, e);
  handleResetToDefault(sheet, row, col, e);
  handleEnableHighlighting(sheet, row, col, e);
  handleDisableNow(sheet, row, col, e);
  handleEnableHighlightingOnSheetsEdit(sheet, row, col, e);
  handleDisableHighlightingOnSheetsEdit(sheet, row, col, e);

  handleSheetVisibilityEnableFeature(sheet, row, col, e);
  handleSheetVisibility(sheet, row, col, e);
  handleSheetVisibilityResetToDefault(sheet, row, col, e);

  handleRguEnableFeature(sheet, row, col, e);
  handleRguUnmergeNow(sheet, row, col, e);
  handleRguResetToDefault(sheet, row, col, e);

  handleCwaEnableFeature(sheet, row, col, e);
  handleCwaApplyNow(sheet, row, col, e);
  handleCwaResetToDefault(sheet, row, col, e);
}
