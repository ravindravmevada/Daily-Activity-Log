function onActivityNotesFixRestoreTickboxesEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (!isDailyLogSheet(sheetName)) return;

  const aiFixCol = columnLetterToIndex(ACTIVITY_NOTES_AI_FIX_AND_RESTORE.TICKBOX_COL);
  const aiRestoreCol = columnLetterToIndex(ACTIVITY_NOTES_AI_FIX_AND_RESTORE.UNDO_TICKBOX_COL);

  if (col === aiFixCol) {
    handleActivityNotesAiFix(sheet, row, e);
    return;
  }

  if (col === aiRestoreCol) {
    handleActivityNotesAiRestore(sheet, row, e);
    return;
  }
}
