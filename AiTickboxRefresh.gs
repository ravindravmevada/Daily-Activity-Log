function handleAiTickboxRefresh(controlsSheet, row, col) {
  const runRC = cellRefToRowCol(AI_TICKBOX_REFRESH.RUN_TICKBOX);
  if (row !== runRC.row || col !== runRC.col) return;
  if (controlsSheet.getRange(AI_TICKBOX_REFRESH.RUN_TICKBOX).getValue() !== true) return;

  if (controlsSheet.getRange(AI_TICKBOX_REFRESH.ENABLE_TICKBOX).getValue() !== true) {
    controlsSheet.getRange(AI_TICKBOX_REFRESH.RUN_TICKBOX).setValue(false);
    return;
  }

  const targetName = normalizeTabName(controlsSheet.getRange(AI_TICKBOX_REFRESH.TARGET_SHEET_CELL).getValue());
  const target = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetName);

  if (target) {
    refreshAiTickboxesForSheet(target);
  }

  controlsSheet.getRange(AI_TICKBOX_REFRESH.RUN_TICKBOX).setValue(false);
}

function refreshAiTickboxesForSheet(sheet) {
  const hasNotesCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_ACTIVITY_NOTES.HAS_NOTES_COL);
  const fixColLetter = ACTIVITY_NOTES_AI_FIX_AND_RESTORE.TICKBOX_COL;
  const undoColLetter = ACTIVITY_NOTES_AI_FIX_AND_RESTORE.UNDO_TICKBOX_COL;
  const dataStartRow = 3;
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return;

  const numRows = lastRow - dataStartRow + 1;
  const hasNotesVals = sheet.getRange(dataStartRow, hasNotesCol, numRows, 1).getValues();

  const enabledA1 = [];
  const disabledA1 = [];

  for (let i = 0; i < numRows; i++) {
    const r = dataStartRow + i;
    const v = hasNotesVals[i][0];
    if (v === "Yes") {
      enabledA1.push(fixColLetter + r, undoColLetter + r);
    } else if (v === "No" || v === "No Status") {
      disabledA1.push(fixColLetter + r, undoColLetter + r);
    }
  }

  if (enabledA1.length > 0) {
    const list = sheet.getRangeList(enabledA1);
    list.insertCheckboxes();
    list.setFontColor("#ffffff");
    list.setHorizontalAlignment("center");
  }

  if (disabledA1.length > 0) {
    const list = sheet.getRangeList(disabledA1);
    list.clearDataValidations();
    list.setValue("❌");
    list.setFontColor("#ff0000");
    list.setHorizontalAlignment("center");
  }
}
