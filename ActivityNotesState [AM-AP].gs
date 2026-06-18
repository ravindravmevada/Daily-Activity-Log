function handleActivityNotesState(sheet, row) {
  const hasNotesCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_ACTIVITY_NOTES.HAS_NOTES_COL);
  const notesCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_ACTIVITY_NOTES.NOTES_COL);

  const hasNotesVal = sheet.getRange(row, hasNotesCol).getValue();
  const notesCell = sheet.getRange(row, notesCol);
  const current = notesCell.getValue();

  if (hasNotesVal === "No Status") {
    notesCell.setValue("Activity Notes Not Decided 🟡");
  } else if (hasNotesVal === "No") {
    notesCell.setValue("No Activity Notes 🔴");
  } else if (hasNotesVal === "Yes") {
    if (current === "Activity Notes Not Decided 🟡" || current === "No Activity Notes 🔴") {
      notesCell.setValue("");
    }
  }

  setAiTickboxesEnabled(sheet, row, hasNotesVal === "Yes");
}

function setAiTickboxesEnabled(sheet, row, enabled) {
  const fixCol = columnLetterToIndex(ACTIVITY_NOTES_AI_FIX_AND_RESTORE.TICKBOX_COL);
  const undoCol = columnLetterToIndex(ACTIVITY_NOTES_AI_FIX_AND_RESTORE.UNDO_TICKBOX_COL);
  const fixCell = sheet.getRange(row, fixCol);
  const undoCell = sheet.getRange(row, undoCol);

  if (enabled) {
    fixCell.insertCheckboxes();
    undoCell.insertCheckboxes();
    fixCell.setFontColor("#ffffff").setHorizontalAlignment("center");
    undoCell.setFontColor("#ffffff").setHorizontalAlignment("center");
  } else {
    fixCell.clearDataValidations();
    undoCell.clearDataValidations();
    fixCell.setValue("❌").setFontColor("#ff0000").setHorizontalAlignment("center");
    undoCell.setValue("❌").setFontColor("#ff0000").setHorizontalAlignment("center");
  }
}
