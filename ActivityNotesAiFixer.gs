function handleActivityNotesAiFix(sheet, row, e) {
  if (e && e.value !== "TRUE" && e.value !== true) return;

  const hasNotesCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_ACTIVITY_NOTES.HAS_NOTES_COL);
  if (sheet.getRange(row, hasNotesCol).getValue() !== "Yes") return;

  const tickCol = columnLetterToIndex(ACTIVITY_NOTES_AI_FIX_AND_RESTORE.TICKBOX_COL);
  const tickCell = sheet.getRange(row, tickCol);

  const notesCol = columnLetterToIndex(ACTIVITY_NOTES_AI_FIX_AND_RESTORE.NOTES_COL);
  const notesCell = sheet.getRange(row, notesCol);
  const original = notesCell.getValue();

  if (typeof original !== "string" || original.trim() === "") {
    tickCell.setValue(false);
    return;
  }

  const existingNote = notesCell.getNote();
  if (!existingNote || existingNote.indexOf(ACTIVITY_NOTES_AI_FIX_AND_RESTORE.BACKUP_PREFIX) !== 0) {
    notesCell.setNote(ACTIVITY_NOTES_AI_FIX_AND_RESTORE.BACKUP_PREFIX + original);
  }

  try {
    const fixed = fixNoteWithGemini(original);
    if (fixed && fixed.trim() !== "") {
      notesCell.setValue(fixed.trim());
      notesCell.setNote(ACTIVITY_NOTES_AI_FIX_AND_RESTORE.BACKUP_PREFIX + original);
    }
  } catch (err) {
    notesCell.setNote(ACTIVITY_NOTES_AI_FIX_AND_RESTORE.BACKUP_PREFIX + original + "\n[AI fix failed: " + err.message + "]");
  }

  tickCell.setValue(false);
}

function handleActivityNotesAiRestore(sheet, row, e) {
  if (e && e.value !== "TRUE" && e.value !== true) return;

  const hasNotesCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_ACTIVITY_NOTES.HAS_NOTES_COL);
  if (sheet.getRange(row, hasNotesCol).getValue() !== "Yes") return;

  const undoCol = columnLetterToIndex(ACTIVITY_NOTES_AI_FIX_AND_RESTORE.UNDO_TICKBOX_COL);
  const undoCell = sheet.getRange(row, undoCol);

  const notesCol = columnLetterToIndex(ACTIVITY_NOTES_AI_FIX_AND_RESTORE.NOTES_COL);
  const notesCell = sheet.getRange(row, notesCol);
  const note = notesCell.getNote();

  if (note && note.indexOf(ACTIVITY_NOTES_AI_FIX_AND_RESTORE.BACKUP_PREFIX) === 0) {
    let original = note.substring(ACTIVITY_NOTES_AI_FIX_AND_RESTORE.BACKUP_PREFIX.length);
    const failMarker = original.indexOf("\n[AI fix failed:");
    if (failMarker !== -1) original = original.substring(0, failMarker);
    notesCell.setValue(original);
    notesCell.clearNote();
  }

  undoCell.setValue(false);
}

function fixNoteWithGemini(noteText) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(ACTIVITY_NOTES_AI_FIX_AND_RESTORE.API_KEY_PROPERTY);
  if (!apiKey) throw new Error("API key not set");

  const url = "https://generativelanguage.googleapis.com/v1beta/models/" +
              ACTIVITY_NOTES_AI_FIX_AND_RESTORE.GEMINI_MODEL + ":generateContent?key=" + apiKey;

  const payload = {
    contents: [{ parts: [{ text: ACTIVITY_NOTES_AI_FIX_AND_RESTORE.PROMPT + "\n\n" + noteText }] }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const maxAttempts = 3;
  const retryDelayMs = 2000;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();

    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      if (!data.candidates || !data.candidates[0] || !data.candidates[0].content ||
          !data.candidates[0].content.parts || !data.candidates[0].content.parts[0] ||
          typeof data.candidates[0].content.parts[0].text !== "string") {
        throw new Error("No text returned");
      }
      return data.candidates[0].content.parts[0].text;
    }

    if (code === 503 && attempt < maxAttempts) {
      Utilities.sleep(retryDelayMs);
      continue;
    }

    throw new Error("Gemini HTTP " + code);
  }
}
