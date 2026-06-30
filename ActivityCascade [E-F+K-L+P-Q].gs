function handleActivityCascade(sheet, row, col) {
  const mainLabelStartCol   = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_START_COL);
  const subLabelStartCol    = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.SUB_GROUP_LABEL_MERGE_START_COL);
  const phaseStartCol       = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.ACTIVITY_PHASE_DROPDOWN_START_COL);

  const isMainLabel = (col >= mainLabelStartCol && col <= columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_END_COL));
  const isSubLabel  = (col >= subLabelStartCol  && col <= columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.SUB_GROUP_LABEL_MERGE_END_COL));
  const isPhaseCol  = (col >= phaseStartCol     && col <= columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.ACTIVITY_PHASE_DROPDOWN_END_COL));

  if (isMainLabel) {
    const mainMerges   = sheet.getRange(row, mainLabelStartCol).getMergedRanges();
    const isMainGroup  = mainMerges.length > 0 && mainMerges[0].getNumRows() > 1;
    const categoryName = sheet.getRange(isMainGroup ? mainMerges[0].getRow() : row, mainLabelStartCol).getValue();

    if (isMainGroup) {
      const startRow = mainMerges[0].getRow();
      const numRows  = mainMerges[0].getNumRows();
      applyActivityDropdownToGroupRows(sheet, startRow, numRows, categoryName);
    } else {
      applyActivityDropdownToUnmergedRow(sheet, row, categoryName);
    }
    return;
  }

  if (isSubLabel) {
    const subMerges  = sheet.getRange(row, subLabelStartCol).getMergedRanges();
    const isSubGroup = subMerges.length > 0 && subMerges[0].getNumRows() > 1;
    const activityName = sheet.getRange(isSubGroup ? subMerges[0].getRow() : row, subLabelStartCol).getValue();

    if (isSubGroup) {
      const startRow = subMerges[0].getRow();
      const numRows  = subMerges[0].getNumRows();
      applyActivityPhaseDropdownToSubGroupRows(sheet, startRow, numRows, activityName);
    } else {
      applyActivityPhaseDropdownToUnmergedRow(sheet, row, activityName);
    }
    return;
  }
}

// ── apply Activity dropdown to main group rows (bottom to top) ───────────────
function applyActivityDropdownToGroupRows(sheet, startRow, numRows, categoryName) {
  // Inside a main group, activity dropdown is K-Q (narrowed — G-J are tickboxes+sub-number)
  const kqStartCol    = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_ROW_DROPDOWN_NARROWED_START_COL);
  const kqEndCol      = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.UNMERGED_ROW_DROPDOWN_END_COL);
  const kqNumCols     = kqEndCol - kqStartCol + 1;
  const klStartCol    = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.SUB_GROUP_LABEL_MERGE_START_COL);
  const klEndCol      = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.SUB_GROUP_LABEL_MERGE_END_COL);
  const klNumCols     = klEndCol - klStartCol + 1;
  const phaseStartCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.ACTIVITY_PHASE_DROPDOWN_START_COL);
  const phaseEndCol   = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.ACTIVITY_PHASE_DROPDOWN_END_COL);
  const phaseNumCols  = phaseEndCol - phaseStartCol + 1;
  const actSourceRange = getActivitySourceRange(categoryName);
  const activities     = getActivitiesForCategory(categoryName);
  const firstActivity  = activities.length > 0 ? activities[0] : '';

  // check if any H sub-groups exist in this group
  let hasSubGroups = false;
  for (let r = 0; r < numRows; r++) {
    const klMerges = sheet.getRange(startRow + r, klStartCol).getMergedRanges();
    if (klMerges.length > 0 && klMerges[0].getNumRows() > 1) { hasSubGroups = true; break; }
  }

  if (!hasSubGroups) {
    // ── FAST PATH: no sub-groups — apply to entire K-Q range in one batch ───
    const fullKqRange = sheet.getRange(startRow, kqStartCol, numRows, kqNumCols);
    if (actSourceRange && firstActivity !== '') {
      const actValidation = SpreadsheetApp.newDataValidation()
        .requireValueInRange(actSourceRange, true)
        .setAllowInvalid(false)
        .build();
      fullKqRange.setDataValidation(actValidation);
      fullKqRange.setValue(firstActivity);
    } else {
      fullKqRange.clearDataValidations();
      fullKqRange.clearContent();
    }
    return;
  }

  // ── SLOW PATH: H sub-groups exist — handle row by row ────────────────────
  const visited = new Set();
  for (let r = numRows - 1; r >= 0; r--) {
    const actualRow = startRow + r;
    if (visited.has(actualRow)) continue;

    const klMerges   = sheet.getRange(actualRow, klStartCol).getMergedRanges();
    const isSubGroup = klMerges.length > 0 && klMerges[0].getNumRows() > 1;

    if (isSubGroup) {
      const subStart = klMerges[0].getRow();
      const subSize  = klMerges[0].getNumRows();
      for (let sr = subStart; sr < subStart + subSize; sr++) visited.add(sr);

      const klRange = sheet.getRange(subStart, klStartCol, subSize, klNumCols);
      if (actSourceRange && firstActivity !== '') {
        const actValidation = SpreadsheetApp.newDataValidation()
          .requireValueInRange(actSourceRange, true)
          .setAllowInvalid(false)
          .build();
        klRange.setDataValidation(actValidation);
        klRange.setValue(firstActivity);
      } else {
        klRange.clearDataValidations();
        klRange.clearContent();
      }

      const phaseSourceRange = getActivityPhaseSourceRange(firstActivity);
      const phases           = getActivityPhasesForActivity(firstActivity);
      const firstPhase       = phases.length > 0 ? phases[0] : '';
      const phaseRange       = sheet.getRange(subStart, phaseStartCol, subSize, phaseNumCols);
      if (phaseSourceRange && firstPhase !== '') {
        const phaseValidation = SpreadsheetApp.newDataValidation()
          .requireValueInRange(phaseSourceRange, true)
          .setAllowInvalid(false)
          .build();
        phaseRange.setDataValidation(phaseValidation);
        phaseRange.setValue(firstPhase);
      } else {
        phaseRange.clearDataValidations();
        phaseRange.clearContent();
      }
    } else {
      visited.add(actualRow);
      const kqRange = sheet.getRange(actualRow, kqStartCol, 1, kqNumCols);
      if (actSourceRange && firstActivity !== '') {
        const actValidation = SpreadsheetApp.newDataValidation()
          .requireValueInRange(actSourceRange, true)
          .setAllowInvalid(false)
          .build();
        kqRange.setDataValidation(actValidation);
        kqRange.setValue(firstActivity);
      } else {
        kqRange.clearDataValidations();
        kqRange.clearContent();
      }
    }
  }
}

// ── apply Activity dropdown to solo unmerged row (K-Q) ───────────────────────
function applyActivityDropdownToUnmergedRow(sheet, row, categoryName) {
  const kqStartCol  = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.UNMERGED_ROW_DROPDOWN_START_COL);
  const kqEndCol    = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.UNMERGED_ROW_DROPDOWN_END_COL);
  const kqNumCols   = kqEndCol - kqStartCol + 1;
  const klStartCol  = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.SUB_GROUP_LABEL_MERGE_START_COL);
  const sourceRange = getActivitySourceRange(categoryName);
  const activities  = getActivitiesForCategory(categoryName);
  const firstActivity = activities.length > 0 ? activities[0] : '';

  // check if this row has a sub-group (K-L vertically merged)
  const klMerges   = sheet.getRange(row, klStartCol).getMergedRanges();
  const isSubGroup = klMerges.length > 0 && klMerges[0].getNumRows() > 1;

  if (isSubGroup) {
    // sub-group exists — only update K-L label value, never touch merge structure
    const klRange = sheet.getRange(klMerges[0].getRow(), klMerges[0].getColumn(), klMerges[0].getNumRows(), klMerges[0].getNumColumns());
    if (sourceRange && firstActivity !== '') {
      const validation = SpreadsheetApp.newDataValidation()
        .requireValueInRange(sourceRange, true)
        .setAllowInvalid(false)
        .build();
      klRange.setDataValidation(validation);
      klRange.setValue(firstActivity);
    } else {
      klRange.clearDataValidations();
      klRange.clearContent();
    }
    return;
  }

  // no sub-group — K-Q is one merged cell, safe to break and re-merge
  const kqRange   = sheet.getRange(row, kqStartCol, 1, kqNumCols);
  const allMerges = kqRange.getMergedRanges();
  for (let i = 0; i < allMerges.length; i++) {
    const m = allMerges[i];
    sheet.getRange(m.getRow(), m.getColumn(), m.getNumRows(), m.getNumColumns()).breakApart();
  }
  try { kqRange.breakApart(); } catch(e) {}

  kqRange.merge();
  kqRange.setFontWeight('normal').setFontColor('#ffffff').setFontSize(11);
  kqRange.setHorizontalAlignment('center').setVerticalAlignment('middle');
  kqRange.setBorder(true, true, true, true, false, false, '#ffffff', SpreadsheetApp.BorderStyle.SOLID);

  if (sourceRange && firstActivity !== '') {
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sourceRange, true)
      .setAllowInvalid(false)
      .build();
    kqRange.setDataValidation(validation);
    kqRange.clearContent();
    kqRange.setValue(firstActivity);
  } else {
    kqRange.clearDataValidations();
    kqRange.clearContent();
  }
}

// ── apply Activity Phase dropdown to sub-group rows (P-Q per row) ─────────────
function applyActivityPhaseDropdownToSubGroupRows(sheet, startRow, numRows, activityName) {
  const phaseStartCol  = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.ACTIVITY_PHASE_DROPDOWN_START_COL);
  const phaseEndCol    = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.ACTIVITY_PHASE_DROPDOWN_END_COL);
  const phaseNumCols   = phaseEndCol - phaseStartCol + 1;
  const sourceRange    = getActivityPhaseSourceRange(activityName);
  const phases         = getActivityPhasesForActivity(activityName);
  const fullPhaseRange = sheet.getRange(startRow, phaseStartCol, numRows, phaseNumCols);

  if (phases.length > 0 && sourceRange) {
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sourceRange, true)
      .setAllowInvalid(false)
      .build();
    fullPhaseRange.setDataValidation(validation);
    fullPhaseRange.setValue(phases[0]);
  } else {
    fullPhaseRange.clearDataValidations();
    fullPhaseRange.clearContent();
  }
}

// ── apply Activity Phase dropdown to single unmerged row (P-Q) ────────────────
function applyActivityPhaseDropdownToUnmergedRow(sheet, row, activityName) {
  const phaseStartCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.ACTIVITY_PHASE_DROPDOWN_START_COL);
  const phaseEndCol   = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.ACTIVITY_PHASE_DROPDOWN_END_COL);
  const phaseNumCols  = phaseEndCol - phaseStartCol + 1;
  const sourceRange   = getActivityPhaseSourceRange(activityName);
  const phases        = getActivityPhasesForActivity(activityName);
  const phaseRange    = sheet.getRange(row, phaseStartCol, 1, phaseNumCols);

  phaseRange.clearDataValidations();
  phaseRange.clearContent();

  if (phases.length > 0 && sourceRange) {
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sourceRange, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(row, phaseStartCol).setDataValidation(validation);
    sheet.getRange(row, phaseStartCol).setValue(phases[0]);
  }
}

// ── get Activity source range from Central Dropdown Source - Activity ──────────
function getActivitySourceRange(categoryName) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEET_REGISTRY.CENTRAL_ACTIVITY_SOURCE_SHEET_NAME);
  if (!sourceSheet) return null;

  const headerRow   = CENTRAL_SOURCE_CONFIG.HEADER_ROW;
  const dataStartRow = CENTRAL_SOURCE_CONFIG.DATA_START_ROW;
  const lastCol     = sourceSheet.getLastColumn();
  const headers     = sourceSheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === categoryName) {
      const lastDataRow = sourceSheet.getLastRow();
      return sourceSheet.getRange(dataStartRow, i + 1, lastDataRow - dataStartRow + 1, 1);
    }
  }
  return null;
}

// ── get activities array for a category ───────────────────────────────────────
function getActivitiesForCategory(categoryName) {
  const sourceRange = getActivitySourceRange(categoryName);
  if (!sourceRange) return [];
  return sourceRange.getValues().flat().filter(v => v !== "");
}

// ── get Activity Phase source range from Central Dropdown Source - Activity Phase
function getActivityPhaseSourceRange(activityName) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEET_REGISTRY.CENTRAL_ACTIVITY_PHASE_SOURCE_SHEET_NAME);
  if (!sourceSheet) return null;

  const headerRow    = CENTRAL_SOURCE_CONFIG.HEADER_ROW;
  const dataStartRow = CENTRAL_SOURCE_CONFIG.DATA_START_ROW;
  const lastCol      = sourceSheet.getLastColumn();
  const headers      = sourceSheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === activityName) {
      const lastDataRow = sourceSheet.getLastRow();
      return sourceSheet.getRange(dataStartRow, i + 1, lastDataRow - dataStartRow + 1, 1);
    }
  }
  return null;
}

// ── get phases array for an activity ──────────────────────────────────────────
function getActivityPhasesForActivity(activityName) {
  const sourceRange = getActivityPhaseSourceRange(activityName);
  if (!sourceRange) return [];
  return sourceRange.getValues().flat().filter(v => v !== "");
}
