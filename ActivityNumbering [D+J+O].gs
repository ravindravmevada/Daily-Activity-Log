// ── Shared helper: detect data rows via col A checkbox validation ─────────────
// For merged rows (inside a group), getDataValidations only returns validation
// on the top cell of the merge. So we also mark all rows within a D-column merge
// as data rows by reading the merge ranges directly.
function getDataRowMap(sheet) {
  const dataStartRow = 3;
  const lastRow      = sheet.getLastRow();
  if (lastRow < dataStartRow) return { dataStartRow, lastRow, isDataRow: {} };

  const insertColIdx  = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INSERT_ROW_ABOVE);
  const mainNumColIdx = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INCREMENT_NUMBER);
  const numDataRows   = lastRow - dataStartRow + 1;

  // detect via col A validation (works for unmerged rows)
  const indicatorVals = sheet.getRange(dataStartRow, insertColIdx, numDataRows, 1).getDataValidations();
  const isDataRow     = {};
  for (let i = 0; i < numDataRows; i++) {
    const v = indicatorVals[i][0];
    if (v && v.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      isDataRow[dataStartRow + i] = true;
    }
  }

  // for merged rows: col A validation only exists on the top cell of the merge
  // so also mark all rows inside a D-column merge as data rows
  const mainMerges = sheet.getRange(dataStartRow, mainNumColIdx, numDataRows, 1).getMergedRanges();
  for (let i = 0; i < mainMerges.length; i++) {
    const m = mainMerges[i];
    if (m.getNumRows() < 2) continue;
    for (let r = m.getRow(); r < m.getRow() + m.getNumRows(); r++) {
      isDataRow[r] = true;
    }
  }

  return { dataStartRow, lastRow, isDataRow };
}

// ── Shared helper: write a number cell ───────────────────────────────────────
function writeNum(sheet, row, col, value) {
  const cell = sheet.getRange(row, col);
  cell.setNumberFormat('@');
  cell.setValue(String(value));
  cell.setFontWeight('bold').setFontColor('#f7ef00');
}

// ── D: main numbers bottom to top ────────────────────────────────────────────
function resequenceDNumbers(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const { dataStartRow, lastRow, isDataRow } = getDataRowMap(sheet);
  if (lastRow < dataStartRow) return;

  const mainNumColIdx = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INCREMENT_NUMBER);
  const numDataRows   = lastRow - dataStartRow + 1;
  const mainMerges    = sheet.getRange(dataStartRow, mainNumColIdx, numDataRows, 1).getMergedRanges();
  const mainMergeMap  = {};
  for (let i = 0; i < mainMerges.length; i++) {
    const m = mainMerges[i], top = m.getRow(), size = m.getNumRows();
    if (size < 2) continue;
    mainMergeMap[top] = { isTop: true };
    for (let r = top + 1; r < top + size; r++) mainMergeMap[r] = { isTop: false };
  }

  let mainCounter = 1;
  for (let row = lastRow; row >= dataStartRow; row--) {
    if (!isDataRow[row]) continue;
    const mainInfo = mainMergeMap[row];
    if (mainInfo && !mainInfo.isTop) continue;
    sheet.getRange(row, mainNumColIdx).setValue(mainCounter);
    mainCounter++;
  }
}

// ── J: sub numbers within each main group ────────────────────────────────────
function resequenceJNumbers(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const { dataStartRow, lastRow, isDataRow } = getDataRowMap(sheet);
  if (lastRow < dataStartRow) return;

  const mainNumColIdx = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INCREMENT_NUMBER);
  const subNumColIdx  = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_INCREMENT_COL);
  const numDataRows   = lastRow - dataStartRow + 1;

  // clear ALL J validations upfront — G-Q dropdown bleeds into J on default rows
  sheet.getRange(dataStartRow, subNumColIdx, numDataRows, 1).clearDataValidations();

  const mainMerges   = sheet.getRange(dataStartRow, mainNumColIdx, numDataRows, 1).getMergedRanges();
  const mainMergeMap = {};
  for (let i = 0; i < mainMerges.length; i++) {
    const m = mainMerges[i], top = m.getRow(), size = m.getNumRows();
    if (size < 2) continue;
    mainMergeMap[top] = { isTop: true, size };
    for (let r = top + 1; r < top + size; r++) mainMergeMap[r] = { isTop: false };
  }

  const subMerges   = sheet.getRange(dataStartRow, subNumColIdx, numDataRows, 1).getMergedRanges();
  const subMergeMap = {};
  for (let i = 0; i < subMerges.length; i++) {
    const m = subMerges[i], top = m.getRow(), size = m.getNumRows();
    if (size < 2) continue;
    subMergeMap[top] = { isTop: true, size };
    for (let r = top + 1; r < top + size; r++) subMergeMap[r] = { isTop: false };
  }

  for (let row = dataStartRow; row <= lastRow; row++) {
    if (!isDataRow[row]) continue;
    const mainInfo = mainMergeMap[row];
    if (!mainInfo || !mainInfo.isTop) continue;

    const groupTop = row;
    const groupEnd = groupTop + mainInfo.size - 1;
    const groupNum = sheet.getRange(groupTop, mainNumColIdx).getValue();

    // check if any H sub-groups exist inside this group
    let hasSubGroups = false;
    for (let r = groupTop; r <= groupEnd; r++) {
      if (subMergeMap[r] && subMergeMap[r].isTop) { hasSubGroups = true; break; }
    }

    if (!hasSubGroups) {
      // no H sub-groups: write J to every data row, bottom=.1 top=.highest
      let counter = 1;
      for (let r = groupEnd; r >= groupTop; r--) {
        if (!isDataRow[r]) continue;
        writeNum(sheet, r, subNumColIdx, groupNum + '.' + counter);
        counter++;
      }
    } else {
      // H sub-groups exist: slot logic — one number per slot (solo row or H sub-group)
      const slots   = [];
      const visited = new Set();
      for (let r = groupEnd; r >= groupTop; r--) {
        if (!isDataRow[r] || visited.has(r)) continue;
        const subInfo = subMergeMap[r];
        if (subInfo && !subInfo.isTop) { visited.add(r); continue; }
        if (subInfo && subInfo.isTop) {
          for (let sr = r; sr < r + subInfo.size; sr++) visited.add(sr);
          slots.push({ type: 'subgroup', topRow: r });
        } else {
          visited.add(r);
          slots.push({ type: 'solo', topRow: r });
        }
      }
      for (let s = 0; s < slots.length; s++) {
        writeNum(sheet, slots[s].topRow, subNumColIdx, groupNum + '.' + (s + 1));
      }
    }
  }
}

// ── O: sub-sub numbers within each H sub-group ───────────────────────────────
function resequenceONumbers(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const { dataStartRow, lastRow, isDataRow } = getDataRowMap(sheet);
  if (lastRow < dataStartRow) return;

  const subNumColIdx    = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_INCREMENT_COL);
  const subSubNumColIdx = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_SUB_INCREMENT_COL);
  const numDataRows     = lastRow - dataStartRow + 1;

  // clear ALL O validations upfront
  sheet.getRange(dataStartRow, subSubNumColIdx, numDataRows, 1).clearDataValidations();

  const subMerges   = sheet.getRange(dataStartRow, subNumColIdx, numDataRows, 1).getMergedRanges();
  const subMergeMap = {};
  for (let i = 0; i < subMerges.length; i++) {
    const m = subMerges[i], top = m.getRow(), size = m.getNumRows();
    if (size < 2) continue;
    subMergeMap[top] = { isTop: true, size };
  }

  for (let row = dataStartRow; row <= lastRow; row++) {
    if (!isDataRow[row]) continue;
    const subInfo = subMergeMap[row];
    if (!subInfo || !subInfo.isTop) continue;

    const sgTop  = row;
    const sgEnd  = sgTop + subInfo.size - 1;
    const subVal = sheet.getRange(sgTop, subNumColIdx).getValue();

    let counter = 1;
    for (let r = sgEnd; r >= sgTop; r--) {
      if (!isDataRow[r]) continue;
      writeNum(sheet, r, subSubNumColIdx, subVal + '.' + counter);
      sheet.getRange(r, subSubNumColIdx)
        .setBorder(true, true, true, true, true, true, '#ffffff', SpreadsheetApp.BorderStyle.SOLID);
      counter++;
    }
  }
}

// ── Master: call all three ────────────────────────────────────────────────────
function resequenceActivityNumbers(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  resequenceDNumbers(sheet);
  resequenceJNumbers(sheet);
  resequenceONumbers(sheet);
}

function manualResequenceActivityNumbers() {
  resequenceActivityNumbers(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet());
}
