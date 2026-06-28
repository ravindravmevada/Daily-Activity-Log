function resequenceActivityNumbers(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const dataStartRow    = 3;
  const lastRow         = sheet.getLastRow();
  if (lastRow < dataStartRow) return;

  const insertColIdx    = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INSERT_ROW_ABOVE);
  const mainNumColIdx   = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INCREMENT_NUMBER);
  const subNumColIdx    = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_INCREMENT_COL);
  const subSubNumColIdx = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_SUB_INCREMENT_COL);

  // detect data rows via checkbox in col A
  const numDataRows   = lastRow - dataStartRow + 1;
  const indicatorVals = sheet.getRange(dataStartRow, insertColIdx, numDataRows, 1).getDataValidations();
  const isDataRow     = {};
  for (let i = 0; i < numDataRows; i++) {
    const v = indicatorVals[i][0];
    if (v && v.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      isDataRow[dataStartRow + i] = true;
    }
  }

  // read main group merges via col D
  const mainMerges   = sheet.getRange(dataStartRow, mainNumColIdx, numDataRows, 1).getMergedRanges();
  const mainMergeMap = {};
  for (let i = 0; i < mainMerges.length; i++) {
    const m = mainMerges[i], top = m.getRow(), size = m.getNumRows();
    if (size < 2) continue;
    mainMergeMap[top] = { isTop: true, size };
    for (let r = top + 1; r < top + size; r++) mainMergeMap[r] = { isTop: false, parentTop: top };
  }

  // read sub-group merges via col J
  const subMerges   = sheet.getRange(dataStartRow, subNumColIdx, numDataRows, 1).getMergedRanges();
  const subMergeMap = {};
  for (let i = 0; i < subMerges.length; i++) {
    const m = subMerges[i], top = m.getRow(), size = m.getNumRows();
    if (size < 2) continue;
    subMergeMap[top] = { isTop: true, size };
    for (let r = top + 1; r < top + size; r++) subMergeMap[r] = { isTop: false, parentTop: top };
  }

  // write text number preserving e.g. 67.10
  function writeNum(row, col, value) {
    const cell = sheet.getRange(row, col);
    cell.setNumberFormat('@');
    cell.setValue(String(value));
    cell.setFontWeight('bold').setFontColor('#f7ef00');
  }

  // ── PASS 1: assign main numbers bottom to top ─────────────────────────────
  let mainCounter = 1;
  for (let row = lastRow; row >= dataStartRow; row--) {
    if (!isDataRow[row]) continue;
    const mainInfo = mainMergeMap[row];
    if (mainInfo && !mainInfo.isTop) continue; // skip non-top group members
    sheet.getRange(row, mainNumColIdx).setValue(mainCounter);
    mainCounter++;
  }

  // ── PASS 2: sub numbers within each main group ────────────────────────────
  for (let row = dataStartRow; row <= lastRow; row++) {
    if (!isDataRow[row]) continue;
    const mainInfo = mainMergeMap[row];
    if (!mainInfo || !mainInfo.isTop) continue;

    const groupTop  = row;
    const groupSize = mainInfo.size;
    const groupEnd  = groupTop + groupSize - 1;
    const groupNum  = sheet.getRange(groupTop, mainNumColIdx).getValue();

    // build slots scanning BOTTOM TO TOP
    // slot = one logical entry in J column (solo row OR sub-group top)
    // slots[0] = bottom-most slot → gets counter 1
    // slots[last] = top-most slot → gets highest counter
    const slots = [];
    const visited = new Set();

    for (let r = groupEnd; r >= groupTop; r--) {
      if (!isDataRow[r]) continue;
      if (visited.has(r)) continue;

      const subInfo = subMergeMap[r];

      if (subInfo && !subInfo.isTop) {
        // non-top sub-group member — skip, parent top handles it
        visited.add(r);
        continue;
      }

      if (subInfo && subInfo.isTop) {
        // sub-group top — collect its member rows bottom to top
        const sgTop  = r;
        const sgSize = subInfo.size;
        const sgRows = [];
        for (let sr = sgTop + sgSize - 1; sr >= sgTop; sr--) {
          if (isDataRow[sr]) sgRows.push(sr);
          visited.add(sr);
        }
        // sgRows is bottom-to-top order for sub-sub numbering
        slots.push({ type: 'subgroup', topRow: sgTop, rows: sgRows });
      } else {
        // solo row
        visited.add(r);
        slots.push({ type: 'solo', topRow: r });
      }
    }

    // slots[0] = bottom slot → counter 1
    // slots[last] = top slot → counter = slots.length
    for (let s = 0; s < slots.length; s++) {
      const slot     = slots[s];
      const subCount = s + 1; // 1 = bottom, highest = top
      const subVal   = groupNum + '.' + subCount;

      if (slot.type === 'solo') {
        writeNum(slot.topRow, subNumColIdx, subVal);
      } else {
        // write sub val to J top cell of sub-group
        writeNum(slot.topRow, subNumColIdx, subVal);
        // sub-sub numbers: rows is bottom-to-top, so rows[0]=bottom=1
        for (let si = 0; si < slot.rows.length; si++) {
          const ssCount = si + 1; // bottom=1, top=highest
          writeNum(slot.rows[si], subSubNumColIdx, subVal + '.' + ssCount);
          sheet.getRange(slot.rows[si], subSubNumColIdx)
            .setBorder(true, true, true, true, true, true, '#ffffff', SpreadsheetApp.BorderStyle.SOLID);
        }
      }
    }
  }
}

function manualResequenceActivityNumbers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  resequenceActivityNumbers(sheet);
}
