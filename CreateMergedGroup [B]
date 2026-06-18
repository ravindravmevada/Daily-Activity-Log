function handleCreateMergedGroupFromRange(sheet, row, col) {
  const editedCell = sheet.getRange(row, col);
  const value = editedCell.getValue();
  if (value !== true) return;

  const lastRow = sheet.getLastRow();
  const triggerColRange = sheet.getRange(1, col, lastRow, 1);
  const triggerValues = triggerColRange.getValues();

  const tickedRows = [];
  for (let i = 0; i < triggerValues.length; i++) {
    if (triggerValues[i][0] === true) {
      tickedRows.push(i + 1);
    }
  }

  if (tickedRows.length < 2) return;

  let groupStartRow = Math.min.apply(null, tickedRows);
  let groupEndRow = Math.max.apply(null, tickedRows);

  const labelColForExpand = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_START_COL);
  for (let i = 0; i < tickedRows.length; i++) {
    const mr = sheet.getRange(tickedRows[i], labelColForExpand).getMergedRanges();
    if (mr && mr.length > 0) {
      const m = mr[0];
      const top = m.getRow();
      const bottom = m.getRow() + m.getNumRows() - 1;
      if (top < groupStartRow) groupStartRow = top;
      if (bottom > groupEndRow) groupEndRow = bottom;
    }
  }

  const groupNumRows = groupEndRow - groupStartRow + 1;

  const incrementColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INCREMENT_NUMBER);
  const subIncrementColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_INCREMENT_COL);
  const verticalMergeColIndexes = DAILY_ACTIVITY_LOG_COLS.GROUP_VERTICAL_MERGE_COLS.map(columnLetterToIndex);
  const labelMergeStartColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_START_COL);
  const labelMergeEndColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_END_COL);
  const labelMergeNumCols = labelMergeEndColIndex - labelMergeStartColIndex + 1;
  const individualTickboxColIndexes = DAILY_ACTIVITY_LOG_COLS.GROUP_INDIVIDUAL_TICKBOX_COLS.map(columnLetterToIndex);
  const dropdownNarrowedStartColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_ROW_DROPDOWN_NARROWED_START_COL);
  const dropdownEndColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_ROW_DROPDOWN_END_COL);
  const dropdownNarrowedNumCols = dropdownEndColIndex - dropdownNarrowedStartColIndex + 1;

  const rowBelowGroupD = sheet.getRange(groupEndRow + 1, incrementColIndex).getValue();
  let newGroupDValue = null;
  if (typeof rowBelowGroupD === "number") {
    newGroupDValue = rowBelowGroupD + 1;
  }

  const groupColumnEnd = dropdownEndColIndex;
  const groupRowRange = sheet.getRange(groupStartRow, 1, groupNumRows, groupColumnEnd);
  const existingRowMerges = groupRowRange.getMergedRanges();
  const mergesToRestore = [];

  for (let i = 0; i < existingRowMerges.length; i++) {
    const m = existingRowMerges[i];
    const mergeEndCol = m.getColumn() + m.getNumColumns() - 1;
    if (mergeEndCol > groupColumnEnd) {
      mergesToRestore.push({
        startRow: m.getRow(),
        startCol: m.getColumn(),
        numRows: m.getNumRows(),
        numCols: m.getNumColumns()
      });
    }
    m.breakApart();
  }

  for (let i = 0; i < tickedRows.length; i++) {
    sheet.getRange(tickedRows[i], col).setValue(false);
  }

  for (let i = 0; i < verticalMergeColIndexes.length; i++) {
    const mergeColIndex = verticalMergeColIndexes[i];
    const targetRange = sheet.getRange(groupStartRow, mergeColIndex, groupNumRows, 1);
    targetRange.merge();
    targetRange.setVerticalAlignment("middle");
  }

  const labelMergeRange = sheet.getRange(groupStartRow, labelMergeStartColIndex, groupNumRows, labelMergeNumCols);
  labelMergeRange.merge();
  labelMergeRange.setVerticalAlignment("middle");

  if (newGroupDValue !== null) {
    sheet.getRange(groupStartRow, incrementColIndex).setValue(newGroupDValue);
  }

  const tickboxValidation = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
  for (let i = 0; i < individualTickboxColIndexes.length; i++) {
    const tickboxColIndex = individualTickboxColIndexes[i];
    const tickboxColLetter = DAILY_ACTIVITY_LOG_COLS.GROUP_INDIVIDUAL_TICKBOX_COLS[i];
    const tickboxRange = sheet.getRange(groupStartRow, tickboxColIndex, groupNumRows, 1);
    tickboxRange.setDataValidation(tickboxValidation);
    tickboxRange.setValue(false);

    if (tickboxColLetter === "G") {
      tickboxRange.setFontColor("#ff6d01");
    } else if (tickboxColLetter === "H") {
      tickboxRange.setFontColor("#f7ef00");
    }
  }

  const groupDForSubNumbers = newGroupDValue !== null ? newGroupDValue : sheet.getRange(groupStartRow, incrementColIndex).getValue();
  const subIncrementRange = sheet.getRange(groupStartRow, subIncrementColIndex, groupNumRows, 1);
  subIncrementRange.clearDataValidations();
  for (let r = 0; r < groupNumRows; r++) {
    const positionFromBottom = groupNumRows - r;
    const subNumber = parseFloat(groupDForSubNumbers + "." + positionFromBottom);
    sheet.getRange(groupStartRow + r, subIncrementColIndex).setValue(subNumber);
  }
  subIncrementRange.setFontWeight("bold");
  subIncrementRange.setFontColor("#f7ef00");

  for (let r = 0; r < groupNumRows; r++) {
    const actualRow = groupStartRow + r;
    const narrowedRange = sheet.getRange(actualRow, dropdownNarrowedStartColIndex, 1, dropdownNarrowedNumCols);
    narrowedRange.merge();
  }

  const labelCellValue = sheet.getRange(groupStartRow, labelMergeStartColIndex).getValue();
  if (labelCellValue && labelCellValue !== "") {
    applySubcategoryDropdownToGroupRows(sheet, groupStartRow, groupNumRows, labelCellValue);
  }

  for (let i = 0; i < mergesToRestore.length; i++) {
    const m = mergesToRestore[i];
    for (let r = 0; r < m.numRows; r++) {
      const restoreRange = sheet.getRange(m.startRow + r, m.startCol, 1, m.numCols);
      restoreRange.merge();
    }
  }

  for (let i = 0; i < individualTickboxColIndexes.length; i++) {
    const tickboxColIndex = individualTickboxColIndexes[i];
    const tickboxRange = sheet.getRange(groupStartRow, tickboxColIndex, groupNumRows, 1);
    tickboxRange.setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
  }

  const subBorderRange = sheet.getRange(groupStartRow, subIncrementColIndex, groupNumRows, 1);
  subBorderRange.setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);

  for (let r = 0; r < groupNumRows; r++) {
    const actualRow = groupStartRow + r;
    const narrowedRange = sheet.getRange(actualRow, dropdownNarrowedStartColIndex, 1, dropdownNarrowedNumCols);
    narrowedRange.setBorder(true, true, true, true, false, false, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
  }

  updateRowDuration(sheet, groupStartRow);
}
