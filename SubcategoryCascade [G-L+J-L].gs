function handleCategoryChangeCascade(sheet, row, col) {
  const labelMergeStartColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_START_COL);
  const labelCell = sheet.getRange(row, labelMergeStartColIndex);
  const labelMerges = labelCell.getMergedRanges();

  let groupStartRow, groupNumRows, isMergedGroup;
  if (labelMerges && labelMerges.length > 0 && labelMerges[0].getNumRows() > 1) {
    const m = labelMerges[0];
    groupStartRow = m.getRow();
    groupNumRows = m.getNumRows();
    isMergedGroup = true;
  } else {
    groupStartRow = row;
    groupNumRows = 1;
    isMergedGroup = false;
  }

  const categoryName = sheet.getRange(groupStartRow, labelMergeStartColIndex).getValue();

  if (isMergedGroup) {
    applySubcategoryDropdownToGroupRows(sheet, groupStartRow, groupNumRows, categoryName);
  } else {
    applySubcategoryDropdownToUnmergedRow(sheet, groupStartRow, categoryName);
  }
}

function applySubcategoryDropdownToGroupRows(sheet, groupStartRow, groupNumRows, categoryName) {
  const dropdownNarrowedStartColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_ROW_DROPDOWN_NARROWED_START_COL);
  const dropdownEndColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_ROW_DROPDOWN_END_COL);
  const dropdownNarrowedNumCols = dropdownEndColIndex - dropdownNarrowedStartColIndex + 1;

  const subcategories = getSubcategoriesForCategory(categoryName);

  for (let r = 0; r < groupNumRows; r++) {
    const actualRow = groupStartRow + r;
    const narrowedRange = sheet.getRange(actualRow, dropdownNarrowedStartColIndex, 1, dropdownNarrowedNumCols);

    if (subcategories.length > 0) {
      const validation = SpreadsheetApp.newDataValidation()
        .requireValueInList(subcategories, true)
        .setAllowInvalid(false)
        .build();
      narrowedRange.setDataValidation(validation);
      narrowedRange.clearContent();
      narrowedRange.setValue(subcategories[0]);
    } else {
      narrowedRange.clearDataValidations();
      narrowedRange.clearContent();
    }
  }
}

function applySubcategoryDropdownToGroupRows(sheet, groupStartRow, groupNumRows, categoryName) {
  const dropdownNarrowedStartColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_ROW_DROPDOWN_NARROWED_START_COL);
  const dropdownEndColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_ROW_DROPDOWN_END_COL);
  const dropdownNarrowedNumCols = dropdownEndColIndex - dropdownNarrowedStartColIndex + 1;

  const subcategories = getSubcategoriesForCategory(categoryName);
  const sourceRange = getSubcategorySourceRange(categoryName);

  for (let r = 0; r < groupNumRows; r++) {
    const actualRow = groupStartRow + r;
    const narrowedRange = sheet.getRange(actualRow, dropdownNarrowedStartColIndex, 1, dropdownNarrowedNumCols);

    if (subcategories.length > 0 && sourceRange) {
      const validation = SpreadsheetApp.newDataValidation()
        .requireValueInRange(sourceRange, true)
        .setAllowInvalid(false)
        .build();
      narrowedRange.setDataValidation(validation);
      narrowedRange.clearContent();
      narrowedRange.setValue(subcategories[0]);
    } else {
      narrowedRange.clearDataValidations();
      narrowedRange.clearContent();
    }
  }
}

function applySubcategoryDropdownToUnmergedRow(sheet, row, categoryName) {
  const dropdownStartColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.UNMERGED_ROW_DROPDOWN_START_COL);
  const dropdownEndColIndex = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.UNMERGED_ROW_DROPDOWN_END_COL);
  const dropdownNumCols = dropdownEndColIndex - dropdownStartColIndex + 1;

  const subcategories = getSubcategoriesForCategory(categoryName);
  const sourceRange = getSubcategorySourceRange(categoryName);

  const dropdownRange = sheet.getRange(row, dropdownStartColIndex, 1, dropdownNumCols);
  dropdownRange.clearDataValidations();
  dropdownRange.clearContent();

  const topLeftCell = sheet.getRange(row, dropdownStartColIndex);

  if (subcategories.length > 0 && sourceRange) {
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sourceRange, true)
      .setAllowInvalid(false)
      .build();
    topLeftCell.setDataValidation(validation);
    topLeftCell.setValue(subcategories[0]);
  }
}
