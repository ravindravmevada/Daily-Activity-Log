function getSubcategoriesForCategory(categoryName) {
  if (!categoryName || categoryName === "") return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEET_NAMES.LINKED_DROPDOWN_SOURCE_SHEET);
  if (!sourceSheet) return [];

  const headerRow = LINKED_DROPDOWN_SOURCE_SHEET_CONFIG.HEADER_ROW;
  const dataStartRow = LINKED_DROPDOWN_SOURCE_SHEET_CONFIG.DATA_START_ROW;
  const lastCol = sourceSheet.getLastColumn();
  const lastRow = sourceSheet.getLastRow();

  const headers = sourceSheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  let matchedColIndex = -1;
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === categoryName) {
      matchedColIndex = i + 1;
      break;
    }
  }

  if (matchedColIndex === -1) return [];
  if (lastRow < dataStartRow) return [];

  const subcategoryValues = sourceSheet.getRange(dataStartRow, matchedColIndex, lastRow - dataStartRow + 1, 1).getValues();

  const subcategories = [];
  for (let i = 0; i < subcategoryValues.length; i++) {
    const v = subcategoryValues[i][0];
    if (v !== "" && v !== null) {
      subcategories.push(v);
    }
  }

  return subcategories;
}

function getSubcategorySourceRange(categoryName) {
  if (!categoryName || categoryName === "") return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEET_NAMES.LINKED_DROPDOWN_SOURCE_SHEET);
  if (!sourceSheet) return null;

  const headerRow = LINKED_DROPDOWN_SOURCE_SHEET_CONFIG.HEADER_ROW;
  const dataStartRow = LINKED_DROPDOWN_SOURCE_SHEET_CONFIG.DATA_START_ROW;
  const lastCol = sourceSheet.getLastColumn();
  const headers = sourceSheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  let matchedColIndex = -1;
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === categoryName) { matchedColIndex = i + 1; break; }
  }
  if (matchedColIndex === -1) return null;

  const maxRows = sourceSheet.getMaxRows();
  const numRows = maxRows - dataStartRow + 1;
  if (numRows < 1) return null;

  return sourceSheet.getRange(dataStartRow, matchedColIndex, numRows, 1);
}
