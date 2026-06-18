function handleDeviceViaCascade(sheet, row) {
  const fromCol = columnLetterToIndex(DEVICE_VIA_CASCADE.LOGGED_FROM_COL);
  const viaCol = columnLetterToIndex(DEVICE_VIA_CASCADE.LOGGED_VIA_COL);

  const deviceVal = sheet.getRange(row, fromCol).getValue();
  if (!deviceVal || deviceVal === "") return;

  const defaultVia = getDefaultViaForDevice(deviceVal);
  if (defaultVia === null || defaultVia === "") return;

  sheet.getRange(row, viaCol).setValue(defaultVia);
}

function getDefaultViaForDevice(deviceName) {
  if (!deviceName || deviceName === "") return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(DEVICE_VIA_CASCADE.SOURCE_SHEET);
  if (!sourceSheet) return null;

  const headerRow = DEVICE_VIA_CASCADE.HEADER_ROW;
  const dataStartRow = DEVICE_VIA_CASCADE.DATA_START_ROW;
  const lastCol = sourceSheet.getLastColumn();
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < dataStartRow) return null;

  const headers = sourceSheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  let deviceColIndex = -1;
  let viaColIndex = -1;
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === DEVICE_VIA_CASCADE.DEVICE_HEADER) deviceColIndex = i + 1;
    if (headers[i] === DEVICE_VIA_CASCADE.VIA_HEADER) viaColIndex = i + 1;
  }
  if (deviceColIndex === -1 || viaColIndex === -1) return null;

  const numRows = lastRow - dataStartRow + 1;
  const devices = sourceSheet.getRange(dataStartRow, deviceColIndex, numRows, 1).getValues();
  const vias = sourceSheet.getRange(dataStartRow, viaColIndex, numRows, 1).getValues();

  for (let i = 0; i < devices.length; i++) {
    if (devices[i][0] === deviceName) return vias[i][0];
  }
  return null;
}
