function handleTickboxTimestampToggle(sheet, row, tickboxCol) {
  const capturedTimestamp = getFormattedDateTime();

  const tickboxCell = sheet.getRange(row, tickboxCol);
  const tickValue = tickboxCell.getValue();

  if (tickValue === true) {
    tickboxCell.clearDataValidations();
    tickboxCell.setValue(capturedTimestamp);
  } else if (tickValue === "" || tickValue === null) {
    const validation = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
    tickboxCell.setDataValidation(validation);
    tickboxCell.setValue(false);
  }
}
