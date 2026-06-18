function handleHasTimestampState(sheet, row, hasTimestampCol, timestampCol, label) {
  const hasTimestampCell = sheet.getRange(row, hasTimestampCol);
  const timestampCell = sheet.getRange(row, timestampCol);

  const hasTimestampValue = hasTimestampCell.getValue();

  if (hasTimestampValue === "No Status") {
    timestampCell.clearDataValidations();
    timestampCell.setValue(label + " Not Decided 🟡");
  } else if (hasTimestampValue === "Not Applicable") {
    timestampCell.clearDataValidations();
    timestampCell.setValue("Not Applicable 🔴");
  } else if (hasTimestampValue === "No") {
    timestampCell.clearDataValidations();
    timestampCell.setValue("No " + label + " 🔴");
  } else if (hasTimestampValue === "Yes" || hasTimestampValue === "Yes (Approx)") {
    const validation = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
    timestampCell.setDataValidation(validation);
    timestampCell.setValue(false);
  }
}
