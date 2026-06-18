function handleHasTimestampStatus(sheet, row, hasTimestampCol, timestampStatusCol) {
  const hasTimestampCell = sheet.getRange(row, hasTimestampCol);
  const timestampStatusCell = sheet.getRange(row, timestampStatusCol);

  const hasTimestampValue = hasTimestampCell.getValue();

  if (hasTimestampValue === "No Status") {
    timestampStatusCell.setValue("🟨");
  } else if (hasTimestampValue === "Not Applicable" || hasTimestampValue === "No") {
    timestampStatusCell.setValue("🟥");
  } else if (hasTimestampValue === "Yes" || hasTimestampValue === "Yes (Approx)") {
    timestampStatusCell.setValue("🟩");
  }
}
