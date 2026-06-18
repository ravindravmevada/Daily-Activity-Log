function setFeatureStatusCell(namedRangeKey, isEnabled) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statusRange = ss.getRangeByName(namedRangeKey);
  if (!statusRange) return;

  if (isEnabled) {
    statusRange.setValue("✅ Feature is enabled")
      .setBackground("#434343")
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setFontWeight("bold")
      .setFontStyle("normal")
      .setFontColor("#00ff00")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);
  } else {
    statusRange.setValue("✨ Feature Status ✨")
      .setBackground("#434343")
      .setFontFamily("Lexend")
      .setFontSize(11)
      .setFontWeight("normal")
      .setFontStyle("normal")
      .setFontColor("#f3f3f3")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);
  }
}

function setFeatureNotEnabledWarning(namedRangeKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statusRange = ss.getRangeByName(namedRangeKey);
  if (!statusRange) return;
  statusRange.setValue("⚠️ Enable the feature first")
    .setBackground("#434343")
    .setFontFamily("Lexend")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontStyle("italic")
    .setFontColor("#faab17")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
}
