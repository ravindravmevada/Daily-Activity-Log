function isDailyLogSheet(sheetName) {
  return SHEET_REGISTRY.DAILY_ACTIVITY_LOG_SHEET_NAME_PATTERN.test(sheetName) ||
         SHEET_REGISTRY.DAILY_ACTIVITY_LOG_EXTRA_SHEET_NAMES.includes(sheetName);
}
