function parseTimestampString(timestampStr) {
  if (typeof timestampStr !== "string" || timestampStr === "") return null;
  if (timestampStr.indexOf("Not Applicable") !== -1) return null;
  if (timestampStr.indexOf("Not Decided") !== -1) return null;

  const parts = timestampStr.split(" - ");
  if (parts.length !== 2) return null;

  const datePart = parts[0].trim();
  const timePart = parts[1].trim();

  const combined = datePart + " " + timePart;
  const parsed = new Date(combined);
  if (isNaN(parsed.getTime())) return null;
  return parsed;
}

function formatDurationMs(ms) {
  if (ms < 0) ms = 0;
  const totalSec = Math.floor(ms / 1000);
  const h = Math.floor(totalSec / 3600);
  const m = Math.floor((totalSec % 3600) / 60);
  const s = totalSec % 60;

  if (h > 0) return h + "h " + m + "m " + s + "s";
  if (m > 0) return m + "m " + s + "s";
  return s + "s";
}

function buildDurationRichText(durationText, fontColor) {
  const boldStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor(fontColor).build();
  const normalStyle = SpreadsheetApp.newTextStyle().setBold(false).setForegroundColor(fontColor).build();

  const builder = SpreadsheetApp.newRichTextValue().setText(durationText);

  for (let i = 0; i < durationText.length; i++) {
    const ch = durationText.charAt(i);
    if (ch === "h" || ch === "m" || ch === "s") {
      builder.setTextStyle(i, i + 1, normalStyle);
    } else if (ch >= "0" && ch <= "9") {
      builder.setTextStyle(i, i + 1, boldStyle);
    } else {
      builder.setTextStyle(i, i + 1, normalStyle);
    }
  }

  return builder.build();
}

function buildPlainRichText(text, fontColor) {
  const style = SpreadsheetApp.newTextStyle().setBold(false).setForegroundColor(fontColor).build();
  return SpreadsheetApp.newRichTextValue().setText(text).setTextStyle(style).build();
}

function calculateRowDuration(sheet, row) {
  const startCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[0].timestampCol);
  const endCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[1].timestampCol);

  const startVal = sheet.getRange(row, startCol).getValue();
  const endVal = sheet.getRange(row, endCol).getValue();

  const startDate = parseTimestampString(startVal);
  const endDate = parseTimestampString(endVal);

  if (!startDate || !endDate) return null;
  return endDate.getTime() - startDate.getTime();
}

function isTimestampFilled(value) {
  return parseTimestampString(value) !== null;
}

function computeRowDisplay(sheet, row) {
  const startHasCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[0].hasTimestampCol);
  const endHasCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[1].hasTimestampCol);
  const startTsCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[0].timestampCol);
  const endTsCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[1].timestampCol);

  const startHasVal = sheet.getRange(row, startHasCol).getValue();
  const endHasVal = sheet.getRange(row, endHasCol).getValue();
  const startTsVal = sheet.getRange(row, startTsCol).getValue();
  const endTsVal = sheet.getRange(row, endTsCol).getValue();

  const isYesLike = function(v) { return v === "Yes" || v === "Yes (Approx.)"; };
  const isNoLike = function(v) { return v === "No" || v === "Not Applicable"; };

  if (isNoLike(startHasVal) || isNoLike(endHasVal)) {
    return { text: "🔴", category: "RED", durMs: null };
  }

  const startYesWithTs = isYesLike(startHasVal) && isTimestampFilled(startTsVal);
  const endYesWithTs = isYesLike(endHasVal) && isTimestampFilled(endTsVal);

  if (isYesLike(startHasVal) && isYesLike(endHasVal) && startYesWithTs && endYesWithTs) {
    const startDate = parseTimestampString(startTsVal);
    const endDate = parseTimestampString(endTsVal);
    const durMs = endDate.getTime() - startDate.getTime();
    return { text: formatDurationMs(durMs), category: "GREEN", durMs: durMs };
  }

  if (startYesWithTs || endYesWithTs) {
    return { text: "🟢", category: "GREEN", durMs: null };
  }

  return { text: "🟡", category: "YELLOW", durMs: null };
}

function findGroupRange(sheet, row) {
  const labelMergeStartCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_LABEL_MERGE_START_COL);

  const labelCell = sheet.getRange(row, labelMergeStartCol);
  const mergedRanges = labelCell.getMergedRanges();
  if (!mergedRanges || mergedRanges.length === 0) return null;

  const m = mergedRanges[0];
  const startRow = m.getRow();
  const numRows = m.getNumRows();
  if (numRows < 2) return null;

  return { startRow: startRow, numRows: numRows };
}

function isDurationText(text) {
  return text.indexOf("h ") !== -1 || text.indexOf("m ") !== -1 || text.charAt(text.length - 1) === "s";
}

function writeCellWithStyle(range, text, fontColor) {
  let richValue;
  if (isDurationText(text)) {
    richValue = buildDurationRichText(text, fontColor);
  } else {
    richValue = buildPlainRichText(text, fontColor);
  }
  range.setRichTextValue(richValue);
  range.setHorizontalAlignment("center");
  range.setVerticalAlignment("middle");
}

/*
function updateRowDuration(sheet, row) {
  const perRowCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.PER_ROW_DURATION_COL);
  const groupTotalCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.GROUP_TOTAL_DURATION_COL);
  const subRowFontColor = DAILY_ACTIVITY_LOG_DURATION_COLS.SUB_ROW_FONT_COLOR;
  const mergedResultFontColor = DAILY_ACTIVITY_LOG_DURATION_COLS.MERGED_RESULT_FONT_COLOR;

  const groupRange = findGroupRange(sheet, row);

  if (groupRange === null) {
    const horizontalMergeRange = sheet.getRange(row, perRowCol, 1, 2);
    horizontalMergeRange.breakApart();
    horizontalMergeRange.merge();

    const display = computeRowDisplay(sheet, row);
    writeCellWithStyle(horizontalMergeRange, display.text, mergedResultFontColor);

    horizontalMergeRange.setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
    return;
  }

  const startRow = groupRange.startRow;
  const numRows = groupRange.numRows;

  const aeAfBlock = sheet.getRange(startRow, perRowCol, numRows, 2);
  aeAfBlock.breakApart();

  let totalMs = 0;
  let hasAnyDuration = false;
  let hasRed = false;
  let allGreen = true;

  for (let r = 0; r < numRows; r++) {
    const subRow = startRow + r;
    const aeCell = sheet.getRange(subRow, perRowCol);

    const display = computeRowDisplay(sheet, subRow);
    writeCellWithStyle(aeCell, display.text, subRowFontColor);

    if (display.category === "RED") hasRed = true;
    if (display.category !== "GREEN") allGreen = false;

    if (display.durMs !== null) {
      totalMs += display.durMs;
      hasAnyDuration = true;
    }
  }

  const afMergeRange = sheet.getRange(startRow, groupTotalCol, numRows, 1);
  afMergeRange.merge();

  let afValue;
  if (hasAnyDuration) {
    afValue = formatDurationMs(totalMs);
  } else if (hasRed) {
    afValue = "🔴";
  } else if (allGreen) {
    afValue = "🟢";
  } else {
    afValue = "🟡";
  }

  writeCellWithStyle(afMergeRange, afValue, mergedResultFontColor);

  sheet.getRange(startRow, perRowCol, numRows, 1)
    .setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
  afMergeRange
    .setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
}*/

function updateRowDuration(sheet, row) {
  const perRowCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.PER_ROW_DURATION_COL);
  const groupTotalCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.GROUP_TOTAL_DURATION_COL);
  const activityTypeCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.ACTIVITY_TYPE_COL);
  const parallelValue = DAILY_ACTIVITY_LOG_DURATION_COLS.PARALLEL_VALUE;
  const subRowFontColor = DAILY_ACTIVITY_LOG_DURATION_COLS.SUB_ROW_FONT_COLOR;
  const mergedResultFontColor = DAILY_ACTIVITY_LOG_DURATION_COLS.MERGED_RESULT_FONT_COLOR;
  const parallelFontColor = DAILY_ACTIVITY_LOG_DURATION_COLS.PARALLEL_FONT_COLOR;

  const groupRange = findGroupRange(sheet, row);

  if (groupRange === null) {
    const startHasCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[0].hasTimestampCol);
    const endHasCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[1].hasTimestampCol);
    const startHasVal = sheet.getRange(row, startHasCol).getValue();
    const endHasVal = sheet.getRange(row, endHasCol).getValue();

    const isYesLike = function(v) { return v === "Yes" || v === "Yes (Approx)" || v === "Yes (Approx.)"; };
    const isNoLike = function(v) { return v === "No" || v === "Not Applicable"; };
    const allNoStatus = !isYesLike(startHasVal) && !isYesLike(endHasVal) && !isNoLike(startHasVal) && !isNoLike(endHasVal);

    const pairRange = sheet.getRange(row, perRowCol, 1, 2);
    const isParallel = (sheet.getRange(row, activityTypeCol).getValue() === parallelValue);

    if (allNoStatus) {
      pairRange.breakApart();
      writeCellWithStyle(sheet.getRange(row, perRowCol), "🟡", mergedResultFontColor);
      writeCellWithStyle(sheet.getRange(row, groupTotalCol), "🟡", mergedResultFontColor);
      pairRange.setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
    } else {
      pairRange.breakApart();
      pairRange.merge();
      const display = computeRowDisplay(sheet, row);
      const indFontColor = isParallel ? parallelFontColor : mergedResultFontColor;
      writeCellWithStyle(pairRange, display.text, indFontColor);
      pairRange.setBorder(true, true, true, true, false, false, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
    }
    return;
  }

  const startRow = groupRange.startRow;
  const numRows = groupRange.numRows;

  const aeAfBlock = sheet.getRange(startRow, perRowCol, numRows, 2);
  aeAfBlock.breakApart();

  let totalMs = 0;
  let hasAnyDuration = false;
  let hasRed = false;
  let allGreen = true;

  for (let r = 0; r < numRows; r++) {
    const subRow = startRow + r;
    const aeCell = sheet.getRange(subRow, perRowCol);

    const display = computeRowDisplay(sheet, subRow);
    const isParallel = (sheet.getRange(subRow, activityTypeCol).getValue() === parallelValue);
    const rowFontColor = isParallel ? parallelFontColor : subRowFontColor;
    writeCellWithStyle(aeCell, display.text, rowFontColor);

    if (display.category === "RED") hasRed = true;
    if (display.category !== "GREEN") allGreen = false;

    if (display.durMs !== null && !isParallel) {
      totalMs += display.durMs;
      hasAnyDuration = true;
    }
  }

  const afMergeRange = sheet.getRange(startRow, groupTotalCol, numRows, 1);
  afMergeRange.merge();

  let afValue;
  if (hasAnyDuration) {
    afValue = formatDurationMs(totalMs);
  } else if (hasRed) {
    afValue = "🔴";
  } else if (allGreen) {
    afValue = "🟢";
  } else {
    afValue = "🟡";
  }

  writeCellWithStyle(afMergeRange, afValue, mergedResultFontColor);

  sheet.getRange(startRow, perRowCol, numRows, 1)
    .setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
  afMergeRange
    .setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
}
