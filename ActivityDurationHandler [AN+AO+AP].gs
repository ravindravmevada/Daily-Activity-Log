function parseTimestampString(timestampStr) {
  if (typeof timestampStr !== "string" || timestampStr === "") return null;
  if (timestampStr.indexOf("Not Applicable") !== -1) return null;
  if (timestampStr.indexOf("Not Decided") !== -1) return null;
  const parts = timestampStr.split(" - ");
  if (parts.length !== 2) return null;
  const parsed = new Date(parts[0].trim() + " " + parts[1].trim());
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

function buildDurationRichText(text, fontColor) {
  const bold   = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor(fontColor).build();
  const normal = SpreadsheetApp.newTextStyle().setBold(false).setForegroundColor(fontColor).build();
  const builder = SpreadsheetApp.newRichTextValue().setText(text);
  for (let i = 0; i < text.length; i++) {
    const ch = text.charAt(i);
    builder.setTextStyle(i, i + 1, (ch >= "0" && ch <= "9") ? bold : normal);
  }
  return builder.build();
}

function buildPlainRichText(text, fontColor) {
  const style = SpreadsheetApp.newTextStyle().setBold(false).setForegroundColor(fontColor).build();
  return SpreadsheetApp.newRichTextValue().setText(text).setTextStyle(style).build();
}

function isDurationText(text) {
  return text.indexOf("h ") !== -1 || text.indexOf("m ") !== -1 || text.charAt(text.length - 1) === "s";
}

function writeCellWithStyle(range, text, fontColor) {
  const rich = isDurationText(text) ? buildDurationRichText(text, fontColor) : buildPlainRichText(text, fontColor);
  range.setRichTextValue(rich);
  range.setHorizontalAlignment("center");
  range.setVerticalAlignment("middle");
}

function isTimestampFilled(value) {
  return parseTimestampString(value) !== null;
}

function computeRowDisplay(sheet, row) {
  const startHasCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[0].hasTimestampCol);
  const endHasCol   = columnLetterToIndex(DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[1].hasTimestampCol);
  const startTsCol  = columnLetterToIndex(DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[0].timestampCol);
  const endTsCol    = columnLetterToIndex(DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS[1].timestampCol);

  const startHasVal = sheet.getRange(row, startHasCol).getValue();
  const endHasVal   = sheet.getRange(row, endHasCol).getValue();
  const startTsVal  = sheet.getRange(row, startTsCol).getValue();
  const endTsVal    = sheet.getRange(row, endTsCol).getValue();

  const isYesLike = v => v === "Yes" || v === "Yes (Approx)" || v === "Yes (Approx.)";
  const isNoLike  = v => v === "No" || v === "Not Applicable";

  if (isNoLike(startHasVal) || isNoLike(endHasVal))
    return { text: "🔴", category: "RED", durMs: null };

  const startFilled = isYesLike(startHasVal) && isTimestampFilled(startTsVal);
  const endFilled   = isYesLike(endHasVal)   && isTimestampFilled(endTsVal);

  if (isYesLike(startHasVal) && isYesLike(endHasVal) && startFilled && endFilled) {
    const durMs = parseTimestampString(endTsVal).getTime() - parseTimestampString(startTsVal).getTime();
    return { text: formatDurationMs(durMs), category: "GREEN", durMs };
  }
  if (startFilled || endFilled) return { text: "🟢", category: "GREEN", durMs: null };
  return { text: "🟡", category: "YELLOW", durMs: null };
}

function findMainGroupRange(sheet, row) {
  const mainNumCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INCREMENT_NUMBER);
  try {
    const merges = sheet.getRange(row, mainNumCol).getMergedRanges();
    if (!merges || merges.length === 0) return null;
    const m = merges[0];
    if (m.getNumRows() < 2) return null;
    return { startRow: m.getRow(), numRows: m.getNumRows() };
  } catch(e) { return null; }
}

function findSubGroupRange(sheet, row) {
  const subNumCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.GROUP_SUB_INCREMENT_COL);
  try {
    const merges = sheet.getRange(row, subNumCol).getMergedRanges();
    if (!merges || merges.length === 0) return null;
    const m = merges[0];
    if (m.getNumRows() < 2) return null;
    return { startRow: m.getRow(), numRows: m.getNumRows() };
  } catch(e) { return null; }
}

// ── Called when timestamp edited → update AN only ─────────────────────────
function updateRowDuration(sheet, row) {
  const anCol      = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.PER_ROW_DURATION_COL);
  const actTypeCol = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.ACTIVITY_TYPE_COL);
  const parallelVal = DAILY_ACTIVITY_LOG_DURATION_COLS.PARALLEL_VALUE;
  const perRowColor = DAILY_ACTIVITY_LOG_DURATION_COLS.SUB_ROW_FONT_COLOR;
  const mainColor   = DAILY_ACTIVITY_LOG_DURATION_COLS.MERGED_RESULT_FONT_COLOR;
  const parColor    = DAILY_ACTIVITY_LOG_DURATION_COLS.PARALLEL_FONT_COLOR;
  const WHITE       = "#ffffff";
  const SOLID       = SpreadsheetApp.BorderStyle.SOLID;

  const isParallel = sheet.getRange(row, actTypeCol).getValue() === parallelVal;
  const display    = computeRowDisplay(sheet, row);
  const mainGroup  = findMainGroupRange(sheet, row);
  const color      = isParallel ? parColor : (mainGroup ? perRowColor : mainColor);

  const anCell = sheet.getRange(row, anCol);
  writeCellWithStyle(anCell, display.text, color);
  anCell.setBorder(true, true, true, true, true, true, WHITE, SOLID);
}

// ── Called when B creates main group → AN-AO horizontal per row + AP vertical
function updateMainGroupDuration(sheet, groupStartRow, groupNumRows) {
  const anCol     = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.PER_ROW_DURATION_COL);
  const aoCol     = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.SUB_GROUP_TOTAL_DURATION_COL);
  const apCol     = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.GROUP_TOTAL_DURATION_COL);
  const mainColor = DAILY_ACTIVITY_LOG_DURATION_COLS.MERGED_RESULT_FONT_COLOR;
  const parColor  = DAILY_ACTIVITY_LOG_DURATION_COLS.PARALLEL_FONT_COLOR;
  const WHITE     = "#ffffff";
  const SOLID     = SpreadsheetApp.BorderStyle.SOLID;
  const actTypeCol  = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.ACTIVITY_TYPE_COL);
  const parallelVal = DAILY_ACTIVITY_LOG_DURATION_COLS.PARALLEL_VALUE;

  // safely break all AN-AO-AP for full group range first
  const fullRange  = sheet.getRange(groupStartRow, anCol, groupNumRows, 3);
  const allMerges  = fullRange.getMergedRanges();
  for (let i = 0; i < allMerges.length; i++) {
    const m = allMerges[i];
    sheet.getRange(m.getRow(), m.getColumn(), m.getNumRows(), m.getNumColumns()).breakApart();
  }
  try { fullRange.breakApart(); } catch(e) {}

  // AN-AO = horizontal merge per row + per row duration value
  let totalMs = 0, hasAny = false, hasRed = false, allGreen = true;
  for (let r = 0; r < groupNumRows; r++) {
    const curRow     = groupStartRow + r;
    const isParallel = sheet.getRange(curRow, actTypeCol).getValue() === parallelVal;
    const display    = computeRowDisplay(sheet, curRow);
    const color      = isParallel ? parColor : mainColor;

    // AN-AO horizontal merge per row
    const anaoRange = sheet.getRange(curRow, anCol, 1, 2);
    anaoRange.merge();
    writeCellWithStyle(anaoRange, display.text, color);
    anaoRange.setBorder(true, true, true, true, false, false, WHITE, SOLID);

    if (display.category === "RED")   hasRed = true;
    if (display.category !== "GREEN") allGreen = false;
    if (display.durMs !== null && !isParallel) { totalMs += display.durMs; hasAny = true; }
  }

  // AP = vertical merge across full group range
  const apRange = sheet.getRange(groupStartRow, apCol, groupNumRows, 1);
  apRange.merge();
  const text = hasAny ? formatDurationMs(totalMs) : hasRed ? "🔴" : allGreen ? "🟢" : "🟡";
  writeCellWithStyle(apRange, text, mainColor);
  apRange.setBorder(true, true, true, true, true, true, WHITE, SOLID);
}

// ── Called when H creates sub-group → break AN-AO horizontal, AO vertical, AP untouched
function updateSubGroupDuration(sheet, subStartRow, subNumRows) {
  const anCol    = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.PER_ROW_DURATION_COL);
  const aoCol    = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.SUB_GROUP_TOTAL_DURATION_COL);
  const subColor = DAILY_ACTIVITY_LOG_DURATION_COLS.SUB_GROUP_FONT_COLOR;
  const perColor = DAILY_ACTIVITY_LOG_DURATION_COLS.SUB_ROW_FONT_COLOR;
  const parColor = DAILY_ACTIVITY_LOG_DURATION_COLS.PARALLEL_FONT_COLOR;
  const WHITE    = "#ffffff";
  const SOLID    = SpreadsheetApp.BorderStyle.SOLID;
  const actTypeCol  = columnLetterToIndex(DAILY_ACTIVITY_LOG_DURATION_COLS.ACTIVITY_TYPE_COL);
  const parallelVal = DAILY_ACTIVITY_LOG_DURATION_COLS.PARALLEL_VALUE;

  // break AN-AO horizontal merges for sub-group range (AP untouched!)
  const anaoRange  = sheet.getRange(subStartRow, anCol, subNumRows, 2);
  const anaoMerges = anaoRange.getMergedRanges();
  for (let i = 0; i < anaoMerges.length; i++) {
    const m = anaoMerges[i];
    sheet.getRange(m.getRow(), m.getColumn(), m.getNumRows(), m.getNumColumns()).breakApart();
  }
  try { anaoRange.breakApart(); } catch(e) {}

  // AN = individual per row
  let totalMs = 0, hasAny = false, hasRed = false, allGreen = true;
  for (let r = 0; r < subNumRows; r++) {
    const curRow     = subStartRow + r;
    const isParallel = sheet.getRange(curRow, actTypeCol).getValue() === parallelVal;
    const display    = computeRowDisplay(sheet, curRow);
    const color      = isParallel ? parColor : perColor;

    // AN individual
    const anCell = sheet.getRange(curRow, anCol);
    writeCellWithStyle(anCell, display.text, color);
    anCell.setBorder(true, true, true, true, true, true, WHITE, SOLID);

    if (display.category === "RED")   hasRed = true;
    if (display.category !== "GREEN") allGreen = false;
    if (display.durMs !== null && !isParallel) { totalMs += display.durMs; hasAny = true; }
  }

  // AO = vertical merge across sub-group range
  const aoRange = sheet.getRange(subStartRow, aoCol, subNumRows, 1);
  const aoMerges = aoRange.getMergedRanges();
  for (let i = 0; i < aoMerges.length; i++) {
    const m = aoMerges[i];
    sheet.getRange(m.getRow(), m.getColumn(), m.getNumRows(), m.getNumColumns()).breakApart();
  }
  try { aoRange.breakApart(); } catch(e) {}
  aoRange.merge();
  const text = hasAny ? formatDurationMs(totalMs) : hasRed ? "🔴" : allGreen ? "🟢" : "🟡";
  writeCellWithStyle(aoRange, text, subColor);
  aoRange.setBorder(true, true, true, true, true, true, WHITE, SOLID);
  // AP = untouched ✅
}

function refreshDurationValuesForSheet(sheet) {
  const dataStartRow = 3;
  const lastRow      = sheet.getLastRow();
  if (lastRow < dataStartRow) return;
  const insertCol    = columnLetterToIndex(DAILY_ACTIVITY_LOG_COLS.INSERT_ROW_ABOVE);
  const vals         = sheet.getRange(dataStartRow, insertCol, lastRow - dataStartRow + 1, 1).getDataValidations();
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i][0];
    if (v && v.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      updateRowDuration(sheet, dataStartRow + i);
    }
  }
}
