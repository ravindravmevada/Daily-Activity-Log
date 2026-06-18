const SHEET_NAMES = {
  DAILY_ACTIVITY_LOG: [
    "20-05-2026",
    "21-05-2026",
    "22-05-2026",
    "23-05-2026",
    "24-05-2026",
    "25-05-2026",
    "26-05-2026",
    "27-05-2026",
    "28-05-2026",
    "29-05-2026",
    "30-05-2026",
    "31-05-2026",
    "01-06-2026",
    "02-06-2026",
    "03-06-2026",
    "04-06-2026",
    "05-06-2026",
    "06-06-2026",
    "07-06-2026",
    "08-06-2026",
    "09-06-2026",
    "10-06-2026",
    "11-06-2026",
    "12-06-2026",
    "13-06-2026",
    "14-06-2026",
    "15-06-2026",
    "16-06-2026",
    "17-06-2026",
    "18-06-2026",
    "Test",
    "Template"
  ],
  LINKED_DROPDOWN_SOURCE_SHEET: "Dropdown Source - Linked Data"
};

const LINKED_DROPDOWN_SOURCE_SHEET_CONFIG = {
  HEADER_ROW: 3,
  DATA_START_ROW: 4
};

const DAILY_ACTIVITY_LOG_COLS = {
  INSERT_ROW_ABOVE: "A",
  ROW_HIGHLIGHT_COLS: ["C", "H"],
  INCREMENT_NUMBER: "D",
  INSERT_MERGED_GROUP_ROW_ABOVE: "G",
  MERGED_GROUP_LABEL_COL: "E",
  MERGED_GROUP_INCREMENT_NUMBER: "I",
  MERGED_GROUP_SHARED_MERGE_COLS: ["A", "B", "C", "D"],
  CREATE_GROUP_TRIGGER_COL: "B",
  GROUP_VERTICAL_MERGE_COLS: ["A", "B", "C", "D"],
  GROUP_LABEL_MERGE_START_COL: "E",
  GROUP_LABEL_MERGE_END_COL: "F",
  GROUP_INDIVIDUAL_TICKBOX_COLS: ["G", "H"],
  GROUP_SUB_INCREMENT_COL: "I",
  GROUP_ROW_DROPDOWN_ORIGINAL_START_COL: "G",
  GROUP_ROW_DROPDOWN_NARROWED_START_COL: "J",
  GROUP_ROW_DROPDOWN_END_COL: "L",
  UNMERGED_ROW_DROPDOWN_START_COL: "G",
  UNMERGED_ROW_DROPDOWN_END_COL: "L"
};

const DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS = [
  {
    label: "Activity Start Timestamp",
    hasTimestampCol: "S",   // was Q
    timestampCol: "U",      // was S
    timestampStatusCol: "Y" // was W
  },
  {
    label: "Activity End Timestamp",
    hasTimestampCol: "Z",    // was X
    timestampCol: "AB",      // was Z
    timestampStatusCol: "AF" // was AD
  }
];

const DAILY_ACTIVITY_LOG_ACTIVITY_NOTES = {
  HAS_NOTES_COL: "AK",  // was AI
  NOTES_COL: "AM",      // was AK
};

const DAILY_ACTIVITY_LOG_DURATION_COLS = {
  PER_ROW_DURATION_COL: "AI",      // was AG
  GROUP_TOTAL_DURATION_COL: "AJ",  // was AH
  PENDING_TEXT: "Pending",
  SUB_ROW_FONT_COLOR: "#a3e635",        // cyan – per-row AE "#00ffff",
  MERGED_RESULT_FONT_COLOR: "#faab17",   // amber – merged totals
  ACTIVITY_TYPE_COL: "AG",
  PARALLEL_VALUE: "Parallel",
  PARALLEL_FONT_COLOR: "#00ffff"
};

const HIGHLIGHT_BG_COLOR = "#666666";

const DAILY_ACTIVITY_LOG_ACTIVITY_ID = {
  ID_COL: "BB",
  PREFIX: "ACT"
};

const ACTIVITY_NOTES_AI_FIX_AND_RESTORE = {
  TICKBOX_COL: "AQ",
  UNDO_TICKBOX_COL: "AR",
  NOTES_COL: "AM",
  GEMINI_MODEL: "gemini-2.5-flash",
  API_KEY_PROPERTY: "GEMINI_API_KEY",
  BACKUP_PREFIX: "ORIGINAL: ",
  PROMPT: "Rewrite the following activity note in clear, correct English. Translate any non-English content (Hindi, Gujarati, Hinglish, Gujlish, or any language mix) fully into English. Write exactly like a quick personal diary or daily log note - casual, natural, human. Keep the original feeling and context intact. No formal language, no corporate tone, no over-explanation. Just a simple honest note the way a real person would write it for themselves. No first-person pronouns (no 'I', 'my', 'me'). Keep the same meaning and facts — do not add or remove information. Do not end with a period. Return ONLY the final rewritten note, with no quotes, labels, or explanation."
};


const DEVICE_VIA_CASCADE = {
  LOGGED_FROM_COL: "AT",
  LOGGED_VIA_COL: "AX",
  SOURCE_SHEET: "Dropdown Source - Linked Data",
  HEADER_ROW: 3,
  DATA_START_ROW: 4,
  DEVICE_HEADER: "Activity Logged From",
  VIA_HEADER: "Activity Logged Via"
};
