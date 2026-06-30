const SHEET_REGISTRY = {
  DAILY_ACTIVITY_LOG_SHEET_NAME_PATTERN:    /^\d{2}-\d{2}-\d{4}$/,
  DAILY_ACTIVITY_LOG_EXTRA_SHEET_NAMES:     ["Template", "Test"],
  CENTRAL_CATEGORY_SOURCE_SHEET_NAME:       "Dropdown Source - Category",
  CENTRAL_ACTIVITY_SOURCE_SHEET_NAME:       "Dropdown Source - Activity",
  CENTRAL_ACTIVITY_PHASE_SOURCE_SHEET_NAME: "Dropdown Source - Activity Phase",
  LINKED_DROPDOWN_SOURCE_SHEET:             "Dropdown Source - Category"
};

const CENTRAL_SOURCE_CONFIG = {
  HEADER_ROW:     3,
  DATA_START_ROW: 4
};

const LINKED_DROPDOWN_SOURCE_SHEET_CONFIG = {
  HEADER_ROW:     3,
  DATA_START_ROW: 4
};

const DAILY_ACTIVITY_LOG_COLS = {
  INSERT_ROW_ABOVE:                      "A",
  CREATE_GROUP_TRIGGER_COL:              "B",
  ROW_HIGHLIGHT_COLS:                    ["C", "I", "N"],
  INCREMENT_NUMBER:                      "D",
  GROUP_LABEL_MERGE_START_COL:           "E",
  GROUP_LABEL_MERGE_END_COL:             "F",
  GROUP_VERTICAL_MERGE_COLS:             ["A", "B", "C", "D"],
  GROUP_SHARED_MERGE_COLS:               ["A", "B", "C", "D"],

  INSERT_MERGED_GROUP_ROW_ABOVE:         "G",
  MERGED_GROUP_LABEL_COL:                "E",
  MERGED_GROUP_INCREMENT_NUMBER:         "J",
  CREATE_SUB_GROUP_TRIGGER_COL:          "H",
  SUB_GROUP_ROW_HIGHLIGHT_COL:           "I",
  GROUP_SUB_INCREMENT_COL:               "J",
  GROUP_INDIVIDUAL_TICKBOX_COLS:         ["G", "H", "I"],
  SUB_GROUP_LABEL_MERGE_START_COL:       "K",
  SUB_GROUP_LABEL_MERGE_END_COL:         "L",
  SUB_GROUP_VERTICAL_MERGE_COLS:         ["G", "H", "I", "J"],

  GROUP_ROW_DROPDOWN_ORIGINAL_START_COL: "G",
  GROUP_ROW_DROPDOWN_NARROWED_START_COL: "K",
  GROUP_ROW_DROPDOWN_END_COL:            "Q",
  UNMERGED_ROW_DROPDOWN_START_COL:       "G",   // default blank row: G-Q is the activity merge
  UNMERGED_ROW_DROPDOWN_END_COL:         "Q",

  INSERT_SUB_MERGED_GROUP_ROW_ABOVE:     "M",
  SUB_SUB_ROW_HIGHLIGHT_COL:             "N",
  GROUP_SUB_SUB_INCREMENT_COL:           "O",
  SUB_SUB_GROUP_INDIVIDUAL_TICKBOX_COLS: ["M", "N"],
  ACTIVITY_PHASE_DROPDOWN_START_COL:     "P",
  ACTIVITY_PHASE_DROPDOWN_END_COL:       "Q",
  SUB_SUB_GROUP_VERTICAL_MERGE_COLS:     ["M", "N", "O"],
};

const DAILY_ACTIVITY_LOG_TIMESTAMP_GROUPS = [
  {
    label:              "Activity Start Timestamp",
    hasTimestampCol:    "X",
    timestampCol:       "Z",
    timestampStatusCol: "AD"
  },
  {
    label:              "Activity End Timestamp",
    hasTimestampCol:    "AE",
    timestampCol:       "AG",
    timestampStatusCol: "AK"
  }
];

const DAILY_ACTIVITY_LOG_DURATION_COLS = {
  PER_ROW_DURATION_COL:         "AN",
  SUB_GROUP_TOTAL_DURATION_COL: "AO",
  GROUP_TOTAL_DURATION_COL:     "AP",
  PENDING_TEXT:                 "Pending",
  SUB_ROW_FONT_COLOR:           "#a3e635",
  SUB_GROUP_FONT_COLOR:         "#00ffff",
  MERGED_RESULT_FONT_COLOR:     "#faab17",
  ACTIVITY_TYPE_COL:            "AL",
  PARALLEL_VALUE:               "Parallel",
  PARALLEL_FONT_COLOR:          "#00ffff"
};

const DAILY_ACTIVITY_LOG_ACTIVITY_NOTES = {
  HAS_NOTES_COL: "AQ",
  NOTES_COL:     "AS",
};

const HIGHLIGHT_BG_COLOR = "#666666";

const DAILY_ACTIVITY_LOG_ACTIVITY_ID = {
  ID_COL: "BA",
  PREFIX: "ACT"
};

const ACTIVITY_NOTES_AI_FIX_AND_RESTORE = {
  TICKBOX_COL:      "AW",
  UNDO_TICKBOX_COL: "AX",
  NOTES_COL:        "AS",
  GEMINI_MODEL:     "gemini-2.5-flash",
  API_KEY_PROPERTY: "GEMINI_API_KEY",
  BACKUP_PREFIX:    "ORIGINAL: ",
  PROMPT: "Rewrite the following activity note in clear, correct English. Translate any non-English content (Hindi, Gujarati, Hinglish, Gujlish, or any language mix) fully into English. Write exactly like a quick personal diary or daily log note - casual, natural, human. Keep the original feeling and context intact. No formal language, no corporate tone, no over-explanation. Just a simple honest note the way a real person would write it for themselves. No first-person pronouns (no 'I', 'my', 'me'). Keep the same meaning and facts — do not add or remove information. Do not end with a period. Return ONLY the final rewritten note, with no quotes, labels, or explanation."
};

const DEVICE_VIA_CASCADE = {
  LOGGED_FROM_COL: "AZ",
  LOGGED_VIA_COL:  "BD",
  SOURCE_SHEET:    "Dropdown Source - Category",
  HEADER_ROW:      3,
  DATA_START_ROW:  4,
  DEVICE_HEADER:   "Activity Logged From",
  VIA_HEADER:      "Activity Logged Via"
};
