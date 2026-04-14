# Cache: GanttColorSync Apps Script Reference
**Last updated:** 2026-04-13

---

## File Location
- Workspace: `FTC-MENTOR-GANT-DIRECTOR/GanttColorSync.gs`
- Installed in: Google Sheets → Extensions → Apps Script

## Purpose
Reads cell background colors in the Gantt timeline area and auto-updates columns D (start), E (end), F (duration) for each task row.

## CFG Object (Configuration Constants)
```javascript
const CFG = {
  SHEET_NAME:     'גאנט עבודה FLL שנתי',
  DATE_ROW:       5,   // Month labels (merged) + date anchors
  WEEK_ROW:       6,   // Week labels  שבוע X
  DAY_ROW:        7,   // Day names    א ב ג ד ה
  DAY_NUM_ROW:    8,   // Day numbers  1 2 3 … 31
  FIRST_DATA_ROW: 9,   // First task row
  FIRST_COL:      9,   // Column I — first timeline column
  COL_START:      4,   // Column D — תאריך התחלה
  COL_END:        5,   // Column E — תאריך סיום
  COL_DURATION:   6,   // Column F — משך
  COL_TASK_NUM:   1,   // Column A — section-header detection
  EMPTY_BG: Set([null, '', '#ffffff', '#ffffffff', '#f3f3f3', '#efefef', '#f1f3f4', '#ffffff00'])
};
```

## Key Functions

| Function | Purpose | Trigger |
|----------|---------|---------|
| `onChange(e)` | Auto-sync on any sheet change | Installable onChange trigger |
| `onOpen()` | Creates "🔄 סנכרון גאנט" menu | Sheet open |
| `manualSync()` | Full sync, shows count alert | Menu button |
| `getSelectedDateRange()` | Shows date range of selected cells | Menu button |
| `setupYear()` | **NEW** Interactive year setup with calendar grid | Menu button |
| `buildYearGrid_()` | **NEW** Builds full-year calendar (1/1-31/12) with colors | Called by setupYear |
| `findFirstThursdayOfYear_()` | **NEW** ISO week 1 calculation | Helper |
| `getWeekPositionInMonth_()` | **NEW** Week position for color coding | Helper |
| `getWeekColor_()` | **NEW** Color map by week position | Helper |
| `restructureToFullWeeks()` | One-time: rebuilds week structure from Dec 28, 2025 | Menu (run once) |
| `validateWeeks()` | Checks day labels match actual dates | Menu button |
| `installTrigger()` | Sets up automatic onChange trigger | Menu button |
| `removeTrigger()` | Removes auto trigger | Menu button |
| `fixDayNumFormat()` | Fixes row 8 percentage display bug | Utility |

## Core Engine: syncAllTaskRows_(sheet)
1. Acquires document lock (15s timeout)
2. Batch-reads: row 5 dates, all backgrounds, column A values
3. Builds date map via `buildDateMap_()`
4. For each non-header task row: `resolveBarDates_()` → writes D/E/F
5. Releases lock

## buildDateMap_(row5Values, numCols)
- Finds anchor dates from row 5 merged cells
- Auto-detects 7-day vs 5-day column structure
- Fills in all column dates by interpolation
- Returns: Date[] array (one per timeline column)

## resolveBarDates_(rowBgs, dates)
- Scans background colors left-to-right
- Skips EMPTY_BG colors + Fri/Sat dates
- Returns: `{ startDate, endDate, duration }`

## restructureToFullWeeks()
- Breaks all merges in header rows
- Inserts Dec 28–31 columns if needed (detects by checking if row 7 col I = 'ה')
- Ensures minimum 406 columns (58 weeks × 7)
- Rebuilds rows 5–8 (month labels, week labels, day names, day numbers)
- Sets Fri+Sat columns as hidden
- Idempotent — safe to run multiple times

## Installation Guide
Full guide in: `FTC-MENTOR-GANT-DIRECTOR/GanttColorSync_Installation.md`
Owner account: office1.aido@gmail.com
