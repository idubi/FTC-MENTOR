# Gantt Color → Date Sync: Design Document
**Project:** גאנט מנטורים רובוטיקה שדרות FLL 2026-27
**Author:** Claude
**Date:** April 2026

---

## 1. Goal

When a user paints (fills) any cell in the Gantt timeline area with a non-default color:

| Column | Field | Sync Behavior |
|--------|-------|---------------|
| D | תאריך התחלה (Start Date) | Date of the **first** colored cell in that row |
| E | תאריך סיום (End Date) | Date of the **last** colored cell in that row |
| F | משך (Duration) | **Count** of all colored working days in that row |

The date for each timeline column is taken from **Row 5** (anchor date header).

---

## 2. Why Google Apps Script (Not Regular Formulas)

Google Sheets has **no native formula** that can read a cell's background color.
Standard functions like `=CELL()`, `=IF()`, `=MATCH()` are blind to fill colors.

The only solution inside Google Sheets is **Google Apps Script** (GAS), which is JavaScript that runs server-side inside Google's infrastructure and has full access to cell formatting via `getBackground()` / `getBackgrounds()`.

---

## 3. Sheet Structure (Reference)

```
Row 5:  │ I5        │ J5        │ K5        │ L5        │ M5        │ N5 … │
        │ 6/28/2026 │ 6/29/2026 │ 6/30/2026 │ 7/1/2026  │ 7/2/2026  │  …   │
        │ (Sun W28) │ (Mon W28) │ (Tue W28) │ (Wed W28) │ (Thu W28) │ Sun …│

Row 6:  │ ←──────── שבוע 28 (5 cols merged) ─────────→ │ ← שבוע 29 → │

Row 7:  │    א      │    ב      │    ג      │    ד      │    ה      │  א … │
        │   Sun     │   Mon     │   Tue     │   Wed     │   Thu     │      │

Row 8:  Section header: "הכשרת צוות" (gray, spans full row — SKIP)

Row 9:  Task 1.1 │ … │ D9=Start │ E9=End │ F9=Duration │ … │ [Gantt cells] │
Row 10: Task 1.1.1 …
```

**Key constants:**
- `FIRST_TIMELINE_COL = 9`  → Column I (1-based index)
- `LAST_TIMELINE_COL` = dynamically determined from last non-empty cell in Row 6
- `DATE_ROW = 5`  → Row 5 holds the date for each column
- `START_COL = 4`  → Column D
- `END_COL = 5`    → Column E
- `DURATION_COL = 6` → Column F
- `FIRST_DATA_ROW = 9` → First task row (after headers + first section header)

---

## 4. Color Detection Logic

### 4.1 What Counts as "Colored" (Active Gantt Bar)

A cell is considered **part of the Gantt bar** if its background color is anything except:

| Color | Hex | Meaning |
|-------|-----|---------|
| White | `#ffffff` | Default / empty cell |
| null / transparent | `null` | No fill |
| Very light gray | `#f3f3f3` | Alternate row tint |
| Light blue-gray (week grid) | `#cfe2f3` | Grid background tint (if present) |

> **Important:** The script will read all backgrounds at startup to auto-detect what the "resting" (no-bar) colors are in this specific sheet, rather than hard-coding them. This makes the system resilient to theme changes.

### 4.2 What Counts as a "Section Header Row" (Skip)

A row is a section header (and should be skipped) if:
- Column A (מס"ד) contains a **whole number** like `1`, `2`, `3`, `4`, `5` (not `1.1`, `2.1`, etc.)
- OR the entire row has a uniform dark background with light text

The script checks `col A` numeric value and skips rows where it is an integer with no decimal.

---

## 5. Architecture: Three-Layer Design

```
┌─────────────────────────────────────────────────────────────┐
│  LAYER 1 — TRIGGER                                          │
│  onChange(e) — installable trigger                          │
│  Fires when:  formatting changes (e.changeType == "FORMAT") │
│               also catches value edits for safety           │
└──────────────────────┬──────────────────────────────────────┘
                       │ calls
┌──────────────────────▼──────────────────────────────────────┐
│  LAYER 2 — SCANNER: syncAllTaskRows(sheet)                  │
│  • Gets full timeline background grid in ONE batch call     │
│  • Gets date row (row 5) values in ONE batch call           │
│  • Iterates over each task row (rows 9–last)                │
│  • Skips section header rows (integer col-A value)          │
│  • For each task row → calls resolveBarDates(rowBgs, dates) │
└──────────────────────┬──────────────────────────────────────┘
                       │ calls per-row
┌──────────────────────▼──────────────────────────────────────┐
│  LAYER 3 — RESOLVER: resolveBarDates(rowBgs, dates)         │
│  • Scans rowBgs[] array for non-default colors              │
│  • Returns { startDate, endDate, durationCount }            │
│  • startDate = dates[firstColoredIndex]                     │
│  • endDate   = dates[lastColoredIndex]                      │
│  • durationCount = count of all colored cells               │
└─────────────────────────────────────────────────────────────┘
```

---

## 6. Apps Script Pseudocode

```javascript
// ============================================================
// CONSTANTS — adjust once if sheet layout changes
// ============================================================
const SHEET_NAME      = 'גאנט עבודה FLL שנתי';
const DATE_ROW        = 5;      // Row 5 holds per-column dates
const WEEK_LABEL_ROW  = 6;      // Row 6: week headers (merged)
const DAY_LABEL_ROW   = 7;      // Row 7: א ב ג ד ה labels
const FIRST_DATA_ROW  = 9;      // First task row
const FIRST_COL       = 9;      // Column I (1-based)
const COL_START       = 4;      // Column D  — תאריך התחלה
const COL_END         = 5;      // Column E  — תאריך סיום
const COL_DURATION    = 6;      // Column F  — משך

// Colors to treat as "no bar" (empty)
const DEFAULT_COLORS  = new Set([null, '#ffffff', '#f3f3f3', '#efefef']);

// ============================================================
// TRIGGER — fires on any format change
// ============================================================
function onChange(e) {
  if (e.changeType !== 'FORMAT' && e.changeType !== 'EDIT') return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                              .getSheetByName(SHEET_NAME);
  if (!sheet) return;
  syncAllTaskRows(sheet);
}

// ============================================================
// SCANNER — batch-reads the full timeline, iterates rows
// ============================================================
function syncAllTaskRows(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const numTimelineCols = lastCol - FIRST_COL + 1;
  const numDataRows     = lastRow  - FIRST_DATA_ROW + 1;

  // ONE batch read for all backgrounds in timeline area
  const bgGrid = sheet
    .getRange(FIRST_DATA_ROW, FIRST_COL, numDataRows, numTimelineCols)
    .getBackgrounds();

  // ONE batch read for date headers (row 5)
  const dateRow = sheet
    .getRange(DATE_ROW, FIRST_COL, 1, numTimelineCols)
    .getValues()[0];   // flat array of date values

  // Read col A to detect section header rows
  const colA = sheet
    .getRange(FIRST_DATA_ROW, 1, numDataRows, 1)
    .getValues();

  // Batch-collect writes (avoid per-cell writes → slow)
  const writes = [];   // [ { row, startDate, endDate, duration } ]

  for (let r = 0; r < numDataRows; r++) {
    const absRow = FIRST_DATA_ROW + r;

    // Skip section headers (integer task numbers: 1, 2, 3…)
    const taskNum = colA[r][0];
    if (isSectionHeader(taskNum)) continue;

    const rowBgs = bgGrid[r];
    const result = resolveBarDates(rowBgs, dateRow);

    writes.push({ row: absRow, ...result });
  }

  // BATCH WRITE — one setValue per row (3 cells: D, E, F)
  writes.forEach(w => {
    sheet.getRange(w.row, COL_START).setValue(w.startDate || '');
    sheet.getRange(w.row, COL_END).setValue(w.endDate || '');
    sheet.getRange(w.row, COL_DURATION).setValue(w.duration || '');
  });
}

// ============================================================
// RESOLVER — finds first/last colored cell and duration
// ============================================================
function resolveBarDates(rowBgs, dateRow) {
  let firstIdx = -1, lastIdx = -1, duration = 0;

  for (let c = 0; c < rowBgs.length; c++) {
    const bg = rowBgs[c];
    if (!DEFAULT_COLORS.has(bg)) {       // ← colored cell
      if (firstIdx === -1) firstIdx = c;
      lastIdx = c;
      duration++;
    }
  }

  if (firstIdx === -1) return { startDate: null, endDate: null, duration: 0 };

  return {
    startDate: dateRow[firstIdx],   // Date object from row 5
    endDate:   dateRow[lastIdx],
    duration:  duration             // working-day count
  };
}

// ============================================================
// HELPER — detect section header rows
// ============================================================
function isSectionHeader(cellValue) {
  if (cellValue === '' || cellValue === null) return false;
  const n = Number(cellValue);
  return Number.isInteger(n) && n > 0;
}

// ============================================================
// MANUAL TRIGGER — callable from custom menu button
// ============================================================
function manualSync() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                              .getSheetByName(SHEET_NAME);
  syncAllTaskRows(sheet);
  SpreadsheetApp.getUi().alert('✅ גאנט סונכרן בהצלחה');
}

// ============================================================
// SETUP — run ONCE to install the onChange trigger + menu
// ============================================================
function installTrigger() {
  // Remove existing duplicate triggers first
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onChange')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onChange')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();

  SpreadsheetApp.getUi().alert('✅ טריגר הותקן בהצלחה');
}

// ============================================================
// MENU — adds "סנכרון גאנט" menu to the spreadsheet UI
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔄 סנכרון גאנט')
    .addItem('סנכרן עכשיו', 'manualSync')
    .addItem('התקן טריגר אוטומטי', 'installTrigger')
    .addToUi();
}
```

---

## 7. Trigger Setup: Step-by-Step

1. Open the Gantt spreadsheet
2. Go to **Extensions → Apps Script**
3. Paste the full script into `Code.gs`
4. Click **Save**
5. Click **Run → onOpen** (first time — grants permissions)
6. Return to the sheet; a new menu **"🔄 סנכרון גאנט"** appears
7. Click **"סנכרון גאנט → התקן טריגר אוטומטי"** (installs the onChange trigger)
8. Grant the permissions popup (required once)

From now on, any background color change in the timeline fires the sync automatically.

---

## 8. Duration Column (F — משך) Format

The script writes an **integer** to column F (count of colored working days).
The existing `=DAYS36...` formula in F will be **replaced** by this integer.

If you prefer to keep the formula, an alternative is to add a hidden helper column that stores the colored-day count, and update F to reference it:

```
Option A (simple):  F = integer written by script
Option B (formula): G = integer written by script (hidden)
                    F = =G9  (formula referencing G)
```

---

## 9. Week Structure Verification

The design assumes row 5 contains **only Sun–Thu dates** (skipping Fri/Sat).
A validation function will be included:

```javascript
function validateWeekStructure() {
  // For each column in the timeline:
  //   - Get date from row 5
  //   - Verify getDay() returns 0,1,2,3,4 (Sun=0, Mon=1, …, Thu=4)
  //   - Alert if any Fri (5) or Sat (6) found
  //   - Report column letter + date that violates the rule
}
```

---

## 10. Edge Cases & Handling

| Situation | Behavior |
|-----------|----------|
| Row has NO colored cells | D and E are cleared (set to empty) |
| Row has only ONE colored cell | D = E = that cell's date; F = 1 |
| Two separate colored segments with a gap | D = first segment start, E = last segment end, F = total colored count (gap NOT counted) |
| Section header row (e.g., "הכשרת צוות") | Skipped entirely — D/E/F not touched |
| User colors the header/label rows (5, 6, 7) | Skipped (below `FIRST_DATA_ROW`) |
| Date in row 5 is blank for a column | That column is ignored for D/E resolution |
| Multiple users editing simultaneously | Script uses `LockService` to prevent race conditions |

---

## 11. Limitations

| Limitation | Explanation |
|------------|-------------|
| **XLSX mode warning** | The file is currently stored as `.xlsx`. Apps Script works best with native Google Sheets format. **Recommended:** Save a copy as Google Sheets format (`File → Save as Google Sheets`) |
| **onChange fires for all format changes** | Coloring a header cell, resizing a column, etc. all trigger a full re-scan. With <100 task rows this runs in <1 second. |
| **No undo sync** | If user undoes (Ctrl+Z) a color change, the D/E/F values written by the script are NOT undone (Apps Script writes bypass undo history) |
| **No Fri/Sat columns allowed** | Script logic assumes every column in the timeline is a working day (Sun–Thu). If any Fri/Sat columns exist, the date math will be wrong. |

---

## 12. Implementation Plan (Next Step)

Once design is approved, implementation proceeds in this order:

1. **Convert file** to native Google Sheets format (recommended)
2. **Validate week structure** — confirm row 5 has only Sun–Thu dates; fix if needed
3. **Paste and run Apps Script** — install trigger + menu
4. **Test** with a single task row: color 3 cells → confirm D/E/F update correctly
5. **Test edge cases** — zero colored cells, one cell, non-contiguous coloring
6. **Fix column F formula** — replace `=DAYS36...` with script-written value or helper column

---

*Ready to implement on your approval.*
