/**
 * ════════════════════════════════════════════════════════════
 *  Gantt Color → Date Sync  |  גאנט רובוטיקה שדרות FLL 2026-27
 * ════════════════════════════════════════════════════════════
 *
 *  WEEK STRUCTURE (after restructure):
 *  Each week = 7 columns (Sun Mon Tue Wed Thu | Fri Sat hidden).
 *  Week 1: Dec 28, 2025 (Sun) – Jan 1, 2026 (Thu) — includes pre-Jan 2026 days.
 *  Week 58: Jan 31, 2027 (Sun) – Feb 4, 2027 (Thu).
 *  Fri + Sat columns EXIST but are HIDDEN.
 *
 *  HEADER ROWS layout (columns I onward):
 *    Row 5  DATE_ROW    — Month anchor dates, merged per month,
 *                         displayed as "MMMM yyyy" (ינואר 2026 …)
 *    Row 6  WEEK_ROW    — שבוע 1 / שבוע 2 … merged per visible week (5 cols)
 *    Row 7  DAY_ROW     — א ב ג ד ה (ו ש in hidden cols)
 *    Row 8  DAY_NUM_ROW — 1  2  3  4  5  6  7 … (day of month)
 *    Row 9  FIRST_DATA  — first task row
 *
 *  SETUP (run once per year):
 *    1. Paste into Extensions → Apps Script, Save
 *    2. Open spreadsheet → menu "🔄 סנכרון גאנט"
 *    3. Click "🗓 אתחל שנה עבודה" → enter year → grid is built
 *    4. Run "התקן טריגר אוטומטי"
 *    5. Done — painting cells now auto-updates D / E / F
 *
 *  v2 FIXES:
 *    - Year prompt always shown (both menu items)
 *    - Exact columns for year only (no 58-week padding)
 *    - Fri+Sat: show all first, hide exactly 2 per week
 *    - Israeli public holidays highlighted gray (#c0c0c0)
 * ════════════════════════════════════════════════════════════
 */

// ─── CONFIGURATION ───────────────────────────────────────────
const CFG = {
  SHEET_NAME:     'גאנט עבודה FLL שנתי',

  DATE_ROW:       5,   // Month labels (merged) + date anchors
  WEEK_ROW:       6,   // Week labels  שבוע X (merged per week)
  DAY_ROW:        7,   // Day names    א ב ג ד ה (ו ש hidden)
  DAY_NUM_ROW:    8,   // Day numbers  1  2  3 … 28 29 30 31
  FIRST_DATA_ROW: 9,   // First task row

  FIRST_COL:      9,   // Column I — first timeline column (1-based)

  COL_START:      4,   // Column D — תאריך התחלה
  COL_END:        5,   // Column E — תאריך סיום
  COL_DURATION:   6,   // Column F — משך (working-day count)
  COL_TASK_NUM:   1,   // Column A — section-header detection

  EMPTY_BG: new Set([
    null, '', '#ffffff', '#ffffffff',
    '#f3f3f3', '#efefef', '#f1f3f4', '#ffffff00',
  ]),
};


// ════════════════════════════════════════════════════════════
//  TRIGGER
// ════════════════════════════════════════════════════════════
function onChange(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(CFG.SHEET_NAME);
    if (!sheet) return;
    syncAllTaskRows_(sheet);
  } catch (err) {
    console.error('GanttSync onChange:', err.message);
  }
}


// ════════════════════════════════════════════════════════════
//  MENU
// ════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔄 סנכרון גאנט')
    .addItem('סנכרן עכשיו',                        'manualSync')
    .addItem('📅 טווח תאריכים של בחירה',          'getSelectedDateRange')
    .addSeparator()
    .addItem('🗓 אתחל שנה עבודה',                 'setupYear')
    .addItem('בנה מחדש שבועות (חד פעמי)',          'restructureToFullWeeks')
    .addItem('אמת מבנה שבועות',                    'validateWeeks')
    .addSeparator()
    .addItem('התקן טריגר אוטומטי',                 'installTrigger')
    .addItem('הסר טריגר אוטומטי',                  'removeTrigger')
    .addToUi();
}


// ════════════════════════════════════════════════════════════
//  MANUAL SYNC
// ════════════════════════════════════════════════════════════
function manualSync() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CFG.SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ גיליון לא נמצא:\n' + CFG.SHEET_NAME);
    return;
  }
  const count = syncAllTaskRows_(sheet);
  SpreadsheetApp.getUi().alert('✅ סנכרון הושלם\n' + count + ' שורות עודכנו');
}


// ════════════════════════════════════════════════════════════
//  GET SELECTED DATE RANGE
// ════════════════════════════════════════════════════════════
function getSelectedDateRange() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== CFG.SHEET_NAME) {
    SpreadsheetApp.getUi().alert('❌ אנא עבור לגיליון:\n' + CFG.SHEET_NAME);
    return;
  }
  const range = sheet.getActiveRange();
  if (!range) { SpreadsheetApp.getUi().alert('❌ אין בחירה פעילה'); return; }

  const selStartCol   = range.getColumn();
  const selEndCol     = range.getLastColumn();
  const lastCol       = sheet.getLastColumn();
  const timelineStart = Math.max(selStartCol, CFG.FIRST_COL);
  const timelineEnd   = Math.min(selEndCol, lastCol);

  if (timelineStart > timelineEnd) {
    SpreadsheetApp.getUi().alert('❌ הבחירה אינה על עמודות ציר הזמן (עמודה I ואילך)');
    return;
  }

  const totalCols = timelineEnd - CFG.FIRST_COL + 1;
  const row5Raw   = sheet.getRange(CFG.DATE_ROW, CFG.FIRST_COL, 1, totalCols).getValues()[0];
  const allDates  = buildDateMap_(row5Raw, totalCols);

  let startDate = null, endDate = null, dayCount = 0;
  const selOffset = timelineStart - CFG.FIRST_COL;
  const selCols   = timelineEnd - timelineStart + 1;

  for (let i = selOffset; i < selOffset + selCols; i++) {
    const d = allDates[i];
    if (!d) continue;
    const dow = d.getDay();
    if (dow === 5 || dow === 6) continue;   // skip Fri/Sat
    if (!startDate) startDate = d;
    endDate = d;
    dayCount++;
  }

  if (!startDate) { SpreadsheetApp.getUi().alert('❌ לא נמצאו תאריכים בטווח'); return; }

  const tz  = Session.getScriptTimeZone();
  const fmt = d => Utilities.formatDate(d, tz, 'dd/MM/yyyy');
  SpreadsheetApp.getUi().alert(
    '📅 טווח תאריכים נבחר\n\n' +
    'התחלה:     ' + fmt(startDate) + '\n' +
    'סיום:      ' + fmt(endDate)   + '\n' +
    'ימי עבודה: ' + dayCount
  );
}


// ════════════════════════════════════════════════════════════
//  YEAR SETUP — Interactive month/year range initialization
//  v4: Two prompts in MM/YYYY format (from-month / to-month).
//      Grid spans exactly from start-month to end-month.
//      Week numbering begins at the correct ISO week for the
//      start month (not always 1), and resets to 1 at each
//      new year's week-1 boundary.
// ════════════════════════════════════════════════════════════
function setupYear() {
  const ui = SpreadsheetApp.getUi();

  // ── Prompt 1: from MM/YYYY ────────────────────────────────
  const r1 = ui.prompt(
    '🗓 אתחל גאנט — חודש/שנה התחלה',
    'הכנס חודש ושנת התחלה בפורמט MM/YYYY\n(לדוגמה: 09/2026)',
    ui.ButtonSet.OK_CANCEL
  );
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const p1 = parseMonthYear_(r1.getResponseText().trim());
  if (!p1) {
    ui.alert('❌ פורמט לא תקין.\nהכנס בפורמט MM/YYYY — לדוגמה: 09/2026'); return;
  }
  const { month: startMonth, year: startYear } = p1;

  // ── Prompt 2: to MM/YYYY ──────────────────────────────────
  const defaultEnd = String(startMonth).padStart(2, '0') + '/' + startYear;
  const r2 = ui.prompt(
    '🗓 אתחל גאנט — חודש/שנה סיום',
    'הכנס חודש ושנת סיום בפורמט MM/YYYY\n(לדוגמה: ' + defaultEnd + ')',
    ui.ButtonSet.OK_CANCEL
  );
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const p2 = parseMonthYear_(r2.getResponseText().trim());
  if (!p2) {
    ui.alert('❌ פורמט לא תקין.\nהכנס בפורמט MM/YYYY — לדוגמה: ' + defaultEnd); return;
  }
  const { month: endMonth, year: endYear } = p2;

  if (endYear * 12 + endMonth < startYear * 12 + startMonth) {
    ui.alert('❌ חודש הסיום חייב להיות שווה או אחרי חודש ההתחלה'); return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CFG.SHEET_NAME);
  if (!sheet) { ui.alert('❌ גיליון לא נמצא: ' + CFG.SHEET_NAME); return; }

  buildYearGrid_(sheet, startYear, startMonth, endYear, endMonth);

  const fmtMY = (m, y) => String(m).padStart(2, '0') + '/' + y;
  const rangeLabel = (startYear === endYear && startMonth === endMonth)
    ? fmtMY(startMonth, startYear)
    : fmtMY(startMonth, startYear) + ' – ' + fmtMY(endMonth, endYear);
  ui.alert(
    '✅ גאנט ' + rangeLabel + ' אותחל!\n\n' +
    'שורות 5-7 מולאו, שישי+שבת הוסתרו.\n' +
    'מיספור שבועות לפי מיקום בשנה (מתחדש בשנה חדשה).\n' +
    'ניתן להוסיף משימות החל משורה ' + CFG.FIRST_DATA_ROW + '.'
  );
}


// ════════════════════════════════════════════════════════════
//  YEAR GRID BUILDER  (v4 — month/year range, correct initial
//  week number, resets to 1 at each new year's ISO week 1)
// ════════════════════════════════════════════════════════════
function buildYearGrid_(sheet, startYear, startMonth, endYear, endMonth) {
  // ── 1. Compute grid boundaries ────────────────────────────
  //  Grid starts on the Sunday of the week containing the 1st
  //  of startMonth.  Grid ends on the Saturday of the week
  //  containing the last day of endMonth.
  const startMonthFirst = new Date(startYear, startMonth - 1, 1);
  const gridStart = new Date(startMonthFirst);
  gridStart.setDate(startMonthFirst.getDate() - startMonthFirst.getDay()); // back to Sunday

  const endMonthLast = new Date(endYear, endMonth, 0);   // day-0 of month+1 = last day of endMonth
  const gridEnd = new Date(endMonthLast);
  gridEnd.setDate(endMonthLast.getDate() + (6 - endMonthLast.getDay())); // forward to Saturday

  const totalDays  = Math.round((gridEnd - gridStart) / 86400000) + 1;
  const totalWeeks = totalDays / 7;
  const totalCols  = totalDays;

  // ── 2. Compute initial week number for gridStart ──────────
  //  The Thursday of gridStart's week determines which year's
  //  ISO numbering applies (handles late-Dec and early-Jan edges).
  const gridStartThu = new Date(gridStart);
  gridStartThu.setDate(gridStart.getDate() + 4);           // Thu of this week
  const gridStartThuYear = gridStartThu.getFullYear();
  const week1SunOfInitYear = getWeek1Sunday_(gridStartThuYear);
  let weekNum = Math.round((gridStart - week1SunOfInitYear) / (7 * 86400000)) + 1;

  // ── 3. Pre-compute week-1 Sundays for every year in grid ──
  //  When we land on one of these Sundays, weekNum resets to 1.
  const firstGridYear = gridStart.getFullYear();
  const lastGridYear  = gridEnd.getFullYear();
  const week1Sundays  = new Set();
  for (let y = firstGridYear; y <= lastGridYear; y++) {
    week1Sundays.add(dateKey_(getWeek1Sunday_(y)));
  }

  // ── 4. Merge holiday sets for all years in grid ───────────
  const holidays = new Set();
  for (let y = firstGridYear; y <= lastGridYear; y++) {
    getIsraeliHolidays_(y).forEach(h => holidays.add(h));
  }

  // ── 5. Range bounds for "other month" gray coloring ───────
  //  Days before startMonth/startYear or after endMonth/endYear
  //  get the dimmed gray background.
  const rangeStartKey = startYear * 12 + (startMonth - 1);
  const rangeEndKey   = endYear   * 12 + (endMonth   - 1);

  // ── 6. Build column arrays ────────────────────────────────
  const row5vals = new Array(totalCols).fill('');
  const row6vals = new Array(totalCols).fill('');
  const row7vals = new Array(totalCols).fill('');
  const bgColors = new Array(totalCols).fill('#ffffff');

  let cur       = new Date(gridStart);
  let prevMonth = -1;

  for (let i = 0; i < totalCols; i++) {
    const month      = cur.getMonth();
    const dayOfMonth = cur.getDate();
    const dow        = cur.getDay();
    const curYear    = cur.getFullYear();

    // "other range" = outside the requested month/year window
    const curMonthKey    = curYear * 12 + month;
    const isOtherRange   = curMonthKey < rangeStartKey || curMonthKey > rangeEndKey;
    const isHoliday      = holidays.has(dateKey_(cur));

    // Reset week counter at week-1 Sunday of each year.
    // (If gridStart itself is a week-1 Sunday, weekNum is already
    //  initialised correctly and the reset is a harmless no-op.)
    if (dow === 0 && week1Sundays.has(dateKey_(cur))) {
      weekNum = 1;
    }

    // Row 5: month anchor on first column of each new month.
    // Use noon (12:00) to avoid UTC-boundary shift in Israel (UTC+2):
    // midnight Jan 1 local = Dec 31 22:00 UTC → Sheets would show December.
    if (month !== prevMonth) {
      row5vals[i] = new Date(curYear, month, 1, 12, 0, 0);
      prevMonth   = month;
    }

    // Row 6: week label only on Sundays
    if (dow === 0) row6vals[i] = 'שבוע ' + weekNum;

    // Row 7: day-of-month
    row7vals[i] = dayOfMonth;

    // Background
    if (isOtherRange) {
      bgColors[i] = '#d8d8d8';          // gray — outside requested range
    } else if (isHoliday) {
      bgColors[i] = '#c0c0c0';          // darker gray — public holiday
    } else {
      bgColors[i] = getWeekColor_(getWeekPosInMonth_(cur));
    }

    cur.setDate(cur.getDate() + 1);
    if (dow === 6) weekNum++;           // advance weekNum after each Saturday
  }

  // ── 4. Adjust sheet column count ─────────────────────────
  //   Show all existing timeline columns first (clears stale hide state),
  //   then add missing columns or hide surplus ones.
  const existingTimelineCols = Math.max(0, sheet.getLastColumn() - CFG.FIRST_COL + 1);
  if (existingTimelineCols > 0) {
    sheet.showColumns(CFG.FIRST_COL, existingTimelineCols);
  }
  const neededLastCol = CFG.FIRST_COL + totalCols - 1;
  if (sheet.getLastColumn() < neededLastCol) {
    sheet.insertColumnsAfter(sheet.getLastColumn(),
      neededLastCol - sheet.getLastColumn());
  }

  // ── 5. Clear header rows 5-7 (break merges + content + format) ──
  const hdrCols  = Math.max(totalCols, existingTimelineCols);
  const hdrRange = sheet.getRange(CFG.DATE_ROW, CFG.FIRST_COL, 3, hdrCols);
  hdrRange.breakApart();
  hdrRange.clearContent();
  hdrRange.clearFormat();

  // ── 6. Write Row 5 — Month labels (merged spans) ─────────
  const r5 = sheet.getRange(CFG.DATE_ROW, CFG.FIRST_COL, 1, totalCols);
  r5.setValues([row5vals]);
  r5.setNumberFormat('MMMM yyyy');

  // Merge each month's columns
  let spanStart = 0;
  for (let i = 1; i <= totalCols; i++) {
    if (i === totalCols || row5vals[i] !== '') {
      const span    = i - spanStart;
      const absCol  = CFG.FIRST_COL + spanStart;
      if (span > 1) sheet.getRange(CFG.DATE_ROW, absCol, 1, span).merge();
      sheet.getRange(CFG.DATE_ROW, absCol, 1, span)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setFontWeight('bold')
        .setFontSize(10);
      spanStart = i;
    }
  }

  // ── 7. Write Row 6 — Week numbers (merged 5 visible cols) ──
  const r6 = sheet.getRange(CFG.WEEK_ROW, CFG.FIRST_COL, 1, totalCols);
  r6.setValues([row6vals]);

  for (let w = 0; w < totalWeeks; w++) {
    const sunAbsCol = CFG.FIRST_COL + w * 7;
    // Merge only the 5 working-day columns (Sun–Thu); Fri+Sat will be hidden
    sheet.getRange(CFG.WEEK_ROW, sunAbsCol, 1, 5)
      .merge()
      .setHorizontalAlignment('center')
      .setFontWeight('bold')
      .setFontSize(9);
  }

  // ── 8. Write Row 7 — Day-of-month numbers ────────────────
  sheet.getRange(CFG.DAY_ROW, CFG.FIRST_COL, 1, totalCols)
    .setValues([row7vals])
    .setHorizontalAlignment('center')
    .setFontSize(8)
    .setNumberFormat('0');

  // ── 9. Apply background colors to all 3 header rows ──────
  r5.setBackgrounds([bgColors]);
  r6.setBackgrounds([bgColors]);
  sheet.getRange(CFG.DAY_ROW, CFG.FIRST_COL, 1, totalCols).setBackgrounds([bgColors]);

  // ── 10. Hide Fri+Sat exactly: 2 cols per week, every week ──
  //   Sheet already fully shown from step 4.
  for (let w = 0; w < totalWeeks; w++) {
    const friAbsCol = CFG.FIRST_COL + w * 7 + 5;   // index 5 = Friday
    sheet.hideColumns(friAbsCol, 2);                 // Fri + Sat
  }

  SpreadsheetApp.flush();
}


// ════════════════════════════════════════════════════════════
//  HELPER: First Thursday of year  (ISO week-1 anchor)
// ════════════════════════════════════════════════════════════
function findFirstThursdayOfYear_(year) {
  const jan1       = new Date(year, 0, 1);
  const dow        = jan1.getDay();            // 0=Sun … 6=Sat
  const daysToThu  = (4 - dow + 7) % 7;       // 0 if Jan1 is Thu
  return new Date(year, 0, 1 + daysToThu);
}


// ════════════════════════════════════════════════════════════
//  HELPER: Sunday that starts ISO week 1 of a given year
//  (The Thursday of that week is the first Thursday of the year)
// ════════════════════════════════════════════════════════════
function getWeek1Sunday_(year) {
  const thu = findFirstThursdayOfYear_(year);
  const sun = new Date(thu);
  sun.setDate(thu.getDate() - thu.getDay());   // back to Sunday (getDay() = 4 for Thu)
  return sun;
}


// ════════════════════════════════════════════════════════════
//  HELPER: Parse "MM/YYYY" string → { month, year } or null
//  Accepts 1 or 2 digit month (e.g. "9/2026" or "09/2026")
// ════════════════════════════════════════════════════════════
function parseMonthYear_(text) {
  const m = text.match(/^(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  const month = parseInt(m[1], 10);
  const year  = parseInt(m[2], 10);
  if (month < 1 || month > 12)   return null;
  if (year  < 2000 || year > 2100) return null;
  return { month, year };
}


// ════════════════════════════════════════════════════════════
//  HELPER: Week position in month (1–5) based on week's Sunday
// ════════════════════════════════════════════════════════════
function getWeekPosInMonth_(date) {
  // Step back to the Sunday of date's week
  const weekSun  = new Date(date);
  weekSun.setDate(date.getDate() - date.getDay());

  // Step back to the Sunday that anchors this month (on or before 1st)
  const monthStart    = new Date(weekSun.getFullYear(), weekSun.getMonth(), 1);
  const anchorSun     = new Date(monthStart);
  anchorSun.setDate(monthStart.getDate() - monthStart.getDay());

  const pos = Math.round((weekSun - anchorSun) / (7 * 86400000)) + 1;
  return Math.min(Math.max(pos, 1), 5);
}


// ════════════════════════════════════════════════════════════
//  HELPER: Background color by week position in month
// ════════════════════════════════════════════════════════════
function getWeekColor_(pos) {
  return [
    '#dce9f5',   // pos 1 — light blue
    '#ddf2e3',   // pos 2 — light green
    '#fdf5d4',   // pos 3 — light yellow
    '#fde8d4',   // pos 4 — light orange
    '#f5d4ea',   // pos 5 — light pink
  ][pos - 1] || '#ffffff';
}


// ════════════════════════════════════════════════════════════
//  HELPER: Date → 'YYYY-MM-DD' key for Set lookup
// ════════════════════════════════════════════════════════════
function dateKey_(d) {
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return d.getFullYear() + '-' + mm + '-' + dd;
}


// ════════════════════════════════════════════════════════════
//  ISRAELI PUBLIC HOLIDAYS  (gray / disabled columns)
//  Update this list each year.  Format: 'YYYY-MM-DD'
//
//  Includes: Chagim, Yom HaAtzmaut, Yom HaZikaron, Yom HaShoah
//  Hol HaMoed days included (school usually closed).
//  Hanukkah NOT included (school continues).
// ════════════════════════════════════════════════════════════
function getIsraeliHolidays_(year) {
  const ALL = {
    2025: [
      '2025-03-13', '2025-03-14',             // Purim (Thu-Fri)
      '2025-04-12',                            // Erev Pesach (half-day)
      '2025-04-13', '2025-04-14',             // Pesach days 1-2
      '2025-04-15', '2025-04-16', '2025-04-17', '2025-04-18', // Hol HaMoed
      '2025-04-19', '2025-04-20',             // Pesach days 7-8
      '2025-05-01',                            // Yom HaShoah
      '2025-05-07',                            // Yom HaZikaron
      '2025-05-08',                            // Yom HaAtzmaut
      '2025-06-01', '2025-06-02',             // Shavuot
      '2025-08-12',                            // Tisha B'Av
      '2025-09-22', '2025-09-23',             // Rosh HaShanah
      '2025-10-01',                            // Yom Kippur
      '2025-10-06', '2025-10-07', '2025-10-08', '2025-10-09',
      '2025-10-10', '2025-10-11', '2025-10-12',// Sukkot + Hol HaMoed
      '2025-10-13',                            // Shemini Atzeret
      '2025-10-14',                            // Simchat Torah
    ],
    2026: [
      '2026-03-05',                            // Purim (Thu)
      '2026-04-21',                            // Erev Pesach (half-day, Tue)
      '2026-04-22', '2026-04-23',             // Pesach days 1-2 (Wed-Thu)
      '2026-04-24', '2026-04-25', '2026-04-27', '2026-04-28', // Hol HaMoed (Fri hidden, Mon-Tue)
      '2026-04-29',                            // Pesach day 7 (Wed)
      '2026-04-30',                            // Yom HaShoah (Thu)
      '2026-05-06',                            // Yom HaZikaron (Wed)
      '2026-05-07',                            // Yom HaAtzmaut (Thu)
      '2026-06-11', '2026-06-12',             // Shavuot (Thu-Fri)
      '2026-08-04',                            // Tisha B'Av (Tue)
      '2026-09-21', '2026-09-22',             // Rosh HaShanah (Mon-Tue)
      '2026-09-30',                            // Yom Kippur (Wed)
      '2026-10-05', '2026-10-06', '2026-10-07', '2026-10-08',
      '2026-10-09', '2026-10-11',             // Sukkot + Hol HaMoed
      '2026-10-12',                            // Shemini Atzeret (Mon)
      '2026-10-13',                            // Simchat Torah (Tue)
    ],
    2027: [
      '2027-03-02',                            // Ta'anit Esther (Mon, optional)
      '2027-03-03', '2027-03-04',             // Purim (Tue-Wed)
      '2027-04-11',                            // Erev Pesach
      '2027-04-12', '2027-04-13',             // Pesach 1-2
      '2027-04-14', '2027-04-15', '2027-04-18',// Hol HaMoed
      '2027-04-19',                            // Pesach 7
      '2027-04-21',                            // Yom HaShoah
      '2027-04-28',                            // Yom HaZikaron
      '2027-04-29',                            // Yom HaAtzmaut
      '2027-05-31', '2027-06-01',             // Shavuot
      '2027-07-22',                            // Tisha B'Av
      '2027-09-11', '2027-09-12',             // Rosh HaShanah
      '2027-09-20',                            // Yom Kippur
      '2027-09-25', '2027-09-26', '2027-09-27', '2027-09-28',
      '2027-09-29', '2027-10-01',             // Sukkot + Hol HaMoed
      '2027-10-02',                            // Shemini Atzeret
      '2027-10-03',                            // Simchat Torah
    ],
  };
  return new Set(ALL[year] || []);
}


// ════════════════════════════════════════════════════════════
//  RESTRUCTURE — legacy redirect to setupYear
//  (kept so old bookmarks still work; always prompts for year)
// ════════════════════════════════════════════════════════════
function restructureToFullWeeks() {
  setupYear();
}


// ════════════════════════════════════════════════════════════
//  FIX DAY NUMBER FORMAT (utility — row 8 percentage fix)
// ════════════════════════════════════════════════════════════
function fixDayNumFormat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert('Sheet not found'); return; }
  const lastCol = sheet.getLastColumn();
  const numCols = lastCol - CFG.FIRST_COL + 1;
  if (numCols <= 0) return;
  const row8 = sheet.getRange(8, CFG.FIRST_COL, 1, numCols);
  row8.clearFormat();
  row8.setNumberFormat('0');
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Done! Row 8 format fixed.');
}


// ════════════════════════════════════════════════════════════
//  CORE ENGINE
// ════════════════════════════════════════════════════════════
function syncAllTaskRows_(sheet) {
  const lock = LockService.getDocumentLock();
  try { lock.waitLock(15000); }
  catch (e) { console.log('GanttSync: lock timeout'); return 0; }

  try {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < CFG.FIRST_DATA_ROW || lastCol < CFG.FIRST_COL) return 0;

    const numTimeCols = lastCol - CFG.FIRST_COL + 1;
    const numDataRows = lastRow - CFG.FIRST_DATA_ROW + 1;

    const row5Raw = sheet.getRange(CFG.DATE_ROW, CFG.FIRST_COL, 1, numTimeCols).getValues()[0];
    const dates   = buildDateMap_(row5Raw, numTimeCols);
    const allBg   = sheet.getRange(CFG.FIRST_DATA_ROW, CFG.FIRST_COL, numDataRows, numTimeCols).getBackgrounds();
    const colA    = sheet.getRange(CFG.FIRST_DATA_ROW, CFG.COL_TASK_NUM, numDataRows, 1).getValues();

    let updatedCount = 0;
    for (let r = 0; r < numDataRows; r++) {
      const taskNum = colA[r][0];
      if (isSectionHeader_(taskNum)) continue;
      if (taskNum === '' || taskNum === null || taskNum === undefined) continue;

      const { startDate, endDate, duration } = resolveBarDates_(allBg[r], dates);
      const absRow = CFG.FIRST_DATA_ROW + r;

      const dCell = sheet.getRange(absRow, CFG.COL_START);
      const eCell = sheet.getRange(absRow, CFG.COL_END);
      const fCell = sheet.getRange(absRow, CFG.COL_DURATION);

      if (startDate) dCell.setValue(startDate).setNumberFormat('d/m/yy');
      else           dCell.clearContent();
      if (endDate)   eCell.setValue(endDate).setNumberFormat('d/m/yy');
      else           eCell.clearContent();
      fCell.setValue(duration > 0 ? duration : '');
      updatedCount++;
    }

    SpreadsheetApp.flush();
    return updatedCount;
  } finally {
    lock.releaseLock();
  }
}


// ════════════════════════════════════════════════════════════
//  DATE MAP BUILDER — auto-detects 7-day vs 5-day structure
// ════════════════════════════════════════════════════════════
function buildDateMap_(row5Values, numCols) {
  const dates = new Array(numCols).fill(null);

  for (let c = 0; c < numCols; c++) {
    const v = row5Values[c];
    if (v instanceof Date && !isNaN(v.getTime())) dates[c] = new Date(v);
  }

  let anchorDate = null, anchorIdx = -1;
  for (let c = 0; c < numCols; c++) {
    if (dates[c]) { anchorDate = dates[c]; anchorIdx = c; break; }
  }
  if (!anchorDate) { console.warn('GanttSync: no anchor in row 5'); return dates; }

  // Detect structure: find the first anchor pair with >= 7 days between them
  // (skips very short partial-month spans like Dec 28-31 = 4 days).
  // 7-day/week structure: colSpan ≈ daySpan (1 col per calendar day).
  // 5-day/week structure: colSpan ≈ daySpan * 5/7 (1 col per working day).
  let use7Day = true;
  let prevDate = anchorDate, prevIdx = anchorIdx;
  for (let c = anchorIdx + 1; c < numCols; c++) {
    if (!dates[c]) continue;
    const colSpan = c - prevIdx;
    const daySpan = Math.round((dates[c] - prevDate) / 86400000);
    if (daySpan >= 7) {
      // Reliable pair: 7-day/week → colSpan ≈ daySpan; 5-day/week → colSpan < daySpan
      use7Day = (colSpan >= Math.round(daySpan * 0.85));
      break;
    }
    // Too short a span — keep looking
    prevDate = dates[c]; prevIdx = c;
  }

  for (let c = 0; c < numCols; c++) {
    if (dates[c]) continue;
    const off = c - anchorIdx;
    dates[c] = use7Day
      ? addCalendarDays_(anchorDate, off)
      : addWorkingDays_(anchorDate, off);
  }
  return dates;
}


// ════════════════════════════════════════════════════════════
//  BAR DATE RESOLVER — skips Fri/Sat
// ════════════════════════════════════════════════════════════
function resolveBarDates_(rowBgs, dates) {
  let firstIdx = -1, lastIdx = -1, duration = 0;
  for (let c = 0; c < rowBgs.length; c++) {
    let bg = (rowBgs[c] || '').toString().toLowerCase().trim();
    if (bg.length === 9 && bg.startsWith('#')) bg = '#' + bg.slice(1, 7);
    if (CFG.EMPTY_BG.has(bg) || bg === '') continue;
    const d = dates[c];
    if (d && (d.getDay() === 5 || d.getDay() === 6)) continue;
    if (firstIdx === -1) firstIdx = c;
    lastIdx = c;
    duration++;
  }
  if (firstIdx === -1) return { startDate: null, endDate: null, duration: 0 };
  return { startDate: dates[firstIdx] || null, endDate: dates[lastIdx] || null, duration };
}


// ════════════════════════════════════════════════════════════
//  SECTION HEADER DETECTOR
// ════════════════════════════════════════════════════════════
function isSectionHeader_(v) {
  if (v === null || v === undefined || v === '') return false;
  const n = Number(v);
  if (isNaN(n) || n <= 0) return false;
  return Number.isInteger(n);
}


// ════════════════════════════════════════════════════════════
//  WEEK VALIDATOR
// ════════════════════════════════════════════════════════════
function validateWeeks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_NAME);
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  const numCols = lastCol - CFG.FIRST_COL + 1;
  const row5Raw = sheet.getRange(CFG.DATE_ROW, CFG.FIRST_COL, 1, numCols).getValues()[0];
  const row7    = sheet.getRange(CFG.DAY_ROW,  CFG.FIRST_COL, 1, numCols).getValues()[0];
  const dates   = buildDateMap_(row5Raw, numCols);
  const HEB     = ['א','ב','ג','ד','ה','ו','ש'];
  const warns   = [];
  let hidden    = 0;
  for (let c = 0; c < numCols; c++) {
    const d = dates[c]; if (!d) continue;
    const dow = d.getDay();
    if (dow === 5 || dow === 6) { hidden++; continue; }
    const lbl = (row7[c] || '').toString().trim();
    if (lbl && lbl !== HEB[dow])
      warns.push(colLetter_(CFG.FIRST_COL+c) + ' (' + fmtDate_(d) + '): "' + lbl + '" → "' + HEB[dow] + '" ⚠️');
  }
  const working = numCols - hidden;
  SpreadsheetApp.getUi().alert(
    warns.length === 0
      ? '✅ מבנה תקין!\n\nעמודות: ' + numCols + '  |  עבודה: ' + working + '  |  מוסתר: ' + hidden
      : '⚠️ ' + warns.length + ' אזהרות:\n' + warns.slice(0,15).join('\n')
  );
}


// ════════════════════════════════════════════════════════════
//  TRIGGER MANAGEMENT
// ════════════════════════════════════════════════════════════
function installTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onChange')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onChange')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange().create();
  SpreadsheetApp.getUi().alert('✅ טריגר הותקן!\nכל שינוי צבע יעדכן D/E/F אוטומטית.');
}

function removeTrigger() {
  const ts = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'onChange');
  ts.forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getUi().alert('✅ ' + ts.length + ' טריגר(ים) הוסרו.');
}


// ════════════════════════════════════════════════════════════
//  ARITHMETIC HELPERS
// ════════════════════════════════════════════════════════════
function addCalendarDays_(baseDate, n) {
  const d = new Date(baseDate.getTime());
  d.setDate(d.getDate() + n);
  return d;
}

function addWorkingDays_(baseDate, n) {
  const d = new Date(baseDate.getTime());
  if (n === 0) return d;
  const step = n > 0 ? 1 : -1;
  let rem = Math.abs(n);
  while (rem > 0) {
    d.setDate(d.getDate() + step);
    if (d.getDay() !== 5 && d.getDay() !== 6) rem--;
  }
  return d;
}


// ════════════════════════════════════════════════════════════
//  UTILITY HELPERS
// ════════════════════════════════════════════════════════════
function colLetter_(n) {
  let s = '';
  while (n > 0) { s = String.fromCharCode(64 + ((n-1)%26+1)) + s; n = Math.floor((n-1)/26); }
  return s;
}

function fmtDate_(d) {
  if (!d) return '?';
  return d.getDate() + '/' + (d.getMonth()+1) + '/' + d.getFullYear();
}
