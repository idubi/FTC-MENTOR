/**
 * ════════════════════════════════════════════════════════════
 *  Gantt Color → Date Sync  |  גאנט רובוטיקה שדרות FLL 2026-27
 * ════════════════════════════════════════════════════════════
 *
 *  WEEK STRUCTURE (after restructure):
 *  Each week = 7 columns (Sun Mon Tue Wed Thu | Fri Sat hidden).
 *  Week 1 starts Jan 1, 2026 (Thu) — first visible day.
 *  Fri + Sat columns EXIST but are HIDDEN.
 *
 *  HEADER ROWS layout (columns I onward):
 *    Row 5  DATE_ROW    — Month anchor dates, merged per month,
 *                         displayed as "MMMM yyyy" (ינואר 2026 …)
 *    Row 6  WEEK_ROW    — שבוע 1 / שבוע 2 … merged per 5-col week
 *    Row 7  DAY_ROW     — א ב ג ד ה (ו ש in hidden cols)
 *    Row 8  DAY_NUM_ROW — 1  2  3  4  5  6  7 … (day of month)
 *    Row 9  FIRST_DATA  — first task row
 *
 *  SETUP (run once per year):
 *    1. Paste into Extensions → Apps Script, Save
 *    2. Open spreadsheet → menu "🔄 סנכרון גאנט"
 *    3. Run "🗓 בנה מחדש שבועות (חד פעמי)" — confirm Yes
 *    4. Run "התקן טריגר אוטומטי"
 *    5. Done — painting cells now auto-updates D / E / F
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
    .addItem('התקן טריגר אוטומטי',                 'installTrigger')
    .addItem('הסר טריגר אוטומטי',                  'removeTrigger')
    .addSeparator()
    .addItem('🗓 בנה מחדש שבועות (חד פעמי)',      'restructureToFullWeeks')
    .addItem('אמת מבנה שבועות',                    'validateWeeks')
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
//  ONE-TIME RESTRUCTURE
//  Converts 5-day/week  →  7-day/week (Fri+Sat hidden).
//  Starts from Jan 1, 2026 (existing col I).
//  Rebuilds header rows 5-8 with full convention:
//    Row 5 = Month names (merged)
//    Row 6 = Week numbers (merged)
//    Row 7 = Day names א–ש
//    Row 8 = Day numbers 1–31
// ════════════════════════════════════════════════════════════
function restructureToFullWeeks() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert('❌ גיליון לא נמצא'); return; }
  const ui = SpreadsheetApp.getUi();

  const resp = ui.alert(
    '🗓 בניית מחדש של מבנה שבועות',
    'הפעולה תוסיף עמודות מוסתרות לכל שישי ושבת\n' +
    'ותבנה מחדש את שורות הכותרת (5–8).\n\n' +
    'הפעלה חד-פעמית בלבד! להמשיך?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  // ── 0. Baseline ─────────────────────────────────────────
  const lastCol     = sheet.getLastColumn();
  const numTimeCols = lastCol - CFG.FIRST_COL + 1;
  if (numTimeCols < 5) {
    ui.alert('❌ ציר הזמן קצר מדי — אולי הפעולה כבר בוצעה?');
    return;
  }
  console.log('restructure: numTimeCols = ' + numTimeCols);

  // The existing structure: 5-day/week starting Jan 1, 2026 (Thu).
  // Thursdays sit at offsets 0, 5, 10, 15 … from FIRST_COL.
  const ganttStart = new Date(2026, 0, 1);   // Jan 1, 2026

  // ── 1. Collect Thursday column numbers (right → left) ───
  const thuCols = [];
  for (let off = 0; off < numTimeCols; off += 5) {
    const d = addWorkingDays_(ganttStart, off);
    if (d.getDay() === 4) thuCols.push(CFG.FIRST_COL + off);
  }
  console.log('restructure: ' + thuCols.length + ' Thursdays');

  // ── 2. Insert 2 hidden cols (Fri+Sat) after each Thu ────
  for (let i = thuCols.length - 1; i >= 0; i--) {
    sheet.insertColumnsAfter(thuCols[i], 2);
    sheet.hideColumns(thuCols[i] + 1, 2);
  }

  // ── 3. Recalculate totals ────────────────────────────────
  const newLastCol     = sheet.getLastColumn();
  const newNumTimeCols = newLastCol - CFG.FIRST_COL + 1;
  console.log('restructure: newNumTimeCols = ' + newNumTimeCols);

  // ── 4. Break apart existing merges in header rows ───────
  const headerRange = sheet.getRange(CFG.DATE_ROW, CFG.FIRST_COL, 4, newNumTimeCols);
  headerRange.breakApart();

  // ── 5. Build per-column date array (Jan 1 + n days) ─────
  const colDates = [];
  for (let c = 0; c < newNumTimeCols; c++) {
    colDates.push(addCalendarDays_(ganttStart, c));
  }

  // ── 6. ROW 5 — Month labels (merged, Date value) ─────────
  const MONTH_HEB = [
    'ינואר','פברואר','מרץ','אפריל','מאי','יוני',
    'יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'
  ];
  sheet.getRange(CFG.DATE_ROW, CFG.FIRST_COL, 1, newNumTimeCols).clearContent();

  let mStart = 0;
  let mMonth = colDates[0].getMonth();
  let mYear  = colDates[0].getFullYear();
  for (let c = 0; c <= newNumTimeCols; c++) {
    const isEnd = (c === newNumTimeCols) ||
                  (colDates[c].getMonth() !== mMonth || colDates[c].getFullYear() !== mYear);
    if (isEnd) {
      const span = c - mStart;
      const absCol = CFG.FIRST_COL + mStart;
      // Store a real Date so buildDateMap_ picks it up
      const anchorDate = new Date(mYear, mMonth, 1);
      sheet.getRange(CFG.DATE_ROW, absCol).setValue(anchorDate);
      sheet.getRange(CFG.DATE_ROW, absCol).setNumberFormat('MMMM yyyy');
      if (span > 1) sheet.getRange(CFG.DATE_ROW, absCol, 1, span).merge();
      sheet.getRange(CFG.DATE_ROW, absCol, 1, span)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setFontWeight('bold');
      if (c < newNumTimeCols) {
        mStart = c;
        mMonth = colDates[c].getMonth();
        mYear  = colDates[c].getFullYear();
      }
    }
  }

  // ── 7. ROW 6 — Week labels (merged per 5 visible cols) ───
  sheet.getRange(CFG.WEEK_ROW, CFG.FIRST_COL, 1, newNumTimeCols).clearContent();
  // Week 1: cols 0–2 (Jan1 Thu + Fri hidden + Sat hidden) = 3 cols
  // Week 2: cols 3–9 (Jan4 Sun … Jan10 Sat hidden) = 7 cols, etc.
  // We detect week starts by Sunday (getDay()===0) or the very first col
  let weekNum = 1;
  let wStart  = 0;
  for (let c = 0; c <= newNumTimeCols; c++) {
    const isEnd = (c === newNumTimeCols) ||
                  (c > 0 && colDates[c].getDay() === 0);  // Sunday = new week
    if (isEnd) {
      const span    = c - wStart;
      const absCol  = CFG.FIRST_COL + wStart;
      // Visible span = remove trailing hidden cols (Fri+Sat at end of week)
      let visSpan = 0;
      for (let j = wStart; j < c; j++) {
        if (colDates[j].getDay() !== 5 && colDates[j].getDay() !== 6) visSpan++;
        else break; // Fri/Sat are at the END of the week — stop counting
      }
      visSpan = Math.max(1, visSpan);
      sheet.getRange(CFG.WEEK_ROW, absCol).setValue('שבוע ' + weekNum);
      if (visSpan > 1) sheet.getRange(CFG.WEEK_ROW, absCol, 1, visSpan).merge();
      sheet.getRange(CFG.WEEK_ROW, absCol, 1, visSpan)
        .setHorizontalAlignment('center')
        .setFontWeight('bold');
      weekNum++;
      wStart = c;
    }
  }

  // ── 8. ROW 7 — Day names א–ש ────────────────────────────
  const HEB_DAYS = ['א','ב','ג','ד','ה','ו','ש'];
  const dayNames = colDates.map(d => HEB_DAYS[d.getDay()]);
  sheet.getRange(CFG.DAY_ROW, CFG.FIRST_COL, 1, newNumTimeCols)
    .setValues([dayNames])
    .setHorizontalAlignment('center');

  // ── 9. ROW 8 — Day numbers 1–31 ─────────────────────────
  const dayNums = colDates.map(d => d.getDate());
  sheet.getRange(CFG.DAY_NUM_ROW, CFG.FIRST_COL, 1, newNumTimeCols)
    .setValues([dayNums])
    .setHorizontalAlignment('center')
    .setFontSize(8);

  SpreadsheetApp.flush();

  ui.alert(
    '✅ מבנה שבועות עודכן!\n\n' +
    'שורות 5–8 נבנו מחדש:\n' +
    '  שורה 5 — שמות חודשים (ממוזג)\n' +
    '  שורה 6 — שבועות (ממוזג)\n' +
    '  שורה 7 — שמות ימים א–ש\n' +
    '  שורה 8 — מספרי ימים 1–31\n\n' +
    'כעת הרץ "סנכרן עכשיו" לעדכון תאריכי המשימות.'
  );
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

  // Detect 7-day vs 5-day by comparing two adjacent anchors
  let anchorDate2 = null, anchorIdx2 = -1;
  for (let c = anchorIdx + 1; c < numCols; c++) {
    if (dates[c]) { anchorDate2 = dates[c]; anchorIdx2 = c; break; }
  }

  let use7Day = true;
  if (anchorDate2 && anchorIdx2 > anchorIdx) {
    const colSpan = anchorIdx2 - anchorIdx;
    const daySpan = Math.round((anchorDate2 - anchorDate) / 86400000);
    use7Day = (daySpan <= colSpan + 2);   // ~1 day/col → calendar; ~1.4 day/col → working
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
//  SECTION HEADER DETECTOR
// ════════════════════════════════════════════════════════════
function isSectionHeader_(v) {
  if (v === null || v === undefined || v === '') return false;
  const n = Number(v);
  if (isNaN(n) || n <= 0) return false;
  return Number.isInteger(n);
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



function fixDayNumFormat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('גאנט עבודה FLL שנתי');
  if (!sheet) { SpreadsheetApp.getUi().alert('Sheet not found'); return; }
  const lastCol = sheet.getLastColumn();
  const firstCol = 9;
  const numCols = lastCol - firstCol + 1;
  if (numCols <= 0) return;
  const row8 = sheet.getRange(8, firstCol, 1, numCols);
  row8.clearFormat();
  row8.setNumberFormat('0');
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Done!');
}


function diagHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('גאנט עבודה FLL שנתי');
  const firstCol = 9; // Column I
  const lastCol  = sheet.getLastColumn();
  const numCols  = lastCol - firstCol + 1;

  // Row 5 – date anchors (non-empty cells only)
  const r5 = sheet.getRange(5, firstCol, 1, numCols).getValues()[0];
  const r6 = sheet.getRange(6, firstCol, 1, numCols).getValues()[0];
  const r7 = sheet.getRange(7, firstCol, 1, numCols).getValues()[0];
  const r8 = sheet.getRange(8, firstCol, 1, numCols).getValues()[0];

  // Collect non-empty anchors from row 5
  const anchors = [];
  for (let i = 0; i < numCols; i++) {
    if (r5[i] instanceof Date || (r5[i] && r5[i] !== '')) {
      const val = r5[i] instanceof Date
        ? Utilities.formatDate(r5[i], 'Asia/Jerusalem', 'dd/MM/yyyy EEE')
        : String(r5[i]);
      anchors.push('col' + (firstCol+i) + '=' + val);
    }
  }

  // Week labels row 6
  const weeks = [];
  for (let i = 0; i < numCols; i++) {
    if (r6[i] && r6[i] !== '') weeks.push('col' + (firstCol+i) + '=' + r6[i]);
  }

  // First 20 day numbers row 8
  const days = [];
  for (let i = 0; i < Math.min(numCols,30); i++) {
    if (r8[i] !== '' && r8[i] !== null) days.push('col'+(firstCol+i)+'='+r8[i]);
  }

  // Last 10 day numbers
  const lastDays = [];
  for (let i = numCols-20; i < numCols; i++) {
    if (r8[i] !== '' && r8[i] !== null) lastDays.push('col'+(firstCol+i)+'='+r8[i]);
  }

  // Last week labels
  const lastWeeks = [];
  for (let i = numCols-25; i < numCols; i++) {
    if (r6[i] && r6[i] !== '') lastWeeks.push('col'+(firstCol+i)+'='+r6[i]);
  }

  console.log('TOTAL_COLS:', numCols, 'LAST_COL:', lastCol);
  console.log('ROW5_ANCHORS:', JSON.stringify(anchors));
  console.log('ROW6_WEEKS_FIRST:', JSON.stringify(weeks.slice(0,15)));
  console.log('ROW6_WEEKS_LAST:', JSON.stringify(lastWeeks));
  console.log('ROW8_DAYS_FIRST:', JSON.stringify(days));
  console.log('ROW8_DAYS_LAST:', JSON.stringify(lastDays));
}
