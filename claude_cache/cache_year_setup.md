# Cache: Year Setup Process
**Last updated:** 2026-04-13

---

## Overview
"Initialize Year" button creates a full year grid (1/1 to 31/12) organized by weeks, with automatic month labels, week numbers, day numbers, and color-coding.

## User Interaction (v3 — multi-year)
1. User clicks "🗓 אתחל שנה עבודה"
2. **Prompt 1:** "שנת התחלה" (e.g. 2026)
3. **Prompt 2:** "שנת סיום" (e.g. 2027, or same year for single year)
4. System clears rows 5-7 and builds continuous grid
5. Week numbering resets to 1 at each new year's ISO week 1

## Week Numbering Logic (ISO 8601-like)
**Week 1** = week containing the **first Thursday** of the year
**Week N (last)** = week containing the **last Sunday** of the year

Examples:
- 2026: Jan 1 is Thursday → Dec 28, 2025 (Sun) starts week 1
- 2027: Jan 1 is Friday → Dec 27, 2026 (Sun) starts week 1
- 2024: Jan 1 is Monday → Dec 25, 2023 (Sun) starts week 1

## Column Structure
- 7 columns per week (Sun–Sat)
- Visible: 5 columns (Sun–Thu = א ב ג ד ה)
- Hidden: 2 columns (Fri–Sat = ו ש)

## Output Grid
Starting from column I (1-based = column 9):

### Row 5: Month Labels
- Cell value: Date object (first day of month section)
- Format: "MMMM yyyy" (ינואר 2026)
- Merged: across all columns of that month section
- Alignment: center, bold

### Row 6: Week Numbers
- Cell value: "שבוע 1", "שבוע 2", etc.
- Merged: across visible columns of that week (5 cols)
- Alignment: center, bold

### Row 7: Day of Month Numbers
- Cell value: 1–31 (calendar day number)
- No merge
- Alignment: center
- Font size: 8

## Color Coding (Week Position in Month)
Each week within a month has a color based on its **position in that month**, not the month itself:

- **Week position 1** (1st week touching month) → Color A (light blue)
- **Week position 2** (2nd week touching month) → Color B (light green)
- **Week position 3** (3rd week touching month) → Color C (light yellow)
- **Week position 4** (4th week touching month) → Color D (light orange)
- **Week position 5** (5th week touching month) → Color E (light pink)

This pattern **repeats across all months** — every month's 1st week = Color A, etc.

### Special: Days from Other Years
If a week contains days from the previous or next year (gray background):
- Column background: #f3f3f3 (light gray)
- Text: visible but dim
- Still participates in week/month grid

## Algorithm: buildYearGrid(year)
```
1. Find firstThursday = first Thursday of year
2. weekStartDate = Sunday of that week (firstThursday - 4 days)
3. Find lastSunday = last Sunday of year
4. weekEndDate = Sunday + 6 days (Saturday of that week)

5. totalDays = weekEndDate - weekStartDate
6. Initialize arrays: dates[], monthLabels[], weekNumbers[], dayNumbers[], colors[]

7. FOR each day from weekStartDate to weekEndDate:
     dates[colIndex] = day
     monthLabels[colIndex] = month name (if first day of month)
     weekNumbers[colIndex] = week number
     dayNumbers[colIndex] = day of month (1-31)
     colors[colIndex] = (day is from other year) ? gray : colorForWeekPosition()

8. Build merges for row 5 (month labels) and row 6 (week labels)
9. Write all values and colors to sheet
10. Hide Fri+Sat columns per week
```

## Pseudo-code: Determine Week Position in Month
```
function getWeekPositionInMonth(date) {
  // Given a date, how many weeks into its month does it fall?
  monthStart = new Date(date.year, date.month, 1)
  firstSunday = (monthStart.day === 0) ? monthStart : 
                monthStart - (monthStart.day - 1) days
  weekStartOfDate = date - date.day days (Sunday of week containing date)
  
  weekPosition = (weekStartOfDate - firstSunday) / 7 + 1
  return Math.min(weekPosition, 5)  // Cap at 5 for edge weeks
}
```

## Color Map (v2)
```javascript
// Week position colors (repeating per month)
[
  '#dce9f5',   // pos 1 — light blue
  '#ddf2e3',   // pos 2 — light green
  '#fdf5d4',   // pos 3 — light yellow
  '#fde8d4',   // pos 4 — light orange
  '#f5d4ea',   // pos 5 — light pink
]

const GRAY_OTHER_YEAR = '#d8d8d8';   // days outside requested year
const GRAY_HOLIDAY    = '#c0c0c0';   // Israeli public holidays (disabled)
```

## v2 Bug Fixes (2026-04-13)
1. Removed `minCols = 58*7` — now uses exact weeks for year (no 2027 bleed)
2. `restructureToFullWeeks()` redirects to `setupYear()` — always prompts for year
3. Fri+Sat: `showColumns` ALL first, then `hideColumns(friCol, 2)` per week
4. Israeli holidays 2025–2027 hardcoded; shown in gray #c0c0c0

## Integration with Existing Menu
Add to "🔄 סנכרון גאנט" menu:
- Separator
- "🗓 אתחל שנה עבודה" → `setupYear()`

## After Year Setup
Once rows 5–8 are populated:
- Tasks can be added starting at row 9
- Color-syncing (GanttColorSync) auto-updates D/E/F columns
- Gantt bars can be painted with any color (colors differ from week background)
