# Cache: Gantt Sheet Structure & Layout
**Last updated:** 2026-04-13

---

## Sheet Name
`גאנט עבודה FLL שנתי`

## Project
FLL Robotics Mentors Gantt — שדרות FLL 2026-27

## Header Rows (columns I onward)

| Row | Purpose | Content |
|-----|---------|---------|
| 5 | DATE_ROW | Month anchor dates, merged per month, format "MMMM yyyy" (e.g. ינואר 2026) |
| 6 | WEEK_ROW | Week labels: שבוע 1, שבוע 2 … merged per visible week (5 cols) |
| 7 | DAY_ROW | Hebrew day names: א ב ג ד ה (ו ש in hidden cols) |
| 8 | DAY_NUM_ROW | Day-of-month numbers: 1, 2, 3 … 28, 29, 30, 31 |
| 9+ | FIRST_DATA_ROW | Task rows begin here |

## Column Layout (A–H = metadata, I+ = timeline)

| Col | Letter | Purpose |
|-----|--------|---------|
| 1 | A | מס"ד — Task/section number (integer = section header, decimal = task) |
| 2 | B | Task name / description |
| 3 | C | (varies) |
| 4 | D | תאריך התחלה — Start Date (auto-synced by script) |
| 5 | E | תאריך סיום — End Date (auto-synced by script) |
| 6 | F | משך — Duration in working days (auto-synced by script) |
| 7 | G | (varies) |
| 8 | H | (varies) |
| 9+ | I+ | Timeline / Gantt bar area — each column = 1 calendar day |

## Week Structure
- Each week = 7 columns (Sun–Sat)
- Visible: Sun(א) Mon(ב) Tue(ג) Wed(ד) Thu(ה) = 5 columns
- Hidden: Fri(ו) Sat(ש) = 2 columns (exist but hidden)
- Week 1 starts: Dec 28, 2025 (Sunday)
- Week 58 ends: Feb 4, 2027 (Thursday)
- Israeli work week: Sun–Thu

## Section Headers vs Task Rows
- Section header: Column A = whole integer (1, 2, 3, 4, 5) → SKIPPED by sync
- Task row: Column A = decimal (1.1, 2.3, etc.) → synced
- Empty Column A → skipped

## Color Detection
A cell is part of a Gantt bar if its background is NOT in the "empty" set:
```
null, '', '#ffffff', '#ffffffff', '#f3f3f3', '#efefef', '#f1f3f4', '#ffffff00'
```
Any other color = active bar cell.

## Date Map Logic
- Row 5 holds anchor dates (merged cells, one per month section)
- `buildDateMap_()` interpolates all column dates from these anchors
- Auto-detects 7-day vs 5-day structure using anchor pair distance
- For 7-day: offset = calendar days from anchor
- For 5-day: offset = working days from anchor

## Key Dimensions
- Timeline starts at column I (index 9, 1-based)
- Minimum 58 weeks × 7 cols = 406 timeline columns
- Fri/Sat skipped in bar resolution (getDay() === 5 || 6)
