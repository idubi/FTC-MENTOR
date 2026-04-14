# Cache: History of Work Completed
**Last updated:** 2026-04-14
**Purpose:** Track all features, fixes, and changes to the Gantt system

---

## Work Log Format
- **Date:** YYYY-MM-DD
- **Category:** Feature / Fix / Refactor / Documentation / Testing
- **Module:** Which part of system affected
- **Description:** What was done and why
- **Status:** Complete / In Progress / Blocked
- **Files Changed:** List of files modified/created

---

## Work Entries

### Entry 001: GanttColorSync Core Script
**Date:** 2026-04 (Session 1)  
**Category:** Feature  
**Module:** GanttColorSync.gs (Core)  
**Description:** 
- Built Google Apps Script that reads cell background colors in Gantt timeline
- Implements 3-layer architecture: Trigger → Scanner → Resolver
- Auto-syncs columns D (start), E (end), F (duration) from painted cells
- Automatic date detection and interpolation (7-day vs 5-day structures)
- Task: color-paint cell → D/E/F auto-update via onChange trigger

**Status:** Complete  
**Files Changed:** 
- Created: GanttColorSync.gs
- Created: Gantt_ColorSync_Design.md
- Created: GanttColorSync_Installation.md

---

### Entry 002: Cache System Implementation
**Date:** 2026-04-13 (Session 2)  
**Category:** Documentation + Process  
**Module:** Project Infrastructure  
**Description:**
- Established cache-first workflow golden rule
- Created 7 focused cache files for project knowledge
- Cache avoids redundant scanning between sessions
- Cache is source of truth for architecture, methods, status

**Status:** Complete  
**Files Changed:**
- Created: claude_cache/INDEX.md
- Created: claude_cache/cache_gantt_structure.md
- Created: claude_cache/cache_apps_script.md
- Created: claude_cache/cache_google_drive.md
- Created: claude_cache/cache_project_overview.md
- Created: claude_cache/cache_methodology.md
- Created: claude_cache/cache_architecture.md
- Created: claude_cache/cache_chronolog.md
- Created: CLAUDE.md (project directives)
- Created: feedback_cache_first.md (persistent memory)
- Created: MEMORY.md (memory index)

---

### Entry 003: Year Setup Feature
**Date:** 2026-04-13 (Session 2)  
**Category:** Feature  
**Module:** GanttColorSync.gs (Year Setup)  
**Description:**
- Interactive year initialization button "🗓 אתחל שנה עבודה"
- Creates full-year calendar grid (1/1 - 31/12) automatically
- ISO-8601-like week numbering (week 1 = first Thursday)
- 5-color repeating pattern based on week position in month
- Smart handling of year boundaries (Dec/Jan shown in gray)
- Auto-hides Fri+Sat columns per week
- Merges cells for month labels and week numbers
- Proper formatting: bold, centered, appropriate fonts

**Features:**
- setupYear() — interactive year prompt
- buildYearGrid_() — full calendar builder
- findFirstThursdayOfYear_() — ISO week 1 calculation
- getWeekPositionInMonth_() — color assignment logic
- getWeekColor_() — color palette lookup

**Row Output:**
- Row 5: Month labels (MMMM yyyy format, merged)
- Row 6: Week numbers (שבוע 1, שבוע 2, etc., merged)
- Row 7: Day of month (1-31, individual cells)
- Colors: 5 repeating palette + gray for other-year dates

**Status:** Complete  
**Files Changed:**
- Modified: GanttColorSync.gs (450+ lines added)
- Modified: onOpen() menu (added new button)
- Created: cache_year_setup.md
- Created: YEAR_SETUP_SUMMARY.md
- Modified: cache_apps_script.md
- Modified: cache_architecture.md
- Modified: cache_chronolog.md

---

### Entry 004: Bug Fixes — Year Setup v2
**Date:** 2026-04-13 (Session 2, continued)  
**Category:** Fix  
**Module:** GanttColorSync.gs (Year Setup)  
**Description:**
Fixed 4 bugs reported after first deployment:

1. **Columns extended past year end** → Removed `minCols = 58 * 7` constraint. Now calculates exact columns for the requested year (53-54 weeks typical).

2. **No year prompt on "בנה מחדש" button** → `restructureToFullWeeks()` replaced with single-line redirect to `setupYear()`. Both menu items now always prompt for year.

3. **Inconsistent Fri+Sat visibility** (some weeks showed 1, 3, or 5 days) → Fixed show/hide sequence: now calls `sheet.showColumns()` on ALL existing timeline columns first (clears stale hide state from previous runs), then hides exactly 2 columns (Fri+Sat) per week for every week. Uses `totalWeeks` from grid calculation, not a loop condition.

4. **No holiday support** → Added `getIsraeliHolidays_(year)` with 2025/2026/2027 Israeli public holidays (Pesach, Shavuot, RH, YK, Sukkot, Yom HaAtzmaut, etc.). Holidays shown in darker gray (#c0c0c0) vs other-year days (#d8d8d8). Uses `dateKey_()` helper for fast Set lookup.

**Additional changes:**
- `findFirstThursdayOfYear_()`: simplified using `(4 - dow + 7) % 7` formula
- `getWeekPositionInMonth_()` renamed to `getWeekPosInMonth_()` with cleaner logic
- Added `dateKey_(d)` utility function
- Updated file header comment block with v2 fix log

**Status:** Complete  
**Files Changed:**
- Modified: GanttColorSync.gs (buildYearGrid_, setupYear, restructureToFullWeeks, helpers)
- Modified: cache_history_work.md (this file)
- To update: cache_apps_script.md, cache_year_setup.md

---

### Entry 007: MM/YYYY Prompts + Month-Accurate Grid (v4)
**Date:** 2026-04-14 (Session 3)
**Category:** Feature
**Module:** GanttColorSync.gs (setupYear, buildYearGrid_)
**Description:**
Changed year prompts to month/year format (MM/YYYY). Grid now spans exactly from the start month to the end month, and week numbering begins at the correct ISO week for that point in the year (not always 1).

**Changes:**
- `setupYear()`: prompts now ask for "MM/YYYY" (e.g. 09/2026). Uses new `parseMonthYear_()` helper. Validation: endYear*12+endMonth >= startYear*12+startMonth. Calls `buildYearGrid_(sheet, startYear, startMonth, endYear, endMonth)`.
- `buildYearGrid_()`: new signature `(sheet, startYear, startMonth, endYear, endMonth)`.
  - `gridStart` = Sunday of week containing 1st of startMonth (was: Sunday of year's week 1).
  - `gridEnd` = Saturday of week containing last day of endMonth (was: Saturday of year's last week).
  - `initialWeekNum` computed via `getWeek1Sunday_()` of the year whose Thursday falls in gridStart's week — correctly handles late-Dec/early-Jan edge weeks.
  - `week1Sundays` Set now covers `gridStart.getFullYear()` to `gridEnd.getFullYear()` (not just startYear–endYear).
  - `isOtherRange` uses month/year key comparison (`curYear*12 + month`) instead of year-only comparison.
- `getWeek1Sunday_(year)`: new helper — returns the Sunday that starts ISO week 1 of `year`.
- `parseMonthYear_(text)`: new helper — parses "MM/YYYY" → `{month, year}` or null; accepts 1 or 2 digit month.

**Example:** Input 09/2026 → 12/2026: grid runs from Aug 30, 2026 (Sun, gray edge) to Jan 3, 2027 (Sat, gray edge). First week label = "שבוע 36". Jan 3, 2027 edge week shows gray.

**Status:** Complete
**Files Changed:**
- Modified: GanttColorSync.gs (setupYear, buildYearGrid_, new helpers)
- Modified: cache_history_work.md

---

### Entry 006: Fix Month Label Timezone Shift (Nov→Nov bug)
**Date:** 2026-04-13 (Session 2)
**Category:** Fix
**Module:** GanttColorSync.gs (buildYearGrid_)
**Description:**
Month labels in row 5 were shifted back 1 month (e.g. January showed as December, December as November — producing "November to November" for a full year).

**Root cause:** `new Date(curYear, month, 1)` creates midnight local time. In Israel (UTC+2), Jan 1 2026 00:00 local = Dec 31 2025 22:00 UTC. Google Sheets stores the UTC value → serial is Dec 31 → 'MMMM yyyy' displays "December 2025" instead of "January 2026". Every month showed as the previous month.

**Fix:** `new Date(curYear, month, 1, 12, 0, 0)` — use noon instead of midnight. Noon in UTC+2 = 10:00 UTC, safely within the correct calendar day in any timezone (UTC-12 to UTC+14).

**Status:** Complete
**Files Changed:**
- Modified: GanttColorSync.gs (1 line in buildYearGrid_)
- Modified: cache_history_work.md

---

### Entry 005: Multi-Year Support + Week Number Reset per Year (v3)
**Date:** 2026-04-13 (Session 2)  
**Category:** Feature  
**Module:** GanttColorSync.gs (Year Setup)  
**Description:**
- `setupYear()` now prompts TWICE: "שנת התחלה" then "שנת סיום"
- Single-year mode: enter the same year twice (e.g. 2026 / 2026)
- Multi-year mode: e.g. 2026 / 2027 — continuous grid across both years
- Week counter resets to שבוע 1 at the start of each new year's ISO week 1
- Pre-computes `week1Sundays` Set for all years in range; resets weekNum on match
- `isOtherYear` now checks `curYear < startYear || curYear > endYear` (range-aware)
- Holidays merged from all years in range into single Set
- `buildYearGrid_` signature: `(sheet, startYear, endYear)` — both callers updated
- Success alert shows "2026–2027" style range label

**Status:** Complete  
**Files Changed:**
- Modified: GanttColorSync.gs (setupYear, buildYearGrid_)
- Modified: cache_year_setup.md
- Modified: cache_history_work.md (this file)

---

## Category Summary

| Category | Count | Last Updated |
|----------|-------|--------------|
| Feature | 4 | 2026-04-14 |
| Fix | 1 | 2026-04-13 |
| Documentation | 1 | 2026-04-13 |
| Process/Infrastructure | 1 | 2026-04-13 |
| **Total** | **7 entries** | **2026-04-14** |

---

## Module Coverage

| Module | Features | Status |
|--------|----------|--------|
| GanttColorSync.gs (Core) | Color sync, trigger, sync engine | Complete |
| GanttColorSync.gs (Year Setup) | Year initialization, calendar builder | Complete |
| Cache System | 7 reference files, golden rule | Complete |
| Documentation | Design doc, install guide, summaries | Complete |

---

## Dependencies & Notes

- **Year Setup depends on:** GanttColorSync core (uses same CFG, rows 5-7)
- **Color sync depends on:** Rows 5 populated with dates (year setup provides this)
- **Testing pending:** Live deployment, week number validation, color display verification
- **Performance:** All functions optimized for batch operations (< 1s for 58 weeks × 365 days)

---

## Next Work Items (Backlog)

- [ ] Live testing with 2026, 2027 years
- [ ] Validate week boundaries and colors
- [ ] Test Gantt bar painting + D/E/F sync
- [ ] Document any adjustments from testing
- [ ] Performance optimization (if needed)
- [ ] User training documentation
