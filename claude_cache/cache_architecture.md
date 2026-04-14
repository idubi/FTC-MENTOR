# Cache: System Architecture
**Last updated:** 2026-04-13

---

## GanttColorSync Architecture

### Three-Layer Design
```
┌─────────────────────────────────────────────────┐
│  LAYER 1 — TRIGGER                               │
│  onChange(e)                                      │
│  Fires on: any sheet change (format/edit)        │
└──────────────────┬──────────────────────────────┘
                   │
┌──────────────────▼──────────────────────────────┐
│  LAYER 2 — SCANNER: syncAllTaskRows_(sheet)      │
│  • Acquires document lock (15s)                  │
│  • Batch-reads row 5 dates + all backgrounds     │
│  • Builds date map from anchor dates             │
│  • Iterates task rows, skips section headers     │
│  • Writes D/E/F per row                          │
└──────────────────┬──────────────────────────────┘
                   │
┌──────────────────▼──────────────────────────────┐
│  LAYER 3 — RESOLVER: resolveBarDates_(row, dates)│
│  • Scans row backgrounds left→right              │
│  • Skips empty colors + Fri/Sat                  │
│  • Returns {startDate, endDate, duration}        │
└─────────────────────────────────────────────────┘
```

### Data Flow
```
Cell Color Paint → onChange trigger → syncAllTaskRows_()
  → batch read backgrounds (one API call)
  → batch read row 5 dates (one API call)
  → buildDateMap_() interpolates all column dates
  → per row: resolveBarDates_() → first/last colored → D/E/F
```

### Date Map Algorithm
```
1. Scan row 5 for anchor dates (merged month cells)
2. Find first anchor pair with ≥7 day span
3. Detect structure:
   - colSpan ≈ daySpan → 7-day (calendar days)
   - colSpan < daySpan → 5-day (working days)
4. Fill all dates by interpolation from anchors
```

### Week Restructure Flow
```
restructureToFullWeeks():
  1. Break all merges in rows 5-8
  2. Detect if Dec 28-31 cols exist (check row 7 col I)
  3. Insert 4 cols if missing
  4. Ensure ≥406 cols (58 weeks × 7)
  5. Build per-column date array from GANTT_START
  6. Show all cols, then hide Fri+Sat per week
  7. Write row 5: month labels (merged)
  8. Write row 6: week labels (merged per 5 visible)
  9. Write row 7: day names א–ש
  10. Write row 8: day numbers 1–31
```

### Concurrency
- Uses `LockService.getDocumentLock()` with 15s timeout
- Prevents race conditions with multiple simultaneous editors

### Performance
- Batch reads (one call for all backgrounds, one for dates)
- Per-row writes to D/E/F (could be optimized to batch writes)
- < 1 second for ~100 task rows

---

## Year Setup Architecture (NEW)

### Flow: setupYear()
```
User clicks "🗓 אתחל שנה עבודה"
  ↓
Prompt: "הכנס שנה (YYYY):"
  ↓
Validate year (2000-2100)
  ↓
buildYearGrid_(sheet, year)
  ↓
Display success alert
```

### buildYearGrid() Algorithm
```
1. findFirstThursdayOfYear_(year) → ISO week 1 anchor
2. weekStart = Sunday of that week
3. weekEnd = Saturday of week containing last Sunday of year
4. totalDays = weekEnd - weekStart
5. Ensure ≥ 406 columns (58 weeks × 7)

6. FOR each day in range:
     - Build dates[], dayNums[] arrays
     - Calculate week position in month → color from palette
     - Mark other-year days with #f3f3f3 (gray)

7. Write all arrays to rows 5, 6, 7
8. Merge month cells (row 5)
9. Merge week cells (row 6, 5-col spans for visible weeks)
10. Hide Fri+Sat columns per week
```

### Color Coding Strategy
- **5 color positions** representing week position in month (1st, 2nd, 3rd, 4th, 5th)
- **Repeats every month** — Jan week 1 = same color as Feb week 1, Mar week 1, etc.
- **Palette**: Light blue, light green, light yellow, light orange, light pink
- **Other-year dates**: Gray (#f3f3f3) overrides position color
