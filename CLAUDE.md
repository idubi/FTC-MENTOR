# Project Directives — FTC-MENTOR-GANT-DIRECTOR

## GOLDEN RULE: Cache-First Workflow

**Before doing ANY work, ALWAYS read the relevant `claude_cache/` files first.**

1. Start every session by reading `claude_cache/INDEX.md`
2. Read the specific `cache_[subject].md` files relevant to the current task
3. Use cache as the primary source of truth — avoid redundant scanning
4. After completing work, UPDATE the relevant cache files with new findings
5. Add new cache files for new subjects using `cache_[subject].md` naming

### Cache Location
```
FTC-MENTOR-GANT-DIRECTOR/claude_cache/
├── INDEX.md                  ← Read this first (quick lookup)
├── cache_gantt_structure.md  ← Sheet layout, rows, columns, colors
├── cache_apps_script.md      ← GanttColorSync.gs reference
├── cache_architecture.md     ← System design, data flows, algorithms
├── cache_google_drive.md     ← File locations, Drive URLs, access
├── cache_project_overview.md ← Goals, decisions, status
├── cache_methodology.md      ← Workflows, conventions, tools
└── cache_chronolog.md        ← Timeline of all work done
```

## Project Context
- FLL Robotics Mentors Gantt — שדרות FLL 2026-27
- Google Sheets Gantt with color-based task bars
- Apps Script (GanttColorSync.gs) auto-syncs D/E/F from colors
- Israeli work week: Sun–Thu
- Google Drive folder: https://drive.google.com/drive/folders/1XQ3QszXpyiAJP7yt1eFjevXr3jU9i_yE

## Key Rules
- Use `google-sheets-pm` skill for spreadsheet/PM tasks
- Cache files are small, focused, searchable — one subject per file
- Chronolog tracks all processes and decisions over time
- Architecture cache documents system design and data flows
- Always update cache when important things change
- **Update `cache_history_work.md` after EVERY feature, fix, or change** — document what was done, why, and which files changed

## Work Tracking Rule
Every time work is completed (feature, fix, refactor, documentation, testing):
1. Add entry to `claude_cache/cache_history_work.md`
2. Include: Date, Category, Module, Description, Status, Files Changed
3. Keep chronological order
4. Update category summary and module coverage tables
5. This ensures future sessions have complete work history
