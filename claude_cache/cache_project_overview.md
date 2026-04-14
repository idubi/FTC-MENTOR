# Cache: Project Overview & Context
**Last updated:** 2026-04-13

---

## Project
FTC/FLL Robotics Mentors Gantt Director — שדרות FLL 2026-27

## Goal
Manage and automate a Gantt chart in Google Sheets for FLL robotics mentoring program scheduling. The Gantt uses cell background colors to represent task bars, and an Apps Script automatically syncs date columns from those colors.

## User Profile
- Ido (ido.bistry@gmail.com)
- Fullstack/DevOps/AI developer, system analyst, technologist
- Fluent in JS/TS/Python/Java, SQL, shell, K8s, Docker, CI/CD, AI/ML
- Owner account for sheets: office1.aido@gmail.com

## Key Decisions Made
1. **Color-based Gantt sync** — cells painted with any non-default color = active bar
2. **Israeli work week** — Sun–Thu (Fri/Sat hidden columns)
3. **Apps Script (GAS)** — only way to read cell bg colors in Google Sheets
4. **Week 1 starts Dec 28, 2025** (Sunday before Jan 1, 2026)
5. **58 weeks minimum** — covers through Feb 4, 2027
6. **Section headers** = integer in col A (skipped); tasks = decimal (synced)
7. **Columns D/E/F** auto-populated by script (overrides any formulas)

## Current Status
- GanttColorSync.gs: Written, tested design approved
- Installation guide: Created (Hebrew + English)
- Design document: Complete
- Restructure function: Built for one-time week structure fix
- Next: Installation on live sheet, testing, iteration
