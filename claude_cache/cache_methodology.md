# Cache: Workflows, Conventions & Methodology
**Last updated:** 2026-04-13

---

## Golden Rule: Cache-First Workflow
1. **Before scanning** — always check `claude_cache/` files first
2. **After completing work** — update relevant cache files
3. **Cache = source of truth** for project knowledge between sessions
4. **Small focused files** — one subject per file, easy to search

## File Naming Convention
```
cache_[subject].md
```
Examples: `cache_gantt_structure.md`, `cache_apps_script.md`

## When to Update Cache
- New architectural decisions
- Changed configurations or layouts
- New files/scripts created
- Process changes or new workflows
- Bug fixes that change behavior

## Session Workflow
1. Read relevant `cache_*.md` files
2. Understand context without re-scanning everything
3. Do the work
4. Update cache files with any new findings
5. Save deliverables to workspace folder

## Skills Available
- `google-sheets-pm` — Google Sheets + project management
- `xlsx` — Excel file handling
- `pptx` — PowerPoint
- `docx` — Word documents
- `pdf` — PDF handling

## Tools & Connectors
- Google Drive MCP: search and fetch docs
- Chrome browser: for direct sheet interaction if needed
- Bash shell: for file operations, code execution
- Read/Write/Edit: for local file management
