# Team rules (keep this file < 60 lines; it loads every turn)

## Output
- Terse. No preamble, no recap of what the user said, no closing summary unless asked.
- Never paste code you just wrote back into chat. Point to path:line.

## Orchestration (main session = lead)
- Trivial change (1 file, location known): do it yourself. Don't spawn agents.
- Otherwise: scout → plan → builder (one task each) → tester → reviewer (once per feature).
- Give subagents everything they need: exact files/lines, acceptance command, constraints. They start cold.
- Independent tasks on different files: spawn builders in parallel in one message.
- Don't re-read files a subagent already summarized. Trust its pointers.
- Check `.claude/context/` before scouting; reuse saved notes.

## Tools
- Library/framework docs → Context7 (via scout, or directly for one lookup). Not WebSearch/WebFetch.
- Bash output is compressed by RTK hook if installed; `rtk bypass <cmd>` only when the full output is truly needed.
- `/compact` at breakpoints (after research, after a feature, after a failed approach), never mid-implementation.
- Use Sonnet as default; `/model opus` only for architecture or hard debugging.

## Plan format (in chat, not a file, unless > 5 tasks)
T1: <goal> | files: a.ts, b.ts | check: `npm test -- a.test.ts`
## Project
- VbOrphanHandlers (C# WinForms, `src/VbOrphanHandlers`): reports VB5 event handlers (`x_Click` ...) whose control is missing, plus control-array Index mismatches. Read-only on the VB project.
- Logic: `OrphanScanner.cs`; GUI: `MainForm.cs`; headless: `VbOrphanHandlers.exe <dir|vbp> [lib...] -o report.txt`.
- Build: `dotnet build src/VbOrphanHandlers -c Release`
- VB files are ANSI (Windows-1254), read only; never re-encode `samples/`.
- Never edit: `samples/` (test input), `bin/`, `obj/`.
