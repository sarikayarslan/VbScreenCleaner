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
- VB5/VB6 project cleaner (C#): `analyze <project dir>` reports unused screens/code, `clean <project dir> <target dir>` writes a cleaned copy + report. Never writes to the source project.
- Stack: .NET 9, C#, WinForms GUI (`src/VbScreenCleaner`), Core, Cli, xUnit. Solution: VbScreenCleaner.sln
- Build: `dotnet build VbScreenCleaner.sln -c Release`
- Test: `dotnet test tests/VbScreenCleaner.Tests`
- Run: `dotnet run --project src/VbScreenCleaner.Cli -- analyze samples/SampleVb5`
- VB files are ANSI (Windows-1254), read/written via `VbProject.FileEncoding`; never re-encode `samples/`.
- Never edit: `samples/` (test input), `bin/`, `obj/`.
