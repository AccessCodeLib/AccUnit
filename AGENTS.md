# AGENTS.md

Instructions for AI coding agents. Project-specific rules take precedence over general ones.

## Project overview
- Goal: AccUnit is a unit-test framework for Access/Excel VBA code. A C# COM core (`source/`) runs the tests. The Access/Excel add-ins load it registration-free and start the tests, and the VBE add-in (`vbe-add-In/`) provides the UI to start tests, show results and generate test procedures.
- Repository layout:
  - `access-add-in/`: Access add-in (VBA, exported with msaccess-vcs-addin). Loads the C# core without installation, starts tests (also usable as test starter, for example from msaccess-vcs) and can install the VBE add-in. Has its own AGENTS.md.
  - `source/`: C# solution `AccessCodeLib.AccUnit.sln` (framework core). Has its own AGENTS.md.
  - `vbe-add-In/`: C# solution `AccessCodeLib.AccUnit.VbeAddIn.sln` (VBE add-in). Has its own AGENTS.md.
  - `excel-add-in/`: Excel add-in (VBA); no own AGENTS.md.
  - `examples/`, `binaries/`: examples and released binaries; do not modify.
- Rules of the subfolder AGENTS.md apply in addition to this file.

## Language
- Reply in the language the user writes in. If unclear, use the Windows display language (`Get-UICulture`).
- Code, identifiers, code comments, commit messages and branch names are always in English.

## Way of working
- Before changing files, propose a short plan and wait for confirmation. Exceptions: read-only access, analysis, and temporary files outside the project.
- Ask when requirements are unclear instead of making assumptions. With several possible solutions, give a recommendation with reasons.
- Keep changes small and traceable. No unrequested refactoring of existing code. Report findings outside the task, do not fix them.
- Follow the existing style and conventions of the project.
- Never delete, overwrite, commit or push without explicit approval. An approval applies only to that action, not to later ones.
- After changes, summarize briefly: what changed, what was tested, what is open.

## Git
- Commit messages in imperative mood, subject line max. ~72 characters.
- No force push, no rewriting of already pushed history.
- No attribution to AI tools: no `Co-Authored-By: ...` trailer in commit messages, no "Generated with ..." line in pull request descriptions.

## Files and environment
- Keep the encoding and line endings (CRLF) of existing files.
- In PowerShell 5.1 always specify the encoding explicitly when writing files.
- Windows 11. Take the regional settings (decimal separator, date format) into account when parsing and formatting. For persistence and data exchange use locale-independent formats.

## Clean code (all languages)
Principles of Clean Code / Clean Code Developers, as rules for everyday work:

### Structure
- **Single responsibility (SRP):** each procedure/method does exactly one thing, each class/module has one reason to change. Split long procedures into named steps.
- **Separation of concerns:** keep logic, UI and data access apart. Event handlers contain no logic; they only call well-named procedures.
- **Keep code testable:** logic does not depend on UI or on concrete data access; dependencies are passed in.

### Simplicity
- **KISS:** choose the simplest solution that works and can be read at once.
- **YAGNI:** build only what is needed now; no speculative options, parameters or layers.
- **DRY:** no duplicated logic - extract shared code instead of copying it. Do not abstract before the duplication is real.
- **Principle of least astonishment:** names, behavior and side effects match what a reader expects.

### Readability
- Meaningful, intention-revealing names; no abbreviations that need explanation (established prefixes of the project's style guide are allowed).
- No magic numbers or strings: use named constants or enums.
- Comments explain why, not what. Remove commented-out and dead code in the parts you change; report it elsewhere.
- Tell, don't ask; avoid long call chains (law of Demeter).

### Safety net and discipline
- **Errors:** never swallow errors silently. Handle an error where it can be handled, otherwise pass it on.
- **Tests:** logic is covered by automated tests; refactor only with tests as safety net.
- **Boy scout rule within the task:** leave code you touch a little cleaner. Findings outside the task are reported, not fixed (no unrequested rewrites).
- Work in small, understandable steps.

## Secure coding (all languages)
- Treat every input as untrusted (user input, form fields, files, command-line arguments, database content, results of other systems). Validate type, length, range and format; prefer allow-lists over block-lists.
- Prevent injection: never concatenate unchecked values or identifiers into SQL, filter or expression strings, shell commands or file paths.
  Use parameter queries; where dynamic SQL is unavoidable, use the safe building blocks of the platform or project.
- No secrets in code, repository, logs or test data (passwords, tokens, connection strings with credentials). Keep them in `.env.local` or the secret store of the environment; keep such files out of version control.
- Error messages shown to users contain no internal details (SQL text, paths, stack information). Details go to the log.
- Least privilege: request only the permissions and access that are needed; do not widen them to make a problem disappear.
- Handle files and paths defensively: build paths from validated parts, check that the target is where you expect it, never delete or overwrite based on unchecked input.
- Do not invent your own cryptography, hashing or authentication; use proven platform or library functions.
- Use third-party and library code only from known sources; report outdated versions instead of updating them unasked.
- Do not execute strings as code (`Eval`, `Application.Run` with user input, `Shell`, `Invoke-Expression` and similar) unless the content is fully under your control.
