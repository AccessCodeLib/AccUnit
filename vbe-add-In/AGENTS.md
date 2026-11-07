# AGENTS.md

Additional rules for this folder. The rules of the parent `AGENTS.md` (language, way of working, git, files) apply as well.

## Project specifics
- Role: VBE add-in UI of AccUnit (start tests, select tests, show results in a tree view, generate test procedures).
- Solution `AccessCodeLib.AccUnit.VbeAddIn.sln`: classic .NET Framework 4.8 add-in project; it also references the projects in `..\source`. Test project `WpfUserControlsTests` is SDK-style (`net8.0`, NUnit).
- Build the solution with MSBuild/Visual Studio (NuGet restore, then build the `.sln`), not `dotnet build`. `dotnet test` is possible for the SDK-style test project.
- The Access files in this folder (`LocalTest*.accdb`, `NW2.accdb`, `TestAC.accdb`) are only for manual testing of the add-in. They are binary; do not edit or open them without asking.
- `packages\` and `lib\` contain third-party/binary content; do not modify.

## .NET / C#
- Follow common .NET naming conventions and the style of the existing solution.
- `async`/`await` is allowed for the UI (forms, WPF), as long as it causes no problems with the VBE control. Where VBA or the VBE is remote-controlled (COM calls into or out of VBA/VBE: threading and COM apartment issues), it is risky and must be checked carefully before use.
- For persistence and data exchange use `CultureInfo.InvariantCulture`.
- After changes, build the solution with MSBuild and run the existing tests (NUnit/VSTest); `dotnet test` only for the SDK-style test project.
