# AGENTS.md

Additional rules for this folder. The rules of the parent `AGENTS.md` (language, way of working, git, files) apply as well.

## Project specifics
- Role: framework core of AccUnit (C# COM library that runs the tests); used by the Access/Excel add-ins and the VBE add-in.
- AccUnit supports row tests like NUnit (one test procedure, several data rows); keep this behavior compatible when changing the test engine.
- Tests must not run in parallel against one open Access application (parallel runs fail). Do not introduce parallel test execution or parallel VBA calls.
- Solution `AccessCodeLib.AccUnit.sln`: classic .NET Framework projects (core and Common projects target 3.5, test projects 4.8, NUnit). Two test projects are SDK-style (`net48`).
- The core is called from Access/VBA without installation (CLR hosting from the Access add-in, no registry entries, no admin rights, `Access.exe.manifest` must not be touched). It must keep working this way.
- Target framework: the core is being migrated from .NET Framework 3.5 to 4.8 (branch `feature/net48-core`), because 3.5 is an optional Windows feature that needs admin rights to enable, while 4.8 ships with Windows 10 (1903+) and Windows 11. Stay on .NET Framework 4.8; do not move to .NET (Core)/.NET 8+ (no registration-free hosting without a runtime installation). Until the migration is finished, projects that still target 3.5 must not use newer runtime features. Do not use newer language features (`async`/`await`, nullable reference types, ...) without asking.
- The migration includes the CLR hosting in the add-ins (`NetComDomain` in `access-add-in/` and `excel-add-in/`): it has to start CLR 4 (`CLRCreateInstance`/`ICLRMetaHost`) instead of CLR 2 (`CorRuntimeHost`).
- Build with MSBuild/Visual Studio (NuGet restore, then build the `.sln`), not `dotnet build`. Tests run with NUnit/VSTest, as in `.azure-pipelines`.
- `packages\` and `lib\` contain third-party/binary content; do not modify.

## .NET / C#
- Follow common .NET naming conventions and the style of the existing solution.
- Do not use `async`/`await` or nullable reference types in the 3.5 core (see Project specifics). `async`/`await` is risky where VBA is remote-controlled (COM calls into or out of VBA, VBE automation: threading and COM apartment issues); such cases must be checked carefully before use.
- For persistence and data exchange use `CultureInfo.InvariantCulture`.
- After changes, build the solution with MSBuild and run the existing tests (NUnit/VSTest). Do not use `dotnet build` for the classic projects.
