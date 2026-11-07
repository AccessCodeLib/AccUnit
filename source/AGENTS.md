# AGENTS.md

Additional rules for this folder. The rules of the parent `AGENTS.md` (language, way of working, git, files) apply as well.

## Project specifics
- Role: framework core of AccUnit (C# COM library that runs the tests); used by the Access/Excel add-ins and the VBE add-in.
- AccUnit supports row tests like NUnit (one test procedure, several data rows); keep this behavior compatible when changing the test engine.
- Tests must not run in parallel against one open Access application (parallel runs fail). Do not introduce parallel test execution or parallel VBA calls.
- Solution `AccessCodeLib.AccUnit.sln`: classic .NET Framework projects (core and Common projects target 3.5, test projects 4.8, NUnit). Two test projects are SDK-style (`net48`).
- The core targets .NET Framework 3.5 on purpose: it is called from Access/VBA via COM without installation (registration-free). Do not change target framework or use newer language/runtime features (`async`/`await`, nullable reference types, ...) there without asking. A migration may be considered later.
- Build with MSBuild/Visual Studio (NuGet restore, then build the `.sln`), not `dotnet build`. Tests run with NUnit/VSTest, as in `.azure-pipelines`.
- `packages\` and `lib\` contain third-party/binary content; do not modify.

## .NET / C#
- Follow common .NET naming conventions and the style of the existing solution.
- Do not use `async`/`await` or nullable reference types in the 3.5 core (see Project specifics). `async`/`await` is risky where VBA is remote-controlled (COM calls into or out of VBA, VBE automation: threading and COM apartment issues); such cases must be checked carefully before use.
- For persistence and data exchange use `CultureInfo.InvariantCulture`.
- After changes, build the solution with MSBuild and run the existing tests (NUnit/VSTest). Do not use `dotnet build` for the classic projects.
