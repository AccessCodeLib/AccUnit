## Installing the add-in
1. Run AccUnitLoader_Install.vbs

    The add-in is copied to %appdata%\Microsoft\AddIns.
    The script lets you choose whether the add-in is used as accda or compiled to accde.
2. In Access, install the add-in 'ACLib AccUnit Loader' via Database Tools -- Add-ins -- Add-in Manager.

## Preparing an application for AccUnit
1. Open the add-in.
2. In the add-in form, select the folder the files are copied to.
3. Export the files from the add-in to that folder.
4. (optional) Set the tlb reference in the current application (button in the add-in form).
5. Insert the AccUnit factory module into the application.

## Building the add-in from source (development)
The AccUnit files (DLLs, tlb) and the icon are not stored in the repository.
Releases are created by the GitHub workflow `.github/workflows/release.yml`: it builds the DLLs,
creates `AccUnitLoader.accda` with msaccess-vcs and loads the files into the table `USys_AppFiles`.

Local build:
1. Build `source` into `AccUnitLoader.accda` with msaccess-vcs ("Build From Source").
2. Provide the AccUnit files in one folder: extract `AccUnit.*.bin.zip` from a release,
   or build the C# solutions and copy the output of `source\AccUnit\bin\<configuration>`
   and `vbe-add-In\AccUnit.VbeAddIn\bin\<configuration>` into one folder.
   Required: AccUnit.dll, AccUnit.tlb, AccessCodeLib.Common.Tools.dll, AccessCodeLib.Common.VBIDETools.dll,
   AccUnit.VbeAddIn.dll, AccessCodeLib.Common.VbeUserControlHost.dll
3. Load the files into `USys_AppFiles` from the immediate window of the add-in (VBA editor of AccUnitLoader.accda):

       ImportAccUnitFilesFromFolder "<folder with the AccUnit files>"
       ImportAppIconFile "<repository>\access-add-in\resources\ACLib.ico"

   `ImportAccUnitFiles` (without parameter) uses the AccUnit folder configured in the add-in.

## Testing the add-in
The tests of the add-in (e.g. version comparison) are in `AccUnitLoaderTests.accdb.src`.
1. Run `AccUnitLoaderTests.accdb.src\Copy-ModulesUnderTest.ps1` (copies the modules under test from `source\modules`).
2. Build `AccUnitLoaderTests.accdb` with msaccess-vcs ("Build From Source").
3. Run the tests with the AccUnit add-in.
