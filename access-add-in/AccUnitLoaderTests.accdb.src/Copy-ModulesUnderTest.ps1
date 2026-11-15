# Copies the add-in modules under test into this source folder (they are ignored by git),
# so the test database can be built from source with msaccess-vcs.
# Usage: .\Copy-ModulesUnderTest.ps1, then "Build From Source" of AccUnitLoaderTests.accdb.

$moduleNames = @('modVersionCompare')

$addInModulesFolder = Join-Path $PSScriptRoot '..\source\modules'
$testModulesFolder = Join-Path $PSScriptRoot 'modules'

foreach ($moduleName in $moduleNames) {
    $sourceFile = Join-Path $addInModulesFolder "$moduleName.bas"
    if (-not (Test-Path -LiteralPath $sourceFile)) {
        throw "Module file not found: $sourceFile"
    }
    Copy-Item -LiteralPath $sourceFile -Destination $testModulesFolder -Force
    Write-Host "Copied $moduleName.bas"
}
