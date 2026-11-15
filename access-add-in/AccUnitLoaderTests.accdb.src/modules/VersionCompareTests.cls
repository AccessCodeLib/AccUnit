VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VersionCompareTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

' Module under test: modVersionCompare from ..\..\source\modules
' (copied by Copy-ModulesUnderTest.ps1 before the test database is built)

'AccUnit:Row("0.10.0.0", "0.10.0.0", 0).Name("Equal")
'AccUnit:Row("0.9.11.0", "0.10.0.0", -1).Name("MinorPartIsComparedNumerically")
'AccUnit:Row("0.10.0.0", "0.9.11.0", 1).Name("NewerMinorWithLowerPatch")
'AccUnit:Row("0.9.12.0", "0.9.11.0", 1).Name("NewerPatch")
'AccUnit:Row("0.99.99", "1.0", -1).Name("MajorPartWins")
'AccUnit:Row("0.10", "0.10.0.0", 0).Name("MissingPartsCountAsZero")
'AccUnit:Row("0.10.0.1", "0.10", 1).Name("AdditionalPartIsNewer")
'AccUnit:Row("2026.11.12", "2025.06.27", 1).Name("DateVersion")
Public Sub CompareVersions_ReturnsOrderOfVersions(ByVal Version1 As String, ByVal Version2 As String, ByVal Expected As Long)

   Dim Actual As Long

   Actual = modVersionCompare.CompareVersions(Version1, Version2)

   Assert.That Actual, Iz.EqualTo(Expected)

End Sub
