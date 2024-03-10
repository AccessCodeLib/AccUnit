Attribute VB_Name = "defGlobal_AccUnitLoader"
'---------------------------------------------------------------------------------------
' Modul: defGlobal_AccUnitLoader
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Public Enum CodeLibElementType  'angelehnt an Enum vbext_ComponentType
   clet_StdModule = 1           ' = vbext_ComponentType.vbext_ct_StdModule
   clet_ClassModule = 2         ' = vbext_ComponentType.vbext_ct_ClassModule
   clet_Form = 101              ' = vbext_ComponentType.vbext_ct_Document + 1
   clet_Report = 102            ' = vbext_ComponentType.vbext_ct_Document + 2
End Enum

Public Enum CodeLibImportMode
   clim_ImportMissingItems = 0  ' überschreibt keine vorhandene Access-Objekte in der Anwendung
   clim_ImportSelectedOnly = 1  ' nur die ausgewählte Datei wird importiert (keine Abhängigkeistprüfung)
   clim_ImportAllUsedItems = 2  ' auch vorhandene Access-Objekte werden überschrieben
End Enum

Public Type CodeLibInfoReference
   Name As String
   Major As Long
   Minor As Long
   GUID As String
End Type

Public Type CodeLibInfo
   Name As String
   Type As CodeLibElementType
   RepositoryFile As String
   LocalFile As String
   RepositoryFileReplacement As String
   Dependency() As String
   References() As CodeLibInfoReference
   TestFiles() As String
   ExecuteList() As String
   LicenseFile As String
   Description As String
End Type

Public Enum TestReportOutput
   DebugPrint = 1
End Enum

'Standard-Icon
Public ACLibIconFileName As String 'Nur Dateiname inkl. Dateierweiterung, aber ohne vollständigen Pfad
