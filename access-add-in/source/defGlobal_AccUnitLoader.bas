Attribute VB_Name = "defGlobal_AccUnitLoader"
'---------------------------------------------------------------------------------------
' Modul: defGlobal_AccUnitLoader
'---------------------------------------------------------------------------------------
'/**
' <summary>
' AccUnitLoader
' </summary>
' <remarks>
' Indiviuell gestaltete Config-Module nicht in das Repositiory laden.
' </remarks>
' \ingroup ACLibAddInImportWizard
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/AccUnitLoader/defGlobal_AccUnitLoader.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/AccUnitLoader/AccUnitConfiguration.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

' Integrierte Erweiterungen
Private Const EXTENSION_KEY_ACLIBFILEMANAGER As String = "ACLibAccUnitStarter"
Private Const EXTENSION_KEY_AccUnitConfiguration As String = "AccUnitConfiguration"


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


'Standard-Icon
Public ACLibIconFileName As String 'Nur Dateiname inkl. Dateierweiterung, aber ohne vollständigen Pfad

Public Property Get CurrentAccUnitConfiguration() As AccUnitConfiguration

   Set CurrentAccUnitConfiguration = CurrentApplication.Extensions(EXTENSION_KEY_AccUnitConfiguration)
   
End Property

Public Property Get AccUnitFileNames() As Variant()

   AccUnitFileNames = Array( _
                        ACCUNIT_TYPELIB_FILE, _
                        ACCUNIT_DLL_FILE, _
                        "AccessCodeLib.Common.Tools.dll", _
                        "AccessCodeLib.Common.VBIDETools.dll", _
                        "AccessCodeLib.Common.VBIDETools.XmlSerializers.dll", _
                        "Interop.TLI.dll", _
                        "Microsoft.Vbe.Interop.dll")

End Property

Public Sub ExportAccUnitFiles(Optional ByVal lBit As Long = 0)

   Dim accFileName As Variant
   Dim sBit As String
   Dim DllPath As String
   
   If lBit = 0 Then
      lBit = GetCurrentAccessBitSystem
   End If
   
   sBit = CStr(lBit)
   DllPath = CurrentAccUnitConfiguration.AccUnitDllPath

   With CurrentApplication.Extensions("AppFile")
      For Each accFileName In AccUnitFileNames
         .CreateAppFile accFileName, DllPath & accFileName, "BitInfo", sBit
      Next
   End With

End Sub

Public Sub ImportAccUnitFiles(Optional ByVal lBit As Long = 0)

   Dim accFileName As Variant
   Dim sBit As String
   Dim DllPath As String
   
   If lBit = 0 Then
      lBit = GetCurrentAccessBitSystem
   End If
   
   sBit = CStr(lBit)
   DllPath = CurrentAccUnitConfiguration.AccUnitDllPath

   With CurrentApplication.Extensions("AppFile")
      For Each accFileName In AccUnitFileNames
         .SaveAppFile accFileName, DllPath & accFileName, True, , , "BitInfo", sBit
      Next
   End With

End Sub

Public Function GetCurrentAccessBitSystem() As Long

#If VBA7 Then
#If Win64 Then
      GetCurrentAccessBitSystem = 64
#Else
      GetCurrentAccessBitSystem = 32
#End If
#Else
      GetCurrentAccessBitSystem = 32
#End If

End Function
