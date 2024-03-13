Attribute VB_Name = "AccUnitLoaderConfigProcedures"
Option Explicit
Option Compare Text

' Integrierte Erweiterungen
Private Const EXTENSION_KEY_AccUnitConfiguration As String = "AccUnitConfiguration"

Public Property Get CurrentAccUnitConfiguration() As AccUnitConfiguration
   Set CurrentAccUnitConfiguration = CurrentApplication.Extensions(EXTENSION_KEY_AccUnitConfiguration)
End Property

Public Sub AddAccUnitTlbReference()
   RemoveAccUnitTlbReference
   CurrentVbProject.References.AddFromFile CurrentAccUnitConfiguration.AccUnitDllPath & "\AccessCodeLib.AccUnit.tlb"
End Sub

Public Sub RemoveAccUnitTlbReference()

   Dim ref As VBIDE.Reference
   Dim RefName As String

   With CurrentVbProject
      For Each ref In .References
On Error Resume Next
         RefName = ref.Name
         If Err.Number <> 0 Then
            Err.Clear
            RefName = vbNullString
         End If
On Error GoTo 0
         If RefName = "AccUnit" Then
            .References.Remove ref
            Exit Sub
         End If
      Next
   End With

End Sub

Public Sub InsertFactoryModule()

   Dim Configurator As AccUnit.Configurator

   With New AccUnitLoaderFactory
      Set Configurator = .Configurator
   End With

   Configurator.InsertAccUnitLoaderFactoryModule AccUnitTlbReferenceExists, True, CurrentVbProject, Application
   Set Configurator = Nothing

On Error Resume Next
   Application.RunCommand acCmdCompileAndSaveAllModules

End Sub

Private Function AccUnitTlbReferenceExists() As Boolean

   Dim ref As VBIDE.Reference
   Dim RefName As String

   For Each ref In CurrentVbProject.References
On Error Resume Next
      RefName = ref.Name
      If Err.Number <> 0 Then
         Err.Clear
         RefName = vbNullString
      End If
On Error GoTo 0
      If RefName = "AccUnit" Then
         AccUnitTlbReferenceExists = True
         Exit Function
      End If
   Next

End Function

Public Sub ImportTestClasses()

   Dim Configurator As AccUnit.Configurator

   With New AccUnitLoaderFactory
      Set Configurator = .Configurator
   End With

   Configurator.InsertAccUnitLoaderFactoryModule AccUnitTlbReferenceExists, False, CurrentVbProject, Application
   Configurator.ImportTestClasses
   Set Configurator = Nothing

On Error Resume Next
   Application.RunCommand acCmdCompileAndSaveAllModules

End Sub

Public Sub ExportTestClasses()

   Dim Configurator As AccUnit.Configurator

   With New AccUnitLoaderFactory
      Set Configurator = .Configurator
   End With

   Configurator.ExportTestClasses
   Set Configurator = Nothing

End Sub

Public Sub RemoveTestEnvironment(ByVal RemoveTestModules As Boolean, Optional ByVal SaveTestModules As Boolean = True)

   Dim Configurator As AccUnit.Configurator

   With New AccUnitLoaderFactory
      Set Configurator = .Configurator
   End With

   Configurator.RemoveTestEnvironment RemoveTestModules, SaveTestModules, CurrentVbProject
   Set Configurator = Nothing

On Error Resume Next
   Application.RunCommand acCmdCompileAndSaveAllModules

End Sub


Public Property Get AccUnitFileNames() As Variant()

   AccUnitFileNames = Array( _
                        ACCUNIT_TYPELIB_FILE, _
                        ACCUNIT_DLL_FILE, _
                        "AccessCodeLib.Common.Tools.dll", _
                        "AccessCodeLib.Common.VBIDETools.dll", _
                        "AccessCodeLib.Common.VBIDETools.XmlSerializers.dll", _
                        "Microsoft.Vbe.Interop.dll")
   '                       "Interop.VBA.dll"
End Property

Public Sub ExportAccUnitFiles(Optional ByVal lBit As Long = 0)

   Dim accFileName As Variant
   Dim sBit As String
   Dim DllPath As String

On Error GoTo HandleErr

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

ExitHere:
   Exit Sub

HandleErr:
   If accFileName = "AccessCodeLib.AccUnit.tlb" Then
      Resume Next
   End If
   Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

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

   If lBit = 32 Then
      DllPath = Replace(DllPath, "x64", "x86")
   ElseIf lBit = 64 Then
      DllPath = Replace(DllPath, "x86", "x64")
   End If

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

Public Function AutomatedTestRun(Optional ByRef FailedMessage As String) As Boolean

   Dim Success As Boolean
   Dim TestSummary As AccUnit.ITestSummary

   AddAccUnitTlbReference
   InsertFactoryModule
   ImportTestClasses

   SetFocusToImmediateWindow

   Set TestSummary = GetAccUnitFactory.TestSuite(LogFile + DebugPrint).AddFromVBProject.Run.Summary
   Success = TestSummary.Success

   RemoveTestEnvironment True

   If Not Success Then
      FailedMessage = TestSummary.Failed & " of " & TestSummary.Total & " Tests failed"
   End If

   AutomatedTestRun = Success

End Function

Private Sub SetFocusToImmediateWindow()
   Dim VbeWin As VBIDE.Window
   For Each VbeWin In Application.VBE.Windows
      If VbeWin.Type = vbext_wt_Immediate Then
         If Not VbeWin.Visible Then
            VbeWin.Visible = True
         End If
         VbeWin.SetFocus
         Exit Sub
      End If
   Next
End Sub
