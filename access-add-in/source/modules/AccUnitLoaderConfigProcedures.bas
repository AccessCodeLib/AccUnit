﻿Attribute VB_Name = "AccUnitLoaderConfigProcedures"
Option Explicit
Option Compare Text

#Const AccUnitEarlyBinding = 0

#If AccUnitEarlyBinding Then
Public Property Get CurrentAccUnitConfiguration() As AccUnitConfiguration
#Else
Public Property Get CurrentAccUnitConfiguration() As Object
#End If
   Set CurrentAccUnitConfiguration = New AccUnitConfiguration
End Property

Public Sub AddAccUnitTlbReference()
   RemoveAccUnitTlbReference
   modVbProject.CurrentVbProject.References.AddFromFile CurrentAccUnitConfiguration.AccUnitDllPath & "\" & ACCUNIT_TYPELIB_FILE
End Sub

Public Sub RemoveAccUnitTlbReference()

   Dim ref As VBIDE.Reference
   Dim RefName As String

   With modVbProject.CurrentVbProject
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

#If AccUnitEarlyBinding Then
   Dim Configurator As AccUnit.Configurator
#Else
   Dim Configurator As Object
#End If

   With New AccUnitLoaderFactory
      Set Configurator = .Configurator
   End With

   Configurator.InsertAccUnitLoaderFactoryModule AccUnitTlbReferenceExists, True, modVbProject.CurrentVbProject, Application
   Set Configurator = Nothing

On Error Resume Next
   Application.RunCommand acCmdCompileAndSaveAllModules

End Sub

Private Function AccUnitTlbReferenceExists() As Boolean

   Dim ref As VBIDE.Reference
   Dim RefName As String

   For Each ref In modVbProject.CurrentVbProject.References
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

#If AccUnitEarlyBinding Then
   Dim Configurator As AccUnit.Configurator
#Else
   Dim Configurator As Object
#End If

   With New AccUnitLoaderFactory
      Set Configurator = .Configurator
   End With

   Configurator.InsertAccUnitLoaderFactoryModule AccUnitTlbReferenceExists, False, modVbProject.CurrentVbProject, Application
   Configurator.ImportTestClasses
   Set Configurator = Nothing

On Error Resume Next
   Application.RunCommand acCmdCompileAndSaveAllModules

End Sub

Public Sub ExportTestClasses()

#If AccUnitEarlyBinding Then
   Dim Configurator As AccUnit.Configurator
#Else
   Dim Configurator As Object
#End If

   With New AccUnitLoaderFactory
      Set Configurator = .Configurator
   End With

   Configurator.ExportTestClasses
   Set Configurator = Nothing

End Sub

Public Sub RemoveTestEnvironment(ByVal RemoveTestModules As Boolean, Optional ByVal SaveTestModules As Boolean = True)

#If AccUnitEarlyBinding Then
   Dim Configurator As AccUnit.Configurator
#Else
   Dim Configurator As Object
#End If

   With New AccUnitLoaderFactory
      Set Configurator = .Configurator
   End With

   Configurator.RemoveTestEnvironment RemoveTestModules, SaveTestModules, modVbProject.CurrentVbProject
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
                        "AccUnit.VbeAddIn.dll", _
                        "AccessCodeLib.Common.VbeUserControlHost.dll")

End Property

Public Sub ExportAccUnitFiles()

   Dim AccUnitFileName As Variant
   Dim DllPath As String

On Error GoTo HandleErr

   DllPath = CurrentAccUnitConfiguration.AccUnitDllPath

   With modApplication.CurrentApplication.Extensions("AppFile")
      For Each AccUnitFileName In AccUnitFileNames
         .CreateAppFile AccUnitFileName, DllPath & AccUnitFileName
      Next
   End With

ExitHere:
   Exit Sub

HandleErr:
   If AccUnitFileName = ACCUNIT_TYPELIB_FILE Then
      Resume Next
   End If
   Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Public Sub ImportAccUnitFiles()

   Dim AccFileName As Variant
   Dim DllPath As String

   DllPath = CurrentAccUnitConfiguration.AccUnitDllPath

   With modApplication.CurrentApplication.Extensions("AppFile")
      For Each AccFileName In AccUnitFileNames
         .SaveAppFile AccFileName, DllPath & AccFileName, True
      Next
   End With

End Sub

Public Function AutomatedTestRunVCS() As Variant

    Dim ResultMessage As String
    Dim Success As Boolean

    Success = AutomatedTestRun(ResultMessage, TestReportOutput.DebugPrint + TestReportOutput.MsAccessVCS, False)
    If Success Then
        AutomatedTestRunVCS = "Success: " & ResultMessage
    Else
        AutomatedTestRunVCS = "Failed: " & ResultMessage
    End If

End Function

Public Function AutomatedTestRun(Optional ByRef ResultMessage As String, _
                                 Optional ByVal TestReportOutputTo As TestReportOutput = TestReportOutput.LogFile + TestReportOutput.DebugPrint, _
                                 Optional ByVal SetFocusToImmediateWindowBeforeTestStart As Boolean = True) As Boolean

   Dim Success As Boolean

#If AccUnitEarlyBinding Then
   Dim TestSummary As AccUnit.ITestSummary
#Else
   Dim TestSummary As Object
#End If

   AddAccUnitTlbReference
   InsertFactoryModule
   ImportTestClasses

   If SetFocusToImmediateWindowBeforeTestStart Then
      SetFocusToImmediateWindow
   End If

   Set TestSummary = AccUnitLoaderFactoryCall.GetAccUnitFactory.TestSuite(TestReportOutputTo).AddFromVBProject.Run.Summary
   Success = TestSummary.Success

   RemoveTestEnvironment True

   If Not Success Then
      ResultMessage = (TestSummary.Failed + TestSummary.Error) & " of " & TestSummary.Total & " tests failed"
   ElseIf TestSummary.Ignored > 0 Then
      ResultMessage = TestSummary.Ignored & " of " & TestSummary.Total & " tests ignored"
   Else
      ResultMessage = TestSummary.Total & " tests passed"
   End If

   AutomatedTestRun = Success

End Function

Private Sub SetFocusToImmediateWindow()
   Dim VbeWin As VBIDE.Window
   For Each VbeWin In Application.VBE.Windows
      If VbeWin.Type = VBIDE.vbext_WindowType.vbext_wt_Immediate Then
         If Not VbeWin.Visible Then
            VbeWin.Visible = True
         End If
         VbeWin.SetFocus
         Exit Sub
      End If
   Next
End Sub
