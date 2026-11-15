Attribute VB_Name = "AccUnitLoaderConfigProcedures"
Option Explicit
Option Compare Text

#Const AccUnitEarlyBinding = 0

Private Const APPFILE_ID_APPICON As String = "AppIcon"
Private Const ERR_ACCUNIT_FILE_MISSING As Long = vbObjectError + 5001

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
   ImportAccUnitFilesFromFolder CurrentAccUnitConfiguration.AccUnitDllPath
End Sub

' Loads all AccUnit library files from SourceFolder into usys_AppFiles.
' Used by the release build (msaccess-vcs-build app-config), so the folder is checked
' completely before the table is changed.
Public Sub ImportAccUnitFilesFromFolder(ByVal SourceFolder As String)

   Dim FolderPath As String
   Dim AccFileName As Variant

   FolderPath = FolderPathWithBackslash(SourceFolder)
   CheckAccUnitFilesExist FolderPath

   For Each AccFileName In AccUnitFileNames
      SaveAppFileOrRaise AccFileName, FolderPath & AccFileName, True
   Next

End Sub

' Loads the application icon shown in the add-in form into usys_AppFiles.
Public Sub ImportAppIconFile(ByVal IconFilePath As String)

   If Not FileTools.FileExists(IconFilePath) Then
      Err.Raise ERR_ACCUNIT_FILE_MISSING, "ImportAppIconFile", "Icon file not found: " & IconFilePath
   End If

   SaveAppFileOrRaise APPFILE_ID_APPICON, IconFilePath, False

End Sub

Private Sub CheckAccUnitFilesExist(ByVal FolderPath As String)

   Dim AccFileName As Variant
   Dim MissingFiles As String

   If Not FileTools.DirExists(FolderPath) Then
      Err.Raise ERR_ACCUNIT_FILE_MISSING, "CheckAccUnitFilesExist", "Folder not found: " & FolderPath
   End If

   For Each AccFileName In AccUnitFileNames
      If Not FileTools.FileExists(FolderPath & AccFileName) Then
         MissingFiles = MissingFiles & ", " & AccFileName
      End If
   Next

   If Len(MissingFiles) > 0 Then
      Err.Raise ERR_ACCUNIT_FILE_MISSING, "CheckAccUnitFilesExist", _
                "Missing files in " & FolderPath & ": " & Mid$(MissingFiles, 3)
   End If

End Sub

Private Sub SaveAppFileOrRaise(ByVal FileID As String, ByVal FilePath As String, ByVal SaveVersion As Boolean)

   ' SaveAppFile returns False (without error) for an empty file
   If Not modApplication.CurrentApplication.Extensions("AppFile").SaveAppFile(FileID, FilePath, SaveVersion) Then
      Err.Raise ERR_ACCUNIT_FILE_MISSING, "SaveAppFileOrRaise", "File could not be saved in usys_AppFiles: " & FilePath
   End If

End Sub

Private Function FolderPathWithBackslash(ByVal FolderPath As String) As String

   If Len(FolderPath) > 0 And VBA.Right$(FolderPath, 1) <> "\" Then
      FolderPath = FolderPath & "\"
   End If
   FolderPathWithBackslash = FolderPath

End Function

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
                                 Optional ByVal SetFocusToImmediateWindowBeforeTestStart As Boolean = True, _
                                 Optional ByVal JUnitXmlFilePath As String = vbNullString) As Boolean

   Dim Success As Boolean

#If AccUnitEarlyBinding Then
   Dim TestSummary As AccUnit.ITestSummary
#Else
   Dim TestSummary As Object
#End If

   Set VBE.ActiveVBProject = CurrentVbProject

   AddAccUnitTlbReference
   InsertFactoryModule
   ImportTestClasses

   If SetFocusToImmediateWindowBeforeTestStart Then
      SetFocusToImmediateWindow
   End If

   Set TestSummary = AccUnitLoaderFactoryCall.GetAccUnitFactory.TestSuite(TestReportOutputTo, JUnitXmlFilePath).AddFromVBProject.Run.Summary
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
