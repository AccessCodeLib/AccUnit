Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Form: AccUnitLoaderForm
'---------------------------------------------------------------------------------------
'
' Wizard Formular to config AccUnit Loader
'
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/AccUnitLoaderForm.frm</file>
'  <description>Wizard Formular to config AccUnit Loader</description>
'  <use>%AppFolder%/source/defGlobal_AccUnitLoader.bas</use>
'  <use>file/FileTools.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

' verwendete Erweiterungen
Private Const EXTENSION_KEY_APPFILE As String = "AppFile"
Private Const APPFILE_PROPNAME_APPICON As String = "AppIcon"

Private Const ShowSuccessInfoTimerInterval As Long = 4000

Private RunLoadAfterAddInInstalledTimer As Boolean

Private Sub ShowErrorHandlerInfo(ByVal ProcName As String)
   Me.labInfo.Caption = "Error " & Err.Number & " (" & Err.Description & ") in procedure " & ProcName
End Sub

Private Sub cmdClose_Click()
   DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdExportAccUnitFiles_Click()

On Error GoTo HandleErr

   ExportAccUnitFiles
   Me.labInfo.Caption = "AccUnit files exported"
   Me.TimerInterval = ShowSuccessInfoTimerInterval

ExitHere:
   Exit Sub

HandleErr:
   ShowErrorHandlerInfo "cmdExportAccUnitFiles_Click"
   Resume ExitHere

End Sub

Private Sub cmdInsertFactoryModule_Click()

On Error GoTo HandleErr

   InsertFactoryModule
   Me.labInfo.Caption = "Factory module has been updated"
   Me.TimerInterval = ShowSuccessInfoTimerInterval

ExitHere:
   Exit Sub

HandleErr:
   ShowErrorHandlerInfo "cmdInsertFactoryModule_Click"
   Resume ExitHere

End Sub

Private Sub cmdInstallVbeAddIn_Click()

On Error GoTo HandleErr

   If Me.cmdInstallVbeAddIn.Tag <> "Installed" Then
      OpenVbeInstallShortcutMenu
      Exit Sub
   End If

   With New VbeAddInInstaller
      .UnRegisterAccUnitVbeAddIn
      VBE.Addins.Update
      SetInstallVbeAddInCaption False
      Me.cmdLoadVbeAddIn.Enabled = False
      SetLoadVbeAddInCaption False

   End With

   Me.labInfo.Caption = "AccUnit VBE Add-in has been unregistered"
   Me.TimerInterval = ShowSuccessInfoTimerInterval

ExitHere:
   Exit Sub

HandleErr:
   ShowErrorHandlerInfo "cmdInstallVbeAddIn_Click"
   Resume ExitHere

End Sub

Private Function OpenVbeInstallShortcutMenu() As Long

   Dim mnu As WinApiShortcutMenu
   Dim SuccessMessage As String
   Dim Installed As Boolean

On Error GoTo HandleErr

   Set mnu = New WinApiShortcutMenu

   With mnu
      Set .MenuControl = Me.cmdInstallVbeAddIn
      Set .AccessForm = Me
      .ControlSection = acDetail

      .AddMenuItem 1, "Install Add-In (Load on Startup)"
      .AddMenuItem 2, "Install Add-In (don't load)"

   End With

   With New VbeAddInInstaller
      Select Case mnu.OpenMenu
         Case 1
            .RegisterAccUnitVbeAddIn CurrentAccUnitConfiguration.AccUnitDllPath, 3
            Installed = True
            RunLoadAfterAddInInstalledTimer = True
         Case 2
            .RegisterAccUnitVbeAddIn CurrentAccUnitConfiguration.AccUnitDllPath, 0
            Installed = True
         Case Else
            '
      End Select
   End With

   SetInstallVbeAddInCaption Installed
   If Installed Then
      SuccessMessage = "AccUnit VBE Add-in installed"
      Me.cmdLoadVbeAddIn.Enabled = True
   Else
      SuccessMessage = "Canceled"
   End If

   Me.labInfo.Caption = SuccessMessage
   If RunLoadAfterAddInInstalledTimer Then
      Me.TimerInterval = 250
   Else
      Me.TimerInterval = ShowSuccessInfoTimerInterval
   End If
   Set mnu = Nothing

ExitHere:
   Exit Function

HandleErr:
   ShowErrorHandlerInfo "OpenVbeInstallShortcutMenu"
   Resume ExitHere

End Function

Private Sub LoadAddInAfterInstall()
   VBE.Addins.Update
   With New VbeAddInInstaller
      .LoadAccUnitAddIn True
      SetLoadVbeAddInCaption .AddInIsLoaded
   End With
   Me.TimerInterval = ShowSuccessInfoTimerInterval

   Me.labInfo.Caption = "AccUnit VBE Add-in has been unregistered"
   Me.TimerInterval = ShowSuccessInfoTimerInterval

End Sub

Private Sub cmdLoadVbeAddIn_Click()

   Dim IsLoaded As Boolean

On Error GoTo HandleErr

   VBE.Addins.Update

   With New VbeAddInInstaller
      .LoadAccUnitAddIn Not .AddInIsLoaded
      IsLoaded = .AddInIsLoaded
      SetLoadVbeAddInCaption IsLoaded
   End With

   If IsLoaded Then
      Me.labInfo.Caption = "AccUnit VBE Add-in loaded"
   Else
      Me.labInfo.Caption = "AccUnit VBE Add-in unloaded"
   End If
   Me.TimerInterval = ShowSuccessInfoTimerInterval

ExitHere:
   Exit Sub

HandleErr:
   ShowErrorHandlerInfo "cmdLoadVbeAddIn_Click"
   Resume ExitHere

End Sub

Private Sub SetInstallVbeAddInCaption(ByVal IsInstalled As Boolean)

   Dim ButtonCaption As String
   Dim ButtonTag As String

   If IsInstalled Then
      ButtonCaption = "Uninstall AccUnit VBE Add-in"
      ButtonTag = "Installed"
   Else
      ButtonCaption = "Install AccUnit VBE Add-in"
      ButtonTag = vbNullString
   End If

   With Me.cmdInstallVbeAddIn
      .Caption = ButtonCaption
      .Tag = ButtonTag
   End With

End Sub

Private Sub SetLoadVbeAddInCaption(ByVal IsLoaded As Boolean)

   Dim ButtonCaption As String

   If IsLoaded Then
      ButtonCaption = "Unload AccUnit VBE Add-in"
   Else
      ButtonCaption = "Load AccUnit VBE Add-in"
   End If

   Me.cmdLoadVbeAddIn.Caption = ButtonCaption

End Sub

Private Sub cmdOpenMenu_Click()
   OpenImportFileShortcutMenu
End Sub

Private Function OpenImportFileShortcutMenu() As Long

   Dim mnu As WinApiShortcutMenu
   Dim SuccessMessage As String

On Error GoTo HandleErr

   Set mnu = New WinApiShortcutMenu

   With mnu
      Set .MenuControl = Me.cmdOpenMenu
      Set .AccessForm = Me
      .ControlSection = acDetail

#If DEVMODE = 1 Then
      .AddMenuItem -1, "For AccUnit developers:", MF_STRING + MF_GRAYED
      .AddMenuItem 11, "Import AccUnit files from directory"
      .AddMenuItem -2, "", MF_SEPARATOR
#End If

      .AddMenuItem 21, "Export AccUnit files to directory"

      .AddMenuItem -3, "", MF_SEPARATOR
      .AddMenuItem 31, "Remove test environment incl. test classes"
      .AddMenuItem 32, "Remove test environment (keep test classes)"

      .AddMenuItem -4, "", MF_SEPARATOR
      .AddMenuItem 41, "Export test classes"
      .AddMenuItem 42, "Import test classes"

   End With

   Select Case mnu.OpenMenu
      Case 11
         ImportAccUnitFiles
         SuccessMessage = "AccUnit files imported"
      Case 21
         ExportAccUnitFiles
         SuccessMessage = "AccUnit files exported"
      Case 31
         RemoveTestEnvironment True
         SuccessMessage = "Test environment end test classes removed"
      Case 32
         RemoveTestEnvironment False
         SuccessMessage = "Test environment removed"
      Case 41
         ExportTestClasses
         SuccessMessage = "Test classes exported"
      Case 42
         ImportTestClasses
         SuccessMessage = "Test classes imported"
      Case Else
         '
   End Select

   Me.labInfo.Caption = SuccessMessage
   Me.TimerInterval = ShowSuccessInfoTimerInterval

   Set mnu = Nothing

ExitHere:
   Exit Function

HandleErr:
   ShowErrorHandlerInfo "ImportAccUnitFiles"
   Resume ExitHere

End Function

Private Sub cmdSelectAccUnitDllPath_Click()

   Dim SelectedAccUnitDllPath As String

On Error GoTo HandleErr

   SelectedAccUnitDllPath = SelectFolder(Nz(Me.txtAccUnitDllPath, vbNullString), "Lokalen Repository-Ordner auswählen", , False, 1)

   If Len(SelectedAccUnitDllPath) > 0 Then
      If Right$(SelectedAccUnitDllPath, 1) = "\" Then
         SelectedAccUnitDllPath = Left$(SelectedAccUnitDllPath, Len(SelectedAccUnitDllPath) - 1)
      End If

      SetAccUnitDllPath SelectedAccUnitDllPath

   End If

ExitHere:
   Exit Sub

HandleErr:
   ShowErrorHandlerInfo "cmdSelectAccUnitDllPath_Click"
   Resume ExitHere

End Sub

Private Sub SetEnableMode()

   Dim bolPathExists As Boolean
   bolPathExists = Len(Me.txtAccUnitDllPath.Value & vbNullString) > 0

   Me.cmdSetAccUnitTlbReferenz.Enabled = bolPathExists
   Me.cmdInsertFactoryModule.Enabled = bolPathExists

End Sub

Private Sub cmdSetAccUnitTlbReferenz_Click()

On Error GoTo HandleErr

   AddAccUnitTlbReference
   Me.labInfo.Caption = ACCUNIT_TYPELIB_FILE & " reference inserted"
   Me.TimerInterval = ShowSuccessInfoTimerInterval

ExitHere:
   Exit Sub

HandleErr:
   ShowErrorHandlerInfo "cmdSetAccUnitTlbReferenz_Click"
   Resume ExitHere

End Sub

Private Sub cmdRemoveAccUnitTlbReferenz_Click()

On Error GoTo HandleErr

   RemoveAccUnitTlbReference
   Me.labInfo.Caption = ACCUNIT_TYPELIB_FILE & " reference removed"
   Me.TimerInterval = ShowSuccessInfoTimerInterval

ExitHere:
   Exit Sub

HandleErr:
   ShowErrorHandlerInfo "cmdRemoveAccUnitTlbReferenz_Click"
   Resume ExitHere

End Sub

Private Sub cmdUserSettings_Click()
   DoCmd.OpenForm "AccUnitUserSettings", acNormal, , , , acDialog
End Sub

Private Sub Form_Load()

On Error GoTo ErrMissingPath

   With CurrentApplication
      Me.Caption = .ApplicationTitle & "  " & VBA.ChrW(&H25AA) & "  Version " & .Version
   End With

   LoadIconFromAppFiles

   With CurrentAccUnitConfiguration
      Me.txtAccUnitDllPath.Value = .AccUnitDllPath
   End With

   Dim VbeAddInIsInstalled As Boolean
   With New VbeAddInInstaller
      VbeAddInIsInstalled = .AddInIsInstalled
      SetInstallVbeAddInCaption VbeAddInIsInstalled
      Me.cmdLoadVbeAddIn.Enabled = VbeAddInIsInstalled
      SetLoadVbeAddInCaption .AddInIsLoaded
   End With

   SetEnableMode

   Exit Sub

ErrMissingPath:
   ShowErrorHandlerInfo Err.Description
   Resume Next

End Sub

Private Sub Form_Open(Cancel As Integer)

   Dim ReferenceFixed As Boolean
   Dim ReferenceFixedMessage As String

   modTypeLibCheck.CheckAccUnitTypeLibFile modVbProject.CodeVBProject, ReferenceFixed, ReferenceFixedMessage
   If VBA.Len(ReferenceFixedMessage) Then
      Me.labInfo.Caption = ReferenceFixedMessage
   End If

End Sub

Private Sub Form_Timer()

   If RunLoadAfterAddInInstalledTimer Then
      RunLoadAfterAddInInstalledTimer = False
      LoadAddInAfterInstall
      Exit Sub
   End If

   Me.TimerInterval = 0
   Me.labInfo.Caption = vbNullString
End Sub

Private Sub Form_Unload(ByRef Cancel As Integer)
On Error Resume Next
   DisposeCurrentApplicationHandler
End Sub

Private Sub txtAccUnitDllPath_AfterUpdate()
   SetAccUnitDllPath Me.txtAccUnitDllPath & vbNullString
End Sub

Private Sub SetAccUnitDllPath(ByRef NewRoot As String)

   CurrentAccUnitConfiguration.AccUnitDllPath = NewRoot

   'damit mögliche Modifikationen aus CurrentAccUnitConfiguration übernommen werden:
   Me.txtAccUnitDllPath.Value = CurrentAccUnitConfiguration.AccUnitDllPath

   SetEnableMode

End Sub

Private Sub txtAccUnitDllPath_BeforeUpdate(ByRef Cancel As Integer)

   Dim NewPath As String

   NewPath = Me.txtAccUnitDllPath & ""

   If Len(NewPath) > 0 Then
      If Not DirExists(NewPath) Then
         If MsgBox("Directory does not exist." & vbNewLine & "Create directory?", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
            FileTools.CreateDirectory NewPath
         Else
            Cancel = True
         End If
      End If
   End If

End Sub

Private Sub LoadIconFromAppFiles()

   Dim IconFilePath As String
   Dim IconFileName As String

   'Latebindung, damit ApplicationHandler_AppFile-Klasse nicht vorhanden sein muss
   Dim AppFile As Object ' ... ApplicationHandler_AppFile

   If Val(SysCmd(acSysCmdAccessVer)) <= 9 Then 'Abbruch, da Ac00 sonst abstürzt
      Exit Sub
   End If

   Set AppFile = CurrentApplication.Extensions(EXTENSION_KEY_APPFILE)

   'Textbox binden
   If Not (AppFile Is Nothing) Then
      IconFileName = ACLibIconFileName
      IconFilePath = CurrentAccUnitConfiguration.ACLibConfig.ACLibConfigDirectory

      If Len(ACLibIconFileName) = 0 Then 'nur Temp-Datei erzeugen
         IconFileName = Me.Name & ".ico"
         IconFilePath = TempPath
      End If

      IconFilePath = IconFilePath & IconFileName

      If Len(Dir$(IconFilePath)) = 0 Then
         If Not AppFile.CreateAppFile(APPFILE_PROPNAME_APPICON, IconFilePath) Then
            Exit Sub
         End If
      End If

      WinAPI.Image.SetFormIconFromFile Me, IconFilePath

   End If

End Sub
