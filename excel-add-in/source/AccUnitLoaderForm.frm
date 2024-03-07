VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AccUnitLoaderForm 
   Caption         =   "ACLib - AccUnit Loader"
   ClientHeight    =   4473
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   9373
   OleObjectBlob   =   "AccUnitLoaderForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "AccUnitLoaderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private m_RemoveInfoTextMaxTimer As Double

Private m_OpenMenuMouse_X As Long
Private m_OpenMenuMouse_Y As Long

Private Sub ShowErrorHandlerInfo(ByVal ProcName As String)
   m_RemoveInfoTextMaxTimer = 0
   Me.labInfo.Caption = "Error " & Err.Number & " (" & Err.Description & ") in procedure " & ProcName
End Sub

Private Sub ShowSuccessInfo(ByVal InfoText As String)
   Me.labInfo.Caption = InfoText
   m_RemoveInfoTextMaxTimer = Timer + ShowSuccessInfoTimerInterval / 1000
End Sub

Private Sub cmdOpenMenu_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
   m_OpenMenuMouse_X = x
   m_OpenMenuMouse_Y = Y
End Sub

Private Sub UserForm_Initialize()

   CheckAccUnitTypeLibFile CodeVBProject

   With CurrentApplication
      Me.Caption = .ApplicationTitle & " (Version " & .Version & ")"
   End With
   
'   LoadIconFromAppFiles
   
   With CurrentAccUnitConfiguration
On Error GoTo ErrMissingPath
      Me.txtAccUnitDllPath.Value = .AccUnitDllPath
On Error GoTo 0
   End With
   
   SetEnableMode
   
   Exit Sub
   
ErrMissingPath:
   Resume Next

End Sub

Private Sub UserForm_Terminate()
   On Error Resume Next
   DisposeCurrentApplicationHandler
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
   If m_RemoveInfoTextMaxTimer > 0 Then
      If Timer >= m_RemoveInfoTextMaxTimer Then
         m_RemoveInfoTextMaxTimer = 0
         Me.labInfo.Caption = vbNullString
      End If
   End If
End Sub

Private Property Get CurrentAccUnitDllPath() As String
   CurrentAccUnitDllPath = Me.txtAccUnitDllPath.Value
End Property

Private Sub cmdInsertFactoryModule_Click()

On Error GoTo HandleErr

   InsertFactoryModule
   ShowSuccessInfo "Factory module has been updated"
   
ExitHere:
   Exit Sub

HandleErr:
   ShowErrorHandlerInfo "cmdInsertFactoryModule_Click"
   Resume ExitHere
   
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
      Set .AccessForm = Me
      .UserFormCaption = Me.Caption
      
      Set .MenuControl = Me.cmdOpenMenu
      
     ' .ControlSection = acDetail

#If DEVMODE = 1 Then
      .AddMenuItem -99, "", MF_SEPARATOR
      .AddMenuItem -1, "For AccUnit developers:", MF_STRING + MF_GRAYED
      .AddMenuItem 11, "Import AccUnit files from directory"
#End If

If ThisWorkbook.CustomDocumentProperties.Count = 10 Then
      .AddMenuItem -2, "", MF_SEPARATOR
      .AddMenuItem 21, "Export AccUnit files to directory"
      .AddMenuItem 22, "Remove AccUnit files from Add-In file"
End If

      .AddMenuItem -3, "", MF_SEPARATOR
      .AddMenuItem 31, "Remove test environment incl. test classes"
      .AddMenuItem 32, "Remove test environment (keep test classes)"

      .AddMenuItem -4, "", MF_SEPARATOR
      .AddMenuItem 41, "Export test classes"
      .AddMenuItem 42, "Import test classes"

   End With

   Select Case mnu.OpenMenu(m_OpenMenuMouse_X, m_OpenMenuMouse_Y)
      Case 11
         ImportAccUnitFiles
         SuccessMessage = "AccUnit files imported"
      Case 21
         ExportAccUnitFiles
         SuccessMessage = "AccUnit files exported"
      Case 22
         RemoveAccUnitFilesFromAddInStorage
         SuccessMessage = "AccUnit files removed from Add-In file"
      Case 31
         RemoveTestEnvironment True
         SetEnableMode
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

   If Len(SuccessMessage) > 0 Then
      ShowSuccessInfo SuccessMessage
   End If
   
   Set mnu = Nothing

ExitHere:
   Exit Function

HandleErr:
   ShowErrorHandlerInfo "ImportAccUnitFiles"
   Resume ExitHere

End Function

Private Sub cmdExportFilesToFolder_Click()
   
On Error GoTo HandleErr

   ExportAccUnitFiles
   ShowSuccessInfo "AccUnit files exported"
   
ExitHere:
   Exit Sub

HandleErr:
   ShowErrorHandlerInfo "cmdInsertFactoryModule_Click"
   Resume ExitHere
   
End Sub

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
   With Me.cmdInsertFactoryModule
      .Enabled = bolPathExists
      If bolPathExists Then
         .Enabled = (ThisWorkbook.CustomDocumentProperties.Count = 10)
      End If
   End With

End Sub

Private Sub cmdSetAccUnitTlbReferenz_Click()
   
On Error GoTo HandleErr

   AddAccUnitTlbReference
   ShowSuccessInfo "AccUnit.tlb reference inserted"
   
ExitHere:
   Exit Sub

HandleErr:
   ShowErrorHandlerInfo "cmdSetAccUnitTlbReferenz_Click"
   Resume ExitHere
   
End Sub

Private Sub cmdRemoveAccUnitTlbReferenz_Click()

On Error GoTo HandleErr

   RemoveAccUnitTlbReference
   ShowSuccessInfo "AccUnit.tlb reference removed"
   
ExitHere:
   Exit Sub

HandleErr:
   ShowErrorHandlerInfo "cmdRemoveAccUnitTlbReferenz_Click"
   Resume ExitHere

End Sub



Private Sub cmdUserSettings_Click()
'   DoCmd.OpenForm "AccUnitUserSettings", acNormal, , , , acDialog
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

Private Sub txtAccUnitDllPath_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

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

'   Dim IconFilePath As String
'   Dim IconFileName As String
'
'   'Latebindung, damit ApplicationHandler_AppFile-Klasse nicht vorhanden sein muss
'   Dim AppFile As Object ' ... ApplicationHandler_AppFile
'
'   If Val(SysCmd(acSysCmdAccessVer)) <= 9 Then 'Abbruch, da Ac00 sonst abstürzt
'      Exit Sub
'   End If
'
'   Set AppFile = CurrentApplication.Extensions(EXTENSION_KEY_APPFILE)
'
'   'Textbox binden
'   If Not (AppFile Is Nothing) Then
'      IconFileName = ACLibIconFileName
'      IconFilePath = CurrentAccUnitConfiguration.ACLibConfig.ACLibConfigDirectory
'
'      If Len(ACLibIconFileName) = 0 Then 'nur Temp-Datei erzeugen
'         IconFileName = Me.Name & ".ico"
'         IconFilePath = TempPath
'      End If
'
'      IconFilePath = IconFilePath & IconFileName
'
'      If Len(Dir$(IconFilePath)) = 0 Then
'         If Not AppFile.CreateAppFile(APPFILE_PROPNAME_APPICON, IconFilePath) Then
'            Exit Sub
'         End If
'      End If
'
'      WinAPI.Image.SetFormIconFromFile Me, IconFilePath
'
'   End If
   
End Sub

