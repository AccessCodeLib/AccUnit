Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Form: AccUnitUserSettings
'---------------------------------------------------------------------------------------
'
' Wizard form to set AccUnit User Settings
'
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/AccUnitUserSettings.frm</file>
'  <description>Wizard form to set AccUnit User Settings</description>
'  <use>%AppFolder%/source/defGlobal_AccUnitLoader.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private m_UserSettings As AccUnit.IUserSettings

Private Sub cmdCancel_Click()
   DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdCommit_Click()
   SaveDataToSettings
   DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Load()
   InitSettings
End Sub

Private Sub InitSettings()

   Dim Configurator As AccUnit.Configurator

   With New AccUnitLoaderFactory
      Set Configurator = .Configurator
   End With

   Set m_UserSettings = Configurator.UserSettings

   Set Configurator = Nothing

   LoadDataFromSettings

End Sub

Private Sub LoadDataFromSettings()
   With m_UserSettings
      Me.txtImportExportFolder.Value = .ImportExportFolder
      Me.txtTemplateFolder.Value = .TemplateFolder
      Me.txtTestClassNameFormat.Value = .TestClassNameFormat
      Me.txtTestMethodTemplate.Value = Replace(.TestMethodTemplate, vbTab, TabSpaces)
   End With
End Sub

Private Sub SaveDataToSettings()
   With m_UserSettings
      .ImportExportFolder = Me.txtImportExportFolder.Value
      .TemplateFolder = Me.txtTemplateFolder.Value
      .TestClassNameFormat = Me.txtTestClassNameFormat.Value
      .TestMethodTemplate = Replace(Me.txtTestMethodTemplate.Value, TabSpaces, vbTab)
      .Save
   End With
End Sub

Private Property Get TabSpaces() As String
   TabSpaces = String(VBETabWidth, " ")
End Property

Private Property Get VBETabWidth() As Long

   Static TabWidth As Long

   Dim RegPath As String
   RegPath = "HKEY_CURRENT_USER\Software\Microsoft\VBA\" & Replace(VBE.Version, ".0", ".") & "\Common\TabWidth"

   If TabWidth = 0 Then
      With CreateObject("WScript.Shell")
         TabWidth = Val(.RegRead(RegPath))
      End With
   End If

   VBETabWidth = TabWidth

End Property
