VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AccUnitUserSettings 
   Caption         =   "ACLib - AccUnit: Settings"
   ClientHeight    =   5572
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   10620
   OleObjectBlob   =   "AccUnitUserSettings.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "AccUnitUserSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
   Unload Me
End Sub

Private Sub cmdCommit_Click()
   SaveDataToSettings
   Unload Me
End Sub

Private Sub UserForm_Initialize()
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
   RegPath = "HKEY_CURRENT_USER\Software\Microsoft\VBA\" & Replace(Application.VBE.Version, ".0", ".") & "\Common\TabWidth"

   If TabWidth = 0 Then
      With CreateObject("WScript.Shell")
         TabWidth = Val(.RegRead(RegPath))
      End With
   End If
   
   VBETabWidth = TabWidth
   
End Property
