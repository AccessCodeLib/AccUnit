VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccUnitConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: AccUnitConfiguration
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Konfigurationseinstellungen verwalten
' </summary>
' <remarks>
' </remarks>
'\ingroup addins_shared
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/AccUnitConfiguration.cls</file>
'  <use>base/ApplicationHandler.cls</use>
'  <use>data/SqlTools.bas</use>
'  <use>file/FileTools.bas</use>
'  <use>data/dao/modDAO_Tools.bas</use>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private m_DaoSqlTools As SqlTools

Private Const EXTENSION_KEY As String = "AccUnitConfiguration"

#Const ADODB_EARLYBINDING = 0
'ADODB wird hier über Late binding eingesetzt, da es nur zum Erstellen der Tabelle genutzt wird

Private m_ACLibConfig As ACLibConfiguration

' Base config
Private Const PROPNAME_ACCUNITDLLPATH As String = "AccUnitDllPath"

' Hilfsvariablen
Private m_AccUnitDllPath As String   ' allg. Repository-Verzeichnis
Private m_PrivateRepositoryPath As String ' privates Verzeichnis (nicht in CodeLib enthalten)
Private m_ImportTestDefaultValue As Long
Private m_ACLibPropertyDb As DAO.Database

'---------------------------------------------------------------------------------------
' Standard-Initialisierung von Erweiterungen
'---------------------------------------------------------------------------------------

Private WithEvents m_ApplicationHandler As ApplicationHandler
Attribute m_ApplicationHandler.VB_VarHelpID = -1

Public Property Set ApplicationHandlerRef(ByRef ObjRef As ApplicationHandler)
   Set m_ApplicationHandler = ObjRef
End Property

Public Property Get ExtensionKey() As String
   ExtensionKey = EXTENSION_KEY
End Property

'---------------------------------------------------------------------------------------
' Standard-Ereignisbehandlung von Erweiterungen
'---------------------------------------------------------------------------------------

' CheckExtension
Private Sub m_ApplicationHandler_CheckExtension(ByVal ExtensionKeyToCheck As String, ByRef Exists As Boolean)
   If ExtensionKeyToCheck = EXTENSION_KEY Then Exists = True
End Sub

' ExtensionLookup
Private Sub m_ApplicationHandler_ExtensionLookup(ByVal ExtensionKeyToCheck As String, ByRef ExtensionReference As Object)
   If ExtensionKeyToCheck = EXTENSION_KEY Then
      Set ExtensionReference = Me
   End If
End Sub

'ExtensionPropertyLookup
Private Sub m_ApplicationHandler_ExtensionPropertyLookup( _
                     ByVal ExtensionKeyToCheck As String, ByVal PropertyName As String, _
                     ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
   If ExtensionKeyToCheck = EXTENSION_KEY Then
      GetExtensionPropertyLookup PropertyName, ResumeMode, ResumeMessage
   End If
End Sub

' AfterDispose
Private Sub m_ApplicationHandler_AfterDispose(ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
'=> Referenz in m_ApplicationHandler auf Nothing setzen
   Set m_ApplicationHandler = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Ergänzungen für Erweiterung: AccUnitConfiguration
'---------------------------------------------------------------------------------------


Public Property Get ACLibConfig() As ACLibConfiguration
   If m_ACLibConfig Is Nothing Then
      Set m_ACLibConfig = New ACLibConfiguration
   End If
   Set ACLibConfig = m_ACLibConfig
End Property

Private Sub GetExtensionPropertyLookup(ByVal PropertyName As String, ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

   ResumeMode = AppResumeMode_Completed

   Select Case PropertyName
      Case PROPNAME_ACCUNITDLLPATH
         ResumeMessage = AccUnitDllPath

      Case Else 'Property wurde nicht erkannt
         ResumeMode = AppResumeMode_Error

   End Select

End Sub

Public Property Get AccUnitDllPathPropertyName() As String
   AccUnitDllPathPropertyName = PROPNAME_ACCUNITDLLPATH
End Property

Public Property Get AccUnitDllPath() As String

   If Len(m_AccUnitDllPath) = 0 Then
      m_AccUnitDllPath = ACLibConfig.GetACLibGlobalProperty(PROPNAME_ACCUNITDLLPATH)
      If Len(m_AccUnitDllPath) > 0 Then
         If Not DirExists(m_AccUnitDllPath) Then
            Err.Raise vbObjectError, "AccUnitConfiguration.AccUnitDllPath", "Das Verzeichnis '" & m_AccUnitDllPath & "' ist nicht vorhanden!"
            m_AccUnitDllPath = vbNullString
         End If
         If VBA.Right$(m_AccUnitDllPath, 1) <> "\" Then
            m_AccUnitDllPath = m_AccUnitDllPath & "\"
            ACLibConfig.SetACLibGlobalProperty PROPNAME_ACCUNITDLLPATH, m_AccUnitDllPath
         End If
      End If
   End If

   If Len(m_AccUnitDllPath) = 0 Then
      m_AccUnitDllPath = DefaultAccUnitDllPath
   End If

   AccUnitDllPath = m_AccUnitDllPath

End Property

Public Property Let AccUnitDllPath(ByVal NewPath As String)

   If Len(NewPath) > 0 Then
      If Right$(NewPath, 1) <> "\" Then
         NewPath = NewPath & "\"
      End If
   End If
   m_AccUnitDllPath = NewPath
   ACLibConfig.SetACLibGlobalProperty PROPNAME_ACCUNITDLLPATH, m_AccUnitDllPath

End Property

Private Property Get DefaultAccUnitDllPath() As String
   DefaultAccUnitDllPath = DefaultAccUnitLibFolder & "\"
End Property


