VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccUnitConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Konfigurationseinstellungen der CodeLib verwalten"
'---------------------------------------------------------------------------------------
' Class: AccUnitConfiguration
'---------------------------------------------------------------------------------------
'
' AccUnit configuration
'
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Private m_DaoSqlTools As SqlTools

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

Public Property Get ACLibConfig() As ACLibConfiguration
   If m_ACLibConfig Is Nothing Then
      Set m_ACLibConfig = New ACLibConfiguration
   End If
   Set ACLibConfig = m_ACLibConfig
End Property

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
