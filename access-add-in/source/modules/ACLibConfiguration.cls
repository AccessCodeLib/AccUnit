VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ACLibConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Konfigurationseinstellungen der CodeLib verwalten"
'---------------------------------------------------------------------------------------
' Class: _codelib.addins.shared.ACLibConfiguration
'---------------------------------------------------------------------------------------
'
' Manage AccessCodeLib configuration settings
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/ACLibConfiguration.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>data/SqlTools.bas</use>
'  <use>file/modFiles.bas</use>
'  <use>data/modSQL_Tools.bas</use>
'  <use>data/dao/modDAO_Tools.bas</use>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_DaoSqlTools As SqlTools

Private Const EXTENSION_KEY As String = "ACLibConfiguration"

#Const ADODB_EARLYBINDING = 0
'ADODB wird hier über Late binding eingesetzt, da es nur zum Erstellen der Tabelle genutzt wird


' Grundeinstellungen
Private Const ACLIB_CONFIG_ROOTFOLDERNAME As String = "AccessCodeLib"
Private Const ACLIB_CONFIG_ROOTFOLDERNAME_DEPR As String = "AccessCodeLibrary"
Private Const ACLIB_CONFIG_DATABASENAME As String = "ACLib_Config"
Private Const ACLIB_CONFIG_TABLEDEFNAME As String = "ACLib_ConfigTable"

Private Const SQL_CONFIG_TABLE_FIELD_PROPNAME As String = "PropName"
Private Const SQL_CONFIG_TABLE_FIELD_PROPVALUE As String = "PropValue"
Private Const SQL_SELECT_PROPERTYVALUE As String = _
      "select " & SQL_CONFIG_TABLE_FIELD_PROPNAME & ", " & SQL_CONFIG_TABLE_FIELD_PROPVALUE & _
      " From " & ACLIB_CONFIG_TABLEDEFNAME & " where " & SQL_CONFIG_TABLE_FIELD_PROPNAME & " = [?]"

Private Const SQL_CREATETABLE_CONFIGTABLE As String = _
            "CREATE TABLE " & ACLIB_CONFIG_TABLEDEFNAME & _
            "([PropName] varchar(255) WITH COMPRESSION NOT NULL," & _
            " [PropValue] varchar(255) WITH COMPRESSION," & _
            " [PropRemarks] text WITH COMPRESSION," & _
            " CONSTRAINT PK_" & ACLIB_CONFIG_TABLEDEFNAME & " PRIMARY KEY ([PropName]))"

Private Const PROPNAME_LOCALREPOSITORYROOT As String = "LocalRepositoryRoot"
Private Const PROPNAME_PRIVATEREPOSITORYROOT As String = "PrivateRepositoryRoot"
Private Const PROPNAME_IMPORTTESTDEFAULTVALUE As String = "ImportTestDefaultValue"
Private Const PROPNAME_GITHUBAUTHPERSONALACCESSTOKEN As String = "GitHubAuthPersonalAccessToken"

' Hilfsvariablen
Private m_LocalRepositoryPath As String   ' allg. Repository-Verzeichnis
Private m_PrivateRepositoryPath As String ' privates Verzeichnis (nicht in CodeLib enthalten)
Private m_ImportTestDefaultValue As Long
Private m_GitHubAuthPersonalAccessToken As String
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
' Ergänzungen für Ereiterung: ACLibConfiguration
'---------------------------------------------------------------------------------------

Private Sub GetExtensionPropertyLookup(ByVal PropertyName As String, ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

   ResumeMode = AppResumeMode_Completed

   Select Case PropertyName
      Case PROPNAME_LOCALREPOSITORYROOT
         ResumeMessage = LocalRepositoryPath

      Case PROPNAME_PRIVATEREPOSITORYROOT
         ResumeMessage = PrivateRepositoryPath

      Case Else 'Property wurde nicht erkannt
         ResumeMode = AppResumeMode_Error

   End Select

End Sub

Public Property Get LocalRepositoryPath() As String

   If Len(m_LocalRepositoryPath) = 0 Then
      m_LocalRepositoryPath = GetACLibGlobalProperty(PROPNAME_LOCALREPOSITORYROOT)
      If Len(m_LocalRepositoryPath) > 0 Then
         If Not DirExists(m_LocalRepositoryPath) Then
            Err.Raise vbObjectError, "ACLibConfiguration.LocalRepositoryPath", "Das Verzeichnis '" & m_LocalRepositoryPath & "' ist nicht vorhanden!"
            m_LocalRepositoryPath = vbNullString
         End If
         If Right$(m_LocalRepositoryPath, 1) <> "\" Then
            m_LocalRepositoryPath = m_LocalRepositoryPath & "\"
            SetACLibGlobalProperty PROPNAME_LOCALREPOSITORYROOT, m_LocalRepositoryPath
         End If
      End If
   End If

   LocalRepositoryPath = m_LocalRepositoryPath

End Property

Public Property Let LocalRepositoryPath(ByVal NewPath As String)

   If Len(NewPath) > 0 Then
      If Right$(NewPath, 1) <> "\" Then
         NewPath = NewPath & "\"
      End If
   End If
   m_LocalRepositoryPath = NewPath
   SetACLibGlobalProperty PROPNAME_LOCALREPOSITORYROOT, m_LocalRepositoryPath

End Property


Public Property Get PrivateRepositoryPath() As String

   If Len(m_PrivateRepositoryPath) = 0 Then
      m_PrivateRepositoryPath = GetACLibGlobalProperty(PROPNAME_PRIVATEREPOSITORYROOT)
      If Len(m_PrivateRepositoryPath) > 0 Then
         If Not DirExists(m_PrivateRepositoryPath) Then
            Err.Raise vbObjectError, "ACLibConfiguration.PrivateRepositoryPath", "Das Verzeichnis '" & m_PrivateRepositoryPath & "' ist nicht vorhanden!"
            m_PrivateRepositoryPath = vbNullString
         End If
         If Right$(m_PrivateRepositoryPath, 1) <> "\" Then
            m_PrivateRepositoryPath = m_PrivateRepositoryPath & "\"
            SetACLibGlobalProperty PROPNAME_PRIVATEREPOSITORYROOT, m_PrivateRepositoryPath
         End If
      End If
   End If

   PrivateRepositoryPath = m_PrivateRepositoryPath

End Property

Public Property Let PrivateRepositoryPath(ByVal NewPath As String)

   If Len(NewPath) > 0 Then
      If Right$(NewPath, 1) <> "\" Then
         NewPath = NewPath & "\"
      End If
   End If
   m_PrivateRepositoryPath = NewPath
   SetACLibGlobalProperty PROPNAME_PRIVATEREPOSITORYROOT, m_PrivateRepositoryPath

End Property

Public Property Get ImportTestsDefaultValue() As Boolean
'm_ImportTestDefaultValue: 0 = noch nicht abgefragt
'                          1 = false
'                          2 = true

   If m_ImportTestDefaultValue = 0 Then
      m_ImportTestDefaultValue = Val(GetACLibGlobalProperty(PROPNAME_IMPORTTESTDEFAULTVALUE)) + 1
   End If
   ImportTestsDefaultValue = (m_ImportTestDefaultValue = 2)

End Property

Public Property Let ImportTestsDefaultValue(ByVal NewValue As Boolean)

   m_ImportTestDefaultValue = Abs(NewValue) + 1
   SetACLibGlobalProperty PROPNAME_IMPORTTESTDEFAULTVALUE, Abs(NewValue)

End Property

Public Property Get GitHubAuthPersonalAccessToken() As String
'm_GitHubAuthPersonalAccessToken: vbnullstring = noch nicht abgefragt

   If StrPtr(m_GitHubAuthPersonalAccessToken) = 0 Then
      m_GitHubAuthPersonalAccessToken = GetACLibGlobalProperty(PROPNAME_GITHUBAUTHPERSONALACCESSTOKEN) & ""
   End If
   GitHubAuthPersonalAccessToken = m_GitHubAuthPersonalAccessToken

End Property

Public Property Let GitHubAuthPersonalAccessToken(ByVal NewValue As String)

   m_GitHubAuthPersonalAccessToken = NewValue
   SetACLibGlobalProperty PROPNAME_GITHUBAUTHPERSONALACCESSTOKEN, m_GitHubAuthPersonalAccessToken

End Property

Friend Function GetACLibGlobalProperty(ByRef PropertyName As String) As String

   Dim rst As DAO.Recordset
   Dim SelectSql As String

   SelectSql = Replace(SQL_SELECT_PROPERTYVALUE, "[?]", DaoSqlTool.TextToSqlText(PropertyName))
   Set rst = ACLibPropertyDb.OpenRecordset(SelectSql)
   If Not rst.EOF Then
      GetACLibGlobalProperty = Nz(rst.Fields(SQL_CONFIG_TABLE_FIELD_PROPVALUE), vbNullString)
   Else
      GetACLibGlobalProperty = vbNullString
   End If
   rst.Close

End Function

Friend Function SetACLibGlobalProperty(ByRef PropertyName As String, ByRef NewValue As String) As String

   Dim rst As DAO.Recordset
   Dim SelectSql As String

   SelectSql = Replace(SQL_SELECT_PROPERTYVALUE, "[?]", DaoSqlTool.TextToSqlText(PropertyName))
   Set rst = ACLibPropertyDb.OpenRecordset(SelectSql)
   If rst.EOF Then
      rst.AddNew
      rst.Fields(SQL_CONFIG_TABLE_FIELD_PROPNAME).Value = PropertyName
   Else
      rst.Edit
   End If
   rst.Fields(SQL_CONFIG_TABLE_FIELD_PROPVALUE).Value = NewValue
   rst.Update
   rst.Close

End Function

Private Property Get ACLibPropertyDb() As DAO.Database

   If m_ACLibPropertyDb Is Nothing Then
      If CheckConfigTableDef Then
         Set m_ACLibPropertyDb = CodeDb
      End If
   End If
   Set ACLibPropertyDb = m_ACLibPropertyDb

End Property

#If ADODB_EARLYBINDING Then
Private Function CreateConfigTable(ByRef TargetConnection As ADODB.Connection) As Boolean
#Else
Private Function CreateConfigTable(ByRef TargetConnection As Object) As Boolean
#End If

   TargetConnection.Execute SQL_CREATETABLE_CONFIGTABLE
   CreateConfigTable = True

End Function


Private Function CheckConfigTableDef() As Boolean

   Dim db As DAO.Database
   Dim tdf As DAO.TableDef
   Dim ConfigDataPath As String

   Set db = CodeDb

   If Not TableDefExists(ACLIB_CONFIG_TABLEDEFNAME, db) Then

      Set tdf = db.CreateTableDef(ACLIB_CONFIG_TABLEDEFNAME)
      tdf.Connect = ";Database=" & ACLibConfigDatabaseFile
      tdf.SourceTableName = ACLIB_CONFIG_TABLEDEFNAME
      db.TableDefs.Append tdf

   Else

      ConfigDataPath = Mid$(db.TableDefs(ACLIB_CONFIG_TABLEDEFNAME).Connect, Len(";Database=") + 1)
      If ConfigDataPath <> ACLibConfigDatabaseFile Then
         With db.TableDefs(ACLIB_CONFIG_TABLEDEFNAME)
            .Connect = ";Database=" & ACLibConfigDatabaseFile
            .RefreshLink
         End With
      End If

   End If

   Set db = Nothing

   CheckConfigTableDef = True

End Function

Public Property Get ACLibConfigDirectory() As String

   Dim strPath As String

   strPath = Environ("Appdata") & "\" & ACLIB_CONFIG_ROOTFOLDERNAME & "\"
   If Len(Dir$(strPath, vbDirectory)) = 0 Then
      MkDir strPath
   End If

   ACLibConfigDirectory = strPath

End Property

Private Property Get ACLibConfigDirectoryDepr() As String

   Dim strPath As String

   strPath = Environ("Appdata") & "\" & ACLIB_CONFIG_ROOTFOLDERNAME_DEPR & "\"

   ACLibConfigDirectoryDepr = strPath

End Property

Private Property Get ACLibConfigDatabaseFile() As String

   Dim db As DAO.Database
   Dim strDbFileExt As String
   Dim strDbFile As String
   Dim strDbFileDepr As String
   Dim bolCreateConfigTable As Boolean

#If ADODB_EARLYBINDING = 1 Then
   Dim cnn As ADODB.Connection
#Else
   Dim cnn As Object
#End If

   strDbFileExt = CodeDb.Name
   strDbFileExt = Mid$(strDbFileExt, InStrRev(strDbFileExt, "."))
   If Left$(strDbFileExt, 5) = ".accd" Then
      strDbFileExt = ".accdu"
   Else
      strDbFileExt = ".mdt"
   End If
   strDbFile = ACLibConfigDirectory & ACLIB_CONFIG_DATABASENAME & strDbFileExt

   ' Try transfer file from deprecated folder path:
   If Len(Dir$(strDbFile)) = 0 Then
      strDbFileDepr = ACLibConfigDirectoryDepr & ACLIB_CONFIG_DATABASENAME & strDbFileExt
      If Len(Dir$(strDbFileDepr)) > 0 Then
         FileCopy strDbFileDepr, strDbFile
      End If
   End If

   If Len(Dir$(strDbFile)) = 0 Then

      'Datenbank anlegen
      If CodeDb.Version = "4.0" Then
         Set db = DBEngine.CreateDatabase(strDbFile, dbLangGeneral, dbVersion40)
      Else
         Set db = DBEngine.CreateDatabase(strDbFile, dbLangGeneral)
      End If
      db.Close

      bolCreateConfigTable = True

   Else 'Prüfen ob Config-Tabelle vorhanden ist

      Set db = DBEngine.OpenDatabase(strDbFile)
      bolCreateConfigTable = Not TableDefExists(ACLIB_CONFIG_TABLEDEFNAME, db)
      db.Close

   End If

   If bolCreateConfigTable Then
      'Tabelle erstellen
#If ADODB_EARLYBINDING = 1 Then
      Set cnn = New ADODB.Connection
#Else
      Set cnn = CreateObject("ADODB.Connection")
#End If
      cnn.ConnectionString = Replace(CodeProject.Connection.ConnectionString, CodeDb.Name, strDbFile)
      cnn.Open
      CreateConfigTable cnn
      cnn.Close
      Set cnn = Nothing
   End If

   ACLibConfigDatabaseFile = strDbFile

End Property

Private Property Get DaoSqlTool()
   If m_DaoSqlTools Is Nothing Then
      Set m_DaoSqlTools = SqlTools.Clone("\#yyyy-mm-dd\#", "True", "*")
   End If
   Set DaoSqlTool = m_DaoSqlTools
End Property

Private Sub Class_Terminate()
   Set m_DaoSqlTools = Nothing
End Sub
