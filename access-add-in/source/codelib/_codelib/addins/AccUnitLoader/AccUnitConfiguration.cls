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
'  <file>_codelib/addins/AccUnitLoader/AccUnitConfiguration.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>data/SqlTools.bas</use>
'  <use>file/FileTools.bas</use>
'  <use>data/dao/modDAO_Tools.bas</use>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_DaoSqlTools As SqlTools

Private Const EXTENSION_KEY As String = "AccUnitConfiguration"

#Const ADODB_EARLYBINDING = 0
'ADODB wird hier über Late binding eingesetzt, da es nur zum Erstellen der Tabelle genutzt wird


' Grundeinstellungen
Private Const ACLIB_CONFIG_ROOTFOLDERNAME As String = "AccessCodeLibrary"
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
' Ergänzungen für Ereiterung: AccUnitConfiguration
'---------------------------------------------------------------------------------------

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
      m_AccUnitDllPath = GetACLibGlobalProperty(PROPNAME_ACCUNITDLLPATH)
      If Len(m_AccUnitDllPath) > 0 Then
         If Not DirExists(m_AccUnitDllPath) Then
            Err.Raise vbObjectError, "AccUnitConfiguration.AccUnitDllPath", "Das Verzeichnis '" & m_AccUnitDllPath & "' ist nicht vorhanden!"
            m_AccUnitDllPath = vbNullString
         End If
         If VBA.Right$(m_AccUnitDllPath, 1) <> "\" Then
            m_AccUnitDllPath = m_AccUnitDllPath & "\"
            SetACLibGlobalProperty PROPNAME_ACCUNITDLLPATH, m_AccUnitDllPath
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
   SetACLibGlobalProperty PROPNAME_ACCUNITDLLPATH, m_AccUnitDllPath
  
End Property

Private Property Get DefaultAccUnitDllPath() As String
   DefaultAccUnitDllPath = DefaultAccUnitLibFolder & "\"
End Property

Private Function GetACLibGlobalProperty(ByRef PropertyName As String) As String
   
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

Private Function SetACLibGlobalProperty(ByRef PropertyName As String, ByRef NewValue As String) As String
   
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

   Set db = CodeDb
   
   If Not TableDefExists(ACLIB_CONFIG_TABLEDEFNAME, db) Then
   
      Set tdf = db.CreateTableDef(ACLIB_CONFIG_TABLEDEFNAME)
      tdf.Connect = ";Database=" & ACLibConfigDatabaseFile
      tdf.SourceTableName = ACLIB_CONFIG_TABLEDEFNAME
      db.TableDefs.Append tdf
      
   ElseIf Len(VBA.Dir$(VBA.Mid$(db.TableDefs(ACLIB_CONFIG_TABLEDEFNAME).Connect, VBA.Len(";Database=") + 1))) = 0 Then
   
      With db.TableDefs(ACLIB_CONFIG_TABLEDEFNAME)
         .Connect = ";Database=" & ACLibConfigDatabaseFile
         .RefreshLink
      End With
      
   End If
   
   Set db = Nothing
   
   CheckConfigTableDef = True

End Function

Public Property Get ACLibConfigDirectory() As String
   
   Dim strPath As String

   strPath = VBA.Environ("Appdata") & "\" & ACLIB_CONFIG_ROOTFOLDERNAME & "\"
   If Len(VBA.Dir$(strPath, vbDirectory)) = 0 Then
      VBA.MkDir strPath
   End If
   
   ACLibConfigDirectory = strPath

End Property

Private Property Get ACLibConfigDatabaseFile() As String
   
   Dim db As DAO.Database
   Dim strDbFile As String
   Dim bolCreateConfigTable As Boolean
   
#If ADODB_EARLYBINDING = 1 Then
   Dim cnn As ADODB.Connection
#Else
   Dim cnn As Object
#End If

   strDbFile = CodeDb.Name
   strDbFile = VBA.Mid$(strDbFile, VBA.InStrRev(strDbFile, "."))
   If VBA.Left$(strDbFile, 5) = ".accd" Then
      strDbFile = ".accdu"
   Else
      strDbFile = ".mdt"
   End If
   strDbFile = ACLibConfigDirectory & ACLIB_CONFIG_DATABASENAME & strDbFile
   
   If Len(VBA.Dir$(strDbFile)) = 0 Then

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
      cnn.ConnectionString = VBA.Replace(CodeProject.Connection.ConnectionString, CodeDb.Name, strDbFile)
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
