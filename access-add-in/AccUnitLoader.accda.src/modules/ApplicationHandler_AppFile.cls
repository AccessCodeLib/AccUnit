VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler_AppFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Erweiterung für ApplicationHandler-Klasse: Anwendungsspezifische Dateien verwalten"
'---------------------------------------------------------------------------------------
' Class: base.ApplicationHandler_File
'---------------------------------------------------------------------------------------
'
' Extension for ApplicationHandler class: manage application-specific files
'
' Author:
'     Josef Poetzl
'
' Remarks:
'     Activation in the _config_Application module:
'     <code>ExtensionsCollectionReferenz.Add New ApplicationHandler_AppFile</code>
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/ApplicationHandler_AppFile.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>api/winapi/WinApiFileInfo.cls</use>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const EXTENSION_KEY As String = "AppFile"

Private Const EXTENSION_PROPERTY_APPFILETABLENAME As String = "AppFileTableName"
Private Const TABLE_APPFILES As String = "usys_AppFiles"
Private Const TABLE_FIELD_ID As String = "id"
Private Const TABLE_FIELD_VERSION As String = "version"
Private Const TABLE_FIELD_FILE As String = "file"


'---------------------------------------------------------------------------------------
' Standard initialization of extensions
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
' Standard event handling of extensions
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
'=> Set reference in m_ApplicationHandler to Nothing
   Set m_ApplicationHandler = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Additions for extension: ApplicationHandler_AppFile
'---------------------------------------------------------------------------------------
'

Private Sub GetExtensionPropertyLookup(ByVal PropertyName As String, _
                                       ByRef ResumeMode As ApplicationHandlerResumeModes, _
                                       ByRef ResumeMessage As Variant)

   ResumeMode = AppResumeMode_Completed

   Select Case PropertyName
      Case EXTENSION_PROPERTY_APPFILETABLENAME
         ResumeMessage = TABLE_APPFILES

      Case Else 'Property wurde nicht erkannt
         ResumeMode = AppResumeMode_Error

   End Select

End Sub

Public Property Get AppFileTableName() As String
   AppFileTableName = TABLE_APPFILES
End Property

'---------------------------------------------------------------------------------------
' Function: CreateAppFile
'---------------------------------------------------------------------------------------
'
' Extract file from usys_AppFiles
'
' Parameters:
'     FileID   - Record identifier
'     FileName - File name to be created
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function CreateAppFile(ByVal FileID As String, ByVal FileName As String, _
                     Optional ByVal ExtFilterFieldName As String, Optional ExtFilterValue As Variant) As Boolean

   Dim Binfile() As Byte
   Dim FieldSize As Long
   Dim fld As DAO.Field
   Dim SelectSql As String

   SelectSql = "select " & TABLE_FIELD_FILE & " from " & TABLE_APPFILES & " where " & TABLE_FIELD_ID & "='" & Replace(FileID, "'", "''") & "'"
   If Len(ExtFilterFieldName) > 0 Then
      SelectSql = SelectSql & " and " & ExtFilterFieldName & " = '" & Replace(ExtFilterValue, "'", "''") & "'"
   End If

   With CodeDb.OpenRecordset(SelectSql)
      If Not .EOF Then

         Set fld = .Fields(0)
         FieldSize = fld.FieldSize

         If FieldSize > 0 Then
            ReDim Binfile(FieldSize - 1)
            Binfile = fld.GetChunk(0, FieldSize)
            CreateFileFromByteArray FileName, Binfile
            CreateAppFile = True
         End If

      End If
      .Close
   End With

End Function

Private Sub CreateFileFromByteArray(ByVal FileName As String, ByRef Binfile() As Byte)

   Dim FileNo As Integer

   FileTools.CreateDirectoryIfMissing FileTools.GetDirFromFullFileName(FileName)

   FileNo = FreeFile
   Open FileName For Binary As #FileNo
   Put #FileNo, , Binfile()
   Close #FileNo

End Sub

'---------------------------------------------------------------------------------------
' Function: SaveAppFile
'---------------------------------------------------------------------------------------
'
' Datei in usys_AppFiles speichern
'
' Parameters:
'     FileID         - Identifier in table
'     FileName       - File name to be created
'     SaveVersion    - (optional) Read version from file (e. g. from dll file)
'     ExtFieldName   - (optional) append additional data to data field: Field name
'     ExtFieldValue  - (optional) append additional data to data field: Value
'
' Returns:
'     Boolean  - True: File was saved in table
'
'---------------------------------------------------------------------------------------
Public Function SaveAppFile(ByVal FileID As String, ByVal FileName As String, _
                   Optional ByVal SaveVersion As Boolean = False, _
                   Optional ByVal ExtFieldName As String, Optional ByVal ExtFieldValue As Variant, _
                   Optional ByVal ExtFilterFieldName As String, Optional ExtFilterValue As Variant) As Boolean

   Dim FileNo As Integer
   Dim Binfile() As Byte
   Dim ArraySize As Long
   Dim Version As String
   Dim FileInfo As WinApiFileInfo
   Dim SelectSql As String

On Error GoTo HandleErr

   FileNo = FreeFile
   Open FileName For Binary As #FileNo
   ArraySize = LOF(FileNo) - 1
   If ArraySize >= 0 Then
      ReDim Binfile(ArraySize)
      Get #FileNo, , Binfile()
   End If
   Close #FileNo

   If ArraySize < 0 Then
      SaveAppFile = False
      Exit Function
   End If

   SelectSql = "select * from " & TABLE_APPFILES & " where id='" & Replace(FileID, "'", "''") & "'"
   If Len(ExtFilterFieldName) > 0 Then
      SelectSql = SelectSql & " and " & ExtFilterFieldName & " = '" & Replace(ExtFilterValue, "'", "''") & "'"
   End If

   Dim rst As DAO.Recordset
   Set rst = CodeDb.OpenRecordset(SelectSql)
   If rst.EOF Then
      rst.AddNew
      rst.Fields("id") = FileID
      If Len(ExtFilterFieldName) > 0 Then
         rst.Fields(ExtFilterFieldName) = ExtFilterValue
      End If
   Else
      rst.Edit
   End If
   rst.Fields("file").AppendChunk Binfile
   If SaveVersion Then
      Set FileInfo = New WinApiFileInfo
      Version = FileInfo.GetFileVersion(FileName)
          Set FileInfo = Nothing
      If Len(Version) = 0 Then
         Version = VBA.Format$(Nz(FileDateTime(FileName), vbNullString), "yyyy.mm.dd")
      End If
      rst.Fields(TABLE_FIELD_VERSION) = Version
   End If

   If Len(ExtFieldName) > 0 Then
      rst.Fields(ExtFieldName).Value = ExtFieldValue
   End If

   rst.Update
   rst.Close
   Set rst = Nothing

   SaveAppFile = True

ExitHere:
   Exit Function

HandleErr:
   Dim TabCreateErrCnt As Long
   Dim ErrNumber As Long, ErrDescription As String

   ErrNumber = Err.Number
   ErrDescription = Err.Description

   If ErrNumber = 3078 And TabCreateErrCnt = 0 Then
      TabCreateErrCnt = 1
      Err.Clear
      If CreateAppFileTable Then
         Resume
      End If
   End If

   Err.Raise ErrNumber, "ExtAppFile.SaveAppFile", ErrDescription

End Function

Private Function CreateAppFileTable() As Boolean

   Dim CreateTableSql As String

   CreateTableSql = "create table " & TABLE_APPFILES & " ( " & TABLE_FIELD_ID & " varchar(20) primary key, " & _
                    TABLE_FIELD_VERSION & " varchar(10), " & TABLE_FIELD_FILE & " image)"
   CodeDb.Execute CreateTableSql, dbFailOnError

   CreateAppFileTable = (Err.Number = 0)

End Function


'---------------------------------------------------------------------------------------
' Event handling of m_ApplicationHandler
'

' AppFileBeforeCreateFile
Private Sub m_ApplicationHandler_AppFileBeforeCreateFile(ByVal FileID As String, ByVal FileName As String, _
                              ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

   Dim Completed As Boolean

   Completed = CreateAppFile(FileID, FileName)
   If Completed Then
      ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed
   Else 'Fehler rückmelden
      ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Error
   End If

End Sub

' AppFileBeforeSaveFile
Private Sub m_ApplicationHandler_AppFileBeforeSaveFile(ByVal FileID As String, ByVal FileName As String, _
                              ByVal SaveVersion As Boolean, ByRef ResumeMode As ApplicationHandlerResumeModes, _
                              ByRef ResumeMessage As Variant, _
                              ByVal ExtFieldName As String, ByVal ExtFieldValue As Variant)

   Dim Completed As Boolean

   Completed = SaveAppFile(FileID, FileName, SaveVersion, ExtFieldName, ExtFieldValue)
   If Completed Then
      ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed
   Else 'resume error
      ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Error
   End If

End Sub
