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
'  <file>base/ApplicationHandler_AppFileCustomDocumentProperties.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>text/Base64Converter.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
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

   Dim PropertyName As String
   
   PropertyName = FileID
   If ExtFilterFieldName = "BitInfo" Then
      PropertyName = PropertyName & ".x" & ExtFilterValue
   End If

   CreateAppFile = LoadFromCustomDocumentProperty(PropertyName, FileName)

End Function

Private Function LoadFromCustomDocumentProperty(ByVal PropertyName As String, ByVal TargetFilePath As String) As Boolean

   Dim Base64FileString As String
   Dim Prop As Object
   
   Set Prop = GetCustomProperty(PropertyName)
   If Prop Is Nothing Then
      Err.Raise vbObjectError, "LoadFromCustomDocumentProperties", "CustomDocumentProperty not exits"
   End If
   
   Base64FileString = Prop.Value
   
   If Len(Base64FileString) = 0 Then
      Err.Raise vbObjectError, "LoadFromCustomDocumentProperties", "CustomDocumentProperty is empty"
   End If
   
   With New Base64Converter
      .ConvertBase64ToFile TargetFilePath, Base64FileString
   End With
   
   LoadFromCustomDocumentProperty = True
   
End Function

Private Function GetCustomProperty(ByVal PropertyName As String) As Object

   Dim Prop As Object
   
   For Each Prop In ThisWorkbook.CustomDocumentProperties
      If Prop.Name = PropertyName Then
         Set GetCustomProperty = Prop
         Exit Function
      End If
   Next

End Function


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

   Dim PropertyName As String
   Dim Version As String
   Dim FileInfo As WinApiFileInfo
   
   PropertyName = FileID
   If ExtFilterFieldName = "BitInfo" Then
      PropertyName = PropertyName & ".x" & ExtFilterValue
   End If
   
   If SaveVersion Then
      Set FileInfo = New WinApiFileInfo
      Version = FileInfo.GetFileVersion(FileName)
          Set FileInfo = Nothing
      If Len(Version) = 0 Then
         Version = VBA.Format$(Nz(FileDateTime(FileName), vbNullString), "yyyy.mm.dd")
      End If
      SaveCustomDocumentProperty PropertyName & ".Version", Version
   End If

   SaveAppFile = SaveAsCustomDocumentProperty(PropertyName, FileName)

End Function

Private Function SaveAsCustomDocumentProperty(ByVal PropertyName As String, ByVal FilePath As String) As Boolean
   
   Dim Base64FileString As String
   
   With New Base64Converter
      Base64FileString = .ConvertFileToBase64(FilePath)
   End With
   
   SaveAsCustomDocumentProperty = SaveCustomDocumentProperty(PropertyName, Base64FileString)
   
End Function

Private Function SaveCustomDocumentProperty(ByVal PropertyName As String, ByRef PropertyValue As String) As Boolean
   
   Dim CustProp As Object
   
   Set CustProp = GetCustomProperty(PropertyName)
   If CustProp Is Nothing Then
      Set CustProp = ThisWorkbook.CustomDocumentProperties.Add( _
                                         Name:=PropertyName, _
                                         Value:=vbNullString, _
                                         LinkToContent:=False, _
                                         Type:=msoPropertyTypeString)
   End If

   CustProp.Value = PropertyValue
   
   SaveCustomDocumentProperty = True
   
End Function

'---------------------------------------------------------------------------------------
' Function: RemoveAppFileFormAddInStorage
'---------------------------------------------------------------------------------------
'
' Remove file from Add-In (CustomDocumentProperties)
'
' Parameters:
'     FileID   - Record identifier
'     FileName - File name to be created
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function RemoveAppFileFromAddInStorage(ByVal FileID As String, _
                     Optional ByVal ExtFilterFieldName As String, Optional ExtFilterValue As Variant) As Boolean

   Dim PropertyName As String
   
   PropertyName = FileID
   If ExtFilterFieldName = "BitInfo" Then
      PropertyName = PropertyName & ".x" & ExtFilterValue
   End If

   RemoveAppFileFromAddInStorage = RemoveCustomDocumentProperties(PropertyName)

End Function

Private Function RemoveCustomDocumentProperties(ByVal PropertyName As String) As Boolean

   Dim Prop As Object
   
   Set Prop = GetCustomProperty(PropertyName)
   If Not Prop Is Nothing Then
      Prop.Delete
   End If
   
   RemoveCustomDocumentProperties = True
   
End Function

Public Function GetStoredAppFileVersion(ByVal FileID As String, _
                               Optional ByVal ExtFilterFieldName As String, _
                               Optional ByVal ExtFilterValue As Variant) As String
   
   Dim PropertyName As String
   Dim Prop As Object
   
   PropertyName = FileID
   If ExtFilterFieldName = "BitInfo" Then
      PropertyName = PropertyName & ".x" & ExtFilterValue
   End If
   PropertyName = PropertyName & ".Version"
   
   Set Prop = GetCustomProperty(PropertyName)
   If Prop Is Nothing Then
      GetStoredAppFileVersion = vbNullString
      Debug.Print "CustomDocumentProperty "; PropertyName; " not exits"
      Exit Function
      'Err.Raise vbObjectError, "GetStoredAppFileVersion", "CustomDocumentProperty not exits"
   End If
   
   GetStoredAppFileVersion = Prop.Value
   
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


