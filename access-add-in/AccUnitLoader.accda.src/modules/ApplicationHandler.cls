VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Hauptsteuerung der Anwendung"
'---------------------------------------------------------------------------------------
' Class: base.ApplicationHandler
'---------------------------------------------------------------------------------------
'
' Main control of the application
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/ApplicationHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Public Enum ApplicationHandlerResumeModes
   [AppResumeMode_NoResponse] = 0   ' 0 ... no answer
   AppResumeMode_Completed = 1      ' Request was successfully processed
   AppResumeMode_Error = 2          ' Event was accepted, but errors occurred
   AppResumeMode_Cancel = 4         ' Event was accepted, but further processing should be stopped
End Enum

Public Enum ApplicationHandlerLogType
   [_AppLogType_Unknown] = 0
   AppLogType_Error = 1
   AppLogType_Tracing = 2
End Enum


' Integrated extensions
Private Const EXTENSION_KEY_APPFILE As String = "AppFile"

' Error codes
Private Const ERR_CLASS_ID As Long = 1000

Public Enum ApplicationHandlerErrors
   AppErr_EventInterfaceMissing = vbObjectError + ERR_CLASS_ID + 1

End Enum

' API definitions
Private Const GCL_HBRBACKGROUND As Long = -10
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ERASE As Long = &H4
Private Const SYSCOLOR_COLOR_APPWORKSPACE As Long = 12

#If VBA7 Then

Private Declare PtrSafe Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" ( _
       ByVal HWnd1 As Long, ByVal HWnd2 As Long, _
       ByVal Lpsz1 As String, ByVal Lpsz2 As String) As Long

Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32.dll" (ByVal CrColor As Long) As Long

Private Declare PtrSafe Function SetClassLong Lib "user32" Alias "SetClassLongA" ( _
       ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare PtrSafe Function RedrawWindow Lib "user32" ( _
       ByVal Hwnd As Long, LprcUpdate As Any, _
       ByVal HrgnUpdate As Long, ByVal FuRedraw As Long) As Long

Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal NumIndex As Long) As Long

#Else

Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" ( _
       ByVal HWnd1 As Long, ByVal HWnd2 As Long, _
       ByVal Lpsz1 As String, ByVal Lpsz2 As String) As Long

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal CrColor As Long) As Long

Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" ( _
       ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function RedrawWindow Lib "user32" ( _
       ByVal Hwnd As Long, LprcUpdate As Any, _
       ByVal HrgnUpdate As Long, ByVal FuRedraw As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal Index As Long) As Long

#End If

' Auxiliary constants
Private Const LOG_FILE As String = "log.txt"

Private Const DBPROPNAME_APPTITLE As String = "AppTitle"
Private Const DBPROPNAME_APPICON As String = "AppIcon"
Private Const DBPROPNAME_USE_APPICON_FOR_FRMRPT As String = "UseAppIconForFrmRpt"
Private Const DBPROPNAME_STARTUPFORM As String = "StartUpForm"
Private Const DBPROPNAME_STARTUPMENUBAR As String = "StartUpMenuBar"


'Auxiliary variables
Private m_AppDb As DAO.Database              ' Replacement for CurrentDb or CodeDb, see Prop: AppDb

Private m_ApplicationName As String          ' Application name cache (short)
Private m_ApplicationFullName As String      ' Application name cache (long)
Private m_APPLICATIONVERSION As String       ' Version number to be displayed

Private m_TransferValue As Variant
Private m_PublicPath As String               ' Default directory for file explorer

Private m_ApplicationStartFormName As String ' Form that is opened in Start method

Private m_CustomMDIBackgroundColor As Long

'Events
Public Event BeforeDispose(ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
Public Event AfterDispose(ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

Public Event CheckExtension(ByVal Key As String, ByRef Exists As Boolean)

Public Event ExtensionLookup(ByVal Key As String, ByRef ExtensionReference As Object)

Public Event ExtensionInstance(ByVal ExtensionKey As String, ByVal InstanzKey As String, ByRef ExtensionReference As Object)

Public Event ExtensionPropertyLookup(ByVal Key As String, ByVal PropertyName As String, _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

Public Event ExtensionProcedureCall(ByVal Key As String, ByVal ProcedureName As String, _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant, _
            ByRef Params() As Variant)

Public Event AppFileBeforeCreateFile( _
            ByVal sFileID As String, ByVal sFileName As String, _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

Public Event AppFileBeforeSaveFile( _
            ByVal sFileID As String, ByVal sFileName As String, ByVal SaveVersion As Boolean, _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant, _
            ByVal ExtFieldName As String, ByVal ExtFieldValue As Variant)

Public Event UpdateApplication( _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

Public Event CheckUpdate( _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

Public Event NewVersionExists( _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

Public Event BeforeOpenStartForm( _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

Public Event BeforeStartApplication( _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

Public Event AfterStartApplication( _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

'Tracing interface
Public Event NewAppLog( _
            ByVal LogType As ApplicationHandlerLogType, ByVal Msg As String, ByVal Args As Variant, _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

Public WriteLogToFileIfNoEventResponse As Boolean

'---------------------------------------------------------------------------------------
' Disposable support
'---------------------------------------------------------------------------------------
Private m_Disposed As Boolean
Private Sub Class_Terminate()
On Error Resume Next
   If Not m_Disposed Then
      Dispose
   End If
End Sub

Public Sub Dispose(Optional ByRef ResumeMode As ApplicationHandlerResumeModes, Optional ByRef ResumeMessage As Variant)

   Dim TempResumeMode As ApplicationHandlerResumeModes
   Dim TempResumeMessage As Variant

On Error Resume Next

   RaiseEvent BeforeDispose(TempResumeMode, TempResumeMessage)
   If TempResumeMode = AppResumeMode_Cancel Then
      ResumeMode = TempResumeMode
      ResumeMessage = TempResumeMessage
      Exit Sub
   End If

   Set m_AppDb = Nothing

   m_Disposed = True

   RaiseEvent AfterDispose(TempResumeMode, TempResumeMessage)

   ResumeMode = AppResumeMode_Completed
   ResumeMessage = TempResumeMessage

End Sub

'---------------------------------------------------------------------------------------
' Function: Start
'---------------------------------------------------------------------------------------
'
' Application start
'
' Returns:
'     Boolean     - true = success
'
' Remarks:
'     Do not confuse with initialization of ApplicationHandler.
'     This procedure triggers the startup routine for the user.
'
'---------------------------------------------------------------------------------------
Public Function Start(Optional ByRef ResumeMessage As Variant) As Boolean

   Dim CurrentStartFormName As String
   Dim EventResumeMode As ApplicationHandlerResumeModes
   Dim EventResumeMessage As Variant

   'Event interface for application start
   ' ... enables e.g. the execution of a login process
   RaiseEvent BeforeStartApplication(EventResumeMode, EventResumeMessage)
   Select Case EventResumeMode
      Case ApplicationHandlerResumeModes.AppResumeMode_Cancel
         'Cancel start
         ResumeMessage = EventResumeMessage
         Start = False
         Exit Function
      Case ApplicationHandlerResumeModes.AppResumeMode_Completed
         'Start procedure was taken over from extension
         Start = EventResumeMessage
         Exit Function
      Case ApplicationHandlerResumeModes.AppResumeMode_Error
         'Cancel with error message
         Err.Raise vbObjectError, "BeforeStartApplication", EventResumeMessage
         Exit Function
      Case Else
         'continue
   End Select

   'Check update
   If CheckVersionUpdate Then
      If Me.UpdateApplication Then
         Start = False
         Application.Quit acQuitSaveNone
         Exit Function
      End If
   End If

   'Event interface for start form
   '... allows changing the start form by means of extensions
   '    e.g. if a user-specific form is to be opened
   EventResumeMode = 0
   EventResumeMessage = Empty
   RaiseEvent BeforeOpenStartForm(EventResumeMode, EventResumeMessage)
   Select Case EventResumeMode
      Case ApplicationHandlerResumeModes.AppResumeMode_Cancel
         'Do not open startup form
         CurrentStartFormName = vbNullString
      Case ApplicationHandlerResumeModes.AppResumeMode_Completed
         'Name of the StartForm was transferred
         CurrentStartFormName = Nz(EventResumeMessage, vbNullString)
      Case Else
         CurrentStartFormName = Me.ApplicationStartFormName
   End Select

   If Len(CurrentStartFormName) > 0 Then
      DoCmd.OpenForm CurrentStartFormName
   End If

   'Notification about successful application start
   EventResumeMode = 0
   EventResumeMessage = Empty
   RaiseEvent AfterStartApplication(EventResumeMode, EventResumeMessage)

   Start = True

End Function


'---------------------------------------------------------------------------------------
' Property: ApplicationName
'---------------------------------------------------------------------------------------
'
' Name of the current application
'
' Remarks:
'     Sequence of name determination:
'        1. over set value
'        2. from title property with AppDb.Properties("AppTitle")
'        3. from filenames with AppDb.Name
'
'---------------------------------------------------------------------------------------
Public Property Get ApplicationName() As String
   If Len(m_ApplicationName) = 0 Then 'Wert aus Titel-Eigenschaft, da Konstante nicht eingestellt wurde
      On Error Resume Next
      m_ApplicationName = AppDb.Properties("AppTitle").Value
      If Len(m_ApplicationName) = 0 Then 'Wert aus Dateinamen
         m_ApplicationName = AppDb.Name
         m_ApplicationName = Left$(m_ApplicationName, InStrRev(m_ApplicationName, ".") - 1)
      End If
   End If
   ApplicationName = m_ApplicationName
End Property

Public Property Let ApplicationName(ByVal AppName As String)
    m_ApplicationName = AppName
End Property


'---------------------------------------------------------------------------------------
' Property: ApplicationFullName
'---------------------------------------------------------------------------------------
'
' Full name of the current application
'
' Remarks:
'     Sequence of name determination:
'        1. over set value
'        2. with ApplicationName property
'
'---------------------------------------------------------------------------------------
Public Property Get ApplicationFullName() As String
   If Len(m_ApplicationFullName) = 0 Then 'Kurzform verwenden
      m_ApplicationFullName = ApplicationName
   End If
   ApplicationFullName = m_ApplicationFullName
End Property

Public Property Let ApplicationFullName(ByVal AppName As String)
    m_ApplicationFullName = AppName
End Property

'---------------------------------------------------------------------------------------
' Property: ApplicationVersion
'---------------------------------------------------------------------------------------
'
' Application version
'
'---------------------------------------------------------------------------------------
Public Property Get Version() As String

   If Len(m_APPLICATIONVERSION) = 0 Then ' ... aus Properties lesen?
'/**
' @todo Versionskennung aus DB-Eigenschaften bzw. aus den Dateieigenschaften lesen
'**/
      m_APPLICATIONVERSION = ""
   End If

   Version = m_APPLICATIONVERSION

End Property

Public Property Let Version(ByVal AppVersion As String)
    m_APPLICATIONVERSION = AppVersion
End Property

'---------------------------------------------------------------------------------------
' Property: ApplicationStartForm
'---------------------------------------------------------------------------------------
'
' Form opened in the method <Start>
'
'---------------------------------------------------------------------------------------
Public Property Get ApplicationStartFormName() As String
   If StrPtr(m_ApplicationStartFormName) = 0 Then ' ... aus Properties lesen?
      m_ApplicationStartFormName = ""
   End If
   ApplicationStartFormName = m_ApplicationStartFormName
End Property

Public Property Let ApplicationStartFormName(ByVal FormName As String)
    m_ApplicationStartFormName = FormName
End Property

'---------------------------------------------------------------------------------------
' Property: AppDb
'---------------------------------------------------------------------------------------
'
' Replacement for CurrentDb or CodeDb (default: CodeDb).
'
'---------------------------------------------------------------------------------------
Public Property Get AppDb() As DAO.Database
   If m_AppDb Is Nothing Then
      Set m_AppDb = CodeDb
   End If
   Set AppDb = m_AppDb
End Property

Friend Property Set AppDb(ByVal DbRef As DAO.Database)
   Set m_AppDb = DbRef
End Property

Public Function GetTransferValue(Optional ByVal ValueIfNull As Variant = Null, Optional ByVal NoDel As Boolean = False) As Variant
   GetTransferValue = Nz(m_TransferValue, ValueIfNull)
   If Not NoDel Then m_TransferValue = Null
End Function

Public Sub SetTransferValue(ByVal NewValue As Variant)
   m_TransferValue = NewValue
End Sub

Public Property Get PublicPath() As String
   PublicPath = m_PublicPath
End Property

Friend Property Let PublicPath(ByVal NewPublicPath As String)
   m_PublicPath = NewPublicPath
End Property

Public Property Get ApplicationTitle() As String
   ApplicationTitle = AppDb.Properties("AppTitle").Value
End Property

Friend Property Let ApplicationTitle(ByVal AppTitle As String)
   If Len(AppTitle) > 0 Then
      SetAppDbProperty DBPROPNAME_APPTITLE, dbText, AppTitle
   Else
      DeleteAppDbProperty DBPROPNAME_APPTITLE
   End If
   Access.Application.RefreshTitleBar
End Property

Friend Sub SetAppIcon(ByVal AppIconFile As String, Optional ByVal UseAppIconForFrmRpt As Boolean = True)

   Dim CheckVar As Boolean

   If Len(Dir$(AppIconFile)) > 0 Then
      SetAppDbProperty DBPROPNAME_APPICON, dbText, AppIconFile
      SetAppDbProperty DBPROPNAME_USE_APPICON_FOR_FRMRPT, dbBoolean, UseAppIconForFrmRpt
   Else
      '1. Prüfen ob AppFile-Erweiterung vorhanden ist
      RaiseEvent CheckExtension(EXTENSION_KEY_APPFILE, CheckVar)
      If CheckVar Then
         CheckVar = CreateAppFile("AppIcon", AppIconFile)
      End If
      '2. Properties einstellen
      If CheckVar Then
         SetAppDbProperty DBPROPNAME_APPICON, dbText, AppIconFile
         SetAppDbProperty DBPROPNAME_USE_APPICON_FOR_FRMRPT, dbBoolean, UseAppIconForFrmRpt
      Else
         DeleteAppDbProperty DBPROPNAME_APPICON
      End If
   End If
   Application.RefreshTitleBar

End Sub

Friend Property Let StartUpForm(ByVal StartUpFormName As String)
   If Len(StartUpFormName) > 0 Then
      SetAppDbProperty DBPROPNAME_STARTUPFORM, dbText, "Form." & StartUpFormName
   Else
      DeleteAppDbProperty DBPROPNAME_STARTUPFORM
   End If
End Property

Friend Property Let StartUpMenuBar(ByVal MenuBarName As String)
   If Len(MenuBarName) > 0 Then
      SetAppDbProperty DBPROPNAME_STARTUPMENUBAR, dbText, MenuBarName
   Else
      DeleteAppDbProperty DBPROPNAME_STARTUPMENUBAR
   End If
End Property

Public Sub SetApplicationProperty(ByVal PropName As String, ByVal PropValue As Variant, Optional ByVal PropType As Long = dbText)
   SetAppDbProperty PropName, PropType, PropValue
End Sub

Private Sub SetAppDbProperty(ByVal PropName As String, ByVal PropType As Long, ByVal PropValue As Variant)

   Dim db As DAO.Database
   Dim PropCol As DAO.Properties

   Set db = AppDb
   Set PropCol = db.Properties
   If DbPropertyExists(PropCol, PropName) Then
      If Len(PropValue) = 0 Or IsNull(PropValue) Then
         db.Properties.Delete PropName
      Else
         db.Properties(PropName) = PropValue
      End If
   Else
      If Len(PropValue) > 0 Then
         db.Properties.Append db.CreateProperty(PropName, PropType, PropValue)
      End If
   End If

End Sub

Private Sub DeleteAppDbProperty(ByVal PropName As String)

   Dim db As DAO.Database
   Dim PropCol As DAO.Properties

   Set db = AppDb
   Set PropCol = db.Properties

   If DbPropertyExists(PropCol, PropName) Then
      PropCol.Delete PropName
   End If

End Sub

Private Function DbPropertyExists(ByRef PropCol As DAO.Properties, ByVal PropName As String) As Boolean

   Dim Prop As DAO.Property

   For Each Prop In PropCol
      If Prop.Name = PropName Then
         DbPropertyExists = True
         Exit For
      End If
   Next

End Function

Private Sub WriteApplicationLogEntry(ByVal Msg As String)

   Dim LogFile As String
   Dim FileNo As Long

   LogFile = CurrentProject.Path & "\" & LOG_FILE

   FileNo = FreeFile
   Open LogFile For Append As FileNo
   Print #FileNo, Msg
   Close #FileNo

End Sub

Public Sub WriteLog(ByVal Msg As String, _
           Optional ByRef LogType As ApplicationHandlerLogType, _
           Optional ByVal Args As Variant, _
           Optional ByVal WriteToFileIfNoEventResponse As Boolean = False)

   Dim ResumeMode As ApplicationHandlerResumeModes
   Dim EventResumeMessage As Variant

   RaiseEvent NewAppLog(LogType, Msg, Args, ResumeMode, EventResumeMessage)

   If ResumeMode = 0 And (WriteToFileIfNoEventResponse Or WriteLogToFileIfNoEventResponse Or LogType = AppLogType_Error) Then
      'no one has reacted or reported an abort => WriteApplicationLogEntry
      Msg = Now() & ": " & Msg
      If Len(LogType) > 0 Then
         Msg = "(" & GetLogTypeString(LogType) & ") " & Msg
      End If
      WriteApplicationLogEntry Msg
   End If

End Sub

Private Function GetLogTypeString(Optional ByRef LogType As ApplicationHandlerLogType) As String
   Select Case LogType
      Case ApplicationHandlerLogType.AppLogType_Error
         GetLogTypeString = "Error"
      Case ApplicationHandlerLogType.AppLogType_Tracing
         GetLogTypeString = "Tracing"
      Case Else
         GetLogTypeString = "N/A"
   End Select
End Function

'---------------------------------------------------------------------------------------
' Function: CreateAppFile
'---------------------------------------------------------------------------------------
'
' Extract file from AppFile table
'
' Parameters:
'     FileID   - Identifier in table
'     FileName - File name to be created
'
' Returns:
'     Boolean  - True: File was saved
'
' Remarks:
'     Requires extension <base.ApplicationHandler_AppFile> (addressed via event interface)
'
'---------------------------------------------------------------------------------------
Public Function CreateAppFile(ByVal FileID As String, ByVal FileName As String) As Boolean

   Dim EventResumeMode As ApplicationHandlerResumeModes ' Return value from event interface
   Dim EventResumeMessage As Variant                    ' Message, for termination reason (currently not used)

   '---------------------------------------------------
   ' Interface for other classes ...
   ' This makes it possible that the creation of the file takes over another class
   '
   EventResumeMode = 0
   RaiseEvent AppFileBeforeCreateFile(FileID, FileName, EventResumeMode, EventResumeMessage)
   If EventResumeMode <> 0 Then
      CreateAppFile = (EventResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed)
   Else
      Err.Raise ApplicationHandlerErrors.AppErr_EventInterfaceMissing, _
                "CreateAppFile", "Auf das AppFileBeforeCreateFile-Ereignis wurde nicht reagiert"
   End If

End Function

'---------------------------------------------------------------------------------------
' Function: SaveAppFile
'---------------------------------------------------------------------------------------
'
' Save file to AppFile table
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
' Remarks:
'     Requires extension <base.ApplicationHandler_AppFile> (addressed via event interface)
'
'---------------------------------------------------------------------------------------
Friend Function SaveAppFile(ByVal FileID As String, ByVal FileName As String, Optional ByVal SaveVersion As Boolean = False, _
            Optional ByVal ExtFieldName As String, Optional ByVal ExtFieldValue As Variant) As Boolean

   Dim EventResumeMode As ApplicationHandlerResumeModes ' Return value from event interface
   Dim EventResumeMessage As Variant                    ' Message, for termination reason

   '---------------------------------------------------
   ' Event interface AppFileBeforeSaveFile
   '
   EventResumeMode = 0
   RaiseEvent AppFileBeforeSaveFile(FileID, FileName, SaveVersion, EventResumeMode, EventResumeMessage, ExtFieldName, ExtFieldValue)
   If EventResumeMode <> 0 Then
      SaveAppFile = (EventResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed)
   Else
      Err.Raise ApplicationHandlerErrors.AppErr_EventInterfaceMissing, _
                "CreateAppFile", "Auf das AppFileBeforeSaveFile-Ereignis wurde nicht reagiert"
   End If

End Function

Public Function NewerAppVersionExists() As Boolean

   Dim ResumeMode As ApplicationHandlerResumeModes
   Dim ResumeMessage As Boolean

   RaiseEvent NewVersionExists(ResumeMode, ResumeMessage)

   NewerAppVersionExists = ResumeMessage

End Function

Private Function CheckVersionUpdate() As Boolean
' True = Update required

   Dim EventResumeMode As ApplicationHandlerResumeModes
   Dim EventResumeMessage As Variant

   RaiseEvent CheckUpdate(EventResumeMode, EventResumeMessage)
   If EventResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed Then
      CheckVersionUpdate = EventResumeMessage
   Else
      CheckVersionUpdate = False
   End If

End Function

'---------------------------------------------------------------------------------------
' Property: Extensions
'---------------------------------------------------------------------------------------
'
' Try get Reference from extension over event interface
'
' Parameters:
'     ExtensionKey         - Identifier of extension
'
' Returns:
'     Object  - True: File was saved in table
'
' Remarks:
'     Requires extension <base.ApplicationHandler_AppFile> (addressed via event interface)
'
'---------------------------------------------------------------------------------------
Public Property Get Extensions(ByVal ExtensionKey As String) As Object
'Note: collection class waived .. the code is not so elegant, but one less hardly used class

'/**
' * @todo Should you use an interface instead of late binding?
'**/

   Dim ExtRef As Object

   RaiseEvent ExtensionLookup(ExtensionKey, ExtRef)
   Set Extensions = ExtRef

End Property

Public Function GetExtensionInstance(ByVal ExtensionKey As String, Optional ByVal InstanceKey As String = vbNullString) As Object

   Dim ExtRef As Object

   RaiseEvent ExtensionInstance(ExtensionKey, InstanceKey, ExtRef)
   Set GetExtensionInstance = ExtRef

End Function

Public Function GetExtensionProperty(ByVal ExtensionKey As String, ByVal PropertyName As String, _
                            Optional ByVal ValueIfMissing As Variant = Empty) As Variant

   Dim EventResumeMode As ApplicationHandlerResumeModes
   Dim EventResumeMessage As Variant

   RaiseEvent ExtensionPropertyLookup(ExtensionKey, PropertyName, EventResumeMode, EventResumeMessage)

   If EventResumeMode = AppResumeMode_Completed Then
      GetExtensionProperty = EventResumeMessage
   Else
      GetExtensionProperty = ValueIfMissing
   End If

End Function

Public Function UpdateApplication() As Boolean
' True = Close application

' => Outsourced to extension: this makes the usage more flexible
'    and everyone can choose which update method to follow

   Dim EventResumeMode As ApplicationHandlerResumeModes
   Dim EventResumeMessage As Variant

   RaiseEvent UpdateApplication(EventResumeMode, EventResumeMessage)
   If EventResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed Then
      UpdateApplication = EventResumeMessage
   Else
      UpdateApplication = False
   End If

End Function

'---------------------------------------------------------------------------------------
' Sub: CallExtensionProcedure
'---------------------------------------------------------------------------------------
'
' Calling procedures from extensions
'
' Parameters:
'     Key            - Extension identifier
'     ProcedureName  - Name of the procedure to start
'     ResumeMode     - Return identifier
'     ResumeMessage  - Reply
'     Params         - Array for pass and return parameters
'
'---------------------------------------------------------------------------------------
Public Sub CallExtensionProcedure(ByVal Key As String, ByVal ProcedureName As String, _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant, _
            ByRef Params() As Variant)

   RaiseEvent ExtensionProcedureCall(Key, ProcedureName, ResumeMode, ResumeMessage, Params)

End Sub
