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
' Klassenmodul: ApplicationHandler (Josef Pötzl, 2009-12-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hauptsteuerung der Anwendung
' </summary>
' <remarks>
' </remarks>
'\ingroup base
'**/
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
                               ' 0 ... keine Antwort
   AppResumeMode_Completed = 1 ' Anforderung wurde erfolgreich abgearbeitet
   AppResumeMode_Error = 2     ' Event wurde angenommen, aber es traten Fehler auf
   AppResumeMode_Cancel = 4    ' Event wurde angenommen, aber Weiterverabeitung soll gestoppt werden
End Enum

Public Enum ApplicationHandlerLogType
   [_AppLogType_Unknown] = 0
   AppLogType_Error = 1
   AppLogType_Tracing = 2
End Enum


' Integrierte Erweiterungen
Private Const EXTENSION_KEY_APPFILE As String = "AppFile"

' Fehlerkennungen
Private Const ERR_CLASS_ID As Long = 1000
Public Enum ApplicationHandlerErrors
   AppErr_EventInterfaceMissing = vbObjectError + ERR_CLASS_ID + 1

End Enum

' API-Funktionen
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

' Hilfskonstanten
Private Const LOG_FILE As String = "log.txt"

Private Const DBPROPNAME_APPTITLE As String = "AppTitle"
Private Const DBPROPNAME_APPICON As String = "AppIcon"
Private Const DBPROPNAME_USE_APPICON_FOR_FRMRPT As String = "UseAppIconForFrmRpt"
Private Const DBPROPNAME_STARTUPFORM As String = "StartUpForm"
Private Const DBPROPNAME_STARTUPMENUBAR As String = "StartUpMenuBar"


'Hilfsvariablen
Private m_AppDb As DAO.Database              ' Ersatz für CurrentDb bzw. CodeDb, siehe Prop: AppDb

Private m_ApplicationName As String          ' Zwischenspeicher für Anwendungsnamen (kurz)
Private m_ApplicationFullName As String      ' Zwischenspeicher für Anwendungsnamen (lang)

Private m_APPLICATIONVERSION As String       ' anzuzeigende Versionnummer

Private m_TransferValue As Variant
Private m_PublicPath As String               ' Standardverzeichnis für Dateiexplorer

Private m_ApplicationStartFormName As String ' Formular das in Start-Methode geöffnet wird

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

'Tracing-Schnittstelle
Public Event NewAppLog( _
            ByVal LogType As ApplicationHandlerLogType, ByVal Msg As String, ByVal Args As Variant, _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)


Public WriteLogToFileIfNoEventResponse As Boolean

'---------------------------------------------------------------------------------------
' Dispose-Ablauf
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
' Function: Start (Josef Pötzl, 2009-12-15)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Start der Anwendung
' </summary>
' <returns>Boolean</returns>
' <remarks>
' Nicht Verwechseln mit Initialisierung von ApplicationHandler \n
' Diese Prozedur löst die Startroutine für den Anwender aus.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Start(Optional ByRef ResumeMessage As Variant) As Boolean
   
   Dim CurrentStartFormName As String
   Dim EventResumeMode As ApplicationHandlerResumeModes
   Dim EventResumeMessage As Variant
   
   'Evente-Schnittstelle für Anwendungsstart
   ' ... ermöglicht z. B. das Durchführen eines Login-Vorgangs
   RaiseEvent BeforeStartApplication(EventResumeMode, EventResumeMessage)
   Select Case EventResumeMode
      Case ApplicationHandlerResumeModes.AppResumeMode_Cancel
         'Start abbrechen
         ResumeMessage = EventResumeMessage
         Start = False
         Exit Function
      Case ApplicationHandlerResumeModes.AppResumeMode_Completed
         'Startprozedur wurde von Erweiterung übernommen
         Start = EventResumeMessage
         Exit Function
      Case ApplicationHandlerResumeModes.AppResumeMode_Error
         'Mit Fehlermeldung abbrechen
         Err.Raise vbObjectError, "BeforeStartApplication", EventResumeMessage
         Exit Function
      Case Else
         'weitermachen
   End Select

   'Update prüfen
   If CheckVersionUpdate Then
      If Me.UpdateApplication Then
         Start = False
         Application.Quit acQuitSaveNone
         Exit Function
      End If
   End If
   
   'Evente-Schnittstelle für Startformular
   '... ermöglich das Ändern des Startformulars durch ERweiterungen
   '    z. B. falls ein benutzerspezifisches Formulare geöffnet werden soll
   EventResumeMode = 0
   EventResumeMessage = Empty
   RaiseEvent BeforeOpenStartForm(EventResumeMode, EventResumeMessage)
   Select Case EventResumeMode
      Case ApplicationHandlerResumeModes.AppResumeMode_Cancel
         'Startform nicht öffnen
         CurrentStartFormName = vbNullString
      Case ApplicationHandlerResumeModes.AppResumeMode_Completed
         'Name des StartFormulars wurde übertragen
         CurrentStartFormName = Nz(EventResumeMessage, vbNullString)
      Case Else
         CurrentStartFormName = Me.ApplicationStartFormName
   End Select
   
   If Len(CurrentStartFormName) > 0 Then
      DoCmd.OpenForm CurrentStartFormName
   End If

   'Benachrichtigung über erfolgten Anwendungsstart
   EventResumeMode = 0
   EventResumeMessage = Empty
   RaiseEvent AfterStartApplication(EventResumeMode, EventResumeMessage)

   Start = True

End Function


'---------------------------------------------------------------------------------------
' Property: ApplicationName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Name der aktuellen Anwendung
' </summary>
' <returns>String</returns>
' <remarks>
' Reihenfolge der Namensermittlung:
' 1. über gesetzten Wert
' 2. aus Titel-Eigenschaft mit CurrentDb.Properties("AppTitle")
' 3. aus Dateinamen mit CurrentDb.Name
' </remarks>
'**/
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

'---------------------------------------------------------------------------------------
' Property: ApplicationName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Name der aktuellen Anwendung einstellen
' </summary>
' <param name="appName">Anwendungsname</param>
'**/
'---------------------------------------------------------------------------------------
Public Property Let ApplicationName(ByVal AppName As String)
    m_ApplicationName = AppName
End Property


'---------------------------------------------------------------------------------------
' Property: ApplicationFullName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Vollständiger Name der aktuellen Anwendung
' </summary>
' <returns>String</returns>
' <remarks>
' Reihenfolge der Namensermittlung:
' 1. über gesetzten Wert
' 2. über ApplicationName
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ApplicationFullName() As String
   If Len(m_ApplicationFullName) = 0 Then 'Kurzform verwenden
      m_ApplicationFullName = ApplicationName
   End If
   ApplicationFullName = m_ApplicationFullName
End Property

'---------------------------------------------------------------------------------------
' Property: ApplicationFullName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Name der aktuellen Anwendung einstellen
' </summary>
' <param name="appName">Anwendungsname</param>
'**/
'---------------------------------------------------------------------------------------
Public Property Let ApplicationFullName(ByVal AppName As String)
    m_ApplicationFullName = AppName
End Property

'---------------------------------------------------------------------------------------
' Property: ApplicationVersion
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Vollständiger Name der aktuellen Anwendung
' </summary>
' <returns>String</returns>
' <remarks>
' Reihenfolge der Namensermittlung:
' 1. über gesetzten Wert
' 2. über ApplicationName
' </remarks>
'**/
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

'---------------------------------------------------------------------------------------
' Property: ApplicationFullName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Name der aktuellen Anwendung einstellen
' </summary>
' <param name="appName">Anwendungsname</param>
'**/
'---------------------------------------------------------------------------------------
Public Property Let Version(ByVal AppVersion As String)
    m_APPLICATIONVERSION = AppVersion
End Property

'---------------------------------------------------------------------------------------
' Property: ApplicationStartForm
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Formular, das in der Methode Start geöffnet wird
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ApplicationStartFormName() As String
   If StrPtr(m_ApplicationStartFormName) = 0 Then ' ... aus Properties lesen?
      m_ApplicationStartFormName = ""
   End If
   ApplicationStartFormName = m_ApplicationStartFormName
End Property

'---------------------------------------------------------------------------------------
' Property: ApplicationStartForm
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Formular, das in der Methode Start geöffnet wird
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let ApplicationStartFormName(ByVal FormName As String)
    m_ApplicationStartFormName = FormName
End Property


'---------------------------------------------------------------------------------------
' Property: AppDb (Josef Pötzl, 2009-12-13)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ersatz für CurrentDb bzw. CodeDb (Standard: CodeDb)
' </summary>
' <returns>zugewiesene DAO.Database-Instanz</returns>
' <remarks>
' </remarks>
'**/
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
      'niemand hat reagiert bzw. einen Abbruch gemeldet => WriteApplicationLogEntry
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


Public Property Get MdiBackColor() As Long
'/**
' * @bug Ermitteln der Hintergrundfarbe des MDI-Bereichs Funktioniert nicht unter Access 2007
'**/
   If m_CustomMDIBackgroundColor <> 0 Then
      MdiBackColor = m_CustomMDIBackgroundColor
   Else
      MdiBackColor = GetSysColor(SYSCOLOR_COLOR_APPWORKSPACE)
   End If
End Property

Friend Property Let MdiBackColor(ByVal Color As Long)
  
   Dim NewBrush As Long
   Dim hMDI As Long

   m_CustomMDIBackgroundColor = Color
   hMDI = FindWindowEx(Application.hWndAccessApp, 0&, "MDIClient", vbNullString)
   
   'Brush erzeugen
   NewBrush = CreateSolidBrush(m_CustomMDIBackgroundColor)
   'Brush zuweisen
   SetClassLong hMDI, GCL_HBRBACKGROUND, NewBrush
   'Fenster neuzeichnen (gesamtes Fenster inkl. Background)
   RedrawWindow hMDI, ByVal 0&, ByVal 0&, RDW_INVALIDATE Or RDW_ERASE

End Property


'---------------------------------------------------------------------------------------
' Function: CreateAppFile
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datei aus AppFile-Tabelle extrahieren
' </summary>
' <param name="sFileID">Kennung in Tabelle</param>
' <param name="sFileName"></param>
' <returns>True = Datei wurde gespeichert</returns>
' <remarks>
' Erfordert AppFile-Erweiterung (wird über Ereignis-Schnittstelle angesprochen)
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CreateAppFile(ByVal FileID As String, ByVal FileName As String) As Boolean

   Dim EventResumeMode As ApplicationHandlerResumeModes ' Rückgabewert aus Event-Schnittstelle
   Dim EventResumeMessage As Variant                    ' Nachricht, für Abbruchgrund (wird derzeit noch nicht genutzt)

   '---------------------------------------------------
   ' Schnittstelle für andere Klassen .. damit wird es möglich,
   ' dass die Erstellung der Datei eine andere Klasse übernimmt
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
'/
' <summary>
' Datei in AppFile-Tabelle speichern
' </summary>
' <param name="sFileID">Kennung in Tabelle (Feld "id")</param>
' <param name="sFileName">Dateiname</param>
' <param name="saveVersion">Version abspeichern</param>
' <remarks>
' Erfordert ApplicationHandler_File (wird über Ereignis-Schnittstelle angesprochen)
' </remarks>
'/
'---------------------------------------------------------------------------------------
Friend Function SaveAppFile(ByVal FileID As String, ByVal FileName As String, Optional ByVal SaveVersion As Boolean = False, _
            Optional ByVal ExtFieldName As String, Optional ByVal ExtFieldValue As Variant) As Boolean

   Dim EventResumeMode As ApplicationHandlerResumeModes ' Rückgabewert aus Event-Schnittstelle
   Dim EventResumeMessage As Variant                    ' Nachricht, für Abbruchgrund

   '---------------------------------------------------
   ' Event-Schnittstelle AppFileBeforeSaveFile
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


Public Property Get Extensions(ByVal ExtensionKey As String) As Object
'Anm.: auf Auflistungs-Klasse verzichtet .. der Code ist zwar nicht so elegant, dafür ist eine kaum verwendete Klasse weniger

'/**
' * @todo sollte man eine Schnittstelle statt Late binding verwenden?
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
' True = Anwendung schließen
 
' => Ausgelager an Erweiterung: damit wird die Verwendugn flexibler
'    und jeder kann sich aussuchen nach welcher Update-Methode vorgegangen werden soll

   Dim EventResumeMode As ApplicationHandlerResumeModes
   Dim EventResumeMessage As Variant

   RaiseEvent UpdateApplication(EventResumeMode, EventResumeMessage)
   If EventResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed Then
      UpdateApplication = EventResumeMessage
   Else
      UpdateApplication = False
   End If

End Function

Private Function CheckVersionUpdate() As Boolean
' True = Update erforderlich
   
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
' Sub: CallExtensionProcedure (Josef Pötzl, 2010-03-13)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prozeduren aus Erweiterungen aufrufen
' </summary>
' <param name="Key">Extension-Kennung</param>
' <param name="ProcedureName">Name der zu startenden Prozedur</param>
' <param name="ResumeMode">Rückgabekennung</param>
' <param name="ResumeMessage">Antwort</param>
' <param name="Params">Array für Übergabe- und Rückgabeparameter</param>
' <returns></returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub CallExtensionProcedure(ByVal Key As String, ByVal ProcedureName As String, _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant, _
            ByRef Params() As Variant)
            
   RaiseEvent ExtensionProcedureCall(Key, ProcedureName, ResumeMode, ResumeMessage, Params)
  
End Sub

Public Function NewerAppVersionExists() As Boolean

   Dim ResumeMode As ApplicationHandlerResumeModes
   Dim ResumeMessage As Boolean
   
   RaiseEvent NewVersionExists(ResumeMode, ResumeMessage)

   NewerAppVersionExists = ResumeMessage

End Function
