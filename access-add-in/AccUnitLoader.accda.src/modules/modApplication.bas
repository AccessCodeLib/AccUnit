Attribute VB_Name = "modApplication"
Attribute VB_Description = "Standard-Prozeduren für die Arbeit mit ApplicationHandler"
'---------------------------------------------------------------------------------------
' Package: base.modApplication
'---------------------------------------------------------------------------------------
'
' Standard procedures for working with ApplicationHandler
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/modApplication.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>base/_config_Application.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

' Instance of the main control
Private m_ApplicationHandler As ApplicationHandler
Private m_ApplicationName As String  ' Cache for application names
                                     ' if CurrentApplication.ApplicationName is not running

'---------------------------------------------------------------------------------------
' Property: CurrentApplication
'---------------------------------------------------------------------------------------
'
' Property for ApplicationHandler instance (use this property in code)
'
'---------------------------------------------------------------------------------------
Public Property Get CurrentApplication() As ApplicationHandler
   If m_ApplicationHandler Is Nothing Then
      InitApplication
   End If
   Set CurrentApplication = m_ApplicationHandler
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentApplicationName
'---------------------------------------------------------------------------------------
'
' Name of the current application
'
' Remarks:
'     Uses CurrentApplication.ApplicationName
'
'---------------------------------------------------------------------------------------
Public Property Get CurrentApplicationName() As String
' incl. emergency error handler if CurrentApplication is not instantiated

On Error GoTo HandleErr

   CurrentApplicationName = CurrentApplication.ApplicationName

ExitHere:
   Exit Property

HandleErr:
   CurrentApplicationName = GetApplicationNameFromDb
   Resume ExitHere

End Property

Private Function GetApplicationNameFromDb() As String

   If Len(m_ApplicationName) = 0 Then
On Error Resume Next
'1. Value from title property
      m_ApplicationName = CodeDb.Properties("AppTitle").Value
      If Len(m_ApplicationName) = 0 Then
'2. Value from file name
         m_ApplicationName = CodeDb.Name
         m_ApplicationName = Left$(m_ApplicationName, InStrRev(m_ApplicationName, ".") - 1)
      End If
   End If

   GetApplicationNameFromDb = m_ApplicationName

End Function

'---------------------------------------------------------------------------------------
' Sub: TraceLog
'---------------------------------------------------------------------------------------
'
' TraceLog
'
' Parameters:
'     Msg   - Message text
'     Args  - (ParamArray)
'
'---------------------------------------------------------------------------------------
Public Sub TraceLog(ByRef Msg As String, ParamArray Args() As Variant)
   CurrentApplication.WriteLog Msg, ApplicationHandlerLogType.AppLogType_Tracing, Args
End Sub

Private Sub InitApplication()

   Set m_ApplicationHandler = New ApplicationHandler
   Call InitConfig(m_ApplicationHandler)

End Sub


'---------------------------------------------------------------------------------------
' Sub: DisposeCurrentApplicationHandler
'---------------------------------------------------------------------------------------
'
' Destroy instance of ApplicationHandler and the extensions
'
'---------------------------------------------------------------------------------------
Public Sub DisposeCurrentApplicationHandler()

   Dim CheckCnt As Long, MaxCnt As Long

On Error Resume Next

   If Not m_ApplicationHandler Is Nothing Then
      m_ApplicationHandler.Dispose
   End If

   Set m_ApplicationHandler = Nothing

End Sub


'---------------------------------------------------------------------------------------
'
' Auxiliary procedures
Public Sub WriteApplicationLogEntry(ByVal Msg As String, _
           Optional LogType As ApplicationHandlerLogType, _
           Optional ByVal Args As Variant)
   CurrentApplication.WriteLog Msg, LogType, Args
End Sub

Public Property Get PublicPath() As String
   PublicPath = CurrentApplication.PublicPath
End Property
