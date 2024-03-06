Attribute VB_Name = "_initApplication"
'---------------------------------------------------------------------------------------
' Package: base._initApplication
'---------------------------------------------------------------------------------------
'
' Initialising the application
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/_initApplication.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>base/modApplication.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

'-------------------------
' Anwendungseinstellungen
'-------------------------
'
' => see _config_Application
'
'-------------------------

'---------------------------------------------------------------------------------------
' Function: StartApplication
'---------------------------------------------------------------------------------------
'
' Procedure for application start-up
'
' Returns:
'     Boolean - sucess = true
'
'---------------------------------------------------------------------------------------
Public Function StartApplication() As Boolean

On Error GoTo HandleErr

   StartApplication = CurrentApplication.Start

ExitHere:
   Exit Function

HandleErr:
   StartApplication = False
   MsgBox "Application can not be started.", vbCritical, CurrentApplicationName
   Application.Quit acQuitSaveNone
   Resume ExitHere

End Function

Public Sub RestoreApplicationDefaultSettings()
   On Error Resume Next
   CurrentApplication.ApplicationTitle = CurrentApplication.ApplicationFullName
End Sub
