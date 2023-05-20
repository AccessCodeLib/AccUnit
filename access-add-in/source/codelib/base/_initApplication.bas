Attribute VB_Name = "_initApplication"
'---------------------------------------------------------------------------------------
' Modul: _initApplication (2009-07-08)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Initialisierungsaufruf der Anwendung
' </summary>
' <remarks>
' </remarks>
' \ingroup base
' @todo StartApplication-Prozedur für allgemeine Verwendung umschreiben => in Klasse verlagern
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/_initApplication.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>base/modApplication.bas</use>
'  <use>base/defGlobal.bas</use>
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
' => siehe _config_Application
'
'-------------------------

'---------------------------------------------------------------------------------------
' Function: StartApplication
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prozedur für den Anwendungsstart
' </summary>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function StartApplication() As Boolean

On Error GoTo HandleErr

   StartApplication = CurrentApplication.Start

ExitHere:
   Exit Function

HandleErr:
   StartApplication = False
   MsgBox "Anwendung kann nicht gestartet werden.", vbCritical, CurrentApplicationName
   Application.Quit acQuitSaveNone
   Resume ExitHere

End Function

Public Sub RestoreApplicationDefaultSettings()
   On Error Resume Next
   CurrentApplication.ApplicationTitle = CurrentApplication.ApplicationFullName
End Sub
