Attribute VB_Name = "defGlobal"
Attribute VB_Description = "Allgemeine Konstanten und Eigenschaften"
'---------------------------------------------------------------------------------------
' Modul: defGlobal
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Allgemeine Konstanten und Eigenschaften
' </summary>
' <remarks></remarks>
' \ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/defGlobal.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>base/modApplication.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Explicit
Option Compare Text
Option Private Module

'---------------------------------------------------------------------------------------
'
' Konstanten
'


'---------------------------------------------------------------------------------------
'
' Hilfs-Variablen
'


'---------------------------------------------------------------------------------------
'
' Hilfs-Prozeduren
'

'
' Private Hilfsvariablen für die Prozeduren
'
Private m_ApplicationName As String         'Zwischenspeicher für Anwendungsnamen, falls
                                            'CurrentApplication.ApplicationName nicht läuft

'---------------------------------------------------------------------------------------
' Property: CurrentApplicationName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Name der aktuellen Anwendung
' </summary>
' <returns>String</returns>
' <remarks>
' Verwendet CurrentApplication.ApplicationName
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get CurrentApplicationName() As String
' inkl. Notfall-Errorhandler, falls CurrentApplication nicht instanziert ist

On Error GoTo HandleErr

   CurrentApplicationName = CurrentApplication.ApplicationName
      
ExitHere:
   Exit Property

HandleErr:
   CurrentApplicationName = GetApplicationNameFromDb
   Resume ExitHere

End Property

Private Function GetApplicationNameFromDb() As String

   If Len(m_ApplicationName) = 0 Then 'Wert aus Titel-Eigenschaft, da Konstante nicht eingestellt wurde
          On Error Resume Next
      m_ApplicationName = CodeDb.Properties("AppTitle").Value
      If Len(m_ApplicationName) = 0 Then 'Wert aus Dateinamen
         m_ApplicationName = CodeDb.Name
         m_ApplicationName = Left$(m_ApplicationName, InStrRev(m_ApplicationName, ".") - 1)
      End If
   End If
   
   GetApplicationNameFromDb = m_ApplicationName
   
End Function
