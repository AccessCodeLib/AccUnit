Attribute VB_Name = "WinApiToolsFactory"
Attribute VB_Description = "Gebräuchliche WinAPI-Funktionen"
'---------------------------------------------------------------------------------------
' Modul: WinApiToolsFactory
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erzeugt Instanz von WinApiTools
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/WinApiToolsFactory.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>api/winapi/WinApiTools.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit
Option Private Module

Private m_WinApi As WinApiTools

Public Property Get WinAPI() As WinApiTools
   If m_WinApi Is Nothing Then
      Set m_WinApi = New WinApiTools
   End If
   Set WinAPI = m_WinApi
End Property
