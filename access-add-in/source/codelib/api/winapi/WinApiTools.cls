VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: WinApiTools
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Sammlung von WinAPI-Tools
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'
'<codelib>
'  <file>api/winapi/WinApiTools.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>api/winapi/WinApiShellTools.cls</use>
'  <use>api/winapi/WinApiImageTools.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private m_Shell As WinApiShellTools
Private m_Image As WinApiImageTools

Public Property Get Shell() As WinApiShellTools
   If m_Shell Is Nothing Then
      Set m_Shell = New WinApiShellTools
   End If
   Set Shell = m_Shell
End Property

Public Property Get Image() As WinApiImageTools
   If m_Image Is Nothing Then
      Set m_Image = New WinApiImageTools
   End If
   Set Image = m_Image
End Property
