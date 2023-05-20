VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler_ExtensionCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: ApplicationHandler_ExtensionCollection
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Einbindung der Erweiterungen für ApplicationHandler-Framework
' </summary>
' <remarks>
' Verwendung im _config_Application-Modul:\n
' <code>
'   Set m_Extensions = New ApplicationHandler_ExtensionCollection
'   With m_Extensions
'      Set .ApplicationHandler = oCurrentAppHandler
'
'      ' Erweiterungen laden (z. B. ApplicationHandler_DbConnection):
'      .Add New ApplicationHandler_DbConnection
'
'   End With
' </code>
' </remarks>
'\ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/ApplicationHandler_ExtensionCollection.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private WithEvents m_ApplicationHandler As ApplicationHandler
Attribute m_ApplicationHandler.VB_VarHelpID = -1
Private m_Extension As Collection
Private m_Disposed As Boolean

'---------------------------------------------------------------------------------------
' Sub: Extensions
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Collection der Erweiterungen
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get Extensions() As Collection
   If m_Extension Is Nothing Then
      Set m_Extension = New Collection
   End If
   Set Extensions = m_Extension
End Property

Public Property Get ApplicationHandler() As ApplicationHandler
   Set ApplicationHandler = m_ApplicationHandler
End Property

Public Property Set ApplicationHandler(ByVal NewRef As ApplicationHandler)
   Set m_ApplicationHandler = NewRef
End Property

'---------------------------------------------------------------------------------------
' Sub: Add
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erweiterung zu Collection hinzufügen
' </summary>
' <param name="objRef">Referenz auf Instanz der Erweiterung</param>
' <remarks>
' Referenz wird in Collection abgelegt, damit keine zusätzliche (manuelle)
' Referenzspeicherung notwendig ist.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub Add(ByRef ExtensionRef As Object)
   Set ExtensionRef.ApplicationHandlerRef = CurrentApplication
   Extensions.Add ExtensionRef, ExtensionRef.ExtensionKey
End Sub

Public Sub Dispose()
   
   Dim MaxCnt As Long
   Dim CheckCnt As Long

On Error Resume Next

   If Not (m_Extension Is Nothing) Then
      MaxCnt = m_Extension.Count * 2 'nur zur Sicherheit falls wider Erwarten m_Extension.Remove eine Endlosschleife bringen würde
      Do While m_Extension.Count > 0 Or CheckCnt > MaxCnt
         m_Extension.Remove 1
         CheckCnt = CheckCnt + 1
      Loop
      Set m_Extension = Nothing
   End If

End Sub

Private Sub Class_Terminate()
On Error Resume Next
   Dispose
End Sub

Private Sub m_ApplicationHandler_AfterDispose(ResumeMode As ApplicationHandlerResumeModes, ResumeMessage As Variant)
   Dispose
End Sub
