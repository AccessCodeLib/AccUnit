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
' Class: base.ApplicationHandler_ExtensionCollection
'---------------------------------------------------------------------------------------
'
' Integration of the extensions for ApplicationHandler framework
'
' Author:
'     Josef Poetzl
'
' Remarks:
'     Use in the _config_Application module:
' <code>
'   Set m_Extensions = New ApplicationHandler_ExtensionCollection
'   With m_Extensions
'      Set .ApplicationHandler = oCurrentAppHandler
'
'      ' Load extensions (e.g. ApplicationHandler_DbConnection):
'      .Add New ApplicationHandler_DbConnection
'
'   End With
' </code>
'
'---------------------------------------------------------------------------------------

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
' Property: Extensions
'---------------------------------------------------------------------------------------
'
' Collection of extensions
'
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
'
' Add extension to collection
'
' Parameters:
'     ExtensionRef - Reference to instance of extension
'
' Remarks:
' Reference is stored in Collection so that no additional (manual) reference storage is necessary.
'
'---------------------------------------------------------------------------------------
Public Sub Add(ByVal ExtensionRef As Object)
   Set ExtensionRef.ApplicationHandlerRef = CurrentApplication
   Extensions.Add ExtensionRef, ExtensionRef.ExtensionKey
End Sub

Public Sub Dispose()
   
   Dim MaxCnt As Long
   Dim CheckCnt As Long

On Error Resume Next

   If Not (m_Extension Is Nothing) Then
      MaxCnt = m_Extension.Count * 2 'just to be on the safe side in case m_Extension.Remove would bring an infinite loop against expectations
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
