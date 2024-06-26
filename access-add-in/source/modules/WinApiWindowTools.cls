﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiWindowTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: api.winapi.WinApiWindowTools
'---------------------------------------------------------------------------------------
'
' WinAPI functions: Window Handle & Co.
'
' Author:
'     Josef Poetzl
'
' Source info:
'---------------------------------------------------------------------------------------
'| The procedures GetMDI, GetHeaderSection, GetDetailSection, GetFooterSection and GetControl
'| are taken from the AEK10 lecture by Jörg Ostendorp
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/WinApiWindowTools.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Type POINTAPI
   x As Long
   Y As Long
End Type

Private Declare PtrSafe Function ClientToScreen Lib "user32.dll" ( _
         ByVal hwnd As LongPtr, _
         ByRef lpPoint As POINTAPI _
      ) As Long

Private Declare PtrSafe Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" ( _
         ByVal HWnd1 As LongPtr, _
         ByVal HWnd2 As LongPtr, _
         ByVal Lpsz1 As String, _
         ByVal Lpsz2 As String _
      ) As LongPtr

'---------------------------------------------------------------------------------------
' Function: GetMDI
'---------------------------------------------------------------------------------------
'
' Determines the handle of the MDI client window
'
' Returns:
'     Handle (LongPtr)
'
'---------------------------------------------------------------------------------------
Public Function GetMDI() As LongPtr
   Dim h As LongPtr
   h = Application.hWndAccessApp
   'Find the first (and only) "MDIClient" child window of the application window.
   GetMDI = FindWindowEx(h, 0&, "MDIClient", vbNullString)
End Function

'---------------------------------------------------------------------------------------
' Function: GetHeaderSection
'---------------------------------------------------------------------------------------
'
' Determines the handle for the header area of a form
'
' Parameters:
'     Hwnd - Handle of the form (Form.Hwnd)
'
' Returns:
'     LongPtr
'
'---------------------------------------------------------------------------------------
Public Function GetHeaderSection(ByVal hwnd As LongPtr) As LongPtr
   Dim h As LongPtr
   'Determine first "OFormsub" child window of the form (fhwnd)
   h = FindWindowEx(hwnd, 0&, "OformSub", vbNullString)
   GetHeaderSection = h
End Function

'---------------------------------------------------------------------------------------
' Function: GetDetailSection
'---------------------------------------------------------------------------------------
'
' Determines the handle for the detail area of a form
'
' Parameters:
'     Hwnd - Handle of the form (Form.Hwnd)
'
' Returns:
'     LongPtr
'
'---------------------------------------------------------------------------------------
Public Function GetDetailSection(ByVal hwnd As LongPtr) As LongPtr
   Dim h As LongPtr
   'Determine first "OFormsub" child window of the form (fhwnd), starting after the header area.
   h = GetHeaderSection(hwnd)
   h = FindWindowEx(hwnd, h, "OformSub", vbNullString)
   GetDetailSection = h
End Function

'---------------------------------------------------------------------------------------
' Function: GetFooterSection
'---------------------------------------------------------------------------------------
'
' Determines the handle for the footer of a form
'
' Parameters:
'     Hwnd - Handle of the form (Form.Hwnd)
'
' Returns:
'     LongPtr
'
'---------------------------------------------------------------------------------------
Public Function GetFooterSection(ByVal hwnd As Long) As LongPtr
   Dim h As LongPtr
   'Determine first "OFormsub" child window of the form (fhwnd), starting after the detail area.
   h = GetDetailSection(hwnd)
   h = FindWindowEx(hwnd, h, "OformSub", vbNullString)
   GetFooterSection = h
End Function

'---------------------------------------------------------------------------------------
' Function: GetControl
'---------------------------------------------------------------------------------------
'
' Determines the handle of any control
'
' Parameters:
'     FormRef     - Form reference
'     Hwnd        - Handle of the area on which the control is located (Header, Detail, Footer)
'     ClassName   - Name of the window class of the control
'     ControlName - Name des Controls
'
' Returns:
'     LongPtr
'
'---------------------------------------------------------------------------------------
Public Function GetControl(ByRef FormRef As Access.Form, ByVal hwnd As LongPtr, _
                           ByVal ClassName As String, ByVal ControlName As String) As LongPtr

   'Exitieren mehrere Controls der gleichen Klasse auf einem Formular, z.B. TabControls, besteht das Problem, daß
   'deren Reihenfolge nicht definiert ist (anders also als bei den Sektionsfenstern)
   'In diesem Fall kann man alle Kindfenster dieser Klasse in einer Schleife durchlaufen
   'und z.B. prüfen, ob die Position des Fensters des zurückgegebenen Handles
   'mit der des Access-Steuerelementes übereinstimmt.
   'Nachfolgend wird hierfür die undokumentierte Funktion accHittest verwendet.
   'Dieser werden als Parameter die Screenkoordinaten der linken oberen Ecke eines
   'Steuerelementes übergeben. Befindet sich dort ein Objekt, erhält man dieses als Rückgabewert.
   'Ist der Name des Objektes identisch mit dem übergebenen Steuerelementnamen, so
   'hat man das Handle ermittelt:

On Error Resume Next

   Dim h As LongPtr
   Dim obj As Object
   Dim pt As POINTAPI

   h = 0

   Do
      'Erstes (h=0)/nächstes (h<>0) Control auf dem Sektionsfenster ermitteln
      h = FindWindowEx(hwnd, h, ClassName, vbNullString)

      'Bildschirmkoordinaten dieses Controls ermitteln
      'dafür die Punktkoordinaten aus dem letzten Durchlauf zurücksetzen, sonst wird addiert!
      pt.x = 0
      pt.Y = 0
      ClientToScreen h, pt

      'Objekt bei den Koordinaten ermitteln
      Set obj = FormRef.accHitTest(pt.x, pt.Y)

      'Wenn Objektname = Tabname Ausstieg aus der Schleife
      If obj.Name = ControlName Then
         Exit Do
      End If
   Loop While h <> 0

   'Handle zurückgeben
   GetControl = h

End Function
