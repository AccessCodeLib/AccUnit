VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiImageTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: api.winapi.WinApiImageTools
'---------------------------------------------------------------------------------------
'
' WinAPI functions: Image & Co.
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/WinApiImageTools.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const WM_MSG_SETICON As Long = &H80
Private Const WM_PARAM_ICON_SMALL As Long = 0

Private Const IMAGE_ICON As Long = 1
Private Const LR_LOADFROMFILE As Long = &H10

#If VBA7 Then

Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
   ByVal Hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Declare PtrSafe Function LoadImage Lib "user32" Alias "LoadImageA" ( _
   ByVal hInst As Long, _
   ByVal lpszName As String, _
   ByVal uType As Long, _
   ByVal cxDesired As Long, _
   ByVal cyDesired As Long, _
   ByVal fuLoad As Long) As Long

#Else

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
   ByVal Hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" ( _
   ByVal hInst As Long, _
   ByVal lpszName As String, _
   ByVal uType As Long, _
   ByVal cxDesired As Long, _
   ByVal cyDesired As Long, _
   ByVal fuLoad As Long) As Long

#End If

'---------------------------------------------------------------------------------------
' Sub: SetFormIconFromFile
'---------------------------------------------------------------------------------------
'
' Spezielles Icon für ein Formular einstellen
'
' Parameters:
'     FormRef        - Reference to the Access.Form
'     IconFilePath   - Icon file (full path specification)
'
'---------------------------------------------------------------------------------------
Public Sub SetFormIconFromFile(ByRef FormRef As Access.Form, ByVal IconFilePath As String)

On Error Resume Next ' ... Error message would only interfere with this "unimportant" function

   Const ICONPIXELSIZE As Long = 16

   Dim ImageHandle As Long

   ImageHandle = LoadImage(0, IconFilePath, IMAGE_ICON, _
                           ICONPIXELSIZE, ICONPIXELSIZE, LR_LOADFROMFILE)
   If ImageHandle <> 0 Then
      SendMessage FormRef.Hwnd, WM_MSG_SETICON, WM_PARAM_ICON_SMALL, ByVal ImageHandle
   End If

End Sub
