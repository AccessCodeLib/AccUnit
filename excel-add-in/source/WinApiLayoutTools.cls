VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiLayoutTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: api.winapi.WinApiLayoutTools
'---------------------------------------------------------------------------------------
'
' WinAPI functions for layout design
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/WinApiLayoutTools.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Const GCL_HBRBACKGROUND As Long = -10
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ERASE As Long = &H4

Private Const HWND_DESKTOP As LongPtr = 0
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Const SM_CXVSCROLL As Long = 2

Private Declare PtrSafe Function CreateSolidBrush _
      Lib "gdi32.dll" ( _
      ByVal CrColor As Long _
      ) As LongPtr

Private Declare PtrSafe Function RedrawWindow _
      Lib "user32" ( _
      ByVal hWnd As LongPtr, _
      LprcUpdate As Any, _
      ByVal HrgnUpdate As Long, _
      ByVal FuRedraw As Long _
      ) As Long

Private Declare PtrSafe Function SetClassLong _
      Lib "User32.dll" _
      Alias "SetClassLongA" ( _
      ByVal hWnd As LongPtr, _
      ByVal nIndex As Long, _
      ByVal dwNewLong As LongPtr _
      ) As Long

Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long

Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal DC As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal DC As LongPtr, ByVal Index As Long) As LongPtr

Private Declare PtrSafe Function GetDpiForSystem Lib "User32.dll" () As Long


'---------------------------------------------------------------------------------------
' Sub: SetBackColor
'---------------------------------------------------------------------------------------
'
' Set background color of a window
'
' Parameters:
'     Hwnd  - Window handle
'     Color - Color number
'
'---------------------------------------------------------------------------------------
Public Sub SetBackColor(ByVal hWnd As LongPtr, ByVal Color As Long)
  
   Dim NewBrush As LongPtr
   
   'Create Brush
   NewBrush = CreateSolidBrush(Color)
   'Assign Brush
   SetClassLong hWnd, GCL_HBRBACKGROUND, NewBrush
   'Redraw window (entire window incl. background)
   RedrawWindow hWnd, ByVal 0&, ByVal 0&, RDW_INVALIDATE Or RDW_ERASE

End Sub

'---------------------------------------------------------------------------------------
' Function: TwipsPerPixelX
'---------------------------------------------------------------------------------------
'
' Width of a pixel in twips
'
' Returns:
'     Single
'
' Remarks:
'     http://support.microsoft.com/kb/94927
'
'---------------------------------------------------------------------------------------
Public Function TwipsPerPixelX() As Single
   Dim DC As LongPtr
   DC = GetDC(HWND_DESKTOP)
   TwipsPerPixelX = 1440& / GetDeviceCaps(DC, LOGPIXELSX) '* GetDpiForSystem / 96
   ReleaseDC HWND_DESKTOP, DC
End Function

'---------------------------------------------------------------------------------------
' Function: TwipsPerPixelY
'---------------------------------------------------------------------------------------
'
' Height of a pixel in twips
'
' Returns:
'     Single
'
' Remarks:
'     http://support.microsoft.com/kb/94927
'
'---------------------------------------------------------------------------------------
Public Function TwipsPerPixelY() As Single
   Dim DC As LongPtr
   DC = GetDC(HWND_DESKTOP)
   Debug.Print "GetDpiForSystem", GetDpiForSystem
   TwipsPerPixelY = 1440& / GetDeviceCaps(DC, LOGPIXELSY)
   ReleaseDC HWND_DESKTOP, DC
End Function

'---------------------------------------------------------------------------------------
' Function: GetScrollbarWidth
'---------------------------------------------------------------------------------------
'
' Width of the scroll bar
'
' Returns:
'     Single
'
'---------------------------------------------------------------------------------------
Public Function GetScrollbarWidth() As Single
   GetScrollbarWidth = GetSystemMetrics(SM_CXVSCROLL) * TwipsPerPixelX
End Function

'---------------------------------------------------------------------------------------
' Function: GetTwipsFromPixel
'---------------------------------------------------------------------------------------
'
' Converts pixels to twips
'
' Parameters:
'     Pixel - Number of pixels
'
' Returns:
'     Long
'
' Remarks:
'     GetTwipsFromPixel = TwipsPerPixelX * pixel
'
'---------------------------------------------------------------------------------------
Public Function GetTwipsFromPixel(ByVal Pixel As Long) As Long
   GetTwipsFromPixel = TwipsPerPixelX * Pixel
End Function

'---------------------------------------------------------------------------------------
' Function: GetPixelFromTwips
'---------------------------------------------------------------------------------------
'
' Converts twips to pixels
'
' Parameters:
'     Twips - Number of twips
'
' Returns:
'     Long
'
' Remarks:
'  GetPixelFromTwips = Twips / TwipsPerPixelX
'
'---------------------------------------------------------------------------------------
Public Function GetPixelFromTwips(ByVal Twips As Long) As Long
        GetPixelFromTwips = Twips / TwipsPerPixelX
End Function

Public Function DpiScale() As Double
   DpiScale = GetDpiForSystem / 96
End Function
