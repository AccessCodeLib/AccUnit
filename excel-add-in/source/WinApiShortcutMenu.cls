VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiShortcutMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: api.winapi.WinApiShortcutMenu
'---------------------------------------------------------------------------------------
'
' ShortcutMenu via Win API .. for Excel
'
' Author:
'     Josef Poetzl
'
' @todo Clean up code and create doc
'
' Remarks:
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' | Adapted variant from API example by Jörg Ostendorp of AEK10
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/WinApiShortcutMenu_Excel.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>api/winapi/WinApiLayoutTools.cls</use>
'  <use>api/winapi/WinApiWindowTools.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Type POINTAPI
   x As Long
   Y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

'Private m_Helper As LongPtr 'Auxiliary window handle

Private Type APIMSG
   hWnd As LongPtr
   Message As LongPtr
   wParam As Long
   lParam As Long
   Time As LongPtr
   pt As POINTAPI
End Type

'Auxiliary window
Private Declare PtrSafe Function CreateWindowEx _
      Lib "User32.dll" _
      Alias "CreateWindowExA" ( _
      ByVal dwExStyle As Long, _
      ByVal lpClassname As String, _
      ByVal lpWindowName As String, _
      ByVal dwStyle As Long, _
      ByVal x As Long, _
      ByVal Y As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long, _
      ByVal hwndParent As LongPtr, _
      ByVal hMenu As LongPtr, _
      ByVal hInstance As LongPtr, _
      ByRef lpParam As Any) As Long

'Standard menu functions
Private Declare PtrSafe Function CreatePopupMenu _
      Lib "User32.dll" () As Long

Private Declare PtrSafe Function AppendMenu _
      Lib "User32.dll" _
      Alias "AppendMenuA" ( _
      ByVal hMenu As LongPtr, _
      ByVal wFlags As Long, _
      ByVal wIDNewItem As LongPtr, _
      ByVal lpNewItem As Any _
      ) As Long

Private Declare PtrSafe Function TrackPopupMenu _
      Lib "User32.dll" ( _
      ByVal hMenu As LongPtr, _
      ByVal wFlags As Long, _
      ByVal x As Long, _
      ByVal Y As Long, _
      ByVal nReserved As Long, _
      ByVal hWnd As LongPtr, _
      lprc As Any _
      ) As LongPtr 'lprc As RECT

Private Declare PtrSafe Function DestroyMenu _
      Lib "User32.dll" ( _
         ByVal hMenu As LongPtr _
      ) As Long

'Message Loop
Private Declare PtrSafe Function GetMessage _
      Lib "User32.dll" _
      Alias "GetMessageA" ( _
      ByRef lpMsg As APIMSG, _
      ByVal hWnd As LongPtr, _
      ByVal wMsgFilterMin As Long, _
      ByVal wMsgFilterMax As Long _
      ) As Long

Private Declare PtrSafe Function DispatchMessage _
      Lib "User32.dll" _
      Alias "DispatchMessageA" ( _
      ByRef lpMsg As APIMSG _
      ) As Long

Private Declare PtrSafe Function TranslateMessage _
      Lib "User32.dll" (ByRef lpMsg As APIMSG) As Long

Private Declare PtrSafe Function GetWindowRect _
      Lib "User32.dll" ( _
         ByVal hWnd As LongPtr, _
         ByRef lpRect As RECT _
      ) As LongPtr
      
Private Declare PtrSafe Function SetMenuDefaultItem _
      Lib "user32" ( _
         ByVal hMenu As LongPtr, _
         ByVal uItem As LongPtr, _
         ByVal fByPos As LongPtr _
      ) As LongPtr
   
      
'Message
Private Const WM_COMMAND As Long = &H111
Private Const WM_MOUSELEAVE As Long = &H2A3

'Menu or item styles

Public Enum MenuItemStyle
   MF_SEPARATOR = &H800&
   MF_STRING = &H0&
   MF_GRAYED = &H1&
   MF_BITMAP = &H4&
   MF_POPUP = &H10&
   MF_BYCOMMAND = &H0&
   MF_MENUBARBREAK = &H20&
   MF_SYSMENU = &H2000&
   MF_CHECKED = &H8&
End Enum

'Menu alignment
Private Const TPM_BOTTOMALIGN As Long = &H20&
'Private Const TPM_TOPALIGN As Long = &H0&

'Auxiliary window
Private Const WS_CHILD As Long = &H40000000


Private Declare PtrSafe Function FindWindowEx Lib "User32.dll" Alias "FindWindowExA" ( _
         ByVal HWnd1 As LongPtr, _
         ByVal HWnd2 As LongPtr, _
         ByVal Lpsz1 As String, _
         ByVal Lpsz2 As String _
      ) As LongPtr
      
      
'---- additions:

Private Type WinAPIMenuHandlerItem
   ItemNumber As Long
   ItemText As String
   ItemType As Long
   SubMenu As Long
   DefaultItem As Boolean
End Type

Private m_MenuItems() As WinAPIMenuHandlerItem
Private m_ItemCnt As Long
Private m_AccessForm As MSForms.UserForm
Private m_MenuControl As MSForms.control
Private m_Section As Long
Private m_WindowTools As WinApiWindowTools
Private m_LayoutTools As WinApiLayoutTools


Public UserFormCaption As String

Private Sub Class_Initialize()
On Error Resume Next
   m_ItemCnt = 0
End Sub

Private Sub Class_Terminate()
On Error Resume Next
   Set m_MenuControl = Nothing
   Set m_AccessForm = Nothing
End Sub

Public Property Get Object() As WinApiShortcutMenu
   Set Object = Me
End Property

Private Property Get WindowTools() As WinApiWindowTools
   If m_WindowTools Is Nothing Then
      Set m_WindowTools = New WinApiWindowTools
   End If
   Set WindowTools = m_WindowTools
End Property

Private Property Get LayoutTools() As WinApiLayoutTools
   If m_LayoutTools Is Nothing Then
      Set m_LayoutTools = New WinApiLayoutTools
   End If
   Set LayoutTools = m_LayoutTools
End Property


Public Property Let ControlSection(ByVal SectionIndex As Long)
On Error Resume Next
   m_Section = SectionIndex
End Property

Public Property Get MenuControl() As Object
   Set MenuControl = m_MenuControl
End Property

Public Property Set MenuControl(ByVal MnuCtl As control)
   
   Dim ParentObj As Object

   Set m_MenuControl = MnuCtl

   If m_AccessForm Is Nothing Then
      Set ParentObj = m_MenuControl.Parent
      Set m_AccessForm = ParentObj
   End If

End Property

Public Property Get AccessForm() As UserForm
   Set AccessForm = m_AccessForm
End Property

Public Property Set AccessForm(ByRef FrmRef As UserForm)
   Set m_AccessForm = FrmRef
End Property

'---------------------------------------------------------------------------------------
' Sub: AddMenuItem
'---------------------------------------------------------------------------------------
'
' Create menu item
'
' Parameters:
'     MenuItemNumber - Number of the menu item
'     ItemText       - Menu item text
'     ItemType       - Menu item type
'     SubMenu        - Number of the assigned submenu
'
'---------------------------------------------------------------------------------------
Public Sub AddMenuItem(ByVal MenuItemNumber As Long, ByVal ItemText As String, _
              Optional ByVal ItemType As MenuItemStyle = MF_STRING, _
              Optional ByVal SubMenu As Long = 0, _
              Optional ByVal DefaultItem As Boolean = False)
              
   m_ItemCnt = m_ItemCnt + 1
   ReDim Preserve m_MenuItems(m_ItemCnt)
   
   With m_MenuItems(m_ItemCnt)
      .ItemNumber = MenuItemNumber
      .ItemText = ItemText
      .ItemType = ItemType
      .SubMenu = SubMenu
      .DefaultItem = DefaultItem
   End With

End Sub

'---------------------------------------------------------------------------------------
' Function: OpenMenu
'---------------------------------------------------------------------------------------
'
' Öffnet das Popup-Menü
'
' Parameters:
'     X - desired X position ... can be omitted if Control was specified
'     X - desired Y-position ... can be omitted if Control was specified
'
' Returns:
'     Long - Number of the selected menu item
'
'---------------------------------------------------------------------------------------
Public Function OpenMenu(Optional ByVal x As Single = 0, Optional ByVal Y As Single = 0) As Long

   Dim MenuHwnd As LongPtr
   Dim FormHwnd As LongPtr
   Dim TmpHwnd As LongPtr
   Dim hWnd As LongPtr

   Dim Message As APIMSG
   Dim RcHelper As RECT
   Dim ButtonLeft As Long
   Dim ButtonTop As Long
   Dim XPos As Long
   Dim YPos As Long
   
   Dim AccFormRect As RECT
   Dim ScaleFactorX As Double
   Dim ScaleFactorY As Double
   Dim DummyBorderWidth As Long

   'Otherwise, nothing is displayed when ENTER is pressed:
   DoEvents
   
   'Create empty menu
   MenuHwnd = CreatePopupMenu
   
   'Add entries
   Dim i As Long
   
   For i = 1 To m_ItemCnt
      With m_MenuItems(i)
         If .SubMenu > 0 Then
            If TmpHwnd = 0 Then
               TmpHwnd = CreatePopupMenu
            End If
            If .ItemType And MF_POPUP = MF_POPUP Then
               AppendMenu MenuHwnd, MF_STRING Or MF_POPUP, TmpHwnd, .ItemText
               TmpHwnd = 0
            Else
               AppendMenu TmpHwnd, .ItemType, .ItemNumber, .ItemText
            End If
         Else
            AppendMenu MenuHwnd, .ItemType, .ItemNumber, .ItemText
         End If
      
         If .DefaultItem Then SetMenuDefaultItem MenuHwnd, .ItemNumber - 1, &H400&
      
      End With
      
   Next i
   'Show menu
   
   FormHwnd = GetUserFormHwnd(m_AccessForm)
   hWnd = FindWindowEx(FormHwnd, 0, "F3 Server 53f80000", "")
   If hWnd = 0 Then
   '   Debug.Print "use FormHwnd"
      hWnd = FormHwnd
      DummyBorderWidth = 5
   End If
 '  Stop
   GetWindowRect hWnd, RcHelper
   
   
   ScaleFactorX = (RcHelper.Right - RcHelper.Left) / (m_AccessForm.InsideWidth + DummyBorderWidth)
   ScaleFactorY = (RcHelper.Bottom - RcHelper.Top) / m_AccessForm.InsideHeight
   
   'Debug.Print ScaleFactorX, ScaleFactorY
   ButtonLeft = (m_MenuControl.Left + x * 0) * ScaleFactorX
   ButtonTop = (m_MenuControl.Top + Y * 0) * ScaleFactorY
   
   XPos = RcHelper.Left + ButtonLeft - 1
   YPos = RcHelper.Top + ButtonTop + 1
  
   TrackPopupMenu MenuHwnd, TPM_BOTTOMALIGN, XPos, YPos, ByVal 0&, FormHwnd, ByVal 0&
   
   If FormHwnd = 0 Then
      DestroyMenu MenuHwnd
      Exit Function
   End If
   
   Dim RepeatGetMessage As Boolean
   Do ' Security sleep in case there are problems with windows message. (WM_MOUSELEAVE]
      RepeatGetMessage = False
      GetMessage Message, FormHwnd, ByVal 0&, ByVal 0&
      TranslateMessage Message
      DispatchMessage Message
      
      'Evaluate message
      If Message.Message = WM_COMMAND Then
         OpenMenu = Message.wParam
      ElseIf Message.Message = WM_MOUSELEAVE Then
         RepeatGetMessage = True
      Else
         OpenMenu = 0
      End If
   Loop While RepeatGetMessage
   
   'Destroy menu
   DestroyMenu MenuHwnd

End Function

Private Function GetUserFormHwnd(ByVal UserFormRef As UserForm) As LongPtr

    Dim hWnd As LongPtr
    Dim FormHwnd As LongPtr
    Dim FormCaption As String
    
    FormCaption = UserFormRef.Caption
    If Len(FormCaption) = 0 Then
      FormCaption = UserFormCaption
    End If
    
    FormHwnd = Application.hWnd
    ' Attempt to find the window
    hWnd = FindWindowEx(0, 0, "ThunderDFrame", FormCaption)
    'hWnd = FindWindow("ThunderDFrame", FormCaption)
    If hWnd = 0 Then hWnd = FindWindowEx(0, 0, "ThunderXFrame", FormCaption)
    
    GetUserFormHwnd = hWnd
    
End Function

