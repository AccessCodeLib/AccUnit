VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbeAddInInstaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Const SW_HIDE = 0

Private Const AddInProgId As String = "AccUnit.VbeAddIn.Connect"
Private Const AddInClassGuid As String = "{F15F18C3-CA43-421E-9585-6A04F51C5786}"

Private Const VbeUserControlHostProgId As String = "AccLib.VbeUserControlHost"
Private Const VbeUserControlHostClassGuid As String = "{030A1F2F-4E0B-4041-A7F5-C4C0B94BAF07}"

Public Enum VbeAddInLoadBehavior
   LoadBehavior_Disabled = 0
   LoadBehavior_LoadOnStartup = 1
   LoadBehavior_Loaded = 2
End Enum

Public Sub RegisterAccUnitVbeAddIn(ByVal DllPath As String, ByVal LoadBehavior As Long)

   Const FullClassName As String = "AccessCodeLib.AccUnit.VbeAddIn.Connect"
   Const AssemblyName As String = "AccUnit.VbeAddIn"
   Const TargetRuntimeVersion As String = "v4.0.30319"
   Const AssemblyFile As String = AssemblyName & ".dll"
   Const AddInFriendlyName As String = "AccUnit VBE Add-in"
   Const AddInDescription As String = "VBE Add-in for AccUnit Framework"

#If Win64 Then
   Const Bitness As Long = 64
   Const AddInsFolder As String = "Addins64"
#Else
   Const Bitness As Long = 32
   Const AddInsFolder As String = "Addins"
#End If

   Dim AssemblyVersion As String
   With New ApplicationHandler_AppFile
      AssemblyVersion = .GetStoredAppFileVersion("AccUnit.VbeAddIn.dll")
   End With

   If Right(DllPath, 1) <> "\" Then
      DllPath = DllPath & "\"
   End If

' Add-in COM dll
   RegisterNetComDll FullClassName, AddInClassGuid, AddInProgId, _
                     AssemblyName, DllPath & AssemblyFile, AssemblyVersion, Bitness, _
                     TargetRuntimeVersion

' Host
   RegisterVbeUserControlHost DllPath, Bitness

' VBE Add-in
   RegisterVbeAddIn AddInProgId, LoadBehavior, AddInFriendlyName, AddInDescription

End Sub

Private Sub RegisterVbeUserControlHost(ByVal DllPath As String, ByVal Bitness As Long)

   Const FullClassName As String = "AccessCodeLib.Common.VBIDETools.VbeUserControlHost"
   Const AssemblyName As String = "AccessCodeLib.Common.VbeUserControlHost"
   Const TargetRuntimeVersion As String = "v4.0.30319"
   Const AssemblyFile As String = AssemblyName & ".dll"

   Dim AssemblyVersion As String
   With New ApplicationHandler_AppFile
      AssemblyVersion = .GetStoredAppFileVersion("AccessCodeLib.Common.VbeUserControlHost.dll")
   End With

' COM dll
   RegisterNetComDll FullClassName, VbeUserControlHostClassGuid, VbeUserControlHostProgId, _
                     AssemblyName, DllPath & AssemblyFile, AssemblyVersion, Bitness, _
                     TargetRuntimeVersion

End Sub

Private Sub RegisterVbeAddIn(ByVal ProgId As String, _
                           ByVal LoadBehavior As Long, ByVal FriendlyName As String, ByVal Description As String)

#If Win64 Then
   Const AddInsFolder As String = "Addins64"
#Else
   Const AddInsFolder As String = "Addins"
#End If

   AddRegItem "HKCU\Software\Microsoft\VBA\VBE\6.0\" & AddInsFolder & "\" & ProgId, "LoadBehavior", "REG_DWORD", "0x0000000" & LoadBehavior
   AddRegItem "HKCU\Software\Microsoft\VBA\VBE\6.0\" & AddInsFolder & "\" & ProgId, "FriendlyName", "REG_SZ", FriendlyName
   AddRegItem "HKCU\Software\Microsoft\VBA\VBE\6.0\" & AddInsFolder & "\" & ProgId, "Description", "REG_SZ", Description


'%windir%\system32\REG ADD "HKCU\Software\Microsoft\VBA\VBE\6.0\%AddInsFolder%\%ProgId%" /v "LoadBehavior" /t REG_DWORD /d "0x00000003" /f
'%windir%\system32\REG ADD "HKCU\Software\Microsoft\VBA\VBE\6.0\%AddInsFolder%\%ProgId%" /v "FriendlyName" /t REG_SZ /d "AccUnit VBE Add-in" /f
'%windir%\system32\REG ADD "HKCU\Software\Microsoft\VBA\VBE\6.0\%AddInsFolder%\%ProgId%" /v "Description" /t REG_SZ /d "AccUnit VBE Add-in" /f

End Sub

Private Sub RegisterNetComDll(ByVal FullClassName As String, ByVal ClassGuid As String, ByVal ProgId As String, _
                              ByVal AssemblyName As String, ByVal AssemblyPath As String, ByVal AssemblyVersion As String, ByVal Bitness As Long, _
                              ByVal TargetRuntimeVersion As String)

   AddRegItem "HKCU\Software\Classes\" & ProgId, vbNullString, "REG_SZ", FullClassName
   AddRegItem "HKCU\Software\Classes\" & ProgId & "\CLSID", vbNullString, "REG_SZ", ClassGuid
   AddRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid, vbNullString, "REG_SZ", FullClassName, Bitness
   AddRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid & "\ProgId", vbNullString, "REG_SZ", ProgId, Bitness
   AddRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid & "\InprocServer32", vbNullString, "REG_SZ", "mscoree.dll", Bitness
   AddRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid & "\InprocServer32", "ThreadingModel", "REG_SZ", "Both", Bitness
   AddRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid & "\InprocServer32", "Class", "REG_SZ", FullClassName, Bitness
   AddRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid & "\InprocServer32", "Assembly", "REG_SZ", AssemblyName & ", Version=" & AssemblyVersion & ", Culture=neutral, PublicKeyToken=null", Bitness
   AddRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid & "\InprocServer32", "RuntimeVersion", "REG_SZ", TargetRuntimeVersion, Bitness
   AddRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid & "\InprocServer32", "CodeBase", "REG_SZ", "file:///" & Replace(AssemblyPath, "\", "/"), Bitness

End Sub

Private Sub AddRegItem(ByVal RegPath As String, _
                       ByVal RegValue As String, ByVal RegDataType As String, ByVal RegData As String, _
              Optional ByVal Bitness As Long = 0)

   Const ParamStringTemplate As String = "ADD ""{RegPath}"" /v ""{RegValue}"" /t {RegDataType} /d ""{RegData}"" /reg:{Bitness} /f"

   Dim ParamString As String

   ParamString = ParamStringTemplate
   ParamString = Replace(ParamString, "{RegPath}", RegPath)

   If Len(RegValue) = 0 Then
      ParamString = Replace(ParamString, "/v ""{RegValue}""", "/ve")
   Else
      ParamString = Replace(ParamString, "{RegValue}", RegValue)
   End If

   ParamString = Replace(ParamString, "{RegDataType}", RegDataType)
   ParamString = Replace(ParamString, "{RegData}", RegData)

   If Bitness = 0 Then
      ParamString = Replace(ParamString, "/reg:{Bitness} ", vbNullString)
   Else
      ParamString = Replace(ParamString, "{Bitness}", Bitness)
   End If

   RunRegExe ParamString

End Sub

Private Sub RunRegExe(ByVal Params As String)

   Dim RegExeFilePath As String
   RegExeFilePath = Environ("windir") & "\System32\REG.exe"
   ShellExecute Application.hWndAccessApp, "open", RegExeFilePath, Params, vbNullString, SW_HIDE

End Sub

Public Sub UnRegisterAccUnitVbeAddIn()

#If Win64 Then
   Const Bitness As Long = 64
   Const AddInsFolder As String = "Addins64"
#Else
   Const Bitness As Long = 32
   Const AddInsFolder As String = "Addins"
#End If

   If Me.AddInIsInstalled Then
      Me.LoadAccUnitAddIn False
   End If

   UnRegisterVbeAddIn AddInProgId, Bitness
   UnRegisterComDll AddInClassGuid, AddInProgId, Bitness
   UnRegisterComDll VbeUserControlHostClassGuid, VbeUserControlHostProgId, Bitness

End Sub

Private Sub UnRegisterComDll(ByVal ClassGuid As String, ByVal ProgId As String, Bitness As Long)

   DelRegItem "HKCU\Software\Classes\" & ProgId
   DelRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid, Bitness

End Sub

Private Sub UnRegisterVbeAddIn(ByVal ProgId As String, Bitness As Long)

   Dim AddInsFolder As String

   If Bitness = 64 Then
      AddInsFolder = "Addins64"
   Else
      AddInsFolder = "Addins"
   End If

   DelRegItem "HKCU\Software\Microsoft\VBA\VBE\6.0\" & AddInsFolder & "\" & ProgId

End Sub

Private Sub DelRegItem(ByVal RegPath As String, _
              Optional ByVal Bitness As Long = 0)

   Const ParamStringTemplate As String = "DELETE ""{RegPath}"" /reg:{Bitness} /f"

   Dim ParamString As String

   ParamString = ParamStringTemplate
   ParamString = Replace(ParamString, "{RegPath}", RegPath)

   If Bitness = 0 Then
      ParamString = Replace(ParamString, "/reg:{Bitness} ", vbNullString)
   Else
      ParamString = Replace(ParamString, "{Bitness}", Bitness)
   End If

   RunRegExe ParamString

End Sub


Public Property Get AddInIsInstalled() As Boolean
   AddInIsInstalled = FindAccUnitAddIn()
End Property

Public Property Get AddInIsLoaded() As Boolean

   Dim VbeAddIn As VBIDE.AddIn

   If Not FindAccUnitAddIn(VbeAddIn) Then
      AddInIsLoaded = False
      Exit Property
   End If

   AddInIsLoaded = VbeAddIn.Connect

End Property

Public Sub LoadAccUnitAddIn(ByVal LoadAddIn As Boolean)

   Dim VbeAddIn As VBIDE.AddIn

   If Not FindAccUnitAddIn(VbeAddIn) Then
      Exit Sub
   End If

   If LoadAddIn Then
      SetFocusToImmediateWindow
   End If

   VbeAddIn.Connect = LoadAddIn

End Sub

Private Sub SetFocusToImmediateWindow()
   Dim VbeWin As VBIDE.Window
   For Each VbeWin In Application.VBE.Windows
      If VbeWin.Type = VBIDE.vbext_WindowType.vbext_wt_Immediate Then
         If Not VbeWin.Visible Then
            VbeWin.Visible = True
         End If
         VbeWin.SetFocus
         Exit Sub
      End If
   Next
End Sub

Private Function FindAccUnitAddIn(Optional ByRef AddInRef As VBIDE.AddIn) As Boolean

   Dim VbeAddIn As VBIDE.AddIn

   For Each VbeAddIn In VBE.Addins
      If VbeAddIn.ProgId = AddInProgId Then
         Set AddInRef = VbeAddIn
         FindAccUnitAddIn = True
      End If
   Next

End Function
