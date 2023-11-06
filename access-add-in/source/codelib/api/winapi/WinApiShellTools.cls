VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiShellTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: api.winapi.WinApiShellTools
'---------------------------------------------------------------------------------------
'
' WinAPI functions: ShellExecuteA & Co.
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/WinApiShellTools.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Const SE_ERR_NOTFOUND As Long = 2
Private Const SE_ERR_NOASSOC  As Long = 31
Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const NORMAL_PRIORITY_CLASS As Long = &H20

Private Type StartUpInfo
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type ProcessInformation
   hProcess As Long
   hThread As Long
   dwProcessID As Long
   dwThreadID As Long
End Type

Private Const INFINITE As Long = &HFFFFFFFF ' = -1&
Private Const WAIT_TIMEOUT As Long = &H102&

#If VBA7 Then

Private Declare PtrSafe Function ShellExecuteA Lib "shell32.dll" ( _
   ByVal Hwnd As Long, _
   ByVal lOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long
   
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long

Private Declare PtrSafe Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" ( _
   ByVal lpBuffer As String, _
   ByVal nSize As Long) As Long

Private Declare PtrSafe Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
   ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
   ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   lpEnvironment As Any, ByVal lpCurrentDirectory As String, _
   lpStartupInfo As StartUpInfo, lpProcessInformation As ProcessInformation) As Long

Private Declare PtrSafe Function WaitForInputIdle Lib "user32" (ByVal Process As Long, ByVal DwMilliseconds As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal Object As Long) As Long
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal Handle As Long, ByVal DwMilliseconds As Long) As Long

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal DwMilliseconds As Long)

#Else

Private Declare Function ShellExecuteA Lib "shell32.dll" ( _
   ByVal Hwnd As Long, _
   ByVal lOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long
   
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" ( _
   ByVal lpBuffer As String, _
   ByVal nSize As Long) As Long

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
   ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
   ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   lpEnvironment As Any, ByVal lpCurrentDirectory As String, _
   lpStartupInfo As StartUpInfo, lpProcessInformation As ProcessInformation) As Long

Private Declare Function WaitForInputIdle Lib "user32" (ByVal Process As Long, ByVal DwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal Object As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal Handle As Long, ByVal DwMilliseconds As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal DwMilliseconds As Long)
#End If


'---------------------------------------------------------------------------------------
' Encapsulations
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Function: ShellExecuteOpenFile
'---------------------------------------------------------------------------------------
'
' Open file with ShellExecute
'
' Parameters:
'     FilePath       - full file name (incl. directory)
'     ApiOperation   - (optional) "open", "print", ..., Default: vbNullString
'     ShowCmd        - (optional) VbAppWinStyle, Default: vbNormalFocus
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function Execute(ByVal FilePath As String, _
               Optional ByVal ApiOperation As String = vbNullString, _
               Optional ByVal ShowCmd As VbAppWinStyle = vbNormalFocus) As Boolean

   Dim Ret As Long
   Dim Directory As String
   Dim DeskWin As Long
   
   If Len(FilePath) = 0 Then
      Execute = False
      Exit Function
   Else
      DeskWin = GetDesktopWindow()
      Ret = ShellExecuteA(DeskWin, ApiOperation, FilePath, vbNullString, vbNullString, ShowCmd)
   End If
   
   If Ret = SE_ERR_NOTFOUND Then
      MsgBox "File not found" & vbNewLine & vbNewLine & _
             FilePath
      Execute = False
      Exit Function
   ElseIf Ret = SE_ERR_NOASSOC Then
      'If the file extension is not yet known, the "Open with..." dialog is displayed.
      Directory = Space$(260)
      Ret = GetSystemDirectory(Directory, Len(Directory))
      Directory = Left$(Directory, Ret)
      Call ShellExecuteA(DeskWin, vbNullString, "RUNDLL32.EXE", "shell32.dll, OpenAs_RunDLL " & _
         FilePath, Directory, ShowCmd)
   End If
   
   Execute = True

End Function

'---------------------------------------------------------------------------------------
' Function: ExecuteAsAdmin
'---------------------------------------------------------------------------------------
'
' Execute command with extended administrative privilege
'
' Parameters:
'     File - full file name incl. directory
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function ExecuteAsAdmin(ByVal File As String) As Boolean
    ExecuteAsAdmin = Me.Execute(File, "runas")
End Function

'---------------------------------------------------------------------------------------
' Function: ShellExecuteSendMail
'---------------------------------------------------------------------------------------
'
' Send email with standard program
'
' Parameters:
'     SendTo   - Recipient address
'     Subject  - Subject line
'     Body     - Email text
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function ShellExecuteSendMail(ByVal SendTo As String, _
                                     ByVal Subject As String, _
                                     ByVal Body As String) As Boolean

   Dim Ret As Long
   Dim lpFile As String
   
   If Len(SendTo) = 0 Then
      ShellExecuteSendMail = False
      Exit Function
   End If
   
   If Subject > vbNullString Then
      lpFile = "subject=" & Subject
   End If
   If Body > vbNullString Then
      If lpFile > vbNullString Then
         lpFile = lpFile & "&body=" & Body
      Else
         lpFile = "body=" & Body
      End If
   End If
   If lpFile > vbNullString Then
       lpFile = "mailto:" & SendTo & "?" & lpFile
   Else
      lpFile = "mailto:" & SendTo
   End If

   Ret = ShellExecuteA(GetDesktopWindow(), "open", lpFile, vbNullString, vbNullString, vbNormalFocus)
   ShellExecuteSendMail = (Ret <> 0)

End Function

'---------------------------------------------------------------------------------------
' Function: LaunchAppSynchronous
'---------------------------------------------------------------------------------------
'
' Anwnedung Synchron ausführen
'
' Parameters:
'     ExecutablePathAndName   - Executable file
'     Param                   - (optional) Start parameters of the application
'     ShowCommand             - (optional) window mode, Default: vbNormalFocus
'
' Returns:
'     Boolean
'
' Remarks:
'     Code hält so lange an, bis die gestartete Anwendung beendet wurde.
'
'---------------------------------------------------------------------------------------
Public Function LaunchAppSynchronous(ByVal ExecutablePathAndName As String, _
                     Optional ByVal Param As String = vbNullString, _
                     Optional ByVal ShowCommand As Long = vbNormalFocus) As Boolean
   
   'http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=3716&lngWId=1

   Dim Response As Long
   Dim StartUpInf As StartUpInfo
   Dim ProcInfo As ProcessInformation

   LaunchAppSynchronous = False

   With StartUpInf
      .cb = Len(StartUpInf)
      .lpReserved = vbNullString
      .lpDesktop = vbNullString
      .lpTitle = vbNullString
      .dwFlags = STARTF_USESHOWWINDOW
      .wShowWindow = ShowCommand
   End With

   'Launch the application by creating a new process
   Response = CreateProcess(vbNullString, ExecutablePathAndName & " " & Param, 0, 0, True, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, StartUpInf, ProcInfo)

   If Response Then
      'Wait for the application to terminate before moving on
      Call WaitForTermination(ProcInfo)
      LaunchAppSynchronous = True
   Else
      LaunchAppSynchronous = False
   End If

End Function

Private Sub WaitForTermination(ByRef ProcessInfo As ProcessInformation)
   'This wait routine allows other application events
   'to be processed while waiting for the process to complete.
   
   Dim Response As Long
   'Let the process initialize
   Call WaitForInputIdle(ProcessInfo.hProcess, INFINITE)
   'We don't need the thread handle so get rid of it
   Call CloseHandle(ProcessInfo.hThread)
   'Wait for the application to end

   Do
      Response = WaitForSingleObject(ProcessInfo.hProcess, 0)
      If Response <> WAIT_TIMEOUT Then
         'No timeout, app is terminated
         Exit Do
      End If
      DoEvents
      Loop While True

      'Kill the last handle of the process
      Call CloseHandle(ProcessInfo.hProcess)

End Sub
