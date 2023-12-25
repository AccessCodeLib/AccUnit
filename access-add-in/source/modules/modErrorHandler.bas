Attribute VB_Name = "modErrorHandler"
Attribute VB_Description = "Prozeduren für die Fehlerbehandlung"
'---------------------------------------------------------------------------------------
' Package: base.modErrorHandler
'---------------------------------------------------------------------------------------
'
' Error handling procedures
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/modErrorHandler.bas</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

'---------------------------------------------------------------------------------------
' Enum: ACLibErrorHandlerMode
'---------------------------------------------------------------------------------------
'
' ErrorHandler Modes (error handling variants)
'
'   aclibErrRaise    - (0) Pass error to application
'   aclibErrMsgBox   - (1) Show error in MsgBox
'   aclibErrIgnore   - (2) Ignore error, do not display any message
'   aclibErrFile     - (4) Write error information to file
'
' Remarks:
'   The values {0,1,2} exclude each other. The value 4 (aclibErrFile) can be added arbitrarily to {0,1,2}.
'   Example: Init aclibErrRaise + aclibErrFile
'
Public Enum ACLibErrorHandlerMode
   [_aclibErr_default] = -1
   aclibErrRaise = 0&    'Pass error to application
   aclibErrMsgBox = 1&   'MsgBox
   aclibErrIgnore = 2&   'ignore error, do not display any message
   aclibErrFile = 4&     'Output to file
End Enum

'---------------------------------------------------------------------------------------
' Enum: ACLibErrorResumeMode
'---------------------------------------------------------------------------------------
'
' Processing parameters in case of errors
'
'     aclibErrExit         - (0) Termination (function exit)
'     aclibErrResume       - (1) Resume, Problem fixed externally
'     aclibErrResumeNext   - (2) Resume next, continue working in the code at the next point
'
' Remarks:
'     Used for error events
'
Public Enum ACLibErrorResumeMode
   aclibErrExit = 0       'Termination (function exit)
   aclibErrResume = 1     'Resume, Problem fixed externally
   aclibErrResumeNext = 2 'Resume next, continue working in the code at the next point
End Enum

'---------------------------------------------------------------------------------------
' Enum: ACLibErrorNumbers
'---------------------------------------------------------------------------------------
Public Enum ACLibErrorNumbers
   ERRNR_NOOBJECT = vbObjectError + 1001
   ERRNR_NOCONFIG = vbObjectError + 1002
   ERRNR_INACTIVE = vbObjectError + 1003
   ERRNR_FORBIDDEN = vbObjectError + 9001
End Enum

'Default settings:
Private Const DEFAULT_ERRORHANDLERMODE As Long = ACLibErrorHandlerMode.[_aclibErr_default]
Private Const DEFAULT_ERRORRESUMEMODE As Long = ACLibErrorResumeMode.aclibErrExit

Private Const ERRORSOURCE_DELIMITERSYMBOL As String = "->"

'Auxiliary variables
Private m_DefaultErrorHandlerMode As Long
Private m_ErrorHandlerLogFile As String

'---------------------------------------------------------------------------------------
' Property: DefaultErrorHandlerMode
'---------------------------------------------------------------------------------------
'
' Default behaviour of error handling
'
'---------------------------------------------------------------------------------------
Public Property Get DefaultErrorHandlerMode() As ACLibErrorHandlerMode
On Error Resume Next
    DefaultErrorHandlerMode = m_DefaultErrorHandlerMode
End Property

Public Property Let DefaultErrorHandlerMode(ByVal ErrMode As ACLibErrorHandlerMode)
    m_DefaultErrorHandlerMode = ErrMode
End Property

'---------------------------------------------------------------------------------------
' Property: ErrorHandlerLogFile
'---------------------------------------------------------------------------------------
'
' Log file for error message
'
'---------------------------------------------------------------------------------------
Public Property Get ErrorHandlerLogFile() As String
    ErrorHandlerLogFile = m_ErrorHandlerLogFile
End Property

Public Property Let ErrorHandlerLogFile(ByVal Path As String)
'/**
' * @todo: Checking for the existence of the file or at least the directory
'**/
    m_ErrorHandlerLogFile = Path
End Property

'---------------------------------------------------------------------------------------
' Function: HandleError
'---------------------------------------------------------------------------------------
'
' Standard procedure for error handling
'
' Parameters:
'     ErrNumber"></param>
'     ErrSource"></param>
'     ErrDescription"></param>
'     ErrHandlerMode"></param>
'
' Returns:
'     ACLibErrorResumeMode
'
' Remarks:
'Example:
'<code>
'Private Sub ExampleProc()
'
'On Error GoTo HandleErr
'
'[...]
'
'ExitHere:
'On Error Resume Next
'   Exit Sub
'
'HandleErr:
'   Select Case HandleError(Err.Number, "ExampleProc", Err.Description)
'   Case ACLibErrorResumeMode.aclibErrResume
'      Resume
'   Case ACLibErrorResumeMode.aclibErrResumeNext
'      Resume Next
'   Case Else
'      Resume ExitHere
'   End Select
'
'End Sub
'<code>
'
'---------------------------------------------------------------------------------------
Public Function HandleError(ByVal ErrNumber As Long, ByVal ErrSource As String, _
                   Optional ByVal ErrDescription As String, _
                   Optional ByVal ErrHandlerMode As ACLibErrorHandlerMode = DEFAULT_ERRORHANDLERMODE _
            ) As ACLibErrorResumeMode
'Here it would also be possible to activate another ErrorHandler (e.g. ErrorHandler class).

   If ErrHandlerMode = ACLibErrorHandlerMode.[_aclibErr_default] Then
      ErrHandlerMode = m_DefaultErrorHandlerMode
   End If

   HandleError = ProcHandleError(ErrNumber, ErrSource, ErrDescription, ErrHandlerMode)

End Function

Private Function ProcHandleError(ByRef ErrNumber As Long, ByRef ErrSource As String, _
                                 ByRef ErrDescription As String, _
                                 ByVal ErrHandlerMode As ACLibErrorHandlerMode _
             ) As ACLibErrorResumeMode

   Dim NewErrSource As String
   Dim NewErrDescription As String
   Dim CurrentErrSource As String

   NewErrDescription = Err.Description
   CurrentErrSource = Err.Source

On Error Resume Next

   NewErrSource = ErrSource
   If Len(NewErrSource) = 0 Then
      NewErrSource = CurrentErrSource
   ElseIf CurrentErrSource <> GetApplicationVbProjectName Then
      NewErrSource = NewErrSource & ERRORSOURCE_DELIMITERSYMBOL & CurrentErrSource
   End If

   If Len(ErrDescription) > 0 Then
      NewErrDescription = ErrDescription
   End If

   'Output to file
   If (ErrHandlerMode And ACLibErrorHandlerMode.aclibErrFile) Then
      PrintToFile ErrNumber, NewErrSource, NewErrDescription
      ErrHandlerMode = ErrHandlerMode - ACLibErrorHandlerMode.aclibErrFile
   End If

'Error handler
   Err.Clear
On Error GoTo 0
   Select Case ErrHandlerMode
      Case ACLibErrorHandlerMode.aclibErrRaise     ' Passing to the application
         Err.Raise ErrNumber, NewErrSource, NewErrDescription
      Case ACLibErrorHandlerMode.aclibErrMsgBox    ' show Msgbox
         ShowErrorMessage ErrNumber, NewErrSource, NewErrDescription
      Case ACLibErrorHandlerMode.aclibErrIgnore    'Skip error
         '
      Case Else '(should never actually occur) ... pass on to application
         Err.Raise ErrNumber, NewErrSource, NewErrDescription
   End Select

   'return resume mode
   ProcHandleError = DEFAULT_ERRORRESUMEMODE ' This will help when using a class

End Function

Public Sub ShowErrorMessage(ByVal ErrNumber As Long, ByRef ErrSource As String, ByRef ErrDescription As String)

   Dim ErrMsgBoxTitle As String
   Dim Pos As Long
   Dim TempString As String

On Error Resume Next

   Const LineBreakPos As Long = 50

   Pos = InStr(1, ErrSource, ERRORSOURCE_DELIMITERSYMBOL, vbBinaryCompare)
   If Pos > 1 Then
      ErrMsgBoxTitle = Left$(ErrSource, Pos - 1)
   Else
      ErrMsgBoxTitle = ErrSource
   End If

   If Len(ErrSource) > LineBreakPos Then
      Pos = InStr(LineBreakPos, ErrSource, ERRORSOURCE_DELIMITERSYMBOL)
      If Pos > 0 Then
         Do While Pos > 0
            TempString = TempString & Left$(ErrSource, Pos - 1) & vbNewLine
            ErrSource = Mid$(ErrSource, Pos)
            Pos = InStr(LineBreakPos, ErrSource, ERRORSOURCE_DELIMITERSYMBOL)
         Loop
         ErrSource = TempString & ErrSource
      End If
   End If

   VBA.MsgBox "Error " & ErrNumber & ": " & vbNewLine & ErrDescription & vbNewLine & vbNewLine & "(" & ErrSource & ")", _
         vbCritical + vbSystemModal + vbMsgBoxSetForeground, ErrMsgBoxTitle

End Sub

Private Sub PrintToFile(ByRef ErrNumber As Long, ByRef ErrSource As String, _
                        ByRef ErrDescription As String)

   Dim FileSource As String
   Dim f As Long
   Dim WriteToFile As Boolean
   Dim PathToErrLogFile As String

On Error Resume Next

   WriteToFile = True

   FileSource = "[" & ErrSource & "]"
   PathToErrLogFile = ErrorHandlerLogFile
   If Len(PathToErrLogFile) = 0 Then
      PathToErrLogFile = CurrentProject.Path & "\Error.log"
   End If
   f = FreeFile
   Open PathToErrLogFile For Append As #f
      Print #f, Format$(Now(), _
            "yyyy-mm-tt hh:nn:ss "); FileSource; _
            " Error "; CStr(ErrNumber); ": "; ErrDescription
   Close #f

End Sub

Private Function GetApplicationVbProjectName() As String

   Static VbProjectName As String

   Dim DbFile As String
   Dim vbp As Object

On Error Resume Next

   If Len(VbProjectName) = 0 Then
      VbProjectName = Access.VBE.ActiveVBProject.Name
      DbFile = CurrentDb.Name
      'Do not use UNCPath => Code module has no dependencies
      If Access.VBE.ActiveVBProject.FileName <> DbFile Then
         For Each vbp In Access.VBE.VBProjects
            If vbp.FileName = DbFile Then
               VbProjectName = vbp.Name
            End If
         Next
      End If
   End If
   GetApplicationVbProjectName = VbProjectName

End Function
