VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: api.winapi.WinApiFileInfo
'---------------------------------------------------------------------------------------
'
' Read file information with Win-API
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/WinApiFileInfo.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
' Code based on http://support.microsoft.com/kb/509493/
'
Option Compare Text
Option Explicit

#If VBA7 Then

Private Declare PtrSafe Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" ( _
  ByVal lptstrFilename As String, _
  lpdwHandle As LongPtr) As Long

Private Declare PtrSafe Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" ( _
  ByVal lptstrFilename As String, _
  ByVal dwHandle As LongPtr, _
  ByVal dwLen As LongPtr, _
  lpData As Any _
  ) As Long

Private Declare PtrSafe Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" ( _
  pBlock As Any, _
  ByVal lpSubBlock As String, _
  lplpBuffer As Any, _
  puLen As Long _
  ) As Long

Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
  Dest As Any, _
  ByVal Source As LongPtr, _
  ByVal Length As Long)

#Else

Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" ( _
  ByVal lptstrFilename As String, _
  lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" ( _
  ByVal lptstrFilename As String, _
  ByVal dwHandle As Long, _
  ByVal dwLen As Long, _
  lpData As Any _
  ) As Long

Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" ( _
  pBlock As Any, _
  ByVal lpSubBlock As String, _
  lplpBuffer As Any, _
  puLen As Long _
  ) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
  Dest As Any, _
  ByVal Source As Long, _
  ByVal Length As Long)

#End If

Private Type VS_FIXEDFILEINFO
  dwSignature As Long
  dwStrucVersion As Long
  dwFileVersionMS As Long
  dwFileVersionLS As Long
  dwProductVersionMS As Long
  dwProductVersionLS As Long
  dwFileFlagsMask As Long
  dwFileFlags As Long
  dwFileOS As Long
  dwFileType As Long
  dwFileSubtype As Long
  dwFileDateMS As Long
  dwFileDateLS As Long
End Type

Private Type FILEINFOOUT
  FileVersion As String
  ProductVersion As String
End Type

'---------------------------------------------------------------------------------------
' Function: GetFileVersion
'---------------------------------------------------------------------------------------
'
' Determines the version of a file
'
' Parameters:
'     FilePath - full path to the file
'
' Returns:
'     Version identifier
'
' Remarks:
'     Useful for reading versions from dll files
'
'--------------------------------------------------------------------------------------
Public Function GetFileVersion(ByVal FilePath As String) As String
   Dim VerInfo As FILEINFOOUT
   If GetVersion(FilePath, VerInfo) Then
      GetFileVersion = VerInfo.FileVersion
   Else
      GetFileVersion = vbNullString
   End If
End Function

Private Function GetVersion(ByVal FilePath As String, _
                            ByRef GetFileInfo As FILEINFOOUT) As Boolean

  Dim Ret As Long, Size As Long, Handle As LongPtr
  Dim VerBufLen As Long, VerPointer As LongPtr
  Dim FileInfo As VS_FIXEDFILEINFO
  Dim BufferString() As Byte

  Size = GetFileVersionInfoSize(FilePath, Handle)
  If Size = 0 Then
    GetVersion = False
    Exit Function
  End If
  
  ReDim BufferString(Size)
  Ret = GetFileVersionInfo(FilePath, 0&, Size, BufferString(0))
  If Ret = 0 Then
    GetVersion = False
    Exit Function
  End If
  
  Ret = VerQueryValue(BufferString(0), "\", VerPointer, VerBufLen)
  If Ret = 0 Then
    GetVersion = False
    Exit Function
  End If
  
  Call MoveMemory(FileInfo, VerPointer, Len(FileInfo))
  
  With FileInfo
  
    GetFileInfo.FileVersion = _
      Trim$(Str$((.dwFileVersionMS And &HFFFF0000) \ &H10000)) & "." & _
      Trim$(Str$(.dwFileVersionMS And &HFFFF&)) & "." & _
      Trim$(Str$((.dwFileVersionLS And &HFFFF0000) \ &H10000)) & "." & _
      Trim$(Str$(.dwFileVersionLS And &HFFFF&))
    
    GetFileInfo.ProductVersion = _
      Trim$(Str$((.dwProductVersionMS And &HFFFF0000) \ &H10000)) & "." & _
      Trim$(Str$(.dwProductVersionMS And &HFFFF&)) & "." & _
      Trim$(Str$((.dwProductVersionLS And &HFFFF0000) \ &H10000)) & "." & _
      Trim$(Str$(.dwProductVersionLS And &HFFFF&))
      
  End With
  
  GetVersion = True

End Function
