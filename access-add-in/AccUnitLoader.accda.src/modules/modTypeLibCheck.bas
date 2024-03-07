Attribute VB_Name = "modTypeLibCheck"
'---------------------------------------------------------------------------------------
' Module: modTypeLibCheck
'---------------------------------------------------------------------------------------
'/**
' <summary>
' TypeLib-Referenz setzen
' </summary>
' <remarks>
' </remarks>
' \ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/modTypeLibCheck.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>file/FileTools.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

#Const EARLYBINDING = 1

Private Const EXTENSION_KEY_APPFILE As String = "AppFile"

Public Property Get DefaultAccUnitLibFolder() As String
   Dim FilePath As String
   FilePath = CodeVBProject.FileName
   FilePath = Left(FilePath, InStrRev(FilePath, "\"))
   DefaultAccUnitLibFolder = FilePath & "lib"
End Property

Public Sub CheckAccUnitTypeLibFile(Optional ByVal VBProjectRef As VBProject = Nothing)

   Dim LibPath As String
   Dim LibFile As String

   LibPath = GetAccUnitLibPath(True)
   LibFile = LibPath & ACCUNIT_TYPELIB_FILE
   FileTools.CreateDirectory LibPath

   If Not FileTools.FileExists(LibFile) Then
      ExportTlbFile LibFile
   End If

On Error Resume Next
   If VBProjectRef Is Nothing Then
      Set VBProjectRef = CodeVBProject
   End If

   CheckMissingReference VBProjectRef

End Sub

Private Function GetAccUnitLibPath(Optional ByVal BackSlashAtEnd As Boolean = False) As String

   Dim LibPath As String
   Dim LibFile As String

   With CurrentAccUnitConfiguration
On Error GoTo ErrMissingPath
      LibPath = .AccUnitDllPath
On Error GoTo 0
   End With

   If Len(LibPath) = 0 Then
      LibPath = DefaultAccUnitLibFolder
   End If

   If BackSlashAtEnd Then
      If Right(LibPath, 1) <> "\" Then
         LibPath = LibPath & "\"
      End If
   End If

   GetAccUnitLibPath = LibPath

   Exit Function

ErrMissingPath:
   Resume Next

End Function

Private Sub ExportTlbFile(ByVal LibFile As String)
   With CurrentApplication.Extensions(EXTENSION_KEY_APPFILE)
      .CreateAppFile ACCUNIT_TYPELIB_FILE, LibFile
   End With
End Sub

Private Sub CheckMissingReference(ByVal VBProjectRef As VBProject)

   Dim AccUnitRefExists As Boolean
   Dim ref As Object
   Dim RefName As String

   With VBProjectRef
      For Each ref In .References
On Error Resume Next
         RefName = ref.Name
         If Err.Number <> 0 Then
            Err.Clear
            RefName = vbNullString
         End If
On Error GoTo 0
         If RefName = "AccUnit" Then
            AccUnitRefExists = True
            Exit Sub
         End If
      Next
   End With

   AddAccUnitTlbReference VBProjectRef

End Sub

Private Sub AddAccUnitTlbReference(ByVal VBProjectRef As VBProject)
   VBProjectRef.References.AddFromFile GetAccUnitLibPath(True) & ACCUNIT_TYPELIB_FILE
End Sub

Private Sub RemoveAccUnitTlbReference(ByVal VBProjectRef As VBProject)

   Dim ref As Object
   Dim RefName As String

   For Each ref In VBProjectRef.References
On Error Resume Next
      RefName = ref.Name
      If Err.Number <> 0 Then
         Err.Clear
         RefName = vbNullString
      End If
On Error GoTo 0

      If ref.IsBroken Then
         VBProjectRef.References.Remove ref
      ElseIf RefName = "AccUnit" Then
         VBProjectRef.References.Remove ref
         Exit Sub
      End If
   Next

End Sub
