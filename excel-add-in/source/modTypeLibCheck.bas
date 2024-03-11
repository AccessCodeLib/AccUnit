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
   FilePath = VBA.Left(FilePath, VBA.InStrRev(FilePath, "\"))
   DefaultAccUnitLibFolder = FilePath & "lib"
End Property

Public Sub CheckAccUnitTypeLibFile(ByVal VBProjectRef As VBProject, Optional ByRef ReferenceFixed As Boolean)

   Dim LibPath As String
   Dim LibFile As String
   Dim FileFixed As Boolean
   
   LibPath = GetAccUnitLibPath(True)
   LibFile = LibPath & ACCUNIT_TYPELIB_FILE
   FileTools.CreateDirectory LibPath

   If Not FileTools.FileExists(LibFile) Then
      FileFixed = True
      ExportTlbFile LibFile
   End If

On Error Resume Next
   CheckMissingReference VBProjectRef, ReferenceFixed
   
   ReferenceFixed = ReferenceFixed Or FileFixed

End Sub

Public Function GetAccUnitLibPath(Optional ByVal BackSlashAtEnd As Boolean = False) As String

   Dim LibPath As String
   Dim LibFile As String
   
   With CurrentAccUnitConfiguration
On Error GoTo ErrMissingPath
      LibPath = .AccUnitDllPath
On Error GoTo 0
   End With

   If VBA.Len(LibPath) = 0 Then
      LibPath = DefaultAccUnitLibFolder
   End If
   
   If BackSlashAtEnd Then
      If VBA.Right(LibPath, 1) <> "\" Then
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
      .CreateAppFile ACCUNIT_TYPELIB_FILE, LibFile, "BitInfo", CStr(GetCurrentVbaBitSystem)
   End With
End Sub

Private Sub CheckMissingReference(ByVal VBProjectRef As VBProject, Optional ByRef ReferenceFixed As Boolean)

   Dim AccUnitRefExists As Boolean
   Dim ref As Object

   With VBProjectRef
      For Each ref In .References
         If ref.Name = "AccUnit" Then
            AccUnitRefExists = True
            Exit Sub
         End If
      Next
   End With

   AddAccUnitTlbReference VBProjectRef
   ReferenceFixed = True

End Sub

Private Sub AddAccUnitTlbReference(ByVal VBProjectRef As VBProject)
   VBProjectRef.References.AddFromFile GetAccUnitLibPath(True) & ACCUNIT_TYPELIB_FILE
End Sub

Private Sub RemoveAccUnitTlbReference(ByVal VBProjectRef As VBProject)

   Dim ref As Object

   For Each ref In VBProjectRef.References
      If ref.IsBroken Then
         VBProjectRef.References.Remove ref
      ElseIf ref.Name = "AccUnit" Then
         VBProjectRef.References.Remove ref
         Exit Sub
      End If
   Next

End Sub
