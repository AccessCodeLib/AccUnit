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
   Dim ExportFile As Boolean
   Dim FileFixed As Boolean
   
   LibPath = GetAccUnitLibPath(True)
   LibFile = LibPath & ACCUNIT_TYPELIB_FILE
   FileTools.CreateDirectory LibPath

   ExportFile = Not FileTools.FileExists(LibFile)
   If Not ExportFile Then
      If Not CheckAccUnitVersion(LibFile) Then
         RemoveAccUnitTlbReference VBProjectRef
         ExportFile = True
      End If
   End If
   
   If ExportFile Then
      FileFixed = True
      ExportTlbFile LibFile
   End If

On Error Resume Next
   If VBProjectRef Is Nothing Then
      Set VBProjectRef = CodeVBProject
   End If

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
   ReferenceFixed = True

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

Private Function CheckAccUnitVersion(ByVal AccUnitTlbFilePath As String) As Boolean
   
   Dim AccUnitDllPath As String
   
   AccUnitDllPath = VBA.Replace(AccUnitTlbFilePath, ".tlb", ".dll")
   
   If FileTools.FileExists(AccUnitDllPath) Then
      CheckAccUnitVersion = CheckDllVersion(AccUnitDllPath)
      Exit Function
   End If
   
   CheckAccUnitVersion = CheckTlbVersion(AccUnitTlbFilePath)
   
End Function

Private Function CheckDllVersion(ByVal AccUnitDllFilePath As String) As Boolean
   
   Dim InstalledFileVersion As String
   Dim SourceTableFileVersion As String

   With New WinApiFileInfo
      InstalledFileVersion = .GetFileVersion(AccUnitDllFilePath)
   End With
   
   With CurrentApplication.Extensions(EXTENSION_KEY_APPFILE)
      SourceTableFileVersion = .GetStoredAppFileVersion(ACCUNIT_DLL_FILE, "BitInfo", VBA.CStr(GetCurrentVbaBitSystem))
   End With
   
   CheckDllVersion = (CompareVersions(InstalledFileVersion, SourceTableFileVersion) >= 0)
   
End Function

Private Function CheckTlbVersion(ByVal AccUnitTlbFilePath As String) As Boolean
   
   Dim InstalledFileVersion As String
   Dim SourceTableFileVersion As String

   InstalledFileVersion = VBA.Format(VBA.FileDateTime(AccUnitTlbFilePath), "yyyy\.mm\.dd")
   
   With CurrentApplication.Extensions(EXTENSION_KEY_APPFILE)
      SourceTableFileVersion = .GetStoredAppFileVersion(ACCUNIT_TYPELIB_FILE, "BitInfo", VBA.CStr(GetCurrentVbaBitSystem))
   End With
   
   CheckTlbVersion = (CompareVersions(InstalledFileVersion, SourceTableFileVersion) >= 0)
   
End Function

Private Function CompareVersions(ByVal Version1 As String, ByVal Version2 As String) As Long

   Dim Version1Parts() As String
   Dim Version2Parts() As String
   Dim i As Long

   If VBA.StrComp(Version1, Version2, vbTextCompare) = 0 Then
      CompareVersions = 0
      Exit Function
   End If
   
   If Len(Version1) = 0 Then
      CompareVersions = -1
      Exit Function
   ElseIf Len(Version2) = 0 Then
      CompareVersions = 1
      Exit Function
   End If

   Version1Parts = VBA.Split(Version1, ".")
   Version2Parts = VBA.Split(Version2, ".")

   For i = 0 To UBound(Version1Parts)
      If VBA.Val(Version1Parts(i)) > VBA.Val(Version2Parts(i)) Then
         CompareVersions = 1
         Exit Function
      End If
   Next
   
   CompareVersions = -1

End Function

