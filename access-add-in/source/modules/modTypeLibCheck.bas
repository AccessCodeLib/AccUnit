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
Option Compare Database
Option Explicit

#Const EARLYBINDING = 0

Private Const EXTENSION_KEY_APPFILE As String = "AppFile"

Public Property Get DefaultAccUnitLibFolder() As String
   DefaultAccUnitLibFolder = CodeProject.Path & "\lib"
End Property

Public Sub CheckAccUnitTypeLibFile()

   Dim LibPath As String
   Dim LibFile As String

   LibPath = DefaultAccUnitLibFolder & "\"
   LibFile = LibPath & ACCUNIT_TYPELIB_FILE

   FileTools.CreateDirectory LibPath

   If Not FileTools.FileExists(LibFile) Then
      ExportTlbFile LibFile
   End If


On Error Resume Next
   CheckMissingReference

End Sub

Private Sub ExportTlbFile(ByVal LibFile As String)
   With CurrentApplication.Extensions(EXTENSION_KEY_APPFILE)
      .CreateAppFile ACCUNIT_TYPELIB_FILE, LibFile
   End With
End Sub

Private Sub CheckMissingReference()

   Dim AccUnitRefExists As Boolean
   Dim ref As Object

   With CodeDbProject
      For Each ref In .References
         If ref.Name = "AccUnit" Then
            AccUnitRefExists = True
            Exit Sub
         End If
      Next
   End With

   AddAccUnitTlbReference

End Sub

Private Sub AddAccUnitTlbReference()
   CodeDbProject.References.AddFromFile CodeProject.Path & "\lib\" & ACCUNIT_TYPELIB_FILE
End Sub

Private Sub RemoveAccUnitTlbReference()

   Dim ref As Object

   For Each ref In CodeDbProject.References
      If ref.IsBroken Then
         CodeDbProject.References.Remove ref
      ElseIf ref.Name = "AccUnit" Then
         CodeDbProject.References.Remove ref
         Exit Sub
      End If
   Next

End Sub

#If EARLYBINDING Then
Private Property Get CodeDbProject() As VBIDE.VBProject
#Else
Private Property Get CodeDbProject() As Object
#End If

#If EARLYBINDING Then
   Dim Proj As VBProject
#Else
   Dim Proj As Object
#End If
   Dim strCodeDbName As String
   Dim objCodeVbProject As Object

   Set objCodeVbProject = VBE.ActiveVBProject
   'Prüfen, ob das richtige VbProject gewählt wurde (muss das von CodeDb sein)
   strCodeDbName = UncPath(CodeDb.Name)
   If objCodeVbProject.FileName <> strCodeDbName Then
      Set objCodeVbProject = Nothing
      For Each Proj In VBE.VBProjects
         If Proj.FileName = strCodeDbName Then
            Set objCodeVbProject = Proj
            Exit For
         End If
      Next
   End If

   Set CodeDbProject = objCodeVbProject

End Property
