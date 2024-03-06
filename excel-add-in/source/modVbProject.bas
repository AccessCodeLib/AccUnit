Attribute VB_Name = "modVbProject"
'---------------------------------------------------------------------------------------
' Module: modVbProject
'---------------------------------------------------------------------------------------
'/**
' <summary>
' VBProject ermitteln
' </summary>
' <remarks>
' </remarks>
' \ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/modVbProject.bas</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

#Const EARLYBINDING = 1

Private m_CurrentVbProject As Object

#If EARLYBINDING Then
Public Property Get CurrentVbProject() As VBIDE.VBProject
#Else
Public Property Get CurrentVbProject() As Object
#End If

#If EARLYBINDING Then
   Dim Proj As VBProject
#Else
   Dim Proj As Object
#End If
   Dim strCurrentDbName As String

   If m_CurrentVbProject Is Nothing Then
      Set m_CurrentVbProject = Application.VBE.ActiveVBProject
      If Application.VBE.VBProjects.Count > 1 Then
      'Pr�fen, ob das richtige VbProject gew�hlt wurde (muss das vom aktiven Workbook sein)
         strCurrentDbName = Application.ActiveWorkbook.FullName
         If m_CurrentVbProject.FileName <> strCurrentDbName Then
            Set m_CurrentVbProject = Nothing
            For Each Proj In Application.VBE.VBProjects
               If Proj.FileName = strCurrentDbName Then
                  Set m_CurrentVbProject = Proj
                  Exit For
               End If
            Next
         End If
      End If
   End If

   Set CurrentVbProject = m_CurrentVbProject

End Property

#If EARLYBINDING Then
Public Property Get CodeVBProject() As VBIDE.VBProject
#Else
Public Property Get CodeVBProject() As Object
#End If

   Set CodeVBProject = Application.ThisWorkbook.VBProject

End Property

