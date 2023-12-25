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
Option Compare Database
Option Explicit

#Const EARLYBINDING = 0

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
      Set m_CurrentVbProject = VBE.ActiveVBProject
      'Prüfen, ob das richtige VbProject gewählt wurde (muss das von CurrentDb sein)
      strCurrentDbName = UncPath(CurrentDb.Name)
      If m_CurrentVbProject.FileName <> strCurrentDbName Then
         Set m_CurrentVbProject = Nothing
         For Each Proj In VBE.VBProjects
            If Proj.FileName = strCurrentDbName Then
               Set m_CurrentVbProject = Proj
               Exit For
            End If
         Next
      End If
   End If

   Set CurrentVbProject = m_CurrentVbProject

End Property
