Attribute VB_Name = "AccUnitLoaderRibbonCallBacks"
Option Explicit
Option Compare Text

Public Sub ShowAccUnitLoaderForm()
   With New AccUnitLoaderForm
      .Show 0
   End With
End Sub

Public Sub ShowAccUnitLoaderFormRCB(RibbonControl As Object)
   ShowAccUnitLoaderForm
End Sub

Public Sub AddAccUnitTlbReferenceRCB(RibbonControl As Object)
   AddAccUnitTlbReference
End Sub

Public Sub RemoveAccUnitTlbReferenceRCB(RibbonControl As Object)
   RemoveAccUnitTlbReference
End Sub

Public Sub InsertFactoryModuleRCB(RibbonControl As Object)
   CheckAccUnitTypeLibFile CodeVBProject
   InsertFactoryModule
End Sub

Public Sub ImportTestClassesRCB(RibbonControl As Object)
   CheckAccUnitTypeLibFile CodeVBProject
   ImportTestClasses
End Sub

Public Sub ExportTestClassesRCB(RibbonControl As Object)
   ExportTestClasses
End Sub

Public Sub RemoveTestEnvironmentKeepTestsRCB(RibbonControl As Object)
   RemoveTestEnvironment False
End Sub

Public Sub RemoveTestEnvironmentDelTestsRCB(RibbonControl As Object)
   RemoveTestEnvironment True
End Sub

Public Sub TestSuiteRunAllFromVBProjectRCB(RibbonControl As Object)
   GetAccUnitFactory.DebugPrintTestSuite.AddFromVBProject.Run
   SetFocusToImmeidateWindow
End Sub

Private Sub SetFocusToImmeidateWindow()
   Dim VbeWin As VBIDE.Window
   For Each VbeWin In Application.VBE.Windows
      If VbeWin.Type = vbext_wt_Immediate Then
         If Not VbeWin.Visible Then
            VbeWin.Visible = True
         End If
         VbeWin.SetFocus
         Exit Sub
      End If
   Next
End Sub
