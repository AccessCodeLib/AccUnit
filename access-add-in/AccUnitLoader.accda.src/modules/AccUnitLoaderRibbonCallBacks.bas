Attribute VB_Name = "AccUnitLoaderRibbonCallBacks"
Option Explicit
Option Compare Text

Public Sub ShowAccUnitLoaderForm()
   StartApplication
End Sub

Public Sub ShowAccUnitLoaderFormRCB(Optional RibbonControl As Object)
   ShowAccUnitLoaderForm
End Sub

Public Sub AddAccUnitTlbReferenceRCB(Optional RibbonControl As Object)
   AddAccUnitTlbReference
End Sub

Public Sub RemoveAccUnitTlbReferenceRCB(Optional RibbonControl As Object)
   RemoveAccUnitTlbReference
End Sub

Public Sub InsertFactoryModuleRCB(Optional RibbonControl As Object)
   CheckAccUnitTypeLibFile CodeVBProject
   InsertFactoryModule
End Sub

Public Sub ImportTestClassesRCB(Optional RibbonControl As Object)
   CheckAccUnitTypeLibFile CodeVBProject
   ImportTestClasses
End Sub

Public Sub ExportTestClassesRCB(Optional RibbonControl As Object)
   ExportTestClasses
End Sub

Public Sub RemoveTestEnvironmentKeepTestsRCB(Optional RibbonControl As Object)
   RemoveTestEnvironment False
End Sub

Public Sub RemoveTestEnvironmentDelTestsRCB(Optional RibbonControl As Object)
   RemoveTestEnvironment True
End Sub

Public Sub TestSuiteRunAllFromVBProjectRCB(Optional RibbonControl As Object)
   GetAccUnitFactory.DebugPrintTestSuite.AddFromVBProject.Run
   SetFocusToImmediateWindow
End Sub

Private Sub SetFocusToImmediateWindow()
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
