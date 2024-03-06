Attribute VB_Name = "AccUnitLoaderConfigProcedures"
Option Explicit
Option Compare Text

Public Sub AddAccUnitTlbReference()
   RemoveAccUnitTlbReference
   CurrentVbProject.References.AddFromFile CurrentAccUnitConfiguration.AccUnitDllPath & "\AccessCodeLib.AccUnit.tlb"
End Sub

Public Sub RemoveAccUnitTlbReference()

   Dim ref As Reference
   
   With CurrentVbProject
      For Each ref In .References
         If Len(ref.GUID) > 0 Then
         ElseIf ref.Name = "AccUnit" Then
            .References.Remove ref
            Exit Sub
         End If
      Next
   End With
   
End Sub

Public Sub InsertFactoryModule()

   Dim Configurator As AccUnit.Configurator
   
   With New AccUnitLoaderFactory
      Set Configurator = .Configurator
   End With
   
   Configurator.InsertAccUnitLoaderFactoryModule AccUnitTlbReferenceExists, True, CurrentVbProject, Application
   Set Configurator = Nothing
   
On Error Resume Next
'   Application.RunCommand acCmdCompileAndSaveAllModules

End Sub

Private Function AccUnitTlbReferenceExists() As Boolean

   Dim ref As Reference
   
   For Each ref In CurrentVbProject.References
      If Len(ref.GUID) > 0 Then
      ElseIf ref.Name = "AccUnit" Then
         AccUnitTlbReferenceExists = True
         Exit Function
      End If
   Next
   
End Function

Public Sub ImportTestClasses()

   Dim Configurator As AccUnit.Configurator
   
   With New AccUnitLoaderFactory
      Set Configurator = .Configurator
   End With
   
   Configurator.InsertAccUnitLoaderFactoryModule AccUnitTlbReferenceExists, False, CurrentVbProject, Application
   Configurator.ImportTestClasses
   Set Configurator = Nothing
   
On Error Resume Next
'   Application.RunCommand acCmdCompileAndSaveAllModules
   
End Sub

Public Sub ExportTestClasses()

   Dim Configurator As AccUnit.Configurator
   
   With New AccUnitLoaderFactory
      Set Configurator = .Configurator
   End With
   
   Configurator.ExportTestClasses
   Set Configurator = Nothing
   
End Sub

Public Sub RemoveTestEnvironment(ByVal RemoveTestModules As Boolean)

   Dim Configurator As AccUnit.Configurator
   
   With New AccUnitLoaderFactory
      Set Configurator = .Configurator
   End With
   
   Configurator.RemoveTestEnvironment RemoveTestModules, , CurrentVbProject
   Set Configurator = Nothing
   
On Error Resume Next
'   Application.RunCommand acCmdCompileAndSaveAllModules

End Sub
