Attribute VB_Name = "_Example"
Option Compare Database
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Public Sub RunCodeCoverageTestExample()

   'CodeCoverage config:
   With CodeCoverageTest("ExampleClass")
      'TestSuite:
      With .AddByClassName("ExampleClassTests")
         'Run tests:
         .Run
      End With

   End With

   'wait to see CodeCoverageTracker code:
   'Sleep 5000

   Debug.Print CodeCoverageTracker.GetReport("*", "Method2", True)

   'Remove CodeCoverageTracker code:
   CodeCoverageTracker.Clear "ExampleClass"

End Sub

Private Sub TestRef()

   Dim r As Reference
   For Each r In Application.References
      Debug.Print r.Name, r.IsBroken, r.Guid
   Next

End Sub
