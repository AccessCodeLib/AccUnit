VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestResultReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Interface: ITestResultReporter
'---------------------------------------------------------------------------------------
'
' Interface for TestResultReporter classes
'
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Public Enum TestReportOutput
   DebugPrint = 1
   LogFile = 2
End Enum

Public Property Set TestResultCollector(ByVal TestResultCollectorToListen As TestResultCollector)
'
End Property

