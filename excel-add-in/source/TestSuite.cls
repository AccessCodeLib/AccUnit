VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: TestSuite
'---------------------------------------------------------------------------------------
'
' AccUnitLoader TestSuite for Excel
'
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Implements AccUnit.IVBATestSuite

Private WithEvents m_TestSuite As AccUnit.VbaTestSuite
Attribute m_TestSuite.VB_VarHelpID = -1
Private m_TestResultCollector As TestResultCollector
Private m_CodeCoverageTracker As AccUnit.CodeCoverageTracker

Private m_TestResultReporters As Collection

Private Sub Class_Initialize()
   Set m_TestResultReporters = New Collection
End Sub

Private Sub Class_Terminate()
   m_TestSuite.Dispose
   Set m_TestSuite = Nothing
End Sub

Public Sub Init(ByVal NewTestSuite As AccUnit.VbaTestSuite, _
                ByVal NewTestResultCollector As TestResultCollector)
   Set m_TestSuite = NewTestSuite
   Set m_TestResultCollector = NewTestResultCollector
   Set m_TestSuite.TestResultCollector = m_TestResultCollector
End Sub

Public Sub AddTestResultReporter(ByVal Reporter As ITestResultReporter)
   Set Reporter.TestResultCollector = m_TestResultCollector
   m_TestResultReporters.Add Reporter
End Sub

Public Property Get TestSuite() As AccUnit.VbaTestSuite
   Set TestSuite = m_TestSuite
End Property

Private Function IVBATestSuite_AppendTestResultReporter(ByVal Reporter As AccUnit.ITestResultReporter) As AccUnit.IVBATestSuite
   m_TestSuite.AppendTestResultReporter Reporter
End Function

Private Sub m_TestSuite_TestSuiteFinished(ByVal Summary As AccUnit.ITestSummary)
    m_TestResultCollector.TestSuiteFinished Summary
End Sub

Private Sub m_TestSuite_TestSuiteStarted(ByVal TestSuite As AccUnit.ITestSuite)
   m_TestResultCollector.TestSuiteStarted TestSuite
End Sub

Private Sub m_TestSuite_TestTraceMessage(ByVal Message As String, ByVal CodeConverageTracker As CodeCoverageTracker)
   m_TestResultCollector.PrintTestTraceMessage Message, CodeConverageTracker
End Sub

Public Sub PrintSummary(Optional ByVal PrintTestResults As Boolean = False)
   m_TestResultCollector.PrintSummary m_TestSuite.Summary, PrintTestResults
End Sub

Public Property Set CodeCoverageTracker(ByVal NewRef As Object)
   Set m_CodeCoverageTracker = NewRef
End Property

'####################################
' IVBATestSuite
'
Private Function IVBATestSuite_Reset(Optional ByVal mode As AccUnit.ResetMode = 1&) As AccUnit.IVBATestSuite
   m_TestSuite.Reset mode
   Set IVBATestSuite_Reset = Me
End Function

Private Property Set IVBATestSuite_ActiveVBProject(ByVal RHS As VBIDE.VBProject)
   Set m_TestSuite.ActiveVBProject = RHS
End Property

Private Property Get IVBATestSuite_ActiveVBProject() As VBIDE.VBProject
   Set IVBATestSuite_ActiveVBProject = m_TestSuite.ActiveVBProject
End Property

Public Function Add(ByVal testToAdd As Object) As TestSuite
   m_TestSuite.Add testToAdd
   Set Add = Me
End Function

Private Function IVBATestSuite_Add(ByVal testToAdd As Object) As AccUnit.IVBATestSuite
   Set IVBATestSuite_Add = Add(testToAdd)
End Function

Public Function AddByClassName(ByVal ClassName As String) As TestSuite
   m_TestSuite.AddByClassName ClassName
   Set AddByClassName = Me
End Function

Private Function IVBATestSuite_AddByClassName(ByVal ClassName As String) As AccUnit.IVBATestSuite
   Set IVBATestSuite_AddByClassName = AddByClassName(ClassName)
End Function

Public Function AddFromVBProject() As TestSuite
   m_TestSuite.AddFromVBProject
   Set AddFromVBProject = Me
End Function

Private Function IVBATestSuite_AddFromVBProject() As AccUnit.IVBATestSuite
   Set IVBATestSuite_AddFromVBProject = AddFromVBProject()
End Function

Private Sub IVBATestSuite_Dispose()
   m_TestSuite.Dispose
End Sub

Private Property Set IVBATestSuite_HostApplication(ByVal RHS As Object)
   Set m_TestSuite.HostApplication = RHS
End Property

Private Property Get IVBATestSuite_HostApplication() As Object
   Set IVBATestSuite_HostApplication = m_TestSuite.HostApplication
End Property

Private Property Get IVBATestSuite_Name() As String
   IVBATestSuite_Name = "TestSuite"
End Property

Public Function Run() As AccUnit.IVBATestSuite
   m_TestSuite.Run
   Set Run = Me
End Function

Private Function IVBATestSuite_Run() As AccUnit.IVBATestSuite
   Set IVBATestSuite_Run = Run()
End Function

Public Function SelectTests(ByVal TestNameFilter As Variant) As TestSuite
   m_TestSuite.SelectTests TestNameFilter
   Set SelectTests = Me
End Function

Private Function IVBATestSuite_SelectTests(ByVal TestNameFilter As Variant) As AccUnit.IVBATestSuite
   Set IVBATestSuite_SelectTests = SelectTests(TestNameFilter)
End Function

Public Function Filter(ByVal FilterTags As Variant) As TestSuite
   m_TestSuite.Filter FilterTags
   Set Filter = Me
End Function

Private Function IVBATestSuite_Filter(ByVal FilterTags As Variant) As AccUnit.IVBATestSuite
   Set IVBATestSuite_Filter = Filter(FilterTags)
End Function

Private Property Get Summary() As AccUnit.ITestSummary
   Set Summary = m_TestSuite.Summary
End Property

Private Property Get IVBATestSuite_Summary() As AccUnit.ITestSummary
   Set IVBATestSuite_Summary = m_TestSuite.Summary
End Property

Private Property Set IVBATestSuite_TestResultCollector(ByVal RHS As AccUnit.ITestResultCollector)
   Set m_TestSuite.TestResultCollector = RHS
End Property

Private Property Get IVBATestSuite_TestResultCollector() As AccUnit.ITestResultCollector
   Set IVBATestSuite_TestResultCollector = m_TestSuite.TestResultCollector
End Property

Private Property Set IVBATestSuite_TestRunner(ByVal RHS As AccUnit.ITestRunner)
   Set m_TestSuite.TestRunner = RHS
End Property

Private Property Get IVBATestSuite_TestRunner() As AccUnit.ITestRunner
   Set IVBATestSuite_TestRunner = m_TestSuite.TestRunner
End Property

Private Property Get IVBATestSuite_TestClassGenerator() As AccUnit.ITestClassGenerator
   Set IVBATestSuite_TestClassGenerator = m_TestSuite.TestClassGenerator
End Property




