VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Class: TestSuite
'---------------------------------------------------------------------------------------
'
' AccUnitLoader TestSuite for Access
'
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Implements AccUnit.AccessTestSuite

Private WithEvents m_TestSuite As AccUnit.AccessTestSuite
Attribute m_TestSuite.VB_VarHelpID = -1
Private m_CodeCoverageTracker As AccUnit.CodeCoverageTracker

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
   m_TestSuite.Dispose
   Set m_TestSuite = Nothing
End Sub

Public Sub Init(ByVal NewTestSuite As AccUnit.AccessTestSuite)
   Set m_TestSuite = NewTestSuite
End Sub

Public Sub AppendTestResultReporter(ByVal Reporter As ITestResultReporter)
   m_TestSuite.AppendTestResultReporter Reporter
End Sub

Public Property Get TestSuite() As AccUnit.AccessTestSuite
   Set TestSuite = m_TestSuite
End Property

Private Sub m_TestSuite_TestSuiteStarted(ByVal TestSuite As AccUnit.ITestSuite)
'
End Sub

Private Sub m_TestSuite_TestTraceMessage(ByVal Message As String, ByVal CodeCoverageTracker As AccUnit.ICodeCoverageTracker)
'
End Sub

Private Sub m_TestSuite_TestSuiteFinished(ByVal Summary As AccUnit.ITestSummary)
'
End Sub

Public Property Set CodeCoverageTracker(ByVal NewRef As Object)
   Set m_CodeCoverageTracker = NewRef
End Property

'####################################
' IAccessTestSuite
'
Private Function AccessTestSuite_Reset(Optional ByVal mode As AccUnit.ResetMode = 1&) As AccUnit.IAccessTestSuite
   m_TestSuite.Reset mode
   Set AccessTestSuite_Reset = Me
End Function

Private Property Set AccessTestSuite_ActiveVBProject(ByVal RHS As Object)
   Set m_TestSuite.ActiveVBProject = RHS
End Property

Private Property Get AccessTestSuite_ActiveVBProject() As Object
   Set AccessTestSuite_ActiveVBProject = m_TestSuite.ActiveVBProject
End Property

Public Function Add(ByVal testToAdd As Object) As TestSuite
   m_TestSuite.Add testToAdd
   Set Add = Me
End Function

Private Function AccessTestSuite_Add(ByVal testToAdd As Object) As AccUnit.IAccessTestSuite
   Set AccessTestSuite_Add = Add(testToAdd)
End Function

Public Function AddByClassName(ByVal ClassName As String) As TestSuite
   m_TestSuite.AddByClassName ClassName
   Set AddByClassName = Me
End Function

Private Function AccessTestSuite_AddByClassName(ByVal ClassName As String) As AccUnit.IAccessTestSuite
   Set AccessTestSuite_AddByClassName = AddByClassName(ClassName)
End Function

Public Function AddFromVBProject() As TestSuite
   m_TestSuite.AddFromVBProject
   Set AddFromVBProject = Me
End Function

Private Function AccessTestSuite_AddFromVBProject() As AccUnit.IAccessTestSuite
   Set AccessTestSuite_AddFromVBProject = AddFromVBProject()
End Function

Private Sub AccessTestSuite_Dispose()
   m_TestSuite.Dispose
End Sub

Private Property Set AccessTestSuite_HostApplication(ByVal RHS As Object)
   Set m_TestSuite.HostApplication = RHS
End Property

Private Property Get AccessTestSuite_HostApplication() As Object
   Set AccessTestSuite_HostApplication = m_TestSuite.HostApplication
End Property

Public Property Get Name() As String
   Name = "TestSuite"
End Property

Private Property Get AccessTestSuite_Name() As String
   AccessTestSuite_Name = "AccUnitLoader TestSuite"
End Property

Public Function Run() As AccUnit.IAccessTestSuite
   m_TestSuite.Run
   Set Run = Me
End Function

Private Function AccessTestSuite_Run() As AccUnit.IAccessTestSuite
   Set AccessTestSuite_Run = Run()
End Function

Public Function SelectTests(ByVal TestNameFilter As Variant) As TestSuite
   m_TestSuite.SelectTests TestNameFilter
   Set SelectTests = Me
End Function

Private Function AccessTestSuite_SelectTests(ByVal TestNameFilter As Variant) As AccUnit.IAccessTestSuite
   Set AccessTestSuite_SelectTests = SelectTests(TestNameFilter)
End Function

Public Function Filter(ByVal FilterTags As Variant) As TestSuite
   m_TestSuite.Filter FilterTags
   Set Filter = Me
End Function

Private Function AccessTestSuite_Filter(ByVal FilterTags As Variant) As AccUnit.IAccessTestSuite
   Set AccessTestSuite_Filter = Filter(FilterTags)
End Function

Private Property Get Summary() As AccUnit.ITestSummary
   Set Summary = m_TestSuite.Summary
End Property

Private Property Get AccessTestSuite_Summary() As AccUnit.ITestSummary
   Set AccessTestSuite_Summary = m_TestSuite.Summary
End Property

Private Property Set AccessTestSuite_TestResultCollector(ByVal RHS As AccUnit.ITestResultCollector)
   Set m_TestSuite.TestResultCollector = RHS
End Property

Private Property Get AccessTestSuite_TestResultCollector() As AccUnit.ITestResultCollector
   Set AccessTestSuite_TestResultCollector = m_TestSuite.TestResultCollector
End Property

Private Function AccessTestSuite_AppendTestResultReporter(ByVal Reporter As AccUnit.ITestResultReporter) As AccUnit.IAccessTestSuite
   AppendTestResultReporter Reporter
End Function

Private Property Set AccessTestSuite_TestRunner(ByVal RHS As AccUnit.ITestRunner)
   Set m_TestSuite.TestRunner = RHS
End Property

Private Property Get AccessTestSuite_TestRunner() As AccUnit.ITestRunner
   Set AccessTestSuite_TestRunner = m_TestSuite.TestRunner
End Property

Private Property Get AccessTestSuite_TestClassGenerator() As AccUnit.ITestClassGenerator
   Set AccessTestSuite_TestClassGenerator = m_TestSuite.TestClassGenerator
End Property
