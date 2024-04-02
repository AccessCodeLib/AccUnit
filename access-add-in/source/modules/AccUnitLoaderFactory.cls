VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccUnitLoaderFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Class: AccUnitLoaderFactory
'---------------------------------------------------------------------------------------
'
' Factory for AccUnit elements
'
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Private m_AccUnitFactory As Object
Private m_MatchResultCollector As Object

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
On Error Resume Next
   If Not (m_MatchResultCollector Is Nothing) Then
      Set m_MatchResultCollector = Nothing
   End If

   If Not (m_AccUnitFactory Is Nothing) Then
      'm_AccUnitFactory.Dispose ' COM-sichtbar machen?
      Set m_AccUnitFactory = Nothing
   End If

End Sub

Private Property Get AccUnitFactory() As Object 'AccUnit.AccUnitFactory
   If m_AccUnitFactory Is Nothing Then
      InitAccUnitFactory
   End If
   Set AccUnitFactory = m_AccUnitFactory
End Property

Private Sub InitAccUnitFactory()

   With New NetComDomain
      .DefaultNameSpace = "AccessCodeLib.AccUnit.Interop"
      .LibraryFile = CurrentAccUnitConfiguration.AccUnitDllPath & ACCUNIT_DLL_FILE
      Set m_AccUnitFactory = .CreateObject("AccUnitFactory")
   End With

End Sub

Public Sub Init(ByVal AssertMatchResultCollector As Object)
   Set m_MatchResultCollector = AssertMatchResultCollector
End Sub

Public Function Assert() As Object ' AccUnit.Assert

   Set Assert = AccUnitFactory.Assert
   If Not (m_MatchResultCollector Is Nothing) Then
      Set Assert.MatchResultCollector = m_MatchResultCollector
   End If

End Function

Public Property Get ConstraintBuilder() As Object ' AccUnit.ConstraintBuilder
    Set ConstraintBuilder = AccUnitFactory.ConstraintBuilder
End Property

Public Property Get Configurator() As Object 'AccUnit.Configurator
   Set Configurator = AccUnitFactory.Configurator(CurrentVbProject)
End Property

Public Property Get AccessTestSuite() As Object 'AccUnit.AccessTestSuite
   Set AccessTestSuite = AccUnitFactory.AccessTestSuite(Application)
End Property

Public Property Get TestSuite(Optional ByVal TestReportOutputTo As TestReportOutput = TestReportOutput.DebugPrint) As Object 'TestSuite
   Set TestSuite = New TestSuite

   With TestSuite
      .Init AccessTestSuite
   End With

   AppendTestResultReporter TestSuite, TestReportOutputTo

End Property

Private Sub AppendTestResultReporter(ByVal TestSuite As TestSuite, ByVal TestReportOutputTo As TestReportOutput)

   If (TestReportOutputTo And TestReportOutput.DebugPrint) = TestReportOutput.DebugPrint Then
      TestSuite.AppendTestResultReporter New DebugPrintTestResultReporter
   End If

   If (TestReportOutputTo And TestReportOutput.LogFile) = TestReportOutput.LogFile Then
      TestSuite.AppendTestResultReporter New LogFileTestResultReporter
   End If

   If (TestReportOutputTo And TestReportOutput.MsAccessVCS) = TestReportOutput.MsAccessVCS Then
      TestSuite.AppendTestResultReporter New MsAccessVcsTestResultReporter
   End If

End Sub

Public Property Get CodeCoverageTracker() As Object 'AccUnit.CodeCoverageTracker
    Set CodeCoverageTracker = AccUnitFactory.CodeCoverageTracker(CurrentVbProject)
End Property

Public Property Get ErrorTrappingObserver() As AccUnit.AccessErrorTrappingObserver
    Set ErrorTrappingObserver = AccUnitFactory.AccessErrorTrappingObserver(Application)
End Property
