VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccUnitAssertBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' Notwendig für Dispose-Aufruf

'Implements AccUnit.Assert

Private m_Assert As Object 'AccUnit.Assert

Private Sub Assert_Dispose()
   DisposeAssert
End Sub

Private Property Set Assert_MatchResultCollector(ByVal MatchResultCollector As AccUnit.IMatchResultCollector)
   Set m_Assert.MatchResultCollector = MatchResultCollector
End Property

Private Property Get Assert_MatchResultCollector() As AccUnit.IMatchResultCollector
   Set Assert_MatchResultCollector = m_Assert.MatchResultCollector
End Property

Private Sub Assert_That(ByVal Actual As Variant, ByVal IsConstraint As AccUnit.IConstraintBuilder, Optional ByVal InfoText As String = vbNullString)
   m_Assert.That Actual, IsConstraint, InfoText
End Sub

Private Sub Class_Terminate()
   DisposeAssert
End Sub

Public Property Get This() As AccUnitAssertBridge
   Set This = Me
End Property

Public Sub Init(ByVal NewAssert As AccUnit.Assert, Optional ByVal MatchResultCollector As IMatchResultCollector = Nothing)
   Set m_Assert = NewAssert
   Set m_Assert.MatchResultCollector = MatchResultCollector
End Sub

Private Sub DisposeAssert()
   If Not (m_Assert Is Nothing) Then
      m_Assert.Dispose
      Set m_Assert = Nothing
   End If
End Sub

Public Sub That(ByVal Actual As Variant, ByVal IsConstraint As AccUnit.ConstraintBuilder, _
       Optional ByVal InfoText As String = vbNullString)
   m_Assert.That Actual, IsConstraint, InfoText
End Sub


