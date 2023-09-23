VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Klassenmodul: TestSummary
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Test-Summary
' </summary>P
' <remarks>
' </remarks>
'\ingroup addins_shared
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/TestSummary.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_Failed As Long
Private m_Ignored As Long
Private m_Passed As Long
Private m_Total As Long

Public Sub Reset()
   m_Failed = 0
   m_Ignored = 0
   m_Passed = 0
   m_Total = 0
End Sub

Public Sub SaveResult(ByVal TestResult As AccUnit.ITestResult)
   With TestResult
      If .IsIgnored Then
         m_Ignored = m_Ignored + 1
      ElseIf .IsFailure Or .IsError Then
         m_Failed = m_Failed + 1
      ElseIf .IsPassed Then
         m_Passed = m_Passed + 1
      End If
   End With
   m_Total = m_Total + 1
End Sub

Public Property Get ElapsedTime() As Double
   ElapsedTime = 0
End Property

Public Property Get Failed() As Long
   Failed = m_Failed
End Property

Public Property Get Ignored() As Long
   Ignored = m_Ignored
End Property

Public Property Get Passed() As Long
   Passed = m_Passed
End Property

Public Property Get Total() As Long
   Total = m_Total
End Property
