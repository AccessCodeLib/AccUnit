VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugPrintMatchResultCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Klassenmodul: DebugPrintMatchResultCollector
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Debug.Print-Ausgabe von Assert.That
' </summary>
' <remarks>
' </remarks>
'\ingroup addins_shared
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/DebugPrintMatchResultCollector.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Implements AccUnit.IMatchResultCollector

Private m_UseRaiseErrorForFailedMatch As Boolean

Private m_IgnorePassedText As Boolean
Private m_FailedCounter As Long
Private m_PassedCounter As Long

Public Property Get This() As DebugPrintMatchResultCollector
   Set This = Me
End Property

Public Sub Reset()
   m_FailedCounter = 0
   m_PassedCounter = 0
End Sub

Public Sub Init(ByVal ShowPassedText As Boolean, Optional ByVal UseRaiseErrorForFailedMatch As Boolean = True)
   m_IgnorePassedText = Not ShowPassedText
   m_UseRaiseErrorForFailedMatch = UseRaiseErrorForFailedMatch
   Reset
End Sub

Public Sub PrintSummary(Optional ByVal ResetCounter As Boolean = False)

   Debug.Print String(40, "-")
   Debug.Print "Tests: "; m_PassedCounter + m_FailedCounter;
   Debug.Print , "Passed: "; m_PassedCounter
   Debug.Print , "Failed: "; m_FailedCounter
   Debug.Print String(40, "-")
      
   If ResetCounter Then
      Reset
   End If
   
End Sub

Private Sub IMatchResultCollector_Add(ByVal Result As AccUnit.IMatchResult, _
                             Optional ByVal InfoText As String = vbNullString)
                             
   If Result.Match Then
      m_PassedCounter = m_PassedCounter + 1
      If Not m_IgnorePassedText Then
         Debug.Print , "passed", InfoText
      End If
   Else
      m_FailedCounter = m_FailedCounter + 1
      
      Debug.Print , "failed", Result.Text, "Expected: " & FormatValues(Result.Expected) & " but was: " & FormatValues(Result.Actual), InfoText
      ' Debug.Print InfoText, "failed", result.FormattedText
   End If
   
End Sub

Private Function FormatValues(ByVal Value2Format As Variant) As String
   
   If VarType(Value2Format) = vbString Then
      FormatValues = """" & Value2Format & """"
   Else
      FormatValues = Value2Format
   End If
   
End Function

Private Property Get IMatchResultCollector_IgnoreFailedMatchAfterAdd() As Boolean
   IMatchResultCollector_IgnoreFailedMatchAfterAdd = Not m_UseRaiseErrorForFailedMatch
End Property
