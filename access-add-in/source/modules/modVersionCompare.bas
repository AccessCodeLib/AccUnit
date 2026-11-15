Attribute VB_Name = "modVersionCompare"
'---------------------------------------------------------------------------------------
' Module: modVersionCompare
'---------------------------------------------------------------------------------------
'
' Compare version strings, e.g. file versions ("0.10.0.0") or date versions ("2026.11.12")
'
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit
Option Private Module

Private Const VersionPartSeparator As String = "."

' Returns 1 if Version1 is newer than Version2, -1 if it is older and 0 if both are equal.
' Parts are compared numerically from left to right; missing parts count as 0.
Public Function CompareVersions(ByVal Version1 As String, ByVal Version2 As String) As Long

   Dim Version1Parts() As String
   Dim Version2Parts() As String
   Dim LastPartIndex As Long
   Dim PartIndex As Long
   Dim PartDifference As Double

   Version1Parts = VBA.Split(Version1, VersionPartSeparator)
   Version2Parts = VBA.Split(Version2, VersionPartSeparator)

   LastPartIndex = UBound(Version1Parts)
   If UBound(Version2Parts) > LastPartIndex Then
      LastPartIndex = UBound(Version2Parts)
   End If

   For PartIndex = 0 To LastPartIndex
      PartDifference = VersionPartValue(Version1Parts, PartIndex) - VersionPartValue(Version2Parts, PartIndex)
      If PartDifference <> 0 Then
         CompareVersions = VBA.Sgn(PartDifference)
         Exit Function
      End If
   Next

   CompareVersions = 0

End Function

Private Function VersionPartValue(ByRef VersionParts() As String, ByVal PartIndex As Long) As Double
   If PartIndex <= UBound(VersionParts) Then
      VersionPartValue = VBA.Val(VersionParts(PartIndex))
   End If
End Function
