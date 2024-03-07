Attribute VB_Name = "modAccessFunctions"
Option Explicit
Option Compare Text
Option Private Module

Public Enum AcObjectType
   acTable = 0
   acQuery = 1
End Enum

Public Function Nz(ByVal Value As Variant, ByVal ValueIfNull As Variant) As Variant
   If IsNull(Value) Then
      Value = ValueIfNull
   End If
   Nz = Value
End Function
