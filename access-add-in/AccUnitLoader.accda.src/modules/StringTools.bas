Attribute VB_Name = "StringTools"
Attribute VB_Description = "String-Hilfsfunktionen"
'---------------------------------------------------------------------------------------
' Package: text.StringTools
'---------------------------------------------------------------------------------------
'
' Text functions
'
' Author:
'     Josef Poetzl, Sten Schmidt
'
' Remarks:
'     Use DisableReplaceVbaStringFunctions = 1 in conditional compilation arguments (in vbe project properties)
'     to disable replacement of VBA.Format function
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>text/StringTools.bas</file>
'  <license>_codelib/license.bas</license>
'  <test>_test/text/StringToolsTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

'---------------------------------------------------------------------------------------
' Enum: TrimOption
'---------------------------------------------------------------------------------------
'
' Available trim options for the trim function
'
'  TrimStart - (1) Remove leading spaces from a string variable
'  TrimEnd   - (2) Remove trailing spaces from a string variable
'  TrimBoth  - (3) Remove leading and trailing spaces
'
'---------------------------------------------------------------------------------------
Public Enum TrimOption
    TrimStart = 1
    TrimEnd = 2
    TrimBoth = 3
End Enum

'---------------------------------------------------------------------------------------
' Function: IsNullOrEmpty
'---------------------------------------------------------------------------------------
'
' Specifies whether the passed value is null or an empty string
'
' Parameters:
'     ValueToTest    - Value to be checked
'     IgnoreSpaces   - Ignore spaces at the beginning and end
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function IsNullOrEmpty(ByVal ValueToTest As Variant, Optional ByVal IgnoreSpaces As Boolean = False) As Boolean

   Dim TempValue As String

   If IsNull(ValueToTest) Then
      IsNullOrEmpty = True
      Exit Function
   End If

   TempValue = CStr(ValueToTest)

   If IgnoreSpaces Then
      TempValue = Trim$(TempValue)
   End If

   IsNullOrEmpty = (Len(TempValue) = 0)

End Function

'---------------------------------------------------------------------------------------
' Function: FormatText
'---------------------------------------------------------------------------------------
'
' Inserts the passed parameters into the placeholder {0..n} of the format text
'
' Parameters:
'     FormatString   - Text format with placeholder ... Example: "XYZ{0}, {1}"
'     Args           - Passing parameters in suitable order
'
' Returns:
'     String
'
'---------------------------------------------------------------------------------------
Public Function FormatText(ByVal FormatString As String, ParamArray Args() As Variant) As String

   Dim Arg As Variant
   Dim Temp As String
   Dim i As Long

#If USELOCALIZATION = 1 Then
   FormatString = L10n.Text(FormatString)
#End If

   Temp = FormatString
   For Each Arg In Args
      Temp = Replace(Temp, "{" & i & "}", CStr(Arg))
      i = i + 1
   Next

   FormatText = Temp

End Function

'---------------------------------------------------------------------------------------
' Function: Format
'---------------------------------------------------------------------------------------
'
' Replaces the VBA format function
' Extension: [h] or [hh] for hour display over 24
'
' Parameters:
'     Expression        - The value to format
'     FormatString      - A valid named or user-defined format expression incl. extension for hours display over 24 (for standard format instructions see VBA.Format)
'     FirstDayOfWeek    - Passed on to VBA.Format
'     FirstWeekOfYear   - Passed on to VBA.Format
'
' Returns:
'     String
'
'---------------------------------------------------------------------------------------
#If DisableReplaceVbaStringFunctions = 0 Then
Public Function Format(ByVal Expression As Variant, Optional ByVal FormatString As Variant, _
              Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
              Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) As String
   Format = FormatX(Expression, FormatString, FirstDayOfWeek, FirstWeekOfYear)
End Function
#End If

Public Function FormatX(ByVal Expression As Variant, Optional ByVal FormatString As Variant, _
              Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
              Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) As String

   Dim Hours As Long

   If Not IsDate(Expression) Then
      If IsNumeric(Expression) Then ' falls Zeitberechnung übergeben wird, ohne auf CDate zu konvertieren
         If InStr(1, Replace(FormatString, "[hh]", "[h]"), "[h]") > 0 Then
            Expression = CDate(Expression)
         End If
      End If
   End If

   If IsDate(Expression) Then
      If InStr(1, FormatString, "[h", vbTextCompare) > 0 Then
         Hours = Fix(Round(CDate(Expression) * 24, 1))
         If Abs(Hours) < 24 Then
            FormatString = Replace(FormatString, "[hh]", "hh", , , vbTextCompare)
            FormatString = Replace(FormatString, "[h]", "h", , , vbTextCompare)
         Else
            FormatString = Replace(FormatString, "[hh]", "[h]", , , vbTextCompare)
            FormatString = Replace(FormatString, "[h]", Replace(CStr(Hours), "0", "\0"), , , vbTextCompare)
         End If
      End If
   End If

   FormatX = VBA.Format$(Expression, FormatString, FirstDayOfWeek, FirstWeekOfYear)

End Function

'---------------------------------------------------------------------------------------
' Function: PadLeft
'---------------------------------------------------------------------------------------
'
' Left padding of a string
'
' Parameters:
'     Value       - String to be filled in
'     TotalWidth  - Total length of the resulting string
'     PadChar     - (optional) Character to be padded with; Default: " "
'
' Returns:
'     String
'
' Remarks:
'     If the length of value is greater than or equal to totalWidth, the result is limited to totalWidth characters
'
'---------------------------------------------------------------------------------------
Public Function PadLeft(ByVal Value As String, ByVal TotalWidth As Integer, Optional ByVal PadChar As String = " ") As String
    PadLeft = VBA.Right$(VBA.String$(TotalWidth, PadChar) & Value, TotalWidth)
End Function

'---------------------------------------------------------------------------------------
' Function: PadRight
'---------------------------------------------------------------------------------------
'
' Right padding of a string
'
' Parameters:
'     Value       - String to be filled in
'     TotalWidth  - Total length of the resulting string
'     PadChar     - (optional) Character to be padded with; Default: " "
'
' Returns:
'     String
'
' Remarks:
'     If the length of value is greater than or equal to totalWidth, the result is limited to totalWidth characters
'
'---------------------------------------------------------------------------------------
Public Function PadRight(ByVal Value As String, ByVal TotalWidth As Integer, Optional ByVal PadChar As String = " ") As String
    PadRight = VBA.Left$(Value & VBA.String$(TotalWidth, PadChar), TotalWidth)
End Function

'---------------------------------------------------------------------------------------
' Function: Contains
'---------------------------------------------------------------------------------------
'
' Indicates whether SearchValue occurs in the CheckValue string
'
' Parameters:
'     CheckValue  - String to be searched
'     SearchValue - String to be searched for
'
' Returns:
'     Boolean
'
' Remarks:
'     Returns True if SearchValue is contained in CheckValue or SearchValue has the value vbNullString
'
'---------------------------------------------------------------------------------------
Public Function Contains(ByVal CheckValue As String, ByVal SearchValue As String) As Boolean
    Contains = VBA.InStr(1, CheckValue, SearchValue, vbTextCompare) > 0
End Function

'---------------------------------------------------------------------------------------
' Function: EndsWith
'---------------------------------------------------------------------------------------
'
' Indicates whether the string CheckValue ends with SearchValue
'
' Parameters:
'     CheckValue  - String to be searched
'     SearchValue - String to be searched for
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function EndsWith(ByVal CheckValue As String, ByVal SearchValue As String) As Boolean
    EndsWith = VBA.Right$(CheckValue, VBA.Len(SearchValue)) = SearchValue
End Function

'---------------------------------------------------------------------------------------
' Function: StartsWith
'---------------------------------------------------------------------------------------
'
' Indicates whether the string CheckValue starts with SearchValue
'
' Parameters:
'     CheckValue  - String to be searched
'     SearchValue - String to be searched for
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function StartsWith(ByVal CheckValue As String, ByVal SearchValue As String) As Boolean
    StartsWith = VBA.Left$(CheckValue, VBA.Len(SearchValue)) = SearchValue
End Function

'---------------------------------------------------------------------------------------
' Function: Length
'---------------------------------------------------------------------------------------
'
' Returns the number of characters in Value
'
' Parameters:
'     Value - String to be checked
'
' Returns:
'     Long - Anzahl Zeichen von Value
'
'---------------------------------------------------------------------------------------
Public Function Length(ByVal Value As String) As Long
    Length = VBA.Len(Value)
End Function

'---------------------------------------------------------------------------------------
' Function: Concat
'---------------------------------------------------------------------------------------
'
' Appends the string ValueB to the string ValueA.
'
' Parameters:
'     ValueA - Base string
'     ValueB - String to be append at end of A
'
' Returns:
'     String - ValueB appended to ValueA
'
'---------------------------------------------------------------------------------------
Public Function Concat(ByVal ValueA As String, ByVal ValueB As String) As String
    Concat = ValueA & ValueB
End Function

'---------------------------------------------------------------------------------------
' Function: Trim
'---------------------------------------------------------------------------------------
'
' Removes leading and/or trailing spaces from a string
'
' Replaces the function VBA.Trim().
'
' Parameters:
'     Value    - String to be trimmed
'     TrimType - Trim options (at start, at end or both)
'
' Returns:
'     String
'
'---------------------------------------------------------------------------------------
Public Function Trim(ByVal Value As String, Optional ByVal TrimType As TrimOption = TrimOption.TrimBoth) As String

    Select Case TrimType
        Case TrimOption.TrimBoth
            Trim = VBA.Trim$(Value)
            Exit Function
        Case TrimOption.TrimStart
            Trim = VBA.LTrim$(Value)
            Exit Function
        Case TrimOption.TrimEnd
            Trim = VBA.RTrim(Value)
            Exit Function
        Case Else
            Trim = Value
            Exit Function
    End Select

End Function

'---------------------------------------------------------------------------------------
' Function: Substring
'---------------------------------------------------------------------------------------
'
' Returns a part of the string Value starting at the position StartIndex and having the length Length.
'
' Parameters:
'     Value       - String
'     StartIndex  - Start position in the string
'     Length      - Number of characters to be returned
'
' Returns:
'     String
'
' Remarks:
'     StartIndex is null terminated, analogous to String.Substring() in .NET
'
'---------------------------------------------------------------------------------------
Public Function SubString(ByVal Value As String, ByVal StartIndex As Long, Optional ByVal Length As Long = 0) As String
    If Length = 0 Then Length = StringTools.Length(Value) - StartIndex
    SubString = VBA.Mid$(Value, StartIndex + 1, Length)
End Function

'---------------------------------------------------------------------------------------
' Function: InsertAt
'---------------------------------------------------------------------------------------
'
' Setzt die Zeichenfolge InsertValue an der Position Pos ein
'
' Parameters:
'     Value       - String
'     InsertValue - String to be inserted
'     Pos         - Position at which the string is to be inserted (Pos is zero-terminated).
'
' Returns:
'     String
'
'---------------------------------------------------------------------------------------
Public Function InsertAt(ByVal Value As String, ByVal InsertValue As String, ByVal Pos As Long) As String
    InsertAt = VBA.Mid$(Value, 1, Pos) & InsertValue & StringTools.SubString(Value, Pos)
End Function

'---------------------------------------------------------------------------------------
' Function: Replicate
'---------------------------------------------------------------------------------------
'
' Repeat string
'
' Parameters:
'     Value    - The string to be repeated
'     Number   - Number of repetitions
'
' Returns:
'     String
'
' Remarks:
'     Replicate("abc", 3) creates "abcabcabc"
'
'---------------------------------------------------------------------------------------
Public Function Replicate(ByVal Value As String, ByVal Number As Long) As String

   Dim ValueLen As Long
   Dim TempString As String
   Dim i As Long

   If Number = 0 Then
      Replicate = vbNullString
      Exit Function
   ElseIf Number = 1 Then
      Replicate = Value
      Exit Function
   End If

   ValueLen = Len(Value)

   If ValueLen = 0 Then
      Replicate = vbNullString
      Exit Function
   ElseIf ValueLen = 1 Then
      Replicate = String$(Number, Value)
      Exit Function
   End If

   TempString = String$(Number * ValueLen, Chr(0))

   For i = 1 To Number
      Mid$(TempString, (i - 1) * ValueLen + 1, ValueLen) = Value
   Next

   Replicate = TempString

End Function
