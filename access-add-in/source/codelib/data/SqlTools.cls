VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Attribute VB_PredeclaredId = True
'---------------------------------------------------------------------------------------
' Klasse: SqlTools
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' SQL-Hilfsfunktionen
' </summary>
' <remarks>Attribute VB_PredeclaredId = True einstellen, damit SqlTools ohne explizite Instanzierung genutzt werden kann</remarks>
'
' \warning Nicht vergessen: Parameter für Datumsformat, Boolean und WildCard für das DBMS einstellen
'
' \ingroup data
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/SqlTools.cls</file>
'  <license>_codelib/license.bas</license>
'  <replace>data/SqlTools.bas</replace>
'  <test>_test/data/SqlToolsTests.cls</test>
'  <test>_test/data/SqlToolsBuildCriteriaTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Enum SqlToolsErrorNumbers
   ERRNR_NOCONFIG = vbObjectError + 1
End Enum

' Standard-Werte für Methodenparameter
Private Const SQL_DEFAULT_TEXTDELIMITER As String = "'"
Private Const SQL_DEFAULT_DATEFORMAT As String = ""     ' "" => SqlDateFormat-Eigenschaft wird verwenden.
                                                        '    Zum Deaktivieren Wert eintragen (z. B. "\#yyyy\-mm\-dd\#"),
                                                        '    dann wird dieser Wert als Standardeintrag verwendet.
Private Const SQL_DEFAULT_BOOLTRUESTRING As String = "" ' "" => SqlBooleanTrueString wird verwendet.
                                                        '    Zum Deaktivieren Wert eintragen (z. B. "True oder 1")
Private Const SQL_DEFAULT_WILDCARD As String = "%"      ' * = Standardwert,
                                                        ' benötigte Abweichungen über SqlWildCardString einstellen
                                                   
Private Const SqlAndConcatString As String = " And "
Private Const SqlOrConcatString As String = " Or "

Private m_SqlDateFormat As String
Private m_SqlBooleanTrueString As String
Private m_SqlWildCardString As String

Private Const ResultTextIfNull As String = "Null"

Public Enum SqlRelationalOperators
   [_IgnoreAll] = &H80000000
   SQL_Not = 1
   SQL_Equal = 2
   SQL_LessThan = 4
   SQL_GreaterThan = 8
   SQL_Like = 256
   SQL_Between = 512
   SQL_In = 1024
   SQL_Add_WildCardSuffix = 2048
   SQL_Add_WildCardPrefix = 4096
   SQL_SplitValueToArray = 8192
   SQL_AllowSqlDirect = 16384
   SQL_UseLikeBehavior = 65536
End Enum

Public Enum SqlFieldDataType
   SQL_Boolean = 1
   SQL_Numeric = 2
   SQL_Text = 3
   SQL_Date = 4
End Enum

Public Enum SqlLogicalOperator
   [_SQL_Default] = 0
   SQL_And = 1
   SQL_Or = 2
   SQL_CommaSeparator = 3
End Enum

' SQL-Dialekt-Voreinstellungen
Public Property Get DAO() As SqlTools
   Set DAO = Me.Clone("\#yyyy-mm-dd hh:nn:ss\#", "True", "*")
End Property

Public Property Get TSql() As SqlTools
   Set TSql = Me.Clone("'yyyymmdd hh:nn:ss'", "1", "%")
End Property

' Konfiguration für den SQL-Dialekt
Public Property Get SqlWildCardString() As String
   If Len(m_SqlWildCardString) > 0 Then
      SqlWildCardString = m_SqlWildCardString
   Else
      SqlWildCardString = SQL_DEFAULT_WILDCARD
   End If
End Property

Public Property Let SqlWildCardString(ByVal NewValue As String)
   m_SqlWildCardString = NewValue
End Property

Public Property Get SqlDateFormat() As String
   If Len(m_SqlDateFormat) > 0 Then
      SqlDateFormat = m_SqlDateFormat
   Else
      SqlDateFormat = SQL_DEFAULT_DATEFORMAT
   End If
End Property

Public Property Let SqlDateFormat(ByVal NewValue As String)
   m_SqlDateFormat = NewValue
End Property

Public Property Get SqlBooleanTrueString() As String
   If Len(m_SqlBooleanTrueString) > 0 Then
      SqlBooleanTrueString = m_SqlBooleanTrueString
   Else
      SqlBooleanTrueString = SQL_DEFAULT_BOOLTRUESTRING
   End If
End Property

Public Property Let SqlBooleanTrueString(ByVal NewValue As String)
   m_SqlBooleanTrueString = NewValue
End Property

'---------------------------------------------------------------------------------------
' Function: Clone
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Neue Instanz mit Grundeinstellungen der aktuellen Instanz erstellen.
' </summary>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Clone(Optional ByVal NewSqlDateFormat As String = SQL_DEFAULT_DATEFORMAT, _
                      Optional ByVal NewSqlBooleanTrueString As String = SQL_DEFAULT_BOOLTRUESTRING, _
                      Optional ByVal NewSqlWildCardString As String = SQL_DEFAULT_WILDCARD) As SqlTools


   If Len(NewSqlDateFormat) = 0 Then NewSqlDateFormat = Me.SqlDateFormat
   If Len(NewSqlBooleanTrueString) = 0 Then NewSqlBooleanTrueString = Me.SqlBooleanTrueString
   If Len(NewSqlWildCardString) = 0 Then NewSqlWildCardString = Me.SqlWildCardString

   Set Clone = NewInstance(NewSqlDateFormat, NewSqlBooleanTrueString, NewSqlWildCardString)

End Function

Public Function NewInstance(ByVal NewSqlDateFormat As String, _
                            ByVal NewSqlBooleanTrueString As String, _
                            ByVal NewSqlWildCardString As String) As SqlTools
   
   Dim NewInst As SqlTools

   Set NewInst = New SqlTools
   With NewInst
      .SqlDateFormat = NewSqlDateFormat
      .SqlBooleanTrueString = NewSqlBooleanTrueString
      .SqlWildCardString = NewSqlWildCardString
   End With

   Set NewInstance = NewInst

End Function

'---------------------------------------------------------------------------------------
' Function: DateToSqlText
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Werte in String für SQL-Anweisung umwandeln, die per VBA zusammengesetzt wird.
' </summary>
' <param name="Value">Übergabewert</param>
' <param name="FieldDataType">Datentyp des zu konvertierenden Wertes</param>
' <returns>String</returns>
'**/
'---------------------------------------------------------------------------------------
Public Function ConvertToSqlText(ByVal Value As Variant, _
                                 ByVal FieldDataType As SqlFieldDataType) As String
                     
   Select Case FieldDataType
      Case SqlFieldDataType.SQL_Text
         ConvertToSqlText = TextToSqlText(Value)
      Case SqlFieldDataType.SQL_Numeric
         ConvertToSqlText = NumberToSqlText(Value)
      Case SqlFieldDataType.SQL_Date
         ConvertToSqlText = DateToSqlText(Value)
      Case SqlFieldDataType.SQL_Boolean
         ConvertToSqlText = BooleanToSqlText(Value)
      Case Else
         Err.Raise vbObjectError, "SqlTools.ConvertToSqlText", "FieldDataType '" & FieldDataType & "' not supported"
   End Select
   
End Function

'---------------------------------------------------------------------------------------
' Function: TextToSqlText
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Text für SQL-Anweisung aufbereiten.
' </summary>
' <param name="Value">Übergabewert</param>
' <param name="Delimiter">Begrenzungszeichen für Text-Werte. (In den meisten DBMS wird ' als Begrenzungszeichen verwendet.)</param>
' <param name="WithoutLeftRightDelim">Nur Begrenzungszeichnen innerhalb des Werte verdoppeln, Eingrenzung jedoch nicht setzen.</param>
' <returns>String</returns>
' <remarks>
' Beispiel: strSQL = "select ... from tabelle where Feld = " & TextToSqlText("ab'cd")
'        => strSQL = "select ... from tabelle where Feld = 'ab''cd'"
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function TextToSqlText(ByVal Value As Variant, _
                     Optional ByVal Delimiter As String = SQL_DEFAULT_TEXTDELIMITER, _
                     Optional ByVal WithoutLeftRightDelim As Boolean = False) As String
   
   Dim Result As String
   
   If IsNull(Value) Then
      TextToSqlText = ResultTextIfNull
      Exit Function
   End If
   
   Result = Replace$(Value, Delimiter, Delimiter & Delimiter)
   If Not WithoutLeftRightDelim Then
      Result = Delimiter & Result & Delimiter
   End If
   
   TextToSqlText = Result

End Function

'---------------------------------------------------------------------------------------
' Function: DateToSqlText
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datumswert in String für SQL-Anweisung umwandeln, die per VBA zusammengesetzt wird.
' </summary>
' <param name="Value">Übergabewert</param>
' <param name="FormatString">Datumsformat (von DBMS abhängig!)</param>
' <returns>String</returns>
'**/
'---------------------------------------------------------------------------------------
Public Function DateToSqlText(ByVal Value As Variant, _
                     Optional ByVal FormatString As String = SQL_DEFAULT_DATEFORMAT) As String

   If IsNull(Value) Then
      DateToSqlText = ResultTextIfNull
      Exit Function
   End If
   
   If Not IsDate(Value) Then
      Err.Raise vbObjectError, "SqlTools.DateToSqlText", "Der Wert '" & Value & "' vom Parameter Value ist kein Datumswert!"
   End If

   If Len(FormatString) = 0 Then
      FormatString = SqlDateFormat
      If Len(FormatString) = 0 Then
         Err.Raise SqlToolsErrorNumbers.ERRNR_NOCONFIG, "DateToSqlText", "date format is not defined"
      End If
   End If
   
   DateToSqlText = VBA.Format$(Value, FormatString)

End Function

'---------------------------------------------------------------------------------------
' Function: NumberToSqlText
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Zahl für SQL-Text aufbereiten
' </summary>
' <param name="Value">Übergabewert</param>
' <returns>String</returns>
' <remarks>
' Durch Str-Funktion wird . statt , verwendet.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function NumberToSqlText(ByVal Value As Variant) As String

   Dim Result As String

   If IsNull(Value) Then
      NumberToSqlText = ResultTextIfNull
      Exit Function
   End If
   
   Value = ConvertToNumeric(Value)
   
   Result = Trim$(Str$(Value))
   If Left(Result, 1) = "." Then
      Result = "0" & Result
   End If
   
   NumberToSqlText = Result
   
End Function

Private Function ConvertToNumeric(ByVal Value As Variant) As Variant
   
   Const CheckNumber As Double = 1.23
   
   Dim CheckText As String
   
   If IsNull(Value) Then
      ConvertToNumeric = Null
      Exit Function
   ElseIf CStr(Value) = vbNullString Then
      ConvertToNumeric = Null
      Exit Function
   End If
   
   CheckText = CStr(CheckNumber)
   If InStr(1, CheckText, ",") > 0 Then
      If InStr(1, Value, ".") > 0 Then
         Value = Replace(Value, ".", ",")
         Do While Value Like "*,*,*"
            Value = Replace(Value, ",", vbNullString, 1, 1)
         Loop
      End If
   Else
      If InStr(1, Value, ",") > 0 Then
         Value = Replace(Value, ",", ".")
         Do While Value Like "*.*.*"
            Value = Replace(Value, ".", vbNullString, 1, 1)
         Loop
      End If
   End If
   
   ConvertToNumeric = CDbl(Value)
   
End Function

'---------------------------------------------------------------------------------------
' Function: BooleanToSqlText
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Boolean für SQL-Text aufbereiten
' </summary>
' <param name="Value">Übergabewert</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function BooleanToSqlText(ByVal Value As Variant, _
                        Optional ByVal TrueString As String = SQL_DEFAULT_BOOLTRUESTRING) As String

   If IsNull(Value) Then
      BooleanToSqlText = ResultTextIfNull
      Exit Function
   End If

   If Value = True Then
      If Len(TrueString) = 0 Then
         TrueString = SqlBooleanTrueString
         If Len(TrueString) = 0 Then
            Err.Raise SqlToolsErrorNumbers.ERRNR_NOCONFIG, "BooleanToSqlText", "boolean string for true is not defined"
         End If
      End If
      BooleanToSqlText = TrueString
   Else
      BooleanToSqlText = "0"
   End If
   
End Function

'---------------------------------------------------------------------------------------
' Function: BuildCriteria
'---------------------------------------------------------------------------------------
'/**
' <summary>
' SQL-Kriterium erstellen
' </summary>
' <param name="FieldName">Feldname in der Datenquelle, die gefiltert werden soll</param>
' <param name="RelationalOperator">Vergleichsoperator (=, <=, usw.)</param>
' <param name="FilterValue">Filterwert (kann einzelner Wert oder auch Array mit Werten sein)</param>
' <param name="FilterValue2">Optionale 2. Filterwert (für Between)</param>
' <param name="IgnoreValue">Jener Wert, für den keine Filterbedingung erzeugt werden soll. (Array-Übergabe von Werten möglich)</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function BuildCriteria(ByVal FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
                              ByVal RelationalOperator As SqlRelationalOperators, _
                              ByVal FilterValue As Variant, _
                     Optional ByVal FilterValue2 As Variant = Null, _
                     Optional ByVal IgnoreValue As Variant = Null, _
                     Optional ByVal DisableIgnoreNullValue As Boolean = False) As String

   Dim FilterValueString As String
   Dim OperatorString As String
   Dim Criteria As String
   Dim Criteria1 As String
   Dim Criteria2 As String
   Dim TempArr() As String

   If (RelationalOperator And [_IgnoreAll]) = [_IgnoreAll] Then
      Exit Function
   End If
   
   If Not IsArray(FilterValue) Then
     
      If FilterValue = "{NULL}" Or FilterValue = "{LEER}" Or FilterValue = "{EMPTY}" Then
         FilterValue = Null
         DisableIgnoreNullValue = True
      End If
      
      If FilterValue2 = "{NULL}" Or FilterValue2 = "{LEER}" Or FilterValue2 = "{EMPTY}" Then
         FilterValue2 = Null
         DisableIgnoreNullValue = True
      End If
      
      If (RelationalOperator And SQL_AllowSqlDirect) = SQL_AllowSqlDirect Then
         If FilterValue Like "{*@*}" Then ' Idee von Ulrich: Anwender schreibt SQL-Ausdruck
            Criteria = Replace(Mid(FilterValue, 2, Len(FilterValue) - 2), "@", FieldName)
            BuildCriteria = Criteria
            Exit Function
         End If
      End If
      
   End If
   
   If NullFilterOrEmptyFilter(FieldName, FieldDataType, RelationalOperator, Nz(FilterValue, FilterValue2), IgnoreValue, Criteria, DisableIgnoreNullValue) Then
      BuildCriteria = Criteria
      Exit Function
   End If
   
   If (RelationalOperator And SQL_SplitValueToArray) = SQL_SplitValueToArray Then
      If InStr(1, FilterValue, ";") > 0 Then
         FilterValue = Split(CharTrim(FilterValue, ";"), ";")
      End If
      RelationalOperator = RelationalOperator Xor SQL_SplitValueToArray
   End If
   
   'Sonderfälle:
   If Not IsArray(FilterValue) Then
     
      If FieldDataType = SQL_Numeric Then
      
         If FilterValue = "*" And RelationalOperator = SQL_Equal Then
            BuildCriteria = BuildCriteria(FieldName, FieldDataType, SQL_Not, Null, Null, 0, True)
            Exit Function
         End If
         
         If IsNull(FilterValue2) Then
         
            FilterValue = Trim(FilterValue)
            
            If FilterValue Like "[0-9]*..*[0-9]*" Or FilterValue Like "[+-][0-9]*..*[0-9]*" Then
               TempArr = Split(FilterValue, "..")
               BuildCriteria = BuildCriteria(FieldName, FieldDataType, SQL_Between, Trim(TempArr(0)), Trim(TempArr(1)), IgnoreValue, DisableIgnoreNullValue)
               Exit Function
            End If
         
            If FilterValue Like "[0-9]*-*[0-9]*" Or FilterValue Like "[+-][0-9]*-*[0-9]*" Then    ' convert to a..b
               If Left(FilterValue, 1) = "-" Then
                  FilterValue = "{M}" & Mid(FilterValue, 2)
               End If
               FilterValue = Replace(FilterValue, "  ", " ")
               FilterValue = Replace(FilterValue, "- -", "--")
               FilterValue = Replace(FilterValue, "--", "-{M}")
               FilterValue = Replace(FilterValue, "-", "..")
               FilterValue = Replace(FilterValue, "{M}", "-")
               
               TempArr = Split(FilterValue, "..")
               BuildCriteria = BuildCriteria(FieldName, FieldDataType, SQL_Between, Trim(TempArr(0)), Trim(TempArr(1)), IgnoreValue, DisableIgnoreNullValue)
               Exit Function
               
            End If
         
            If FilterValue Like "*[0-9]" & DecimalMarker & "*[*]" Then
            If (RelationalOperator And SQL_Add_WildCardSuffix) = 0 Then
               BuildCriteria = BuildCriteria(FieldName, FieldDataType, RelationalOperator + SQL_Add_WildCardSuffix, FilterValue, FilterValue2, IgnoreValue, DisableIgnoreNullValue)
               Exit Function
            End If
            End If
         
         End If
         
         ConfigNumericSpecials RelationalOperator, FilterValue, FilterValue2
         
      End If
      
   End If


   If TryBuildInCriteria(FieldName, FieldDataType, RelationalOperator, FilterValue, IgnoreValue, Criteria) Then
      BuildCriteria = Criteria
      Exit Function
   End If

   If TryBuildArrayCriteria(FieldName, FieldDataType, RelationalOperator, FilterValue, IgnoreValue, Criteria) Then
      BuildCriteria = Criteria
      Exit Function
   End If

   If TryBuildBetweenCriteria(FieldName, FieldDataType, RelationalOperator, FilterValue, FilterValue2, IgnoreValue, Criteria) Then
      BuildCriteria = Criteria
      Exit Function
   End If

   If (RelationalOperator And SQL_Like) = SQL_Like Then
      If SqlWildCardString <> "*" Then
         If InStr(1, FilterValue, "*") > 0 Then
            FilterValue = Replace(FilterValue, "[*]", "@@@|||STAR|||@@@")
            FilterValue = Replace(FilterValue, "*", SqlWildCardString)
            FilterValue = Replace(FilterValue, "@@@|||STAR|||@@@", "*")
         End If
      End If
   End If

   If (RelationalOperator And SQL_Add_WildCardSuffix) = SQL_Add_WildCardSuffix Then
      If (RelationalOperator And SQL_Like) = SQL_Like Then
         FilterValue = FilterValue & SqlWildCardString
      ElseIf (FieldDataType And SQL_Date) = SQL_Date Then
         If (RelationalOperator And SQL_LessThan) = 0 Then ' kein < daher:  >, >= oder nur =
            If (RelationalOperator And SQL_GreaterThan) = SQL_GreaterThan Then
               ' nichts ändern => >= DataValue  / SQL_Add_WildCardSuffix ist nicht logisch
            Else ' ganzen Tag berücksichtigen FieldName >= DateValue and FieldName < DateAdd("d", 1, FilterValue))
               BuildCriteria = BuildCriteria(FieldName, FieldDataType, SQL_GreaterThan + SQL_Equal, FilterValue) & _
                               SqlAndConcatString & _
                               BuildCriteria(FieldName, FieldDataType, SQL_LessThan, DateAdd("d", 1, CDate(CLng(FilterValue))))
               Exit Function
            End If
         Else
            If (RelationalOperator And SQL_Equal) = SQL_Equal Then
               RelationalOperator = RelationalOperator - SQL_Equal
            End If
            FilterValue = DateAdd("d", 1, CDate(CLng(FilterValue)))
         End If
      ElseIf (FieldDataType And SQL_Numeric) = SQL_Numeric Then
         If (RelationalOperator And SQL_LessThan) = 0 Then ' kein < daher:  >, >= oder nur =
            If (RelationalOperator And SQL_GreaterThan) = SQL_GreaterThan Then
               If FilterValue Like "*[,.]*[*]" Then
                  FilterValue = Replace(FilterValue, "*", 0)
               ElseIf FilterValue Like "*[*]" Then
                  FilterValue = Replace(FilterValue, "*", vbNullString)
               End If
               ' nichts ändern => >= Zahl  / SQL_Add_WildCardSuffix ist nicht logisch
            Else ' nachfolgende Dezimalwerte berücksichtigen FieldName >= Zahl and FieldName < (Zahl + x)
               If FilterValue Like "-*[*]" Then
                  If FilterValue Like "*[,.]*[*]" Then
                     FilterValue2 = Replace(FilterValue, "*", 0)
                  Else
                     FilterValue2 = Replace(FilterValue, "*", vbNullString)
                  End If
                  Criteria1 = BuildCriteria(FieldName, FieldDataType, SQL_GreaterThan, GetNextDigitNumber(FilterValue, True))
                  Criteria2 = BuildCriteria(FieldName, FieldDataType, SQL_LessThan + SQL_Equal, FilterValue2)
               Else
                  Criteria1 = BuildCriteria(FieldName, FieldDataType, SQL_GreaterThan + SQL_Equal, FilterValue)
                  Criteria2 = BuildCriteria(FieldName, FieldDataType, SQL_LessThan, GetNextDigitNumber(FilterValue))
               End If
               BuildCriteria = Criteria1 & SqlAndConcatString & Criteria2
               Exit Function
            End If
         Else
            If (RelationalOperator And SQL_Equal) = SQL_Equal Then
               RelationalOperator = RelationalOperator - SQL_Equal
            End If
            FilterValue = GetNextDigitNumber(FilterValue)
         End If
      End If
   End If

   If (RelationalOperator And SQL_Add_WildCardPrefix) = SQL_Add_WildCardPrefix Then
      If (RelationalOperator And SQL_Like) = SQL_Like Then
         FilterValue = SqlWildCardString & FilterValue
      End If
   End If

   FilterValueString = ConvertToSqlText(FilterValue, FieldDataType)

   If (RelationalOperator And SQL_Like) = SQL_Like Then
      OperatorString = " Like "
      If (RelationalOperator And SQL_Not) = SQL_Not Then
         OperatorString = " Not" & OperatorString
      End If
      BuildCriteria = FieldName & OperatorString & FilterValueString
      Exit Function
   End If
   
   OperatorString = GetRelationalOperatorString(RelationalOperator)

   Criteria = FieldName & " " & OperatorString & " " & FilterValueString
   If (RelationalOperator And SQL_Not) = SQL_Not Then
      Criteria = "Not " & Criteria
   End If

   BuildCriteria = Criteria

End Function

Private Function ConfigNumericSpecials( _
                              ByRef RelationalOperator As SqlRelationalOperators, _
                              ByRef FilterValue As Variant, _
                              ByRef FilterValue2 As Variant)

   If Left(FilterValue, 1) = "<" Then
      If ((RelationalOperator And SQL_Equal) = SQL_Equal) Then
         RelationalOperator = RelationalOperator - SQL_Equal
      End If
      RelationalOperator = RelationalOperator Or SQL_LessThan
      FilterValue = Mid(FilterValue, 2)
   End If
   
   If Left(FilterValue, 1) = ">" Then
      If ((RelationalOperator And SQL_Equal) = SQL_Equal) Then
         RelationalOperator = RelationalOperator - SQL_Equal
      End If
      RelationalOperator = RelationalOperator Or SQL_GreaterThan
      FilterValue = Mid(FilterValue, 2)
   End If

   If Left(FilterValue, 1) = "=" Then
      RelationalOperator = RelationalOperator Or SQL_Equal
      FilterValue = Mid(FilterValue, 2)
   End If
   
   If Right(FilterValue, 1) = "*" Then
      RelationalOperator = RelationalOperator Or SQL_Add_WildCardSuffix
   ElseIf Right(FilterValue2, 1) = "*" Then
      RelationalOperator = RelationalOperator Or SQL_Add_WildCardSuffix
   End If

End Function

Private Function GetNextDigitNumber(ByVal Z As Variant, Optional AddToAbsoluteValue As Boolean = False) As Double
   
   Dim TestString As String
   Dim KommaPos As Long
   Dim digits As Long
   Dim IsNegativ As Boolean

   Const AdditionalDecDigit As String = "1"
   Const AdditionalDecDigitKorr As Double = 0.1
   
   TestString = Trim(CStr(ConvertToNumeric(Replace(CStr(Z), "*", AdditionalDecDigit))))
   
   If Left(TestString, 1) = "-" And (Not AddToAbsoluteValue) Then
      GetNextDigitNumber = CDbl(Replace(CStr(Z), "*", vbNullString))
      Exit Function
   End If
   
   If Left(TestString, 1) = "-" Then
      IsNegativ = True
   End If
   
   KommaPos = InStrRev(TestString, DecimalMarker)
   If KommaPos = 0 Then ' nächste Ganzzahl
      If AddToAbsoluteValue And IsNegativ Then
         GetNextDigitNumber = CDbl(Replace(CStr(Z), "*", vbNullString)) - 1
      Else
         GetNextDigitNumber = CDbl(Replace(CStr(Z), "*", vbNullString)) + 1
      End If
      Exit Function
   End If
   
   digits = Len(TestString) - KommaPos - 1
   
   If Left(TestString, 1) = "-" Then
      IsNegativ = True
   End If
   
   If AddToAbsoluteValue And IsNegativ Then
      GetNextDigitNumber = CDbl(TestString) + AdditionalDecDigitKorr / 10 ^ digits - AdditionalDecDigitKorr / 10 ^ (digits - 1)
   Else
      GetNextDigitNumber = CDbl(TestString) + (1 - AdditionalDecDigitKorr) / 10 ^ digits
   End If
   
End Function

Private Property Get DecimalMarker() As String

   Static DecChar As String
   Dim x As String
   
   If Len(DecChar) = 0 Then
      x = Trim(CStr(1.2))
      DecChar = Mid(x, 2, 1)
   End If
   
   DecimalMarker = DecChar

End Property

Private Function CharTrim(ByVal ValueToTrim As String, ByVal TrimChar As String) As String

   Dim TrimString As String
   
   TrimString = " " & TrimChar
   Do While InStr(1, ValueToTrim, TrimString)
      ValueToTrim = Replace(ValueToTrim, TrimString, TrimChar)
   Loop
   
   TrimString = TrimChar & " "
   Do While InStr(1, ValueToTrim, TrimString)
      ValueToTrim = Replace(ValueToTrim, TrimString, TrimChar)
   Loop
   
   CharTrim = ValueToTrim

End Function

Friend Function GetRelationalOperatorString(ByRef RelationalOperator As SqlRelationalOperators) As String

   Dim OperatorString As String
   Dim op As SqlRelationalOperators
   
   If (RelationalOperator And SQL_In) = SQL_In Then
      OperatorString = OperatorString & "In"
      If (RelationalOperator And SQL_Not) = SQL_Not Then
         OperatorString = "Not " & OperatorString
      End If
      GetRelationalOperatorString = OperatorString
      Exit Function
   End If
   
   
   If (RelationalOperator And SQL_Not) = SQL_Not Then
      
      op = RelationalOperator Xor SQL_Not
      
      If op = SqlRelationalOperators.SQL_Equal Then ' => "=" zu "<>" .. null berücksichtigen?
         RelationalOperator = SQL_LessThan + SQL_GreaterThan
      ElseIf op = SQL_GreaterThan + SQL_LessThan Then ' => "<>" zu "=" .. null berücksichtigen?
         RelationalOperator = SQL_Equal
      Else
         RelationalOperator = RelationalOperator Xor SQL_Not
         If (op And SQL_Equal) = SQL_Equal Then
            RelationalOperator = RelationalOperator Xor SQL_Equal
         Else
            RelationalOperator = RelationalOperator Or SQL_Equal
         End If
         If (op And SQL_LessThan) = SQL_LessThan Then
            RelationalOperator = RelationalOperator Xor SQL_LessThan
            RelationalOperator = RelationalOperator Or SQL_GreaterThan
         End If
         If (op And SQL_GreaterThan) = SQL_GreaterThan Then
            RelationalOperator = RelationalOperator Xor SQL_GreaterThan
            RelationalOperator = RelationalOperator Or SQL_LessThan
         End If
      End If
   End If

   If (RelationalOperator And SQL_LessThan) = SQL_LessThan Then
      OperatorString = OperatorString & "<"
   End If
   
   If (RelationalOperator And SQL_GreaterThan) = SQL_GreaterThan Then
      OperatorString = OperatorString & ">"
   End If

   If (RelationalOperator And SQL_Equal) = SQL_Equal Then
      OperatorString = OperatorString & "="
   End If

   GetRelationalOperatorString = OperatorString

End Function

Private Function TryBuildArrayCriteria(ByRef FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
                                       ByRef RelationalOperator As SqlRelationalOperators, _
                                       ByRef FilterValue As Variant, _
                                       ByRef IgnoreValue As Variant, _
                                       ByRef Criteria As String) As Boolean
   Dim itm As Variant
   Dim ItmCriteria As String
   
   Dim arrFilterValue() As Variant
   
   If Not IsArray(FilterValue) Then
      Exit Function
   End If

   'Kriterien über Or verbinden
   For Each itm In FilterValue
      ItmCriteria = BuildCriteria(FieldName, FieldDataType, RelationalOperator, itm, , IgnoreValue)
      If Len(ItmCriteria) > 0 Then
         Criteria = Criteria & SqlOrConcatString & ItmCriteria
      End If
   Next
   If Len(Criteria) > 0 Then
      Criteria = Mid(Criteria, Len(SqlOrConcatString) + 1) ' 1. Or wegschneiden
   End If

   TryBuildArrayCriteria = True

End Function

Private Function TryBuildInCriteria(ByRef FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
                                    ByRef RelationalOperator As SqlRelationalOperators, _
                                    ByRef FilterValue As Variant, _
                                    ByRef IgnoreValue As Variant, _
                                    ByRef Criteria As String) As Boolean

   Dim OperatorString As String
   Dim FilterValueString As String

   If (RelationalOperator And SQL_In) = 0 Then
      Exit Function
   End If

   If IsArray(FilterValue) Then
      FilterValueString = GetValueArrayString(FilterValue, FieldDataType, ",", IgnoreValue)
   ElseIf VarType(FilterValue) = vbString Then ' Value ist bereits die Auflistung als String
      FilterValueString = FilterValue
   Else
      FilterValueString = ConvertToSqlText(FilterValue, FieldDataType)
   End If

   OperatorString = " In "
   If (RelationalOperator And SQL_Not) = SQL_Not Then
      OperatorString = " Not" & OperatorString
   End If

   If Len(FilterValueString) > 0 Then
   
      If RemoveNullFromInValueString(FilterValueString) Then
         Criteria = FieldName & " Is Null"
         If Len(FilterValueString) > 0 Then
            Criteria = Criteria & " Or " & FieldName & OperatorString & "(" & FilterValueString & ")"
         End If
      Else
         Criteria = FieldName & OperatorString & "(" & FilterValueString & ")"
      End If
    
   End If

   TryBuildInCriteria = True

End Function

Private Function TryBuildBetweenCriteria(ByRef FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
                                         ByRef RelationalOperator As SqlRelationalOperators, _
                                         ByRef FilterValue As Variant, _
                                         ByRef FilterValue2 As Variant, _
                                         ByRef IgnoreValue As Variant, _
                                         ByRef Criteria As String) As Boolean
   
   Dim Criteria1 As String
   Dim Criteria2 As String
   
   If (RelationalOperator And SQL_Between) = False Then
      TryBuildBetweenCriteria = False
      Exit Function
   End If
   
   If (RelationalOperator And SQL_Not) = SQL_Not Then 'Bedingung umdrehen
      Criteria1 = BuildCriteria(FieldName, FieldDataType, SQL_LessThan, FilterValue, , IgnoreValue)
      Criteria2 = BuildCriteria(FieldName, FieldDataType, SQL_GreaterThan, FilterValue2, , IgnoreValue)
      Criteria = Criteria1 & SqlAndConcatString & Criteria2
      TryBuildBetweenCriteria = True
      Exit Function
   End If
   
   If FieldDataType = SQL_Numeric Then
      If FilterValue2 Like "<=*" Then 'wegschneiden
         FilterValue2 = Mid(FilterValue2, 3)
      ElseIf FilterValue2 Like "<*" Then
         FilterValue2 = Mid(FilterValue2, 2)
         Criteria1 = BuildCriteria(FieldName, FieldDataType, SQL_GreaterThan + SQL_Equal, FilterValue)
         Criteria2 = BuildCriteria(FieldName, FieldDataType, SQL_LessThan, FilterValue2)
         Criteria = Criteria1 & SqlAndConcatString & Criteria2
         TryBuildBetweenCriteria = True
         Exit Function
      End If
   End If

   If IsNull(FilterValue2) Or IsMissing(FilterValue2) Or ValuesAreEqual(FieldDataType, FilterValue2, IgnoreValue) Then
      RelationalOperator = SQL_GreaterThan + SQL_Equal
   ElseIf IsNull(FilterValue) Or ValuesAreEqual(FieldDataType, FilterValue, IgnoreValue) Then
      RelationalOperator = SQL_LessThan + SQL_Equal
      FilterValue = FilterValue2
      FilterValue2 = GetCheckedIgnoreValue(IgnoreValue)
   ElseIf (FieldDataType And SQL_Date) = SQL_Date And (RelationalOperator And SQL_Add_WildCardSuffix) Then
      Criteria1 = BuildCriteria(FieldName, FieldDataType, SQL_GreaterThan + SQL_Equal, FilterValue)
      Criteria2 = BuildCriteria(FieldName, FieldDataType, SQL_LessThan + SQL_Equal + SQL_Add_WildCardSuffix, FilterValue2)
      Criteria = Criteria1 & SqlAndConcatString & Criteria2
      TryBuildBetweenCriteria = True
      Exit Function
   ElseIf (FieldDataType And SQL_Numeric) = SQL_Numeric And (RelationalOperator And SQL_Add_WildCardSuffix) Then
      Criteria1 = BuildCriteria(FieldName, FieldDataType, SQL_GreaterThan + SQL_Equal, FilterValue)
      Criteria2 = BuildCriteria(FieldName, FieldDataType, SQL_LessThan + SQL_Equal + SQL_Add_WildCardSuffix, FilterValue2)
      Criteria = Criteria1 & SqlAndConcatString & Criteria2
      TryBuildBetweenCriteria = True
      Exit Function
   Else
      Criteria = FieldName & " Between " & ConvertToSqlText(FilterValue, FieldDataType) & SqlAndConcatString & ConvertToSqlText(FilterValue2, FieldDataType)
      TryBuildBetweenCriteria = True
      Exit Function
   End If

End Function

Private Function GetCheckedIgnoreValue(ByVal IgnoreValue As Variant) As Variant
   If IsArray(IgnoreValue) Then
      GetCheckedIgnoreValue = IgnoreValue(LBound(IgnoreValue))
   Else
      GetCheckedIgnoreValue = IgnoreValue
   End If
End Function

Private Function NullFilterOrEmptyFilter(ByVal FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
                                          ByVal RelationalOperator As SqlRelationalOperators, _
                                         ByVal Value As Variant, ByVal IgnoreValue As Variant, _
                                         ByRef NullFilterString As String, _
                                Optional ByVal DisableIgnoreNullValue As Boolean = False) As Boolean
   
   If IsObject(IgnoreValue) Then
   If IgnoreValue Is Nothing Then
      If IsNull(Value) Then
         If (RelationalOperator And SQL_Not) = SQL_Not Then
            NullFilterString = FieldName & " Is Not Null"
         Else
            NullFilterString = FieldName & " Is Null"
         End If
         NullFilterOrEmptyFilter = True
      Else
         NullFilterOrEmptyFilter = False
      End If
      Exit Function
   End If
   End If
   
   If IsNull(Value) Then
      If DisableIgnoreNullValue Then
         NullFilterString = FieldName & " Is Null"
      ElseIf Not ValuesAreEqual(FieldDataType, Value, IgnoreValue) Then
         NullFilterString = FieldName & " Is Null"
      End If
      NullFilterOrEmptyFilter = True
   ElseIf IsArray(Value) Then
      Dim CheckArray() As Variant
On Error Resume Next
      CheckArray = Value
      If Err.Number = 0 Then
      If (0 / 1) + (Not Not CheckArray) = 0 Then
         NullFilterOrEmptyFilter = True
         Exit Function
      End If
      Else
         Err.Clear
         Dim ArraySize As Long
         ArraySize = UBound(Value)
         If Err.Number <> 0 Then
            Err.Clear
            NullFilterOrEmptyFilter = True
            Exit Function
         End If
      End If
   Else
      NullFilterOrEmptyFilter = ValuesAreEqual(FieldDataType, Value, IgnoreValue)
   End If
   
   If (RelationalOperator And SQL_Not) = SQL_Not Then
         NullFilterString = Replace(NullFilterString, "Is Null", "Is Not Null")
   End If

End Function

Private Function ValuesAreEqual(ByVal FieldDataType As SqlFieldDataType, ByVal Value As Variant, ByVal Value2 As Variant) As Boolean
   
   If IsArray(Value2) Then
      ValuesAreEqual = ArrayContains(FieldDataType, Value2, Value)
   ElseIf IsNull(Value) Then
      ValuesAreEqual = IsNull(Value2)
   ElseIf IsNull(Value2) Then
      ValuesAreEqual = False
   Else
      Select Case FieldDataType
         Case SqlFieldDataType.SQL_Text
            ValuesAreEqual = (VBA.StrComp(Value, Value2, vbTextCompare) = 0)
         Case SqlFieldDataType.SQL_Numeric
            ValuesAreEqual = (CDbl(Value) = CDbl(Value2))
         Case SqlFieldDataType.SQL_Date
            ValuesAreEqual = (CDate(Value) = CDate(Value2))
         Case SqlFieldDataType.SQL_Boolean
            ValuesAreEqual = (CBool(Value) = CBool(Value2))
         Case Else
            ValuesAreEqual = (Value = Value2)
      End Select
   End If

End Function

Private Function ArrayContains(ByVal FieldDataType As SqlFieldDataType, ByVal ArrayToCheck As Variant, ByVal SearchValue As Variant) As Boolean
   
   Dim i As Long

   If IsNull(SearchValue) Then
      ArrayContains = ArrayContainsNull(ArrayToCheck)
      Exit Function
   End If

   For i = LBound(ArrayToCheck) To UBound(ArrayToCheck)
      If ValuesAreEqual(FieldDataType, ArrayToCheck(i), SearchValue) Then
         ArrayContains = True
         Exit Function
      End If
   Next

   ArrayContains = False

End Function

Private Function ArrayContainsNull(ByVal ArrayToCheck As Variant) As Boolean
   
   Dim i As Long

   For i = LBound(ArrayToCheck) To UBound(ArrayToCheck)
      If IsNull(ArrayToCheck(i)) Then
         ArrayContainsNull = True
         Exit Function
      End If
   Next

   ArrayContainsNull = False

End Function

Private Function GetValueArrayString(ByVal Value As Variant, ByVal FieldDataType As SqlFieldDataType, _
                                     ByVal Delimiter As String, ByVal IgnoreValue As Variant) As String
   
   Dim i As Long
   Dim s As String

   For i = LBound(Value) To UBound(Value)
      If IsArray(IgnoreValue) Then
         If ArrayContains(FieldDataType, IgnoreValue, Value(i)) Then
         Else
            s = s & Delimiter & ConvertToSqlText(Value(i), FieldDataType)
         End If
      Else
         If Value(i) = IgnoreValue Then
         ElseIf IsNull(Value(i)) And IsNull(IgnoreValue) Then
         Else
            s = s & Delimiter & ConvertToSqlText(Value(i), FieldDataType)
         End If
      End If
   Next
   If Len(s) > 0 And Len(Delimiter) > 0 Then
      s = Mid(s, Len(Delimiter) + 1)
   End If
   GetValueArrayString = s

End Function

Private Function RemoveNullFromInValueString(ByRef ValueString As String) As Boolean

   Const NullCheckString As String = ",Null,"
   Dim TestString As String
   
   TestString = "," & ValueString & ","

   If Not (InStr(1, TestString, NullCheckString) > 0) Then
      RemoveNullFromInValueString = False
      Exit Function
   End If
   
   TestString = Replace(TestString, NullCheckString, ",")
   
   If Len(TestString) > 1 Then
      ValueString = Mid(TestString, 2, Len(TestString) - 2)
   Else
      ValueString = vbNullString
   End If
   
   RemoveNullFromInValueString = True

End Function
