VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaoHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: data.dao.DaoHandler
'---------------------------------------------------------------------------------------
'
' DAO data connection methods
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/dao/DaoHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'  <test>_test/data/dao/DaoHandlerTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

#If VBA7 Then
   Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal Bytes As Long)
#Else
   Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal Bytes As Long)
#End If

'DAO Database
Private m_DaoDb As DAO.Database

'Default values for optional DAO enum parameters, e.g. in OpenRecordset.
' (if these are not set, DAO standard applies)
Private m_RecordsetTypeEnumDefault As DAO.RecordsetTypeEnum
Private m_RecordsetOptionEnumDefault As DAO.RecordsetOptionEnum
Private m_LockTypeEnumDefault As DAO.LockTypeEnum

'#############################################################
' Group: Events

'---------------------------------------------------------------------------------------
' Event: ErrorMissingCurrentDb
'---------------------------------------------------------------------------------------
' Event for missing CurrentDb reference
'
' Parameters:
'     Msg             - Message about the cause of the event triggering. (Not yet used)
'     NewCurrentDbRef - DAO.Database reference to be used for CurrentDb
'
' Remarks:
'     The required DAO.Database reference can be passed on via the ByRef parameter NewCurrentDbRef.
'
'---------------------------------------------------------------------------------------
Public Event ErrorMissingCurrentDb(ByVal Msg As String, ByRef NewCurrentDbRef As DAO.Database)


'#############################################################
' Group: Class control

'
' Init / Terminate
' ----------------
'
'Private Sub Class_Initialize()
''
'End Sub

Private Sub Class_Terminate()
   Dispose
End Sub

'---------------------------------------------------------------------------------------
' Sub: Dispose
'---------------------------------------------------------------------------------------
'
' Remove object references
'
'---------------------------------------------------------------------------------------
Public Sub Dispose()
   Set m_DaoDb = Nothing
   m_RecordsetTypeEnumDefault = 0
   m_RecordsetOptionEnumDefault = 0
   m_LockTypeEnumDefault = 0
End Sub

'---------------------------------------------------------------------------------------
' Sub: InitRecordsetEnumDefaultValues
'---------------------------------------------------------------------------------------
'
' Set default values for recordset enums, which are used when parameters cannot be set.
'
' Parameters:
'     RecordsetType  -  DAO.RecordsetTypeEnum
'     Options        -  DAO.RecordsetOptionEnum
'     LockEdit       -  DAO.LockTypeEnum
'
' Remarks:
'     If 0 is set, standard of DAO is used (0 parameters are passed on to DAO as "IsMissing" variant value).
'
'---------------------------------------------------------------------------------------
Public Sub InitRecordsetEnumDefaultValues( _
                        Optional ByRef RecordsetType As DAO.RecordsetTypeEnum, _
                        Optional ByRef Options As DAO.RecordsetOptionEnum, _
                        Optional ByRef LockEdit As DAO.LockTypeEnum)

   m_RecordsetTypeEnumDefault = RecordsetType
   m_RecordsetOptionEnumDefault = Options
   m_LockTypeEnumDefault = LockEdit

End Sub

'#############################################################
' Group: Data Connection

'---------------------------------------------------------------------------------------
' Property: CurrentDb
'---------------------------------------------------------------------------------------
'
' CurrentDbC variant
'
' Returns:
'     DAO.Database
'
' Remarks:
'     CurrentDbC variant which, unlike the original version by Michael Kaplan,
'     does not automatically set the DAO database reference to Access.Application.CurrentDb,
'     but requests it via an event if no reference exists yet.
'     This allows any database reference to be passed via the main application.
'     (e.g. if a temp database is to serve as CurrentDb).
'
'---------------------------------------------------------------------------------------
Public Property Get CurrentDb() As DAO.Database
    If (m_DaoDb Is Nothing) Then
        'Error-Event auslösen und hoffen, dass neue Referenz geliefert wird
        RaiseEvent ErrorMissingCurrentDb("DaoHandler.CurrentDb: unavailable CurrentDb-Database", m_DaoDb)
        If (m_DaoDb Is Nothing) Then
        'Use CurrentDb of the application if DaoHandler "self-instantiating" (VB_PredeclaredId = True) is used.
        'Attention! ... this check may have the disadvantage that from here on VB_PredeclaredId takes effect,
        '               although it may not be needed.
        If Me Is DaoHandler Then 'applies only if instance is identical to VB_PredeclaredId instance
        '                         Every other instance of DaoHandler has different memory address (ObjPtr(Me) <> ObjPtr(DaoHandler)).
           Set Me.CurrentDb = Application.CurrentDb
        End If
        End If
    End If
    Set CurrentDb = m_DaoDb
End Property

Public Property Set CurrentDb(ByRef NewCurrentDb As DAO.Database)
    Set m_DaoDb = NewCurrentDb
End Property

Friend Property Get CurrentDbReferenceOnly() As DAO.Database
    Set CurrentDbReferenceOnly = m_DaoDb
End Property

'#############################################################
' Group: Data Operations

'---------------------------------------------------------------------------------------
' Function: Execute
'---------------------------------------------------------------------------------------
'
' Execute SQL statement
'
' Parameters:
'     CommandText - SQL statement (String)
'     Options     - DAO.RecordsetOptionEnum
'
' Returns:
'     RecordsAffected (Long)
'
'---------------------------------------------------------------------------------------
Public Function Execute(ByVal Query As String, _
               Optional ByVal Options As DAO.RecordsetOptionEnum) As Long

    If (Options And dbSeeChanges) = 0 Then
        Options = Options + dbSeeChanges
    End If
    Me.CurrentDb.Execute Query, Options
    Execute = Me.CurrentDb.RecordsAffected

End Function

'---------------------------------------------------------------------------------------
' Function: ExecuteParamSQL
'---------------------------------------------------------------------------------------
'
' Execute SQL statement with parameters
'
' Parameters:
'     SqlText     - SQL statement
'     QueryParams - Values in suitable order as ParamArray or as array or 2-dimensional parameter array (Array(n,1) ... x(n,0) = parameter name, x(n,1) = parameter value) ... generateable with GetParamDefArray or GetNamedParamDefArray
'
' Returns:
'     RecordsAffected (Long)
'
'---------------------------------------------------------------------------------------
Public Function ExecuteParamSql(ByVal SqlText As String, _
                           ParamArray QueryParams() As Variant) As Long

   Dim qdf As DAO.QueryDef

On Error GoTo HandleErr

   If IsMissing(QueryParams) Then
      Set qdf = Me.ParamQueryDefSql(SqlText)
   Else
      Set qdf = Me.ParamQueryDefSql(SqlText, QueryParams)
   End If

   ExecuteParamSql = ExecuteQueryDef(qdf)

On Error GoTo 0

   qdf.Close
   Set qdf = Nothing

ExitHere:
   Exit Function

HandleErr:
   If Not (qdf Is Nothing) Then
      qdf.Close
      Set qdf = Nothing
   End If
   Err.Raise Err.Number, "ExecuteParamSQL:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'---------------------------------------------------------------------------------------
' Function: ExecuteQueryDefByName
'---------------------------------------------------------------------------------------
'
' Execute saved query (optionally with parameters)
'
' Parameters:
'     QueryName   - QueryDef name
'     QueryParams - Values in suitable order as ParamArray or as array
'                   or 2-dimensional parameter array (Array(n,1) ... x(n,0) = parameter name, x(n,1) = parameter value) ... generateable with GetParamDefArray or GetNamedParamDefArray
'
' Returns:
'     RecordsAffected (Long)
'
'---------------------------------------------------------------------------------------
Public Function ExecuteQueryDefByName(ByVal QueryName As String, _
                                 ParamArray QueryParams() As Variant) As Long

   Dim qdf As DAO.QueryDef

On Error GoTo HandleErr

   If IsMissing(QueryParams) Then
      Set qdf = Me.ParamQueryDefByName(QueryName)
   Else
      Set qdf = Me.ParamQueryDefByName(QueryName, QueryParams)
   End If

   ExecuteQueryDefByName = ExecuteQueryDef(qdf)

On Error GoTo 0

   qdf.Close
   Set qdf = Nothing

ExitHere:
   Exit Function

HandleErr:
   If Not (qdf Is Nothing) Then
      qdf.Close
      Set qdf = Nothing
   End If
   Err.Raise Err.Number, "ExecuteParamQdfByName:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'---------------------------------------------------------------------------------------
' Function: ExecuteQueryDef
'---------------------------------------------------------------------------------------
'
' Fill QueryDef object with parameter values and then execute it
'
' Parameters:
'     QdfRef      - QueryDef reference
'     QueryParams - Values in suitable order as ParamArray or as array
'                   or 2-dimensional parameter array (Array(n,1) ... x(n,0) = parameter name, x(n,1) = parameter value) ... generateable with GetParamDefArray or GetNamedParamDefArray
'
' Returns:
'     RecordsAffected (Long)
'
'---------------------------------------------------------------------------------------
Public Function ExecuteQueryDef(ByVal QdfRef As DAO.QueryDef, _
                       ParamArray QueryParams() As Variant) As Long

   Dim QdfParamDefs() As Variant

   If Not IsMissing(QueryParams) Then
      QdfParamDefs = GetParamDefArray(QueryParams)
      FillQueryDefParameters QdfRef, QdfParamDefs
   End If

   QdfRef.Execute dbFailOnError

   ExecuteQueryDef = QdfRef.RecordsAffected

End Function

Private Sub FillQueryDefParameters(ByVal QdfRef As DAO.QueryDef, ByVal QdfParamDefs As Variant)

   Dim ParamDefsArraySize As Long
   Dim ParamIndex As Variant
   Dim i As Long

   ParamDefsArraySize = UBound(QdfParamDefs, 1)

   With QdfRef
      For i = 0 To ParamDefsArraySize
         ParamIndex = QdfParamDefs(i, 0)
         If Len(ParamIndex) > 0 Then
            .Parameters(ParamIndex).Value = QdfParamDefs(i, 1)
         End If
      Next i
   End With

End Sub

'---------------------------------------------------------------------------------------
' Function: ParamQueryDefByName
'---------------------------------------------------------------------------------------
'
' Open saved query as QueryDef object incl. passed parameter values
'
' Parameters:
'     QueryName   - name of query with parameters
'     QueryParams - Values in suitable order as ParamArray or as array
'                   or 2-dimensional parameter array (Array(n,1) ... x(n,0) = parameter name, x(n,1) = parameter value) ... generateable with GetParamDefArray or GetNamedParamDefArray
'
' Returns:
'     DAO.QueryDef
'
'---------------------------------------------------------------------------------------
Public Function ParamQueryDefByName(ByVal QueryName As String, ParamArray QueryParams() As Variant) As DAO.QueryDef

   Dim qdf As DAO.QueryDef
   Dim QdfParamDefs() As Variant

   Set qdf = Me.CurrentDb.QueryDefs(QueryName)
   If Not IsMissing(QueryParams) Then
      QdfParamDefs = GetParamDefArray(QueryParams)
      FillQueryDefParameters qdf, QdfParamDefs
   End If

   Set ParamQueryDefByName = qdf

End Function

'---------------------------------------------------------------------------------------
' Function: ParamQueryDefSql
'---------------------------------------------------------------------------------------
'
' Temporäres QueryDef-Objekt ink. übergebenen Parameterwerten öffnen
'
' Parameters:
'     ParamSqlText   - SQL statement for the temporary QueryDef object
'     QueryParams    - Values in suitable order as ParamArray or as array
'                      or 2-dimensional parameter array (Array(n,1) ... x(n,0) = parameter name, x(n,1) = parameter value) ... generateable with GetParamDefArray or GetNamedParamDefArray
'
' Returns:
'     DAO.QueryDef
'
'---------------------------------------------------------------------------------------
Public Function ParamQueryDefSql(ByVal ParamSqlText As String, ParamArray QueryParams() As Variant) As DAO.QueryDef

   Dim qdf As DAO.QueryDef
   Dim QdfParamDefs() As Variant

   Set qdf = Me.CurrentDb.CreateQueryDef("", ParamSqlText)
   If Not IsMissing(QueryParams) Then
      QdfParamDefs = GetParamDefArray(QueryParams)
      FillQueryDefParameters qdf, QdfParamDefs
   End If

   Set ParamQueryDefSql = qdf

End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordset
'---------------------------------------------------------------------------------------
'
' Open DAO.Recordset
'
' Parameters:
'     Source            - SQL statement or table or query name
'     RecordsetType     - DAO.RecordsetTypeEnum (Standard: dbOpenDynaset)
'     RecordsetOptions  - DAO.RecordsetOptionEnum (Standard: dbSeeChanges)
'     LockEdit          - DAO.LockTypeEnum (Standard: dbOptimistic)
'
' Returns:
'     DAO.Recordset
'
'---------------------------------------------------------------------------------------
Public Function OpenRecordset(ByVal Source As String, _
                     Optional ByVal RecordsetType As DAO.RecordsetTypeEnum = dbOpenDynaset, _
                     Optional ByVal RecordsetOptions As DAO.RecordsetOptionEnum = DAO.RecordsetOptionEnum.dbSeeChanges, _
                     Optional ByVal LockEdit As DAO.LockTypeEnum) As DAO.Recordset

   If (RecordsetOptions And dbSeeChanges) = 0 Then
      RecordsetOptions = RecordsetOptions + dbSeeChanges
   End If

   ' Replace0WithIsMissing ... so that IsMissing for Variant data type works on original DAO openrecordset
   Set OpenRecordset = Me.CurrentDb.OpenRecordset(Source, _
                                                  Replace0WithIsMissing(RecordsetType), _
                                                  Replace0WithIsMissing(RecordsetOptions), _
                                                  Replace0WithIsMissing(LockEdit))

End Function

Private Function Replace0WithIsMissing(ByVal CheckValue As Long) As Variant
   If CheckValue = 0 Then
      Replace0WithIsMissing = GetIsMissing()
   Else
      Replace0WithIsMissing = CheckValue
   End If
End Function

Private Function GetIsMissing(Optional ByVal DoNotFillThisParam As Variant) As Variant
   GetIsMissing = DoNotFillThisParam
End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordsetParamSQL
'---------------------------------------------------------------------------------------
'
' Open parameter SQL statement (uses temporary QueryDef)
'
' Parameters:
'     ParamSqlText      - SQL statement with parameters (parameter declaration)
'     RecordsetType     - DAO.RecordsetTypeEnum (Standard: dbOpenDynaset)
'     RecordsetOptions  - DAO.RecordsetOptionEnum (Standard: dbSeeChanges)
'     LockEdit          - DAO.LockTypeEnum (Standard: dbOptimistic)
'     QueryParams       - Values in suitable order as ParamArray or as array
'                         or 2-dimensional parameter array (Array(n,1) ... x(n,0) = parameter name, x(n,1) = parameter value) ... generateable with GetParamDefArray or GetNamedParamDefArray
'
' Returns:
'     DAO.Recordset
'
'---------------------------------------------------------------------------------------
Public Function OpenRecordsetParamSql(ByVal ParamSqlText As String, _
                     ByVal RecordsetType As DAO.RecordsetTypeEnum, _
                     ByVal Options As DAO.RecordsetOptionEnum, _
                     ByVal LockEdit As DAO.LockTypeEnum, _
                     ParamArray QueryParams() As Variant) As DAO.Recordset

   Dim qdf As DAO.QueryDef
   Dim QdfParamDefs As Variant '

   QdfParamDefs = GetParamDefArray(QueryParams)

On Error GoTo HandleErr

   Set qdf = Me.CurrentDb.CreateQueryDef("", ParamSqlText)
   Set OpenRecordsetParamSql = OpenRecordsetQueryDef(qdf, QdfParamDefs, RecordsetType, Options, LockEdit)

On Error GoTo 0

   qdf.Close
   Set qdf = Nothing

ExitHere:
   Exit Function

HandleErr:
   If Not (qdf Is Nothing) Then
      qdf.Close
      Set qdf = Nothing
   End If
   Err.Raise Err.Number, "OpenRecordsetParamSQL:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordsetParamSql2
'---------------------------------------------------------------------------------------
'
' Open parameter SQL statement (like OpenRecordsetParamSQL, but with default values for RecordsetType, RecordsetOptions and LockEdit)
'
' Parameters:
'     ParamSqlText   - SQL statement with parameters (parameter declaration)
'     QueryParams    - Values in suitable order as ParamArray or as array
'                      or 2-dimensional parameter array (Array(n,1) ... x(n,0) = parameter name, x(n,1) = parameter value) ... generateable with GetParamDefArray or GetNamedParamDefArray'
'
' Returns:
'     DAO.Recordset
'
'---------------------------------------------------------------------------------------
Public Function OpenRecordsetParamSql2(ByVal ParamSqlText As String, _
                                       ParamArray QueryParams() As Variant) As DAO.Recordset

   Set OpenRecordsetParamSql2 = OpenRecordsetParamSql(ParamSqlText, _
                                                      m_RecordsetTypeEnumDefault, m_RecordsetOptionEnumDefault, m_LockTypeEnumDefault, _
                                                      QueryParams)

End Function

'---------------------------------------------------------------------------------------
' Function: GetParamDefArray
'---------------------------------------------------------------------------------------
'
' Creates 2-dimensional parameter array (e.g. for OpenRecordsetQueryDef)
'
' Parameters:
'     ParamValues - Parmeter values in matching order (without index identifier!)
'
' Returns:
'     Variant (Array)
'
'---------------------------------------------------------------------------------------
Public Function GetParamDefArray(ParamArray ParamValues() As Variant) As Variant

   Dim i As Long
   Dim QdfParamDefs() As Variant
   Dim ArraySize As Long
   Dim ParamValueArray As Variant  ' !!! nicht als Array deklarieren, da sonst ParamValueArray = ParamValueArray(0) fehltschlägt, wenn ParamValueArray(0) ein Array ist
   Dim ArrayDimCnt As Long, SubArrayDimCnt As Long

   ArraySize = UBound(ParamValues)
   If ArraySize = -1 Then
      GetParamDefArray = QdfParamDefs 'nicht dimensionierten Array zurückgeben
      Exit Function
   End If

   ParamValueArray = ParamValues
   ArrayDimCnt = ArrayDimensionCount(ParamValueArray)
   If ArrayDimCnt = 2 Then
      GetParamDefArray = ParamValues
      Exit Function
   End If

   SubArrayDimCnt = ArrayDimensionCount(ParamValueArray(0))
   If SubArrayDimCnt <> -1 Then

      Do While SubArrayDimCnt <> -1

         ParamValueArray = ParamValueArray(0)

         If SubArrayDimCnt = 2 Then  ' !!! /** \todo CHECK SubArrayDimCnt = 2176 .. ist das immer richtig? **/
         'Bereits aufbereiteter 2-dim. Array
                  GetParamDefArray = ParamValueArray
                  Exit Function
         End If

         SubArrayDimCnt = ArrayDimensionCount(ParamValueArray(0))

      Loop

      ArraySize = UBound(ParamValueArray)

   End If

   ReDim QdfParamDefs(ArraySize, 1)

   For i = 0 To ArraySize
      QdfParamDefs(i, 0) = i
      QdfParamDefs(i, 1) = ParamValueArray(i)
   Next

   GetParamDefArray = QdfParamDefs

End Function

Private Function ArrayDimensionCount(ByRef ArrayToCheck As Variant) As Integer
' siehe http://msdn.microsoft.com/de-de/library/bb978902.aspx .. Funktioniert nicht, wenn Array in Variant eingebettet ist

'   Dim lngCnt As Integer
'   Dim lpSADescriptor As Long, lpSafeArray As Long
'
'On Error Resume Next
'
'   If Not IsArray(ArrayToCheck) Then
'      ArrayDimensionCount = -1
'      Exit Function
'   End If
'
'   ' Adresse des SAFEARRAY-Deskriptors holen
'   CopyMemory lpSADescriptor, ByVal VarPtr(ArrayToCheck) + 8, 4
'   ' Adresse der SAFEARRAY-Struktur holen
'   CopyMemory lpSafeArray, ByVal lpSADescriptor, 4
'
'   If lpSafeArray <> 0 Then
'      ' Die ersten zwei Bytes enthalten die Anzahl der Dimensionen:
'      CopyMemory lngCnt, ByVal lpSafeArray, 2
'   End If
'
'   ArrayDimensionCount = lngCnt

   Dim ArraySize As Long
   Dim DimCnt As Long

   DimCnt = 0

   If Not IsArray(ArrayToCheck) Then
      ArrayDimensionCount = -1
      Exit Function
   End If

On Error Resume Next

   Do
      DimCnt = DimCnt + 1
      ArraySize = UBound(ArrayToCheck, DimCnt)
   Loop While Err.Number = 0
   Err.Clear

On Error GoTo 0

   DimCnt = DimCnt - 1

   ArrayDimensionCount = DimCnt

End Function

'---------------------------------------------------------------------------------------
' Function: GetNamedParamDefArray
'---------------------------------------------------------------------------------------
'
' Creates 2-dimensional parameter array (e.g. for OpenRecordsetQueryDef)
'
' Parameters:
'     ParamValues - Parameter pairs: alternating parameter index and parameter value
'
' Returns:
'     Variant (Array)
'
'---------------------------------------------------------------------------------------
Public Function GetNamedParamDefArray(ParamArray ParamIndexAndValues() As Variant) As Variant

   Dim i As Long
   Dim QdfParamDefs() As Variant
   Dim ArraySize As Long

   ArraySize = UBound(ParamIndexAndValues)
   If ArraySize = -1 Then
      GetNamedParamDefArray = QdfParamDefs 'nicht dimensionierten Array zurückgeben
      Exit Function
   End If

   'Prüfen, ob immer Paare (Index / Wert) möglich sind
   If (((ArraySize + 1) \ 2) * 2) <> (ArraySize + 1) Then
      Err.Raise vbObjectError, "DaoHandler.GetNamedParamDefArray", "ParamArray muss immer Paare für Parameter-Index und Parameter-Wert enthalten."
   End If

   ArraySize = ArraySize \ 2

   ReDim QdfParamDefs(ArraySize, 1)

   For i = 0 To ArraySize
      QdfParamDefs(i, 0) = ParamIndexAndValues(i * 2)
      QdfParamDefs(i, 1) = ParamIndexAndValues(i * 2 + 1)
   Next

   GetNamedParamDefArray = QdfParamDefs

End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordsetQueryDefByName
'---------------------------------------------------------------------------------------
'
' Open recordset from saved query (optionally with parameters)
'
' Parameters:
'     QueryName         - Name of the saved query
'     QueryParams       - Values in suitable order as ParamArray or as array
'                         or 2-dimensional parameter array (Array(n,1) ... x(n,0) = parameter name, x(n,1) = parameter value) ... generateable with GetParamDefArray or GetNamedParamDefArray'
'     RecordsetType     - DAO.RecordsetTypeEnum (Standard: dbOpenDynaset)
'     RecordsetOptions  - DAO.RecordsetOptionEnum (Standard: dbSeeChanges)
'     LockEdit          - DAO.LockTypeEnum (Standard: dbOptimistic)
'
' Returns:
'     DAO.Recordset
'
'---------------------------------------------------------------------------------------
Public Function OpenRecordsetQueryDefByName(ByVal QueryName As String, _
                             Optional ByVal QueryParams As Variant, _
                             Optional ByVal RecordsetType As DAO.RecordsetTypeEnum, _
                             Optional ByVal Options As DAO.RecordsetOptionEnum, _
                             Optional ByVal LockEdit As DAO.LockTypeEnum) As DAO.Recordset

   Dim qdf As DAO.QueryDef

On Error GoTo HandleErr

   Set qdf = Me.CurrentDb.QueryDefs(QueryName)
   Set OpenRecordsetQueryDefByName = OpenRecordsetQueryDef(qdf, QueryParams, RecordsetType, Options, LockEdit)

On Error GoTo 0

   qdf.Close
   Set qdf = Nothing

ExitHere:
   Exit Function

HandleErr:
   If Not (qdf Is Nothing) Then
      qdf.Close
      Set qdf = Nothing
   End If
   Err.Raise Err.Number, "OpenQueryDefRecordsetByName:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordsetQueryDef
'---------------------------------------------------------------------------------------
'
' Recordset von gespeicherter Abfrage (optional mit Parametern) öffnen
'
' Parameters:
'     QdfRef            - QueryDef reference
'     QueryParams       - Values in suitable order as ParamArray or as array
'                         or 2-dimensional parameter array (Array(n,1) ... x(n,0) = parameter name, x(n,1) = parameter value) ... generateable with GetParamDefArray or GetNamedParamDefArray'
'     RecordsetType     - DAO.RecordsetTypeEnum (Standard: dbOpenDynaset)
'     RecordsetOptions  - DAO.RecordsetOptionEnum (Standard: dbSeeChanges)
'     LockEdit          - DAO.LockTypeEnum (Standard: dbOptimistic)
'
' Returns:
'     DAO.Recordset
'
'---------------------------------------------------------------------------------------
Public Function OpenRecordsetQueryDef(ByVal QdfRef As DAO.QueryDef, _
                             Optional ByVal QueryParams As Variant, _
                             Optional ByVal RecordsetType As DAO.RecordsetTypeEnum, _
                             Optional ByVal RecordsetOptions As DAO.RecordsetOptionEnum, _
                             Optional ByVal LockEdit As DAO.LockTypeEnum) As DAO.Recordset

   If Not IsMissing(QueryParams) Then
      FillQueryDefParameters QdfRef, QueryParams
   End If

   Set OpenRecordsetQueryDef = QdfRef.OpenRecordset(Replace0WithIsMissing(RecordsetType), _
                                                 Replace0WithIsMissing(RecordsetOptions), _
                                                 Replace0WithIsMissing(LockEdit))

End Function

'---------------------------------------------------------------------------------------
' Function: LookupSQL
'---------------------------------------------------------------------------------------
'
' Lookup DAO.Recordset replacement function for DLookup (passing a SQL statement)
'
' Parameters:
'     SqlText     - SQL statement
'     Index       - Field selection if the first field of the recordset is not to be returned.
'                   The index is passed in the same way as DAO.Recordset.Fields(index).
'     ValueIfNull - Return value if record is missing or data field value is Null (default: Null).
'
' Returns:
'     Variant  - Null, if SQL statement does not return a record.
'
' Remarks:
'     The SQL statement must be written in the Jet/ACE SQL dialect.
'
'---------------------------------------------------------------------------------------
Public Function LookupSql(ByVal SqlText As String, _
                 Optional ByVal Index As Variant = 0&, _
                 Optional ByVal ValueIfNull As Variant = Null) As Variant

   Dim rst As DAO.Recordset

On Error GoTo HandleErr

   Set rst = Me.OpenRecordset(SqlText, dbOpenForwardOnly, dbSeeChanges, dbReadOnly)
   With rst
      If .EOF Then
         LookupSql = ValueIfNull
      Else
         LookupSql = Nz(.Fields(Index), ValueIfNull)
      End If
      .Close
   End With
   Set rst = Nothing

ExitHere:
   Exit Function

HandleErr:
   If Not (rst Is Nothing) Then
      rst.Close
      Set rst = Nothing
   End If

   Err.Raise Err.Number, "LookupSQL:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'---------------------------------------------------------------------------------------
' Function: Lookup
'---------------------------------------------------------------------------------------
'
' DAO.Recordset replacement function for DLookup
'
' Parameters:
'     Expr        - Field name or calculated expression (select part)
'     Domain      - Table, query or SQL expression for sub-query (FROM part)
'     Criteria    - (Optional) Criteria string (WHERE part)
'     ValueIfNull - Return value if record is missing or data field value is Null (default: Null).
'
' Returns:
'     Variant
'
'---------------------------------------------------------------------------------------
Public Function Lookup(ByVal Expr As String, ByVal Domain As String, _
              Optional ByVal Criteria As Variant, _
              Optional ByVal ValueIfNull As Variant = Null) As Variant

   Dim SelectSql As String

   SelectSql = BuildSelectSql(Expr, Domain, Criteria, False)
   Lookup = LookupSql(SelectSql, , ValueIfNull)

End Function

Private Function BuildSelectSql(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant, _
                       Optional ByVal Distinct As Boolean = False)

   Dim SelectSql As String

   SelectSql = "SELECT "
   If Distinct Then
      SelectSql = SelectSql & "Distinct "
   End If
   SelectSql = SelectSql & Expr & " FROM (" & Domain & ")"
   If Not (VarType(Criteria) = vbError) Then
      If Len(Criteria) > 0 Then
         SelectSql = SelectSql & " WHERE " & Criteria
      End If
   End If

   BuildSelectSql = SelectSql

End Function

'---------------------------------------------------------------------------------------
' Function: Count
'---------------------------------------------------------------------------------------
'
' DAO.Recordset replacement function for DCount
'
' Parameters:
'     Expr        - Field name or calculated expression (select part)
'     Domain      - Table, query or SQL expression for sub-query (FROM part)
'     Criteria    - (optional) Criteria string (WHERE part)
'
' Returns:
'     Long
'
'---------------------------------------------------------------------------------------
Public Function Count(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant, _
             Optional ByVal Distinct As Boolean = False) As Long

   If Distinct Then
      If Expr <> "*" Then
         Domain = "(" & BuildSelectSql(Expr, Domain, Criteria, True) & ")"
         Criteria = vbNullString
      End If
   End If

   Count = Nz(Me.Lookup("Count(" & Expr & ")", Domain, Criteria), 0)

End Function

'---------------------------------------------------------------------------------------
' Function: Max
'---------------------------------------------------------------------------------------
'
' DAO.Recordset replacement function for DMax
'
' Parameters:
'     Expr        - Field name or calculated expression (select part)
'     Domain      - Table, query or SQL expression for sub-query (FROM part)
'     Criteria    - (Optional) Criteria string (WHERE part)
'
' Returns:
'     Variant
'
'---------------------------------------------------------------------------------------
Public Function Max(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
    Max = Me.Lookup("Max(" & Expr & ")", Domain, Criteria)
End Function

'---------------------------------------------------------------------------------------
' Function: Min
'---------------------------------------------------------------------------------------
'
' DAO.Recordset replacement function for DMin
'
' Parameters:
'     Expr        - Field name or calculated expression (select part)
'     Domain      - Table, query or SQL expression for sub-query (FROM part)
'     Criteria    - (Optional) Criteria string (WHERE part)
'
' Returns:
'     Variant
'
'---------------------------------------------------------------------------------------
Public Function Min(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
    Min = Me.Lookup("Min(" & Expr & ")", Domain, Criteria)
End Function

'---------------------------------------------------------------------------------------
' Function: Sum
'---------------------------------------------------------------------------------------
'
' DAO.Recordset replacement function for DSum
'
' Parameters:
'     Expr        - Field name or calculated expression (select part)
'     Domain      - Table, query or SQL expression for sub-query (FROM part)
'     Criteria    - (Optional) Criteria string (WHERE part)
'
' Returns:
'     Variant
'
'---------------------------------------------------------------------------------------
Public Function Sum(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
    Sum = Me.Lookup("Sum(" & Expr & ")", Domain, Criteria)
End Function

'---------------------------------------------------------------------------------------
' Function: Exists
'---------------------------------------------------------------------------------------
'
' Check if record exists
'
' Parameters:
'     Domain   - data source (FROM part)
'     Criteria - (Optional) Criteria string (WHERE part)
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function Exists(ByVal Domain As String, _
              Optional ByVal Criteria As String = vbNullString) As Boolean

   Dim strSQL As String

   strSQL = "SELECT True FROM " & Domain
   If Len(Criteria) > 0 Then
      strSQL = strSQL & " WHERE " & Criteria
   End If
   Exists = (LookupSql(strSQL, , False) = True)

End Function

'---------------------------------------------------------------------------------------
' Function: InsertIdentityReturn
'---------------------------------------------------------------------------------------
'
' Execute insert SQL statement and return last identity value (auto value)
'
' Parameters:
'     InsertSQL   - Insert SQL statement
'
' Returns:
'     Identity value or NULL if no record was appended
'
' Remarks:
'     Works for Jet only from Jet 4.0 (Access 2000), IdentityTable is only applicable for MSSQL, under MySQL the IDENT_CURRENT function does not exist.
'     (<data.adodb.AdodbHandler::InsertIdentityReturn> is more suitable regarding use for active DBMS.)
'---------------------------------------------------------------------------------------
Public Function InsertIdentityReturn(ByVal InsertSQL As String) As Variant

    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim RecordsAffected As Long

On Error GoTo HandleErr

    Set db = Me.CurrentDb
    db.Execute InsertSQL
    RecordsAffected = db.RecordsAffected

    If RecordsAffected = 0 Then
        InsertIdentityReturn = Null 'es wurde kein neuer DS angefügt
        Exit Function
    End If

    Set rst = db.OpenRecordset("SELECT @@Identity") 'zuletzt zugewiesener Autowert ... funktioniert nicht mit MySQL-Tabellen!
    With rst
        If .EOF Then
            InsertIdentityReturn = Null
        Else
            InsertIdentityReturn = .Fields(0)
        End If
        .Close
    End With
    Set rst = Nothing

ExitHere:
   Exit Function

HandleErr:
   If Not (rst Is Nothing) Then
      rst.Close
      Set rst = Nothing
   End If

   Err.Raise Err.Number, "InsertIdentityReturn:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'---------------------------------------------------------------------------------------
' Function: InsertValues
'---------------------------------------------------------------------------------------
'
' Open recordset and append values
'
' Parameters:
'     Source            - Insert data source
'     FieldsAndValues   - ParamArray: "Field1", Value1, "Field2", Value2, ....
'
' Returns:
'     True if fully executed
'
' Remarks:
'     Since a recordset is used for insertion, there is no data type check in the code -> watch out for implicit conversion.
'
'---------------------------------------------------------------------------------------
Public Function InsertValues(ByVal Source As String, ParamArray FieldsAndValues() As Variant) As Boolean
   InsertValues = InsertRecordValuesArrayReturnFieldValue(Source, vbNullString, FieldsAndValues)
End Function

'---------------------------------------------------------------------------------------
' Function: InsertValuesReturnFieldValue
'---------------------------------------------------------------------------------------
'
' Ropen ecordset and append values, returns value of a defined data field
'
' Parameters:
'     Source                  - Insert data source
'     ReturnFieldNameOrIndex  - Datenfeld für Wertrückgabe</param>
'     FieldsAndValues         - ParamArray: "Field1", Value1, "Field2", Value2, ....
'
' Returns:
'     Value from data field
'
' Remarks:
'     Since a recordset is used for insertion, there is no data type check in the code -> watch out for implicit conversion.
'
'---------------------------------------------------------------------------------------
Public Function InsertValuesReturnFieldValue(ByVal Source As String, ByVal ReturnFieldNameOrIndex As Variant, ParamArray FieldsAndValues() As Variant) As Variant
   InsertValuesReturnFieldValue = InsertRecordValuesArrayReturnFieldValue(Source, ReturnFieldNameOrIndex, FieldsAndValues)
End Function

Private Function InsertRecordValuesArrayReturnFieldValue( _
                        ByVal Source As String, ByVal ReturnField As Variant, _
                        ByVal FieldsAndValues As Variant) As Variant

    Dim lngNum As Long, lngAnz As Long, i As Long, anz As Long
    Dim FieldNames() As Variant, FieldValues() As Variant
    Dim rst As DAO.Recordset
    Dim ErrNo As Long, ErrDesc As String, ErrSource As String

On Error GoTo HandleErr

    lngNum = UBound(FieldsAndValues) - LBound(FieldsAndValues) + 1 'hfArraySize(FieldsAndValues)
    ReDim FieldNames(lngNum \ 2)
    ReDim FieldValues(lngNum \ 2)

    anz = 0
    For i = 0 To (lngNum - 2) Step 2
      FieldNames(anz) = FieldsAndValues(i)
      FieldValues(anz) = FieldsAndValues(i + 1)
      anz = anz + 1
    Next i

    Set rst = Me.OpenRecordset(Source, dbOpenDynaset)
    With rst

      .AddNew
      i = 0
      For i = 0 To (anz - 1)
         .Fields(FieldNames(i)) = FieldValues(i)
      Next i
      .Update

      If Len(ReturnField & vbNullString) > 0 Then
          .Bookmark = .LastModified
          InsertRecordValuesArrayReturnFieldValue = .Fields(ReturnField)
      Else
          InsertRecordValuesArrayReturnFieldValue = True
      End If

      .Close

    End With
    Set rst = Nothing

ExitHere:
    Exit Function

ResumeRaiseError:
On Error GoTo 0
   Err.Raise ErrNo, ErrSource, ErrDesc

HandleErr:
    ErrNo = Err.Number
    ErrDesc = Err.Description
    ErrSource = "InsertRecordValuesArrayReturnFieldValue (" & Err.Source & ")"

    If Not (rst Is Nothing) Then
        rst.Close
    End If
    Set rst = Nothing
    Resume ResumeRaiseError

End Function

'---------------------------------------------------------------------------------------
' Function: NewRecord
'---------------------------------------------------------------------------------------
'
' Append record via recordset and set field values.
'
' Parameters:
'     Source                  - Data source (table name or SQL statement)
'     ReturnFieldName         - Field name whose value is to be returned (e.g. autonumber field).
'     FieldsAndValues         - new data ... ParamArray: "Field1", Value1, "Field2", Value2, ....
'
' Returns:
'     Value from ReturnFieldName, or True/False if no ReturnFieldName was specified.
'
'---------------------------------------------------------------------------------------
Public Function NewRecord(ByVal Source As String, ByVal ReturnFieldName As String, _
                 ParamArray FieldsAndValues() As Variant) As Variant

   Dim vA() As Variant
   vA = ConvertParamArrayToVariantArray(FieldsAndValues)
   NewRecord = NewRecord_FieldArray(Source, ReturnFieldName, vA)

End Function

Private Function ConvertParamArrayToVariantArray(ByVal ArrayToConvert As Variant) As Variant()
   ConvertParamArrayToVariantArray = ArrayToConvert
End Function

Private Function NewRecord_FieldArray(ByVal strSource As String, ByVal ReturnFieldName As String, _
                       ByRef DataParams() As Variant) As Variant

    Dim lngNum As Long, i As Long, anz As Long
    Dim DataFields() As Variant, DataValues() As Variant
    Dim rst As DAO.Recordset

    lngNum = UBound(DataParams) + 1
    ReDim DataFields(lngNum \ 2)
    ReDim DataValues(lngNum \ 2)

    anz = 0
    For i = 0 To (lngNum - 2) Step 2
      DataFields(anz) = DataParams(i)
      DataValues(anz) = DataParams(i + 1)
      anz = anz + 1
    Next i

    Set rst = Me.OpenRecordset(strSource, dbOpenDynaset)

On Error GoTo HandleErr
    With rst
      .AddNew
      i = 0
      For i = 0 To (anz - 1)
         .Fields(DataFields(i)) = DataValues(i)
      Next i

      .Update
      If Len(ReturnFieldName) > 0 Then
          .Bookmark = .LastModified
          NewRecord_FieldArray = .Fields(ReturnFieldName).Value
      Else
          NewRecord_FieldArray = True
      End If
      .Close
    End With
    Set rst = Nothing

ExitHere:
    Exit Function

HandleErr:
   Dim FncErrNr As Long
   Dim FncErrDesc As String

   FncErrNr = Err.Number
   FncErrDesc = Err.Description
   Err.Clear
On Error Resume Next
    If Not (rst Is Nothing) Then
        rst.Close
        Set rst = Nothing
    End If
On Error GoTo 0
   Err.Raise FncErrNr, "NewRecord_FieldArray", FncErrDesc

End Function

'---------------------------------------------------------------------------------------
' Function: CopyRecord
'---------------------------------------------------------------------------------------
'
' Copy data set and change individual values
'
' Parameters:
'     Source            - Data source (table name or SQL statement)
'     StartFieldNumber  - Data field number (=field index+1) from which the data is copied
'     LastFieldNumber   - Data field number (=field index+1) up to which the data are copied. A negative value counts from behind.
'     ReturnFieldName   - Field name whose value is to be returned (e.g. auto number field).
'     DataToChange      - data to be changed ... alternating: field name, field value - ParamArray: "Field1", Value1, "Field2", Value2, ....
'
' Returns:
'     Value from ReturnFieldName, or True/False if no ReturnFieldName was specified.
'
'---------------------------------------------------------------------------------------
Public Function CopyRecord(ByVal Source As String, _
                 ByVal StartFieldNumber As Long, ByVal LastFieldNumber As Long, _
                 ByVal ReturnFieldName As String, _
                 ParamArray DataToChange() As Variant) As Variant

    Dim vA() As Variant
    vA = ConvertParamArrayToVariantArray(DataToChange)

    CopyRecord = CopyRecord_FieldArray(Source, StartFieldNumber, LastFieldNumber, ReturnFieldName, vA)

End Function

Private Function CopyRecord_FieldArray(ByVal Source As String, _
                                ByVal StartFieldNumber As Long, ByVal LastFieldNumber As Long, _
                                ByVal ReturnFieldName As String, ByRef DataToChange() As Variant) As Variant

On Error GoTo HandleErr

   Dim lngNum As Long, i As Long, iDS As Long, anz As Long
   Dim strCtl() As Variant, valCtl() As Variant
   Dim rst As DAO.Recordset
   Dim lngLastCopyField As Long, dsAnz As Long, Start As Long
   Dim fCopyVal() As Variant

   lngNum = UBound(DataToChange) + 1

   anz = 0
   ReDim strCtl((lngNum \ 2) - 1)
   ReDim valCtl((lngNum \ 2) - 1)

   For i = 0 To (lngNum - 2) Step 2
     strCtl(anz) = DataToChange(i)
     valCtl(anz) = DataToChange(i + 1)
     anz = anz + 1
   Next i

   'Recordsetobjekt öffnen
   Set rst = OpenRecordset(Source, dbOpenDynaset)

   If rst.EOF And rst.BOF Then
         rst.Close
         Set rst = Nothing
         CopyRecord_FieldArray = 0
         Exit Function
   End If

   With rst
      .MoveLast
      .MoveFirst
      dsAnz = .RecordCount

      If LastFieldNumber = 0 Then
          lngLastCopyField = .Fields.Count
      ElseIf LastFieldNumber < 0 Then
          lngLastCopyField = .Fields.Count + LastFieldNumber
      Else
          lngLastCopyField = LastFieldNumber
      End If

      ReDim fCopyVal(dsAnz, lngLastCopyField)

      If StartFieldNumber > 1 Then Start = StartFieldNumber - 1 Else Start = 0

      'Werte einlesen
      iDS = 0
      While Not .EOF
          For i = Start To lngLastCopyField - 1
            fCopyVal(iDS, i) = .Fields(i).Value
          Next i
          .MoveNext
          iDS = iDS + 1
      Wend
      .MoveLast

      'Werte kopieren
      For iDS = 0 To (dsAnz - 1)
         .AddNew
         For i = Start To (lngLastCopyField - 1)
           .Fields(i) = fCopyVal(iDS, i)
         Next i
         For i = 0 To (anz - 1)
           .Fields(strCtl(i)) = valCtl(i)
         Next i
         .Update
      Next iDS

      If Len(ReturnFieldName) > 0 Then
          .Bookmark = .LastModified
          CopyRecord_FieldArray = .Fields(ReturnFieldName).Value
      Else
          CopyRecord_FieldArray = True
      End If

      .Close

  End With

ExitHere:
    Exit Function

HandleErr:
   Dim FncErrNr As Long
   Dim FncErrDesc As String

   FncErrNr = Err.Number
   FncErrDesc = Err.Description
   Err.Clear
On Error Resume Next
    If Not (rst Is Nothing) Then
        rst.Close
        Set rst = Nothing
    End If
On Error GoTo 0
   Err.Raise FncErrNr, "DaoHandler.CopyRecord_FieldArray", FncErrDesc

End Function

'---------------------------------------------------------------------------------------
' Function: CopyRecords
'---------------------------------------------------------------------------------------
'
' Recordset öffnen und Werte eines Datensatzes anfügen
'
' Parameters:
'     Source            - Data source (table name or SQL statement)
'     StartFieldNumber  - Data field number (=field index+1) from which the data is copied
'     LastFieldNumber   - Data field number (=field index+1) up to which the data are copied. A negative value counts from behind.
'     DataToChange      - data to be changed ... alternating: field name, field value - ParamArray: "Field1", Value1, "Field2", Value2, ....
'
' Returns:
'     Boolean - True if success
'
' Remarks:
'     Since a recordset is used for insertion, there is no data type check in the code -> watch out for implicit conversion.
'
'---------------------------------------------------------------------------------------
Public Function CopyRecords(ByVal Source As String, _
                            ByVal StartFieldNumber As Long, ByVal LastFieldNumber As Long, _
                            ParamArray DataToChange() As Variant) As Boolean

    Dim maxIdx As Long
    Dim vA() As Variant
    Dim i As Long

    maxIdx = UBound(DataToChange)
    ReDim vA(maxIdx)

    For i = 0 To maxIdx
        vA(i) = DataToChange(i)
    Next i

    CopyRecords = CopyRecordsArrayReturnFieldValue(Source, StartFieldNumber, LastFieldNumber, vbNullString, vA)

End Function

'---------------------------------------------------------------------------------------
' Function: CopyRecordsReturnFieldValue
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datensätze kopieren und Werte ändern
'
' Parameters:
'     Source            - Data source (table name or SQL statement)
'     StartFieldNumber  - Data field number (=field index+1) from which the data is copied
'     LastFieldNumber   - Data field number (=field index+1) up to which the data are copied. A negative value counts from behind.
'     ReturnFieldName   - Field name whose value is to be returned (e.g. auto number field).
'     DataToChange      - data to be changed ... alternating: field name, field value - ParamArray: "Field1", Value1, "Field2", Value2, ....
'
' Returns:
'     Value from ReturnFieldName, or True/False if no ReturnFieldName was specified.
'
' Remarks:
'     Since a recordset is used for insertion, there is no data type check in the code -> watch out for implicit conversion.
'
'---------------------------------------------------------------------------------------
Public Function CopyRecordsReturnFieldValue(ByVal Source As String, _
                           ByVal StartFieldNumber As Long, ByVal LastFieldNumber As Long, _
                           ByVal ReturnFieldName As String, _
                           ParamArray DataToChange() As Variant) As Variant

    Dim maxIdx As Long
    Dim vA() As Variant
    Dim i As Long

    maxIdx = UBound(DataToChange)
    ReDim vA(maxIdx)
    For i = 0 To maxIdx
        vA(i) = DataToChange(i)
    Next i
    CopyRecordsReturnFieldValue = CopyRecordsArrayReturnFieldValue(Source, StartFieldNumber, LastFieldNumber, ReturnFieldName, vA)

End Function

Private Function CopyRecordsArrayReturnFieldValue(ByVal SqlSource As String, _
                           ByVal StartFeldID As Long, ByVal LastFeldID As Long, _
                           ByVal ReturnField As String, ByRef UpdateFieldsAndValues() As Variant) As Variant

On Error GoTo HandleErr

   Dim lngNum As Long, lngAnz As Long, i As Long, iDS As Long, anz As Long
   Dim strCtl() As Variant, valCtl() As Variant
   Dim rst As Object
   Dim lngLastCopyField As Long, dsAnz As Long, Start As Long
   Dim CopyFieldsAndValues() As Variant
   Dim ErrNo As Long, ErrDesc As String, ErrSource As String

   lngNum = UBound(UpdateFieldsAndValues) - LBound(UpdateFieldsAndValues) + 1

   anz = 0
   ReDim strCtl((lngNum \ 2) - 1)
   ReDim valCtl((lngNum \ 2) - 1)

   For i = 0 To (lngNum - 2) Step 2
     strCtl(anz) = UpdateFieldsAndValues(i)
     valCtl(anz) = UpdateFieldsAndValues(i + 1)
     anz = anz + 1
   Next i

   Set rst = Me.OpenRecordset(SqlSource, dbOpenDynaset)

   If rst.EOF And rst.BOF Then
         rst.Close
         Set rst = Nothing
         CopyRecordsArrayReturnFieldValue = 0
         Exit Function
   End If

    With rst
      .MoveLast
      .MoveFirst
      dsAnz = .RecordCount

      If LastFeldID = 0 Then
          lngLastCopyField = .Fields.Count
      ElseIf LastFeldID < 0 Then
          lngLastCopyField = .Fields.Count + LastFeldID
      Else
          lngLastCopyField = LastFeldID
      End If

      ReDim CopyFieldsAndValues(dsAnz, lngLastCopyField)

      If StartFeldID > 1 Then Start = StartFeldID - 1 Else Start = 0

      'Read values
      iDS = 0
      While Not .EOF
          For i = Start To lngLastCopyField - 1
              If .Fields(i).Type <> dbBinary Then
                  CopyFieldsAndValues(iDS, i) = .Fields(i).Value
              End If
          Next i
          .MoveNext
          iDS = iDS + 1
      Wend
      .MoveLast

      'Copy values
      For iDS = 0 To (dsAnz - 1)
          .AddNew
          For i = Start To (lngLastCopyField - 1)
              If .Fields(i).Type <> dbBinary Then
                  .Fields(i) = CopyFieldsAndValues(iDS, i)
              End If
          Next i
          For i = 0 To (anz - 1)
            .Fields(strCtl(i)) = valCtl(i)
          Next i
          .Update
      Next iDS

      If Len(ReturnField & vbNullString) > 0 Then
          .Bookmark = .LastModified
          CopyRecordsArrayReturnFieldValue = .Fields(ReturnField)
      Else
          CopyRecordsArrayReturnFieldValue = True
      End If
      .Close
   End With

ExitHere:
    Set rst = Nothing
    Exit Function

ResumeRaiseError:
On Error GoTo 0
   Err.Raise ErrNo, ErrSource, ErrDesc

HandleErr:
    ErrNo = Err.Number
    ErrDesc = Err.Description
    ErrSource = "CopyRecordsArrayReturnFieldValue (" & Err.Source & ")"

    If Not (rst Is Nothing) Then
        rst.Close
    End If
    Set rst = Nothing
    Resume ResumeRaiseError

End Function
