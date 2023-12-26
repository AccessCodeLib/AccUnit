Attribute VB_Name = "DaoTools"
Attribute VB_Description = "Hilfsfunktionen für den Umgang mit DAO"
'---------------------------------------------------------------------------------------
' Package: data.dao.DaoTools
'---------------------------------------------------------------------------------------
'
' Auxiliary functions for the handling of DAO
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/dao/DaoTools.bas</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'  <test>_test/data/dao/DaoToolsTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Function: TableDefExists
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prüft ob eine Tabelle (TableDef) vorhanden ist
' </summary>
' <param name="TableDefName">Name der Tabelle</param>
' <param name="dbs">DAO.Database-Referenz (falls keine angegeben wurde, wird CodeDb verwendet)</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function TableDefExists(ByVal TableDefName As String, _
                      Optional ByVal DbRef As DAO.Database = Nothing) As Boolean
'Man könnte auch die TableDef-Liste durchlaufen.
'Eine weitere Alternative wäre das Auswerten über cnn.OpenSchema(adSchemaTables, ...)

   TableDefExists = CheckDatabaseObjectExists(acTable, TableDefName, DbRef)

End Function

'---------------------------------------------------------------------------------------
' Function: QueryDefExists
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prüft ob eine Abfrage (QueryDef) vorhanden ist
' </summary>
' <param name="QueryDefName">Name der Abfrage</param>
' <param name="dbs">DAO.Database-Referenz (falls keine angegeben wurde, wird CodeDb verwendet)</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function QueryDefExists(ByVal QueryDefName As String, _
                      Optional ByVal DbRef As DAO.Database = Nothing) As Boolean

   QueryDefExists = CheckDatabaseObjectExists(acQuery, QueryDefName, DbRef)

End Function

Private Function CheckDatabaseObjectExists(ByVal ObjType As AcObjectType, ByVal ObjName As String, _
                      Optional ByVal DbRef As DAO.Database = Nothing) As Boolean

   Dim rst As DAO.Recordset
   Dim FilterString As String
   Dim ObjectTypeFilterString As String

   If DbRef Is Nothing Then
      Set DbRef = CodeDb
   End If

   FilterString = "where Name = '" & Replace(ObjName, "'", "''") & "'"

   Select Case ObjType
      Case AcObjectType.acTable
         ObjectTypeFilterString = "Type IN (1, 4, 6)"
      Case AcObjectType.acQuery
         ObjectTypeFilterString = "Type =5"
   End Select

   If Len(ObjectTypeFilterString) > 0 Then
      FilterString = FilterString & " AND " & ObjectTypeFilterString
   End If

   Set rst = DbRef.OpenRecordset("select Name from MSysObjects " & FilterString, dbOpenForwardOnly, dbReadOnly)
   CheckDatabaseObjectExists = Not rst.EOF
   rst.Close

End Function
