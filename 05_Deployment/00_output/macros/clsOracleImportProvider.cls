VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOracleImportProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
'-------------------------------------------
'-- Oracle Import Provider
'-------------------------------------------
Option Explicit
Implements IImportProvider

Public Property Get IImportProvider_Providers() As String()
    Dim candidateProviders(1 To 2) As String
    candidateProviders(1) = "MSDAORA.1"
    candidateProviders(2) = "OraOLEDB.Oracle.1"
    
    IImportProvider_Providers = candidateProviders
End Property

Public Property Get IImportProvider_SupportSelectDatabase() As Boolean
    IImportProvider_SupportSelectDatabase = False
End Property

Public Function IImportProvider_GetDatabasesSql() As String
    IImportProvider_GetDatabasesSql = ""
End Function

Public Function IImportProvider_GetTablesSql() As String
    Dim syntax As String
    
    syntax = "  SELECT table_name AS name" _
    & LINE & "    FROM User_Tables T" _
    & LINE & "ORDER BY T.TABLE_NAME"
    
    IImportProvider_GetTablesSql = syntax
    IImportProvider_GetTablesSql = "SELECT TABLE_NAME FROM User_Tables T" _
        & LINE & "ORDER BY T.TABLE_NAME"
End Function

Public Function IImportProvider_GetOptions() As clsImportOptions
    If basImport.mImportOptions Is Nothing Then
        Set basImport.mImportOptions = New clsImportOptions
        basImport.mImportOptions.DataSource = "localhost"
        basImport.mImportOptions.DataSourceTip = "e.g. localhost:1521/SID"
    End If
    Set IImportProvider_GetOptions = basImport.mImportOptions
End Function

Public Function IImportProvider_CreateConnectionString(ByVal Provider As String, _
                    ByVal server As String, _
                    ByVal user As String, _
                    ByVal Password As String, _
                    ByVal Database As String) As String
                    
    IImportProvider_CreateConnectionString = "Provider=" & Trim(Provider) _
            & ";Data Source=" & Trim(server) _
            & ";User ID=" & Trim(user) _
            & ";Password=" & Password _
            & ";Persist Security Info=True"
            
End Function

Public Function IImportProvider_GetLogicalTable(conn As ADODB.Connection, tableName As String) As clsLogicalTable
    Dim objTable As clsLogicalTable
    Set objTable = New clsLogicalTable
    
    objTable.tableName = tableName
    Set objTable.PrimaryKey = New clsLogicalPrimaryKey
    Set objTable.Indexes = New Collection
    Set objTable.ForeignKeys = New Collection
    Set objTable.Columns = New Collection
    
    RenderPKAndIndex conn, objTable
    RenderForeignKey conn, objTable
    RenderColumn conn, objTable
    
    '-- Return
    Set IImportProvider_GetLogicalTable = objTable
End Function

Public Sub RenderPKAndIndex(conn As ADODB.Connection, objTable As clsLogicalTable)
    Dim syntax As String
    
    syntax = "SELECT TI.table_name" _
    & LINE & "     , TI.index_name" _
    & LINE & "     , TI.uniqueness" _
    & LINE & "     , TIC.COLUMN_NAME" _
    & LINE & "     , TIC.COLUMN_POSITION" _
    & LINE & "     , TC.CONSTRAINT_TYPE" _
    & LINE & "  FROM User_Indexes TI" _
    & LINE & "  JOIN user_ind_columns TIC" _
    & LINE & "    ON  TI.table_name = TIC.Table_Name" _
    & LINE & "   AND TI.index_name = TIC.Index_Name" _
    & LINE & "  JOIN user_constraints TC" _
    & LINE & "    ON TI.table_name = TC.Table_Name" _
    & LINE & "   AND TI.index_name = TC.Constraint_Name" _
    & LINE & " WHERE lower(TI.TABLE_NAME) = lower({0:table name})" _
    & LINE & " ORDER BY TI.TABLE_NAME" _
    & LINE & "     , TI.INDEX_NAME" _
    & LINE & "     , TIC.Column_Position"

    Dim sSQL                    As String
    sSQL = FormatString(syntax, SQL_ToSQL(objTable.tableName))
    
    Dim oRs                     As ADODB.Recordset
    Dim curIndexName            As String
    Dim objIndex                As clsLogicalIndex

    On Error GoTo Flag_Err

    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    curIndexName = ""

    Do While Not oRs.EOF
        If oRs("CONSTRAINT_TYPE") = "P" Then
            '-- Primary Key
            If Len(objTable.PrimaryKey.PKcolumns) = 0 Then
                objTable.PrimaryKey.PKcolumns = oRs("COLUMN_NAME") & ""
            Else
                objTable.PrimaryKey.PKcolumns = objTable.PrimaryKey.PKcolumns & ", " & oRs("COLUMN_NAME")
            End If
            objTable.PrimaryKey.IsClustered = True
        Else
            '-- Index
            If curIndexName <> (oRs("INDEX_NAME") & "") Then
                Set objIndex = New clsLogicalIndex
                objTable.Indexes.Add objIndex
                
                objIndex.IsClustered = False
                objIndex.isUnique = (oRs("uniqueness") = "UNIQUE")

                curIndexName = oRs("INDEX_NAME") & ""
            End If

            If Len(objIndex.IKColumns) = 0 Then
                objIndex.IKColumns = oRs("COLUMN_NAME") & ""
            Else
                objIndex.IKColumns = objIndex.IKColumns & ", " & oRs("COLUMN_NAME")
            End If
        End If

        '-- Move next record
        oRs.MoveNext
    Loop

    '-- Close record set
    oRs.Close
    Set oRs = Nothing
    Exit Sub
Flag_Err:
    If Not oRs Is Nothing Then oRs.Close
    Set oRs = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub RenderForeignKey(conn As ADODB.Connection, objTable As clsLogicalTable)
    Dim syntax As String
    
    syntax = "SELECT UC.TABLE_NAME" _
    & LINE & "     , UC.CONSTRAINT_NAME" _
    & LINE & "     , UC.delete_rule" _
    & LINE & "     , UCC.column_name" _
    & LINE & "     , UCC.POSITION" _
    & LINE & "     , UCCR.table_name AS REFERENCED_TABLE_NAME" _
    & LINE & "     , UCCR.column_name AS REFERENCED_COLUMN_NAME" _
    & LINE & "  FROM user_constraints UC" _
    & LINE & "  JOIN user_cons_columns UCC" _
    & LINE & "    ON UC.TABLE_NAME        = UCC.TABLE_NAME" _
    & LINE & "   AND UC.CONSTRAINT_NAME   = UCC.CONSTRAINT_NAME" _
    & LINE & "  JOIN user_cons_columns UCCR" _
    & LINE & "    ON UC.r_constraint_name   = UCCR.constraint_name" _
    & LINE & "   AND UCC.position = UCCR.position" _
    & LINE & " WHERE UC.TABLE_NAME = {0:table name}" _
    & LINE & "   AND UC.constraint_type = 'R'" _
    & LINE & " ORDER BY UC.TABLE_NAME" _
    & LINE & "     , UC.CONSTRAINT_NAME" _
    & LINE & "     , UCC.POSITION"

    Dim sSQL                    As String
    sSQL = FormatString(syntax, SQL_ToSQL(objTable.tableName))
    
    Dim oRs             As ADODB.Recordset
    Dim curFKName       As String
    Dim objForeignKey   As clsLogicalForeignKey
    
    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    curFKName = ""

    Do While Not oRs.EOF
        '-- For Foreign Key
        If curFKName <> (oRs("CONSTRAINT_NAME") & "") Then
            Set objForeignKey = New clsLogicalForeignKey
            objTable.ForeignKeys.Add objForeignKey

            objForeignKey.refTableName = oRs("REFERENCED_TABLE_NAME") & ""
            If oRs("DELETE_RULE") <> "NO ACTION" Then
                objForeignKey.OnDelete = "ON DELETE " & oRs("DELETE_RULE")
            Else
                objForeignKey.OnDelete = ""
            End If
          
            objForeignKey.OnUpdate = ""
            
            curFKName = oRs("CONSTRAINT_NAME") & ""
        End If

        If Len(objForeignKey.FKcolumns) > 0 Then
            objForeignKey.FKcolumns = objForeignKey.FKcolumns & ", "
        End If
        objForeignKey.FKcolumns = objForeignKey.FKcolumns & oRs("COLUMN_NAME")
        
        If Len(objForeignKey.RefTableColumns) > 0 Then
            objForeignKey.RefTableColumns = objForeignKey.RefTableColumns & ", "
        End If
        objForeignKey.RefTableColumns = objForeignKey.RefTableColumns & oRs("REFERENCED_COLUMN_NAME")

        '-- Move next record
        oRs.MoveNext
    Loop

    '-- Close record set
    oRs.Close
    Set oRs = Nothing

    Exit Sub
Flag_Err:
    If Not oRs Is Nothing Then oRs.Close
    Set oRs = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub RenderColumn(conn As ADODB.Connection, objTable As clsLogicalTable)
    Dim syntax As String
    
    syntax = "  SELECT C.TABLE_NAME" _
    & LINE & "       , C.COLUMN_NAME" _
    & LINE & "       , C.COLUMN_ID" _
    & LINE & "       , C.DATA_TYPE" _
    & LINE & "       , C.NULLABLE" _
    & LINE & "       , C.CHAR_LENGTH" _
    & LINE & "       , C.DATA_PRECISION" _
    & LINE & "       , C.DATA_SCALE" _
    & LINE & "       , C.DATA_DEFAULT" _
    & LINE & "    FROM User_Tab_Columns C" _
    & LINE & "   WHERE C.TABLE_NAME = {0:table name}" _
    & LINE & "ORDER BY C.TABLE_NAME" _
    & LINE & "       , C.COLUMN_ID"

    Dim sSQL                    As String
    sSQL = FormatString(syntax, SQL_ToSQL(objTable.tableName))
    
    Dim oRs             As ADODB.Recordset
    Dim objColumn       As clsLogicalColumn
    
    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    Do While Not oRs.EOF
        '-- set Column
        Set objColumn = New clsLogicalColumn
        objTable.Columns.Add objColumn
        
        objColumn.columnName = oRs("COLUMN_NAME") & ""
        objColumn.dataType = GetColumnDataType( _
                                            oRs("DATA_TYPE"), _
                                            oRs("CHAR_LENGTH") & "", _
                                            oRs("DATA_PRECISION") & "", _
                                            oRs("DATA_SCALE") & "")
        objColumn.Nullable = (oRs("NULLABLE") = "Y")
        If Not IsNull(oRs("DATA_DEFAULT")) Then
            objColumn.Default = oRs("DATA_DEFAULT")
        Else
            objColumn.Default = ""
        End If

        '-- Move next record
        oRs.MoveNext
    Loop

    '-- Close record set
    oRs.Close
    Set oRs = Nothing

    Exit Sub
Flag_Err:
    If Not oRs Is Nothing Then oRs.Close
    Set oRs = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Function GetColumnDataType(typeName As String, _
                        maxLength As String, _
                        precision As String, _
                        type_scale As String) As String
    Dim dataType As String

    dataType = LCase(typeName)
    Select Case dataType
        Case "varchar2", "nvarchar2"
            dataType = dataType & "(" & maxLength & ")"
        Case "numeric"
            dataType = dataType & "(" & precision & "," & type_scale & ")"
    End Select

    GetColumnDataType = dataType
End Function



