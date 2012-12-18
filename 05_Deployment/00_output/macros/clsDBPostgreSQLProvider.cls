VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBPostgreSQLProvider"
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
Option Explicit
Implements clsIDatabaseProvider

'-----------------------------------------------------------
'-- PostgreSQL
'Syntax:
'  Comment: --
'  Value modifier: '
'  Object modifier: "
' Has concept of inherits
' CREATE LANGUAGE plpgsql;
'
'CREATE INDEX ik_1
'  ON sampletable2
'  USING btree
'  (subitemid);
'ALTER TABLE sampletable2 CLUSTER ON ik_1;
'
'-- Index: ik_2
'
'-- DROP INDEX ik_2;
'
'CREATE UNIQUE INDEX ik_2
'  ON sampletable2
'  USING btree
'  (subitemid);
'-----------------------------------------------------------

Private Const SP_DROP_FK_NAME = "tmp_dbmodelexcel_drop_table_fk"

Public Sub clsIDatabaseProvider_GetSQLCreateTable(objLogicalTableCollection As Collection, _
                            ByVal withColumnDescription As Boolean)
    Dim sSQL                As String
    Dim sqlCreateFK         As String
    Dim sqlAddDescription   As String
    Dim objLogicalTable     As clsLogicalTable
    
    Output_Initialize
    
    '-- Create tables
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                            withColumnDescription, _
                                            sSQL, _
                                            sqlCreateFK, _
                                            sqlAddDescription)
        Output_WriteLine sSQL & ";" & LINE
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & ";" & LINE, 1
        End If
        If withColumnDescription And Len(sqlAddDescription) > 0 Then
            Output_WriteLine sqlAddDescription & ";" & LINE, 2
        End If
    Next
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropTable(objLogicalTableCollection As Collection)
    Dim sSQL            As String
    Dim objLogicalTable As clsLogicalTable
    
    Output_Initialize
    
    '-- Create the store procedue of drop foreign key
    sSQL = Get_SQL_Pre_RemoveFK()
    Output_WriteLine sSQL & LINE
    
    '-- Drop foreign key relation
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_RemoveFK_Single(objLogicalTable)
        Output_WriteLine sSQL & LINE
    Next
    
    '-- Drop the store procedue of drop foreign key
    sSQL = Get_SQL_Post_RemoveFK()
    Output_WriteLine sSQL & LINE
    
    '-- Drop tables
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_DropTable_Single(objLogicalTable)
        Output_WriteLine sSQL & LINE
    Next
    
    '-- Drop tables
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_DropTable_Single(objLogicalTable)
        Output_WriteLine sSQL & ";" & LINE
    Next
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropAndCreateTable(objLogicalTableCollection As Collection, _
                            ByVal withColumnDescription As Boolean)
    Dim sSQL                As String
    Dim sqlCreateFK         As String
    Dim sqlAddDescription   As String
    Dim objLogicalTable     As clsLogicalTable
    
    Output_Initialize
    
    '-- Create the store procedue of drop foreign key
    sSQL = Get_SQL_Pre_RemoveFK()
    Output_WriteLine sSQL & LINE
    
    '-- Drop foreign key relation
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_RemoveFK_Single(objLogicalTable)
        Output_WriteLine sSQL & LINE
    Next
    
    '-- Drop the store procedue of drop foreign key
    sSQL = Get_SQL_Post_RemoveFK()
    Output_WriteLine sSQL & LINE
    
    '-- Drop tables
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_DropTable_Single(objLogicalTable)
        Output_WriteLine sSQL & LINE
    Next
    
    '-- Drop tables
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_DropTable_Single(objLogicalTable)
        Output_WriteLine sSQL & ";" & LINE
    Next
    
    '-- Create tables
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                            withColumnDescription, _
                                            sSQL, _
                                            sqlCreateFK, _
                                            sqlAddDescription)
        Output_WriteLine sSQL & ";" & LINE
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & ";" & LINE, 1
        End If
        If withColumnDescription And Len(sqlAddDescription) > 0 Then
            Output_WriteLine sqlAddDescription & ";" & LINE, 2
        End If
    Next
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLCreateTableIfNotExists(objLogicalTableCollection As Collection)
    Call Err.Raise(1, "PostgreSQL:CreateTableIfNotExists", "PostgreSQL doesn't support CreateTableIfNotExists.")
End Sub

Private Function Get_SQL_RemoveFK_Single(objLogicalTable As clsLogicalTable) As String
    Dim syntaxRemoveFK      As String
    
    syntaxRemoveFK = "-- Drop {0:table name} foreign keys constraints" _
            & LINE & "SELECT {1:sp name for remove FK}('{0:table name}');" _
            
    '-- Return
    Get_SQL_RemoveFK_Single = FormatString(syntaxRemoveFK, objLogicalTable.tableName, SP_DROP_FK_NAME)
End Function

Private Function Get_SQL_Pre_RemoveFK() As String
    Dim sSQL        As String
    
    sSQL = "-- Create temporary store procedue of dropping foreign keys" _
            & LINE & "CREATE or REPLACE FUNCTION tmp_dbmodelexcel_drop_table_fk(tableName character varying) RETURNS integer AS $$" _
            & LINE & "DECLARE" _
            & LINE & "  fkName character varying;" _
            & LINE & "  removeCount integer := 0;" _
            & LINE & "BEGIN" _
            & LINE & "FOR fkName IN SELECT constraint_name" _
            & LINE & "FROM information_schema.table_constraints" _
            & LINE & "WHERE constraint_type = 'FOREIGN KEY'" _
            & LINE & "AND table_name = tableName" _
            & LINE & "ORDER BY constraint_name LOOP" _
            & LINE & "" _
            & LINE & "removeCount := removeCount + 1;" _
            & LINE & "Execute 'ALTER TABLE ' || tableName || ' DROP CONSTRAINT ' || fkName;" _
            & LINE & "END LOOP;" _
            & LINE & "    " _
            & LINE & "RETURN removeCount;" _
            & LINE & "END;" _
            & LINE & "$$ LANGUAGE plpgsql;"
            
    '-- Return
    Get_SQL_Pre_RemoveFK = FormatString(sSQL, SP_DROP_FK_NAME)
End Function

Private Function Get_SQL_Post_RemoveFK() As String
    Dim sSQL        As String
    
    sSQL = "-- Drop the temporary fuction of dropping foreign keys" _
            & LINE & "DROP FUNCTION {0:sp name for dropping FK}(character varying);"
            
    '-- Return
    Get_SQL_Post_RemoveFK = FormatString(sSQL, SP_DROP_FK_NAME)
End Function

Public Sub Get_SQL_CreateTable_Single(objLogicalTable As clsLogicalTable, _
                                ByVal withColumnDescription As Boolean, _
                                ByRef sqlCreateTable As String, _
                                ByRef sqlCreateFK As String, _
                                ByRef sqlAddDescription As String)
    Dim i               As Integer
    
    Dim syntaxTable         As String
    Dim syntaxColumn        As String
    Dim syntaxDefault       As String
    Dim syntaxPrimaryKey    As String
    Dim syntaxUniqueIndex   As String
    Dim syntaxNoUniqueIndex As String
    
    syntaxTable = "CREATE TABLE {0:table name} (" _
                & "{1:columns definition}" _
                & "{2:primary key}" _
                & "{3:indexes}" _
                & LINE & ")"
    syntaxColumn = "  {0:(i = 1 ? space : ,)}{1:column name} {2:data type} {3:nullable} {4:default}"
    syntaxDefault = "DEFAULT {0:default value}"
    
    Dim sqlCoumn            As String
    Dim sqlPrimaryKey       As String
    Dim sqlUniqueIndex      As String
    Dim sqlNoUniqueIndex    As String
    
    '-- Create Columns sql
    sqlCoumn = ""
    For i = 1 To objLogicalTable.Columns.Count
        With objLogicalTable.Columns(i)
            sqlCoumn = sqlCoumn _
                    & LINE & FormatString(syntaxColumn, _
                                IIf(i = 1, " ", ","), _
                                .columnName, _
                                .dataType, _
                                IIf(.Nullable, "NULL", "NOT NULL"), _
                                FormatString(IIf(Len(.Default) = 0, "", syntaxDefault), .Default))
        End With
    Next
    
    syntaxPrimaryKey = "  ,CONSTRAINT {0:primary key name} PRIMARY KEY ({1:columns})"
    syntaxUniqueIndex = "  ,CONSTRAINT {0:index name} UNIQUE {1:columns with bracket}"
    syntaxNoUniqueIndex = "CREATE INDEX {0:index name} ON ""{1:table name}"" {2:columns}"
    '-- Primary key SQL
    sqlPrimaryKey = ""
    With objLogicalTable.PrimaryKey
        If Len(.PKcolumns) > 0 Then
            sqlPrimaryKey = LINE & FormatString(syntaxPrimaryKey, _
                                SQL_Render_PK_Name(objLogicalTable), _
                                 .PKcolumns)
        End If
    End With
    
    '-- Unqiue Indexes SQL
    sqlUniqueIndex = ""
    For i = 1 To objLogicalTable.Indexes.Count
        With objLogicalTable.Indexes(i)
            If .isUnique Then
                sqlUniqueIndex = sqlUniqueIndex _
                        & LINE & FormatString(syntaxUniqueIndex, _
                                    SQL_Render_IK_Name(objLogicalTable, objLogicalTable.Indexes(i), True), _
                                    .IKColumns)
            End If
        End With
    Next

    '-- non-unique Indexes SQL
    sqlNoUniqueIndex = ""
    For i = 1 To objLogicalTable.Indexes.Count
        With objLogicalTable.Indexes(i)
            If Not .isUnique Then
                sqlNoUniqueIndex = sqlNoUniqueIndex _
                        & LINE _
                        & FormatString(syntaxNoUniqueIndex, _
                                SQL_Render_IK_Name(objLogicalTable, objLogicalTable.Indexes(i), False), _
                                objLogicalTable.tableName, _
                                .IKColumns)
            End If
        End With
    Next
    
    '-- Generate table sql
    sqlCreateTable = FormatString(syntaxTable, _
                        objLogicalTable.tableName, _
                        sqlCoumn, _
                        sqlPrimaryKey, _
                        sqlUniqueIndex)
    
    '-- Generate Foreign key SQL
    Dim sqlFK           As String
    sqlFK = "ALTER TABLE {0:Table Name} ADD CONSTRAINT {1:foreign key name} FOREIGN KEY ({2:column[,..n]}) REFERENCES {3:ref_info}"
    
    sqlCreateFK = ""
    If objLogicalTable.ForeignKeys.Count > 0 Then
        For i = 1 To objLogicalTable.ForeignKeys.Count
            With objLogicalTable.ForeignKeys(i)
                sqlCreateFK = sqlCreateFK _
                                & IIf(i = 1, "", ";") _
                                & LINE _
                                & FormatString(sqlFK, _
                                            objLogicalTable.tableName, _
                                            SQL_Render_FK_Name(objLogicalTable, objLogicalTable.ForeignKeys(i)), _
                                            .FKcolumns, _
                                            .RefTableAndColumns & IIf(Len(.fkOption) = 0, "", " " & .fkOption))
            End With
        Next
    End If
    
    '-- Column description
    Dim syntaxColumnDescription   As String
    
    syntaxColumnDescription = "COMMENT ON COLUMN {1:table name}.{2:column name} IS '{0:column note}';"
    sqlAddDescription = ""
    
    If withColumnDescription Then
        For i = 1 To objLogicalTable.Columns.Count
            With objLogicalTable.Columns(i)
                sqlAddDescription = sqlAddDescription _
                        & LINE _
                        & FormatString(syntaxColumnDescription, _
                            .Note, _
                            objLogicalTable.tableName, _
                            .columnName)
            End With
        Next
    End If
    
End Sub

Private Function Get_SQL_DropTable_Single(objLogicalTable As clsLogicalTable) As String
    Dim sSQL            As String
    
    sSQL = "DROP TABLE IF EXISTS {0:table name};"
   
    '-- Return
    Get_SQL_DropTable_Single = FormatString(sSQL, _
                                objLogicalTable.tableName, _
                                objLogicalTable.Description)
End Function

Private Function SQL_Render_TableName(ByVal tableName As String) As String
    SQL_Render_TableName = Replace(Replace(tableName, " ", ""), "_", "")
End Function

Private Function SQL_Render_PK_Name(objLogicalTable As clsLogicalTable) As String
    SQL_Render_PK_Name = "PK_" & SQL_Render_TableName(objLogicalTable.tableName)
End Function

Private Function SQL_Render_FK_Name(objLogicalTable As clsLogicalTable, _
                                objLogicalForeignKey As clsLogicalForeignKey) As String
    SQL_Render_FK_Name = "FK_" & SQL_Render_TableName(objLogicalTable.tableName) _
                            & "_" & Replace(Replace(Replace(objLogicalForeignKey.FKcolumns, " ", ""), "_", ""), ",", "_")
End Function

Private Function SQL_Render_IK_Name(objLogicalTable As clsLogicalTable, _
                                objLogicalIndex As clsLogicalIndex, _
                                isUnique As Boolean) As String
    SQL_Render_IK_Name = IIf(isUnique, "UI_", "I_") _
                            & SQL_Render_TableName(objLogicalTable.tableName) _
                            & "_" & Replace(Replace(Replace(Replace(Replace(objLogicalIndex.IKColumns, _
                                                                    " ", ""), _
                                                            "(", ""), _
                                                    ")", ""), _
                                            "_", ""), _
                                    ",", "_")
End Function

Private Function SQL_Render_DF_Name(objLogicalTable As clsLogicalTable, _
                                objLogicalColumn As clsLogicalColumn) As String
    SQL_Render_DF_Name = "DF_" & SQL_Render_TableName(objLogicalTable.tableName) & "_" & objLogicalColumn.columnName
End Function




