VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBSQLServerProvider"
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
'-- SQL Server
'-----------------------------------------------------------

Public Sub clsIDatabaseProvider_GetSQLCreateTable(objLogicalTableCollection As Collection, ByVal withColumnDescription As Boolean)

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
        Output_WriteLine sSQL & LINE & "GO" & LINE
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & LINE & "GO", 1
        End If
        If withColumnDescription Then
            Output_WriteLine sqlAddDescription & LINE & "GO", 2
        End If
    Next
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropTable(objLogicalTableCollection As Collection)
    Dim sSQL            As String
    Dim objLogicalTable As clsLogicalTable
    
    Output_Initialize
    
    '-- Drop foreign key relation
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_RemoveFK_Single(objLogicalTable)
        Output_WriteLine sSQL & LINE & "GO" & LINE
    Next
    
    '-- Drop tables
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_DropTable_Single(objLogicalTable)
        Output_WriteLine sSQL & LINE & "GO" & LINE
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
    
    '-- Drop foreign key relation
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_RemoveFK_Single(objLogicalTable)
        Output_WriteLine sSQL & LINE & "GO" & LINE
    Next
    
    '-- Drop tables
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_DropTable_Single(objLogicalTable)
        Output_WriteLine sSQL & LINE & "GO" & LINE
    Next
    
    '-- Create tables
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                            withColumnDescription, _
                                            sSQL, _
                                            sqlCreateFK, _
                                            sqlAddDescription)
        Output_WriteLine sSQL & LINE & "GO" & LINE
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & LINE & "GO", 1
        End If
        If withColumnDescription Then
            Output_WriteLine sqlAddDescription & LINE & "GO", 2
        End If
    Next
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLCreateTableIfNotExists(objLogicalTableCollection As Collection)
    Dim sSQL                As String
    Dim sqlCreateFK         As String
    Dim objLogicalTable     As clsLogicalTable
    
    Output_Initialize
    
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTableIfNotExists_Single(objLogicalTable, _
                                            sSQL, _
                                            sqlCreateFK)
        Output_WriteLine sSQL & LINE & "GO" & LINE
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & LINE & "GO" & LINE, 1
        End If
    Next
    
    Output_Copy
End Sub

Private Sub Get_SQL_CreateTableIfNotExists_Single(objLogicalTable As clsLogicalTable, _
                                ByRef sqlCreateTable As String, _
                                ByRef sqlCreateFK As String)
    Dim sSQL            As String
    
    sSQL = "IF NOT EXISTS (" _
        & LINE & "  SELECT * FROM INFORMATION_SCHEMA.TABLES" _
        & LINE & "  WHERE TABLE_TYPE = 'BASE TABLE'" _
        & LINE & "  AND TABLE_SCHEMA = 'dbo'" _
        & LINE & "  AND TABLE_NAME = '{0:table name}'" _
        & LINE & "  )" _
        & LINE & "BEGIN" _
        & LINE & "{2:create table sql}" _
        & LINE & "END"
    
    Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                False, _
                                sqlCreateTable, _
                                sqlCreateFK, _
                                "")
                                
    sqlCreateTable = FormatString(sSQL, _
                    objLogicalTable.tableName, _
                    objLogicalTable.Description, _
                    sqlCreateTable)
    
End Sub

Private Function Get_SQL_RemoveFK_Single(objLogicalTable As clsLogicalTable) As String
    Dim syntaxRemoveFK      As String
    
    syntaxRemoveFK = "-- Remove {0:table name} foreign key constraint" _
            & LINE & "DECLARE {1:@FkName}  SYSNAME" _
            & LINE & "DECLARE fk_cursor CURSOR FOR " _
            & LINE & "SELECT     CONSTRAINT_NAME" _
            & LINE & "FROM       INFORMATION_SCHEMA.TABLE_CONSTRAINTS" _
            & LINE & "WHERE      TABLE_NAME = '{0:table name}'" _
            & LINE & "AND        CONSTRAINT_TYPE = 'FOREIGN KEY'" _
            & LINE & "ORDER BY   CONSTRAINT_NAME" _
            & LINE & "" _
            & LINE & "OPEN fk_cursor" _
            & LINE & "" _
            & LINE & "FETCH NEXT FROM fk_cursor " _
            & LINE & "INTO {1:@FkName}" _
            & LINE & "" _
            & LINE & "WHILE @@FETCH_STATUS = 0" _
            & LINE & "BEGIN" _
            & LINE & "  EXEC('ALTER TABLE [{0:table name}] DROP CONSTRAINT ' + {1:@FkName})" _
            & LINE & "" _
            & LINE & "  FETCH NEXT FROM fk_cursor " _
            & LINE & "  INTO {1:@FkName}" _
            & LINE & "END" _
            & LINE & "" _
            & LINE & "CLOSE fk_cursor" _
            & LINE & "DEALLOCATE fk_cursor"

    '-- Return
    Get_SQL_RemoveFK_Single = FormatString(syntaxRemoveFK, objLogicalTable.tableName, "@FkName")
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
    
    syntaxTable = "CREATE TABLE [{0:table name}] (" _
                & "{1:columns definition}" _
                & "{2:primary key}" _
                & "{3:indexes}" _
                & LINE & ")" _
                & "{4:non unique indexes}"
    syntaxColumn = "  {0:(i = 1 ? space : ,)}[{1:column name}] {2:data type} {3:nullable} {4:default}"
    syntaxDefault = "  CONSTRAINT {0:default name} DEFAULT ({1:default value})"
    
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
                                FormatString(IIf(Len(.Default) = 0, "", syntaxDefault), _
                                    SQL_Render_DF_Name(objLogicalTable, objLogicalTable.Columns(i)), _
                                    .Default))
        End With
    Next
    
    syntaxPrimaryKey = "  ,CONSTRAINT {0:primary key name} PRIMARY KEY {1:clustered tag} ({2:columns})"
    syntaxUniqueIndex = "  ,CONSTRAINT {0:index name} UNIQUE {1:clustered tag} {2:columns with bracket}"
    syntaxNoUniqueIndex = "CREATE {1:clustered tag} INDEX {0:index name} ON [{2:table name}] {3:columns}"
    '-- Primary key SQL
    sqlPrimaryKey = ""
    With objLogicalTable.PrimaryKey
        If Len(.PKcolumns) > 0 Then
            sqlPrimaryKey = LINE & FormatString(syntaxPrimaryKey, _
                                SQL_Render_PK_Name(objLogicalTable), _
                                IIf(.IsClustered, "CLUSTERED", "NONCLUSTERED"), _
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
                                    SQL_Render_IK_Name(objLogicalTable, objLogicalTable.Indexes(i)), _
                                    IIf(.IsClustered, "CLUSTERED", ""), _
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
                        & LINE & FormatString(syntaxNoUniqueIndex, _
                                    SQL_Render_IK_Name(objLogicalTable, objLogicalTable.Indexes(i)), _
                                    IIf(.IsClustered, "CLUSTERED", ""), _
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
                        sqlUniqueIndex, _
                        sqlNoUniqueIndex)
    
    '-- Generate Foreign key SQL
    Dim sqlFKHead       As String
    Dim sqlFK           As String
    sqlFKHead = "ALTER TABLE [{0:Table Name}] ADD"
    sqlFK = "{0:(i=1? :,)} CONSTRAINT {1:foreign key name} FOREIGN KEY ({2:column[,..n]}) REFERENCES {3:ref_info}"
    
    sqlCreateFK = ""
    If objLogicalTable.ForeignKeys.Count > 0 Then
        sqlCreateFK = FormatString(sqlFKHead, objLogicalTable.tableName)
        For i = 1 To objLogicalTable.ForeignKeys.Count
            With objLogicalTable.ForeignKeys(i)
                sqlCreateFK = sqlCreateFK _
                                & LINE & FormatString(sqlFK, _
                                            IIf(i = 1, " ", ","), _
                                            SQL_Render_FK_Name(objLogicalTable, objLogicalTable.ForeignKeys(i)), _
                                            .FKcolumns, _
                                            .RefTableAndColumns & IIf(Len(.fkOption) = 0, "", " " & .fkOption))
            End With
        Next
    End If
    
    '-- Column description
    Dim syntaxColumnDescription   As String
    syntaxColumnDescription = "EXECUTE sp_addextendedproperty N'MS_Description', N'{0:column note}', N'user', N'dbo', N'table', N'{1:table name}', N'column', N'{2:column name}'"
    sqlAddDescription = ""
    
    If withColumnDescription Then
        For i = 1 To objLogicalTable.Columns.Count
            With objLogicalTable.Columns(i)
                sqlAddDescription = sqlAddDescription _
                        & LINE & FormatString(syntaxColumnDescription, _
                            .Note, _
                            objLogicalTable.tableName, _
                            .columnName)
            End With
        Next
    End If
    
End Sub

Private Function Get_SQL_DropTable_Single(objLogicalTable As clsLogicalTable) As String
    Dim sSQL            As String
    
    sSQL = "IF EXISTS (" _
        & LINE & "  SELECT * FROM INFORMATION_SCHEMA.TABLES" _
        & LINE & "  WHERE TABLE_TYPE = 'BASE TABLE'" _
        & LINE & "  AND TABLE_SCHEMA = 'dbo'" _
        & LINE & "  AND TABLE_NAME = '{0:table name}'" _
        & LINE & "  )" _
        & LINE & "BEGIN" _
        & LINE & "  DROP TABLE [dbo].[{0:table name}]" _
        & LINE & "END"
   
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
                                objLogicalIndex As clsLogicalIndex) As String
    SQL_Render_IK_Name = "IK_" & SQL_Render_TableName(objLogicalTable.tableName) _
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
