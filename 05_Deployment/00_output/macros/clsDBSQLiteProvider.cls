VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBSQLiteProvider"
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

Const SQL_Statement_End = ";"
Const ObjectNameStart = """"
Const ObjectNameEnd = """"
Const OutputID_FirstSection = 1
Const OutputID_SecondSection = 2


'-----------------------------------------------------------
'-- SQLite
'-----------------------------------------------------------

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
                                            sqlAddDescription, _
                                            False)
        Output_WriteLine sSQL & LINE
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & LINE & SQL_Statement_End, OutputID_FirstSection
        End If
        If withColumnDescription Then
            Output_WriteLine sqlAddDescription & LINE & SQL_Statement_End, OutputID_SecondSection
        End If
    Next
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropTable(objLogicalTableCollection As Collection)
    Dim sSQL            As String
    Dim objLogicalTable As clsLogicalTable
    
    Output_Initialize
    
    '-- Drop tables
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_DropTable_Single(objLogicalTable)
        Output_WriteLine sSQL & LINE
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
    
    '-- Drop tables
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_DropTable_Single(objLogicalTable)
        Output_WriteLine sSQL & LINE
    Next
    
    '-- Create tables
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                            withColumnDescription, _
                                            sSQL, _
                                            sqlCreateFK, _
                                            sqlAddDescription, _
                                            False)
        Output_WriteLine sSQL & LINE
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & LINE & SQL_Statement_End, 1
        End If
        If withColumnDescription Then
            Output_WriteLine sqlAddDescription & LINE & SQL_Statement_End, 2
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
                                            sSQL)
        Output_WriteLine sSQL & LINE
    Next
    
    Output_Copy
End Sub

Private Sub Get_SQL_CreateTableIfNotExists_Single(objLogicalTable As clsLogicalTable, _
                                ByRef sqlCreateTable As String)
    Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                False, _
                                sqlCreateTable, _
                                "", _
                                "", _
                                True)
End Sub

Public Sub Get_SQL_CreateTable_Single(objLogicalTable As clsLogicalTable, _
                                ByVal withColumnDescription As Boolean, _
                                ByRef sqlCreateTable As String, _
                                ByRef sqlCreateFK As String, _
                                ByRef sqlAddDescription As String, _
                                ByVal IfNotExists As Boolean)
    Dim i                   As Integer
    Dim syntaxTable         As String
    
    syntaxTable = "CREATE TABLE {6:if not exists} " & ObjectNameStart & "{0:table name}" & ObjectNameEnd & " (" _
                & "{1:columns definition}" _
                & "{2:primary key}" _
                & "{3:unique indexes}" _
                & "{4:foreign keys}" _
                & LINE & ")" & SQL_Statement_End _
                & "{5:non unique indexes}"
    
    '-- Create Columns sql
    Dim syntaxColumn        As String
    Dim syntaxDefault       As String
    Dim sqlCoumn            As String
        
    syntaxColumn = "  {0:(i = 1 ? space : ,)}" & ObjectNameStart & "{1:column name}" & ObjectNameEnd & " {2:data type} {3:nullable} {4:default}"
    syntaxDefault = "  CONSTRAINT {0:default name} DEFAULT ({1:default value})"

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
    
    '-- Primary key SQL
    Dim syntaxPrimaryKey    As String
    Dim sqlPrimaryKey       As String
    
    syntaxPrimaryKey = "  ,CONSTRAINT {0:primary key name} PRIMARY KEY ({1:columns})"
    sqlPrimaryKey = ""
    With objLogicalTable.PrimaryKey
        If Len(.PKcolumns) > 0 Then
            sqlPrimaryKey = LINE & FormatString(syntaxPrimaryKey, _
                                SQL_Render_PK_Name(objLogicalTable), _
                                 .PKcolumns)
        End If
    End With
    
    '-- Unqiue Indexes SQL
    Dim syntaxUniqueIndex   As String
    Dim sqlUniqueIndex      As String
    
    syntaxUniqueIndex = "  ,CONSTRAINT {0:index name} UNIQUE {1:columns with brackets}"
    sqlUniqueIndex = ""
    For i = 1 To objLogicalTable.Indexes.Count
        With objLogicalTable.Indexes(i)
            If .isUnique Then
                sqlUniqueIndex = sqlUniqueIndex _
                        & LINE & FormatString(syntaxUniqueIndex, _
                                    SQL_Render_IK_Name(objLogicalTable, objLogicalTable.Indexes(i)), _
                                    .IKColumns)
            End If
        End With
    Next

    '-- non-unique Indexes SQL
    Dim syntaxNoUniqueIndex As String
    Dim sqlNoUniqueIndex    As String
    
    syntaxNoUniqueIndex = "CREATE INDEX {0:index name} ON " _
                    & ObjectNameStart & "{1:table name}" & ObjectNameEnd _
                    & " {2:columns with brackets}" & SQL_Statement_End
    sqlNoUniqueIndex = ""
    For i = 1 To objLogicalTable.Indexes.Count
        With objLogicalTable.Indexes(i)
            If Not .isUnique Then
                sqlNoUniqueIndex = sqlNoUniqueIndex _
                        & LINE & FormatString(syntaxNoUniqueIndex, _
                                    SQL_Render_IK_Name(objLogicalTable, objLogicalTable.Indexes(i)), _
                                    objLogicalTable.tableName, _
                                    .IKColumns)
            End If
        End With
    Next
    
    '-- Generate Foreign key SQL
    sqlCreateFK = ""    '-- SQLite does not support Add FK in Alter Table statement.
    Dim syntaxForeignKey    As String
    Dim sqlForeignKey       As String
    
    syntaxForeignKey = "  ,CONSTRAINT {0:foreign key name} FOREIGN KEY ({1:column[,..n]}) REFERENCES {2:ref_info}"
    sqlForeignKey = ""
    For i = 1 To objLogicalTable.ForeignKeys.Count
        With objLogicalTable.ForeignKeys(i)
            sqlForeignKey = sqlForeignKey _
                            & LINE & FormatString(syntaxForeignKey, _
                                        SQL_Render_FK_Name(objLogicalTable, objLogicalTable.ForeignKeys(i)), _
                                        .FKcolumns, _
                                        .RefTableAndColumns & IIf(Len(.fkOption) = 0, "", " " & .fkOption))
        End With
    Next
    
    '-- Generate table sql
    sqlCreateTable = FormatString(syntaxTable, _
                        objLogicalTable.tableName, _
                        sqlCoumn, _
                        sqlPrimaryKey, _
                        sqlUniqueIndex, _
                        sqlForeignKey, _
                        sqlNoUniqueIndex, _
                        IIf(IfNotExists, "IF NOT EXISTS ", ""))
    
'    YN: NOT Support
'   '-- Column description
'    Dim syntaxColumnDescription   As String
'    syntaxColumnDescription = "EXECUTE sp_addextendedproperty N'MS_Description', N'{0:column note}', N'user', N'dbo', N'table', N'{1:table name}', N'column', N'{2:column name}'"
'    sqlAddDescription = ""
'
'    If withColumnDescription Then
'        For i = 1 To objLogicalTable.Columns.Count
'            With objLogicalTable.Columns(i)
'                sqlAddDescription = sqlAddDescription _
'                        & LINE & FormatString(syntaxColumnDescription, _
'                            .Note, _
'                            objLogicalTable.tableName, _
'                            .columnName)
'            End With
'        Next
'    End If
    
End Sub

Private Function Get_SQL_DropTable_Single(objLogicalTable As clsLogicalTable) As String
    Dim sSQL            As String
    
    sSQL = "DROP TABLE IF EXISTS " & ObjectNameStart & "{0:table name}" & ObjectNameEnd & SQL_Statement_End
   
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
