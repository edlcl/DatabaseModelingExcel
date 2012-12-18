VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBOracleProvider"
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
'-- Orcale
'-----------------------------------------------------------
Private Const SP_REMVOE_FK_NAME = "tmp_dbmodelexcel_drop_table_fk"
Private Const MAX_OBJECT_NAME_LENGTH = 30
Private p_colObjectNames As Collection
'DBMS_OUTPUT.PUT_LINE('');
'sys/123@xiws53 as sysdba
'-------------------------------------

Public Sub clsIDatabaseProvider_GetSQLCreateTable(objLogicalTableCollection As Collection, _
                            ByVal withColumnDescription As Boolean)
    Dim sSQL                As String
    Dim sqlCreateFK         As String
    Dim sqlAddDescription   As String
    Dim objLogicalTable     As clsLogicalTable
    
    Set p_colObjectNames = New Collection
    Output_Initialize
    
    '-- Create tables
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                            withColumnDescription, _
                                            sSQL, _
                                            sqlCreateFK, _
                                            sqlAddDescription)
        Output_WriteLine sSQL & LINE
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & LINE, 1
        End If
        If withColumnDescription Then
            Output_WriteLine sqlAddDescription & LINE, 2
        End If
    Next
    
    Output_Copy
    Set p_colObjectNames = Nothing
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropTable(objLogicalTableCollection As Collection)
    Dim sSQL            As String
    Dim objLogicalTable As clsLogicalTable
    
    Output_Initialize
    
    '-- Create the store procedue of dropping foreign key
    sSQL = Get_SQL_Pre_RemoveFK()
    Output_WriteLine sSQL & LINE & "/"
    
    '-- Drop foreign key relation
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_RemoveFK_Single(objLogicalTable)
        Output_WriteLine sSQL & LINE
    Next
    
    '-- Drop the store procedue of dropping foreign key
    sSQL = Get_SQL_Post_RemoveFK()
    Output_WriteLine sSQL & LINE
    
    '-- Drop tables
    Output_WriteLine Get_SQL_DropTable_Single_Pre
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_DropTable_Single(objLogicalTable)
        Output_WriteLine sSQL & LINE
    Next
    Output_WriteLine Get_SQL_DropTable_Single_Post
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropAndCreateTable(objLogicalTableCollection As Collection, _
                            ByVal withColumnDescription As Boolean)
    Dim sSQL                As String
    Dim sqlCreateFK         As String
    Dim sqlAddDescription   As String
    Dim objLogicalTable     As clsLogicalTable
    
    Output_Initialize
    
    '-- Create the store procedue of dropping foreign key
    sSQL = Get_SQL_Pre_RemoveFK()
    Output_WriteLine sSQL & LINE & "/"
    
    '-- Drop foreign key relation
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_RemoveFK_Single(objLogicalTable)
        Output_WriteLine sSQL & LINE
    Next
    
    '-- Drop the store procedue of dropping foreign key
    sSQL = Get_SQL_Post_RemoveFK()
    Output_WriteLine sSQL & LINE
    
    '-- Drop tables
    Output_WriteLine Get_SQL_DropTable_Single_Pre
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_DropTable_Single(objLogicalTable)
        Output_WriteLine sSQL & LINE
    Next
    Output_WriteLine Get_SQL_DropTable_Single_Post
    
    '-- Create tables
    Set p_colObjectNames = New Collection
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                            withColumnDescription, _
                                            sSQL, _
                                            sqlCreateFK, _
                                            sqlAddDescription)
        Output_WriteLine sSQL & LINE
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & LINE, 1
        End If
        If withColumnDescription Then
            Output_WriteLine sqlAddDescription & LINE, 2
        End If
    Next
    
    Output_Copy
    Set p_colObjectNames = Nothing
End Sub

Public Sub clsIDatabaseProvider_GetSQLCreateTableIfNotExists(objLogicalTableCollection As Collection)
    Dim sSQL                As String
    Dim sqlCreateFK         As String
    Dim objLogicalTable     As clsLogicalTable
    
    Set p_colObjectNames = New Collection
    Output_Initialize
    
    Output_WriteLine Get_SQL_CreateTableIfNotExists_Single_Pre
    
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTableIfNotExists_Single(objLogicalTable, _
                                            sSQL, _
                                            sqlCreateFK)
        Output_WriteLine sSQL & LINE
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & LINE, 1
        End If
    Next
    
    Output_WriteLine Get_SQL_CreateTableIfNotExists_Single_Post
    
    Output_Copy
    Set p_colObjectNames = Nothing
End Sub

Private Sub Get_SQL_CreateTableIfNotExists_Single(objLogicalTable As clsLogicalTable, _
                                ByRef sqlCreateTable As String, _
                                ByRef sqlCreateFK As String)
    Dim sSQL            As String
    
    sSQL = "  SELECT COUNT(*) INTO v_table_is_exists" _
        & LINE & "  FROM user_tables" _
        & LINE & "  WHERE lower(table_name) = lower('{0:table name}');" _
        & LINE & "  IF v_table_is_exists != 0 THEN" _
        & LINE & "    execute immediate {2:create table sql}" _
        & LINE & "  END IF;"
    
    Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                False, _
                                sqlCreateTable, _
                                sqlCreateFK, _
                                "")
                                
    sqlCreateTable = FormatString(sSQL, _
                    objLogicalTable.tableName, _
                    objLogicalTable.Description, _
                    Replace(sqlCreateTable, "'", "''"))
    
End Sub

Private Function Get_SQL_CreateTableIfNotExists_Single_Pre() As String
    Dim sSQL            As String
    
    sSQL = "DECLARE" _
        & LINE & "  v_table_is_exists integer;" _
        & LINE & "BEGIN" _

    '-- Return
    Get_SQL_CreateTableIfNotExists_Single_Pre = sSQL
End Function

Private Function Get_SQL_CreateTableIfNotExists_Single_Post() As String
    Dim sSQL            As String
    
    sSQL = "END;" & LINE & "/"
    
    '-- Return
    Get_SQL_CreateTableIfNotExists_Single_Post = sSQL
End Function

Private Function Get_SQL_RemoveFK_Single(objLogicalTable As clsLogicalTable) As String
    Dim syntaxRemoveFK      As String
    
    syntaxRemoveFK = "-- Remove {0:table name} foreign key constraint" _
            & LINE & "CALL {1:sp name for remove FK}('{0:table name}');" _
            
    '-- Return
    Get_SQL_RemoveFK_Single = FormatString(syntaxRemoveFK, objLogicalTable.tableName, SP_REMVOE_FK_NAME)
End Function

Private Function Get_SQL_Pre_RemoveFK() As String
    Dim sSQL        As String
    
    sSQL = "-- temporary store procedue for remove foreign key" _
            & LINE & "CREATE or REPLACE PROCEDURE {0:sp name for remove FK}(" _
            & LINE & "    a_table_name IN VARCHAR2" _
            & LINE & ") IS" _
            & LINE & "  v_fk_name varchar2(250);" _
            & LINE & "  CURSOR c_fk IS" _
            & LINE & "    SELECT UC.constraint_name" _
            & LINE & "      FROM user_constraints UC" _
            & LINE & "     WHERE lower(UC.table_name) = lower(a_table_name)" _
            & LINE & "       AND UC.constraint_type = 'R';" _
            & LINE & "BEGIN" _
            & LINE & "" _
            & LINE & "  OPEN c_fk;" _
            & LINE & "  LOOP" _
            & LINE & "    FETCH c_fk INTO v_fk_name;" _
            & LINE & "    EXIT WHEN c_fk%NOTFOUND;" _
            & LINE & "      EXECUTE IMMEDIATE 'ALTER TABLE ' || a_table_name || ' DROP CONSTRAINT ' || v_fk_name;" _
            & LINE & "    END LOOP;" _
            & LINE & "  CLOSE c_fk;" _
            & LINE & "END {0:sp name for remove FK};"
    '-- Return
    Get_SQL_Pre_RemoveFK = FormatString(sSQL, SP_REMVOE_FK_NAME)
End Function

Private Function Get_SQL_Post_RemoveFK() As String
    Dim sSQL        As String
    
    sSQL = "-- Remove temporary store procedue for remove foreign key" _
            & LINE & "DROP PROCEDURE {0:sp name for remove FK};" _
            
    '-- Return
    Get_SQL_Post_RemoveFK = FormatString(sSQL, SP_REMVOE_FK_NAME)
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
    '-- Create table
    'create table TEST
    '(
    '  ID   NUMBER not null,
    '  NAME VARCHAR2(20) default 'y'
    ')
    ';
    syntaxTable = "CREATE TABLE {0:table name} (" _
                & "{1:columns definition}" _
                & LINE & ");" _
                & LINE & "{2: primary key}" _
                & LINE & "{3: index}" _
                & LINE & "{4: non-unique index}"
    syntaxColumn = "  {0:(i = 1 ? space : ,)}{1:column name} {2:data type} {3:nullable} {4:default}"
    syntaxDefault = " DEFAULT {0:default value}"
    
    syntaxPrimaryKey = "ALTER TABLE {0:table name} ADD CONSTRAINT {1:primary key name} PRIMARY KEY ({2:columns});"
    syntaxUniqueIndex = "ALTER TABLE {0:table name} ADD CONSTRAINT {1:unique index name} UNIQUE {2:columns with bracket};"
    syntaxNoUniqueIndex = "CREATE INDEX {1:index name} ON {0:table name} {2:columns};"
    
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
                                    .Default))
        End With
    Next
    
    '-- Primary key SQL
    sqlPrimaryKey = ""
    With objLogicalTable.PrimaryKey
        If Len(.PKcolumns) > 0 Then
            sqlPrimaryKey = LINE & FormatString(syntaxPrimaryKey, _
                                objLogicalTable.tableName, _
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
                                    objLogicalTable.tableName, _
                                    SQL_Render_IK_Name(objLogicalTable, objLogicalTable.Indexes(i)), _
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
                                    objLogicalTable.tableName, _
                                    SQL_Render_IK_Name(objLogicalTable, objLogicalTable.Indexes(i)), _
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
    Dim sqlFK           As String
    sqlFK = "ALTER TABLE {0:Table Name} ADD CONSTRAINT {1:foreign key name}" _
        & LINE & "  FOREIGN KEY ({2:column[,..n]}) REFERENCES {3:ref_info};"
    
    sqlCreateFK = ""
    If objLogicalTable.ForeignKeys.Count > 0 Then
        For i = 1 To objLogicalTable.ForeignKeys.Count
            With objLogicalTable.ForeignKeys(i)
                sqlCreateFK = sqlCreateFK _
                        & LINE & FormatString(sqlFK, _
                                    objLogicalTable.tableName, _
                                    SQL_Render_FK_Name(objLogicalTable, objLogicalTable.ForeignKeys(i)), _
                                    .FKcolumns, _
                                    .RefTableAndColumns & IIf(Len(.fkOption) = 0, "", " " & .fkOption))
            End With
        Next
    End If
    
    '-- Column description
    Dim syntaxTableComment          As String
    syntaxTableComment = "COMMENT ON TABLE {0:table name} '{1:table comment}';"
    Dim syntaxColumnComment   As String
    syntaxColumnComment = "COMMENT ON COLUMN {0:table name}.{1:column name} IS '{2:column note}';"
    
    sqlAddDescription = FormatString(syntaxTableComment, _
                                objLogicalTable.tableName, _
                                objLogicalTable.Description)
    
    If withColumnDescription Then
        For i = 1 To objLogicalTable.Columns.Count
            With objLogicalTable.Columns(i)
                sqlAddDescription = sqlAddDescription _
                        & LINE & FormatString(syntaxColumnComment, _
                            objLogicalTable.tableName, _
                            .columnName _
                            .Note)
            End With
        Next
    End If
    
End Sub

Private Function Get_SQL_DropTable_Single(objLogicalTable As clsLogicalTable) As String
    Dim sSQL            As String
    
    sSQL = "  SELECT COUNT(*) INTO v_table_is_exists" _
        & LINE & "  FROM user_tables" _
        & LINE & "  WHERE lower(table_name) = lower('{0:table name}');" _
        & LINE & "  IF v_table_is_exists != 0 THEN" _
        & LINE & "    execute immediate 'DROP TABLE {0:table name} PURGE';" _
        & LINE & "  END IF;" _
   
    '-- Return
    Get_SQL_DropTable_Single = FormatString(sSQL, _
                                objLogicalTable.tableName, _
                                objLogicalTable.Description)
End Function

Private Function Get_SQL_DropTable_Single_Pre() As String
    Dim sSQL            As String
    
    sSQL = "DECLARE" _
        & LINE & "  v_table_is_exists integer;" _
        & LINE & "BEGIN"

    '-- Return
    Get_SQL_DropTable_Single_Pre = sSQL
End Function

Private Function Get_SQL_DropTable_Single_Post() As String
    Dim sSQL            As String
    
    sSQL = "END;" & LINE & "/" & LINE
    
    '-- Return
    Get_SQL_DropTable_Single_Post = sSQL
End Function

Private Function SQL_Render_TableName(ByVal tableName As String) As String
    SQL_Render_TableName = Replace(Replace(tableName, " ", ""), "_", "")
End Function

Private Function SQL_Render_PK_Name(objLogicalTable As clsLogicalTable) As String
    SQL_Render_PK_Name = RenderObjectName( _
                            "PK_" & SQL_Render_TableName(objLogicalTable.tableName) _
                            )
End Function

Private Function SQL_Render_FK_Name(objLogicalTable As clsLogicalTable, _
                                objLogicalForeignKey As clsLogicalForeignKey) As String
    SQL_Render_FK_Name = RenderObjectName("FK_" & SQL_Render_TableName(objLogicalTable.tableName) _
                            & "_" & Replace(Replace(Replace(objLogicalForeignKey.FKcolumns, " ", ""), "_", ""), ",", "_") _
                            )
End Function

Private Function SQL_Render_IK_Name(objLogicalTable As clsLogicalTable, _
                                objLogicalIndex As clsLogicalIndex) As String
    SQL_Render_IK_Name = RenderObjectName("IK_" & SQL_Render_TableName(objLogicalTable.tableName) _
                            & "_" & Replace(Replace(Replace(Replace(Replace(objLogicalIndex.IKColumns, _
                                                                    " ", ""), _
                                                            "(", ""), _
                                                    ")", ""), _
                                            "_", ""), _
                                    ",", "_") _
                        )
                                    
End Function

Private Function SQL_Render_DF_Name(objLogicalTable As clsLogicalTable, _
                                objLogicalColumn As clsLogicalColumn) As String
    SQL_Render_DF_Name = RenderObjectName( _
                    "DF_" & SQL_Render_TableName(objLogicalTable.tableName) & "_" & objLogicalColumn.columnName _
                    )
                    
End Function

Private Function RenderObjectName(ByVal ObjectName As String) As String
    Dim newObjectName           As String
    Dim objectNameIndex         As Integer
    newObjectName = ObjectName
    
    If Len(newObjectName) > MAX_OBJECT_NAME_LENGTH Then
        newObjectName = Left(newObjectName, MAX_OBJECT_NAME_LENGTH - 4)
        On Error Resume Next
        objectNameIndex = p_colObjectNames.item(newObjectName)
        If Err.Number <> 0 Then
            Err.Clear
            objectNameIndex = 0
            Call p_colObjectNames.Add(objectNameIndex, newObjectName)
        Else
            objectNameIndex = objectNameIndex + 1
            p_colObjectNames.item(newObjectName) = objectNameIndex
        End If
        newObjectName = newObjectName & Format(objectNameIndex, "0000")
    End If
    
    '-- Return
    RenderObjectName = newObjectName
End Function


