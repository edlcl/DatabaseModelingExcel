Attribute VB_Name = "basReverse_MySQL"
Option Explicit

'-------------------------------------------
'-- My SQL Reverse Module
'-------------------------------------------
Public DRIVER_NAME As String
Public PORT_ID As String
Public SERVER_NAME As String
Public SERVER_DATABASE_NAME As String
Public SERVER_TABLE_NAME As String

Public Function CreateConnection(ByVal server As String, _
                    ByVal database As String, _
                    ByVal user As String, _
                    ByVal password As String, _
                    ByVal driver As String, _
                    ByVal port As String) As ADODB.Connection
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection

    If Len(driver) = 0 Then driver = "{MySQL ODBC 5.1 Driver}"
    If Len(port) = 0 Then port = "3306"
    
    conn.ConnectionString = "Driver=" & Trim(driver) _
            & ";Server=" & Trim(server) _
            & ";Port=" & Trim(port) _
            & ";Database=" & database _
            & ";User=" & Trim(user) _
            & ";Password=" & password _
            & ";OPTION=3"
    
    Set CreateConnection = conn
End Function

Public Function GetLogicalTable(conn As ADODB.Connection, tableName As String) As clsLogicalTable
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
    Set GetLogicalTable = objTable
End Function

Public Sub RenderPKAndIndex(conn As ADODB.Connection, objTable As clsLogicalTable)
    Dim syntax As String
    
    syntax = "   SELECT S.TABLE_NAME" _
    & LINE & "        , S.INDEX_NAME" _
    & LINE & "        , S.SEQ_IN_INDEX" _
    & LINE & "        , S.COLUMN_NAME" _
    & LINE & "        , S.NON_UNIQUE" _
    & LINE & "        , TC.CONSTRAINT_TYPE" _
    & LINE & "     FROM information_schema.STATISTICS S" _
    & LINE & "LEFT JOIN information_schema.TABLE_CONSTRAINTS TC" _
    & LINE & "       ON S.TABLE_SCHEMA = TC.TABLE_SCHEMA" _
    & LINE & "      AND S.TABLE_NAME = TC.TABLE_NAME" _
    & LINE & "      AND S.INDEX_NAME = TC.CONSTRAINT_NAME" _
    & LINE & "    WHERE S.TABLE_SCHEMA = DATABASE()" _
    & LINE & "      AND S.TABLE_NAME = {0:table name}" _
    & LINE & " ORDER BY S.TABLE_NAME" _
    & LINE & "        , S.INDEX_NAME" _
    & LINE & "        , S.SEQ_IN_INDEX;"

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
        If oRs("CONSTRAINT_TYPE") & "" = "PRIMARY KEY" Then
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
                objIndex.IsUnique = (oRs("NON_UNIQUE") = 0)

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
    
    syntax = "SELECT R.TABLE_NAME" _
    & LINE & "     , R.CONSTRAINT_NAME" _
    & LINE & "     , R.UPDATE_RULE" _
    & LINE & "     , R.DELETE_RULE" _
    & LINE & "     , R.REFERENCED_TABLE_NAME" _
    & LINE & "     , K.COLUMN_NAME" _
    & LINE & "     , K.ORDINAL_POSITION" _
    & LINE & "     , K.POSITION_IN_UNIQUE_CONSTRAINT" _
    & LINE & "     , K.REFERENCED_COLUMN_NAME" _
    & LINE & "  FROM information_schema.REFERENTIAL_CONSTRAINTS R" _
    & LINE & "  JOIN information_schema.KEY_COLUMN_USAGE K" _
    & LINE & "    ON R.CONSTRAINT_SCHEMA = K.CONSTRAINT_SCHEMA" _
    & LINE & "   AND R.TABLE_NAME        = K.TABLE_NAME" _
    & LINE & "   AND R.CONSTRAINT_NAME   = K.CONSTRAINT_NAME" _
    & LINE & " WHERE R.CONSTRAINT_SCHEMA = DATABASE()" _
    & LINE & "   AND R.TABLE_NAME = {0:table name}" _
    & LINE & " ORDER BY R.TABLE_NAME" _
    & LINE & "     , R.CONSTRAINT_NAME" _
    & LINE & "     , K.ORDINAL_POSITION;"

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

            objForeignKey.refTableName = oRs("REFERENCED_TABLE_NAME")
            If oRs("DELETE_RULE") <> "RESTRICT" Then
                objForeignKey.OnDelete = "ON DELETE " & oRs("DELETE_RULE")
            Else
                objForeignKey.OnDelete = ""
            End If
            If oRs("UPDATE_RULE") <> "RESTRICT" Then
                objForeignKey.OnUpdate = "ON DELETE " & oRs("UPDATE_RULE")
            Else
                objForeignKey.OnUpdate = ""
            End If
            
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
    & LINE & "       , C.ORDINAL_POSITION" _
    & LINE & "       , C.COLUMN_TYPE" _
    & LINE & "       , C.COLUMN_DEFAULT" _
    & LINE & "       , C.EXTRA" _
    & LINE & "       , C.IS_NULLABLE" _
    & LINE & "       , C.COLUMN_COMMENT" _
    & LINE & "       , C.DATA_TYPE" _
    & LINE & "       , C.CHARACTER_MAXIMUM_LENGTH" _
    & LINE & "       , C.NUMERIC_PRECISION" _
    & LINE & "       , C.NUMERIC_SCALE" _
    & LINE & "    FROM information_schema.`COLUMNS` C" _
    & LINE & "   WHERE C.TABLE_SCHEMA = DATABASE()" _
    & LINE & "     AND C.TABLE_NAME = {0:table name}" _
    & LINE & "ORDER BY C.TABLE_NAME" _
    & LINE & "       , C.ORDINAL_POSITION;"

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
        
        objColumn.columnName = oRs("COLUMN_NAME")
        objColumn.dataType = GetColumnDataType( _
                                            oRs("COLUMN_TYPE"), _
                                            oRs("EXTRA") & "")
        objColumn.Nullable = (oRs("IS_NULLABLE") = "YES")
        objColumn.Default = oRs("COLUMN_DEFAULT") & ""

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

Private Function GetColumnDataType(columnType As String, _
                        extra As String) As String
    Dim dataType As String

    dataType = LCase(columnType)
    If (Len(extra) > 0) Then
        dataType = dataType & " " & extra
    End If
    GetColumnDataType = dataType
End Function
