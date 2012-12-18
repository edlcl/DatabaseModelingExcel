Attribute VB_Name = "basImport_PostgreSQL"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

Public SQL_SERVER_NAME As String
Public SQL_USER_NAME As String
Public SQL_SERVER_DATABASE_NAME As String
Public SQL_SERVER_TABLE_NAME As String

Public Function CreateConnection(ByVal server As String, _
                    ByVal database As String, _
                    ByVal user As String, _
                    ByVal Password As String) As ADODB.Connection
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection

    user = Trim(user)

    conn.ConnectionString = "Provider=PostgreSQL OLE DB Provider" _
            & ";User ID=" & user _
            & ";password=" & Password _
            & ";location=" & database _
            & ";Data Source=" & Trim(server)

    Set CreateConnection = conn
End Function

Public Function GetDatabasesSQL() As String
    Dim syntax As String
    
    syntax = "  SELECT datname AS name" _
    & LINE & "    FROM pg_database" _
    & LINE & "ORDER BY datname"
    
    GetDatabasesSQL = syntax
End Function

Public Function GetTablesSQL() As String
    Dim syntax As String
    
    syntax = "  SELECT table_name AS name" _
    & LINE & "    FROM information_schema.tables" _
    & LINE & "   WHERE table_type = 'BASE TABLE'" _
    & LINE & "ORDER BY table_name"
    
    GetTablesSQL = syntax
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
    
    syntax = "   SELECT c.oid" _
    & LINE & "        , con.conname" _
    & LINE & "        , con.contype" _
    & LINE & "        , con.conkey" _
    & LINE & "     FROM pg_namespace AS nsp" _
    & LINE & "     JOIN pg_class AS c" _
    & LINE & "       ON c.relnamespace = nsp.oid" _
    & LINE & "      AND c.relname = '{0:table name}'" _
    & LINE & "     JOIN pg_constraint AS con" _
    & LINE & "       ON con.connamespace = nsp.oid" _
    & LINE & "      AND con.conrelid = c.oid" _
    & LINE & "      AND con.contype  IN ('p', 'u')" _
    & LINE & "    WHERE nsp.nspname = current_schema()" _
    & LINE & " ORDER BY c.oid" _
    & LINE & "        , con.conname"
'    & LINE & "   SELECT ind.indrelid" _
'    & LINE & "        , ind.indexrelid" _
'    & LINE & "        , ind.indisunique" _
'    & LINE & "        , ind.indisprimary" _
'    & LINE & "        , ind.indisclustered"
'        , ind.indkey
'     FROM pg_namespace AS nsp
'     JOIN pg_class AS c
'       ON c.relnamespace = nsp.oid
'      AND c.relname = 'customercustomerdemo'
'     JOIN pg_index AS ind
'       ON ind.indrelid = c.oid
'    WHERE nsp.nspname = current_schema()
' ORDER BY ind.indrelid
'        , ind.indexrelid
'
'        SELECT att.attnum
'        , att.attname
'     FROM pg_namespace AS nsp
'     JOIN pg_class AS c
'       ON c.relnamespace = nsp.oid
'      AND c.relname = 'customercustomerdemo'
'     JOIN pg_attribute AS att
'       ON att.attrelid = c.oid
'      AND att.attnum > 0
'    WHERE nsp.nspname = current_schema()
' ORDER BY att.attrelid
'        , att.attnum
    Dim sSQL                    As String
    sSQL = FormatString(syntax, objTable.tableName)
    
    Dim oRs                     As ADODB.Recordset
    Dim curIndexID              As Integer
    Dim objIndex                As clsLogicalIndex

    On Error GoTo Flag_Err

    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    curIndexID = 0

    Do While Not oRs.EOF
        If oRs("isPK") = 1 Then
            '-- Primary Key
            If Len(objTable.PrimaryKey.PKcolumns) = 0 Then
                objTable.PrimaryKey.PKcolumns = oRs("ColName") & ""
            Else
                objTable.PrimaryKey.PKcolumns = objTable.PrimaryKey.PKcolumns & ", " & oRs("ColName")
            End If

            objTable.PrimaryKey.IsClustered = (oRs("IsCluster") = 1)
        Else
            '-- Index
            If curIndexID <> oRs("indid") Then
                Set objIndex = New clsLogicalIndex
                objTable.Indexes.Add objIndex
                
                objIndex.IsClustered = (oRs("IsCluster") = 1)
                objIndex.isUnique = (oRs("IsUnique") = 1)

                curIndexID = oRs("indid")
            End If

            If Len(objIndex.IKColumns) = 0 Then
                objIndex.IKColumns = oRs("ColName") & ""
            Else
                objIndex.IKColumns = objIndex.IKColumns & ", " & oRs("ColName")
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
    
    syntax = "   SELECT c.oid" _
    & LINE & "        , con.conkey" _
    & LINE & "        , cf.relname" _
    & LINE & "        , con.confkey" _
    & LINE & "        , con.confupdtype" _
    & LINE & "        , con.confdeltype" _
    & LINE & "        , con.confmatchtype" _
    & LINE & "     FROM pg_namespace AS nsp" _
    & LINE & "     JOIN pg_class AS c" _
    & LINE & "       ON c.relnamespace = nsp.oid" _
    & LINE & "      AND c.relname = '{0:table name}'" _
    & LINE & "     JOIN pg_constraint AS con" _
    & LINE & "       ON con.conrelid = c.oid" _
    & LINE & "      AND con.contype = 'f'" _
    & LINE & "     JOIN pg_class AS cf" _
    & LINE & "       ON cf.relnamespace = nsp.oid" _
    & LINE & "      AND cf.oid = con.confrelid" _
    & LINE & "    WHERE nsp.nspname = current_schema()" _
    & LINE & " ORDER BY c.oid" _
    & LINE & "        , con.conkey"
    
    Dim sSQL                    As String
    sSQL = FormatString(syntax, objTable.tableName)
    
    Dim oRs             As ADODB.Recordset
    Dim curFKID         As Long
    Dim objForeignKey   As clsLogicalForeignKey
    
    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    curFKID = 0

    Do While Not oRs.EOF
        '-- For Foreign Key
        If curFKID <> oRs("id") Then
            Set objForeignKey = New clsLogicalForeignKey
            objTable.ForeignKeys.Add objForeignKey

            objForeignKey.refTableName = oRs("rtName") & ""
            objForeignKey.OnDelete = IIf(oRs("OnDelete") = 1, "ON DELETE CASCADE", "")
            objForeignKey.OnUpdate = IIf(oRs("OnUpdate") = 1, "ON UPDATE CASCADE", "")
            
            curFKID = oRs("id")
        End If

        If Len(objForeignKey.FKcolumns) > 0 Then
            objForeignKey.FKcolumns = objForeignKey.FKcolumns & ", "
        End If
        objForeignKey.FKcolumns = objForeignKey.FKcolumns & oRs("ftColumnName")
        
        If Len(objForeignKey.RefTableColumns) > 0 Then
            objForeignKey.RefTableColumns = objForeignKey.RefTableColumns & ", "
        End If
        objForeignKey.RefTableColumns = objForeignKey.RefTableColumns & oRs("rtColumnName")

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
    
    syntax = "     SELECT c.table_name" _
    & LINE & "       , c.column_name" _
    & LINE & "       , c.ordinal_position" _
    & LINE & "       , c.column_default" _
    & LINE & "       , c.is_nullable" _
    & LINE & "       , c.data_type" _
    & LINE & "       , c.character_maximum_length" _
    & LINE & "       , c.numeric_precision" _
    & LINE & "       , c.numeric_scale" _
    & LINE & "       , c.column_default" _
    & LINE & "       , col_description(o.oid, c.ordinal_position) as column_comment" _
    & LINE & "    FROM information_schema.columns as c" _
    & LINE & "    JOIN pg_class as o" _
    & LINE & "      ON o.relname = c.table_name" _
    & LINE & "    JOIN pg_namespace AS nsp" _
    & LINE & "      ON nsp.nspname = current_schema()" _
    & LINE & "     AND nsp.oid = c.relnamespace" _
    & LINE & "   WHERE c.table_catalog = current_database()" _
    & LINE & "     AND c.table_schema  = current_schema()" _
    & LINE & "     AND c.table_name    = '{0:table name}'" _
    & LINE & "ORDER BY c.table_name" _
    & LINE & "       , c.ordinal_position"

    Dim sSQL                    As String
    sSQL = FormatString(syntax, objTable.tableName)
    
    Dim oRs             As ADODB.Recordset
    Dim objColumn       As clsLogicalColumn
    
    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    Do While Not oRs.EOF
        '-- set Column
        Set objColumn = New clsLogicalColumn
        objTable.Columns.Add objColumn
        
        objColumn.columnName = oRs("column_name") & ""
        objColumn.dataType = GetColumnDataType( _
                                            oRs("data_type"), _
                                            oRs("character_maximum_length"), _
                                            oRs("numeric_precision"), _
                                            oRs("numeric_scale"))
        objColumn.Nullable = oRs("is_nullable")
        objColumn.Default = oRs("column_default") & ""

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
                        maxLength As Integer, _
                        precision As Integer, _
                        type_scale As Integer) As String
    Dim dataType As String

    dataType = LCase(typeName)
    Select Case dataType
        Case "char", "varchar", "binary", "varbinary"
            dataType = dataType & "(" & maxLength & ")"
        Case "nvarchar", "nchar"
            dataType = dataType & "(" & maxLength / 2 & ")"
        Case "numeric"
            dataType = dataType & "(" & precision & "," & type_scale & ")"
    End Select

    GetColumnDataType = dataType
End Function



