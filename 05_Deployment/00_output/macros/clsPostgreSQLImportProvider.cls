VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPostgreSQLImportProvider"
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
Implements IImportProvider

Const ConnParaDatabase = "Database"
Const ProviderOleDb = "PostgreSQL OLE DB Provider"
Const ConnectionStringFormatOleDb = "Provider={0: provider}" _
            & ";User ID={2: user}" _
            & ";password={3: password}" _
            & ";location={4: database}" _
            & ";Data Source={1: server}"
Const ProviderOdbc = "{PostgreSQL ODBC Driver(UNICODE)}"
Const ConnectionStringFormatOdbc = "Driver={0: provider}" _
            & ";Uid={2: user}" _
            & ";Pwd={3: password}" _
            & ";Database={4: database}" _
            & ";Server={1: server}" _
            & ";Port=5432"
            
Public Property Get IImportProvider_ConnectionParameters() As Collection
    Dim parameters As Collection
    Set parameters = New Collection
    Call parameters.Add(ConnParaDatabase)
    
    Set IImportProvider_ConnectionParameters = parameters
End Property

Public Property Get IImportProvider_Providers() As String()
    Dim candidateProviders(1 To 2) As String
    candidateProviders(1) = ProviderOdbc
    candidateProviders(2) = ProviderOleDb
    
    IImportProvider_Providers = candidateProviders
End Property

Public Property Get IImportProvider_SupportSelectDatabase() As Boolean
    IImportProvider_SupportSelectDatabase = True
End Property

Public Function IImportProvider_GetDatabasesSql() As String
    Dim syntax As String
    
    syntax = "  SELECT datname AS name" _
    & LINE & "    FROM pg_database" _
    & LINE & "ORDER BY datname"
    
    IImportProvider_GetDatabasesSql = syntax
End Function

Public Function IImportProvider_GetTablesSql() As String
    Dim syntax As String
    
    syntax = "  SELECT table_name AS name" _
    & LINE & "    FROM information_schema.tables" _
    & LINE & "   WHERE table_type = 'BASE TABLE'" _
    & LINE & "     AND table_schema = 'public'" _
    & LINE & "ORDER BY table_name"
    
    IImportProvider_GetTablesSql = syntax
End Function

Public Function IImportProvider_GetOptions() As clsImportOptions
    If basImport.mImportOptions Is Nothing Then
        Set basImport.mImportOptions = New clsImportOptions
        basImport.mImportOptions.DataSource = "localhost"
        basImport.mImportOptions.DataSourceTip = "e.g. localhost"
        basImport.mImportOptions.UserName = "postgres"
        basImport.mImportOptions.Database = "postgres"
    End If
    Set IImportProvider_GetOptions = basImport.mImportOptions
End Function

Public Function IImportProvider_CreateConnectionString(ByVal Provider As String, _
                    ByVal server As String, _
                    ByVal user As String, _
                    ByVal Password As String, _
                    ByVal Database As String) As String
                    
    Dim connectionStringFormat As String
    Provider = Trim(Provider)
    If LCase(Provider) = LCase(ProviderOdbc) Then
        IImportProvider_CreateConnectionString = FormatString(ConnectionStringFormatOdbc, _
                        Provider, _
                        Trim(server), _
                        Trim(user), _
                        Password, _
                        IIf(Len(Database) = 0, "postgres", Database))
    Else
        IImportProvider_CreateConnectionString = FormatString(ConnectionStringFormatOleDb, _
                        Provider, _
                        Trim(server), _
                        Trim(user), _
                        Password, _
                        Database)
    End If
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
    syntax = "   SELECT c.oid" _
    & LINE & "        , i.indexrelid as indid" _
    & LINE & "        , i.indisunique as IsUnique" _
    & LINE & "        , i.indisprimary as isPK" _
    & LINE & "        , i.indisclustered as IsCluster" _
    & LINE & "        , i.indkey" _
    & LINE & "        , i.indattind" _
    & LINE & "        , a.attnum" _
    & LINE & "        , a.attname as ColName" _
    & LINE & "     FROM pg_namespace AS nsp" _
    & LINE & "     JOIN pg_class AS c" _
    & LINE & "       ON c.relnamespace = nsp.oid" _
    & LINE & "      AND c.relname = '{0:table name}'" _
    & LINE & "     JOIN (" _
    & LINE & "          SELECT indrelid" _
    & LINE & "           , indexrelid" _
    & LINE & "           , indisunique" _
    & LINE & "           , indisprimary" _
    & LINE & "           , indisclustered" _
    & LINE & "           , indkey" _
    & LINE & "           , generate_subscripts(indkey, 1) as indattind" _
    & LINE & "           FROM pg_index" _
    & LINE & "           ) i" _
    & LINE & "       ON i.indrelid = c.oid" _
    & LINE & "     JOIN pg_attribute a"
    syntax = syntax _
    & LINE & "       ON a.attrelid = c.oid" _
    & LINE & "      AND a.attnum > 0" _
    & LINE & "      AND a.attnum = i.indkey[indattind]" _
    & LINE & "    WHERE nsp.nspname = current_schema()" _
    & LINE & " ORDER BY c.oid" _
    & LINE & "        , i.indexrelid" _
    & LINE & "        , i.indattind"

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
    Set oRs = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub RenderForeignKey(conn As ADODB.Connection, objTable As clsLogicalTable)
    Dim syntax As String
    
    syntax = "   SELECT c.oid" _
    & LINE & "        , con.conname as id" _
    & LINE & "        , con.conkey" _
    & LINE & "        , con.confkey" _
    & LINE & "        , con.confupdtype as OnUpdate" _
    & LINE & "        , con.confdeltype as OnDelete" _
    & LINE & "        , con.confmatchtype" _
    & LINE & "        , a.attnum" _
    & LINE & "        , a.attname as ftColumnName" _
    & LINE & "        , cf.relname as rtName" _
    & LINE & "        , af.attnum" _
    & LINE & "        , af.attname as rtColumnName" _
    & LINE & "     FROM pg_namespace AS nsp" _
    & LINE & "     JOIN pg_class AS c" _
    & LINE & "       ON c.relnamespace = nsp.oid" _
    & LINE & "      AND c.relname = '{0:table name}'" _
    & LINE & "     JOIN (" _
    & LINE & "          SELECT conname" _
    & LINE & "               , conrelid" _
    & LINE & "               , confrelid" _
    & LINE & "               , contype" _
    & LINE & "               , conkey" _
    & LINE & "               , confkey" _
    & LINE & "               , confupdtype"
    syntax = syntax _
    & LINE & "               , confdeltype" _
    & LINE & "               , confmatchtype" _
    & LINE & "               , generate_subscripts(conkey, 1) as conattind" _
    & LINE & "           FROM pg_constraint" _
    & LINE & "          WHERE contype = 'f'" _
    & LINE & "           ) AS con" _
    & LINE & "       ON con.conrelid = c.oid" _
    & LINE & "     JOIN pg_attribute a" _
    & LINE & "       ON a.attrelid = c.oid" _
    & LINE & "      AND a.attnum = con.conkey[conattind]" _
    & LINE & "     JOIN pg_class AS cf" _
    & LINE & "       ON cf.relnamespace = nsp.oid" _
    & LINE & "      AND cf.oid = con.confrelid" _
    & LINE & "     JOIN pg_attribute af" _
    & LINE & "       ON af.attrelid = c.oid" _
    & LINE & "      AND af.attnum = con.confkey[conattind]" _
    & LINE & "    WHERE nsp.nspname = current_schema()" _
    & LINE & " ORDER BY c.oid" _
    & LINE & "        , con.conrelid" _
    & LINE & "        , con.conattind"
    
    Dim sSQL                    As String
    sSQL = FormatString(syntax, objTable.tableName)
    
    Dim oRs             As ADODB.Recordset
    Dim curFKID         As String
    Dim objForeignKey   As clsLogicalForeignKey
    
    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    curFKID = ""

    Do While Not oRs.EOF
        '-- For Foreign Key
        If curFKID <> oRs("id") Then
            Set objForeignKey = New clsLogicalForeignKey
            objTable.ForeignKeys.Add objForeignKey

            objForeignKey.refTableName = oRs("rtName") & ""
            objForeignKey.OnDelete = GetForeignKeyAction(oRs("OnDelete"), "ON DELETE")
            objForeignKey.OnUpdate = GetForeignKeyAction(oRs("OnUpdate"), "ON UPDATE")
            
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
    
    syntax = "   SELECT c.oid" _
    & LINE & "        , a.attnum" _
    & LINE & "        , a.attname as column_name" _
    & LINE & "        , not a.attnotnull as is_nullable" _
    & LINE & "        , t.typname as data_type" _
    & LINE & "        , a.atttypmod - 4 as character_maximum_length" _
    & LINE & "        , a.atttypmod / 65536  as numeric_precision" _
    & LINE & "        , a.atttypmod % 65535 - a.atttypmod / 65536 - 4 as numeric_scale" _
    & LINE & "        , ad.adsrc as column_default" _
    & LINE & "     FROM pg_namespace AS nsp" _
    & LINE & "     JOIN pg_class AS c" _
    & LINE & "       ON c.relnamespace = nsp.oid" _
    & LINE & "      AND c.relname = '{0:table name}'" _
    & LINE & "     JOIN pg_attribute a" _
    & LINE & "       ON a.attrelid = c.oid" _
    & LINE & "      AND a.attnum > 0" _
    & LINE & "     JOIN pg_type t" _
    & LINE & "       ON t.oid = a.atttypid" _
    & LINE & "LEFT JOIN pg_attrdef ad" _
    & LINE & "       ON ad.adrelid = c.oid" _
    & LINE & "      AND ad.adnum = a.attnum" _
    & LINE & "    WHERE nsp.nspname = current_schema()" _
    & LINE & " ORDER BY c.oid" _
    & LINE & "        , a.attnum"

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
        'TODO: objColumn.Comment = oRs("column_comment") & ""
        
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
                        maxLength As Long, _
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
            If precision > 0 And type_scale >= 0 Then
                dataType = dataType & "(" & precision & "," & type_scale & ")"
            End If
    End Select

    GetColumnDataType = dataType
End Function

Private Function GetForeignKeyAction(action As String, _
            keyType As String)
    '-- a = no action, r = restrict, c = cascade, n = set null, d = set default
    Dim foreignKeyAction As String
    
    Select Case action
    Case "r"
        foreignKeyAction = keyType & " RESTRICT"
    Case "c"
        foreignKeyAction = keyType & " CASCADE"
    Case "n"
        foreignKeyAction = keyType & " SET NULL"
    Case "d"
        foreignKeyAction = keyType & " SET DEFAULT"
    End Select
    
    GetForeignKeyAction = foreignKeyAction
End Function


