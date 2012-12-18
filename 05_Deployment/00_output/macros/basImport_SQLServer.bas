Attribute VB_Name = "basImport_SQLServer"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

Public SQL_SERVER_NAME As String
Public SQL_SERVER_DATABASE_NAME As String
Public SQL_SERVER_TABLE_NAME As String

Public Function CreateConnection(ByVal server As String, _
                    ByVal database As String, _
                    ByVal user As String, _
                    ByVal Password As String) As ADODB.Connection
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection

    user = Trim(user)

    If Len(user) = 0 Then
        conn.ConnectionString = "Provider=SQLOLEDB" _
                & ";Integrated Security=SSPI" _
                & ";initial catalog=" & database _
                & ";Data Source=" & Trim(server) _
                & ";Connect Timeout=30"
    Else
        conn.ConnectionString = "Provider=SQLOLEDB" _
                & ";UID=" & user _
                & ";PWD=" & Password _
                & ";initial catalog=" & database _
                & ";Data Source=" & Trim(server) _
                & ";Connect Timeout=30"
    End If

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
    
    syntax = "     SELECT sysindexes.indid" _
    & LINE & "          , IndexName = sysindexes.name" _
    & LINE & "          , IsPK = CASE WHEN sysobjects.xtype='PK' THEN 1" _
    & LINE & "                   ELSE 0" _
    & LINE & "                   End" _
    & LINE & "          , IsCluster = INDEXPROPERTY(O.id, sysindexes.[name], 'IsClustered')" _
    & LINE & "          , IsUnique = INDEXPROPERTY(O.id, sysindexes.[name], 'IsUnique')" _
    & LINE & "          , ColName = syscolumns.[name]" _
    & LINE & "       FROM (SELECT id FROM sysobjects WHERE NAME = '{0:table name}') O" _
    & LINE & "       JOIN sysindexes" _
    & LINE & "         ON   sysindexes.id = O.id " _
    & LINE & "        AND  sysindexes.[name] NOT LIKE '_WA_Sys%'" _
    & LINE & "       JOIN sysindexkeys" _
    & LINE & "         ON sysindexkeys.id = sysindexes.id" _
    & LINE & "        AND sysindexkeys.indid = sysindexes.indid" _
    & LINE & "       JOIN syscolumns" _
    & LINE & "         ON syscolumns.id = sysindexkeys.id" _
    & LINE & "        AND syscolumns.colid = sysindexkeys.colid" _
    & LINE & "  LEFT JOIN sysobjects" _
    & LINE & "         ON sysobjects.parent_obj = O.id" _
    & LINE & "        AND sysobjects.[name] = sysindexes.[name]" _
    & LINE & "        AND sysobjects.xtype='PK'" _
    & LINE & "   ORDER BY sysindexes.indid"

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
    
    syntax = "     SELECT id = sysforeignkeys.constid" _
    & LINE & "          , ftColumnName = fkc.[name]" _
    & LINE & "          , rtName = rt.[name]" _
    & LINE & "          , rtColumnName = rc.[name]" _
    & LINE & "          , OnDelete = OBJECTPROPERTY(sysforeignkeys.constid, 'CnstIsDeleteCascade')" _
    & LINE & "          , OnUpdate = OBJECTPROPERTY(sysforeignkeys.constid, 'CnstIsUpdateCascade')" _
    & LINE & "       FROM sysobjects fkt" _
    & LINE & "       JOIN sysforeignkeys" _
    & LINE & "         ON fkt.id = sysforeignkeys.fkeyid" _
    & LINE & "       JOIN syscolumns fkc" _
    & LINE & "         ON fkc.id = fkt.id" _
    & LINE & "        AND fkc.colid = sysforeignkeys.fkey" _
    & LINE & "       JOIN sysobjects rt" _
    & LINE & "         ON rt.id = sysforeignkeys.rkeyid" _
    & LINE & "       JOIN syscolumns rc" _
    & LINE & "         ON rc.id = rt.id" _
    & LINE & "        AND rc.colid = sysforeignkeys.rkey" _
    & LINE & "      WHERE fkt.name = '{0:table name}'" _
    & LINE & "        AND fkt.xtype='U'" _
    & LINE & "   ORDER BY sysforeignkeys.constid" _
    & LINE & "          , sysforeignkeys.keyno"

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
    
    syntax = "    SELECT syscolumns.[name]" _
    & LINE & "         , type_name = systypes.name" _
    & LINE & "         , length = syscolumns.length" _
    & LINE & "         , [precision] = syscolumns.xprec" _
    & LINE & "         , scale =syscolumns.xscale" _
    & LINE & "         , is_identity = ISNULL(COLUMNPROPERTY(sysobjects.id, syscolumns.[name], 'IsIdentity'), 0)" _
    & LINE & "         , identity_incr = IDENT_INCR(sysobjects.[name])" _
    & LINE & "         , identity_seed = IDENT_SEED(sysobjects.[name])" _
    & LINE & "         , is_nullable = syscolumns.isnullable" _
    & LINE & "         , default_definition = syscomments.text" _
    & LINE & "      FROM sysobjects" _
    & LINE & "      JOIN syscolumns" _
    & LINE & "        ON syscolumns.id = sysobjects.id" _
    & LINE & " LEFT JOIN syscomments" _
    & LINE & "        ON syscomments.id = syscolumns.cdefault" _
    & LINE & " LEFT JOIN systypes" _
    & LINE & "        ON systypes.xusertype = syscolumns.xusertype" _
    & LINE & "     WHERE sysobjects.name = '{0:table name}'" _
    & LINE & "       AND sysobjects.xtype='U'" _
    & LINE & "  ORDER BY syscolumns.colid"

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
        
        objColumn.columnName = oRs("name") & ""
        objColumn.dataType = GetColumnDataType( _
                                            oRs("type_name"), _
                                            oRs("length"), _
                                            oRs("precision"), _
                                            oRs("scale"), _
                                            oRs("is_identity"), _
                                            IIf(IsNull(oRs("identity_seed")), 0, oRs("identity_seed")), _
                                            IIf(IsNull(oRs("identity_incr")), 0, oRs("identity_incr")))
        objColumn.Nullable = oRs("is_nullable")
        If Not IsNull(oRs("default_definition")) Then
            objColumn.Default = oRs("default_definition")
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
                        maxLength As Integer, _
                        precision As Integer, _
                        type_scale As Integer, _
                        is_identity As Integer, _
                        seed As Integer, _
                        incr As Integer) As String
    Dim dataType As String

    dataType = LCase(typeName)
    Select Case dataType
        Case "char", "varchar", "binary", "varbinary"
            dataType = dataType & "(" & maxLength & ")"
        Case "nvarchar", "nchar"
            dataType = dataType & "(" & maxLength / 2 & ")"
        Case "int", "bigint"
            If is_identity Then
                dataType = dataType & " IDENTITY (" & seed & "," & incr & ")"
            End If
        Case "numeric"
            dataType = dataType & "(" & precision & "," & type_scale & ")"
    End Select

    GetColumnDataType = dataType
End Function

