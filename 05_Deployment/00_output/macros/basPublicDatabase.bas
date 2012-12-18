Attribute VB_Name = "basPublicDatabase"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

'---------------------------------------------------------------------
'-- Cet connection string form the OLE DB Links dialog.
'---------------------------------------------------------------------
Public Function GetConnectionString(connString As String)
    Dim conn As ADODB.Connection
    Dim MSDASCObj As MSDASC.DataLinks
    
    Set MSDASCObj = New MSDASC.DataLinks
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = connString
    
    MSDASCObj.PromptEdit conn

    GetConnectionString = conn.ConnectionString
End Function

Public Function GetDatabaseProvider(Optional DatabaseType As String) As clsIDatabaseProvider
    If IsMissing(DatabaseType) Then
        DatabaseType = The_Excel_Type
    End If
    
    Select Case DatabaseType
    Case The_Excel_Type_Multiple
        Set GetDatabaseProvider = New clsDBSQLServerProvider
    Case DBName_MySQL
        Set GetDatabaseProvider = New clsDBMySQLProvider
    Case DBName_Oracle
        Set GetDatabaseProvider = New clsDBOracleProvider
    Case DBName_PostgreSQL
        Set GetDatabaseProvider = New clsDBPostgreSQLProvider
    Case DBName_SQLite
        Set GetDatabaseProvider = New clsDBSQLiteProvider
    Case DBName_SQLServer
        Set GetDatabaseProvider = New clsDBSQLServerProvider
    End Select
End Function
