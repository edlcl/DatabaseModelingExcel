Attribute VB_Name = "basSQLCommand"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

'-- Copy create all tables SQL to clipboard
Public Sub CopyAllCreateTableSQL()
    Call basPublicDatabase.GetDatabaseProvider().GetSQLCreateTable(GetAllLogicalTables(), False)
End Sub

'-- Copy create all tables with comment SQL to clipboard
Public Sub CopyAllCreateTableWithDescriptionSQL()
    Call basPublicDatabase.GetDatabaseProvider().GetSQLCreateTable(GetAllLogicalTables(), True)
End Sub

'-- Copy drop all tables SQL to clipboard
Public Sub CopyAllDropTableSQL()
    Call basPublicDatabase.GetDatabaseProvider().GetSQLDropTable(GetAllLogicalTables())
End Sub

'-- Copy drop and create all tables SQL to clipboard
Public Sub CopyAllDropAndCreateTableSQL()
    Call basPublicDatabase.GetDatabaseProvider().GetSQLDropAndCreateTable(GetAllLogicalTables(), False)
End Sub

'-- Copy create all exits tables SQL to clipboard
Public Sub CopyAllCreateTableIfNotExistsSQL()
    Call basPublicDatabase.GetDatabaseProvider().GetSQLCreateTableIfNotExists(GetAllLogicalTables())
End Sub

'-- Copy drop and create all tables with comment SQL to clipboard
Public Sub CopyAllDropAndCreateTableWithDescriptionSQL()
    Call basPublicDatabase.GetDatabaseProvider().GetSQLDropAndCreateTable(GetAllLogicalTables(), True)
End Sub
