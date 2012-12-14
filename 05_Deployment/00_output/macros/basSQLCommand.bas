Attribute VB_Name = "basSQLCommand"
Option Explicit

'-- Copy create all tables SQL to clipboard
Public Sub CopyAllCreateTableSQL()
    Select Case The_Excel_Type
    Case The_Excel_Type_Multiple
        Call basSQL_SQLServer.Get_SQL_CreateTable(GetAllLogicalTables(), False)
    Case DBName_SQLServer
        Call basSQL_SQLServer.Get_SQL_CreateTable(GetAllLogicalTables(), False)
    Case DBName_MySQL
        Call basSQL_MySQL.Get_SQL_CreateTable(GetAllLogicalTables(), False)
    Case DBName_Oracle
        Call basSQL_Oracle.Get_SQL_CreateTable(GetAllLogicalTables(), False)
    End Select
End Sub

'-- Copy create all tables with description SQL to clipboard
Public Sub CopyAllCreateTableWithDescriptionSQL()
    Select Case The_Excel_Type
    Case The_Excel_Type_Multiple
        Call basSQL_SQLServer.Get_SQL_CreateTable(GetAllLogicalTables(), True)
    Case DBName_SQLServer
        Call basSQL_SQLServer.Get_SQL_CreateTable(GetAllLogicalTables(), True)
    Case DBName_MySQL
        Call basSQL_MySQL.Get_SQL_CreateTable(GetAllLogicalTables(), True)
    Case DBName_Oracle
        Call basSQL_Oracle.Get_SQL_CreateTable(GetAllLogicalTables(), True)
    End Select
End Sub

'-- Copy drop all tables SQL to clipboard
Public Sub CopyAllDropTableSQL()
    Select Case The_Excel_Type
    Case The_Excel_Type_Multiple
        Call basSQL_SQLServer.Get_SQL_DropTable(GetAllLogicalTables())
    Case DBName_SQLServer
        Call basSQL_SQLServer.Get_SQL_DropTable(GetAllLogicalTables())
    Case DBName_MySQL
        Call basSQL_MySQL.Get_SQL_DropTable(GetAllLogicalTables())
    Case DBName_Oracle
        Call basSQL_Oracle.Get_SQL_DropTable(GetAllLogicalTables())
    End Select
End Sub

'-- Copy drop and create all tables SQL to clipboard
Public Sub CopyAllDropAndCreateTableSQL()
    Select Case The_Excel_Type
    Case The_Excel_Type_Multiple
        Call basSQL_SQLServer.Get_SQL_DropAndCreateTable(GetAllLogicalTables(), False)
    Case DBName_SQLServer
        Call basSQL_SQLServer.Get_SQL_DropAndCreateTable(GetAllLogicalTables(), False)
    Case DBName_MySQL
        Call basSQL_MySQL.Get_SQL_DropAndCreateTable(GetAllLogicalTables(), False)
    Case DBName_Oracle
        Call basSQL_Oracle.Get_SQL_DropAndCreateTable(GetAllLogicalTables(), False)
    End Select
End Sub

'-- Copy create all exits tables SQL to clipboard
Public Sub CopyAllCreateTableIfNotExistsSQL()
    Select Case The_Excel_Type
    Case The_Excel_Type_Multiple
        Call basSQL_SQLServer.Get_SQL_CreateTableIfNotExists(GetAllLogicalTables())
    Case DBName_SQLServer
        Call basSQL_SQLServer.Get_SQL_CreateTableIfNotExists(GetAllLogicalTables())
    Case DBName_MySQL
        Call basSQL_MySQL.Get_SQL_CreateTableIfNotExists(GetAllLogicalTables())
    Case DBName_Oracle
        Call basSQL_Oracle.Get_SQL_CreateTableIfNotExists(GetAllLogicalTables())
    End Select
End Sub

'-- Copy drop and create all tables with description SQL to clipboard
Public Sub CopyAllDropAndCreateTableWithDescriptionSQL()
    Select Case The_Excel_Type
    Case The_Excel_Type_Multiple
        Call basSQL_SQLServer.Get_SQL_DropAndCreateTable(GetAllLogicalTables(), True)
    Case DBName_SQLServer
        Call basSQL_SQLServer.Get_SQL_DropAndCreateTable(GetAllLogicalTables(), True)
    Case DBName_MySQL
        Call basSQL_MySQL.Get_SQL_DropAndCreateTable(GetAllLogicalTables(), True)
    Case DBName_Oracle
        Call basSQL_Oracle.Get_SQL_DropAndCreateTable(GetAllLogicalTables(), True)
    End Select
End Sub
