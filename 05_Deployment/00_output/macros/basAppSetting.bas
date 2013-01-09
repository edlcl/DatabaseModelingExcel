Attribute VB_Name = "basAppSetting"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit
Public Const APP_NAME                   As String = "Database Modeling Excel"
Public Const APP_VERSION                As String = "5.0.1"

'-- Sheet part
Public Const Sheet_Index                        As Integer = 1      '-- Index of index sheet
Public Const Sheet_Update_History               As Integer = 2      '-- Index of update history sheet
Public Const Sheet_Rule                         As Integer = 3      '-- Index of rule sheet
Public Const Sheet_First_Table                  As Integer = 4      '-- Index of first table sheet

Public Const Table_Sheet_Row_TableComment       As Integer = 2
Public Const Table_Sheet_Col_TableComment       As Integer = 7
Public Const Table_Sheet_Row_TableName          As Integer = 3
Public Const Table_Sheet_Col_TableName          As Integer = 7
Public Const Table_Sheet_Row_PrimaryKey         As Integer = 4
Public Const Table_Sheet_Col_PrimaryKey         As Integer = 7
Public Const Table_Sheet_Row_ForeignKey         As Integer = 5
Public Const Table_Sheet_Col_ForeignKey         As Integer = 7
Public Const Table_Sheet_Row_Index              As Integer = 6
Public Const Table_Sheet_Col_Index              As Integer = 7
Public Const Table_Sheet_Col_Clustered          As Integer = 48
Public Const Table_Sheet_Col_Unique             As Integer = 49
Public Const Table_Sheet_Row_First_Column       As Integer = 9
Public Const Table_Sheet_Col_ColumnID           As Integer = 2
Public Const Table_Sheet_Col_ColumnLabel        As Integer = 3
Public Const Table_Sheet_Col_ColumnName         As Integer = 11
Public Const Table_Sheet_Col_ColumnDataType     As Integer = 22
Public Const Table_Sheet_Col_ColumnNullable     As Integer = 28
Public Const Table_Sheet_Col_ColumnDefault      As Integer = 31
Public Const Table_Sheet_Col_ColumnComment      As Integer = 34
Public Const Table_Sheet_Row_TableStatus        As Integer = 1
Public Const Table_Sheet_Col_TableStatus        As Integer = 2

'-- Table Sheet Value
Public Const Table_Sheet_PK_Clustered           As String = ""
Public Const Table_Sheet_PK_NonClustered        As String = "N"
Public Const Table_Sheet_Index_Clustered        As String = "Y"
Public Const Table_Sheet_Index_NonClustered     As String = ""
Public Const Table_Sheet_Index_Unique           As String = ""
Public Const Table_Sheet_Index_NonUnique        As String = "N"
Public Const Table_Sheet_Nullable               As String = "Yes"
Public Const Table_Sheet_NonNullable            As String = "No"
Public Const Table_Sheet_TableStatus_Ignore     As String = "ignore"

'-- UI
Public Const Table_Sheet_Row_Height             As Integer = 12.75
Public Const Table_Code_Length                  As Integer = 0
Public Const Sheet_NameIsTableDesc              As Boolean = False

'-- Marks
Public Const LINE                               As String = vbCrLf

'-- Databae Type Global variable
Public Const DBName_DB2                         As String = "DB2"
Public Const DBName_MySQL                       As String = "MySQL"
Public Const DBName_Oracle                      As String = "Oracle"
Public Const DBName_PostgreSQL                  As String = "PostgreSQL"
Public Const DBName_SQLite                      As String = "SQLite"
Public Const DBName_SQLServer                   As String = "SQL Server"

'----------- Excel Type Global variable ---------------
Public Const The_Excel_Type_Multiple            As String = "M"
'-- the constant's value will be one of THE_EXCEL_TYPE_MULTIPLE, DBName_SQLServer, DBName_MySQL, or DBName_XXX
Public Const The_Excel_Type                     As String = The_Excel_Type_Multiple
