VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImport_SQLServer 
   Caption         =   "Import"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6060
   OleObjectBlob   =   "frmImport_SQLServer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImport_SQLServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

'=============================
'           SQL Server
'=============================
Private Sub btnRefreshDatabase_Click()
    Dim conn As ADODB.Connection
    Dim sSQL As String
    Dim oRs As ADODB.Recordset
    Dim lastIndex As Integer
    Dim index As Integer

    If Len(Trim(cboSQLServer.text)) = 0 Then
        MsgBox "Please input SQL Server.", vbOKOnly + vbInformation, APP_NAME
        cboSQLServer.SetFocus
        Exit Sub
    End If

    On Error GoTo Flag_Err

    Set conn = basImport_SQLServer.CreateConnection(Trim(cboSQLServer.text), _
                            "master", _
                            Trim(txtUser.text), _
                            Trim(txtPassword.text))
    conn.Open
    sSQL = "SELECT [name] FROM sysdatabases ORDER BY [name]"

    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    lastIndex = -1
    index = -1
    cboDatabase.Clear
    Do While Not oRs.EOF
        index = index + 1
        cboDatabase.AddItem (oRs("name"))
        If oRs("name") = basImport_SQLServer.SQL_SERVER_DATABASE_NAME Then
            lastIndex = index
        End If
        '-- Move next record
        oRs.MoveNext
    Loop

    If cboDatabase.ListCount > 0 Then
        If lastIndex > 0 Then
            cboDatabase.ListIndex = lastIndex
        Else
            cboDatabase.ListIndex = 0
        End If
    End If
    '-- Close record set
    oRs.Close
    conn.Close
    Set oRs = Nothing
    Set conn = Nothing

    '-- Set public variant
    basImport_SQLServer.SQL_SERVER_NAME = cboSQLServer.text
    Exit Sub
Flag_Err:
    If Not oRs Is Nothing Then oRs.Close
    Set oRs = Nothing
    Set conn = Nothing

    Call MsgBoxEx_Error
End Sub

Private Sub btnRefreshTable_Click()
    Dim conn As ADODB.Connection
    Dim sSQL As String
    Dim oRs As ADODB.Recordset
    Dim lastIndex As Integer
    Dim index As Integer
    Dim sTableName As String

    On Error GoTo Flag_Err
    If Cells.Item(Table_Sheet_Row_TableName, _
                    Table_Sheet_Col_TableName).text = "" Then
        sTableName = basImport_SQLServer.SQL_SERVER_TABLE_NAME
    Else
        sTableName = Cells.Item(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).text
    End If

    Set conn = basImport_SQLServer.CreateConnection(Trim(cboSQLServer.text), _
                            Trim(cboDatabase.text), _
                            Trim(txtUser.text), _
                            Trim(txtPassword.text))
    conn.Open
    sSQL = "SELECT [name] FROM sysobjects WHERE xtype='U' ORDER BY [name]"

    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    lastIndex = -1
    index = -1
    lstTables.Clear
    Do While Not oRs.EOF
        index = index + 1
        lstTables.AddItem (oRs("name"))
        If oRs("name") = sTableName Then
            lastIndex = index
        End If
        '-- Move next record
        oRs.MoveNext
    Loop

    If lstTables.ListCount > 0 Then
        If lastIndex > 0 Then
            lstTables.ListIndex = lastIndex
        Else
            lstTables.ListIndex = 0
        End If
    End If

    '-- Set public variant
    basImport_SQLServer.SQL_SERVER_DATABASE_NAME = cboDatabase.text
    '-- Close record set
    oRs.Close
    conn.Close
    Set oRs = Nothing
    Set conn = Nothing

    Exit Sub
Flag_Err:
    If Not oRs Is Nothing Then oRs.Close
    If Not conn Is Nothing Then conn.Close
    Set oRs = Nothing
    Set conn = Nothing

    Call MsgBoxEx_Error
End Sub

Private Sub Init()
    Dim iActiveSheet As Integer
    Dim index As Integer
    Dim iSheet As Integer

    Dim shtCurrent As Worksheet
    On Error GoTo FALG_ERR

    '-- Init
    If Len(basImport_SQLServer.SQL_SERVER_NAME) = 0 Then
        cboSQLServer.text = ""
    Else
        cboSQLServer.text = basImport_SQLServer.SQL_SERVER_NAME
    End If
    cboSQLServer.SetFocus

    '-- init combo sheet
    cboSheet.Clear
    index = 0
    iActiveSheet = -1
    iActiveSheet = ThisWorkbook.ActiveSheet.index - Sheet_First_Table + 1
    For iSheet = Sheet_First_Table To ThisWorkbook.Sheets.Count
        Set shtCurrent = ThisWorkbook.Sheets(iSheet)
        '-- Set Caption = index & tablecaption
        cboSheet.AddItem shtCurrent.Name
        cboSheet.List(index, 1) = shtCurrent.index

        If ThisWorkbook.ActiveSheet.index = shtCurrent.index Then
            iActiveSheet = index
        End If
        index = index + 1
    Next

    If cboSheet.ListCount > 0 Then
        If iActiveSheet >= 0 Then
            cboSheet.ListIndex = iActiveSheet
        Else
            cboSheet.ListIndex = 0
        End If

    End If

FALG_ERR:
    Set shtCurrent = Nothing
End Sub

Private Sub SQLServerImport()
    On Error GoTo Flag_Err

    Dim index               As Integer
    Dim shtTemplate         As Worksheet
    Dim shtCurrent          As Worksheet
    Dim currentSheetIndex   As Integer
    Dim conn                As ADODB.Connection
    Dim clearSheet          As Boolean
    Dim tableName           As String
    Dim table               As clsLogicalTable
    Dim isSetPublicVarient  As Boolean
    isSetPublicVarient = False

    If cboSheet.ListIndex < 0 Then Exit Sub

    clearSheet = (Me.chkClearSheet.value)
    Set conn = basImport_SQLServer.CreateConnection(Trim(cboSQLServer.text), _
                            cboDatabase.text, _
                            Trim(txtUser.text), _
                            Trim(txtPassword.text))
    conn.Open
    
    currentSheetIndex = CInt(cboSheet.List(cboSheet.ListIndex, 1))
    Set shtTemplate = ThisWorkbook.Sheets(currentSheetIndex)
    
    For index = 0 To Me.lstTables.ListCount - 1
        If lstTables.selected(index) Then
            '-- Get Table Definition
            tableName = lstTables.List(index, 0)
            
            '-- Get a sheet which is used to store table information
            
            If Me.chkAutoMatch.value Then
                Set shtCurrent = GetSheetFromTableName(tableName)
                clearSheet = False
                If shtCurrent Is Nothing Then
                    Set shtCurrent = CopyASheet(shtTemplate, , ThisWorkbook.Sheets(currentSheetIndex))
                    currentSheetIndex = shtCurrent.index
                    clearSheet = True
                End If
            Else
                If Me.chkCreateASheet.value Then
                    currentSheetIndex = ThisWorkbook.Sheets.Count
                    Set shtCurrent = CopyASheet(shtTemplate, , ThisWorkbook.Sheets(currentSheetIndex))
                Else
                    Set shtCurrent = ThisWorkbook.Sheets(currentSheetIndex)
                End If
            End If
            
            '-- Set public variant
            If isSetPublicVarient = False Then
                basImport_SQLServer.SQL_SERVER_TABLE_NAME = tableName
                isSetPublicVarient = True
            End If
            Set table = basImport_SQLServer.GetLogicalTable(conn, tableName)
            '-- Write to sheet
            shtCurrent.Select
            Call basTableSheet.SetTableInfoToWorksheet(shtCurrent, table, clearSheet)
        End If
    Next

    '-- Close connection
    conn.Close
    Set conn = Nothing
    
    Exit Sub
Flag_Err:
    Set conn = Nothing

    Call MsgBoxEx_Error
End Sub

Private Sub btnSelectAllTable_Click()
    Call SelectAllListBoxItems(lstTables, True)
End Sub

Private Sub btnUnSelectAllTable_Click()
    Call SelectAllListBoxItems(Me.lstTables, False)
End Sub

Private Sub cboDatabase_Enter()
    Call SelectAllText(cboDatabase)
End Sub

Private Sub cboSheet_Enter()
    Call SelectAllText(cboSheet)
End Sub

Private Sub cboSQLServer_Enter()
    Call SelectAllText(cboSQLServer)
End Sub

Private Sub chkAutoMatch_Change()
    If Me.chkAutoMatch.value Then
        Me.chkCreateASheet.value = True
        Me.chkClearSheet.value = True
    End If
End Sub

Private Sub txtPassword_Enter()
    Call SelectAllText(txtPassword)
End Sub

Private Sub txtUser_Enter()
    Call SelectAllText(txtUser)
End Sub

Private Sub btnImport_Click()
    Call SQLServerImport
End Sub

Private Sub UserForm_Initialize()
    Call Init
End Sub

Private Sub chkCreateASheet_Change()
    If Me.chkCreateASheet.value Then
        Me.chkClearSheet.value = True
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
