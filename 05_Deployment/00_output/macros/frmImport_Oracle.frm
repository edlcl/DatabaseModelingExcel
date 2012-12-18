VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImport_Oracle 
   Caption         =   "Import"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   OleObjectBlob   =   "frmImport_Oracle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImport_Oracle"
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

'-----------------------------------------------------------
'           Oracle
'-----------------------------------------------------------
Option Explicit

Const PageConnectIndex      As Integer = 0
Const PageTablesIndex       As Integer = 1
Const PageOptionIndex       As Integer = 2
Const ConnectText           As String = "Connect >"
Const NextText              As String = "Next >"
Const ImportText            As String = "Import"

Private SkipShowTablesStatisticsInfo As Boolean

Private mDatabaseType As String
Public Property Get DatabaseType() As String
        DatabaseType = mDatabaseType
End Property
Public Property Let DatabaseType(value As String)
        mDatabaseType = value
End Property

Private Sub SetWizardStatus(newPageValue As Integer)
    Dim index As Integer
    
    For index = 0 To Me.MultiPageMain.Pages.Count - 1
        If index = newPageValue Then
            Me.MultiPageMain.Pages(index).Enabled = True
        Else
            Me.MultiPageMain.Pages(index).Enabled = False
        End If
    Next
    Me.MultiPageMain.value = newPageValue
    
    If Me.MultiPageMain.value = 0 Then
        btnPrevious.Enabled = False
    Else
        btnPrevious.Enabled = True
    End If

    If Me.MultiPageMain.value = PageConnectIndex Then
        btnNext.Caption = ConnectText
    ElseIf Me.MultiPageMain.value = PageOptionIndex Then
        btnNext.Caption = ImportText
    Else
        btnNext.Caption = ConnectText
    End If
End Sub

Private Sub DoConnect()
    Dim conn As ADODB.Connection
    Dim sSQL As String
    Dim oRs As ADODB.Recordset
    Dim lastIndex As Integer
    Dim index As Integer
    Dim sTableName As String

    On Error GoTo Flag_Err
    
    If Cells.item(Table_Sheet_Row_TableName, _
                    Table_Sheet_Col_TableName).text = "" Then
        sTableName = basImport_Oracle.GetOptions().LastAccessTableName
    Else
        sTableName = Cells.item(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).text
    End If

    Set conn = New Connection
    If Me.MultiPageConnection.value = 0 Then
        conn.ConnectionString = basImport_Oracle.CreateConnectionString(Trim(cboProvider.text), _
                            Trim(txtServer.text), _
                            Trim(txtUser.text), _
                            Trim(txtPassword.text))
    Else
        conn.ConnectionString = txtConnectionString.text
    End If
    
    conn.Open
    
    sSQL = "SELECT * FROM User_Tables T" _
        & LINE & "ORDER BY T.TABLE_NAME"

    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn

    lastIndex = -1
    index = -1
    lstTables.Clear
    Do While Not oRs.EOF
        index = index + 1
        lstTables.AddItem (oRs("TABLE_NAME"))
        If oRs("TABLE_NAME") = sTableName Then
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

    Call ShowTablesInfo
    '-- Close record set
    oRs.Close
    conn.Close
    Set oRs = Nothing
    Set conn = Nothing

    Call SaveConnectionOptions
    Exit Sub
Flag_Err:
    Set oRs = Nothing
    Set conn = Nothing

    Call MsgBoxEx_Error
End Sub

Private Sub DoImport()
    On Error GoTo Flag_Err

    Dim index               As Integer
    Dim shtTemplate         As Worksheet
    Dim shtCurrent          As Worksheet
    Dim insertSheetIndex    As Integer
    Dim conn                As ADODB.Connection
    Dim tableName           As String
    Dim table               As clsLogicalTable
    Dim isSetPublicVarient  As Boolean
    isSetPublicVarient = False

    If cboSheet.ListIndex < 0 Then Exit Sub

    Dim clearExistedData    As Boolean
    clearExistedData = Me.chkClearExistedData.value
    
    Set conn = New Connection
    If Me.MultiPageConnection.value = 0 Then
        conn.ConnectionString = basImport_Oracle.CreateConnectionString(Trim(cboProvider.text), _
                            Trim(txtServer.text), _
                            Trim(txtUser.text), _
                            Trim(txtPassword.text))
    Else
        conn.ConnectionString = txtConnectionString.text
    End If
    
    conn.Open
    
    insertSheetIndex = CInt(cboSheet.List(cboSheet.ListIndex, 1))
    Set shtTemplate = ThisWorkbook.Sheets(insertSheetIndex)
    
    For index = 0 To Me.lstTables.ListCount - 1
        If lstTables.selected(index) Then
            '-- Get the importing table name
            tableName = lstTables.List(index, 0)
            
            '-- Get the sheet which is used to store the table information
            If Me.optImportModeOverwrite.value Then
                Set shtCurrent = GetSheetFromTableName(tableName)
                clearExistedData = False
                If shtCurrent Is Nothing Then
                    Set shtCurrent = CopyASheet(shtTemplate, , ThisWorkbook.Sheets(insertSheetIndex))
                    insertSheetIndex = shtCurrent.index
                    clearExistedData = True
                End If
            ElseIf Me.optImportModeAlwaysCreateSheets.value Then
                insertSheetIndex = ThisWorkbook.Sheets.Count
                Set shtCurrent = CopyASheet(shtTemplate, , ThisWorkbook.Sheets(insertSheetIndex))
            Else
                Set shtCurrent = ThisWorkbook.Sheets(insertSheetIndex)
            End If
            
            '-- Set public variant
            If isSetPublicVarient = False Then
                basImport_Oracle.GetOptions().LastAccessTableName = tableName
                isSetPublicVarient = True
            End If
            Set table = basImport_Oracle.GetLogicalTable(conn, tableName)
            '-- Write to sheet
            shtCurrent.Select
            Call basTableSheet.SetTableInfoToWorksheet(shtCurrent, table, clearExistedData)
        End If
    Next

    '-- Close connection
    conn.Close
    Set conn = Nothing
    
    Call SaveImportOptions
    Exit Sub
Flag_Err:
    Set conn = Nothing

    Call MsgBoxEx_Error
End Sub

Private Sub GenerateConnectionString()

    Me.txtConnectionString = basImport_Oracle.CreateConnectionString(Trim(cboProvider.text), _
                            Trim(txtServer.text), _
                            Trim(txtUser.text), _
                            Trim(txtPassword.text))

End Sub

Private Sub Init()
    On Error GoTo FALG_ERR
    
    Me.DatabaseType = DBName_Oracle
    Call InitForm
    Call InitConnectionPage
    Call InitOptionPage
    Call SetWizardStatus(PageConnectIndex)
    
    Exit Sub

FALG_ERR:
     Call MsgBoxEx_Error
End Sub

Private Sub InitConnectionPage()
    '-- Active connection page
    Me.MultiPageMain.value = 0
    
    '-- Active connection sub page
    If basImport_Oracle.GetOptions().ConnectionMode = ConnectionModeDataSource Then
        Me.MultiPageConnection.value = 0
        txtServer.SetFocus
    Else
        Me.MultiPageConnection.value = 1
        txtConnectionString.SetFocus
    End If
    
    '-- Init provider list
    cboProvider.Clear
    Call cboProvider.AddItem("MSDAORA.1")
    Call cboProvider.AddItem("OraOLEDB.Oracle.1")
    cboProvider.ListIndex = 0
    cboProvider.text = basImport_Oracle.GetOptions().provider

    '-- fill text box values
    txtServer.text = basImport_Oracle.GetOptions().DataSource
    txtUser.text = basImport_Oracle.GetOptions().UserName
    txtPassword.text = basImport_Oracle.GetOptions().Password
    
    If basImport_Oracle.GetOptions().ConnectionMode = ConnectionModeConnectionString Then
        txtConnectionString.text = basImport_Oracle.GetOptions().ConnectionString
    End If
End Sub

Private Sub InitOptionPage()
    Dim iActiveSheet As Integer
    Dim iSheet As Integer
    Dim index As Integer
    Dim shtCurrent As Worksheet

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
    
    '-- init importing options
    Select Case basImport_Oracle.GetOptions().ImportMode
    Case enmImportMode.ImportModeOverwrite
        Me.optImportModeOverwrite.value = True
    Case enmImportMode.ImportModeAlwaysCreateSheet
        Me.optImportModeAlwaysCreateSheets.value = True
    Case enmImportMode.ImportModeAlwaysUpdate
        Me.optImportModeOnlyUpdateTemplateSheet.value = True
    End Select
    
    Me.chkClearExistedData.value = basImport_Oracle.GetOptions().ClearDataInExistedSheet
End Sub

Private Sub InitForm()
    Me.Caption = "Import from " & DatabaseType
End Sub

Private Sub SaveConnectionOptions()
    If Me.MultiPageConnection.value = 0 Then
        basImport_Oracle.GetOptions().ConnectionMode = ConnectionModeDataSource
    Else
        basImport_Oracle.GetOptions().ConnectionMode = ConnectionModeConnectionString
    End If
    
    basImport_Oracle.GetOptions().provider = Trim(cboProvider.text)
    basImport_Oracle.GetOptions().DataSource = Trim(txtServer.text)
    basImport_Oracle.GetOptions().UserName = Trim(txtUser.text)
    basImport_Oracle.GetOptions().Password = txtPassword.text
    basImport_Oracle.GetOptions().ConnectionString = Trim(txtConnectionString.text)
End Sub

Private Sub SaveImportOptions()
    If Me.optImportModeOverwrite.value Then
        basImport_Oracle.GetOptions().ImportMode = ImportModeOverwrite
    ElseIf Me.optImportModeAlwaysCreateSheets.value Then
        basImport_Oracle.GetOptions().ImportMode = ImportModeAlwaysCreateSheet
    Else
        basImport_Oracle.GetOptions().ImportMode = ImportModeAlwaysUpdate
    End If
    
    basImport_Oracle.GetOptions().ClearDataInExistedSheet = Me.chkClearExistedData.value
End Sub


Private Sub ShowTablesInfo()
    Dim selectTableCount As Integer
    Dim index As Integer
    
    selectTableCount = 0
    For index = 0 To Me.lstTables.ListCount - 1
        If lstTables.selected(index) Then
            selectTableCount = selectTableCount + 1
        End If
    Next
    Me.labTable.Caption = "Select Tables (" & CStr(selectTableCount) & "\" & CStr(Me.lstTables.ListCount) & ")"
End Sub

Private Sub btnConnBuild_Click()
    Me.txtConnectionString.text = basPublicDatabase.GetConnectionString(Me.txtConnectionString.text)
End Sub

Private Sub btnNext_Click()
    On Error GoTo Flag_Err
    
    Select Case Me.MultiPageMain.value
    Case PageConnectIndex
        Call DoConnect
    Case PageTablesIndex
    
    Case PageOptionIndex
        Call DoImport
    End Select
    
    If Me.MultiPageMain.value < PageOptionIndex Then
        Call SetWizardStatus(Me.MultiPageMain.value + 1)
    End If
    
    Exit Sub
Flag_Err:
    Call MsgBoxEx_Error
End Sub

Private Sub btnPrevious_Click()
    On Error GoTo Flag_Err
    
    If Me.MultiPageMain.value > 0 Then
        Me.MultiPageMain.value = Me.MultiPageMain.value - 1
        Call SetWizardStatus(Me.MultiPageMain.value - 1)
    End If
    
    Exit Sub
Flag_Err:
    Call MsgBoxEx_Error
End Sub

Private Sub cboProvider_Change()
    Call GenerateConnectionString
End Sub

Private Sub cboSheet_Enter()
    Call SelectAllText(cboSheet)
End Sub

Private Sub chkSelectTablesAll_Change()

    SkipShowTablesStatisticsInfo = True
    Call SelectAllListBoxItems(Me.lstTables, chkSelectTablesAll.value)
    SkipShowTablesStatisticsInfo = False
    lstTables_Change

End Sub

Private Sub lstTables_Change()
    If Not SkipShowTablesStatisticsInfo Then
        Call ShowTablesInfo
    End If
End Sub

Private Sub optImportModeOnlyUpdateTemplateSheet_Click()
    If optImportModeOnlyUpdateTemplateSheet.value Then
        Me.chkClearExistedData.value = True
    End If
End Sub

Private Sub optImportModeOverwrite_Change()
    If optImportModeOverwrite.value Then
        Me.chkClearExistedData.value = True
    End If
End Sub

Private Sub txtPassword_Change()
    Call GenerateConnectionString
End Sub

Private Sub txtServer_Change()
    Call GenerateConnectionString
End Sub

Private Sub txtServer_Enter()
    Call SelectAllText(txtServer)
End Sub

Private Sub txtPassword_Enter()
    Call SelectAllText(txtPassword)
End Sub

Private Sub txtUser_Change()
    Call GenerateConnectionString
End Sub

Private Sub txtUser_Enter()
    Call SelectAllText(txtUser)
End Sub

Private Sub UserForm_Initialize()
    SkipShowTablesStatisticsInfo = False
    Call Init
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

