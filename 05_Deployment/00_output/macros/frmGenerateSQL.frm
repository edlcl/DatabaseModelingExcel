VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGenerateSQL 
   Caption         =   "Generate for <database>"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   OleObjectBlob   =   "frmGenerateSQL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGenerateSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SelectAllItem(booSel As Boolean)
    Dim index      As Integer
    
    '-- Select all items or deselect all items
    With lstTables
        For index = 0 To .ListCount - 1
            .selected(index) = booSel
        Next
    End With
End Sub

'---------------------------------------------
'-  Get all seleted table information
'---------------------------------------------
Private Function GetSelectedLogicalTables() As Collection
    Dim objLogicalTables    As Collection
    Dim index               As Integer
    Dim iSheet              As Integer
    
    Set objLogicalTables = New Collection
    
    With lstTables
        For index = 0 To .ListCount - 1
            If .selected(index) Then
                iSheet = .List(index, 1)
                Call objLogicalTables.Add(GetTableInfoFromWorksheet(ThisWorkbook.Sheets(iSheet)))
            End If
        Next
    End With
    
    '-- Return
    Set GetSelectedLogicalTables = objLogicalTables
End Function

Private Function CopyCreateTableSQL(ByVal withDescription As Boolean) As String
    Select Case gCurentDatabaseType
    Case DBName_MySQL
        Call basSQL_MySQL.Get_SQL_CreateTable(GetSelectedLogicalTables, withDescription)
    Case DBName_Oracle
        Call basSQL_Oracle.Get_SQL_CreateTable(GetSelectedLogicalTables, withDescription)
    Case Else
        Call basSQL_SQLServer.Get_SQL_CreateTable(GetSelectedLogicalTables, withDescription)
    End Select
End Function

Private Sub CopyDropTableSQL()
    Select Case gCurentDatabaseType
    Case DBName_MySQL
        Call basSQL_MySQL.Get_SQL_DropTable(GetSelectedLogicalTables)
    Case DBName_Oracle
        Call basSQL_Oracle.Get_SQL_DropTable(GetSelectedLogicalTables)
    Case Else
        Call basSQL_SQLServer.Get_SQL_DropTable(GetSelectedLogicalTables)
    End Select
End Sub

Private Sub CopyDropAndCreateTableSQL(ByVal withDescription As Boolean)
    Select Case gCurentDatabaseType
    Case DBName_MySQL
        Call basSQL_MySQL.Get_SQL_DropAndCreateTable(GetSelectedLogicalTables, withDescription)
    Case DBName_Oracle
        Call basSQL_Oracle.Get_SQL_DropAndCreateTable(GetSelectedLogicalTables, withDescription)
    Case Else
        Call basSQL_SQLServer.Get_SQL_DropAndCreateTable(GetSelectedLogicalTables, withDescription)
    End Select
End Sub

Private Function CopyCreateTableIfNotExistsSQL() As String
    Select Case gCurentDatabaseType
    Case DBName_MySQL
        Call basSQL_MySQL.Get_SQL_CreateTableIfNotExists(GetSelectedLogicalTables)
    Case DBName_Oracle
        Call basSQL_Oracle.Get_SQL_CreateTableIfNotExists(GetSelectedLogicalTables)
    Case Else
        Call basSQL_SQLServer.Get_SQL_CreateTableIfNotExists(GetSelectedLogicalTables)
    End Select
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Me.MultiPage1.SelectedItem.index = 0 Then
        If Me.optCreateTableSQL.value Then
            Call CopyCreateTableSQL(Me.chkWithFieldDescription.value)
            
        ElseIf optDropTableSQL.value = True Then
            Call CopyDropTableSQL
            
        ElseIf optDropAndCreateSQL.value = True Then
            Call CopyDropAndCreateTableSQL(Me.chkWithFieldDescription.value)
        
        ElseIf Me.optCreateTableIfNotExistsSQL.value Then
            Call CopyCreateTableIfNotExistsSQL
            
        End If
    End If
    
    '-- Return
    Unload Me
End Sub

Private Sub cmdSelectAll_Click()
    Call SelectAllItem(True)
End Sub

Private Sub cmdSelectNone_Click()
    Call SelectAllItem(False)
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Generate for " & gCurentDatabaseType
    Me.MultiPage1.Pages(0).Caption = gCurentDatabaseType
    
    Dim iSheet      As Integer
    Dim oSheet      As Worksheet
    Dim index       As Integer
    With lstTables
        '-- Get create tables's SQL
        index = 0
        For iSheet = Sheet_First_Table To ThisWorkbook.Sheets.Count
            Set oSheet = ThisWorkbook.Sheets(iSheet)
            If VBA.StrComp( _
                    LCase(TrimEx( _
                        oSheet.Cells.Item(Table_Sheet_Row_TableStatus, Table_Sheet_Col_TableStatus).text)) _
                    , Table_Sheet_TableStatus_Ignore) _
                <> 0 Then
                .AddItem (ThisWorkbook.Sheets(iSheet).Name)
                .List(index, 1) = iSheet
                index = index + 1
            End If
        Next
    End With
    
    '-- Defaut Select ALL
    Call SelectAllItem(True)
End Sub
