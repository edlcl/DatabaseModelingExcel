Attribute VB_Name = "basTableSheet"
Option Explicit

'---------------------------------------------
'-  Get All Table information
'---------------------------------------------
Public Function GetAllLogicalTables() As Collection
    Dim objLogicalTables    As Collection
    Dim iSheet              As Integer
    Dim oSheet              As Worksheet
    
    Set objLogicalTables = New Collection
    
    For iSheet = Sheet_First_Table To ThisWorkbook.Sheets.Count
        Set oSheet = ThisWorkbook.Sheets(iSheet)
        If VBA.StrComp( _
                LCase(TrimEx( _
                    oSheet.Cells.Item(Table_Sheet_Row_TableStatus, Table_Sheet_Col_TableStatus).text)) _
                , Table_Sheet_TableStatus_Ignore) _
            <> 0 Then
            objLogicalTables.Add GetTableInfoFromWorksheet(ThisWorkbook.Sheets(iSheet))
        End If
    Next
    
    '-- Return
    Set GetAllLogicalTables = objLogicalTables
End Function

'---------------------------------------------
'-  Get Table information
'---------------------------------------------
Public Function GetTableInfoFromWorksheet(shtCurrent As Worksheet) As clsLogicalTable
    Dim objTable As clsLogicalTable
    
    Set objTable = New clsLogicalTable
    objTable.tableName = Trim(shtCurrent.Cells.Item(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).text)
    objTable.Description = Trim(shtCurrent.Cells.Item(Table_Sheet_Row_TableDescription, Table_Sheet_Col_TableDescription).text)
   
    Set objTable.PrimaryKey = GetTablePrimaryKey(shtCurrent)
    Set objTable.ForeignKeys = GetTableForeignKeys(shtCurrent)
    Set objTable.Indexes = GetTableIndexes(shtCurrent)
    Set objTable.Columns = GetTableColumns(shtCurrent)
    
    '-- Return
    Set GetTableInfoFromWorksheet = objTable
End Function

'---------------------------------------------
'-  Get Columns information
'---------------------------------------------
Public Function GetTableColumns(shtCurrent As Worksheet) As Collection
    Dim objColumns      As Collection
    Dim objColumn       As clsLogicalColumn
    Dim strCell         As String
    Dim index           As Integer
    
    Set objColumns = New Collection
    index = 1
    Do While (True)
        '-- Get Column name
        strCell = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_First_Column + index - 1, Table_Sheet_Col_ColumnName).text)
        
        '-- if Column name is '', finished Columns search
        If Len(strCell) = 0 Then Exit Do
        
        '-- add a Column information
        Set objColumn = New clsLogicalColumn
        objColumns.Add objColumn
        With objColumn
            '-- Get Column information
            .columnLabel = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_First_Column + index - 1, Table_Sheet_Col_ColumnLabel).text)
            .columnName = strCell
            .dataType = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_First_Column + index - 1, Table_Sheet_Col_ColumnDataType).text)
            .Nullable = IIf(UCase(Trim(shtCurrent.Cells.Item(Table_Sheet_Row_First_Column + index - 1, Table_Sheet_Col_ColumnNullable).text)) = "YES", True, False)
            .Default = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_First_Column + index - 1, Table_Sheet_Col_ColumnDefault).text)
            .Note = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_First_Column + index - 1, Table_Sheet_Col_ColumnNote).text)
        End With
        index = index + 1
    Loop
    
    '-- Return
    Set GetTableColumns = objColumns
End Function

'---------------------------------------------
'-  Get PrimaryKeys information
'---------------------------------------------
Public Function GetTablePrimaryKey(shtCurrent As Worksheet) As clsLogicalPrimaryKey
    Dim objPK           As clsLogicalPrimaryKey
    Dim strCell         As String
    
    Set objPK = New clsLogicalPrimaryKey
    
    strCell = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_PrimaryKey, Table_Sheet_Col_Clustered).text)
    With objPK
        '-- Get PK Columns' information
        .PKcolumns = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_PrimaryKey, Table_Sheet_Col_PrimaryKey).text)
        
        '-- Get clustered information
        .IsClustered = Not (UCase(Trim(strCell)) = "N")
    End With
    
    '-- Return
    Set GetTablePrimaryKey = objPK
End Function

'---------------------------------------------
'-  Get ForeignKeys information
'---------------------------------------------
Public Function GetTableForeignKeys(shtCurrent As Worksheet) As Collection
    Dim objFKs          As Collection
    Dim objFK           As clsLogicalForeignKey
    Dim strCell         As String
    Dim index           As Integer
    Dim strFKArray()    As String
    Dim strFKItem       As String
    Dim intFKItemLen    As Integer
    Dim intPos          As Integer
    
    index = 0
    Set objFKs = New Collection
    
    '-- Get Column name
    strCell = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_ForeignKey, Table_Sheet_Col_ForeignKey).text)
    '-- Split PK infomation into array, one item is a infomation of foreign key
    strFKArray = Split(strCell, ";")
    
    For index = 0 To UBound(strFKArray)
        Set objFK = New clsLogicalForeignKey
        objFKs.Add objFK
        
        strFKItem = TrimEx(strFKArray(index))
        '-- Replace ", "  to ",", to aviod get wrong table name.
        Do While True
            intFKItemLen = Len(strFKItem)
            strFKItem = Replace(strFKItem, ", ", ",")
            If Len(strFKItem) = intFKItemLen Then
                Exit Do
            End If
        Loop
        
        intPos = InStr(1, strFKItem, " ")
        With objFKs(index + 1)
            '-- Get Foreign key's Columnname
            .FKcolumns = Left(strFKItem, intPos - 1)
            '-- get rid of vbcr and vblf
            Do While Left(.FKcolumns, 1) = vbCr Or Left(.FKcolumns, 1) = vbLf
                .FKcolumns = Mid(.FKcolumns, 2)
            Loop
            '-- Get Foreign key's foreign table infomation
            Call SetForeignKeyRefTableAndName(objFKs(index + 1), Mid(strFKItem, intPos + 1))
            
            .FKName = Replace(Replace(.FKcolumns, " ", ""), ",", "$")
        End With
    Next
    
    '-- Return
    Set GetTableForeignKeys = objFKs
End Function

'---------------------------------------------
'-  GetIndexKeys information
'---------------------------------------------
Public Function GetTableIndexes(shtCurrent As Worksheet) As Collection
    Dim objIKs              As Collection
    Dim objIK               As clsLogicalIndex
    Dim strCell             As String
    Dim index               As Integer
    Dim strIKArray()        As String
    Dim strIKUnique()       As String
    Dim strIKClustered()    As String
    Dim strIKItem           As String
    Dim intPos              As Integer
    
    index = 0
    Set objIKs = New Collection
  
    '-- Get Index infomation
    strCell = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_Index, Table_Sheet_Col_Index).text)
    '-- Split index infomation into array, one item is a infomation of index
    strIKArray = Split(strCell, ";")
    
    '-- Get index Unique information
    strCell = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_Index, Table_Sheet_Col_Unique).text)
    '-- Split index's unique infomation into array, one item is a infomation of index's unique
    strIKUnique = Split(strCell, ";")
    
    '-- Get index Clustered information
    strCell = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_Index, Table_Sheet_Col_Clustered).text)
    '-- Split index's Clustered infomation into array, one item is a infomation of index's Clustered
    strIKClustered = Split(strCell, ";")
    
    For index = 0 To UBound(strIKArray)
        Set objIK = New clsLogicalIndex
        objIKs.Add objIK
        
        '-- Get one IK describation
        strIKItem = TrimEx(strIKArray(index))
        With objIK
            '-- Set default information
            .IsClustered = False
            .IsUnique = True
            
            '-- Get index's name and index's columns
            intPos = InStr(1, strIKItem, ",")
            If intPos = 0 Then
                '-- Get index's name
                .IKName = strIKItem
            Else
                '-- Get index's name
                .IKName = Replace(Replace(strIKItem, " ", ""), ",", "$")
            End If
            '-- Get index's Columns infomation
            .IKColumns = "(" & strIKItem & ")"
            
            '-- Is Uniqued?
            If UBound(strIKUnique) >= index Then
                If UCase(TrimEx(strIKUnique(index))) = "N" Then
                    '-- Not Unique flag
                    .IsUnique = False
                End If
            End If
            
            '-- Is Clustered?
            If UBound(strIKClustered) >= index Then
                If UCase(TrimEx(strIKClustered(index))) = "Y" Then
                    '-- Not Unique flag
                    .IsClustered = True
                End If
            End If
        End With
    Next
    
    '-- Return
    Set GetTableIndexes = objIKs
End Function

'---------------------------------------------
'-  Write Table information to worksheet
'---------------------------------------------
Public Sub SetTableInfoToWorksheet(ByVal sh As Worksheet, _
                        ByVal table As clsLogicalTable, _
                        ByVal clearSheet As Boolean)
    
    Dim indexText       As String
    Dim indexClustered  As String
    Dim indexUnique     As String
    
    '-- Set Table Name
    If clearSheet _
        Or sh.Cells(Table_Sheet_Row_TableDescription, _
                Table_Sheet_Col_TableDescription).text = "" Then
                
            sh.Cells.Item(Table_Sheet_Row_TableDescription, _
                Table_Sheet_Col_TableDescription).value = table.tableName
    End If
    sh.Cells(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).value = table.tableName
    
    '-- Set PK
    Call table.GetPrimaryKeyInfoText(indexText, indexClustered)
    sh.Cells(Table_Sheet_Row_PrimaryKey, Table_Sheet_Col_PrimaryKey).value = indexText
    sh.Cells(Table_Sheet_Row_PrimaryKey, Table_Sheet_Col_Clustered).value = indexClustered

    '-- Set Index
    Call table.GetIndexexInfoText(indexText, indexClustered, indexUnique)
    sh.Cells(Table_Sheet_Row_Index, Table_Sheet_Col_Index).value = indexText
    sh.Cells(Table_Sheet_Row_Index, Table_Sheet_Col_Clustered).value = indexClustered
    sh.Cells(Table_Sheet_Row_Index, Table_Sheet_Col_Unique).value = indexUnique

    '-- Set Index Row Height
    sh.Rows(Table_Sheet_Row_Index & ":" & Table_Sheet_Row_Index).Select
    If table.Indexes.Count > 0 Then
        Selection.RowHeight = table.Indexes.Count * Table_Sheet_Row_Height
    Else
        Selection.RowHeight = 1 * Table_Sheet_Row_Height
    End If
    Application.CutCopyMode = False
    
    '-- Set FK
    sh.Cells(Table_Sheet_Row_ForeignKey, Table_Sheet_Col_ForeignKey).value = table.GetForeignKeysText
    
    '-- Set FK Row Height
    sh.Rows(Table_Sheet_Row_ForeignKey & ":" & Table_Sheet_Row_ForeignKey).Select
    If table.ForeignKeys.Count > 0 Then
        Selection.RowHeight = table.ForeignKeys.Count * Table_Sheet_Row_Height
    Else
        Selection.RowHeight = 1 * Table_Sheet_Row_Height
    End If
    Application.CutCopyMode = False
    
    '-- Set Column
    Dim row                 As Integer
    Dim tableColumn         As clsLogicalColumn
    
    row = Table_Sheet_Row_First_Column
    For Each tableColumn In table.Columns
        '-- Render column row
        If row > Table_Sheet_Row_First_Column And sh.Cells(row, Table_Sheet_Col_ColumnID).text = "" Then
            sh.Rows(row - 1 & ":" & row - 1).Select
            Selection.Copy
            sh.Rows(row & ":" & row).Select
            Selection.Insert Shift:=xlDown
            Application.CutCopyMode = False
            sh.Cells(row, Table_Sheet_Col_ColumnLabel).value = ""
            sh.Cells(row, Table_Sheet_Col_ColumnNote).value = ""
        End If
        sh.Range(GetColumnName(Table_Sheet_Col_ColumnLabel) & row, GetColumnName(Table_Sheet_Col_ColumnDataType - 1) & row).Select
        Selection.Interior.ColorIndex = xlNone
        Selection.Font.Bold = False
        '-- render PK
        If table.IsPKColumn(tableColumn.columnName) Then
            sh.Range(GetColumnName(Table_Sheet_Col_ColumnLabel) & row, GetColumnName(Table_Sheet_Col_ColumnDataType - 1) & row).Select
            Selection.Interior.ColorIndex = 15
            Application.CutCopyMode = False
        End If
        '-- render FK
        If table.IsFKColumn(tableColumn.columnName) Then
            sh.Range(GetColumnName(Table_Sheet_Col_ColumnLabel) & row, GetColumnName(Table_Sheet_Col_ColumnDataType - 1) & row).Select
            Selection.Font.Bold = True
            Application.CutCopyMode = False
        End If
        
        '-- set Column
        sh.Cells(row, Table_Sheet_Col_ColumnName).Select
        If clearSheet _
            Or sh.Cells(row, Table_Sheet_Col_ColumnLabel).text = "" Then
            sh.Cells(row, Table_Sheet_Col_ColumnLabel).value = tableColumn.columnName
        End If
        sh.Cells(row, Table_Sheet_Col_ColumnName).value = tableColumn.columnName
        sh.Cells(row, Table_Sheet_Col_ColumnDataType).value = tableColumn.dataType
        sh.Cells(row, Table_Sheet_Col_ColumnNullable).value = IIf(tableColumn.Nullable, Table_Sheet_Nullable, Table_Sheet_NonNullable)
        sh.Cells(row, Table_Sheet_Col_ColumnDefault).value = IIf(Len(tableColumn.Default) > 0, "'" & tableColumn.Default, "")
        If clearSheet Then
            sh.Cells(row, Table_Sheet_Col_ColumnNote).value = ""
        End If

        '-- Move next record
        row = row + 1
    Next
    '-- set left row
    For row = row To row + 2
        If IsColumnRow(sh, row) And clearSheet Then
            SetColumnEmpty sh, row
        Else
            Exit For
        End If
    Next
    If clearSheet Then
        row = row - 1
        For row = row To 32667
            If IsColumnRow(sh, row) Then
                sh.Rows(row & ":" & row).Select
                Selection.Delete Shift:=xlUp
                row = row - 1
            Else
                Exit For
            End If
        Next
    End If
    sh.Cells(1, 1).Select
End Sub

Private Function IsColumnRow(sh As Worksheet, row As Integer) As Boolean
    Dim ma
    Set ma = sh.Range(GetColumnName(Table_Sheet_Col_ColumnLabel) & row).MergeArea
    IsColumnRow = (ma.Address = "$" & GetColumnName(Table_Sheet_Col_ColumnLabel) & "$" & row _
                    & ":$" & GetColumnName(Table_Sheet_Col_ColumnName - 1) & "$" & row)
End Function

Private Function SetColumnEmpty(ByVal sh As Worksheet, _
                        ByVal row As Integer) As Boolean
    sh.Range(GetColumnName(Table_Sheet_Col_ColumnLabel) & row, GetColumnName(Table_Sheet_Col_ColumnDataType - 1) & row).Select
    Selection.Interior.ColorIndex = xlNone
    Selection.Font.Bold = False
    '-- set Column
    sh.Cells(row, Table_Sheet_Col_ColumnLabel).value = ""
    sh.Cells(row, Table_Sheet_Col_ColumnNote).value = ""

    sh.Cells(row, Table_Sheet_Col_ColumnName).value = ""
    sh.Cells(row, Table_Sheet_Col_ColumnDataType).value = ""
    sh.Cells(row, Table_Sheet_Col_ColumnNullable).value = ""
    sh.Cells(row, Table_Sheet_Col_ColumnDefault).value = ""

End Function

Public Function SetForeignKeyRefTableAndName(ByVal foreignKey As clsLogicalForeignKey, _
                                            ByVal strRefTableAndColumns As String)
    Dim refTableName    As String
    Dim refColumns      As String
    Dim fkOption        As String
    
    strRefTableAndColumns = Trim(strRefTableAndColumns)
    
    refTableName = GetStringBefore(strRefTableAndColumns, "(")
    refColumns = GetStringAfter(strRefTableAndColumns, "(")
    fkOption = GetStringAfter(refColumns, ")")
    refColumns = GetStringBefore(refColumns, ")")
    
    foreignKey.refTableName = refTableName
    foreignKey.RefTableColumns = refColumns
    foreignKey.fkOption = fkOption
End Function

Public Function GetSheetFromTableName(ByVal tableName As String) As Worksheet
    Dim sh      As Worksheet
    Dim index   As Integer
    
    tableName = LCase(Trim(tableName))
    For index = Sheet_First_Table To ThisWorkbook.Sheets.Count
        If LCase(ThisWorkbook.Sheets(index).Cells(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).text) = tableName Then
            Set sh = ThisWorkbook.Sheets(index)
            GoTo Exit_Flag
        End If
    Next
    
Exit_Flag:
    Set GetSheetFromTableName = sh
End Function
