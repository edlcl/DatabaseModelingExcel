Attribute VB_Name = "basPublic"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2013, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

Public Enum enmConnectionMode
    ConnectionModeDataSource
    ConnectionModeConnectionString
End Enum

Public Enum enmImportMode
    ImportModeOverwrite
    ImportModeAlwaysCreateSheet
    ImportModeAlwaysUpdate
End Enum

Public Function GetAppVersion() As String
    GetAppVersion = APP_VERSION
End Function

Public Function GetColumnIndex(name As String) As Integer
    Dim colIndex As Integer
    Dim name1 As String
    Dim name2 As String
    
    colIndex = 0
    If Len(name) = 1 Then
        colIndex = Asc(UCase(name)) - Asc("A") + 1
    ElseIf Len(name) > 1 Then
        colIndex = (Asc(UCase(Mid(name, 1, 1))) - Asc("A") + 1) * 26 _
            + Asc(UCase(Mid(name, 2, 1))) - Asc("A") + 1
    End If
        
    '-- return
    GetColumnIndex = colIndex
End Function

Public Function GetColumnName(index As Integer) As String
    Dim colName As String
    Dim name1 As Integer
    Dim name2 As Integer
    
    name1 = index / 26 - 0.5
    name2 = (index Mod 26)
    If name2 = 0 Then name2 = 26
    colName = IIf(name1 = 0, "", Chr(name1 + Asc("A") - 1)) _
        & Chr(name2 + Asc("A") - 1)
        
    '-- return
    GetColumnName = colName
End Function

'---------------------
'-  Get Sheet Name
'---------------------
Public Function GetSheetName(index As Integer) As String
    If (index > ThisWorkbook.Sheets.Count) Or (index < 1) Then
        GetSheetName = ""
    Else
        GetSheetName = ThisWorkbook.Sheets(index).name
    End If
End Function

Public Function CollectionIsContains(ByRef col As Collection, ByVal key As String) As Boolean
    On Error GoTo Flag_Err
    Dim a
    a = col.Item(key)
    CollectionIsContains = True
    Exit Function
Flag_Err:
    CollectionIsContains = False
End Function

Public Function CopyASheet(sourceSheet As Worksheet, Optional Before, Optional After) As Worksheet
    If Not VBA.IsMissing(Before) Then
        sourceSheet.Copy Before
    ElseIf Not VBA.IsMissing(After) Then
        sourceSheet.Copy , After
    Else
        sourceSheet.Copy
    End If
    Set CopyASheet = ThisWorkbook.ActiveSheet
End Function

Public Function SQL_ToSQL(value As Variant) As String
    Dim sSQL As String

    Select Case LCase(typeName(value))
    Case "integer", "long", "double", "single"
        sSQL = CStr(value)
    Case "date"
        Dim d As Date
        d = value
        sSQL = "'"
        sSQL = sSQL & Year(d) & "/" & Month(d) & "/" & Day(d)
        If Hour(d) <> 0 _
            Or Minute(d) <> 0 _
            Or Second(d) <> 0 Then
            sSQL = sSQL & " " & Hour(d) & ":" & Minute(d) & ":" & Second(d)
        End If
        sSQL = sSQL & "'"
    Case "boolean"
        sSQL = IIf(value, "1", "0")
    Case Else
        sSQL = "'" & Replace(CStr(value), "'", "''") & "'"
    End Select
    
    SQL_ToSQL = sSQL
End Function
