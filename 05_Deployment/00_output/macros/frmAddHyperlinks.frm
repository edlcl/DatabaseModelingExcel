VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddHyperlinks 
   Caption         =   "Add Hyperlinks to Tables Sheets in the Index Sheet"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   OleObjectBlob   =   "frmAddHyperlinks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddHyperlinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2013, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

Private Sub DoAddHyperLinks()
    On Error GoTo Flag_Err

    Dim shtIndex    As Worksheet
    Dim startRow    As Integer
    Dim startCol    As Integer
    Dim index       As Integer
    Dim objCell     As Range
        
    startCol = GetColumnIndex(Trim(txtStartColumn.text))
    startRow = VBA.Int(Trim(txtStartRow.text))
    
    Set shtIndex = ThisWorkbook.Sheets(Sheet_Index)
    For index = 2 To ThisWorkbook.Sheets.Count
        Set objCell = shtIndex.Cells(startRow + index - 2, startCol)
        objCell.value = ">>"
        objCell.Hyperlinks.Delete
        Call shtIndex.Hyperlinks.Add(objCell, _
                    "", _
                    "'" & GetSheetName(index) & "'!A1")
    Next
    
    Exit Sub
Flag_Err:
    Call MsgBoxEx_Error
End Sub

Private Sub btnCancel_Click()
    Unload frmAddHyperlinks
End Sub

Private Sub btnOK_Click()
    DoAddHyperLinks
End Sub

Private Sub UserForm_Initialize()
    Me.txtStartColumn.text = "B"
    Me.txtStartRow.text = "4"
End Sub
