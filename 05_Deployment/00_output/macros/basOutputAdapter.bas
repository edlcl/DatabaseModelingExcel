Attribute VB_Name = "basOutputAdapter"
'--------------------------------------------------------------------------------
' This file includes following class:
'  mdlOutputAdapter
'  enmOutputPlace
'--------------------------------------------------------------------------------
' Created By: StevenYang
' Date:       31/03/2006
' Modified History:
' Date     By          Bug ID          Description
' ------   ----------  ----------      ------------
'
'--------------------------------------------------------------------------------
Option Explicit

Enum enmOutputPlace
    oppToClipboard = 0
    oppToExcelSheet
End Enum

Public Output_Place                 As enmOutputPlace
Private outputContentArray()        As String
Const G_Output_SHEET_NAME = "Output"

Public Function Output_Initialize(Optional outputPlace As enmOutputPlace = oppToClipboard)
    Output_Place = outputPlace
    ReDim outputContentArray(0) As String
    ReDim objOutputLine(1) As Long
End Function

Public Function Output_Write(ByVal text As String, Optional outputID As Integer = 0)
    Dim outputContent As String
    
    outputContent = GetOuputContent(outputID)
    
    outputContent = outputContent & text
    
    SetOutputContentToCollection outputID, outputContent
End Function

Public Function Output_WriteLine(ByVal text As String, Optional outputID As Integer = 0)
    Output_Write text & vbCrLf, outputID
End Function

Private Function GetOuputContent(outputID As Integer) As String
    If UBound(outputContentArray) < outputID Then
        ReDim Preserve outputContentArray(outputID) As String
    End If

    GetOuputContent = outputContentArray(outputID)
End Function

Private Function SetOutputContentToCollection(outputID As Integer, outputContent As String)
    outputContentArray(outputID) = outputContent
End Function

Private Function GetAllOuputContentString() As String
    Dim outputContent As String
    Dim index As Integer
    
    For index = 0 To UBound(outputContentArray)
        If index > 0 Then outputContent = outputContent & vbCrLf
        outputContent = outputContent & outputContentArray(index)
    Next
    
    GetAllOuputContentString = outputContent
End Function

Public Function Output_Copy()

    CopyToClipboard GetAllOuputContentString
    
    ReDim outputContentArray(0) As String

    If Output_Place = oppToExcelSheet Then
        Dim sheet As Worksheet
        Set sheet = Sheets(G_Output_SHEET_NAME)
        sheet.Cells.ClearContents
        sheet.Range("A1").Select
        sheet.Paste
    End If
End Function

Public Function GetStringLen(ByRef text As String) As Long
    Dim i As Long
    Dim length As Long

    For i = 1 To Len(text)
        If Asc(Mid(text, i, 1)) < 255 And Asc(Mid(text, i, 1)) >= 0 Then
            length = length + 1
        Else
            length = length + LenB(Mid(text, i, 1))
        End If
    Next
    
    GetStringLen = length
End Function

