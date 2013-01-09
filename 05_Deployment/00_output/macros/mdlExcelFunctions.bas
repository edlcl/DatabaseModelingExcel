Attribute VB_Name = "mdlExcelFunctions"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
'{93E95525-C71D-4753-B7F9-179D9023639B}
Option Explicit

Public Function VBComponent_ExportAll()
    Dim path As String
    path = InputBox("Please select exported folder:", Application.Caption)
    If Len(path) = 0 Then Exit Function
    
    Call VBComponent_ExportAll_Command(path)
End Function

Public Function VBComponent_ExportAll_Command(ByVal a_sPath As String)
    Dim vbCom As VBComponent
    Dim path As String
    Dim suffix As String
    path = a_sPath

    If Len(path) = 0 Then Exit Function
    
    For Each vbCom In ThisWorkbook.VBProject.VBComponents
        Select Case vbCom.Type
            Case vbext_ct_Document
               suffix = ".dls"
            Case vbext_ct_ClassModule
               suffix = ".cls"
            Case vbext_ct_MSForm
               suffix = ".frm"
            Case vbext_ct_StdModule
               suffix = ".bas"
            Case Else
               suffix = ""
        End Select
        If Len(suffix) > 0 _
            And vbCom.CodeModule.CountOfLines > 1 Then
            
            vbCom.Export path & "\" & vbCom.name & suffix
        End If
    Next
End Function

Public Function VBComponent_ImportAll()
    Dim path    As String
    path = InputBox("Please select import file(s) folder:", Application.Caption)
    If Len(path) = 0 Then Exit Function

    Call VBComponent_ImportAll_Command(path)
End Function

Public Function VBComponent_ImportAll_Command(ByVal a_sPath As String)
    Dim vbCom   As VBComponent
    Dim path    As String
    Dim comType As vbext_ComponentType
    path = a_sPath
    If Len(path) = 0 Then Exit Function
    
    Dim fileName As String
    Dim documentName As String
    fileName = Dir(path & "\*.*")

    Do While Len(fileName) > 0
        If LCase(Right(fileName, 4)) = ".dls" Then
            comType = vbext_ct_Document
        ElseIf LCase(Right(fileName, 4)) = ".cls" Then
            comType = vbext_ct_ClassModule
        ElseIf LCase(Right(fileName, 4)) = ".frm" Then
            comType = vbext_ct_MSForm
        ElseIf LCase(Right(fileName, 4)) = ".bas" Then
            comType = vbext_ct_StdModule
        Else
            GoTo FLAG_NEXT
        End If
        
        '-- Skip the file
        If LCase(fileName) = LCase("mdlExcelFunctions.bas") Then GoTo FLAG_NEXT
        
        If comType = vbext_ct_Document Then
            documentName = Left(fileName, Len(fileName) - 4)
            Set vbCom = ThisWorkbook.VBProject.VBComponents(documentName)
            Call vbCom.CodeModule.DeleteLines(1, vbCom.CodeModule.CountOfLines)
            Call vbCom.CodeModule.AddFromFile(path & "\" & fileName)
            Call vbCom.CodeModule.DeleteLines(1, 4)
        Else
            Set vbCom = ThisWorkbook.VBProject.VBComponents.Import(path & "\" & fileName)
        End If
        
FLAG_NEXT:
        fileName = Dir
    Loop
End Function

Public Function VBComponent_CloseAllCodeWindows()
    Dim i As Long
    i = 1
    Do While i <= ThisWorkbook.VBProject.VBE.Windows.Count
        If ThisWorkbook.VBProject.VBE.Windows(i).Type = vbext_wt_CodeWindow Then
            ThisWorkbook.VBProject.VBE.Windows(i).Close
            i = i - 1
        End If
        i = i + 1
    Loop
End Function

Public Function VBComponent_RemoveAll()
    If MsgBox("Are you want to clear all macros in the file[" & ThisWorkbook.name & "]?" _
        , vbYesNo + vbQuestion + vbDefaultButton2, Application.Caption) = vbNo _
    Then Exit Function
    
    Call VBComponent_RemoveAll_Command
End Function

Public Function VBComponent_RemoveAll_Command()
    Dim vbCom As VBComponent

    Dim i As Long
    i = 1
    Do While i <= ThisWorkbook.VBProject.VBComponents.Count
        Set vbCom = ThisWorkbook.VBProject.VBComponents(i)

        If vbCom.Type = vbext_ct_Document Then
            If vbCom.CodeModule.CountOfLines > 0 Then
                vbCom.CodeModule.DeleteLines 1, vbCom.CodeModule.CountOfLines
            End If
        ElseIf IsThisVBComponent(vbCom) Then
        Else
            ThisWorkbook.VBProject.VBComponents.Remove vbCom
            i = i - 1
        End If
        i = i + 1
    Loop
End Function

Private Function IsThisVBComponent(vbCom As VBComponent) As Boolean
    Dim boo As Boolean
    Dim index As Integer
    
    boo = False
    For index = 1 To vbCom.CodeModule.CountOfLines
        If vbCom.CodeModule.Lines(index, 1) = "'{93E95525-C71D-4753-B7F9-179D9023639B}" Then
            boo = True
            Exit For
        End If
        
        '-- only search top 10 lines
        If index >= 10 Then
            Exit For
        End If
    Next
    
    IsThisVBComponent = boo
End Function

Public Function Sheet_GetColumnHeaderText(index As Integer) As String
    Dim colName As String
    Dim name1 As Integer
    Dim name2 As Integer
    
    name1 = index / 26 - 0.5
    name2 = (index Mod 26)
    If name2 = 0 Then name2 = 26
    colName = IIf(name1 = 0, "", Chr(name1 + Asc("A") - 1)) _
        & Chr(name2 + Asc("A") - 1)
        
    '-- return
    Sheet_GetColumnHeaderText = colName
End Function

Private Function renderFormula(ByVal formula As String) As String
    '-- render formula to script
    formula = Replace(formula, vbCrLf, vbLf)
    formula = Replace(formula, vbCr, vbLf)
    formula = Replace(formula, """", """""")
    formula = Replace(formula, vbLf, """ _" & vbLf & vbTab & vbTab & "& vbLf & """)

    renderFormula = formula
End Function
