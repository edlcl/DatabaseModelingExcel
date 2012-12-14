Attribute VB_Name = "basString"
Option Explicit

''' -----------------------------------------------------------------------------
''' <summary>
'''     Trim a string include space, vblf, vbcr.
''' </summary>
''' <param name="text"></param>
''' <returns>trimed string</returns>
''' <remarks>
''' </remarks>
''' -----------------------------------------------------------------------------
Public Function TrimEx(ByVal text As String) As String
    Dim sRet As String
    sRet = Trim(text)
    Do While Len(sRet) > 0
        If Left(sRet, 1) = vbLf Or Left(sRet, 1) = vbCr Then
            sRet = Mid(sRet, 2)
        Else
            Exit Do
        End If
    Loop
    Do While Len(sRet) > 0
        If Right(sRet, 1) = vbLf Or Right(sRet, 1) = vbCr Then
            sRet = Mid(sRet, 1, Len(sRet) - 1)
        Else
            Exit Do
        End If
    Loop

    '-- Return
    TrimEx = sRet
End Function

''' -----------------------------------------------------------------------------
''' <summary>
'''     Format string.
''' </summary>
''' <param name="text"></param>
''' <param name="Args"></param>
''' <returns></returns>
''' <remarks>
'''     format like:
'''     "a{0:description}b{1}c{{0}}"
'''     {0:description} is a tag; {1} is a tag; {{0}} is {0}, a,b,c is a,b,c
''' </remarks>
''' -----------------------------------------------------------------------------
Public Function FormatString(ByVal text As String, ParamArray Args()) As String
    Dim newText             As String
    Dim index               As Long

    Dim textLength          As Long
    Dim argLength           As Integer
    Dim ch                  As String

    Dim tagBeginInNewText   As Long
    Dim tagText             As String
    Dim tagValue            As Integer
    Dim tagLength           As Long

    textLength = Len(text)
    argLength = UBound(Args)
    index = 1

    tagBeginInNewText = -1
    tagText = ""
    tagValue = -1

    Do While index <= textLength
        ch = Mid(text, index, 1)
        Select Case ch
            Case "{"
                If index < textLength Then
                    If Mid(text, index + 1, 1) = "{" Then
                        index = index + 1
                        GoTo FLAG_AddToText
                    End If
                End If
                tagBeginInNewText = Len(newText)
                tagText = ""
                tagValue = -1
            Case "}"
                If index < textLength Then
                    If Mid(text, index + 1, 1) = "}" Then
                        index = index + 1
                        GoTo FLAG_AddToText
                    End If
                End If
                If tagBeginInNewText >= 0 Then
                    If tagValue = -1 Then
                        tagLength = Len(tagText)
                        If tagLength > 0 And tagLength <= 4 Then
                            tagValue = CInt(tagText)
                            If tagValue > argLength Then
                                tagBeginInNewText = -1
                                tagText = ""
                                tagValue = -1
                            End If
                        Else
                            tagBeginInNewText = -1
                            tagText = ""
                            tagValue = -1
                        End If
                    End If
                    If tagValue >= 0 Then
                        newText = Mid(newText, 1, tagBeginInNewText) & Args(tagValue)
                        GoTo FLAG_NEXT
                    End If
                End If
            Case Else
                If tagBeginInNewText >= 0 Then
                    If IsNumeric(ch) And tagValue = -1 Then
                        tagText = tagText & ch
                    ElseIf ch = ":" Then
                        tagLength = Len(tagText)
                        If tagLength > 0 And tagLength <= 4 Then
                            tagValue = CInt(tagText)
                            If tagValue > argLength Then
                                tagBeginInNewText = -1
                                tagText = ""
                                tagValue = -1
                            End If
                        Else
                            tagBeginInNewText = -1
                            tagText = ""
                            tagValue = -1
                        End If
                    End If
                End If
        End Select
FLAG_AddToText:
        newText = newText & ch
FLAG_NEXT:
        index = index + 1
    Loop

    '-- Return
    FormatString = newText
End Function

''' -----------------------------------------------------------------------------
''' <summary>
'''     Get string before indicate string
''' </summary>
''' <param name="text"></param>
''' <param name="subString"></param>
''' <returns></returns>
''' <remarks>
'''     text = "table(col1, col2)"
'''     text2 = "("
'''     return "table"
''' </remarks>
''' -----------------------------------------------------------------------------
Public Function GetStringBefore(ByVal text As String, ByVal text2 As String) As String
    Dim pos     As Long
    pos = VBA.Strings.InStr(text, text2)
    If pos > 0 Then
        text = Mid(text, 1, pos - 1)
    End If
    
    GetStringBefore = text
End Function

''' -----------------------------------------------------------------------------
''' <summary>
'''     Get string after indicate string
''' </summary>
''' <param name="text"></param>
''' <param name="subString"></param>
''' <returns></returns>
''' <remarks>
'''     text = "table(col1, col2"
'''     text2 = "("
'''     return "col1, col2"
''' </remarks>
''' -----------------------------------------------------------------------------
Public Function GetStringAfter(ByVal text As String, ByVal text2 As String) As String
    Dim pos     As Long
    pos = VBA.Strings.InStr(text, text2)
    If pos > 0 Then
        text = Mid(text, pos + 1)
    End If
    
    GetStringAfter = text
End Function

''' -----------------------------------------------------------------------------
''' <summary>
'''     split and trim
''' </summary>
''' <param name="text"></param>
''' <param name="delimiter"></param>
''' <returns></returns>
''' <remarks>
''' </remarks>
''' -----------------------------------------------------------------------------
Public Function SplitAndTrim(ByVal text As String, ByVal delimiter As String) As String()
    Dim str()   As String
    Dim index   As Integer
    str = Split(text, delimiter)
    
    For index = LBound(str) To UBound(str)
        str(index) = Trim(str(index))
    Next
    SplitAndTrim = str
End Function


