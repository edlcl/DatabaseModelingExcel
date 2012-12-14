Attribute VB_Name = "basClipboard"
Option Explicit

Private Declare Function CloseClipboard _
                Lib "user32" _
                () As Long
Private Declare Function OpenClipboard _
                Lib "user32" _
                (ByVal hWnd As Long) _
                As Long
Private Declare Function SetClipboardData _
                Lib "user32" _
                (ByVal wFormat As Long, _
                ByVal Hmem As Long) _
                As Long
Private Declare Function EmptyClipboard _
                Lib "user32" _
                () As Long
Private Declare Function RegisterClipboardFormat _
                Lib "user32" Alias "RegisterClipboardFormatA" _
                (ByVal lpString As String) _
                As Long

Private Declare Function GlobalAlloc _
                Lib "kernel32" _
                (ByVal wFlags As Long, _
                ByVal dwBytes As Long) _
                As Long
Private Declare Function GlobalFree _
                Lib "kernel32" _
                (ByVal Hmem As Long) _
                As Long

Private Declare Function GlobalLock _
                Lib "kernel32" _
                (ByVal Hmem As Long) _
                As Long
Private Declare Function GlobalUnlock _
                Lib "kernel32" ( _
                ByVal Hmem As Long) _
                As Long
Private Declare Sub CopyMemory _
                Lib "kernel32" Alias "RtlMoveMemory" _
                (pDest As Any, _
                pSource As Any, _
                ByVal cbLength As Long)
Private Declare Function GetClipboardData _
                Lib "user32" _
                (ByVal wFormat As Long) _
                As Long
Private Declare Function lstrlen _
                Lib "kernel32" Alias "lstrlenA" _
                (ByVal lpData As Long) _
                As Long
'/*
' * Predefined Clipboard Formats
' */
Public Const CF_TEXT            As Long = 1
Public Const CF_BITMAP          As Long = 2
Public Const CF_METAFILEPICT    As Long = 3
Public Const CF_SYLK            As Long = 4
Public Const CF_DIF             As Long = 5
Public Const CF_TIFF            As Long = 6
Public Const CF_OEMTEXT         As Long = 7
Public Const CF_DIB             As Long = 8
Public Const CF_PALETTE         As Long = 9
Public Const CF_PENDATA         As Long = 10
Public Const CF_RIFF            As Long = 11
Public Const CF_WAVE            As Long = 12
Public Const CF_UNICODETEXT     As Long = 13
Public Const CF_ENHMETAFILE     As Long = 14
'#if(WINVER >= 0x0400)
Public Const CF_HDROP           As Long = 15
Public Const CF_LOCALE          As Long = 16
'#endif /* WINVER >= 0x0400 */
'#if(WINVER >= 0x0500)
Public Const CF_DIBV5           As Long = 17
'#endif /* WINVER >= 0x0500 */

Public Function CopyToClipboard(ByRef text As String)
    Dim myDataObject As DataObject
    '-- Copy result to clipboard
    
    Set myDataObject = New DataObject
    myDataObject.Clear
    myDataObject.SetText text
    myDataObject.PutInClipboard
    Set myDataObject = Nothing

'        If CBool(OpenClipboard(0)) Then
'            EmptyClipboard
'            Dim hMemHandle As Long, lpData As Long, length As Long
'            length = GetStringLen(G_Output_Content)
'            hMemHandle = GlobalAlloc(0, length)
'            If CBool(hMemHandle) Then
'                lpData = GlobalLock(hMemHandle)
'                If lpData <> 0 Then
'                    CopyMemory ByVal lpData, ByVal G_Output_Content, length
'                    GlobalUnlock hMemHandle
'                    EmptyClipboard
'                    SetClipboardData 1, hMemHandle
'                End If
'                GlobalFree hMemHandle
'            End If
'            Call CloseClipboard
'        End If
End Function

