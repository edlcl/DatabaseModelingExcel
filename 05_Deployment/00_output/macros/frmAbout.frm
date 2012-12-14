VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "About Database Modeling Excel"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnDenote_Click()
 On Error GoTo Flag_Err
    Call VBA.Shell("cmd /C ""start http://sourceforge.net/donate/index.php?group_id=171489""")
    
    Exit Sub
Flag_Err:
    Call MsgBoxEx_Error
End Sub

Private Sub txtEmail_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo Flag_Err
    Call VBA.Shell("cmd /C ""start mailto:" & txtEmail & """")
    
    Exit Sub
Flag_Err:
    Call MsgBoxEx_Error
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "About " & APP_NAME
    Me.labName = APP_NAME & " < " & APP_VERSION & " >"
    
    Dim desc As String
    desc = "Support open source!" _
        & vbCrLf & "To contribute more software, make the world better." _
        & vbCrLf & "" _
        & vbCrLf & "What's a reasonable donation?" _
        & vbCrLf & "Even $1 is enough to show your appreciation - give what you can." _
        & vbCrLf & "" _
        & vbCrLf & "If you are a business, using " & APP_NAME & " for professional endeavors, you should consider donating at least $20. If you are using several copies of the program in your company you should perhaps pay an additional $5 for each copy. In the end, pay what you feel is fair."
        
    Me.txtDenoteNote.text = desc
    Me.txtDenoteNote.SelStart = 0
End Sub
