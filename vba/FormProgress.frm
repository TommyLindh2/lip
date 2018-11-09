VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormProgress 
   Caption         =   "Loading"
   ClientHeight    =   816
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5745
   OleObjectBlob   =   "FormProgress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
On Error GoTo ErrorHandler
    
    Call FormHelper.SetFormDefaultColors(Me)
    Me.Progress.BackColor = &H8000000D
Exit Sub

ErrorHandler:
    If Err = 424 Then
        Call LC_FormHelper.SetFormDefaultColors(Me)
        Resume Next
    End If
    Call LC_UI.ShowError("FormProgress.UserForm_Initialize")
End Sub
