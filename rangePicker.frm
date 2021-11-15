VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} rangePicker 
   Caption         =   "rangePicker"
   ClientHeight    =   1755
   ClientLeft      =   -465
   ClientTop       =   -2100
   ClientWidth     =   3315
   OleObjectBlob   =   "rangePicker.frx":0000
End
Attribute VB_Name = "rangePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub CommandButton1_Click()
selectedRange = Me.RefEdit1.Value
Unload Me

End Sub

