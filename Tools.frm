VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tools 
   Caption         =   "Home Page"
   ClientHeight    =   7140
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11505
   OleObjectBlob   =   "Tools.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Tools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DelimButton_Click()
    Unload Tools
    Deliminator.DelimValue = ","
    Deliminator.Show
End Sub

Private Sub EmailButton_Click()
    Unload Tools
    Emailer.Show
End Sub
