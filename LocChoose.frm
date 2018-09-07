VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LocChoose 
   Caption         =   "Localization Tasks"
   ClientHeight    =   2775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "LocChoose.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LocChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ExcelButton_Click()
    With ExcelLoc.ReleaseBox
        .AddItem "ACGL"
        .AddItem "ACDE"
        .AddItem "ACCN"
    End With
    
    ExcelLoc.VersionBox = "##"
    ExcelLoc.ErrorLabel = ""
    
    ExcelLoc.Show
    Unload LocChoose
End Sub

Private Sub ResxButton_Click()
    Locform.Show
    Unload LocChoose
End Sub
