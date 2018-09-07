VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExcelLoc 
   Caption         =   "Excel Loc Generator"
   ClientHeight    =   1710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5535
   OleObjectBlob   =   "ExcelLoc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExcelLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CreateButton_Click()
    
    On Error Resume Next
    If ExcelLoc.ReleaseBox = "" Then
        ExcelLoc.ErrorLabel = "Please select release type"
        Exit Sub
    ElseIf IsError(CInt(ExcelLoc.VersionBox)) Then
        ExcelLoc.ErrorLabel = "Please enter a numeric value"
        Exit Sub
    End If
    On Error GoTo 0
    
    Call ExcelLocFile
End Sub
