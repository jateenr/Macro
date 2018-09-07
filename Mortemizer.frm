VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Mortemizer 
   Caption         =   "PostMortem"
   ClientHeight    =   3885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6390
   OleObjectBlob   =   "Mortemizer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Mortemizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FetchData_Click()
    Call FetchPostMortemData
    Unload Mortemizer
End Sub
