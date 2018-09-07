VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LocForm 
   Caption         =   "Resx Validation"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6855
   OleObjectBlob   =   "LocForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Locform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GoButton_Click()
    Call LocValidation(LocInput.Value)
    Unload Locform
End Sub
