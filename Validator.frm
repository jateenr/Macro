VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Validator 
   Caption         =   "Validation Station"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12165
   OleObjectBlob   =   "Validator.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Validator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ReleaseBox_Change()
    If Left(Validator.ReleaseBox.Value, 2) = "AC" Then
        Validator.TextBox = "##"
    ElseIf Validator.ReleaseBox.Value = "SAP PL" Then
        Validator.TextBox = Format(DateAdd("m", 1, Date), "mm/yy")
    ElseIf Validator.ReleaseBox.Value = "SAP PL - China" Then
        Validator.TextBox = Format(DateAdd("m", 2, Date), "mm/yy")
    End If
End Sub

Private Sub GoButton_Click()
Dim entry As String, Valtype As String, SovCloud As String, Version As String

    Version = Validator.TextBox.Value
    Validator.ErrorLabel.Visible = False
    
    If Left(Validator.ReleaseBox.Value, 2) = "AC" Then
        Valtype = "Direct"
        If Validator.ReleaseBox.Value = "ACGL" Then SovCloud = "Global"
        If Validator.ReleaseBox.Value = "ACDE" Then SovCloud = "Germany"
        If Validator.ReleaseBox.Value = "ACCN" Then SovCloud = "China"
        On Error Resume Next
        If Validator.CreateFile Then
            If Not (IsNumeric(Version)) Then
                Validator.ErrorLabel = "Please enter a numeric version number"
                Validator.ErrorLabel.Visible = True
                Exit Sub
            End If
        End If
        On Error GoTo 0
    ElseIf Left(Validator.ReleaseBox.Value, 3) = "SAP" Then
        Valtype = "EA"
        If Validator.ReleaseBox.Value = "SAP PL" Then SovCloud = "Global"
        If Validator.ReleaseBox.Value = "SAP PL - China" Then SovCloud = "China"
        If Len(Version) <> 5 Then
            Validator.ErrorLabel = "Please enter Pricelist version as mm/yy"
            Validator.ErrorLabel.Visible = True
            Exit Sub
        End If
    End If
    
    Call RunValidations(Valtype, SovCloud, Version)
    Unload Validator
    
End Sub

Private Sub SelectAllOps_Click()
    If SelectAllOps Then
        Validator.CreateFile = True
        Validator.GetProdData = True
        Validator.GetProdDevTest = True
        Validator.UpdateWorking = True
        Validator.UpdateScope = True
    Else
        Validator.CreateFile = False
        Validator.GetProdData = False
        Validator.GetProdDevTest = False
        Validator.UpdateWorking = False
        Validator.UpdateScope = False
    End If
End Sub

Private Sub SelectAllVals_Click()
    If SelectAllVals Then
        Validator.CoreVals = True
        Validator.EventLevel = True
        Validator.ModifyChecks = True
        Validator.Changeset = True
        Validator.NameCheck = True
    Else
        Validator.CoreVals = False
        Validator.EventLevel = False
        Validator.ModifyChecks = False
        Validator.Changeset = False
        Validator.NameCheck = False
    End If
End Sub

Private Sub CreateFile_Click()

    If CreateFile Then
        Validator.UpdateWorking = False
        Validator.UpdateScope = False
        Validator.CoreVals = False
        Validator.EventLevel = False
        Validator.ModifyChecks = False
        Validator.Changeset = False
        Validator.NameCheck = False
    End If
End Sub
