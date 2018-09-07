Attribute VB_Name = "Home"
Sub OpenTools()
    Tools.Show
End Sub

Sub OpenValidator()

    With Validator.ReleaseBox
        .AddItem "ACGL"
        .AddItem "ACDE"
        .AddItem "ACCN"
        .AddItem "SAP PL"
        .AddItem "SAP PL - China"
    End With
    
    Validator.GetProdData = False
    Validator.UpdateScope = True
    Validator.UpdateWorking = True
    Validator.CoreVals = True
    Validator.ModifyChecks = True
    Validator.EventLevel = True
    Validator.NameCheck = False
    Validator.Show
    
End Sub

Sub OpenLoc()
    LocChoose.Show
End Sub

Sub OpenMortem()
    Mortemizer.Show
End Sub
