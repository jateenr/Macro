Attribute VB_Name = "SAPPriceAudit"
Option Explicit

Private i As Double, x As Double, last As Double
Private SAP As Workbook, ASOMS As Workbook, WB As Workbook
Private asomscol As Object, sapcol As Object 'column numbers
Private valtext As String, difftext As String

Function xMatch(lookup_value As Variant, lookup_array As Range)     ' cleaner looking match function
      xMatch = WorksheetFunction.Match(lookup_value, lookup_array, 0)
End Function

Function xCountIf(lookup_array As Range, lookup_value As Variant)     ' cleaner looking match function
       xCountIf = WorksheetFunction.CountIf(lookup_array, lookup_value)
End Function

Sub DiffGrab(ASOMSname As String, SAPname As String)
    If Cells(i, asomscol(ASOMSname)) <> ASOMS.Sheets("SAP").Cells(x, sapcol(SAPname)) Then
        difftext = difftext & ", " & ASOMSname & ": " & ASOMS.Sheets("SAP").Cells(x, sapcol(SAPname))
        valtext = valtext & ", " & ASOMSname
    End If
End Sub

Sub SAPPriceAudit()
Set asomscol = CreateObject("Scripting.Dictionary")
Set sapcol = CreateObject("Scripting.Dictionary")
x = 0

Application.ScreenUpdating = False '++speed

    For Each WB In Workbooks    ' look for changeset file
        If Right(WB.Name, 14) = "SAP Audit.xlsx" Then
            Set ASOMS = WB
            x = x + 1
        ElseIf Right(WB.Name, 14) = "ACD Report.xls" Then
            Set SAP = WB        'look for SAP ACD export file
            x = x + 2
        End If
    Next
    
    If x = 0 Then
        MsgBox ("ASOMS and SAP files not found based on filenames." & vbNewLine & "Please open both the ACD Report and ASOMS Changeset files, " & _
        "With filenames ending in 'ACD Report.xls' and 'SAP Audit.xlsx'")
        Exit Sub
    ElseIf x = 1 Then
        MsgBox ("SAP file not found based on filename ending in 'ACD Report.xls'." & vbNewLine & "Please open the ACD Report file.")
        Exit Sub
    ElseIf x = 2 Then
        MsgBox ("ASOMS file not found based on filename ending in 'SAP Audit.xlsx'." & vbNewLine & "Please open the SAP Audit file and run the Changeset query in the first tab.")
        Exit Sub
    End If
    x = 0
    
    ASOMS.Sheets(1).Activate
    If Range("A1") = "" Or Range("A2") <> "ID" Then
        MsgBox ("Please run the ASOMS Changeset query in A1 of the first sheet" & vbNewLine & "of the SAP Audit worksheet prior to running the macro")
        Exit Sub
    End If
    
    SAP.Sheets(1).Activate
    If Range("C1") = "" Then    'removes blank rows on header
        Rows("1:" & Range("C1").End(xlDown).Row - 1).Delete
    End If
    If Range("A1") = "" Then
        Columns("A").Delete     'removes blank border column
    End If
    If Range("A2") = "" Then    'removes blank buffer row
        Rows(2).Delete
    End If
    
    ASOMS.Activate
    ASOMS.Sheets.Add after:=Sheets(Sheets.Count), Count:=2
    ASOMS.Sheets(Sheets.Count - 1).Name = "ASOMS"
    ASOMS.Sheets(Sheets.Count).Name = "SAP"
    
    ASOMS.Sheets(1).Activate
    For i = 1 To 30
        If Right(Cells(2, i), 4) = "Date" Then   'convert date ASOMS columns to date format before paste
            ASOMS.Sheets("ASOMS").Columns(i).NumberFormat = "m/d/yyyy"
        End If
    Next
    
    ASOMS.Sheets("ASOMS").Activate
    ActiveWindow.FreezePanes = False    'unfreeze view
    ASOMS.Sheets(1).Range("A2:BZ50000").Copy
    ASOMS.Sheets("ASOMS").Range("A1").PasteSpecial (xlValues)
    
    SAP.Sheets(1).Range("A1:BZ50000").Copy
    ASOMS.Sheets("SAP").Range("A1").PasteSpecial (xlValues)
        
    ASOMS.Sheets("ASOMS").Activate
    x = Range("BZ1").End(xlToLeft).Column   'get last column
    Cells(1, x + 1) = "Differences"
    Cells(1, x + 2) = "Validations"
    For i = 1 To x + 2
        asomscol.Add Cells(1, i).Value, i   'grab ASOMS col names
    Next
    
    ASOMS.Sheets("SAP").Activate
    x = Range("BZ1").End(xlToLeft).Column   'get last column
    For i = 1 To x
        sapcol.Add Cells(1, i).Value, i     'grab SAP col names
    Next
    
    ASOMS.Sheets("SAP").Activate
    ' If Offering Code ACP and SCE values don't match, then prompt user to end macro
    If xCountIf(Columns(sapcol("Offering Code")), "ACP") <> xCountIf(Columns(sapcol("Offering Code")), "SCE") Then
        If MsgBox("ACP/SCE Offering Code counts don't match- " & xCountIf(Columns(sapcol("Offering Code")), "ACP") & "/" & xCountIf(Columns(sapcol("Offering Code")), "SCE") & _
            vbNewLine & "Would you like to proceed?", vbYesNo) = 7 Then
            Exit Sub
        End If
    End If
    
    ASOMS.Sheets("SAP").Activate
    last = Range("A500000").End(xlUp).Row
    Range("A1:BZ" & last).RemoveDuplicates Columns:=Array(sapcol("Material Number"), sapcol("Monthly Global Base Price"), sapcol("Offering Code"))  'remove duplicates by SKU, Offering Code, and Price
    
    ASOMS.Sheets("ASOMS").Activate
    last = Range("A500000").End(xlUp).Row
    
    For i = 2 To last
        valtext = ""
        difftext = ""
        
        x = xCountIf(Sheets("SAP").Columns(sapcol("Material Number")), Cells(i, asomscol("Part Number"))) ' count SAP SKU/Offering Code/Price combos per SKU
        If x > 2 Then
            valtext = valtext & ", " & x & " Unique SAP SKU/Offer Code/Price Combos"
        ElseIf x = 0 Then valtext = valtext & ", SKU missing from SAP"
        End If
        
        On Error Resume Next    'bad error handling
        x = xMatch(Cells(i, asomscol("Part Number")), Sheets("SAP").Columns(sapcol("Material Number"))) 'SAP iterator
        
        Call DiffGrab("EA Rate", "Monthly Global Base Price")
        Call DiffGrab("SAP Discontinue Date", "Disco Date")
        
        If Left(Cells(i, asomscol("Part Number")), 2) = "AA" Then  ' if legacy SKU, don't check individual naming fields, as these are blank in SAP
            Call DiffGrab("Product Family", "Ext. ID 1")
            Call DiffGrab("Service Name", "Ext. ID 2")
            Call DiffGrab("Service Type", "Ext. ID 3")
            Call DiffGrab("Region Name", "Ext. ID 4")
            Call DiffGrab("EA Unit of Measure", "Ext. ID 5")
            Call DiffGrab("Material Description", "Ext. ID 6")
        End If
        On Error GoTo 0
        
        If Left(valtext, 1) = "," Then valtext = Mid(valtext, 3, 1000)    'remove leading commas
        If Left(difftext, 1) = "," Then difftext = Mid(difftext, 3, 1000)    'remove leading commas
        Cells(i, asomscol("Differences")) = difftext
        Cells(i, asomscol("Validations")) = valtext
    Next
    
    Columns("A:BZ").AutoFit

    If WorksheetFunction.CountA(ASOMS.Sheets("ASOMS").Columns(asomscol("Validations"))) > 1 Then
        MsgBox (WorksheetFunction.CountA(ASOMS.Sheets("ASOMS").Columns(asomscol("Validations"))) - 1 & " issues detected." & vbNewLine & "Please validate any issues in rightmost columns of the ASOMS sheet.")
    Else: MsgBox ("Audit complete. No issues detected")
    End If
    
    Set WB = Nothing
    Set ASOMS = Nothing
    Set SAP = Nothing
    Application.ScreenUpdating = True '++speed
    
End Sub
