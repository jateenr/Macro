Attribute VB_Name = "Validations"
Option Explicit

Private endcol As Integer
Private last As Double, i As Double, x As Double, y As Double, j As Double, a As Double, b As Double, c As Double
Private RelNum As Variant
Private wrkcol As Object, prdcol As Object, chngcol As Object, namecol As Object
Private str As String, WrkItm As String, keycol As String, val1 As String, val2 As String, val3 As String
Private colnames() As String, ModChecks() As String, Names() As String, badchars() As String, Issues() As String, ThirdParty() As String, Dates() As String, Changes() As String, SWIP() As String, NameChecks() As String, ChngChecks() As String, InputReq() As String
Private fpath As String, fname As String, shtname As String, valtext As String, modtext As String, chngtext As String
Private Wrk As Workbook, Prd As Workbook, Chng As Workbook, WB As Workbook, Namebook As Workbook
Private WS As Worksheet
Private PV As ProtectedViewWindow

Function xMatch(ByVal lookup_value As Variant, ByVal lookup_array As Range)     ' cleaner looking match function
      xMatch = WorksheetFunction.Match(lookup_value, lookup_array, 0)
End Function

Function IsIn(ByVal find_text As String, ByVal within_text As Variant)     ' cleaner looking match function
    On Error Resume Next
    IsIn = IsNumeric(WorksheetFunction.Find(find_text, within_text, 1))
    On Error GoTo 0
End Function

Function xCountIf(ByVal lookup_array As Range, ByVal lookup_value As Variant)     ' cleaner looking countif function
       xCountIf = WorksheetFunction.CountIf(lookup_array, lookup_value)
End Function

Function CountIf2(ByVal head1 As String, ByVal val1 As String, ByVal head2 As String, ByVal val2 As String)
    CountIf2 = WorksheetFunction.CountIfs(Sheets("Working").Columns(ltr(wrkcol(head1))), val1, Sheets("Working").Columns(ltr(wrkcol(head2))), val2)
End Function

Function ltr(ByVal colnum As Integer)
    ltr = Split(Cells(1, colnum).Address, "$")(1)
End Function

Function CellDiff(ByVal ColumnName As String)
    CellDiff = Cells(i, wrkcol(ColumnName)) <> Sheets("Prod").Cells(x, prdcol(ColumnName))
End Function

Sub RunValidations(ByVal Valtype As String, Optional ByVal SovCloud As String, Optional ByVal Version As String)
Set wrkcol = CreateObject("Scripting.Dictionary")
Set prdcol = CreateObject("Scripting.Dictionary")
If Validator.Changeset Then Set chngcol = CreateObject("Scripting.Dictionary")
If Validator.NameCheck Then Set namecol = CreateObject("Scripting.Dictionary")

    If Valtype = "EA" Then
        WrkItm = "Consumption SKU"
    ElseIf Valtype = "Direct" Then
        WrkItm = "Meter"
    End If
    
    Application.ScreenUpdating = False
    
    If Validator.CreateFile Then Call CreateWorkingFile(Valtype, SovCloud, Version)
    
    Call GetWorkingWB
    If Validator.Changeset Then Call GetChangesetWB
    If Validator.NameCheck Then Call GetNameWB(Valtype, SovCloud)
    If Validator.GetProdData Then Call GetReduceProd(Valtype, SovCloud)
    'If Validator.GetProdDevTest Then Call GetProdDevTest(ValType, SovCloud)
    
    If Validator.UpdateWorking Then Call UpdateWorking(Valtype)
    
    If Validator.ModifyChecks And (Wrk.Sheets("Prod").Range("A1") = "" Or IsIn("Server:", Wrk.Sheets("Prod").Range("A1"))) Then
        MsgBox ("Prod data not updated based on A1 of the sheet." & vbNewLine & "Either select 'Update Prod Data' or unselect 'Modify Existing'")
        Set Wrk = Nothing
        Set WB = Nothing            'if Modify existing selected and Prod data not updated, then end procedure
        Set Prd = Nothing
        Set Chng = Nothing
        Application.ScreenUpdating = True
        End
    End If
    
    If Validator.UpdateScope Or Validator.CoreVals Or Validator.EventLevel Or Validator.ModifyChecks Or Validator.Changeset Then
        'if operations besides setup selected, then continue, else end procedure
        
        Call GetColumns(Valtype)
        
        If Validator.UpdateScope Then Call UpdateScope(Valtype, Version)
        
        If Validator.CoreVals Or Validator.EventLevel Or Validator.ModifyChecks Or Validator.Changeset Or Validator.NameCheck Then
            Wrk.Sheets("Working").Activate
            endcol = Range("CZ1").End(xlToLeft).Column + 1
            last = Range("A200000").End(xlUp).Row
            
            For i = 2 To last   'cycle through every row
                valtext = ""
                
                If Validator.CoreVals Then
                    Call Duplicates(Valtype, i)
                    Call Regions(Valtype, i)
                    Call ValidNames(Valtype, SovCloud, i)
                    Call Rates(Valtype, i)
                    Call ChangeTypes(Valtype, i)
                    
                    If Valtype = "Direct" Then                                      'if different event from previous, then run event level checks
                        If Validator.EventLevel And Cells(i, wrkcol("EventId")) <> Cells(i - 1, wrkcol("EventId")) Then Call EventLevelChecks(Valtype, i)
                        Call Tags(i)
                    ElseIf Valtype = "EA" Then
                        If Validator.EventLevel And Cells(i, wrkcol("EventId")) <> Cells(i - 1, wrkcol("EventId")) Then Call EventLevelChecks(Valtype, i)
                        Call Discons(Version, i)
                    End If
                End If
                    
                If Validator.ModifyChecks Then
                    If Valtype = "Direct" Then
                        If Cells(i, wrkcol("Change Type")) = "Modify Existing" Then str = "mod"
                    ElseIf Valtype = "EA" Then
                        If Cells(i, wrkcol("Add/Update")) = "Update" Then str = "mod"
                    End If
                    
                    If str = "mod" Then
                        On Error Resume Next
                        If IsError(xMatch(Cells(i, wrkcol(keycol)), Sheets("Prod").Columns(prdcol(keycol)))) Then
                            valtext = valtext & ", Prod record not found"
                        Else
                        On Error GoTo 0
                            x = xMatch(Cells(i, wrkcol(keycol)), Sheets("Prod").Columns(prdcol(keycol))) 'grab the matching prod data row
                            
                            Call ModNames(Valtype, i, x)
                            Call ModPrice(Valtype, i, x)
                            Call ModStatus(Valtype, i, x)
                            Call ModOther(Valtype, i, x)
                            
                            If Valtype = "EA" Then
                            
                            ElseIf Valtype = "Direct" Then
                                Call ModMeterTag(i, x)
                            End If
        
                        End If
                        On Error GoTo 0
                    End If
                End If
                
                If Validator.Changeset Then
                    Call ChangesetChecks(Valtype, i)
                End If
                
                If Validator.NameCheck Then
                    Call NameCheck(Valtype, i)
                End If
                
                If Left(valtext, 1) = "," Then valtext = Mid(valtext, 3, 500)    'remove leading commas
                Cells(i, wrkcol("Macro Validation")) = valtext
                str = ""
            Next
            
            If Validator.NameCheck And xCountIf(Columns(ltr(wrkcol("Macro Validation"))), "*Name sheet*") = 0 Then
                Namebook.Close savechanges:=False   ' if no Name sheet issues, close Name sheet
            End If
        End If
    End If
    
    Set wrkcol = Nothing
    Set prdcol = Nothing
    Set chngcol = Nothing
    Set namecol = Nothing
    Set Wrk = Nothing
    Set WB = Nothing
    Set WS = Nothing
    Set Prd = Nothing
    Set Chng = Nothing
    Set Namebook = Nothing
    Application.ScreenUpdating = True
    
End Sub

Sub EventLevelChecks(ByVal Valtype As String, ByVal i As Double)
x = 0
    
    x = xMatch(CDbl(Cells(i, wrkcol("EventId"))), Sheets("Original").Columns(ltr(wrkcol("ID"))))
    
    If Valtype = "EA" Then
        val1 = Sheets("Original").Cells(x, wrkcol("State") - 1) 'minus 1 to account for Add/Update added column
        
    ElseIf Valtype = "Direct" Then
        val1 = Sheets("Original").Cells(x, wrkcol("State"))
        
        If Sheets("Original").Cells(x, wrkcol("CP Rate Start Date")) > DateAdd("d", 40, Date) Then
            valtext = valtext & ", Event CP Rate Start Date 40+ days out"
        End If
    End If

    If val1 <> "Approved" And val1 <> "In Progress" Then    'And val1 <> "Reviewed"
        valtext = valtext & ", Unexpected Event State"
    End If
    
End Sub

Sub Duplicates(ByVal Valtype As String, ByVal i As Double)

    If Valtype = "Direct" Then
        If Cells(i, wrkcol("Resource GUID")) <> "" Then
            If xCountIf(Columns(ltr(wrkcol("Resource GUID"))), Cells(i, wrkcol("Resource GUID"))) > 1 Then valtext = valtext & "Duplicate GUID"
        End If
                        'this ugly excessive code just checks if there are dupes between all naming fields
        If WorksheetFunction.CountIfs(Columns(ltr(wrkcol("Service Name"))), Cells(i, wrkcol("Service Name")), Columns(ltr(wrkcol("Service Type"))), Cells(i, wrkcol("Service Type")), _
        Columns(ltr(wrkcol("Resource Name"))), Cells(i, wrkcol("Resource Name")), Columns(ltr(wrkcol("Region Name"))), Cells(i, wrkcol("Region Name")), _
        Columns(ltr(wrkcol("Meter Status"))), "Active") > 1 Then valtext = valtext & ", Duplicate name"
         'Columns(ltr(wrkcol("Direct Unit of Measure"))), Cells(i, wrkcol("Direct Unit of Measure")))
        
    ElseIf Valtype = "EA" Then
        If Cells(i, wrkcol("Part Number")) <> "" Then
            If xCountIf(Columns(ltr(wrkcol("Part Number"))), Cells(i, wrkcol("Part Number"))) > 1 Then valtext = valtext & ", Duplicate Part Number"
        End If
        If xCountIf(Columns(ltr(wrkcol("Material Description"))), Cells(i, wrkcol("Material Description"))) > 1 Then valtext = valtext & ", Duplicate Material Description"
        If xCountIf(Columns(ltr(wrkcol("EA Portal Friendly Name"))), Cells(i, wrkcol("EA Portal Friendly Name"))) > 1 Then valtext = valtext & ", Duplicate EA Portal Friendly Name"
    End If
    
End Sub

Sub Regions(ByVal Valtype As String, ByVal i As Double)
val1 = Cells(i, wrkcol("Azure Instance")).Value
val2 = Cells(i, wrkcol("Region Name")).Value

    If val1 = "Global" And (Left(val2, 2) = "DE" Or val2 = "Gov (US)" Or val2 = "US Gov AZ" Or val2 = "US Gov TX" Or val2 = "US DoD" Or val2 = "USGov" Or val2 = "DoD (US)") Then
        valtext = valtext & ", Region and Instance don't match"
    ElseIf val1 = "USGov" And Not (val2 = "Gov (US)" Or val2 = "US Gov AZ" Or val2 = "US Gov TX" Or val2 = "US DoD" Or val2 = "USGov" Or val2 = "DoD (US)") Then
        valtext = valtext & ", Region and Instance don't match"
    ElseIf val1 = "Germany" And Left(val2, 2) <> "DE" Then
        valtext = valtext & ", Region and Instance don't match"
    End If
    
    If val2 = "US Gov AZ" Then
        If StrComp("AZ", Right(val2, 2), vbBinaryCompare) <> 0 Then valtext = valtext & ", lowercase in region name"
    ElseIf val2 = "US Gov TX" Then
        If StrComp("TX", Right(val2, 2), vbBinaryCompare) <> 0 Then valtext = valtext & ", lowercase in region name"
    End If
    
    If Valtype = "EA" Then
        If (val2 = "UK North" Or val2 = "UK South 2") And Cells(i, wrkcol("Sku Is Permanent Lead")) <> "Yes" Then
            valtext = valtext & ", Whitelisted Region not Perm Lead"
        End If
        If (val1 = "USGov" Or val1 = "Germany") And Cells(i, wrkcol("Sku Sub Type")) = "DevTest" Then valtext = valtext & ", USGov/Germany DevTest"
    
    ElseIf Valtype = "Direct" Then
        If (val2 = "UK North" Or val2 = "UK South 2") And Cells(i, wrkcol("Permanent Lead Status")) <> "Yes" Then
            valtext = valtext & ", Whitelisted Region not Perm Lead"
        End If
        
        If Cells(i, wrkcol("China Only Meter")) = "Yes" And Cells(i, wrkcol("Dev/Test Eligible")) = "Yes" Then valtext = valtext & ", DevTest for China Only"
        
        If val1 = "China" Then  'if China meter check for legacy meters which are not being deprecated
            If (val2 = "" Or val2 = "Zone 1" Or val2 = "Azure Stack") And Cells(i, wrkcol("Meter Status")) = "Active" Then valtext = valtext & ", Legacy China meter not deprecating"
        End If
    End If
    
End Sub

Sub ValidNames(ByVal Valtype As String, ByVal SovCloud As String, ByVal i As Double)
badchars = Split("#,^,!,@,{,},[,],  ," & Chr(151), ",") 'Chr 151 is an em dash
val1 = Cells(i, wrkcol("Launch Stage"))

    For j = 0 To UBound(Names)  'cycle through all text description fields
        For x = 0 To UBound(badchars)
            If IsIn(badchars(x), Cells(i, wrkcol(Names(j)))) Then   ' check cells for bad characters listed above"
                valtext = valtext & ", " & Cells(i, wrkcol(Names(j))) & " contains '" & badchars(x) & "'"
            End If
        Next
    Next
    x = 0
    
    If IsIn("VM", Cells(i, wrkcol("Service Type"))) And IsIn("Cloud Service", Cells(i, wrkcol("Service Type"))) Then valtext = valtext & ", Cloud Service VM not allowed"
    
    If Valtype = "EA" Then
        If Not IsNumeric(Left(Cells(i, wrkcol("EA Unit of Measure")), 1)) Then valtext = valtext & ", UoM without leading number"

        If Cells(i, wrkcol("Add/Update")) = "Add" Or (Cells(i, wrkcol("Has Ea Rate Decreased")) = "Yes" Or Cells(i, wrkcol("Has Ea Rate Increased")) = "Yes") Then
            If IsIn("000", Cells(i, wrkcol(Valtype & " Unit of Measure"))) Then ' look for 1000s instead of 1K/1M
                valtext = valtext & ", 000 in UoM"
            End If
        End If
    
        If SovCloud = "China" And Cells(i, wrkcol("Product Family")) <> "Azure China OE" Then
            valtext = valtext & ", Should have China PFAM"
        End If
        
        If IsIn("Preview", val1) Then
            If Cells(i, wrkcol("Product Family")) <> "Azure Services in Preview" And SovCloud <> "China" Then valtext = valtext & ", Should have Preview PFAM"
            If Not IsIn("Preview", Cells(i, wrkcol("EA Portal Friendly Name"))) Then valtext = valtext & ", 'Preview' not in EAP Friendly Name"
        Else
            If Cells(i, wrkcol("Product Family")) = "Azure Services in Preview" Then valtext = valtext & ", Should not have Preview PFAM"
            If IsIn("Preview", Cells(i, wrkcol("EA Portal Friendly Name"))) Then valtext = valtext & ", 'Preview' should not be in EAP Friendly Name"
        End If
        
        If Cells(i, wrkcol("Service Name")) = "Virtual Machines" And Not IsIn("Windows", Cells(i, wrkcol("Service Type"))) And _
            Cells(i, wrkcol("Sku Sub Type")) = "DevTest" Then valtext = valtext & ", Linux DevTest"
    
    ElseIf Valtype = "Direct" Then
        If Cells(i, wrkcol("Change Type")) = "Create New" Or Cells(i, wrkcol("Has Name Change")) = "Yes" Then
            If IsIn("000", Cells(i, wrkcol(Valtype & " Unit of Measure"))) Then ' look for 1000s instead of 1K/1M
                valtext = valtext & ", 000 in UoM"
            End If
        End If
    End If
End Sub

Sub Rates(ByVal Valtype As String, ByVal i As Double)
val1 = Cells(i, wrkcol("Resource Name"))
x = Cells(i, wrkcol(Valtype & " Rate"))

    If Valtype = "Direct" Then
        If x = 0 Then       ' if zero priced, and resource Name doesn't suggest Free, and not Private Preview, and not China Only, throw exception
            If Cells(i, wrkcol("Change Type")) = "Create New" Or Cells(i, wrkcol("Has Price Increase")) = "Yes" Or Cells(i, wrkcol("Has Price Decrease")) = "Yes" Then
                If Not (IsIn("Free", val1) Or IsIn("Trial", val1) Or IsIn("Delete Operation", val1) _
                    Or Cells(i, wrkcol("Launch Stage")) = "Private Preview" Or Cells(i, wrkcol("China Only Meter")) = "Yes") Then
                    valtext = valtext & ", Unexpected zero price"
                End If
            End If
        Else
            If x <= Cells(i, wrkcol("DevTest Discount Rate")) Then
                valtext = valtext & ", No DevTest discount"
            End If
            If x <= Cells(i, wrkcol("Graduated Tier 1 Discount Rate")) Then
                valtext = valtext & ", No Grad Rate discount"
            End If
            
            If Cells(i, wrkcol("Dev/Test Eligible")) = "Yes" Then
                val2 = Cells(i, wrkcol("DevTest Rate Source Meter ID"))
                If val2 = "" Then   'if DevTest Source is blank throw exception for Modifies
                    'If Cells(i, wrkcol("Change Type")) = "Modify Existing" Then valtext = valtext & ", Expected DevTest source meter"
                Else
                    a = Cells(i, wrkcol("DevTest Discount Rate"))
                    On Error Resume Next
                    b = Cells(xMatch(val2, Columns(ltr(wrkcol("Resource GUID")))), wrkcol("Direct Rate"))
                    c = Sheets("Prod").Cells(xMatch(val2, Sheets("Prod").Columns(ltr(prdcol("Resource GUID")))), prdcol("Direct Rate"))
                    On Error GoTo 0
                    
                    If a <> b And a <> c Then valtext = valtext & ", DevTest Rate does not match source rate"
                    
                End If
            End If
        End If
        
        If Cells(i, wrkcol("Graduated Tier 1 Discount Rate")) <> "" Then
            If Cells(i, wrkcol("Tier Rate for EA Consumption SKU")) = "" Then valtext = valtext & ", Missing graduated Tier Rate for SKU"
            If Cells(i, wrkcol("Has Graduated Rate")) = "No" Then valtext = valtext & ", Unexpected graduated rate"
            For j = 1 To 5      'loop through each graduated rate tier to check is cheaper than previous tier
                If Cells(i, wrkcol("Graduated Tier " & j + 1 & " Discount Rate")) <> "" Then
                    If Cells(i, wrkcol("Graduated Tier " & j & " Discount Rate")) <= Cells(i, wrkcol("Graduated Tier " & j + 1 & " Discount Rate")) Then
                        valtext = valtext & ", Grad Rate " & j + 1 & " not discounted"
                    End If
                End If
            Next
        ElseIf Cells(i, wrkcol("Has Graduated Rate")) = "Yes" Then valtext = valtext & ",  Missing graduated rate"
        End If
        
    ElseIf Valtype = "EA" Then
        If Cells(i, wrkcol("EA Rate")) = 0 Then
            If Cells(i, wrkcol("Change Type")) = "Create New" Or Cells(i, wrkcol("Has Ea Rate Decreased")) = "Yes" Or Cells(i, wrkcol("Has Ea Rate Increased")) = "Yes" Then
                If Not (IsIn("Free", val1) Or IsIn("Trial", val1) Or IsIn("Delete Operation", val1) _
                    Or Cells(i, wrkcol("Launch Stage")) = "Private Preview" Or Cells(i, wrkcol("China Only Meter")) = "Yes") Then
                        valtext = valtext & ", Unexpected zero price"
                End If
            End If
        ElseIf Cells(i, wrkcol("Add/Update")) = "Add" Or (Cells(i, wrkcol("Has Ea Rate Decreased")) = "Yes" Or Cells(i, wrkcol("Has Ea Rate Increased")) = "Yes") Then
            If Cells(i, wrkcol("EA Rate")) < 0.5 Then   ' if EA rate is new or changing check if within expected range
                valtext = valtext & ", EA Rate < .50"
            ElseIf Cells(i, wrkcol("EA Rate")) > 150 Then
                valtext = valtext & ", EA Rate > 150"
            End If
        End If
        If Cells(i, wrkcol("Material Description")) = "" Then
            valtext = valtext & ", Blank Material Description"
        ElseIf Cells(i, wrkcol("Sku Sub Type")) = "DevTest" And x > 0 Then
                If xCountIf(Columns(ltr(wrkcol("Material Description"))), Replace(Cells(i, wrkcol("Material Description")), " DvTst", "")) > 0 Then
                    str = ""
                ElseIf xCountIf(Columns(ltr(wrkcol("Material Description"))), Replace(Cells(i, wrkcol("Material Description")), " DvTst", " Promo")) > 0 Then
                    str = " Promo"
                Else: valtext = valtext & ", DevTest with no Promo or Consumption SKU"
                End If
            If x >= Cells(xMatch(Replace(Cells(i, wrkcol("Material Description")), " DvTst", str), Columns(ltr(wrkcol("Material Description")))), wrkcol("EA Rate")) Then
                valtext = valtext & ", No DevTest Discount"
            End If
        End If
        
        If Cells(i, wrkcol("Discount Slope")) <> "0/0/0/0" Then valtext = valtext & ", Unexpected Discount Slope"
    End If
    
End Sub

Sub Tags(ByVal i As Double)
ThirdParty = Split("Autodesk,Canonical,Citrix,Java Development Environment,OpenLogic,Oracle,RedHat,RHEL,SLES,V-Ray,XenApp,Ubuntu", ",")
SWIP = Split("Linux,VM Support,Citrix,Remote Access Rights,Autodesk,BizTalk,Canonical,OpenLogic,Oracle,Java,Red Hat,RHEL,SLES,SSIS,SQL Server,Xen,Reservation-Windows Svr,V-Ray", ",")
x = 0

    If IsIn("Visual Studio", Cells(i, wrkcol("Service Name"))) And Cells(i, wrkcol("Visual Studio Tag")) <> "Yes" Then
        valtext = valtext & ", Visual Studio Tag missing"
    End If
    
    For j = 0 To UBound(ThirdParty)
        If IsIn(ThirdParty(j), Cells(i, wrkcol("Service Name"))) Or IsIn(ThirdParty(j), Cells(i, wrkcol("Service Type"))) Then x = x + 1
    Next
    
    If Cells(i, wrkcol("3rd Party Tag")) <> "Yes" And x > 0 Then valtext = valtext & ", 3rd Party Tag missing"
    If Cells(i, wrkcol("3rd Party Tag")) = "Yes" And x = 0 Then valtext = valtext & ", Unexpected 3rd Party Tag"
    x = 0
    
    For j = 0 To UBound(SWIP)
        If IsIn(SWIP(j), Cells(i, wrkcol("Service Name"))) Or IsIn(SWIP(j), Cells(i, wrkcol("Service Type"))) Then x = x + 1
    Next
    
    If Cells(i, wrkcol("SW IP Meter")) <> "Yes" And x > 0 Then valtext = valtext & ", SW IP Tag missing"
    If Cells(i, wrkcol("SW IP Meter")) = "Yes" And x = 0 Then valtext = valtext & ", Unexpected SW IP Tag"
    If Cells(i, wrkcol("Change Type")) = "Create New" And x > 0 And Cells(i, wrkcol("Region Name")) <> "" Then valtext = valtext & ", Regional SW IP meter"
    
End Sub

Sub Discons(ByVal Version As String, ByVal i As Double)
val1 = Cells(i, wrkcol("SAP Discontinue Date"))

        If Cells(i, wrkcol("Sku Sub Type")) = "Promo" Then
            If val1 = "12/31/2030" Or val1 = "" Then valtext = valtext & ", Promo SKU without discon date"
        ElseIf val1 <> "12/31/2030" And val1 <> "" Then
            If Cells(i, wrkcol("Add/Update")) = "Add" Then valtext = valtext & ", Discon date for new SKU"
            If val1 <> DateAdd("D", -1, Left(Version, 2) & "/01/" & Right(Version, 2)) Then valtext = valtext & ", Unexpected Discon Date"
        End If          ' if not end of prior month

End Sub

Sub ChangeTypes(ByVal Valtype, ByVal i As Double)
x = 0

    For j = 0 To UBound(Changes)
        If Cells(i, wrkcol(Changes(j))) = "Yes" Then
            x = x + 1
        End If
    Next

    If Valtype = "EA" Then
        If Cells(i, wrkcol("Change Type")) = "Modify Existing" Then
            If Cells(i, wrkcol("Part Number")) = "" And Cells(i, wrkcol("Launch Stage")) = Cells(i, wrkcol("From Launch Stage")) Then
                valtext = valtext & ", Part Number missing for modify"
            End If
        ElseIf Cells(i, wrkcol("Change Type")) = "Create New" Then
            If Cells(i, wrkcol("Part Number")) <> "" And Cells(i, wrkcol("Launch Stage")) = Cells(i, wrkcol("From Launch Stage")) Then
                valtext = valtext & ", Part Number present for new"
            End If
        End If
        
        If x > 0 Then
            If Cells(i, wrkcol("Add/Update")) = "Add" Then valtext = valtext & ", Changes noted for New"
        ElseIf Cells(i, wrkcol("Add/Update")) = "Update" Then
            valtext = valtext & ", No changes noted for Modify"
        End If
        
    ElseIf Valtype = "Direct" Then
        If x > 0 Then
            If Cells(i, wrkcol("Change Type")) = "Create New" And Cells(i, wrkcol("Has Meter Status Changed")) <> "Yes" Then valtext = valtext & ", Changes noted for New"
        ElseIf Cells(i, wrkcol("Change Type")) = "Modify Existing" Then
            valtext = valtext & ", No changes noted for Modify"
        End If
    End If
    
End Sub

Sub Informative(ByVal Valtype As String, ByVal i As Double)
    
    If Valtype = "Direct" Then
        If Cells(i, wrkcol("Has Graduated Rate Change")) = "Yes" And Cells(i, wrkcol("Graduated Tier 1 Discount Rate")) = "" Then valtext = valtext & ",  Grad Rate needs manual deletion"
        If Cells(i, wrkcol("Has DevTest Percent Change")) = "Yes" And Cells(i, wrkcol("DevTest Discount Percentage")) = "" Then valtext = valtext & ",  DevTest needs manual deletion"
    ElseIf Valtype = "EA" Then
    
    End If

End Sub

Sub ModNames(ByVal Valtype As String, ByVal i As Double, ByVal x As Double)
y = 0
modtext = ""

    For j = 0 To UBound(Names)  'cycle through all text description fields
        If CellDiff(Names(j)) Then
            modtext = modtext & ", " & Names(j) & ": " & Sheets("Prod").Cells(x, prdcol(Names(j)))
            y = y + 1
        End If
    Next
    
    If Left(modtext, 1) = "," Then modtext = Mid(modtext, 3, 500)    'remove leading commas
    Cells(i, wrkcol("NameChange")) = modtext
    
    If IsIn("Unit of Measure", Cells(i, wrkcol("NameChange"))) Then
        If InStr(1, Cells(i, wrkcol(Valtype & " Unit of Measure")), " ") > 0 And InStr(1, Sheets("Prod").Cells(x, prdcol(Valtype & " Unit of Measure")), " ") Then
            val1 = Left(Cells(i, wrkcol(Valtype & " Unit of Measure")), InStr(1, Cells(i, wrkcol(Valtype & " Unit of Measure")), " ") - 1)
            val2 = Left(Sheets("Prod").Cells(x, prdcol(Valtype & " Unit of Measure")), InStr(1, Sheets("Prod").Cells(x, prdcol(Valtype & " Unit of Measure")), " ") - 1)
            
            If val1 <> val2 And Not (CellDiff(Valtype & " Rate")) Then
                valtext = valtext & ", UoM change without rate change"
            End If
        End If
        
        If Valtype = "EA" Then
            If Cells(i, wrkcol("Has Ea Uom Changed")) = "No" Then valtext = valtext & ", Unexpected UoM change"
        End If
    End If
    
    If Valtype = "EA" Then

    ElseIf Valtype = "Direct" Then
        If y > 0 And Cells(i, wrkcol("Has Name Change")) = "No" Then valtext = valtext & ", Unexpected name change"
        If y = 0 And Cells(i, wrkcol("Has Name Change")) = "Yes" Then valtext = valtext & ", Missing name change"
        
            'if a Service Type changes to a blank value, the old values need to be deleted from Cayman after the build import
        If Cells(i, wrkcol("Service Type")) = "" And IsIn("Service Type", Cells(i, wrkcol("NameChange"))) Then valtext = valtext & ", Service Type needs manual deletion"
    End If
    
End Sub

Sub ModPrice(ByVal Valtype, ByVal i As Double, ByVal x As Double)
a = Sheets("Working").Cells(i, wrkcol(Valtype & " Rate"))
b = Sheets("Prod").Cells(x, prdcol(Valtype & " Rate"))

    If Valtype = "EA" Then
        val1 = "Has Ea Rate Decreased"
        val2 = "Has Ea Rate Increased"
    ElseIf Valtype = "Direct" Then
        val1 = "Has Price Decrease"
        val2 = "Has Price Increase"
    End If

    If a < b Then
        Cells(i, wrkcol("PriceChange")) = "Decr: " & b
        If Cells(i, wrkcol(val1)) <> "Yes" Then valtext = valtext & ", Unexpected price decrease"
    ElseIf a > b Then
        Cells(i, wrkcol("PriceChange")) = "Incr: " & b
        If Cells(i, wrkcol(val2)) <> "Yes" Then valtext = valtext & ", Unexpected price increase"
    End If
    
End Sub

Sub ModStatus(ByVal Valtype, ByVal i As Double, ByVal x As Double)
    
    If Valtype = "EA" Then
        If CellDiff("SAP Discontinue Date") Then
            modtext = modtext & ", SAP Disco Date: " & Sheets("Prod").Cells(x, prdcol("SAP Discontinue Date"))
        End If
    ElseIf Valtype = "Direct" Then
        If CellDiff("Meter Status") Then
            modtext = modtext & ", Meter Status: " & Sheets("Prod").Cells(x, prdcol("Meter Status"))
            If Cells(i, wrkcol("Has Meter Status Changed")) <> "Yes" Then valtext = valtext & ", Unexpected meter status change"
        ElseIf Cells(i, wrkcol("Has Meter Status Changed")) = "Yes" Then
            valtext = valtext & ", Expected meter status change"
        End If
    End If
    
End Sub

Sub ModPublicDate(ByVal Valtype, ByVal i As Double, ByVal x As Double)
    
    If Valtype = "EA" Then
        If CellDiff("Public Status Date") Then
            modtext = modtext & ", Public Status Date: " & Sheets("Prod").Cells(x, prdcol("Public Status Date"))
            If Cells(i, wrkcol("Has Public Status Date Changed")) <> "Yes" Then valtext = valtext & ", Unexpected Public Status Date change"
        ElseIf Cells(i, wrkcol("Has Public Status Date Changed")) = "Yes" Then
            valtext = valtext & ", Expected Public Status Date change"
        End If
    ElseIf Valtype = "Direct" Then
        If Format(Cells(i, wrkcol("Public Disclosure Date")), "m/d/yyyy") <> Format(Sheets("Prod").Cells(x, prdcol("Public Disclosure Date")), "m/d/yyyy") Then
            modtext = modtext & ", Disclosure Date: " & Sheets("Prod").Cells(x, prdcol("Public Disclosure Date"))
        End If
    End If
    
End Sub

Sub ModRevSku(ByVal i As Double, ByVal x As Double)
    
    If CellDiff("Revenue SKU") Then
        modtext = modtext & ", Revenue SKU: " & Sheets("Prod").Cells(x, prdcol("Revenue SKU"))
        If Cells(i, wrkcol("Has Revenue Sku Change")) <> "Yes" Then valtext = valtext & ", Unexpected revenue SKU change"
    ElseIf Cells(i, wrkcol("Has Revenue Sku Change")) = "Yes" Then
        valtext = valtext & ", Expected revenue SKU change"
    End If
    
End Sub

Sub ModPermLead(ByVal i As Double, ByVal x As Double)
    
    If CellDiff("Sku Is Permanent Lead") Then
        modtext = modtext & ", Sku Is Permanent Lead: " & Sheets("Prod").Cells(x, prdcol("Sku Is Permanent Lead"))
    End If
    
End Sub

Sub ModMeterTag(ByVal i As Double, ByVal x As Double)
val1 = Cells(i, wrkcol("3rd Party Tag"))
val2 = Cells(i, wrkcol("Visual Studio Tag"))
val3 = Sheets("Prod").Cells(x, prdcol("MeterTags"))

    If val1 = "Yes" And Not IsIn("Third Party", val3) Then
        valtext = valtext & ", 3rd Party Tag added "
    ElseIf val1 = "No" And IsIn("Third Party", val3) Then
        valtext = valtext & ", 3rd Party Tag removed"
    End If
    
    If val2 = "Yes" And Not IsIn("Visual Studio", val3) Then
        valtext = valtext & ", Visual Studio Tag added "
    ElseIf val2 = "No" And IsIn("Visual Studio", val3) Then
        valtext = valtext & ", Visual Studio Tag removed"
    End If
    
End Sub

Sub ModOther(ByVal Valtype, ByVal i As Double, ByVal x As Double)
modtext = ""

    Call ModPublicDate(Valtype, i, x)
    Call ModStatus(Valtype, i, x)
    
    If Valtype = "EA" Then
        Call ModPermLead(i, x)
    ElseIf Valtype = "Direct" Then
        Call ModRevSku(i, x)
    End If
    
    If Left(modtext, 1) = "," Then modtext = Mid(modtext, 3, 500)    'remove leading commas
    Cells(i, wrkcol("OtherModifies")) = modtext
    
End Sub

Sub ChangesetChecks(ByVal Valtype As String, ByVal i As Double)
    
    chngtext = ""
    
    If Valtype = "EA" Then
        str = "Material Description"
    ElseIf Valtype = "Direct" Then
        str = "Resource GUID"
    End If
    
    If Valtype = "Direct" And Cells(i, wrkcol("Resource GUID")) = "" Then
        chngtext = chngtext & ", Blank GUID- Update Working File"   'GUIDs are used for comparison, so these need to be added for changeset analysis
    Else
        On Error Resume Next
        If IsError(xMatch(Cells(i, wrkcol(str)), Chng.Sheets("Changeset").Columns(ltr(chngcol(str))))) Then
            chngtext = chngtext & ", Missing from Changeset based on " & str
        ElseIf Valtype = "Direct" And (Cells(i, wrkcol("Meter Status")) = "Deprecated" Or Cells(i, wrkcol("Meter Status")) = "Never Used") Then
            'If deprecting meter, ignore differences between sheets
        Else
        On Error GoTo 0
            x = xMatch(Cells(i, wrkcol(str)), Chng.Sheets("Changeset").Columns(ltr(chngcol(str))))
            
            For j = 0 To UBound(ChngChecks)
                If Cells(i, wrkcol(ChngChecks(j))) <> Chng.Sheets("Changeset").Cells(x, chngcol(ChngChecks(j))) Then
                    If Valtype = "Direct" And (ChngChecks(j) = "Direct Rate" Or ChngChecks(j) = "CP Rate Start Date") And Cells(i, wrkcol("Has Price Decrease")) <> "Yes" And _
                        Cells(i, wrkcol("Has Price Increase")) <> "Yes" And Chng.Sheets("Changeset").Cells(x, chngcol(ChngChecks(j))) = "" Then
                            'Ignore blank values for Price/CP Rate Start Date if no changes noted for meter
                    Else
                    chngtext = chngtext & ", " & ChngChecks(j) & ": " & Chng.Sheets("Changeset").Cells(x, chngcol(ChngChecks(j)))
                    End If
                End If
            Next
            
        End If
        On Error GoTo 0
        
    End If
    
    If Left(chngtext, 1) = "," Then chngtext = Mid(chngtext, 3, 500)    'remove leading commas
    Cells(i, wrkcol("Changeset")) = chngtext
    
    If Len(chngtext) > 0 Then valtext = valtext & ", Changeset Issues"
    
End Sub

Sub NameCheck(ByVal Valtype As String, ByVal i As Double)
x = 0
    
    If Valtype = "EA" Then
        If (Cells(i, wrkcol("Has Material Description Changed")) = "Yes" Or Cells(i, wrkcol("NameChange")) <> "") Then
            On Error Resume Next    'check if SKU missing from Name sheet missing
            If IsError(xMatch(Cells(i, wrkcol("Part Number")), Namebook.Sheets(shtname).Columns(namecol("ConsumptionPartNumber")))) Then
                valtext = valtext & ", SKU not found in Name sheet"
            Else
                On Error GoTo 0
                x = xMatch(Cells(i, wrkcol("Part Number")), Namebook.Sheets("EA Names").Columns(namecol("ConsumptionPartNumber"))) 'get line item in the name sheet
            
                If Cells(i, wrkcol("Material Description")) <> Namebook.Sheets("EA Names").Cells(x, namecol("SAPPartDescriptionFromTemplate")) Then valtext = valtext & ", Mat. Desc. doesn't match Name sheet"
                If Cells(i, wrkcol("EA Portal Friendly Name")) <> Namebook.Sheets("EA Names").Cells(x, namecol("EAPortalDescriptionFromTemplate")) Then valtext = valtext & ", EA Portal name doesn't match Name sheet"
                
            End If
            On Error GoTo 0
        End If
    ElseIf Valtype = "Direct" Then
        If (Cells(i, wrkcol("Has Name Change")) = "Yes" Or Cells(i, wrkcol("NameChange")) <> "") Then
            On Error Resume Next    'check if GUID missing from Name sheet missing
            If IsError(xMatch(Cells(i, wrkcol("Resource GUID")), Namebook.Sheets(shtname).Columns(namecol("MeterId")))) Then
                valtext = valtext & ", GUID not found in Name sheet"
            Else
                On Error GoTo 0
                x = xMatch(Cells(i, wrkcol("Resource GUID")), Namebook.Sheets(shtname).Columns(namecol("MeterId"))) 'get line item in the name sheet
            
                For j = 0 To UBound(NameChecks())
                    If Cells(i, wrkcol(Names(j))) <> Namebook.Sheets(shtname).Cells(x, namecol(NameChecks(j))) Then valtext = valtext & ", " & Names(j) & " does not match Name sheet"
                Next
            End If
            On Error GoTo 0
        End If
    End If


End Sub

Sub checkitz()
    Call RunValidations("Direct", "Global")
    Application.ScreenUpdating = True
End Sub

Sub runitz()
    Application.ScreenUpdating = True
End Sub

Sub GetColumns(ByVal Valtype As String)
        
    If Valtype = "EA" Then
        'text description fields below
        Names = Split("Product Family,Service Name,Service Type,Resource Name,Region Name,EA Unit of Measure,EA Portal Friendly Name,Material Description", ",")
        Dates = Split("SAP Rate Start Date,SAP Discontinue Date,Public Price List Date,Public Status Date", ",")
        Changes = Split("Has Material Description Changed,Has Included Quantity Changed,Has Discount Slope Changed,Has Ea Rate Decreased," & _
            "Has Ea Rate Increased,Has Ea Uom Changed,Has Public Status Date Changed,Has Sku Changed", ",")
        ModChecks = Split("NameChange,PriceChange,OtherModifies", ",")
        ChngChecks = Split("Product Family,Service Name,Service Type,Resource Name,Region Name,EA Unit of Measure,SAP Rate Start Date,SAP Discontinue Date,EA Rate,Discount Slope,Part Number,Resource GUID,Azure Instance,Material Description", ",")
        keycol = "Part Number"
        
    ElseIf Valtype = "Direct" Then
        Sheets("Prod").Activate
        If Range("A1") = "ResourceId" Then
            Cells(1, xMatch("ResourceId", Range("A1:BZ1"))) = "Resource GUID"
            Cells(1, xMatch("Service", Range("A1:BZ1"))) = "Service Name"
            Cells(1, xMatch("ServiceType", Range("A1:BZ1"))) = "Service Type"   'change Cayman column names to match ASOMS
            Cells(1, xMatch("FriendlyName", Range("A1:BZ1"))) = "Resource Name"
            Cells(1, xMatch("Region", Range("A1:BZ1"))) = "Region Name"
            Cells(1, xMatch("UnitofMeasure", Range("A1:BZ1"))) = "Direct Unit of Measure"
            Cells(1, xMatch("Amount", Range("A1:BZ1"))) = "Direct Rate"
            Cells(1, xMatch("Status", Range("A1:BZ1"))) = "Meter Status"
            Cells(1, xMatch("RevenueSKU", Range("A1:BZ1"))) = "Revenue SKU"
            Cells(1, xMatch("DisclosureDate", Range("A1:BZ1"))) = "Public Disclosure Date"
        End If

        'text description fields below
        Names = Split("Service Name,Service Type,Resource Name,Region Name,Direct Unit of Measure", ",")
        NameChecks = Split("NewService,NewServiceType,NewMeterName,Region,NewUnitOfMeasure", ",") ' changes to the above order should be matched
        Dates = Split("CP Rate Start Date,Public Disclosure Date", ",")
        Changes = Split("Has Name Change,Has Revenue Sku Change,Has Price Decrease,Has Price Increase,Has DevTest Percent Change," & _
            "Has Meter Status Changed,Has Meter Status Changed,Has Graduated Rate Change,Has Incl Qty Decrease,Has Incl Qty Increase", ",")
        ModChecks = Split("NameChange,PriceChange,OtherModifies", ",")
        ChngChecks = Split("Service Name,Service Type,Resource Name,Direct Unit of Measure,Region Name,Revenue SKU,Direct Rate,CP Rate Start Date,Meter Status", ",")
        keycol = "Resource GUID"

    End If
    
    Wrk.Sheets("Working").Activate
    last = Range("A200000").End(xlUp).Row
    If last = 1 Then        'check to see if there is data in the Working tab, if not end the sub
        MsgBox ("'Working' tab not populated." & vbNewLine & "Please rerun with appropriate data.")
        Application.ScreenUpdating = True
        End
    End If
    x = Range("A1").End(xlToRight).Column   'grab last column
    
    On Error Resume Next
    If Validator.ModifyChecks Then  'for all modify checks, see if column already present; if not, add to the end of the sheet
        For i = 0 To UBound(ModChecks)
            If IsError(xMatch(ModChecks(i), Range("A1:CZ1"))) Then
                x = x + 1
                Cells(1, x) = ModChecks(i)
            End If
        Next
    End If
    
    x = Range("A1").End(xlToRight).Column   'grab last column
    
    If Validator.Changeset Then
        If IsError(xMatch("Changeset", Range("A1:CZ1"))) Then
            x = x + 1
            Cells(1, x) = "Changeset"
        End If
    End If
    Wrk.Sheets("Working").Activate
    
    Columns(ltr(xMatch("Macro Validation", Range("A1:CZ1")))).Delete    'delete old validation column
    On Error GoTo 0
    
    x = Range("A1").End(xlToRight).Column   'grab last column
    Cells(1, Range("A1").End(xlToRight).Column + 1) = "Macro Validation"   'place validation column at end
    x = x + 1
    
    For i = 1 To x  'add all columns to the collection
        wrkcol.Add Cells(1, i).Value, i
    Next
    
    Columns("A:CZ").NumberFormat = "General" 'correct formatting
    For i = 0 To UBound(Dates)  'adjust for Date formatting
        Columns(ltr(wrkcol(Dates(i)))).NumberFormat = "m/d/yyyy"
    Next
    
    Wrk.Sheets("Prod").Activate
    x = Range("BZ1").End(xlToLeft).Column
    For i = 1 To x
        prdcol.Add Cells(1, i).Value, i     'add prod columns to collection
    Next
    
    Columns("A:CZ").NumberFormat = "General" 'correct formatting
    
    If Valtype = "EA" Then
        For i = 0 To UBound(Dates)  'adjust for Date formatting
            Columns(ltr(prdcol(Dates(i)))).NumberFormat = "m/d/yyyy"
        Next
    End If
    
    If Validator.Changeset Then
        Chng.Sheets("Changeset").Activate
        If Range("A1") = "ID" Then  'grab column header row (either 1st or 2nd)
            y = 1
        ElseIf Range("A2") = "ID" Then
            y = 2
        Else        'throw error if data not in proper place
            MsgBox ("Changeset data not found based on 'ID' header in the first two rows of the first column." & vbNewLine & "Please run the Changeset query in cell 'A1'.")
        End If
        
        x = Range("A" & y).End(xlToRight).Column
        For i = 1 To x
            chngcol.Add Cells(y, i).Value, i        'add Changeset columns to collection
        Next
    End If
    
    If Validator.NameCheck Then
        Namebook.Sheets(shtname).Activate
        
        x = Range("A1").End(xlToRight).Column   'grab last column
        For i = 1 To x
            namecol.Add Cells(1, i).Value, i        'add name sheet columns to collection
        Next
    End If
    
    Wrk.Sheets("Working").Activate
    
End Sub

Sub UpdateScope(ByVal Valtype As String, Optional ByVal Version)

    On Error Resume Next
    If IsError(Sheets("Scope").Name) Then Sheets.Add(after:=Sheets(Sheets.Count)).Name = "Scope"  'if scope sheet doesn't exist, add it
    On Error GoTo 0
    
    Sheets("Scope").Activate
    If Range("A100").End(xlUp).Row = 1 Then 'if scope sheet is blank populate it
            Range("A1") = "Version"
            Range("A3") = "Events"
            Range("A6") = "Create New"
            Range("A7") = "DevTest"
            
        If Valtype = "EA" Then
            Range("A4") = "Total SKUs"
            Range("A8") = "Included Quantities"
            Range("A10") = "Modify Existing"
            Range("A11") = "Price Decreases"
            Range("A12") = "Price Increases"
            Range("A13") = "Mat Desc Changes"
            Range("A14") = "EA UoM Changes"
            Range("A15") = "EA Inc Quantity Changes"
            Range("A16") = "Discontinues"
            'Range("A17") = "No Changes"
            
        ElseIf Valtype = "Direct" Then
            Range("A4") = "Total Meters"
            Range("A8") = "Graduated Rate"
            Range("A9") = "Included Quantities"
            Range("A11") = "Modify Existing"
            Range("A12") = "Price Decreases"
            Range("A13") = "Price Increases"
            Range("A14") = "Name Changes"
            Range("A15") = "DevTest Percent Changes"
            Range("A16") = "Graduated Rate Changes"
            Range("A17") = "Included Quantity Decreases"
            Range("A18") = "Included Quantity Increases"
            Range("A19") = "Meter Status Changes"
            'Range("A20") = "No Changes"
        End If
    End If
    
    Sheets("Scope").Activate
    x = Range("BZ1").End(xlToLeft).Column + 1  'grab version column
    
    If Valtype = "EA" Then
        val1 = DateAdd("D", -1, Left(Version, 2) & "/01/" & Right(Version, 2))
        Cells(1, x) = x - 1
        Cells(3, x) = xCountIf(Sheets("Original").Columns(ltr(xMatch("Work Item Type", Sheets("Original").Range("A2:BZ2")))), "Event")
        Cells(4, x) = WorksheetFunction.CountA(Sheets("Working").Range("A2:A800000"))
        Cells(6, x) = xCountIf(Sheets("Working").Columns(wrkcol("Add/Update")), "Add")
        Cells(7, x) = CountIf2("Add/Update", "Add", "Sku Sub Type", "DevTest")
        Cells(8, x) = CountIf2("Add/Update", "Add", "EA Included Quantity Units", ">0")
        Cells(10, x) = xCountIf(Sheets("Working").Columns(wrkcol("Add/Update")), "Update")
        Cells(11, x) = CountIf2("Add/Update", "Update", "Has Ea Rate Decreased", "Yes")
        Cells(12, x) = CountIf2("Add/Update", "Update", "Has Ea Rate Increased", "Yes")
        Cells(13, x) = CountIf2("Add/Update", "Update", "Has Material Description Changed", "Yes")
        Cells(14, x) = CountIf2("Add/Update", "Update", "Has Ea Uom Changed", "Yes")
        Cells(15, x) = CountIf2("Add/Update", "Update", "Has Included Quantity Changed", "Yes")
        Cells(16, x) = CountIf2("Add/Update", "Update", "SAP Discontinue Date", val1)
    ElseIf Valtype = "Direct" Then
        str = ltr(wrkcol("Change Type"))
        Cells(1, x) = x - 1
        Cells(3, x) = xCountIf(Sheets("Original").Columns(ltr(xMatch("Work Item Type", Sheets("Original").Range("A2:CZ2")))), "Event")
        Cells(4, x) = WorksheetFunction.CountA(Sheets("Working").Range("A2:A800000"))
        Cells(6, x) = xCountIf(Sheets("Working").Columns(str), "Create New")
        Cells(7, x) = CountIf2("Change Type", "Create New", "Dev/Test Eligible", "Yes")
        Cells(8, x) = CountIf2("Change Type", "Create New", "Has Graduated Rate", "Yes")
        Cells(9, x) = CountIf2("Change Type", "Create New", "Included Quantity Units", ">0")
        Cells(11, x) = xCountIf(Sheets("Working").Columns(str), "Modify Existing")
        Cells(12, x) = CountIf2("Change Type", "Modify Existing", "Has Price Decrease", "Yes")
        Cells(13, x) = CountIf2("Change Type", "Modify Existing", "Has Price Increase", "Yes")
        Cells(14, x) = CountIf2("Change Type", "Modify Existing", "Has Name Change", "Yes")
        Cells(15, x) = CountIf2("Change Type", "Modify Existing", "Has DevTest Percent Change", "Yes")
        Cells(16, x) = CountIf2("Change Type", "Modify Existing", "Has Graduated Rate Change", "Yes")
        Cells(17, x) = CountIf2("Change Type", "Modify Existing", "Has Incl Qty Decrease", "Yes")
        Cells(18, x) = CountIf2("Change Type", "Modify Existing", "Has Incl Qty Increase", "Yes")
        Cells(19, x) = CountIf2("Change Type", "Modify Existing", "Has Meter Status Changed", "Yes")
    End If
    
    Columns("A").AutoFit
    x = 0
End Sub

Sub CreateWorkingFile(ByVal Valtype As String, ByVal SovCloud As String, ByVal Version As String)
fpath = "\\microsoft.sharepoint.com@SSL\DavWWWRoot\teams\AzureReleaseOps\Release Implementation Files\"
x = 0
    
    If Valtype = "EA" Then
        fpath = fpath & "SAP\"
        If SovCloud = "Global" Then
            fname = Left(Version, 2) & "-20" & Right(Version, 2) & " PL"
        ElseIf SovCloud = "China" Then
            fname = Left(Version, 2) & "-20" & Right(Version, 2) & " China PL"
        End If
    ElseIf Valtype = "Direct" Then
        fpath = fpath & "Cayman\"
        fname = Validator.ReleaseBox & "." & Version
    End If
    
    fpath = fpath & fname & "\"
    fname = fname & " Working File.xlsx"
    
    For Each WB In Workbooks
        If WB.Name = fname Then x = 1
    Next

    If x = 1 Then
        If MsgBox("Working file for selected release open." & vbNewLine & "Would you like to use this file?", vbYesNo) = vbYes Then
            Exit Sub
        Else
            Set Wrk = Workbooks.Add
            Call InitialSetup(Valtype)
        End If
    Else
        If dir$(fpath, vbDirectory) = "" Then     ' if filepath for release doesn't exist, create it
            MkDir (fpath)
        ElseIf dir$(fpath & fname, vbDirectory) = "" Then 'if working file for release doesn't exist in release folder, create it
            Set Wrk = Workbooks.Add
            Call InitialSetup(Valtype)
            Wrk.SaveAs (fpath & fname)
        Else       'if working file exists, ask to open it
            If MsgBox("Working file exists in release folder." & vbNewLine & "Would you like to use this file?", vbYesNo) = vbNo Then
                Set Wrk = Workbooks.Add
                Call InitialSetup(Valtype)
            Else
                Workbooks.Open (fpath & fname)
            End If
        End If
    End If
End Sub

Sub InitialSetup(ByVal Valtype As String)
Issues = Split("Event ID,Event Name,Service,Change Owner,Meter/SKUs Impacted,Issue,Relevant Fields,Issue Cause,ASOMS Warning?,CR Approver,Severity/Churn,Immediate Workaround,Prevention Plan,Bug Number,Release,Resolution Type,Stage Caught,Release Month", ",")
InputReq = Split("Event ID,Meter/SKUs Impacted,Issue,Relevant Fields,Issue Cause,ASOMS Warning?,Severity/Churn,Immediate Workaround,Bug Number,Resolution Type,Stage Caught,Release Month", ",")

    'set column headers
           
    Wrk.Activate
    Sheets(1).Name = "Original"
    Sheets.Add after:=Sheets(1), Count:=4
    Sheets(2).Name = "Working"
    Sheets(3).Name = "Prod"
    Sheets(4).Name = "Scope"
    Sheets(5).Name = "Issues"
    
    Sheets(5).Activate
    
    For i = 0 To UBound(Issues())   'insert column headers
        Cells(1, i + 1) = Issues(i)
        For j = 0 To UBound(InputReq)
            If Issues(i) = InputReq(j) Then Cells(1, i + 1).Interior.Color = 6740479    'highlight fields which need to be filled by user
        Next
    Next
    
    Columns("A:K").AutoFit

    Sheets(1).Activate
    
End Sub

Sub GetWorkingWB()
x = 0

    If Wrk Is Nothing Then
        For Each WB In Workbooks
            If Right(WB.Name, 17) = "Working File.xlsx" Then
                Set Wrk = WB
                x = x + 1
            End If
        Next
     
        If x = 0 Then
            MsgBox ("Working file not found based on filename, looking at '... Working File'.")
            Application.ScreenUpdating = True
            End
        ElseIf x = 2 Then
            MsgBox ("Multiple working files found based on filename, looking at '... Working File'.")
            Application.ScreenUpdating = True
            End
        End If
    Else
    End If
    
    Wrk.Activate
    
    x = 0

End Sub

Sub GetChangesetWB()   'Look for and assign the Changeset file to variable- redundant with the above, but it'll have to do for now...
x = 0

    If Chng Is Nothing Then
        For Each WB In Workbooks
            If Right(WB.Name, 14) = "Changeset.xlsx" Then
                Set Chng = WB
                x = x + 1
            End If
        Next
     
        If x = 0 Then
            MsgBox ("Changeset file not found based on filename, looking at '... Changeset'." & vbNewLine & "Deselect Changeset if not validating.")
            Application.ScreenUpdating = True
            End
        ElseIf x = 2 Then
            MsgBox ("Multiple changeset files found based on filename, looking at '... Changeset'." & vbNewLine & "Please close additional changesets.")
            Application.ScreenUpdating = True
            End
        End If
    Else
    End If
    
    x = 0
    For Each WS In Chng.Worksheets  'look for Changeset sheet
        If WS.Name = "Changeset" Then
            x = x + 1
        End If
    Next
    If x = 0 Then MsgBox ("Changeset sheet not found in the Changeset file." & vbNewLine & "Please name sheet with Changeset data 'Changeset'.")
    
    Wrk.Activate
    x = 0

End Sub

Sub GetNameWB(ByVal Valtype As String, Optional ByVal SovCloud As String)
x = 0
fpath = "\\srinib2\Public\AzureOffsite\Services\"

    If Valtype = "EA" Then
        fname = "EANewMeterNames.xlsx"
        shtname = "EA Names"
    ElseIf Valtype = "Direct" Then
        fname = "NewMeterNames.xlsx"
        If SovCloud = "China" Then
            shtname = "ChinaMeterNames"
        Else
            shtname = "MeterNewNames"
        End If
    End If
        
    
    For Each WB In Workbooks
        If WB.Name = fname Then x = x + 1
    Next

    If x = 1 Then
        Set Namebook = Workbooks(fname)
    ElseIf dir$(fpath & fname, vbDirectory) <> "" Then 'if EA Names file exists, open and set.
        Workbooks.Open (fpath & fname)
        Set Namebook = Workbooks(fname)
    Else
        MsgBox ("Name file inaccessible. Please check access to " & fpath & fname & ".")
        Exit Sub
    End If

    x = 0
    For Each WS In Namebook.Worksheets  'look for Name sheet
        If WS.Name = shtname Then
            x = x + 1
        End If
    Next
    
    If x = 0 Then   'if name sheet doesn't exist, exit
        MsgBox ("EA Names sheet not found in " & fname & vbNewLine & ". Please check for name change.")
        Exit Sub
    End If
                            
    Wrk.Activate
    x = 0
    
End Sub

Sub UpdateWorking(ByVal Valtype As String)

    Wrk.Sheets("Original").Activate
    If Range("A2") = "" Then
        If MsgBox("'Original' tab empty, data from ASOMS needed to update working tab." & vbNewLine & "Would you like to proceed with data currently on Working tab?", vbYesNo) = vbYes Then
            Exit Sub
        Else
            Application.ScreenUpdating = True
            End
        End If
    End If
    ActiveSheet.Cells.WrapText = False
    'Range("A2:BZ2").AutoFilter
    
    last = Range("A200000").End(xlUp).Row  'get last row
    
    For Each WS In Wrk.Worksheets   'look for working sheet in working file
        If WS.Name = "Working" Then x = x + 1
    Next
    
    If x = 0 Then   'if Working tab doesn't exist, create one, else clear contents of old Prod tab
        Wrk.Sheets.Add(after:=Sheets("Original")).Name = "Working"
    Else
        Wrk.Sheets("Working").UsedRange.ClearContents
    End If
    
    Wrk.Sheets("Original").Range("A2:CZ" & last).Copy
    Wrk.Sheets("Working").Range("A1").PasteSpecial (xlPasteValues)  'paste values to Working sheet
    Application.CutCopyMode = False     ' clear clipboard
    
    Wrk.Sheets("Working").Activate
    endcol = Range("A1").End(xlToRight).Column    'assign last column number for current release
    Range("A1:" & ltr(endcol) & last).AutoFilter Field:=xMatch("Work Item Type", Range("A1:CZ1")), Operator:=xlFilterValues, Criteria1:="<>" & WrkItm
    Application.DisplayAlerts = False
    Range("A1:" & ltr(endcol) & last).Offset(1, 0).SpecialCells(xlCellTypeVisible).Delete   ' delete all line items that are not SKUs/Meters depending on validation
    Application.DisplayAlerts = True
    Range("A1:" & ltr(endcol) & last).AutoFilter
    
    If Valtype = "Direct" Then
        a = xMatch("Title 2", Range("A1:CZ1"))   ' change eventid header and remove "For Event - "
        Cells(1, a) = "EventId"
        For i = 2 To last
            Cells(i, a) = Replace(Cells(i, a), "For Event - ", "")
        Next
    ElseIf Valtype = "EA" Then      ' for EA add a new column to indicate Adds vs Updates
        last = Range("a1").End(xlDown).Row
        Columns(3).Insert
        Range("C1") = "Add/Update"
        a = xMatch("Part Number", Range("A1:CZ1"))  'grab column headers
        b = xMatch("Title", Range("A1:CZ1"))
        Cells(1, b) = "EventId"
        For i = 2 To last
            If Cells(i, a) = "" Then    'if Part Number blank = Add
                Cells(i, 3) = "Add"
            Else
                Cells(i, 3) = "Update"
            End If
            Cells(i, b) = Replace(Cells(i, b), "For Event - ", "")
        Next
        
    End If
    
End Sub

Sub GetReduceProd(ByVal Valtype As String, Optional ByVal SovCloud As String)
Dim sht As String
x = 0
    
    If Application.ProtectedViewWindows.Count > 0 Then
        For Each PV In Application.ProtectedViewWindows     'remove protected view
            PV.Edit
        Next
    End If
    
    Wrk.Activate
    For Each WS In Wrk.Worksheets   'look for Prod sheet in working file
        If WS.Name = "Prod" Then x = x + 1
    Next
    
    If x = 0 Then   'if Prod tab doesn't exist, create one, else clear contents of old Prod tab
        Wrk.Sheets.Add(after:=Sheets(2)).Name = "Prod"
    ElseIf Valtype = "Direct" Then      ' for
        Wrk.Sheets("Prod").UsedRange.ClearContents
    End If
    
    x = 0
    
    If Valtype = "Direct" Then      'for Direct, look for Cayman report
        For Each WB In Workbooks
            If Left(WB.Name, 6) = "Report" Then
                Set Prd = WB
                x = x + 1
            End If
        Next
        
        If x = 0 Then
            MsgBox ("Production data file not found based on filename. Looking for 'Report'.")
            Application.ScreenUpdating = True
            End
        End If
    
        If SovCloud = "Global" Then
            sht = "MS-AZR-0003P"
        ElseIf SovCloud = "Germany" Then
            sht = "MS-AZR-DE-0003P"
        ElseIf SovCloud = "China" Then
            sht = "MS-MC-AZR-0033P"
        End If
    
        Prd.Sheets(sht).Activate
        If Range("A1") <> "ResourceId" Then Rows(1).Delete
        'Sheets(sht).ListObjects("Table1").Unlist
        last = Range("A1").End(xlDown).Row
        Rows(last + 1 & ":200000").Delete       'remove excess table below data
        x = WorksheetFunction.Match("StartDate", Rows("1"), 0)
        
        
        For i = 1 To last
            Cells(i, x) = Format(Cells(i, x), "m/d/yyyy")   'have to change every date so that Excel will recognize it as such
            
        Next
        
        For i = last To 2 Step -1   'remove every value which is not the most recent instance of a meter
            If Cells(i, WorksheetFunction.Match("MinValue", Rows("1"), 0)) <> 0 Or Cells(i, x) <> WorksheetFunction.MaxIfs(Columns(ltr(x)), Columns("A"), Cells(i, 1)) Then
                Rows(i).Delete
            End If
        Next
        
        x = 0
        
        Prd.Sheets(sht).UsedRange.Copy
        Wrk.Sheets("Prod").Range("A1").PasteSpecial (xlPasteValues)
        Application.CutCopyMode = False     'clear clipboard
                
    ElseIf Valtype = "EA" Then
        If Sheets("Prod").Cells(2, 1) = "" Then 'check if the Prod tab is populated
            MsgBox ("Prod data from ASOMS does not appear to be populated based on cell A2." & vbNewLine & "Please download data from ASOMS to the 'Prod' Sheet.")
            If Validator.ModifyChecks Then End  'if modify checks are supposed to run, then end the procedure
        Else
            Sheets("Prod").Activate
            last = Range("A200000").End(xlUp).Row
            Range("A2:CZ" & last).Copy
            Range("A" & last + 2).PasteSpecial (xlPasteValues)
            Rows("1:" & last + 1).Delete
        End If
    End If
    
    x = 0
End Sub

Sub GetProdDevTest(ByVal Valtype As String, ByVal SovCloud As String)
x = 0
    
    If Valtype = "EA" Or SovCloud = "Germany" Then
    
    Else
        If Application.ProtectedViewWindows.Count > 0 Then
            For Each PV In Application.ProtectedViewWindows     'remove protected view
                PV.Edit
            Next
        End If
        
        For Each WB In Workbooks
            If Left(WB.Name, 6) = "Report" Then
                Set Prd = WB
                x = x + 1
            End If
        Next
        If x = 0 Then
            MsgBox ("Production data file not found based on filename. Looking for 'Report'.")
            Application.ScreenUpdating = True
            End
        End If
        
        If SovCloud = "Global" Then
            sht = "MS-AZR-0060P"
        ElseIf SovCloud = "China" Then
            sht = "MS-MC-AZR-0060P"
        End If
        
        x = 0
        
        For Each WS In Prd.Worksheets
            If WS.Name = sht Then x = x + 1
        Next
        If x = 0 Then   'if no DevTest sheet found, prompt to rerun to with proper selections
            MsgBox ("Production DevTest sheet not found based on sheet name = " & sht & vbNewLine & "Please rerun with DevTest data or 'Prod DevTest' not selected.")
            Application.ScreenUpdating = True
            End
        End If
        
        x = 0
        Wrk.Activate
        
        For Each WS In Wrk.Worksheets   'look for Prod sheet in working file
            If WS.Name = "Prod" Then x = x + 1
        Next
        If x = 0 Then   'if no prod sheet, prompt to rerun to get sheet populated
            MsgBox ("Prod sheet not found. Please rerun with Prod Data selected.")
            Application.ScreenUpdating = True
            End
        End If
        
        ElseIf Wrk.Sheets("Prod").Range("A1") = "" Then 'if no prod data, prompt to rerun to get sheet populated
            MsgBox ("Prod sheet empty. Please rerun with Prod Data selected.")
            Application.ScreenUpdating = True
            End
        End If
        
        last = Wrk.Sheets("Prod").Range("A500000").End(xlUp).Row    'get last prod row
        val1 = Range("A1").End(xlToRight).Column + 1   'get last column + 1 of Prod sheet for devtest column
        val2 = WorksheetFunction.Match("Amount", Prd.Sheets(sht).Range("A2:BZ2"))   'get Amount column from Prod Data
        Cells(1, val1) = "DevTest Discount Rate"
        
        For i = 2 To last
            'Cells( i , x) = Prd.Sheets(sht).
        Next
    
    End If
End Sub
