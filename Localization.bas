Attribute VB_Name = "Localization"
Option Explicit
Private fpath As String, fname As String
Private Wrk As Workbook, WB As Workbook
Private i As Integer, j As Integer, x As Integer, extnum As Integer, endcol As Integer

Function ltr(ByVal colnum As Integer)
    ltr = Split(Cells(1, colnum).Address, "$")(1)
End Function

Sub LocValidation(ByVal usrinput As String)

    Application.ScreenUpdating = False '++speed
    fpath = usrinput
    If Left(fpath, 4) = "http" Then fpath = Mid(fpath, InStr(1, fpath, ":") + 1, 500)   'remove htpp if needed
    If Right(fpath, 1) <> "\" Then fpath = fpath & "\"
    
    fname = dir(fpath & "CCSStrings.*.resx")
    i = 3   'header row
    extnum = 8 ' number of extracts desired
    Set Wrk = Workbooks.Add
    Wrk.Activate
    Cells(i, 2) = "Country Code"
    Cells(i, 3) = "Count"
    For j = 1 To extnum     'number of extracts, adding column headers
        Cells(i, 3 + j) = "Extract " & j
    Next
    
    Do While Len(fname) > 0
        i = i + 1
        Application.DisplayAlerts = False
        Workbooks.OpenXML Filename:=fpath & fname, loadoption:=xlXmlLoadImportToList    'import each resx file
        Application.DisplayAlerts = True
        Set WB = ActiveWorkbook
        endcol = WB.Sheets(1).Range("BZ55").End(xlToLeft).Column
        x = WorksheetFunction.CountA(WB.Sheets(1).Columns(ltr(endcol))) - 1 'grab total count of translated values
        Wrk.Sheets(1).Activate
        
        Cells(i, 2) = Replace(Replace(fname, "CCSStrings.", ""), ".resx", "")
        Cells(i, 3) = x
        
        For j = 0 To extnum - 1 ' add 7 values to check via internet translator
            Cells(i, 4 + j) = WB.Sheets(1).Cells(50 + j * 2, endcol)
        Next
        
        WB.Close savechanges:=False
        fname = dir
    Loop
    
    Wrk.Sheets(1).Activate
    Cells(2, 2) = WorksheetFunction.CountA(Range("B4:B100"))    'count files
    x = 0
    
    For i = 5 To Cells(2, 2) + 3
        If Cells(i, 3) <> Cells(4, 3) Then
            x = x + 1   'counts mismatched numbers
            Cells(i, 3).Interior.ColorIndex = 3 'highlights mismatched numbers
        End If
    Next
    
    Cells(2, 3) = "Diffs: " & x    ' number of off counts (should be zero)
    If x > 0 Then Cells(2, 3).Interior.ColorIndex = 3
    If Cells(2, 2) <> 23 Then Cells(2, 2).Interior.ColorIndex = 3
    
    Columns("D:P").ColumnWidth = 20
    
    Wrk.SaveAs (fpath & "Loc Validation.xlsx")
    
    fname = dir(fpath & "CCSStrings.resx")
    If fname = "" Then      'look for the previous resx file
        MsgBox ("CCSStrings.resx not found." & vbNewLine & "Please add the file before zipping.")
    End If
    
    Application.ScreenUpdating = True
    Set WB = Nothing
    Set Wrk = Nothing
    
End Sub

Sub ExcelLocFile()
Dim Outlook As Object, Mail As Object
Dim PV As ProtectedViewWindow
Dim heads() As String, signa As String
Dim col As Object
Set col = CreateObject("Scripting.Dictionary")
x = 0
    
    Application.ScreenUpdating = False '++speed
    
    For Each WB In Workbooks
        If Right(WB.Name, 20) = ExcelLoc.VersionBox & " Working File.xlsx" Then
            Set Wrk = WB
            x = x + 1
        End If
    Next
    
    If x = 0 Then
        MsgBox ("Working File not found based on filename ending in:" & vbNewLine & ExcelLoc.VersionBox & " Working File.xlsx")
        Application.ScreenUpdating = True
        Set WB = Nothing
        Set Wrk = Nothing
        Unload ExcelLoc
        Exit Sub
    End If
    
    x = 0
    fpath = "\\microsoft.sharepoint.com@SSL\DavWWWRoot\teams\AzureReleaseOps\Release Implementation Files\Cayman\" & ExcelLoc.ReleaseBox & "." & ExcelLoc.VersionBox & "\"
    
    If dir$(fpath, vbDirectory) = "" Then     ' if filepath for release doesn't exist, throw exception
        If MsgBox("Filepath not found. Would you like to create the file without saving?" & vbNewLine & fpath, vbYesNo) = vbNo Then
            Application.ScreenUpdating = True
            Set WB = Nothing
            Set Wrk = Nothing
            Unload ExcelLoc
            Exit Sub
        Else
            x = 1
        End If
    End If
    
    Wrk.Sheets("Working").Activate
    For i = 1 To Range("BZ1").End(xlToLeft).Column 'add column names to a collection
        col.Add Cells(1, i).Value, i
    Next
    
    heads = Split("Resource GUID,Service Name,Service Type,Resource Name,Region Name,Direct Unit of Measure,Change Type", ",")
    endcol = UBound(heads) + 1
    
    Set WB = Workbooks.Add
    Wrk.Sheets("Working").Activate
    
    For i = 0 To UBound(heads)
        Wrk.Sheets("Working").Columns(ltr(col(heads(i)))).Copy (WB.Sheets(1).Cells(1, i + 1))
    Next
    
    
    
    WB.Sheets(1).Activate
    Cells(1, endcol) = "New/Update"
    
    For i = 2 To Range("B100000").End(xlUp).Row
        If Cells(i, endcol) = "Create New" Then
            Cells(i, endcol) = "New"
        ElseIf Cells(i, endcol) = "Modify Existing" Then
            Cells(i, endcol) = "Update"
        End If
    Next
    
    Columns("A:K").ColumnWidth = 25
    
    Range("A1:" & ltr(endcol) & "1").Select
    Selection.Borders(xlEdgeBottom).Weight = xlMedium
    Selection.Font.Bold = True
    Selection.Interior.Color = 16764057
    Range("A1").Select

    If x > 0 Then
        Application.ScreenUpdating = True
        Set WB = Nothing
        Set Wrk = Nothing
        Unload ExcelLoc
        Exit Sub
    Else
        WB.SaveAs (fpath & ExcelLoc.ReleaseBox & "." & ExcelLoc.VersionBox & "_LOC_ExcelFile.xlsx")
        
        Do While dir$(fpath & ExcelLoc.ReleaseBox & "." & ExcelLoc.VersionBox & "_LOC_ExcelFile.xlsx", vbDirectory) = ""
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop
        
        Set Outlook = CreateObject("Outlook.Application")
        Set Mail = Outlook.CreateItem(0)     'creates new email
        
        With Mail
            .Display    'adds default signature
        End With
        signa = Mail.body
            
        
        With Mail
            .To = "azuresiteloc@microsoft.com; v-katega@microsoft.com; v-pabarr@microsoft.com"  'sets the contents of the email
            .cc = "AROIT@microsoft.com; qinche@microsoft.com"
            .Subject = ExcelLoc.ReleaseBox & "." & ExcelLoc.VersionBox & " Localization"
            .body = "Hi Lali," & vbNewLine & vbNewLine & "Attached is the Excel localization file for " & ExcelLoc.ReleaseBox & "." & ExcelLoc.VersionBox & _
                ". Please let me know if there are any issues." & vbNewLine & vbNewLine & "Thank you," & vbNewLine & signa
            .attachments.Add (fpath & ExcelLoc.ReleaseBox & "." & ExcelLoc.VersionBox & "_LOC_ExcelFile.xlsx")
        End With
        
        Set Outlook = Nothing
        Set Mail = Nothing
    End If
    
    Application.ScreenUpdating = True
    Set WB = Nothing
    Set Wrk = Nothing
    Unload ExcelLoc
    
End Sub


Sub upadterater()
Application.ScreenUpdating = True
End Sub
