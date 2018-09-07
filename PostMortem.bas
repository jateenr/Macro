Attribute VB_Name = "PostMortem"
Option Explicit

Private last As Double, i As Double, j As Double, x As Double, relcol As Double
Private str As String
Private fpath As String, fname As String, release As String
Private Releases() As String, RelType(2) As String
Private Wrk As Workbook, PM As Workbook, WB As Workbook
Private PV As ProtectedViewWindow

Function ltr(ByVal colnum As Integer)
    ltr = Split(Cells(1, colnum).Address, "$")(1)
End Function

Sub CreateGraphics(ByVal ReleaseDate As String)
    

End Sub

Sub FetchPostMortemData()
    Application.ScreenUpdating = False
    
    Call PostMortemSetup
    Call GetReleaseFolders("Cayman")
    Call GetReleaseFolders("SAP")
    
    Set Wrk = Nothing
    Set PM = Nothing
    Set WB = Nothing
    Application.ScreenUpdating = True
    
End Sub

Sub PostMortemSetup()
x = 0
fpath = "\\microsoft.sharepoint.com@SSL\DavWWWRoot\teams\AzureReleaseOps\Release Implementation Files\"

    For Each WB In Workbooks
        If Right(WB.Name, 11) = "Mortem.xlsx" Then
            x = x + 1
            Set PM = WB
        End If
    Next
    
    If x = 0 Then
        MsgBox ("Post Mortem file not found based on filename." & vbNewLine & "Looking for ending with 'Mortem.xlsx'.")
        Set WB = Nothing
        Application.ScreenUpdating = True
        End
    ElseIf x > 1 Then
        MsgBox ("Multiple Post Mortem files open based on filename." & vbNewLine & "Please close additional file.")
        Set WB = Nothing
        Application.ScreenUpdating = True
        End
    End If
    
    PM.Sheets(1).Activate
    For i = 1 To 50 'get column number for Release
        If Cells(1, i) = "Release" Then
            relcol = i
            Exit For
        End If
    Next
    
    last = Cells(100000, relcol).End(xlUp).Row   'grab last row
    
    Range(ltr(relcol) & "2:" & ltr(relcol) & last).Copy Destination:=Range("A20000")  'copy all values from the release column, then remove duplicates
    Range("A20000:A" & 20000 + last).RemoveDuplicates
    
    ReDim Releases(Range("A50000").End(xlUp).Row - 19999) As String ' resize Releases array for unique Release values
    
    For i = 0 To last - 1
        Releases(i) = Cells(i + 20000, 1)   'add unique releases to the Releases array
    Next
    
    Range("A20000:A50000").Clear 'delete the duplicates from the sheet
    Range("A1").Activate
    
End Sub

Sub GetReleaseFolders(ByVal RelType As String)

    release = dir(fpath & RelType & "\*", vbDirectory)
        
    Do While Len(release) > 0
        x = 0
        
        If Len(release) > 3 And release <> "Manual SAP Calls" Then    'remove bad values like . or .. from processing
        
            For i = 0 To UBound(Releases())
                If Releases(i) = release Then
                    x = 1
                End If
            Next
            
            If x = 0 Then Call GetIssues(RelType, release)
            
        End If
        
        release = dir
    Loop
    
    
End Sub

Sub GetIssues(ByVal RelType As String, ByVal release As String)

    fname = fpath & RelType & "\" & release & "\" & release & " Working File.xlsx"
    Application.DisplayAlerts = False   'keep clipboard and data connection popups from showing
    Set WB = Workbooks.Open(fname, False)

    WB.Sheets("Issues").Range("A2:L30").Copy        'copy issues over to Post Mortem sheet
    PM.Sheets(1).Range("A" & last + 1).PasteSpecial (xlPasteValues)
    WB.Close
    Application.DisplayAlerts = True
    
    PM.Sheets(1).Activate
    last = Range("A10000").End(xlUp).Row
    
    For i = last To 2 Step -1
        If Cells(i, relcol) = "" Then
            Cells(i, relcol) = release
        Else
            Exit For
        End If
    Next
End Sub

Sub GetEvents()

End Sub

Sub GetServiceNames()   'garbage code to get the job done...
Set WB = Workbooks("Book1.xlsx")
WB.Sheets(2).Activate

    last = Range("A100000").End(xlUp).Row
    Cells(1, 6) = "Event Service"
    
    For i = 2 To last
        If Left(Cells(i, 2), 9) = "For Event" <> "Event" Then Cells(i, 3) = Replace(Cells(i, 3), "For Event - ", "")
        If Cells(i, 4) = "Cloud Services" Then Cells(i, 4) = "Virtual Machines"
        If Cells(i, 4) = "Data Management" Then Cells(i, 4) = "Storage"
    Next
    
    For i = 2 To last
        If Cells(i, 2) = "Event" Then
            If WorksheetFunction.CountIf(Range("C1:C" & last), Cells(i, 1)) = WorksheetFunction.CountIfs(Range("C1:C" & last), Cells(i, 1), Range("D1:D" & last), Cells(i + 1, 4)) Then
                Cells(i, 6) = Cells(i + 1, 4)
            Else: Cells(i, 6) = "Multiple"
            End If
        Else
        
        End If
    Next

End Sub
