VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Emailer 
   Caption         =   "Emailer"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8445
   OleObjectBlob   =   "Emailer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Emailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CreateButton_Click()
Dim Outlook As Object, Mail As Object
Dim WB As Workbook, Wrk As Workbook
Dim body As String, signa As String, title As String, col1 As String
Dim rw, col2 As Integer
    
    Application.ScreenUpdating = False
    
    If Not IsNumeric(EventBox.text) Then
        MsgBox ("Non-number submitted for Event Id." & vbNewLine & "Only one event may be sent at a time.")
        Exit Sub    'if a non-number is submiutted under the event Id, throws an error
    End If
    
    For Each WB In Workbooks
        If Right(WB.Name, 17) = "Working File.xlsx" Then
            Set Wrk = WB        ' if working book is open, grabs it for reference
        End If
    Next
    
    On Error Resume Next
    Wrk.Sheets(1).Activate  'assumes user has the ASOMS data on Sheets(1)
    col1 = Split(Cells(1, WorksheetFunction.Match("ID", Rows(2), 0)).Address, "$")(1)  'grabs the column number then letter for the Title column
    col2 = WorksheetFunction.Match("Title" & "*", Rows(2), 0)
    rw = WorksheetFunction.Match(Int(EventBox.text), Columns(col1), 0)  'grabs event title from working file
    title = Cells(rw, col2).Value
    On Error GoTo 0
    
    Set Outlook = CreateObject("Outlook.Application")
    Set Mail = Outlook.CreateItem(0)     'creates new email
    body = "<a href='http://vstfpg07:8080/tfs/ASOMS/Offer/_workitems?_a=edit&id=" & EventBox.text & "'>" & EventBox.text & "</a>"
    
    With Mail
        .Display    'adds default signature
    End With
    signa = Mail.body
        
    With Mail
        .To = ToBox.text    'sets the contents of the email
        .cc = CcBox.text
        .Subject = EventBox.text & " - " & title
        .HTMLbody = "Hi, <p><p>" & body & " - " & title & "<p><p> Thank you,<p>" & signa
    End With
    
    Set Outlook = Nothing
    Set Mail = Nothing
    Unload Emailer
    Application.ScreenUpdating = True
End Sub
