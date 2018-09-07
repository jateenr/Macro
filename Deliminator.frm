VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Deliminator 
   Caption         =   "Deliminator"
   ClientHeight    =   2880
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6045
   OleObjectBlob   =   "Deliminator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Deliminator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CopyButton_Click()

    Application.SendKeys ("^c~")    'copies contents of the Output box to the clipboard
    InputBox "clipboard", , Deliminator.Output.text
    
End Sub

Private Sub DelimRun_Click()
Dim Ray() As Variant
Dim i As Integer, x As Integer, last As Integer
Dim Delim As String, Out As String
Dim cel As Range, Selrange As Range

Application.ScreenUpdating = False '++speed
Delim = Deliminator.DelimValue      ' grab user input for delimiter

    If Selection.Rows.Count < 2 Then    ' if only one cell selected, throw exception
        Deliminator.Output = "Please select more than one cell"
        Deliminator.Output.ForeColor = 192   ' output in red
        Exit Sub
        Else: Deliminator.Output.ForeColor = 0   ' ensure output set to black
    End If
    
    Set Selrange = Application.Selection   'sets selection variable and size
    ReDim Ray(Selrange.Cells.Count)
    
    For Each cel In Selrange.Cells
        i = i + 1
        Ray(i) = cel.Value      'adds each cell to an array
    Next cel
    
    Out = Join(Ray(), Delim)   ' join array into one string using the user-specified delimiter
    Deliminator.Output = Mid(Out, 2, Len(Out)) ' adds to the output box, removing the last value delimiter

End Sub


