Attribute VB_Name = "Loops"
Option Explicit
Sub ForEachLoopWorksheets()
Dim ws As Worksheet

    For Each ws In Worksheets
        ws.Select
        If ws.Name = "Loops" Then
            MsgBox (ws.Name)
        End If
    Next ws
End Sub
Sub ForNextLoopAddSheets()
    Dim numberOfSheets As Integer, counter As Integer
    numberOfSheets = Application.InputBox("How many worksheeets do you want to add?", "Add worksheets", , , , , , 1)
    
    If numberOfSheets = False Then
        Exit Sub
    Else
        For counter = 1 To numberOfSheets
            Worksheets.Add
        Next counter
    End If
    
End Sub
Sub TestForNext()
    Dim i As Long
    Sheets("For Next Loops").Select
    Cells.ClearContents
    
    For i = 1 To 20 Step 2
    Cells(i, 1).Value = i
    Next
End Sub
Sub DeleteBlankRows()
    Dim lastrow As Long, xrow As Long
    xrow = 1
    
    lastrow = Range("A95000").End(xlUp).Row
    
    Do Until xrow = lastrow
        If Cells(xrow, 1).Value = "" Then
            Cells(xrow, 1).Select
            Selection.EntireRow.Delete
            
            xrow = xrow - 1
            lastrow = lastrow - 1
        End If
        xrow = xrow + 1
    Loop
End Sub

Sub DoUntilBlankCell()
    Dim xrow As Long, xcol As Long, lastrow As Long, lastCol As Long
    xrow = 1
    xcol = 1
    
    Do Until Cells(xrow, xcol).Value = ""
        'Cells(xrow, xcol).Select
        xcol = xcol + 1
    Loop
    lastCol = xcol - 1
End Sub
