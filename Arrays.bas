Attribute VB_Name = "Arrays"
Option Explicit


Sub Populate2DArrayFromExcel()
    Dim ExchangeRates(3, 2) As Variant, xrow As Long, xcol As Long, _
    rowIndex As Long, colIndex As Long, i As Long, j As Long
    rowIndex = 0
    colIndex = 0
    xrow = 2
    xcol = 1
    
    'outer loop down over rows
    Do Until Cells(xrow, xcol).Value = ""
        'inner loop across columns
        Do Until Cells(xrow, xcol).Value = ""
            ExchangeRates(rowIndex, colIndex) = Cells(xrow, xcol).Value
            colIndex = colIndex + 1 'increa
            xcol = xcol + 1
        Loop
        xcol = 3 'reset after done with row
        colIndex = 0 'reset 2nd dimension index in array
        xrow = xrow + 1
        rowIndex = rowIndex + 1
    Loop
    
End Sub



Function ConvertToUSD(foreignCurrencySymbol As String, amount As Double) As Double
    Dim ExchangeRates(3, 2) As Variant, i As Integer
    
    ExchangeRates(0, 0) = "Canada"
    ExchangeRates(0, 1) = "CAD"
    ExchangeRates(0, 2) = 1.05
    
    ExchangeRates(1, 0) = "Euro Zone"
    ExchangeRates(1, 1) = "EUR"
    ExchangeRates(1, 2) = 1.2
    
    ExchangeRates(2, 0) = "Japan"
    ExchangeRates(2, 1) = "JPY"
    ExchangeRates(2, 2) = 0.012
    
    ExchangeRates(3, 0) = "Mexico"
    ExchangeRates(3, 1) = "Mxn"
    ExchangeRates(3, 2) = 0.07
    
    For i = 0 To UBound(ExchangeRates, 1)
        If foreignCurrencySymbol = ExchangeRates(i, 1) Then
            ConvertToUSD = amount * ExchangeRates(i, 2)
            
        End If
    Next i
    
End Function

Sub StaticArray()
    Dim names(2) As String
    names(1) = "HI"
    MsgBox (names(1))
End Sub

Sub StaticArrayPopulateAndLoop()
    Dim names(2) As String
    Dim i As Integer
    
    names(0) = "Bob"
    names(1) = "Marie"
    names(2) = "George"
    
    For i = 0 To UBound(names, 1)
        Cells(i + 1, 1).Value = names(i)
    Next i
    
    
End Sub


Sub populate1DArrayFromWorksheet()
    Dim months(11) As String
    Dim i As Integer
    Dim xrow As Long
    i = 0 'variable for the index of the array
    xrow = 2 'variable for the row # on worksheet
    
    Do Until Cells(xrow, 1).Value = ""
        months(i) = Cells(xrow, 1).Value 'populates array
        i = i + 1
        xrow = xrow + 1
    Loop
    
    For i = 0 To UBound(months, 1)
        If months(i) = MonthName(Month(Date)) Then
            MsgBox ("The current month is " & months(i))
        End If
        
    Next i
End Sub
