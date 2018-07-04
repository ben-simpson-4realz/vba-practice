Attribute VB_Name = "Functions"
Option Explicit
Sub TestFunctions()
Dim x As Integer
Dim y As Double
x = Return1()
'MsgBox (x)

y = ConvertToCelcius(100)
MsgBox (y)
End Sub
Function ConvertToCelcius(TempFarenheit As Double) As Double
ConvertToCelcius = (TempFarenheit - 32) * 5 / 9
End Function
Function Return1() As Integer
    Return1 = 1
End Function
