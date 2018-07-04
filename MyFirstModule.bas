Attribute VB_Name = "MyFirstModule"
Sub MySecondMacro()

'runs my first macro
Call MyFirstMacro("Meow", "Neko", 99999)
End Sub
Sub MyFirstMacro(catNoise As String, catName As String, numMeows As Long)
    'This is my first macro
    MsgBox ("My cat's name is " & catName & ", she goes " + catNoise & " " & numMeows & " times.")
End Sub

Sub Vars()
    Dim myFirstVariable As Integer, x As Integer, y As Integer
    
    myFirstVariable = 8
    x = 80
    y = 6000
    
End Sub
