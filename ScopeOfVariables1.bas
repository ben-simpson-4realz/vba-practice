Attribute VB_Name = "ScopeOfVariables1"
Public q As Integer
Dim z As Integer
Sub Sub1()
    Dim x As Integer
    Static y As Integer
    
    x = x + 100
    y = y + 100
    z = z + 100
    q = q + 100
    MsgBox ("x in sub 1: " & x) 'Dies when sub1 ends
    MsgBox ("y in sub 1: " & y) 'Lives after Sub1 ends
    MsgBox ("z in sub 1: " & z) 'Lives after Sub1 ends
    MsgBox ("q in sub 1: " & q) 'Lives after Sub1 ends and across modules
    Call Sub2
    Call GlobalVariable
End Sub
Sub Sub2()

    MsgBox ("x in sub 2: " & x) 'no value
    MsgBox ("y in sub 2: " & y) 'no value
    MsgBox ("z in sub 2: " & z) 'Lives after Sub1 ends
    MsgBox ("q in sub 2: " & q) 'Lives after Sub1 ends and across modules
End Sub
