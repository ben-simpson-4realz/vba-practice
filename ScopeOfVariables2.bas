Attribute VB_Name = "ScopeOfVariables2"
Sub GlobalVariable()
    MsgBox ("x in sub 1: " & x) 'No value
    MsgBox ("y in sub 1: " & y) 'No value
    MsgBox ("z in sub 1: " & z) 'No value
    MsgBox ("q in sub 1: " & q) 'Has a value declared public
End Sub
