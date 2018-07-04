VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_EnterData 
   Caption         =   "Please Enter Data"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8850
   OleObjectBlob   =   "frm_EnterData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_EnterData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_CloseForm_Click()
    Unload frm_EnterData 'closes form
End Sub

Private Sub btn_EnterDataClick_Click()
    Dim xrow As Long
    
    If txt_Name.Value = "" Then
        MsgBox ("You must enter a name.")
        Exit Sub
    End If
    
    If combo_Feeling.Value = "" Or combo_Feeling.Value = "Select" Then
        MsgBox ("You must answer how you feel.")
        Exit Sub
    End If
    
    Sheets("User Form").Select
    'Range("A:z50000").ClearContents
    xrow = Range("lastrow").Row 'find last row to use, (named range under formulas)
    
    'Move data from the form to the worksheet
    Cells(xrow, 1).Value = frm_EnterData.txt_Name.Value
    Cells(xrow, 2).Value = combo_Feeling.Value
End Sub

Private Sub UserForm_Initialize()
    combo_Feeling.AddItem "I feel good."
    combo_Feeling.AddItem "I feel bad."
    combo_Feeling.AddItem "Select"
    combo_Feeling.Value = "Select"
End Sub
