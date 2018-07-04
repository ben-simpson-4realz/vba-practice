VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "basketBallTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private teamName As String

Public Property Get Name() As String
    Name = teamName
End Property

Public Property Let Name(param_name As String)
    teamName = param_name
End Property
