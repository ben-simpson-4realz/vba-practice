Attribute VB_Name = "Classes"
Sub TestClasses()
    Dim bbteam As basketBallTeam
    Set bbteam = New basketBallTeam 'instantiate object
    
    bbteam.Name = "Lakers" 'uses Let
    MsgBox (bbteam.Name) 'Uses Get
    MsgBox (Application.Name)
End Sub

Sub TestObjectBrowser()
    'MsgBox (Application.Sheets.Count)
    Dim ws As Worksheet
    Set ws = Sheets("sheet4")
    MsgBox (ws.Name)
End Sub
