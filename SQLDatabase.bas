Attribute VB_Name = "SQLDatabase"
Option Explicit
Sub SQLDatabase()
    'This code requires a reference to ADO (ActiveX Data Objects Library)
    'ADO lets you connect to any data source like SQL Server, Access, csv files etc..
    ' Data Source <--> OLEDB Provider <--> ADO Object
    
    Dim rs As ADODB.Recordset 'holds data
    Dim cnSQL As ADODB.Connection
    Dim sqlString As String ' holds the SQL text
    Dim sqlCommand As ADODB.Command, prm As Object
    Dim colOffset As Integer
    Dim qf As Object
    Dim numberOfRecordsAffected As Long
    
    colOffset = 0
    Sheets("database").Select
    Cells.ClearContents
    
    Set cnSQL = New ADODB.Connection 'instantiate connection
    'Open DB connection
    cnSQL.Open "Provider=SQLOLEDB.1; Integrated Security = SSPI; Initial Catalog = MyCompany; Data source = .\SQLEXPRESS"
    
    Set sqlCommand = New ADODB.Command
    sqlCommand.ActiveConnection = cnSQL
    sqlCommand.CommandType = adCmdStoredProc 'set command type to be stored procedure
    sqlCommand.CommandText = "getPeopleData"
    Set prm = sqlCommand.CreateParameter("id", adInteger, adParamInput)
    sqlCommand.Parameters.Append prm
    sqlCommand.Parameters("id").Value = 2
    
    'sqlString = "SELECT id, name, age FROM people WHERE name != 'mary'"
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic 'set this to readonly to prevent code from updating db in recordset loop
    rs.Open sqlCommand 'Execute stored procedure
    
    If rs.EOF Then
        MsgBox ("The recordset is empty. rs.eof = " & rs.EOF)
    Else
        'MsgBox ("# of records = " & rs.RecordCount)
        'put headers on worksheet
        For Each qf In rs.Fields
            Range("a1").Offset(0, colOffset).Value = qf.Name
            colOffset = colOffset + 1
        Next qf
        
        Do Until rs.EOF
            'If rs!ID.Value = 3 Then
            '    MsgBox ("Stopped at " & rs!ID.Value)
            'End If
            rs.Fields(2) = rs.Fields(2) + 10 'Change the value in rs and database
            rs.MoveNext 'moves to next row in recordset
        Loop
        rs.MoveFirst 'move back to first record in recordset
        
        'put data from recordset (rs) into worksheet
        ActiveSheet.Cells(2, 1).CopyFromRecordset rs
    End If
    
    rs.Close
    Set rs = Nothing
    
    'INSERT STATEMENT
    'sqlString = "INSERT INTO PEOPLE (name, age) VALUES ('Bitey',16)"
    'cnSQL.Execute sqlString, numberOfRecordsAffected, adCmdText
    'MsgBox ("# of records affectd = " & numberOfRecordsAffected)
    
    'UPDATE STATEMENT
    'sqlString = "UPDATE PEOPLE SET age = 18 WHERE name = 'Ben'"
    'cnSQL.Execute sqlString, numberOfRecordsAffected, adCmdText
    'MsgBox ("# of records affectd = " & numberOfRecordsAffected)
    
End Sub
