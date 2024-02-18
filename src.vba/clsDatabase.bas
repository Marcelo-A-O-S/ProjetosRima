Public conn As ADODB.Connection
Public rs As ADODB.Recordset
Public strConn As String
Public strQuery As String
Private Sub Class_Initialize()
    Me.strConn = "Provider= Microsoft.ACE.OLEDB.12.0;Data Source =" & ThisWorkbook.FullName & ";Extended Properties='Excel 12.0 XML';"
End Sub
Function Init() As ADODB.Connection
    Set Me.conn = New ADODB.Connection
    Me.conn.Open Me.strConn
    Set Init = Me.conn
End Function
Function InsertQuery(ByVal param As String)
    Me.strQuery = param
End Function
Function ExecuteCommand() As ADODB.Recordset
    Set Me.rs = Me.conn.Execute(Me.strQuery)
    Set ExecuteCommand = Me.rs
End Function
Function CloseDb()
    If Me.conn <> Nothing Then
        Me.conn.Close
        Me.conn = Nothing
    End If
    If Me.rs <> Nothing Then
        Me.rs.Close
        Me.conn = Nothing
    End If
    If Me.strConn <> "" Then
        Me.strConn = ""
    End If
End Function