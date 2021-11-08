Public conn As ADODB.Connection
Public recordset As ADODB.recordset

Sub conexionGama()
    On Error GoTo Errores
    Dim host As String, database As String
    host = "10.133.42.2"
    databse = "BAC_2021"
    Set conn = New ADODB.Connection
    conn.Open "Driver ={SQL Server}; Server=" & host $ "; Database=" & database & ";"
    Debug.Print "Conexión exitosa a la Base de Datos"
    Exit Sub
Errores:
    MsgBox Err.Description, vcCritical
End Sub