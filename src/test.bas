Attribute VB_Name = "test"
Sub test_reversion()
    
    Dim r As New Reversion
    
    With r
        .dnis = "4589521"
        Debug.Print TypeName(.dnis)
    End With
    
    
End Sub
Sub test_()
    Dim gk As New Geko
    Dim sql As String
    Dim strCnn As String
    
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\database1.accdb"
    sql = "delete from tabla1 where id = 1"
    
    gk.strConnection = strCnn
    gk.executeCommand (sql)
End Sub

Sub test_insert()
    
    Dim gk As New Geko
    Dim sql As String
    Dim strCnn As String
    
    On Error GoTo Cath
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\database1.accdb"
    sql = "insert into tabla1 values (1,'python', 'soy genial')"
    
    gk.strConnection = strCnn
    gk.executeCommand (sql)
    
    Exit Sub
Cath:
    Debug.Print "ERROR: " & Err.Description
    Debug.Print Err.Number
End Sub

Sub test_read()
    
    Dim pd As New Geko
    Dim sql As String
    Dim strCnn As String
    Dim rs As ADODB.Recordset
    
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\expedienteBase.accdb"
    sql = "select * from reversion"
'    sql = "select max(id) from reversion"
    
    pd.strConnection = strCnn
    pd.showRecordset (sql)
    
    With pd.rs
        If .BOF And .EOF Then
            MsgBox "No se encontaron registros", vbInformation
        Else
        Dim head As String
'            .MoveFirst
'            Do While Not (.EOF)
'                Debug.Print .Fields(0)
'                Debug.Print .Fields(1)
'                Debug.Print .Fields(2)
'                .MoveNext
'            Loop
            For i = 0 To .Fields.Count - 1
'                head = head & "," & .Fields(i).Name
                Debug.Print "valor: "; .Fields(i)
                Debug.Print "Nombre de campo: "; .Fields(i).name
                Debug.Print "Tipo de campo: "; .Fields(i).Type
                Debug.Print "=======" & vbCrLf
            Next i
'            Debug.Print head
        End If
    End With
    
    
End Sub

Sub test_update()
    Dim gk As New Geko
    Dim sql As String
    Dim strCnn As String
    
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\database1.accdb"
    sql = "update tabla1 set name_tb = 'Go', description = 'es rapido' where id = 1"
    
    gk.strConnection = strCnn
    gk.executeCommand (sql)
End Sub
Function getLastRecord() As Integer
    
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strCnn As String

'    On Error GoTo Catch

    ' Configurar la conexión
    Set cnn = New ADODB.Connection
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\expedienteBase.accdb"
    cnn.Open strCnn
    Set rs = cnn.Execute("select * from reversion")
    
    If rs.EOF And rs.BOF Then
        getLastRecord = rs.Fields(0)
    Else
        getLastRecord = 0
    End If
    
End Function
