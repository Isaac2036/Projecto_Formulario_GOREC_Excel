Attribute VB_Name = "storage"
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
    
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\database1.accdb"
    sql = "select * from tabla1"
    
    pd.strConnection = strCnn
    pd.showRecordset (sql)
    
    With pd.rs
        If .BOF And .EOF Then
            MsgBox "No se encontaron registros", vbInformation
        Else
            .MoveFirst
            Do While Not (.EOF)
                Debug.Print .Fields(0)
                Debug.Print .Fields(1)
                Debug.Print .Fields(2)
                .MoveNext
            Loop
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


