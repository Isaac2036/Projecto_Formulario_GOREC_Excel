VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBusquedaUsuarios 
   Caption         =   "Busqueda de expedientes"
   ClientHeight    =   10515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9510.001
   OleObjectBlob   =   "frmBusquedaUsuarios.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmBusquedaUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnBusuqeda_Click()
    Dim partida As String
    Dim expediente As String
    Dim anio As Variant
    Dim list As MSForms.listBox
    
    Set list = Me.ListBox1
    
    With list
        .Clear
        .ColumnCount = 3
    End With
    
    partida = TextBox2.Text
    expediente = TextBox3.Text
    anio = VBA.IIf(IsNumeric(TextBox18.Text), TextBox18.Text, Null)
    
    Call storage.filterByMultipleCriteria(list, partida, anio, expediente)
End Sub

Private Sub btnVer_Click()

'    If txtContenido.Text <> " " Then
'        MsgBox ("Buscando expedientes")
'    End If
    BuscarDatos
    frmViewBusqueda.Show
    
End Sub

Sub BuscarDatos()
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim listBox As MSForms.listBox
    Dim idExpediente As String
    Dim frmDestino As UserForm
    
    ' Obtener el valor del TextBox
    idExpediente = Me.TextBox3.Text
    
    ' Configuración de la conexión a la base de datos
    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\expedienteBase.accdb"
    
    ' Consulta SQL para buscar el expediente
    sql = "SELECT * FROM reversion WHERE expediente = '" & idExpediente & "'"
    
    ' Crear un objeto Recordset para almacenar los datos
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenStatic, adLockReadOnly
    
    
    ' Obtener el formulario destino
    Set frmDestino = frmViewBusqueda
    If frmDestino.ListBox1 Is Nothing Then
        MsgBox "El ListBox1 no existe en el formulario frmViewBusqueda"
        Exit Sub
    End If
    
    ' Obtener el ListBox del formulario destino
    Set listBox = frmDestino.ListBox1
    
    ' Limpiar el ListBox
    listBox.Clear
    
    ' Agregar los datos al ListBox
    If Not rs.EOF Then
        While Not rs.EOF
            listBox.AddItem rs!expediente
            rs.MoveNext
        Wend
    Else
        listBox.AddItem "No se encontró el expediente"
    End If
    
    ' Cerrar el Recordset y la conexión
    rs.Close
    cn.Close
    
    ' Liberar recursos
    Set rs = Nothing
    Set cn = Nothing
    Set frmDestino = Nothing

    
End Sub

Private Sub UserForm_Click()

End Sub
