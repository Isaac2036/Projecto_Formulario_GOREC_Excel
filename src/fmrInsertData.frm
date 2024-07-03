VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmrInsertData 
   Caption         =   "UserForm3"
   ClientHeight    =   11445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17820
   OleObjectBlob   =   "fmrInsertData.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "fmrInsertData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton13_Click()
    Call validateInputs
End Sub

Private Sub CommandButton16_Click()
    Call clearInputs
End Sub

Private Sub Label425_Click()

    Set ctrlFecha = TextBox17
    frm_calendario.Show
     
End Sub

Private Sub UserForm_Initialize()
     Call configFrm
End Sub

Sub validateInputs()
        
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            ctrl.BackColor = -2147483643
            If ctrl.Value = Empty Then
                MsgBox "Por favor complete el campo " & ctrl.Name
                ctrl.SetFocus
                ctrl.BackColor = vbRed
                Exit Sub
            End If
        ElseIf TypeOf ctrl Is MSForms.ComboBox Then
            ctrl.BackColor = -2147483643
            If ctrl.ListIndex = -1 Then
                MsgBox "Por favor selecione un item de la lista"
                ctrl.SetFocus
                ctrl.BackColor = vbRed
                Exit Sub
            End If
        End If
    Next ctrl
    
    Call insertNewRecord
    Call clearInputs
    
End Sub
Sub clearInputs()
        
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.TextBox And ctrl.Enabled Then
           ctrl.Value = Empty
        ElseIf TypeOf ctrl Is MSForms.ComboBox And ctrl.Enabled Then
           ctrl.ListIndex = -1
        End If
    Next ctrl

End Sub
Sub insertNewRecord()
    
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim strCnn As String
    Dim lastId As Integer
    Dim sql As String

    On Error GoTo Catch
    
    lastId = getLastId() + 1
    
    ' Configurar la conexión
    Set cnn = New ADODB.Connection
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\expedienteBase.accdb"
    cnn.Open strCnn

    ' Configurar el comando
    Set cmd = New ADODB.Command
    sql = "INSERT INTO reversion (ID, ETAPA, Serie, USO, ESTADO, Proyecto, Nro_partida, RESOLUCION,"
    sql = sql & " Expediente, anio, Administrados, Dni, Zona, Sector, Barrio, Grupo_Residencial, Manzana, LOTE, Ultimo_documento, Nro_folio, PAQUETE,"
    sql = sql & " ubicacion_expediente, Observacion, Profesional, fecha_atualizacion, Rubro, AREA, Contacto, METRO)"
    sql = sql & " VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    
    With cmd
        .ActiveConnection = cnn
        .CommandText = sql
        .CommandType = adCmdText

        ' Agregar parámetros
         .Parameters.Append .CreateParameter("ID", 2, adParamInput, , lastId)
        .Parameters.Append .CreateParameter("ETAPA", 202, adParamInput, 255, UCase(ComboBox1.Value))
        .Parameters.Append .CreateParameter("Serie", 202, adParamInput, 255, UCase(ComboBox2.Value))
        .Parameters.Append .CreateParameter("USO", 202, adParamInput, 255, UCase(ComboBox3.Value))
        .Parameters.Append .CreateParameter("ESTADO", 202, adParamInput, 255, UCase(ComboBox4.Value))
        .Parameters.Append .CreateParameter("Proyecto", 202, adParamInput, 255, UCase(ComboBox5.Value))
        .Parameters.Append .CreateParameter("Nro_partida", 202, adParamInput, 255, UCase(TextBox1.Text))
        .Parameters.Append .CreateParameter("RESOLUCION", 202, adParamInput, 255, UCase(TextBox2.Text))
        .Parameters.Append .CreateParameter("Expediente", 202, adParamInput, 255, UCase(TextBox3.Text))
        .Parameters.Append .CreateParameter("anio", 5, adParamInput, , 2024)
        .Parameters.Append .CreateParameter("Administrados", 203, adParamInput, 255, UCase(TextBox5.Text))
        .Parameters.Append .CreateParameter("Dni", 203, adParamInput, 255, UCase(TextBox6.Text))
        .Parameters.Append .CreateParameter("Zona", 202, adParamInput, 255, UCase(TextBox7.Text))
        .Parameters.Append .CreateParameter("Sector", 202, adParamInput, 255, UCase(TextBox8.Text))
        .Parameters.Append .CreateParameter("Barrio", 202, adParamInput, 255, UCase(TextBox9.Text))
        .Parameters.Append .CreateParameter("Grupo_Residencial", 3, adParamInput, , TextBox10.Value)
        .Parameters.Append .CreateParameter("Manzana", 202, adParamInput, 255, TextBox11.Text)
        .Parameters.Append .CreateParameter("LOTE", 5, adParamInput, , TextBox12.Value)
        .Parameters.Append .CreateParameter("Ultimo_documento", 203, adParamInput, 255, UCase(TextBox13.Text))
        .Parameters.Append .CreateParameter("Nro_folio", 5, adParamInput, , TextBox14.Value)
        .Parameters.Append .CreateParameter("PAQUETE", 202, adParamInput, 255, UCase(ComboBox6.Value))
        .Parameters.Append .CreateParameter("ubicacion_expediente", 202, adParamInput, 255, UCase(TextBox15.Text))
        .Parameters.Append .CreateParameter("Observacion", 203, adParamInput, 255, UCase(TextBox16.Text))
        .Parameters.Append .CreateParameter("Profesional", 202, adParamInput, 255, ComboBox7.Value)
        .Parameters.Append .CreateParameter("fecha_atualizacion", 7, adParamInput, , CDate(TextBox17.Value))
        .Parameters.Append .CreateParameter("Rubro", 202, adParamInput, 255, UCase(TextBox18.Text))
        .Parameters.Append .CreateParameter("AREA", 202, adParamInput, 255, UCase(ComboBox8.Value))
        .Parameters.Append .CreateParameter("Contacto", 203, adParamInput, 255, UCase(TextBox19.Text))
        .Parameters.Append .CreateParameter("METRO", 202, adParamInput, 255, UCase(TextBox18.Text))
    End With

    ' Ejecutar el comando
    cmd.Execute

    ' Liberar recursos
    Set cmd = Nothing
    cnn.Close
    Set cnn = Nothing

    MsgBox "Se ingreso correctamente la data.", vbInformation
    Exit Sub

Catch:
    MsgBox "Error : " & Err.Description, vbCritical
    Debug.Print "ERROR: " & Err.Description
    Debug.Print Err.Number

End Sub

Sub configFrm()

    'frame1
    With Me.Frame1
        With .ComboBox1
                .AddItem "Pendiente"
                .AddItem "Asignado"
                .AddItem "Evaluando"
                .ListIndex = 0
        End With
        
        With .ComboBox2
            .AddItem "reversion"
            .ListIndex = 0
            .Enabled = False
        End With
        
        With .ComboBox3
            .AddItem "vivienda"
            .AddItem "comercial"
            .AddItem "publico"
            .ListIndex = 0
        End With
        
         With .ComboBox4
            .AddItem "pendiente de asignación"
            .AddItem "pendiente de revisión"
            .AddItem "pendiente notificación"
            .AddItem "asignado especialista"
            .AddItem "asignado jefatura"
            .AddItem "completado"
            .ListIndex = 0
        End With
        
        With .ComboBox5
            .AddItem "PECP"
            .ListIndex = 0
            .Enabled = False
        End With
    End With
    
    'frame 2
    With Me.Frame2
        With .TextBox4
            .Value = Format$(Now(), "yyyy")
            .Enabled = False
        End With
    End With

'    Frame 3
    With Me.Frame3
        With .ComboBox6
            For i = 1 To 9
                .AddItem "GABETA " & i
            Next i
            .ListIndex = 0
        End With
        
        With .TextBox16
            .WordWrap = True
            .MultiLine = False
            .ScrollBars = fmScrollBarsVertical
        End With
        
        With .ComboBox7
            .RowSource = "nombre_abogados"
        End With
        
         With .ComboBox8
            .AddItem "OGP"
            .ListIndex = 0
            .Enabled = False
        End With
        
        With .TextBox17
            .Value = Date
        End With
        
    End With
    
End Sub
Function getLastId() As Integer
    
    Dim pd As New Geko
    Dim sql As String
    Dim strCnn As String
    Dim rs As ADODB.Recordset
    
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\expedienteBase.accdb"
    sql = "select max(id) from reversion"
    
    pd.strConnection = strCnn
    pd.showRecordset (sql)
    
    With pd.rs
        If .BOF And .EOF Then
            getLastId = 0
        Else
            getLastId = .Fields(0)
        End If
    End With
    
End Function


