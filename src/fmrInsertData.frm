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

Private Sub TextBox10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TextBox12_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TextBox14_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TextBox17_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    With TextBox17
        If Not IsDate(.Value) Then
            MsgBox "No es una " & .Value & " fecha valida", vbExclamation
            .Value = Empty
            .SetFocus
        End If
    End With
    
End Sub

Private Sub TextBox20_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub UserForm_Initialize()
     Call configFrm
End Sub

Sub validateInputs()

    Dim backColor As Long
    backColor = -2147483643
    
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            ctrl.backColor = backColor
            If ctrl.Value = Empty Then
                MsgBox "Por favor complete el campo " & ctrl.Name
                ctrl.SetFocus
                ctrl.backColor = vbRed
                Exit Sub
            End If
        ElseIf TypeOf ctrl Is MSForms.ComboBox Then
            ctrl.backColor = backColor
            If ctrl.ListIndex = -1 Then
                MsgBox "Por favor selecione un item de la lista"
                ctrl.SetFocus
                ctrl.backColor = vbRed
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
    Dim sql As String
    Dim r As New Reversion
    
    On Error GoTo Catch
    
    With r
        .id = getLastId() + 1
        .etapa = ComboBox1.Value
        .serie = ComboBox2.Value
        .uso = ComboBox3.Value
        .estado = ComboBox4.Value
        .proyecto = ComboBox5.Value
        .numeroPartida = TextBox1.Value
        .resolucion = TextBox2.Value
        .expedienteHojaRuta = TextBox3.Value
        .anioExpendiente = ComboBox9.Value
        .administrado = TextBox5.Value
        .dnis = TextBox6.Value
        .zona = TextBox7.Value
        .sector = TextBox8.Value
        .barrio = TextBox9.Value
        .grupoResidencial = TextBox10.Value
        .mz = TextBox11.Value
        .lote = TextBox12.Value
        .asuntoUtimoDocumento = TextBox13.Value
        .numeroFolio = TextBox14.Value
        .paquete = ComboBox6.Value
        .ubicacionExpediente = TextBox15.Value
        .observacion = TextBox16.Value
        .profesional = ComboBox7.Value
        .fechaActualizacion = TextBox17.Value
        .rubroComercioActividad = TextBox18.Value
        .area = ComboBox8.Value
        .contacto = TextBox19.Value
        .metro = TextBox20.Value
    End With
        
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
        .Parameters.Append .CreateParameter("ID", 2, adParamInput, , r.id)
        .Parameters.Append .CreateParameter("ETAPA", 202, adParamInput, 255, r.etapa)
        .Parameters.Append .CreateParameter("Serie", 202, adParamInput, 255, r.serie)
        .Parameters.Append .CreateParameter("USO", 202, adParamInput, 255, r.uso)
        .Parameters.Append .CreateParameter("ESTADO", 202, adParamInput, 255, r.estado)
        .Parameters.Append .CreateParameter("Proyecto", 202, adParamInput, 255, r.proyecto)
        .Parameters.Append .CreateParameter("Nro_partida", 202, adParamInput, 255, r.numeroPartida)
        .Parameters.Append .CreateParameter("RESOLUCION", 202, adParamInput, 255, r.resolucion)
        .Parameters.Append .CreateParameter("Expediente", 202, adParamInput, 255, r.expedienteHojaRuta)
        .Parameters.Append .CreateParameter("anio", 5, adParamInput, , r.anioExpendiente)
        .Parameters.Append .CreateParameter("Administrados", 203, adParamInput, 255, r.administrado)
        .Parameters.Append .CreateParameter("Dni", 203, adParamInput, 255, r.dnis)
        .Parameters.Append .CreateParameter("Zona", 202, adParamInput, 255, r.zona)
        .Parameters.Append .CreateParameter("Sector", 202, adParamInput, 255, r.sector)
        .Parameters.Append .CreateParameter("Barrio", 202, adParamInput, 255, r.barrio)
        .Parameters.Append .CreateParameter("Grupo_Residencial", 3, adParamInput, , r.grupoResidencial)
        .Parameters.Append .CreateParameter("Manzana", 202, adParamInput, 255, r.mz)
        .Parameters.Append .CreateParameter("LOTE", 5, adParamInput, , r.lote)
        .Parameters.Append .CreateParameter("Ultimo_documento", 203, adParamInput, 255, r.asuntoUtimoDocumento)
        .Parameters.Append .CreateParameter("Nro_folio", 5, adParamInput, , r.numeroFolio)
        .Parameters.Append .CreateParameter("PAQUETE", 202, adParamInput, 255, r.paquete)
        .Parameters.Append .CreateParameter("ubicacion_expediente", 202, adParamInput, 255, r.ubicacionExpediente)
        .Parameters.Append .CreateParameter("Observacion", 203, adParamInput, 255, r.observacion)
        .Parameters.Append .CreateParameter("Profesional", 202, adParamInput, 255, r.profesional)
        .Parameters.Append .CreateParameter("fecha_atualizacion", 7, adParamInput, , r.fechaActualizacion)
        .Parameters.Append .CreateParameter("Rubro", 202, adParamInput, 255, r.rubroComercioActividad)
        .Parameters.Append .CreateParameter("AREA", 202, adParamInput, 255, r.area)
        .Parameters.Append .CreateParameter("Contacto", 203, adParamInput, 255, r.contacto)
        .Parameters.Append .CreateParameter("METRO", 202, adParamInput, 255, r.metro)
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
    With Me
        .ComboBox1.TabIndex = 0
        .ComboBox2.TabIndex = 1
        .ComboBox3.TabIndex = 2
        .ComboBox4.TabIndex = 3
        .ComboBox5.TabIndex = 4
    End With
    
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
    With Me
        .TextBox1.TabIndex = 5
        .TextBox2.TabIndex = 6
        .TextBox3.TabIndex = 7
        .ComboBox9.TabIndex = 8
        .TextBox5.TabIndex = 9
        .TextBox6.TabIndex = 10
        .TextBox7.TabIndex = 11
        .TextBox8.TabIndex = 12
        .TextBox9.TabIndex = 13
        .TextBox10.TabIndex = 14
        .TextBox11.TabIndex = 15
        .TextBox12.TabIndex = 16
    End With
    
    With Me.Frame2
        With .ComboBox9
            For i = 2010 To 2060
                .AddItem i
            Next i
            
            .Value = 2024
            .MaxLength = 4
        End With
        
        .TextBox10.MaxLength = 4
        .TextBox11.MaxLength = 2
        .TextBox12.MaxLength = 3
        
    End With

'    Frame 3
    With Me
        .TextBox13.TabIndex = 17
        .TextBox14.TabIndex = 18
        .ComboBox6.TabIndex = 19
        .TextBox15.TabIndex = 20
        .TextBox16.TabIndex = 21
        .ComboBox7.TabIndex = 22
        .TextBox17.TabIndex = 23
        .TextBox18.TabIndex = 24
        .ComboBox8.TabIndex = 25
        .TextBox19.TabIndex = 26
        .TextBox20.TabIndex = 27
    End With
    
    With Me.Frame3
        With .ComboBox6
            For i = 1 To 9
                .AddItem "GABETA " & i
            Next i
            .ListIndex = 0
        End With
        
        With .TextBox16
            .WordWrap = True
            .MultiLine = True
            .EnterKeyBehavior = True
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
        
        .TextBox14.MaxLength = 4
        .TextBox17.Value = Date
        
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


