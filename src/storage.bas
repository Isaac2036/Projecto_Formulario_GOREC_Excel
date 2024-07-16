Attribute VB_Name = "storage"
Public r As New Reversion
Sub searchExpediente(list As MSForms.ListBox, Optional partida As String, Optional expediente As String, Optional anio As Integer)

     
    Dim pd As New Geko
    Dim sql As String
    Dim strCnn As String
    Dim rs As ADODB.Recordset
    
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\expedienteBase.accdb"
    sql = "SELECT nro_partida, expediente, anio FROM reversion WHERE nro_partida = '" & partida & "' OR expediente = '" & expediente & "' OR anio = " & anio
    
    pd.strConnection = strCnn
    pd.showRecordset (sql)
    
    list.Clear
    
    With pd.rs
        If .BOF And .EOF Then
            list.AddItem "No se encontraron registros para mostrar."
        Else
            
            list.ColumnCount = 3
            .MoveFirst
            Do While Not (.EOF)
                list.AddItem .Fields(0)
                list.list(list.ListCount - 1, 1) = .Fields(1)
                list.list(list.ListCount - 1, 2) = .Fields(2)
                .MoveNext
            Loop
        End If
    End With
End Sub
Public Function insertNewRecord(frm As UserForm) As Boolean
    
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim strCnn As String
    Dim sql As String
'    Dim r As New Reversion
    
    On Error GoTo Catch
    
    With r
        .id = getLastId() + 1
        .etapa = frm.ComboBox1.Value
        .serie = frm.ComboBox2.Value
        .uso = frm.ComboBox3.Value
        .estado = frm.ComboBox4.Value
        .proyecto = frm.ComboBox5.Value
        .numeroPartida = frm.TextBox1.Value
        .resolucion = frm.TextBox2.Value
        .expedienteHojaRuta = frm.TextBox3.Value
        .anioExpendiente = frm.ComboBox9.Value
        .administrado = frm.TextBox5.Value
        .dnis = frm.TextBox6.Value
        .zona = frm.TextBox7.Value
        .sector = frm.TextBox8.Value
        .barrio = frm.TextBox9.Value
        .grupoResidencial = frm.TextBox10.Value
        .mz = frm.TextBox11.Value
        .lote = frm.TextBox12.Value
        .asuntoUtimoDocumento = frm.TextBox13.Value
        .numeroFolio = frm.TextBox14.Value
        .paquete = frm.ComboBox6.Value
        .ubicacionExpediente = frm.TextBox15.Value
        .observacion = frm.TextBox16.Value
        .profesional = frm.ComboBox7.Value
        .fechaActualizacion = frm.TextBox17.Value
        .rubroComercioActividad = frm.TextBox18.Value
        .area = frm.ComboBox8.Value
        .contacto = frm.TextBox19.Value
        .metro = frm.TextBox20.Value
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

    insertNewRecord = True
    Exit Function

Catch:

    MsgBox "Error : " & Err.Description, vbCritical
    Debug.Print "ERROR: " & Err.Description
    Debug.Print Err.Number

End Function

Public Function insertEditRecord(frm As UserForm) As Boolean
    
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim strCnn As String
    Dim sql As String
'    Dim r As New Reversion
    
    On Error GoTo Catch
    
    With r
        .id = getLastId()
        .etapa = frm.ComboBox1.Value
        .serie = frm.ComboBox2.Value
        .uso = frm.ComboBox3.Value
        .estado = frm.ComboBox4.Value
        .proyecto = frm.ComboBox5.Value
        .numeroPartida = frm.TextBox1.Value
        .resolucion = frm.TextBox2.Value
        .expedienteHojaRuta = frm.TextBox3.Value
        .anioExpendiente = frm.ComboBox9.Value
        .administrado = frm.TextBox5.Value
        .dnis = frm.TextBox6.Value
        .zona = frm.TextBox7.Value
        .sector = frm.TextBox8.Value
        .barrio = frm.TextBox9.Value
        .grupoResidencial = frm.TextBox10.Value
        .mz = frm.TextBox11.Value
        .lote = frm.TextBox12.Value
        .asuntoUtimoDocumento = frm.TextBox13.Value
        .numeroFolio = frm.TextBox14.Value
        .paquete = frm.ComboBox6.Value
        .ubicacionExpediente = frm.TextBox15.Value
        .observacion = frm.TextBox16.Value
        .profesional = frm.ComboBox7.Value
        .fechaActualizacion = frm.TextBox17.Value
        .rubroComercioActividad = frm.TextBox18.Value
        .area = frm.ComboBox8.Value
        .contacto = frm.TextBox19.Value
        .metro = frm.TextBox20.Value
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

    insertNewRecord = True
    Exit Function

Catch:

    MsgBox "Error : " & Err.Description, vbCritical
    Debug.Print "ERROR: " & Err.Description
    Debug.Print Err.Number

End Function

Function viewNewRecord(frm As UserForm) As Boolean
        
        With frm.lstView1
            .Clear
            .ColumnCount = 9
            .ColumnWidths = "60 pt;58 pt;58 pt;150 pt;90 pt;90 pt;80 pt;50 pt;10 pt"
            .AddItem r.etapa
            .list(.ListCount - 1, 1) = r.serie
            .list(.ListCount - 1, 2) = r.uso
            .list(.ListCount - 1, 3) = r.estado
            .list(.ListCount - 1, 4) = r.proyecto
            .list(.ListCount - 1, 5) = r.numeroPartida
            .list(.ListCount - 1, 6) = r.resolucion
            .list(.ListCount - 1, 7) = r.expedienteHojaRuta
            .list(.ListCount - 1, 8) = r.anioExpendiente
        End With
        
        'lstView2
        With frm.lstView2
            .Clear
            .ColumnCount = 10
            .ColumnWidths = "148 pt;74 pt;60 pt;35 pt;30 pt;30 pt;18 pt;80 pt;85 pt;10 pt"
            .AddItem r.administrado
            .list(.ListCount - 1, 1) = r.dnis
            .list(.ListCount - 1, 2) = r.zona
            .list(.ListCount - 1, 3) = r.sector
            .list(.ListCount - 1, 4) = r.barrio
            .list(.ListCount - 1, 5) = r.grupoResidencial
            .list(.ListCount - 1, 6) = r.mz
            .list(.ListCount - 1, 7) = r.lote
            .list(.ListCount - 1, 8) = r.asuntoUtimoDocumento
            .list(.ListCount - 1, 9) = r.numeroFolio
        End With

        'lstView3
        With frm.lstView3
            .Clear
            .ColumnCount = 9
            .ColumnWidths = "60 pt; 100 pt;120 pt;70 pt;60 pt;60 pt;60 pt;60 pt;60 pt"
            .AddItem r.paquete
            .list(.ListCount - 1, 1) = r.ubicacionExpediente
            .list(.ListCount - 1, 2) = r.observacion
            .list(.ListCount - 1, 3) = r.profesional
            .list(.ListCount - 1, 4) = r.fechaActualizacion
            .list(.ListCount - 1, 5) = r.rubroComercioActividad
            .list(.ListCount - 1, 6) = r.area
            .list(.ListCount - 1, 7) = r.contacto
            .list(.ListCount - 1, 8) = r.metro
        End With

End Function
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

'Function editNewRecord(frm As UserForm) As Boolean
'
'    Dim cnn As ADODB.Connection
'    Dim cmd As ADODB.Command
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Dim strCnn As String
'    Dim sql As String
''    Dim r As New Reversion
'
'    ' Configurar la conexión
'    Set cnn = New ADODB.Connection
'    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\expedienteBase.accdb"
'    cnn.Open strCnn
'    rs.Open strCnn, cnn
'
'    ' Configurar el comando
'    Set cmd = New ADODB.Command
'    sql = "SELECT (ID, ETAPA, Serie, USO, ESTADO, Proyecto, Nro_partida, RESOLUCION,"
'    sql = sql & " Expediente, anio, Administrados, Dni, Zona, Sector, Barrio, Grupo_Residencial, Manzana, LOTE, Ultimo_documento, Nro_folio, PAQUETE,"
'    sql = sql & " ubicacion_expediente, Observacion, Profesional, fecha_atualizacion, Rubro, AREA, Contacto, METRO) FROM reversion ORDER BY id DESC LIMIT 1;"
'
'
'    On Error GoTo Catch
'
'        If Not rs.EOF Then
'            Dim datos As New Collection
'            datos.Add rs!campo1
'            datos.Add rs!campo2
'            datos.Add rs!campo3
'            datos.Add rs!campo4
'            datos.Add rs!campo5
'            datos.Add rs!campo6
'            datos.Add rs!campo7
'            datos.Add rs!campo8
'            datos.Add rs!campo9
'            datos.Add rs!campo10
'            datos.Add rs!campo11
'            datos.Add rs!campo12
'            datos.Add rs!campo13
'            datos.Add rs!campo14
'            datos.Add rs!campo15
'            datos.Add rs!campo16
'            datos.Add rs!campo17
'            datos.Add rs!campo18
'            datos.Add rs!campo19
'            datos.Add rs!campo20
'            datos.Add rs!campo21
'            datos.Add rs!campo22
'            datos.Add rs!campo23
'            datos.Add rs!campo24
'            datos.Add rs!campo25
'            datos.Add rs!campo26
'            datos.Add rs!campo27
'            datos.Add rs!campo28
'            datos.Add rs!campo29
'        End If
'
'        Dim frmEditarExp As New frmDestino
'        frmEditarExp.RecibirDatos datos
'        frmDestino.Show
'
'        rs.Close
'        conn.Close
'        Set rs = Nothing
'        Set conn = Nothing
'
'    editNewRecord = True
'    Exit Function
'
'Catch:
'
'    MsgBox "Error : " & Err.Description, vbCritical
'    Debug.Print "ERROR: " & Err.Description
'    Debug.Print Err.Number
'
'
'End Function

