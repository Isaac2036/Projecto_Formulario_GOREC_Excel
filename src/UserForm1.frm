VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Registro de Expedientes"
   ClientHeight    =   14025
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   24225
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub showDataInList()
    
    Dim gk As New Geko
    Dim strCnn As String
    Dim sql As String
    Dim data() As Variant
    Dim reg As Integer
    Dim sqlCount As Integer
    Dim i As Integer
    
    i = getCountRecorset()
    
    If i <> 0 Then i = i - 1
    
    ReDim Data1(i, 16)
    ReDim Data2(i, 9)
    
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\expedienteBase.accdb"
    sql = "select * from Reversion"

    gk.strConnection = strCnn
    gk.showRecordset (sql)
    
    With gk.rs
        If .BOF And .EOF Then
            Debug.Print "No se encontaron registros", vbInformation
        Else
        
            .MoveFirst
            Do While Not (.EOF)
                Data1(reg, 0) = .Fields(0)
                Data1(reg, 1) = .Fields(1)
                Data1(reg, 2) = .Fields(2)
                Data1(reg, 3) = .Fields(3)
                Data1(reg, 4) = .Fields(4)
                Data1(reg, 5) = .Fields(5)
                Data1(reg, 6) = .Fields(6)
                Data1(reg, 7) = .Fields(7)
                Data1(reg, 8) = .Fields(8)
                Data1(reg, 9) = .Fields(9)
                Data1(reg, 10) = .Fields(10)
                Data1(reg, 11) = .Fields(11)
                Data1(reg, 12) = .Fields(12)
                Data1(reg, 13) = .Fields(13)
                Data1(reg, 14) = .Fields(14)
                Data1(reg, 15) = .Fields(15)
                Data1(reg, 16) = .Fields(16)
                Data2(reg, 0) = .Fields(17)
                Data2(reg, 1) = .Fields(18)
                Data2(reg, 2) = .Fields(19)
                Data2(reg, 3) = .Fields(20)
                Data2(reg, 4) = .Fields(21)
                Data2(reg, 5) = .Fields(22)
                Data2(reg, 6) = .Fields(23)
                Data2(reg, 7) = .Fields(24)
                Data2(reg, 8) = .Fields(25)
                Data2(reg, 9) = .Fields(26)
                .MoveNext
                reg = reg + 1
            Loop
            
            With LstExpedientes1
                .ColumnCount = 17
                .List = Data1
            End With
            
            With LstExpedientes1
                .ColumnCount = 10
                .List = Data2
            End With
            
        End If
    
    End With
    
    gk.freeMemory
End Sub
Private Function insertData(etapa As String, serie As String, uso As String, estado As String, proyecto As String, nro_partida As String, resolución As String, Expediente As String, anio As Integer, administrados As String, dni As String, zona As String, sector As String, barrio As String, grupo_residencial As Integer, manzana As String, lote As Integer, ultimo_documento As String, nro_folio As Integer, paquete As String, ubicacion_expediente As String, observacion As String, profesional As String, fecha_actualizacion As Date, rubro As String, area As String, contacto As String, metros As String) As Boolean
    Dim gk As New Geko
    Dim strCnnn As String
    Dim sql As String
    
    On Error GoTo Cath
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\expedienteBase.accdb"
    sql = "insert into Reversion values (" & etapa & ",'" & serie & "', '" & uso & "','" & estado & "','" & proyecto & "','" & nro_partida & "','" & resolucion & "','" & Expediente & "','" & anio & "','" & administrados & "','" & dni & "','" & zona & "','" & sector & "','" & barrio & "','" & grupo_residencial & "','" & manzana & "','" & lote & "','" & ultimo_documento & "','" & nro_folio & "','" & paquete & "','" & ubicacion_expediente & "','" & observacion & "','" & profesional & "','" & fecha_actualizacion & "','" & rubro & "','" & area & "','" & contacto & "','" & metro & "')"
    
    gk.strConnection = strCnn
    gk.executeCommand (sql)
    
    insertData = True
    
    Exit Function
Cath:
    Debug.Print "ERROR: " & Err.Description
    Debug.Print Err.Number
    insertData = False
    
End Function

Private Sub CmdExportarpdf_Click()
    Set r = Sheets("Reversion")
    uf = r.Range("B" & Rows.count).End(xlUp).Row + 1
    Worksheets("Reversion").Range("A4:H" & uf).ClearContents
    Dim Fcc As Date 'Declaracion de variable de tipo fecha
    Fcc = FormatDateTime(Now, vbShortDate) 'asiganos la fecha a la variable
    r.Cells(1, 1) = "GOBIERNO REGIONAL DEL CALLAO" 'Enviamos el nombre de la empresa a la Celda A1
    'Enviamos los datos del ListBox a la hoja Reportes
    For X = 0 To LstExpedientes1.ListCount - 1
        uf = r.Range("B" & Rows.count).End(xlUp).Row + 1
        'LstExpedientes1
        r.Cells(uf, 1).Value = LstExpedientes1.List(X, 0) 'ETAPA
        r.Cells(uf, 2).Value = LstExpedientes1.List(X, 1) 'SERIE
        r.Cells(uf, 3).Value = LstExpedientes1.List(X, 2) 'USO
        r.Cells(uf, 4).Value = LstExpedientes1.List(X, 3) 'ESTADO
        r.Cells(uf, 5).Value = LstExpedientes1.List(X, 4) 'PROYECTO
        r.Cells(uf, 6).Value = LstExpedientes1.List(X, 5) 'N° DE PARTIDA
        r.Cells(uf, 7).Value = LstExpedientes1.List(X, 6) 'RESOLUCION
        r.Cells(uf, 8).Value = LstExpedientes1.List(X, 7) 'EXPEDIENTE
        r.Cells(uf, 9).Value = LstExpedientes1.List(X, 8) 'AÑO
        r.Cells(uf, 10).Value = LstExpedientes1.List(X, 9) 'ADMINISTRADOS
        r.Cells(uf, 11).Value = LstExpedientes1.List(X, 10) 'DNI
        r.Cells(uf, 12).Value = LstExpedientes1.List(X, 11) 'ZONA
        r.Cells(uf, 13).Value = LstExpedientes1.List(X, 12) 'SECTOR
        r.Cells(uf, 14).Value = LstExpedientes1.List(X, 13) 'BARRIO
        r.Cells(uf, 15).Value = LstExpedientes1.List(X, 14) 'GRUPO
        r.Cells(uf, 16).Value = LstExpedientes1.List(X, 15) 'MANZANA
        r.Cells(uf, 17).Value = LstExpedientes1.List(X, 16) 'LOTE
        r.Cells(uf, 18).Value = LstExpedientes1.List(X, 17) 'ULTIMO
        r.Cells(uf, 19).Value = LstExpedientes1.List(X, 18) 'FOLIO
        'LstExpedientes2
        r.Cells(uf, 20).Value = LstExpedientes2.List(X, 0) 'PAQUETE
        r.Cells(uf, 21).Value = LstExpedientes2.List(X, 1) 'UBICACION
        r.Cells(uf, 22).Value = LstExpedientes2.List(X, 2) 'OBSERVACION
        r.Cells(uf, 23).Value = LstExpedientes2.List(X, 3) 'PROFESIONAL
        r.Cells(uf, 24).Value = LstExpedientes2.List(X, 4) 'RUBRO
        r.Cells(uf, 25).Value = LstExpedientes2.List(X, 5) 'AREA
        r.Cells(uf, 26).Value = LstExpedientes2.List(X, 6) 'CONTACTO
        r.Cells(uf, 27).Value = LstExpedientes2.List(X, 7) 'METROS
        
    Next X
    
    'Agregar Firma
    r.Cells(uf + 5, 2).Value = "Nombre y Firma"   'Firma
    
    nArch = InputBox("Escriba el nombre del archivo") 'Cuadro de dialogo para pedir el nombre que le daremos al reporte
    If nArch = Empty Then
    Else
        'Genera Archiv en pdf
        On Error GoTo errB 'Manejador de errores
        Sheets("Reportes").ExportAsFixedFormat Type:=xlTypePDF, Filename:=ThisWorkbook.Path & "\" & nArch & ".pdf", Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Exit Sub
errB:
        MsgBox ("Verifique el archivo pdf esta abierto, cierrelo para generarlo"), vbOKOnly + vbInformation, "Mensaje"
    End If
End Sub

Private Sub CmdImprimir_Click()
    resultado = MsgBox("¿Desea imprimir la lista de expedientes?", vbYesNo + vbQuestion, "Mensaje")
    Select Case resultado
    Case vbYes:
        X = Application.Dialogs(xlDialogPrinterSetup).Show ' Muestra las impresoras instaladas
        If X = False Then Exit Sub
        Sheets("Reversion").Select
        ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Case vbNo:
        Close
     End Select
End Sub

Private Sub CmdLimpiar_Click()
    Call LimpiarCampos 'Llama el procedimiento de limpiar campos
End Sub
Sub LimpiarCampos()
    'Limpiar campos
    TxtEtapa = ""
    TxtSerie = ""
    TxtUso = ""
    TxtEstado = ""
    TxtProyecto = ""
    TxtPartida = ""
    TxtResolucion = ""
    TxtExpediente = ""
    TxtAno = ""
    TxtAdministrados = ""
    TxtDni = ""
    TxtZona = ""
    TxtSector = ""
    TxtBarrio = ""
    TxtGrupo = ""
    TxtManzana = ""
    TxtLote = ""
    TxtUltimo = ""
    TxtFolio = ""
    TxtPaquete = ""
    TxtUbicacion = ""
    TxtObservacion = ""
    TxtProfesional = ""
    TxtActualizacion = ""
    TxtRubro = ""
    TxtArea = ""
    TxtContacto = ""
    TxtMetros = ""
End Sub

Private Sub CommandButton13_Click()

'    Dim etapa As String
'    Dim serie As String
'    Dim uso As String
'    Dim estado As String
'    Dim proyecto As String
'    Dim nro_partida As String
'    Dim resolucion As String
'    Dim expediente As String
'    Dim anio As Integer
'    Dim administrados As String
'    Dim dni As String
'    Dim zona As String
'    Dim sector As String
'    Dim barrio As String
'    Dim grupo_residencial As String
'    Dim manzana As String
'    Dim lote As Integer
'    Dim ultimo_documento As String
'    Dim nro_folio As Integer
'    Dim paquete As String
'    Dim ubicacion_expediente As String
'    Dim observacion As String
'    Dim profesional As String
'    Dim fecha_actualizacion As Date
'    Dim rubro As String
'    Dim area As String
'    Dim contacto As String
'    Dim metros As String
'    Dim result As Boolean
'
'    etapa = TxtEtapa
'    serie = TxtSerie
'    uso = TxtUso
'    estado = TxtEstado
'    proyecto = TxtProyecto
'    nro_partida = TxtPartida
'    resolucion = TxtResolucion
'    expediente = TxtExpediente
'    anio = TxtAnio
'    administrados = TxtAdministrados
'    dni = TxtDni
'    zona = TxtZona
'    sector = TxtSector
'    barrio = TxtBarrio
'    grupo_residencial = TxtGrupo
'    manzana = TxtManzana
'    lote = TxtLote
'    ultimo_documento = TxtUltimo
'    nro_folio = TxtFolio
'    paquete = TxtPaquete
'    ubicacion_expediente = TxtUbicacion
'    observacion = TxtObservacion
'    profesional = TxtProfesional
'    fecha_actualizacion = TxtActualizacion
'    rubro = TxtRubro
'    area = TxtArea
'    contacto = TxtContacto
'    metros = TxtMetros
'
'
'    result = insertData(etapa, serie, uso, estado, proyecto, nro_partida, resolucion, expediente, anio, administrados, dni, zona, sector, barrio, grupo_residencial, manzana, lote, ultimo_documento, nro_folio, paquete, ubicacion_expediente, observacion, profesional, fecha_actualizacion, rubro, area, contacto, metros)
'
'    If result Then
'        MsgBox "Los datos e han guardado", vbInformation
'    Else
'        MsgBox "Error no se pudo guardar los datos", vbExclamation
'    End If
'
'    Call showDataInList
'    Call LimpiarCampos
'
    
End Sub


Private Sub TxtEtapa_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtEtapa = WorksheetFunction.Proper(TxtEtapa)
End Sub

Private Sub TxtEtapa_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtEtapa = WorksheetFunction.Trim(TxtEtapa)
End Sub
Private Sub TxtSerie_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtSerie = WorksheetFunction.Proper(TxtSerie)
End Sub

Private Sub TxtSerie_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtSerie = WorksheetFunction.Trim(TxtSerie)
End Sub
Private Sub TxtUso_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtUso = WorksheetFunction.Proper(TxtUso)
End Sub

Private Sub TxtUso_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtUso = WorksheetFunction.Trim(TxtUso)
End Sub
Private Sub TxtEstado_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtEstado = WorksheetFunction.Proper(TxtEstado)
End Sub

Private Sub TxtEstado_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtEstado = WorksheetFunction.Trim(TxtEstado)
End Sub
Private Sub TxtProyecto_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtProyecto = WorksheetFunction.Proper(TxtProyecto)
End Sub

Private Sub TxtProyecto_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtProyecto = WorksheetFunction.Trim(TxtProyecto)
End Sub
Private Sub TxtPartida_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtPartida = WorksheetFunction.Proper(TxtPartida)
End Sub

Private Sub TxtPartida_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtPartida = WorksheetFunction.Trim(TxtPartida)
End Sub
Private Sub TxtResolucion_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtResolucion = WorksheetFunction.Proper(TxtResolucion)
End Sub

Private Sub TxtResolucion_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtResolucion = WorksheetFunction.Trim(TxtResolucion)
End Sub
Private Sub TxtExpediente_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtExpediente = WorksheetFunction.Proper(TxtExpediente)
End Sub
Private Sub TxtExpediente_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtExpediente = WorksheetFunction.Trim(TxtExpediente)
End Sub
Private Sub TxtAdministrados_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtAdministrados = WorksheetFunction.Proper(TxtAdministrados)
End Sub
Private Sub TxtAdministrados_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtAdministrados = WorksheetFunction.Trim(TxtAdministrados)
End Sub
Private Sub TxtDni_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtDni = WorksheetFunction.Proper(TxtDni)
End Sub
Private Sub TxtDni_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtDni = WorksheetFunction.Trim(TxtDni)
End Sub
Private Sub TxtZona_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtZona = WorksheetFunction.Proper(TxtZona)
End Sub
Private Sub TxtZona_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtZona = WorksheetFunction.Trim(TxtZona)
End Sub
Private Sub TxtSector_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtSector = WorksheetFunction.Proper(TxtSector)
End Sub
Private Sub TxtSector_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtSector = WorksheetFunction.Trim(TxtSector)
End Sub
Private Sub TxtBarrio_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtBarrio = WorksheetFunction.Proper(TxtBarrio)
End Sub
Private Sub TxtBarrio_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtBarrio = WorksheetFunction.Trim(TxtBarrio)
End Sub
Private Sub TxtManzana_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtManzana = WorksheetFunction.Proper(TxtManzana)
End Sub
Private Sub TxtManzana_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtManzana = WorksheetFunction.Trim(TxtManzana)
End Sub
Private Sub TxtUltimo_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtUltimo = WorksheetFunction.Proper(TxtUltimo)
End Sub
Private Sub TxtUltimo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtUltimo = WorksheetFunction.Trim(TxtUltimo)
End Sub
Private Sub TxtPaquete_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtPaquete = WorksheetFunction.Proper(TxtPaquete)
End Sub
Private Sub TxtPaquete_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtPaquete = WorksheetFunction.Trim(TxtPaquete)
End Sub
Private Sub TxtUbicacion_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtUbicacion = WorksheetFunction.Proper(TxtUbicacion)
End Sub
Private Sub TxtUbicacion_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtUbicacion = WorksheetFunction.Trim(TxtUbicacion)
End Sub
Private Sub TxtObservacion_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtObservacion = WorksheetFunction.Proper(TxtObservacion)
End Sub
Private Sub TxtObservacion_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtObservacion = WorksheetFunction.Trim(TxtObservacion)
End Sub
Private Sub TxtProfesional_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtProfesional = WorksheetFunction.Proper(TxtProfesional)
End Sub
Private Sub TxtProfesional_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtProfesional = WorksheetFunction.Trim(TxtProfesional)
End Sub
Private Sub TxtRubro_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtRubro = WorksheetFunction.Proper(TxtRubro)
End Sub
Private Sub TxtRubro_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtRubro = WorksheetFunction.Trim(TxtRubro)
End Sub
Private Sub TxtArea_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtArea = WorksheetFunction.Proper(TxtArea)
End Sub
Private Sub TxtArea_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtArea = WorksheetFunction.Trim(TxtArea)
End Sub
Private Sub TxtContacto_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtContacto = WorksheetFunction.Proper(TxtContacto)
End Sub
Private Sub TxtContacto_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtContacto = WorksheetFunction.Trim(TxtContacto)
End Sub
Private Sub TxtMetros_Change()
    'Funcion para convertir la primera letra en mayusculas
    TxtMetros = WorksheetFunction.Proper(TxtMetros)
End Sub
Private Sub TxtMetros_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Funcion para quitar espacios de mas en cualquier posicion
    TxtMetros = WorksheetFunction.Trim(TxtMetros)
End Sub
Sub MostrarExpediente()
    Dim b As Worksheet
    Set b = ThisWorkbook.Sheets("Reversion") 'Specify the workbook to avoid errors
    
    Dim uf As Long
    uf = b.Cells(b.Rows.count, "A").End(xlUp).Row 'Use Long instead of Integer for row counts
    
    LstExpedientes1.Clear 'Limpiar Listbox1
    LstExpedientes2.Clear 'Limpiar Listbox2

    With LstExpedientes1
        .ColumnCount = 19 'Numero de columnas
    End With

    With LstExpedientes2
        .ColumnCount = 8 'Numero de columnas
    End With
    
    'grega solo la última fila encontrada
    If Not IsEmpty(b.Cells(uf, 1)) Then
        'ñade nuevo item en LstExpedientes1
        With LstExpedientes1
           .AddItem b.Cells(uf, 1) 'columna: ETAPA
           .List(.ListCount - 1, 1) = b.Cells(uf, 2)  ' Columna: SERIE
           .List(.ListCount - 1, 2) = b.Cells(uf, 3) ' Columna: USO
           .List(.ListCount - 1, 3) = b.Cells(uf, 4) ' Columna: ESTADO
           .List(.ListCount - 1, 4) = b.Cells(uf, 5) ' Columna: PROYECTO
           .List(.ListCount - 1, 5) = b.Cells(uf, 6) ' Columna: N° DE PARTIDA
           .List(.ListCount - 1, 6) = b.Cells(uf, 7) ' Columna: RESOLUCION
           .List(.ListCount - 1, 7) = b.Cells(uf, 8) ' Columna: EXPEDIENTE
           .List(.ListCount - 1, 8) = b.Cells(uf, 9) ' Columna: AÑO
           .List(.ListCount - 1, 9) = b.Cells(uf, 10) ' Columna: ADMINISTRADOS
           .List(.ListCount - 1, 10) = b.Cells(uf, 11) ' Columna: DNI
           .List(.ListCount - 1, 11) = b.Cells(uf, 12) ' Columna: ZONA
           .List(.ListCount - 1, 12) = b.Cells(uf, 13) ' Columna: SECTOR
           .List(.ListCount - 1, 13) = b.Cells(uf, 14) ' Columna: BARRIO
           .List(.ListCount - 1, 14) = b.Cells(uf, 15) ' Columna: GRUPO
           .List(.ListCount - 1, 15) = b.Cells(uf, 16) ' Columna: MANZANA
           .List(.ListCount - 1, 16) = b.Cells(uf, 17) ' Columna: LOTE
           .List(.ListCount - 1, 17) = b.Cells(uf, 18) ' Columna: ULTIMO
         End With
    End If

    If Not IsEmpty(b.Cells(uf, 18)) Then
        'ñade nuevo item en LstExpedientes2
        With LstExpedientes2
           .AddItem b.Cells(uf, 19)
           .List(.ListCount - 1, 1) = b.Cells(uf, 20) ' Columna: PAQUETE
           .List(.ListCount - 1, 1) = b.Cells(uf, 21) ' Columna: PAQUETE
           .List(.ListCount - 1, 1) = b.Cells(uf, 22) ' Columna: UBICACION
           .List(.ListCount - 1, 1) = b.Cells(uf, 23) ' Columna: OBSERVACION
           .List(.ListCount - 1, 1) = b.Cells(uf, 24) ' Columna: PROFESIONAL
           .List(.ListCount - 1, 1) = b.Cells(uf, 25) ' Columna: RUBRO
           .List(.ListCount - 1, 1) = b.Cells(uf, 26) ' Columna: AREA
           .List(.ListCount - 1, 1) = b.Cells(uf, 27) ' Columna: CONTACTO
           .List(.ListCount - 1, 1) = b.Cells(uf, 28) ' Columna: METROS
        End With
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim b As Worksheet
    Set b = ThisWorkbook.Sheets("Reversion") 'Specify the workbook to avoid errors
    
    Dim uf As Long
    uf = b.Cells(b.Rows.count, "A").End(xlUp).Row

With Me.LstExpedientes1
    .ColumnCount = 19 'Al iniciar formulario estable listbox1 en 19 column
    .ColumnWidths = "40pt;66pt;47pt;49pt;80pt;78pt;90pt;89pt;35pt;120pt;82pt;75pt;38pt;35pt;30pt;25pt;15pt;40pt"
    
End With
With Me.LstExpedientes2
    .ColumnCount = 10
    .ColumnWidths = "25pt;55pt;47pt;180pt;40pt;120pt;40pt;89pt;20pt;20pt"
End With
    Call toolTips
End Sub
Sub toolTips()
    With Me
        With .CommandButton13
            .ControlTipText = "Guardar"
            .Default = True
        End With
    End With
End Sub

Private Function getCountRecorset() As Integer
    
    Dim gk As New Geko
    Dim strCnn As String
    Dim sql As String
    Dim count As Integer
    
    sql = "select count(id) from Reversion"
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\expedienteBase.accdb"
    
    gk.strConnection = strCnn
    gk.showRecordset (sql)
    
    With gk.rs
        If .BOF And .EOF Then
            getCountRecorset = 0
        Else
            getCountRecorset = .Fields(0)
        End If
    End With
    
    gk.freeMemory
    
End Function


