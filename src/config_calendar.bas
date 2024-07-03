Attribute VB_Name = "config_calendar"
Public pdiaMes As Date
Public diaSemana As Byte
Public nFecha As Date
Public Marcar As Byte
Public SeleccDia As Byte
Public SeleccMes As Byte
Public SeleccAño As Integer
Public SeleccionFecha As Date
Public poscArriba As Double                             'posicion (top) del control que invoco el calendario
Public poscIzquierda As Double                          'posición (left) del control que lo invoco al calendario
Public ctrlFecha As Object
Public nombreComputador$

Sub Limpiar_Etiquetas_Frame1()
    Dim ctrl As control
    
    For Each ctrl In frm_calendario.Frame1.Controls
        ctrl.Caption = ""
    Next ctrl
End Sub
Sub Asignar_Tag()
    Dim ctrl As control
    n = 1
    For Each ctrl In frm_calendario.Frame1.Controls
        ctrl.Tag = n
        n = n + 1
    Next ctrl
End Sub
Sub Configurar_Año()
    Dim i As Integer
    
    For i = 1920 To 2100
         frm_calendario.ComboBox1.AddItem i
    Next i
    
End Sub
Sub Configurar_Meses()
    Dim i As Byte
    
    For i = 0 To 11
        frm_calendario.ComboBox2.AddItem Meses(i)
    Next i
    
End Sub
Sub Año_Mes_Actual()
    
    Dim añoActual As Integer
    Dim mesActual As Byte
    
    añoActual = Year(Date) - CDate("1920")
    mesActual = Month(Date)
    
    With frm_calendario
        .ComboBox1.ListIndex = añoActual
        .ComboBox2.ListIndex = mesActual - 1
    End With
    
End Sub
Sub Inicio_Del_Mes()
'    On Error Resume Next
'    With frm_calendario
'
'        Select Case .Label51.Caption
'             Case Is = ""
'                .Label51.Caption = CDate(1 & "/" & Month(Date) & "/" & Year(Date))
'                .Label52.Caption = Weekday(.Label51, vbTuesday)
'             Case Else
'                 .Label51.Caption = CDate(1 & "/" & .ComboBox2.ListIndex + 1 & "/" & .ComboBox1.Value)
'                .Label52.Caption = Weekday(.Label51, vbTuesday)
'        End Select
'                pdiaMes = .Label51.Caption
'                diaSemana = .Label52.Caption
'    End With
'
'    Call Gráficar_Días
    With frm_calendario

        Select Case pdiaMes
             Case Is = 0
                pdiaMes = CDate(1 & "/" & Month(Date) & "/" & Year(Date))
                diaSemana = Weekday(pdiaMes, vbTuesday)
             Case Else
                pdiaMes = CDate(1 & "/" & .ComboBox2.ListIndex + 1 & "/" & .ComboBox1.Value)
                diaSemana = Weekday(pdiaMes, vbTuesday)
        End Select
            
    End With

    Call Gráficar_Días
    
End Sub
Sub Gráficar_Días()
    Dim ctrl As control
    
    With frm_calendario
        For Each ctrl In .Frame1.Controls
            ctrl.Caption = pdiaMes - diaSemana + n
            nFecha = ctrl.Caption
            n = n + 1
            If Month(nFecha) <> Month(pdiaMes) Then
                ctrl.Caption = ""
            Else
                ctrl.Caption = WorksheetFunction.Text(nFecha, "d")
            End If
        Next ctrl
    End With
    
End Sub
Sub Recorrer_Restaurar_Etiquetas_Fechas()
    Dim ctrl As control
    
    For Each ctrl In frm_calendario.Frame1.Controls
        ctrl.Font.Bold = False
        ctrl.BorderStyle = fmBorderStyleNone
    Next ctrl
End Sub
Sub Recorrer_Marcar_Etiquetas_Fechas()
    Dim ctrl As control
    
    For Each ctrl In frm_calendario.Frame1.Controls
        If ctrl.Caption <> "" And ctrl.Tag = Marcar Then
            ctrl.Font.Bold = True
            ctrl.BorderStyle = fmBorderStyleSingle
        End If
    Next ctrl
End Sub

'POR FIN OBTENEMOS LA FECHA SELCCIONADA ===================================
'===========================================================================
Rem obtenemos la fecha seleccionada y la volcamos en msgbox
Sub Fecha_Seleccionada()
    
    With frm_calendario
        SeleccAño = .ComboBox1.Value
        SeleccMes = .ComboBox2.ListIndex + 1
    End With
    
    SeleccionFecha = (SeleccDia & "/" & SeleccMes & "/" & SeleccAño)
    
'    MsgBox SeleccionFecha
    ctrlFecha = SeleccionFecha
End Sub
Sub Ocultar_Ayuda()
      
      With UserForm2
      
            .Label252.ForeColor = RGB(14, 154, 260)
            .Label252.Font.Underline = True
            .Label252.ControlTipText = "Crear nuevo usuario"
            .MultiPage13.Style = fmTabStyleNone
            .MultiPage13.Visible = False
      
            For i = 254 To 258 Step 2
                  .Controls("Label" & i).ForeColor = RGB(14, 154, 260)
            Next i
            
             For i = 253 To 257 Step 2
                  .Controls("Label" & i).ForeColor = RGB(32, 32, 228)
            Next i
            
      End With
End Sub
Function Meses(ByVal nMes As Byte) As String
    Dim nomMes(0 To 11) As String
    
    nomMeses = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", _
                                "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
    
    Meses = nomMeses(nMes)
End Function

