VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_calendario 
   Caption         =   "Calendario"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3030
   OleObjectBlob   =   "frm_calendario.frx":0000
End
Attribute VB_Name = "frm_calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ComboBox1_Change()
    Call Inicio_Del_Mes
End Sub

Private Sub ComboBox2_Change()
    Call Inicio_Del_Mes
End Sub

Private Sub CommandButton4_Click()
    Unload Me
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Recorrer_Restaurar_Etiquetas_Fechas
End Sub

Private Sub Label50_Click()

End Sub

Private Sub Label8_Click()
    If Label8 <> "" Then
        SeleccDia = Label8.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label8.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label9_Click()
    If Label9 <> "" Then
        SeleccDia = Label9.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label9.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label10_Click()
    If Label10 <> "" Then
        SeleccDia = Label10.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label10.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label11_Click()
    If Label11 <> "" Then
        SeleccDia = Label11.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label11.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label12_Click()
    If Label12 <> "" Then
        SeleccDia = Label12.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label12.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label13_Click()
    If Label13 <> "" Then
        SeleccDia = Label13.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label13.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label14_Click()
    If Label14 <> "" Then
        SeleccDia = Label14.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label14.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label15_Click()
    If Label15 <> "" Then
        SeleccDia = Label15.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label15.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label16_Click()
    If Label16 <> "" Then
        SeleccDia = Label16.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label16.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label17_Click()
    If Label17 <> "" Then
        SeleccDia = Label17.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label17.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label18_Click()
    If Label18 <> "" Then
        SeleccDia = Label18.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label18.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label19_Click()
    If Label19 <> "" Then
        SeleccDia = Label19.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label19.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label20_Click()
    If Label20 <> "" Then
        SeleccDia = Label20.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label20.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label21_Click()
    If Label21 <> "" Then
        SeleccDia = Label21.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label21.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label22_Click()
    If Label22 <> "" Then
        SeleccDia = Label22.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label22.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label23_Click()
    If Label23 <> "" Then
        SeleccDia = Label23.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label23.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label24_Click()
    If Label24 <> "" Then
        SeleccDia = Label24.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label24.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label25_Click()
    If Label25 <> "" Then
        SeleccDia = Label25.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label25.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label26_Click()
    If Label26 <> "" Then
        SeleccDia = Label26.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label26.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label27_Click()
    If Label27 <> "" Then
        SeleccDia = Label27.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label27.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label28_Click()
    If Label28 <> "" Then
        SeleccDia = Label28.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label28_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label28.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label29_Click()
    If Label29 <> "" Then
        SeleccDia = Label29.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label29_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label29.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label30_Click()
    If Label30 <> "" Then
        SeleccDia = Label30.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label30_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label30.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label31_Click()
    If Label31 <> "" Then
        SeleccDia = Label31.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label31.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label32_Click()
    If Label32 <> "" Then
        SeleccDia = Label32.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label32.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label33_Click()
    If Label33 <> "" Then
        SeleccDia = Label33.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label33.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label34_Click()
    If Label34 <> "" Then
        SeleccDia = Label34.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label34.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label35_Click()
    If Label35 <> "" Then
        SeleccDia = Label35.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label35.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label36_Click()
    If Label36 <> "" Then
        SeleccDia = Label36.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label36.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label37_Click()
    If Label37 <> "" Then
        SeleccDia = Label37.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label37.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label38_Click()
    If Label38 <> "" Then
        SeleccDia = Label38.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label38_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label38.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label39_Click()
    If Label39 <> "" Then
        SeleccDia = Label39.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label39_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label39.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label40_Click()
    If Label40 <> "" Then
        SeleccDia = Label40.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label40_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label40.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label41_Click()
    If Label41 <> "" Then
        SeleccDia = Label41.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label41.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label42_Click()
    If Label42 <> "" Then
        SeleccDia = Label42.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label42.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label43_Click()
    If Label43 <> "" Then
        SeleccDia = Label43.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label43_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label43.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label44_Click()
    If Label44 <> "" Then
        SeleccDia = Label44.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label44_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label44.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label45_Click()
    If Label45 <> "" Then
        SeleccDia = Label45.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label45_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label45.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label46_Click()
    If Label46 <> "" Then
        SeleccDia = Label46.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label46_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label46.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label47_Click()
    If Label47 <> "" Then
        SeleccDia = Label47.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label47_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label47.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label48_Click()
    If Label48 <> "" Then
        SeleccDia = Label48.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label48_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label48.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub Label49_Click()
    If Label49 <> "" Then
        SeleccDia = Label49.Caption
        Call Fecha_Seleccionada
    End If
End Sub

Private Sub Label49_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Marcar = Label49.Tag
    Call Recorrer_Restaurar_Etiquetas_Fechas
    Call Recorrer_Marcar_Etiquetas_Fechas
End Sub
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    
    Me.Top = poscArriba
    Me.Left = poscIzquierda
'    Me.Left = poscIzquierda
    
    Call Limpiar_Etiquetas_Frame1
    Call Asignar_Tag
    Call Configurar_Año
    Call Configurar_Meses
    Call Año_Mes_Actual
    
    With Label50
        .Caption = FormatDateTime(Date, vbLongDate)
        .Font.Bold = True
        .Font.Underline = True
    End With
    
    Call Inicio_Del_Mes
    
    CommandButton4.Cancel = True
    
    With Me
        .StartUpPosition = 2
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 
 pdiaMes = Empty
 diaSemana = Empty
 nFecha = Empty
 Marcar = Empty
SeleccDia = Empty
SeleccMes = Empty
SeleccAño = Empty
SeleccionFecha = Empty
poscArriba = Empty
poscIzquierda = Empty

End Sub



