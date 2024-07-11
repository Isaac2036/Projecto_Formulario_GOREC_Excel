VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInsertData 
   Caption         =   "Insertar Expedientes"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16680
   OleObjectBlob   =   "frmInsertData.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmInsertData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub chargePrev()
    
    Dim frmViewP As UserForm
    
    Set frmViewP = frmViewPrev
    
     With frmViewP.ListBox1
        .ColumnCount = 4
        .ColumnWidths = "150"
        
        .AddItem ""
        .AddItem "etapa : "
        .List(.ListCount - 1, 1) = UCase(Me.ComboBox1.Value)
        
        .AddItem "Serie documental/unidad : "
        .List(.ListCount - 1, 1) = UCase(Me.ComboBox2.Value)
        
        .AddItem "Uso : "
        .List(.ListCount - 1, 1) = UCase(Me.ComboBox3.Value)
        
        .AddItem "Estado : "
        .List(.ListCount - 1, 1) = UCase(Me.ComboBox4.Value)
        
        .AddItem "Proyecto y/o subestado : "
        .List(.ListCount - 1, 1) = UCase(Me.ComboBox5.Value)
        
        .AddItem ""
        .AddItem "********************"
        .AddItem ""
        
        .AddItem "N° de partida : "
        .List(.ListCount - 1, 1) = UCase(TextBox1.Value)
        
        .AddItem "Resolución : "
        .List(.ListCount - 1, 1) = UCase(TextBox2.Value)
        
        .AddItem "Exp/Hoja de ruta : "
        .List(.ListCount - 1, 1) = UCase(TextBox3.Value)
        
        .AddItem "Año de expediente : "
        .List(.ListCount - 1, 1) = ComboBox9.Value
        
        .AddItem "Administrado(s) : "
        .List(.ListCount - 1, 1) = UCase(TextBox5.Value)
        
        .AddItem "Dni : "
        .List(.ListCount - 1, 1) = TextBox6.Value
        
        .AddItem "AA.HH/Pueblo/Zona : "
        .List(.ListCount - 1, 1) = UCase(TextBox7.Value)
        
        .AddItem "Sector : "
        .List(.ListCount - 1, 1) = UCase(TextBox8.Value)
        .List(.ListCount - 1, 2) = "Barrio : "
        .List(.ListCount - 1, 3) = UCase(TextBox9.Value)
        
        .AddItem "Grupo Res. : "
        .List(.ListCount - 1, 1) = UCase(TextBox10.Value)
        .List(.ListCount - 1, 2) = "Mz : "
        .List(.ListCount - 1, 3) = UCase(TextBox11.Value)
        
        .AddItem "Lote : "
        .List(.ListCount - 1, 1) = UCase(TextBox12.Value)
        
        .AddItem "Asunto/Ultimo doc. p. : "
        .List(.ListCount - 1, 1) = UCase(TextBox13.Value)
        
        .AddItem ""
        .AddItem "********************"
        .AddItem ""
        
        .AddItem "N° folio : "
        .List(.ListCount - 1, 1) = TextBox14.Value
        
        .AddItem "Paquete : "
        .List(.ListCount - 1, 1) = ComboBox6.Value
        
        .AddItem "Ubicación exp. : "
        .List(.ListCount - 1, 1) = UCase(TextBox15.Value)
        
        .AddItem "Obsevación. : "
        .List(.ListCount - 1, 1) = UCase(TextBox16.Value)
        
        .AddItem "Profesional a cargo : "
        .List(.ListCount - 1, 1) = UCase(ComboBox7.Value)
        
        .AddItem "Fecha de última actualización. : "
        .List(.ListCount - 1, 1) = TextBox17.Value
        
        .AddItem "Rubro/ comercio/ actividad : "
        .List(.ListCount - 1, 1) = TextBox18.Value
        
        .AddItem "Area : "
        .List(.ListCount - 1, 1) = UCase(ComboBox8.Value)
        
        .AddItem "Contacto : "
        .List(.ListCount - 1, 1) = UCase(TextBox19.Value)
        
        .AddItem "Metro : "
        .List(.ListCount - 1, 1) = TextBox20.Value & "M2"
        
    End With
        
    
End Sub

Private Sub CommandButton13_Click()

    If Not utils.isInputEmpty(Me) Then
        
        Call chargePrev
        frmViewPrev.Show
        
    End If
    
    
End Sub
Private Sub CommandButton16_Click()

    Call utils.clearInputs(Me)
    Call utils.configFrm(Me)
    
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

Private Sub TextBox11_Change()

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
     Call utils.configFrm(Me)
End Sub
