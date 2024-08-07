VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInsertData 
   Caption         =   "Insertar Expedientes"
   ClientHeight    =   8145
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
        .ColumnCount = 1
        .ColumnWidths = "120"
        
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(Me.ComboBox1.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(Me.ComboBox2.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(Me.ComboBox3.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(Me.ComboBox4.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(Me.ComboBox5.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(TextBox1.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(TextBox2.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(TextBox3.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = ComboBox9.Value
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = TextBox21.Value
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(TextBox5.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = TextBox6.Value
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(TextBox7.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(TextBox8.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(TextBox9.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(TextBox10.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(TextBox11.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(TextBox12.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(TextBox13.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = TextBox14.Value
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = ComboBox6.Value
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(TextBox15.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(TextBox16.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(ComboBox7.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = TextBox17.Value
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = TextBox18.Value
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(ComboBox8.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = UCase(TextBox19.Value)
        
        .AddItem ""
        .AddItem ""
        .list(.ListCount - 1, 0) = TextBox20.Value & " M2"
        
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

Private Sub Frame6_Click()

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
     Call utils.configFrm(Me)
End Sub
