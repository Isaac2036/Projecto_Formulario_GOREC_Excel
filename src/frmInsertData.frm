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

Private Sub CommandButton13_Click()

    If Not utils.isInputEmpty(Me) Then
        If storage.insertNewRecord(Me) Then
            MsgBox "Datos ingresados correctamente."
            Call utils.clearInputs(Me)
            Call utils.configFrm(Me)
            
            Me.Hide
            frmVerExpediente.Show
            
        Else
          'nada
        End If
        
    End If
    
End Sub

Private Sub CommandButton16_Click()

    Call utils.clearInputs(Me)
    Call utils.configFrm(Me)
    
End Sub

Private Sub Frame1_Click()

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
