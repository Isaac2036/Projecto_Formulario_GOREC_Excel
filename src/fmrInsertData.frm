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

Sub configFrm()

'    Dim marginLeft As Integer
'    Dim withFrame As Integer
'
'    marginLeft = 12
'    withFrame = 432
    
    'Scroll bar
'    With Me
'        .ScrollBars = fmScrollBarsVertical
'        .ScrollHeight = 900
'    End With
    
    'Frame 1
    With Me.Frame1
'        .Height = 174
'        .Top = marginLeft
'        .Width = withFrame
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
'        .Height = 284
'        .Top = 194
'        .Left = marginLeft
'        .Width = withFrame
        .TextBox4.Value = Format$(Now(), "yyyy")
        .Enabled = False
    End With

'    Frame 3
    With Me.Frame3
'        .Height = 284
'        .Top = 484
'        .Left = marginLeft
'        .Width = withFrame
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

Private Sub CommandButton13_Click()
    Call validateInputs
End Sub

Private Sub Frame3_Click()

End Sub

Private Sub Frame6_Click()

End Sub

Private Sub Label412_Click()
'    frm_calendario.Show
End Sub

Private Sub Label420_Click()

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
    
    
End Sub

