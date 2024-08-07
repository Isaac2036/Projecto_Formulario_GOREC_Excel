Attribute VB_Name = "utils"
Function isInputEmpty(frm As UserForm) As Boolean

    Dim backColor As Long
    Dim status As Boolean
    
    status = False
    backColor = -2147483643
    
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            ctrl.backColor = backColor
            If ctrl.Value = Empty Then
                MsgBox "Por favor complete el campo " & ctrl.name
                ctrl.SetFocus
                ctrl.backColor = vbRed
                status = True
                Exit For
            End If
        ElseIf TypeOf ctrl Is MSForms.ComboBox Then
            ctrl.backColor = backColor
            If ctrl.ListIndex = -1 Then
                MsgBox "Por favor selecione un item de la lista"
                ctrl.SetFocus
                ctrl.backColor = vbRed
                status = True
                Exit For
            End If
        End If
    Next ctrl
    
    isInputEmpty = status
    
End Function
Sub clearInputs(frm As UserForm)
        
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is MSForms.TextBox And ctrl.Enabled Then
           ctrl.Value = Empty
        ElseIf TypeOf ctrl Is MSForms.ComboBox And ctrl.Enabled Then
           ctrl.ListIndex = -1
        End If
    Next ctrl

End Sub
Sub configFrm(frm As UserForm)

    'frame1
    With frm
        .ComboBox1.TabIndex = 0
        .ComboBox2.TabIndex = 1
        .ComboBox3.TabIndex = 2
        .ComboBox4.TabIndex = 3
        .ComboBox5.TabIndex = 4
    End With
    
    With frm.Frame1
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
    With frm
        .TextBox1.TabIndex = 5
        .TextBox2.TabIndex = 6
        .TextBox3.TabIndex = 7
        .ComboBox9.TabIndex = 8
        .TextBox21.TabIndex = 9
        .TextBox5.TabIndex = 10
        .TextBox6.TabIndex = 11
        .TextBox7.TabIndex = 12
        .TextBox8.TabIndex = 13
        .TextBox9.TabIndex = 14
        .TextBox10.TabIndex = 15
        .TextBox11.TabIndex = 16
        .TextBox12.TabIndex = 17
    End With
    
    With frm.Frame2
        With .ComboBox9
            For i = 2000 To 2024
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
    With frm
        .TextBox13.TabIndex = 18
        .TextBox14.TabIndex = 19
        .ComboBox6.TabIndex = 20
        .TextBox15.TabIndex = 21
        .TextBox16.TabIndex = 22
        .ComboBox7.TabIndex = 23
        .TextBox17.TabIndex = 24
        .TextBox18.TabIndex = 25
        .ComboBox8.TabIndex = 26
        .TextBox19.TabIndex = 27
        .TextBox20.TabIndex = 28
    End With
    
    With frm.Frame3
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




