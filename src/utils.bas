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
                MsgBox "Por favor complete el campo " & ctrl.Name
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
            .AddItem "pendiente de asignaci�n"
            .AddItem "pendiente de revisi�n"
            .AddItem "pendiente notificaci�n"
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
        .TextBox5.TabIndex = 9
        .TextBox6.TabIndex = 10
        .TextBox7.TabIndex = 11
        .TextBox8.TabIndex = 12
        .TextBox9.TabIndex = 13
        .TextBox10.TabIndex = 14
        .TextBox11.TabIndex = 15
        .TextBox12.TabIndex = 16
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




