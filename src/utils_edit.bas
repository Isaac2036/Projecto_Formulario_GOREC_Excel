Attribute VB_Name = "utils_edit"
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
        .cmBox1.TabIndex = 0
        .cmBox2.TabIndex = 1
        .cmBox3.TabIndex = 2
        .cmBox4.TabIndex = 3
        .cmBox5.TabIndex = 4
    End With
    
    With frm.Frame1
        With .cmBox1
                .AddItem "Pendiente"
                .AddItem "Asignado"
                .AddItem "Evaluando"
                .ListIndex = 0
        End With
        
        With .cmBox2
            .AddItem "reversion"
            .ListIndex = 0
            .Enabled = False
        End With
        
        With .cmBox3
            .AddItem "vivienda"
            .AddItem "comercial"
            .AddItem "publico"
            .ListIndex = 0
        End With
        
         With .cmBox4
            .AddItem "pendiente de asignación"
            .AddItem "pendiente de revisión"
            .AddItem "pendiente notificación"
            .AddItem "asignado especialista"
            .AddItem "asignado jefatura"
            .AddItem "completado"
            .ListIndex = 0
        End With
        
        With .cmBox5
            .AddItem "PECP"
            .ListIndex = 0
            .Enabled = False
        End With
    End With
    
    'frame 2
    With frm
        .txtBox1.TabIndex = 5
        .txtBox2.TabIndex = 6
        .txtBox3.TabIndex = 7
        .cmBox6.TabIndex = 8
        .txtBox4.TabIndex = 9
        .txtBox5.TabIndex = 10
        .txtBox6.TabIndex = 11
        .txtBox7.TabIndex = 12
        .txtBox8.TabIndex = 13
        .txtBox9.TabIndex = 14
        .txtBox10.TabIndex = 15
        .txtBox11.TabIndex = 16
    End With
    
    With frm.Frame2
        With .cmBox6
            For i = 2000 To 2024
                .AddItem i
            Next i
            
            .Value = 2024
            .MaxLength = 4
        End With
        
        .txtBox9.MaxLength = 4
        .txtBox10.MaxLength = 2
        .txtBox11.MaxLength = 3
        
    End With

'    Frame 3
    With frm
        .txtBox12.TabIndex = 17
        .txtBox13.TabIndex = 18
        .cmBox7.TabIndex = 19
        .txtBox14.TabIndex = 20
        .txtBox15.TabIndex = 21
        .cmBox8.TabIndex = 22
        .txtBox16.TabIndex = 23
        .txtBox17.TabIndex = 24
        .cmBox9.TabIndex = 25
        .TextBox21.TabIndex = 26
        .txtBox18.TabIndex = 27
        .txtBox19.TabIndex = 28
    End With
    
    With frm.Frame3
        With .cmBox7
            For i = 1 To 9
                .AddItem "GABETA " & i
            Next i
            .ListIndex = 0
        End With
        
        With .txtBox15
            .WordWrap = True
            .MultiLine = True
            .EnterKeyBehavior = True
            .ScrollBars = fmScrollBarsVertical
        End With
        
        With .cmBox8
            .RowSource = "nombre_abogados"
        End With
        
         With .cmBox9
            .AddItem "OGP"
            .ListIndex = 0
            .Enabled = False
        End With
        
        .txtBox13.MaxLength = 4
        .txtBox16.Value = Date
        
    End With
    
End Sub
