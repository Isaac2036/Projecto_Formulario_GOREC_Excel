Attribute VB_Name = "Módulo1"
Sub test_reversion()
    
    Dim r As New Reversion
    
    With r
        .etapa = "ghdghghdgsh"
        .serie = "45ddsdjs"
        .dnis = 45
    End With
    
End Sub
Sub test_join()
    
    Dim dnis As New Collection
    
    With dnis
        .Add 45899974
        .Add 45898545
    End With
    
   Debug.Print Join(dnis, ",")
    
End Sub
