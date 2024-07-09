VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVerExpediente 
   Caption         =   "Expedientes"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13275
   OleObjectBlob   =   "frmVerExpediente.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmVerExpediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lstView1_Paint()
    Dim i As Integer
    Dim itemText1 As String
    Dim fontSize1 As Integer
    
    For i = 0 To lstView1.ListCount - 1
        itemText1 = lstView1.List(i)
        fontSize1 = GetFontSize1(i) ' función que devuelve el tamaño de fuente para cada item
        lstView1.Font.Size = fontSize1
        lstView1.CurrentX = 0
        lstView1.CurrentY = i * (fontSize1 + 2)
        lstView1.Print itemText1
    Next i
End Sub

Private Function GetFontSize1(index As Integer) As Integer
    ' función que devuelve el tamaño de fuente para cada item
    ' por ejemplo, puedes utilizar un array de tamaños de fuente
    Dim fontSizes1() As Integer
    fontSizes1 = Array(10, 12, 14, 16, 18, 20, 20, 20, 20)
    GetFontSize1 = fontSizes1(index Mod UBound(fontSizes1))
End Function

Private Sub lstView2_Paint()
    Dim i As Integer
    Dim itemText2 As String
    Dim fontSize2 As Integer
    
    For i = 0 To lstView2.ListCount - 1
        itemText2 = lstView2.List(i)
        fontSize2 = GetFontSize2(i) ' función que devuelve el tamaño de fuente para cada item
        lstView2.Font.Size = fontSize2
        lstView2.CurrentX = 0
        lstView2.CurrentY = i * (fontSize2 + 2)
        lstView2.Print itemText2
    Next i
End Sub

Private Function GetFontSize2(index As Integer) As Integer
    ' función que devuelve el tamaño de fuente para cada item
    ' por ejemplo, puedes utilizar un array de tamaños de fuente
    Dim fontSizes2() As Integer
    fontSizes2 = Array(10, 12, 14, 16, 18, 20, 20, 20, 20, 20, 20, 20, 20)
    GetFontSize2 = fontSizes2(index Mod UBound(fontSizes2))
End Function

Private Sub lstView3_Paint()
    Dim i As Integer
    Dim itemText3 As String
    Dim fontSize3 As Integer
    
    For i = 0 To lstView3.ListCount - 1
        itemText3 = lstView3.List(i)
        fontSize3 = GetFontSize3(i) ' función que devuelve el tamaño de fuente para cada item
        lstView3.Font.Size = fontSize3
        lstView3.CurrentX = 0
        lstView3.CurrentY = i * (fontSize2 + 2)
        lstView3.Print itemText3
    Next i
End Sub

Private Function GetFontSize3(index As Integer) As Integer
    ' función que devuelve el tamaño de fuente para cada item
    ' por ejemplo, puedes utilizar un array de tamaños de fuente
    Dim fontSizes3() As Integer
    fontSizes3 = Array(10, 12, 14, 16, 18, 20, 20, 20, 20, 20)
    GetFontSize3 = fontSizes3(index Mod UBound(fontSizes2))
End Function

Private Sub Frame6_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub UserForm_Initialize()
    With Me.lstView1
        .ColumnWidths = 9
    End With
    
    With Me.lstView2
        .ColumnWidths = 10
    End With
        
    With Me.lstView2
        .ColumnWidths = 9
    End With
    
    Call storage.viewNewRecord(Me)
    
End Sub




