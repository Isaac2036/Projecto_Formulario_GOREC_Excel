VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBusquedaAdmin 
   Caption         =   "Busqueda de expedientes"
   ClientHeight    =   11115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12945
   OleObjectBlob   =   "frmBusquedaAdmin.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmBusquedaAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Dim partida As String
    Dim expediente As String
    Dim anio As Integer
    
    partida = TextBox2.Text
    expediente = TextBox3.Text
    anio = VBA.IIf(IsNumeric(TextBox18.Text), TextBox18.Text, 0)
    
    Call storage.searchExpediente(Me.ListBox1, partida, expediente, anio)
    
End Sub
