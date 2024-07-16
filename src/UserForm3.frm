VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    frmInsertData.Show
End Sub

Private Sub CommandButton2_Click()
    frmVerExpediente.Show
End Sub

Private Sub CommandButton4_Click()
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    MsgBox CloseMode
End Sub
