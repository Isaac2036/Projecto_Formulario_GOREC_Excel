VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmViewPrev 
   Caption         =   "UserForm4"
   ClientHeight    =   8220.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6240
   OleObjectBlob   =   "frmViewPrev.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmViewPrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
      If storage.insertNewRecord(frmInsertData) Then
            
        MsgBox "Datos ingresados correctamente."
        Call utils.clearInputs(frmInsertData)
        Call utils.configFrm(frmInsertData)
        Unload Me

    Else
'      nada
    End If

End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub
    
Private Sub UserForm_Initialize()
    
End Sub
