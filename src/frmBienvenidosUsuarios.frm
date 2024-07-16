VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBienvenidosUsuarios 
   Caption         =   "Gobierno Regional del Callao"
   ClientHeight    =   9735.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12960
   OleObjectBlob   =   "frmBienvenidosUsuarios.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmBienvenidosUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmButt1_Click()
    
frmBusquedaUsuarios.Show
    
End Sub

Private Sub cmButt2_Click()

frmInsertData.Show

End Sub

Private Sub CommandButton1_Click()
    ' Termina cualquier compilaci�n en curso
    Application.OnTime Now, "TerminarCompilacion"

    ' Ir a un proyecto VBA espec�fico
    ' Nota: Necesitas conocer el nombre del proyecto y m�dulo al que quieres ir
    ' Este ejemplo asume que el proyecto se llama "MiProyecto" y el m�dulo se llama "MiModulo"

    ' Abre el editor de VBA
    Application.VBE.MainWindow.Visible = True

    ' Selecciona el proyecto y el m�dulo
    Dim vbProj As VBIDE.VBProject
    Set vbProj = Application.VBE.VBProjects("ModeloPrevioReversion.xlsm")
    
    Dim vbMod As VBIDE.CodeModule
    Set vbMod = vbProj.VBComponents("storage").CodeModule

    ' Navega al inicio del m�dulo
    Application.VBE.ActiveCodePane.CodeModule.CodePane.Show
End Sub

Sub TerminarCompilacion()
    ' C�digo para terminar cualquier compilaci�n en curso
    ' Esto depender� de c�mo se est� ejecutando la compilaci�n en tu entorno espec�fico
    ' Podr�as incluir l�gica espec�fica aqu� para asegurarte de que se detenga adecuadamente
    MsgBox "Compilaci�n terminada"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

Application.Visible = True
Unload Me

End Sub

Private Sub Userform_Terminate()
    If Application.Workbooks.Count = 1 Then
        Application.Quit
    Else
        ThisWorkbook.Close True
    End If
End Sub


