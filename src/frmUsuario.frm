VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUsuario 
   Caption         =   "INICIAR SESIÓN"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4335
   OleObjectBlob   =   "frmUsuario.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

Dim nombre As String
Dim contras As String
Dim passw As String
Dim privil As String
Dim numhojas As Integer

On Error GoTo LineaError

Application.ScreenUpdating = False

nombre = ComboBox1.Text 'Variable que lee el nombre del usuario
contras = TextBox1.Text 'Variable que lee la contraseña

If nombre = "" Then 'Si no se coloca el nombre del usuario se muestra un mensaje de error

    MsgBox "DEBE INGRESAR EL NOMBRE DEL USUARIO", vbCritical, "INICIAR SESIÓN"

Else

    'Fórmula para buscar la contraseña en la tabla de usuarios en base a nombre de usuario
    
    passw = Application.WorksheetFunction.VLookup(nombre, [Tabla_Usuarios], 2, False)
    privil = Application.WorksheetFunction.VLookup(nombre, [Tabla_Usuarios], 3, False)
    
    If contras = passw Then 'Si la contraseña ingresada en el textbox es igual al password encontrado pasa lo siguiente:
        
        If privil = "Administrador" Then
            Unload frmUsuario ' Se cierra el formulario
            frmBienvenidosAdmin.Show
        Else
            If privil = "Total" Then
                Unload frmUsuario ' Se cierra el formulario
                frmBienvenidosTotal.Show
            Else
                Unload frmUsuario 'Se cierra el formulario
                frmBienvenidosUsuarios.Show
            End If
        End If
        
    Else
        
        MsgBox "CONTRASEÑA INCORRECTA"
        
    End If

End If

Application.ScreenUpdating = True

Exit Sub
LineaError:
TextBox1.Text = ""
MsgBox "EL NOMBRE QUE HA COLOCADO NO EXISTE", vbCritical, "INICIAR SESIÓN"

End Sub
