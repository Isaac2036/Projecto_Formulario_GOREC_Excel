VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBusquedaAdmin 
   Caption         =   "Busqueda de expedientes"
   ClientHeight    =   9600.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9705.001
   OleObjectBlob   =   "frmBusquedaAdmin.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmBusquedaAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDeleteFiles_Click()
    Dim id As Integer
    Dim answer As VbMsgBoxResult
    Dim tableName As String
    Dim wasDeleted As Boolean
    
    On Error GoTo Catch
    tableName = "reversion"
    With ListBox1
        If .ListIndex <> -1 Then
            id = .list(.ListIndex, 0)
            answer = MsgBox("Se va a eliminar el siguiente expediente con el id " & id & vbCrLf & "¿Desea continuar?", vbYesNo + vbExclamation + vbDefaultButton2)
            
            If answer = vbYes Then
                wasDeleted = storage.deleteFilesForID(id, tableName)
                If wasDeleted Then
                    MsgBox "Expediente con el id " & id & " eliminado."
                End If
            Else
                MsgBox "Operación cancelada."
            End If
        Else
            MsgBox "Debe seleccionar un elemento de la lista"
        End If
    End With
    Exit Sub
Catch:
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
    
End Sub

Private Sub btnSearchFiles_Click()
    Dim partida As String
    Dim expediente As String
    Dim anio As Variant
    Dim list As MSForms.listBox
    
    Set list = Me.ListBox1
    
    With list
        .Clear
        .ColumnCount = 4
    End With
    
    partida = TextBox2.Text
    expediente = TextBox3.Text
    anio = VBA.IIf(IsNumeric(TextBox18.Text), TextBox18.Text, Null)
    
    Call storage.filterByMultipleCriteria(list, partida, anio, expediente)
    Label14.Caption = "Expedientes encontrados " & list.ListCount
End Sub

Private Sub UserForm_Click()

End Sub
