VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchFiles 
   Caption         =   "Busqueda de expedientes"
   ClientHeight    =   9600.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9705.001
   OleObjectBlob   =   "frmSearchFiles.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmSearchFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDeleteFiles_Click()
    Dim numberFiles As String
    Dim answer As VbMsgBoxResult
    Dim tableName As String
    Dim wasDeleted As Boolean
    
    On Error GoTo Catch
    
    tableName = "reversion"
    
    With ListBox1
        If .ListIndex <> -1 Then
            numberFiles = .list(.ListIndex, 0)
            answer = MsgBox("Se va a eliminar el siguiente expediente con número " & numberFiles & vbCrLf & "¿Desea continuar?", vbYesNo + vbExclamation + vbDefaultButton2)
            
            If answer = vbYes Then
                wasDeleted = storage.deleteFilesForNumber(numberFiles, tableName)
                If wasDeleted Then
                    MsgBox "Expediente número " & numberFiles & " fue eliminado."
                    Call btnSearchFiles_Click
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
        .ColumnCount = 3
    End With
    
    partida = TextBox2.Text
    expediente = TextBox3.Text
    anio = VBA.IIf(IsNumeric(TextBox18.Text), TextBox18.Text, Null)
    
    Call storage.filterByMultipleCriteria(list, partida, anio, expediente)
    Label14.Caption = "Expedientes encontrados " & list.ListCount
End Sub

Private Sub btnViewFiles_Click()

    Dim numberFiles As String
    Dim list As MSForms.listBox
    Dim frame As MSForms.frame
    
    On Error GoTo Catch
    
    Set list = frmViewDetailFiles.ListBox1
    Set frame = frmViewDetailFiles.Frame1
    
    With ListBox1
        If .ListIndex <> -1 Then
            numberFiles = .list(.ListIndex, 0)
            Call storage.getFilesForNumber(list, frame, numberFiles)
            frmViewDetailFiles.Show
        Else
            MsgBox "Debe seleccionar un elemento de la lista"
        End If
    End With
    
    Exit Sub
    
Catch:
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
End Sub

Private Sub UserForm_Initialize()

    If privil = "Usuario" Then btnDeleteFiles.Visible = False
    
End Sub
