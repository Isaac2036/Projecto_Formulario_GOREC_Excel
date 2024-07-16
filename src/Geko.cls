VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Geko"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private cnn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public strConnection As String
Private version As String
Const AppName As String = "Geko"

Public Property Get name() As String

    name = AppName
    
End Property

Public Property Get ver() As String

    version = AppName + " 1.1"
    ver = version
    
End Property


Public Function openConnection() As String
    cnn.ConnectionString = strConnection
    openConnection = cnn
    Debug.Print Me.name & ": Conexion abierta " & Now()
End Function

Sub executeCommand(ByVal sql As String)
    
    cnn = openConnection()
    cnn.Open
    
    With rs
          .CursorLocation = adUseServer
          .Open Source:=sql, ActiveConnection:=cnn
    End With
    
    Debug.Print Me.name & ": Consulta ejecutada " & Now()
    
    freeMemory
    
End Sub
Sub executeCommand2(ByVal sql As String)
    
    cnn = openConnection()
    cnn.Open
    
    With rs
          .CursorLocation = adUseServer
          .Open Source:=sql, ActiveConnection:=cnn
    End With
    
    Debug.Print Me.name & ": Consulta ejecutada " & Now()
    
End Sub

Public Function showString(ByVal sql As String) As String
    
    cnn = openConnection()
    cnn.Open
    
    With rs
            .CursorLocation = adUseServer
            .Open Source:=sql, ActiveConnection:=cnn
      End With
    
    showString = rs.getString
    
    Debug.Print Me.name & ": Consulta ejecutada " & Now()
    
    freeMemory
    
End Function

Public Sub showRecordset(ByVal sql As String)

    cnn = openConnection()
    cnn.Open

    With rs
        .CursorLocation = adUseServer
        .Open Source:=sql, ActiveConnection:=cnn
    End With
    
    Debug.Print Me.name & ": Consulta ejecutada " & Now()
    
End Sub

Public Function ExitsRecord(ByVal sql As String) As Boolean

    executeCommand2 sql
    
    If rs.BOF And rs.EOF Then
        ExitsRecord = False
    Else
        ExitsRecord = True
    End If
    
    freeMemory
    
End Function

Public Function getRecordset(ByVal sql As String) As Long
      
      Dim Cuenta!
      
      cnn = openConnection()
      cnn.Open
      
      With rs
            .CursorLocation = adUseClient
            .Open Source:=sql, ActiveConnection:=cnn
            
            Cuenta = rs.RecordCount
            
      End With
      
      freeMemory
      
      getRecordset = Cuenta
      
End Function

Sub downloadData(ByVal sql As String, ByVal name_sheet As String)

      Dim rspta%
      
      rspta = MsgBox("Descargar base de datos?", vbInformation + vbYesNo + vbDefaultButton2, AppName)
      
      If rspta = 6 Then
            executeCommand2 sql
                  Workbooks.Add
                  ActiveSheet.name = name_sheet
                        
                        For i = 0 To rs.Fields.Count - 1
                              Cells(1, i + 1) = rs.Fields(i).name
                        Next i
                        
                        Range("a2").CopyFromRecordset rs
                        Rows(1).Font.Bold = True
                  MsgBox "Tarde un poco pero ya esta tu libro", vbInformation, AppName
            freeMemory
      End If
      

End Sub
Sub LoadCtrlList(ByVal sql As String, ByVal index, ByVal ctrl As MSForms.control)
      '//**Procedimiento que realiza la función de cargar un ListxBox o ComboBox de un formulario
      '//***Esta procedimiento es dependiente de la función queryRecordset(),pero no tienes que llamar el procedimiento lo hará por ti.
      '//****ARGUMENTO_1:archivo_accdb=Nombe del archivo Access con extensión accdb tipo cadena de texto
      '//****ARGUMENTO_2:tabla=consulta SQL de la quiere obtener la cuenta o nombre de la tabla
      '//****ARGUMENTO_3:index=campo que desea cargar en nuestro control, se le puede pasar como número o referenciando directamente el nombre del campo
      '//****ARGUMENTO_4:ctrl=se le debe pasar un control del tipo ListBox o ComboBox
      '//****Ej1:LoadCtrlList "MiPrimerBaseDB.accdb","select * from miTabla",0,Me.ComboBox
      '//****Ej2:LoadCtrlList "MiPrimerBaseDB.accdb","mi primer tabla","mi_campo",UserForm1.ListBox
     
      executeCommand2 sql
            
            If TypeOf ctrl Is MSForms.ComboBox Or TypeOf ctrl Is MSForms.ListBox Then
                  ctrl.Clear
                  
                  rs.MoveFirst
                  Do While Not rs.EOF
                        ctrl.AddItem rs.Fields(index)
                        rs.MoveNext
                    
                  Loop
            End If
            
            freeMemory
            '**Observación para el desarrollador:Al declarar el parámetro index como Byte y ejecutarla el procedimiento muestra error,
            '**pero si lo declaramos como Integer ejecuta sin problemas.
End Sub

Sub freeMemory()
      
      If rs.State <> adStateClosed Then
            rs.Close
      End If

      cnn.Close
      Set cnn = Nothing
      Set rs = Nothing
      Debug.Print Me.name & ": Conexión cerrada " & Now()
End Sub

