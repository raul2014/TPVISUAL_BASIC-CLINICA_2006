Attribute VB_Name = "mdlMain"
Option Explicit

Public Sub main()

' Manejo de errores
On Error GoTo error

'abre el formulario de Apertura
frmSplash.Show

' Creacion de un objeto Conexion
Set conn = New ADODB.Connection
With conn
    'uso el 4.0 para access 2000
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    ' Determinacion de la base de datos a usar
    .ConnectionString = "Data source=" & App.Path & "\datos\base de datos2.mdb"
End With

'Apertura de una conexión
conn.Open

error:
'Si hubo error...
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    End
End If

End Sub

