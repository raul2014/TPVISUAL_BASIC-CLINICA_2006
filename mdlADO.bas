Attribute VB_Name = "mdlADO"
Option Explicit

'conexion: Recibe una conexion abierta, sobre la cual
'crea el objeto command
'query: Es la consulta SQL
'rec: Recibe un recordset por referencia y lo modifica
'según la consulta efectuada


Public Sub consultasql(conexion As ADODB.Connection, query As String, rec As ADODB.Recordset)
'Manejo de Error    'Este procedimiento permite traer registros de la base de datos  atraves del objeto connection,la consulta y el objeto record set
On Error GoTo error

'Creo los Objetos
Dim cmdDatos As New ADODB.Command
With cmdDatos
    ' Enlazo Command con objeto Connection
    Set .ActiveConnection = conexion
    ' Determino el texto de la consulta SQL
    .CommandText = query
    .CommandType = adCmdText
End With

With rec
   .CursorType = adOpenKeyset
   .LockType = adLockOptimistic
   .CursorLocation = adUseClient
End With

' Abro el Recordset teniendo en cuenta la consulta que posee command
rec.Open cmdDatos


error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub


