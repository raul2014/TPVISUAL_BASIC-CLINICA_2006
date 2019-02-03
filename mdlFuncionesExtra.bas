Attribute VB_Name = "mdlFuncionesExtra"
Option Explicit

'Procedimiento para llenar un combo con todas las provincias
'Parametros:
'cmbProvincia: es un combo que recibe por referencia y lo carga
'con todas las provincias de la base de datos

Public Sub cargaComboProvincias(cmbProvincia As ComboBox)

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'limpia el combo
cmbProvincia.Clear

'Hago la consulta
qbusca = " SELECT P.DESCRIPCION AS DESCPROV" & _
         " FROM PROVINCIAS AS P" & _
         " ORDER BY P.DESCRIPCION"
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
While Not rstDatos.EOF
    'carga los items al combo
    cmbProvincia.AddItem rstDatos!DESCPROV
    'avanza al registro siguiente
    rstDatos.MoveNext
Wend

'si no esta vacio el combo pone el primero
If cmbProvincia.ListCount <> 0 Then cmbProvincia.ListIndex = 0

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub

'Funcion para buscar el Codigo de una provincia
'Parametros:
'prov: el nombre de la la provincia
'retorno: el codigo de dicha localidad o -1 si no la encontro

Public Function consultaCodProvincia(prov As String) As Long

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'Hago la consulta
qbusca = " SELECT COD_PROV " & _
         " FROM  PROVINCIAS AS P " & _
         " WHERE P.DESCRIPCION='" & prov & "'"
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
If Not rstDatos.EOF Then
    'retorna
    consultaCodProvincia = rstDatos!COD_PROV
    Exit Function
End If

consultaCodProvincia = -1

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Function
End If
End Function

'Procedimiento para llenar un combo con todas las localidades
'Parametros:
'cmbLoc: es un combo que recibe por referencia y lo carga
'con todas las localidades de la base de datos

Public Sub cargaComboLocalidades(cmbLoc As ComboBox)

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'limpia el combo
cmbLoc.Clear

'Hago la consulta
qbusca = " SELECT L.DESCRIPCION AS DESCLOC" & _
         " FROM LOCALIDADES AS L" & _
         " ORDER BY L.DESCRIPCION"
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
While Not rstDatos.EOF
    'carga los items al combo
    cmbLoc.AddItem rstDatos!DESCLOC
    'avanza al registro siguiente
    rstDatos.MoveNext
Wend

'si no esta vacio el combo pone el primero
If cmbLoc.ListCount <> 0 Then cmbLoc.ListIndex = 0

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub

'LA UTILIZO PARA EL FORMULARIO DE OBSOCIALES
            

'Funcion para buscar el Codigo de una localidad
'Parametros:
'loc: el nombre de la localidad
'retorno: el codigo de dicha localidad o -1 si no la encontro

Public Function consultaCodLocalidad(loc As String) As Long

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'Hago la consulta
qbusca = " SELECT COD_LOC " & _
         " FROM  LOCALIDADES AS L " & _
         " WHERE L.DESCRIPCION='" & loc & "'"
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
If Not rstDatos.EOF Then
    'retorna
    consultaCodLocalidad = rstDatos!COD_LOC
    Exit Function
End If

consultaCodLocalidad = -1

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Function
End If
End Function

'Procedimiento para llenar un combo con todas las obras sociales
'Parametros:
'cmbObSocial: es un combo que recibe por referencia y lo carga
'con todas las localidades de la base de datos

Public Sub cargaComboObrasSociales(cmbObSocial As ComboBox)

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'limpia el combo
cmbObSocial.Clear

'Hago la consulta
qbusca = " SELECT OB.RAZON_SOCIAL AS DESC_OBSOCIAL" & _
         " FROM OBRA_SOCIAL AS OB" & _
         " ORDER BY OB.RAZON_SOCIAL"
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
While Not rstDatos.EOF
    'carga los items al combo
    cmbObSocial.AddItem rstDatos!DESC_OBSOCIAL
    'avanza al registro siguiente
    rstDatos.MoveNext
Wend

'si no esta vacio el combo pone el primero
If cmbObSocial.ListCount <> 0 Then cmbObSocial.ListIndex = 0

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub

'Procedimiento para llenar un combo con todos los medicos
'Parametros:
'cmbMed: es un combo que recibe por referencia y lo carga
'con todas los medicos de la base de datos

Public Sub cargaComboMedicos(cmbMed As ComboBox)

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'limpia el combo
cmbMed.Clear

'Hago la consulta
qbusca = " SELECT MED.NOMBREM" & _
         " FROM MEDICOS AS MED" & _
         " ORDER BY MED.NOMBREM"
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
While Not rstDatos.EOF
    'carga los items al combo
    cmbMed.AddItem rstDatos!NOMBREM
    'avanza al registro siguiente
    rstDatos.MoveNext
Wend

'si no esta vacio el combo pone el primero
If cmbMed.ListCount <> 0 Then cmbMed.ListIndex = 0

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub

'Funcion para buscar el Codigo de una obra social dado la descripsion del mismo
'Parametros:
'obsocial: el nombre de la obra social
'retorno: el codigo de dicha obra social o -1 si no la encontro

Public Function consultaCodObSocial(obsocial As String) As Long

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'Hago la consulta
qbusca = " SELECT OB.COD_OBSOCIAL " & _
         " FROM  OBRA_SOCIAL AS OB " & _
         " WHERE OB.RAZON_SOCIAL='" & obsocial & "'"
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
If Not rstDatos.EOF Then
    'retorna
    consultaCodObSocial = rstDatos!COD_OBSOCIAL
    Exit Function
End If

consultaCodObSocial = -1

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Function
End If
End Function


'Funcion para buscar el Codigo de un medico
'Parametros:
'med: el nombre de la la provincia
'retorno: el codigo del medico o NULL si no la encontro

Public Function consultaCodMedico(med As String) As Variant

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'Hago la consulta
qbusca = " SELECT COD_MLEG " & _
         " FROM  MEDICOS AS M " & _
         " WHERE M.NOMBREM='" & med & "'"
         
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
If Not rstDatos.EOF Then
    'retorna
    consultaCodMedico = rstDatos!COD_MLEG
    Exit Function
End If

consultaCodMedico = Null

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Function
End If
End Function

'Procedimiento para llenar el listbox de especialidades de la tabla doctores
'Parametros:
'lbxEsp: es un listbox que recibe por referencia y lo carga
'con todas las especialidades que correpomden al medico selecionado

Public Sub llenarListboxEspecialidades(lbxEsp As ListBox)

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'limpia el listbox
lbxEsp.Clear

'Hago la consulta
qbusca = " SELECT ESP.DESCRIPCION" & _
         " FROM ESPECIALIDADES AS ESP" & _
         " ORDER BY ESP.COD_ESP"
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
While Not rstDatos.EOF
    'carga los items al combo
    lbxEsp.AddItem rstDatos!DESCRIPCION
    'avanza al registro siguiente
    rstDatos.MoveNext
Wend

'si no esta vacio el combo pone el primero
If lbxEsp.ListCount <> 0 Then lbxEsp.ListIndex = 0

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub

'Procedimiento para llenar el listbox de obras sociales de la tabla doctores
'Parametros:
'lbxObSociales: es un listbox que recibe por referencia y lo carga
'con todas las especialidades que correponden al medico selecionado

Public Sub llenarListboxObrasSociales(lbxObsociales As ListBox)

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'limpia el listbox
lbxObsociales.Clear

'Hago la consulta
qbusca = " SELECT OB.RAZON_SOCIAL" & _
         " FROM OBRA_SOCIAL AS OB" & _
         " ORDER BY OB.COD_OBSOCIAL"
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
While Not rstDatos.EOF
    'carga los items al combo
    lbxObsociales.AddItem rstDatos!RAZON_SOCIAL
    'avanza al registro siguiente
    rstDatos.MoveNext
Wend

'si no esta vacio el combo pone el primero
If lbxObsociales.ListCount <> 0 Then lbxObsociales.ListIndex = 0

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub

'Procedimiento para llenar el listbox de obras sociales de la tabla doctores
'Parametros:
'lbxObSociales: es un listbox que recibe por referencia y lo carga
'con todas las obras sociales que correponden al medico selecionado

Public Sub llenarListboxObrasSocialesSegunMedico(lbxObsociales As ListBox, codmed As Long)

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'limpia el listbox
lbxObsociales.Clear

'Hago la consulta
qbusca = " SELECT  OB.RAZON_SOCIAL" & _
         " FROM OBRA_SOCIAL AS OB,OBRASOCIAL_MEDICOS AS OBMED,MEDICOS AS M" & _
         " WHERE   M.COD_MLEG=OBMED.COD_MLEG" & _
         " AND   OBMED.COD_OBSOCIAL=OB.COD_OBSOCIAL" & _
         " AND OBMED.COD_MLEG=" & codmed & _
         " ORDER BY OB.COD_OBSOCIAL"
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
While Not rstDatos.EOF
    'carga los items al combo
    lbxObsociales.AddItem rstDatos!RAZON_SOCIAL
    'avanza al registro siguiente
    rstDatos.MoveNext
Wend

'si no esta vacio el combo pone el primero
If lbxObsociales.ListCount <> 0 Then lbxObsociales.ListIndex = 0

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub

'Procedimiento para llenar el listbox de especialidades de la tabla doctores
'Parametros:
'lbxEsp: es un listbox que recibe por referencia y lo carga
'con todas las especialidades que correponden al medico selecionado

Public Sub llenarListboxEspecialidadesSegunMedico(lbxEsp As ListBox, codmed As Long)

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'limpia el listbox
lbxEsp.Clear

'Hago la consulta
qbusca = " SELECT  E.DESCRIPCION" & _
         " FROM ESPECIALIDADES AS E,MEDICOS_ESPECIALIDAD AS MEDESP,MEDICOS AS M" & _
         " WHERE   M.COD_MLEG=MEDESP.COD_MLEG" & _
         " AND     MEDESP.COD_ESP=E.COD_ESP" & _
         " AND     MEDESP.COD_MLEG=" & codmed & _
         " ORDER BY E.COD_ESP"
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
While Not rstDatos.EOF
    'carga los items al combo
    lbxEsp.AddItem rstDatos!DESCRIPCION
    'avanza al registro siguiente
    rstDatos.MoveNext
Wend

'si no esta vacio el combo pone el primero
If lbxEsp.ListCount <> 0 Then lbxEsp.ListIndex = 0

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub

'Procedimiento para llenar el listbox de estudios de la tabla historias clinicas
'Parametros:
'lbxEstudios: es un listbox que recibe por referencia y lo carga
'con todas los estudios que correponden al paciente seleccionado

Public Sub llenarListboxEstudiosSegunPaciente(lbxEstudios As ListBox, codPaciente As Long)

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'limpia el listbox
lbxEstudios.Clear

'Hago la consulta
qbusca = " SELECT  E.DESCRIPCION" & _
         " FROM PACIENTES AS P,TURNOS AS T,ESTUDIO_TURNOS AS ESTTUR,ESTUDIOS AS E" & _
         " WHERE   T.COD_TURNO=ESTTUR.COD_TURNO" & _
         " AND     ESTTUR.COD_ESTMED=E.COD_ESTMED" & _
         " AND     P.COD_LEGP=T.COD_LEGP" & _
         " AND     T.COD_LEGP=" & codPaciente & _
         " ORDER BY E.DESCRIPCION"
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
While Not rstDatos.EOF
    'carga los items al combo
    lbxEstudios.AddItem rstDatos!DESCRIPCION
    'avanza al registro siguiente
    rstDatos.MoveNext
Wend

'si no esta vacio el combo pone el primero
If lbxEstudios.ListCount <> 0 Then lbxEstudios.ListIndex = 0

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub

'Procedimiento para llenar un combo con todos los pacientes
'Parametros:
'cmbPacientes: es un combo que recibe por referencia y lo carga
'con todas los pacientes de la base de datos

Public Sub cargaComboPacientes(cmbPacientes As ComboBox)

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'limpia el combo
cmbPacientes.Clear

'Hago la consulta
qbusca = " SELECT P.NOMBRE" & _
         " FROM PACIENTES AS P" & _
         " ORDER BY P.NOMBRE"
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
While Not rstDatos.EOF
    'carga los items al combo
    cmbPacientes.AddItem rstDatos!NOMBRE
    'avanza al registro siguiente
    rstDatos.MoveNext
Wend

'si no esta vacio el combo pone el primero
If cmbPacientes.ListCount <> 0 Then cmbPacientes.ListIndex = 0

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub

'Procedimiento para llenar un combo con todas las especialidades
'Parametros:
'cmbEsp: es un combo que recibe por referencia y lo carga
'con todas las especialidades de la base de datos

Public Sub cargaComboEspecialidades(cmbEsp As ComboBox)

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'limpia el combo
cmbEsp.Clear

'Hago la consulta
qbusca = " SELECT ESP.DESCRIPCION" & _
         " FROM ESPECIALIDADES AS ESP" & _
         " ORDER BY ESP.DESCRIPCION"
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
While Not rstDatos.EOF
    'carga los items al combo
    cmbEsp.AddItem rstDatos!DESCRIPCION
    'avanza al registro siguiente
    rstDatos.MoveNext
Wend

'si no esta vacio el combo pone el primero
If cmbEsp.ListCount <> 0 Then cmbEsp.ListIndex = 0

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub

'Funcion para buscar el Codigo de una especialidad
'Parametros:
'esp: el nombre de la especialidad
'retorno: el codigo de dicha especialidad o -1 si no la encontro

Public Function consultaCodEspecialidad(esp As String) As Long

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'Hago la consulta
qbusca = " SELECT COD_ESP " & _
         " FROM  ESPECIALIDADES AS E " & _
         " WHERE E.DESCRIPCION='" & esp & "'"
         
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
If Not rstDatos.EOF Then
    'retorna
    consultaCodEspecialidad = rstDatos!COD_ESP
    Exit Function
End If

consultaCodEspecialidad = -1

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Function
End If
End Function


'Funcion para buscar el Codigo de una paciente
'Parametros:
'paciente: el nombre del paciente
'retorno: el codigo de dicho paciente o -1 si no la encontro

Public Function consultaCodPaciente(paciente As String) As Long

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'Hago la consulta
qbusca = " SELECT COD_LEGP " & _
         " FROM  PACIENTES AS P " & _
         " WHERE P.NOMBRE='" & paciente & "'"
         
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
If Not rstDatos.EOF Then
    'retorna
    consultaCodPaciente = rstDatos!COD_LEGP
    Exit Function
End If

consultaCodPaciente = -1

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Function
End If
End Function

'Procedimiento para llenar el listbox de estudios de el formulario HISTORIAS CLINICAS
'Parametros:
'lbxEstudios: es un listbox que recibe por referencia y lo carga
'con todas los estudios

Public Sub llenarListboxEstudios(lbxEstudios As ListBox)

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'limpia el listbox
lbxEstudios.Clear

'Hago la consulta
qbusca = " SELECT E.DESCRIPCION" & _
         " FROM ESTUDIOS AS E" & _
         " ORDER BY E.COD_ESTMED"
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
While Not rstDatos.EOF
    'carga los items al combo
    lbxEstudios.AddItem rstDatos!DESCRIPCION
    'avanza al registro siguiente
    rstDatos.MoveNext
Wend

'si no esta vacio el combo pone el primero
If lbxEstudios.ListCount <> 0 Then lbxEstudios.ListIndex = 0

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub


'Funcion para buscar el Codigo de una estudio medico
'Parametros:
'estudio: el nombre del estudio medico
'retorno: el codigo de dicho estudio o -1 si no la encontro

Public Function consultaCodDeESTUDIO(estudio As String) As Long
'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'Hago la consulta
qbusca = " SELECT COD_ESTMED " & _
         " FROM  ESTUDIOS AS EST " & _
         " WHERE EST.DESCRIPCION='" & estudio & "'"
         
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
If Not rstDatos.EOF Then
    'retorna
    consultaCodDeESTUDIO = rstDatos!COD_ESTMED
    Exit Function
End If

consultaCodDeESTUDIO = -1

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Function
End If
End Function

Public Function traer_nombre_De_Medico(cod As Variant) As Variant
'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

If Not IsNull(cod) Then

'Hago la consulta
qbusca = " SELECT NOMBREM " & _
         " FROM  MEDICOS AS M " & _
         " WHERE M.COD_MLEG=" & cod
         
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
If Not rstDatos.EOF Then
    'retorna
    traer_nombre_De_Medico = rstDatos!NOMBREM
    Exit Function
End If

Else
traer_nombre_De_Medico = Null
End If

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Function
End If

End Function
