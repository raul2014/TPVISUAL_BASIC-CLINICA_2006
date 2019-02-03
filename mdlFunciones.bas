Attribute VB_Name = "mdlFunciones"

Option Explicit

Public Sub restriccion_numeros(KeyAscii As Integer)
If KeyAscii = 13 Then
    'ignora la tecla enter
    KeyAscii = 0
    'envia un tab
    SendKeys "{tab}"
End If
'verifica que sea numerico o el backspace
'si no es, ignora la tecla pulsada(pone ascii=0)
If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack) Then
    KeyAscii = 0
End If
End Sub



Public Sub restriccion_solo_letras(KeyAscii As Integer)
If (((KeyAscii > 0) And (KeyAscii < 65)) And KeyAscii <> 32) _
   Or (KeyAscii > 90 And KeyAscii < 97) Or ((KeyAscii > 122)) _
   Or (KeyAscii = vbKeyBack) Then
    
    If KeyAscii <> 8 Then     'KeyAscii = 8 es el retroceso o BackSpace
       KeyAscii = 0
    End If

End If
End Sub

Public Sub habilitar_pacientes(valor As Boolean)
With frmPacientes
.lblLegajo.Enabled = Not valor
.lblCp.Enabled = Not valor
.txtCodigo_postal.Enabled = Not valor
.lblNombre.Enabled = Not valor
.txtNombre.Enabled = Not valor
.lbldni.Enabled = Not valor
.txtdni.Enabled = Not valor
.lblDomicilio.Enabled = Not valor
.txtDomicilio.Enabled = Not valor
.lblLocalidad.Enabled = Not valor
'.cmbLocalidad.Enabled = Not valor
.lblCod_Paciente.Enabled = Not valor
.lblProvincia.Enabled = Not valor
.cmbProvincia.Enabled = Not valor
.lblTelefono.Enabled = Not valor
.txtTelefono.Enabled = Not valor
.lblObra_Social.Enabled = Not valor
.CmbObra_social.Enabled = Not valor
.fraArriva.Enabled = Not valor
.lblSexo.Enabled = Not valor
.frasexo.Enabled = Not valor
.cmdGuardar.Enabled = Not valor
.cmdCancelar.Enabled = Not valor
.optFemenino.Enabled = Not valor
.OptMasculino.Enabled = Not valor

.fraAbajo.Enabled = valor
.cmdNuevo.Enabled = valor
.cmdModificar.Enabled = valor
.cmdEliminar.Enabled = valor
.cmdVolver.Enabled = valor
End With
End Sub



Public Sub desahabilitar_controles_doctores(valor As Boolean)
With frmDoctores
.lblLegajo.Enabled = valor
.lblCp.Enabled = valor
.txtCodigo_postal.Enabled = valor
.lblNombre.Enabled = valor
.txtNombre.Enabled = valor
.lbldni.Enabled = valor
.txtdni.Enabled = valor
.lblDomicilio.Enabled = valor
.txtDomicilio.Enabled = valor
.lblLocalidad.Enabled = valor
.cmbLocalidad.Enabled = valor
.lblProvincia.Enabled = valor
.cmbProvincia.Enabled = valor
.lblEspecialidad.Enabled = valor
.lbxEspecialidad.Enabled = valor
.cmdAgregarE.Enabled = valor
.CmdQuitarE.Enabled = valor
.lblObraSocial.Enabled = valor
.lbxObrasocial.Enabled = valor
.cmdAgregarOS.Enabled = valor
.CmdQuitarOS.Enabled = valor
.fraSuperior.Enabled = valor
End With
End Sub

Public Sub desabilitar_botones_alta_doctores(valor As Boolean)
With frmDoctores
.cmdNuevo.Enabled = valor
.cmdModificar.Enabled = valor
.cmdEliminar.Enabled = valor
.cmdVolver.Enabled = valor
End With
End Sub
'////
Public Sub Controles_doctores(valor As Boolean)
With frmDoctores
.txtTelefono.Enabled = valor
.lblCod_doctores.Enabled = valor
.lblLegajo.Enabled = valor
.lblCp.Enabled = valor
.txtCodigo_postal.Enabled = valor
.lblNombre.Enabled = valor
.txtNombre.Enabled = valor
.lbldni.Enabled = valor
.txtdni.Enabled = valor
.lblDomicilio.Enabled = valor
.txtDomicilio.Enabled = valor
.lblLocalidad.Enabled = valor
.cmbLocalidad.Enabled = valor
.lblProvincia.Enabled = valor
.cmbProvincia.Enabled = valor
.lblEspecialidad.Enabled = valor
.lbxEspecialidad.Enabled = valor
.cmdAgregarE.Enabled = valor
.CmdQuitarE.Enabled = valor
.lblObraSocial.Enabled = valor
.lbxObrasocial.Enabled = valor
.cmdAgregarOS.Enabled = valor
.CmdQuitarOS.Enabled = valor
.fraSuperior.Enabled = valor
.txtTelefono.Enabled = valor
.lblTelefono.Enabled = valor
.lblSexo.Enabled = valor
.Frmsexo.Enabled = valor
.cmdGuardar.Enabled = valor
.cmdCancelar.Enabled = valor

.fraInferior.Enabled = Not valor
.lvwDoctores.Enabled = Not valor
.cmdNuevo.Enabled = Not valor
.cmdModificar.Enabled = Not valor
.cmdEliminar.Enabled = Not valor
.cmdVolver.Enabled = Not valor
End With
End Sub



'////
Public Sub Habilitar_Obras_Sociales(valor As Boolean)
 With frmObras_Sociales
  .txtNombre.Enabled = Not valor
  .txtDomicilio.Enabled = Not valor
  .lblCodOBSocial.Enabled = Not valor
  .cmbProvincia.Enabled = Not valor
  .txtCuit.Enabled = Not valor
  .txtCodigo_postal.Enabled = Not valor
  .cmdGuardar.Enabled = Not valor
  .cmdCancelar.Enabled = Not valor
  .lblN°deLegajo.Enabled = Not valor
  .lblRazon_social.Enabled = Not valor
  .lblDomicilio.Enabled = Not valor
  .lblLocalidad.Enabled = Not valor
  .lblProvincia.Enabled = Not valor
  .lblCuit.Enabled = Not valor
  .lblCp.Enabled = Not valor
  .fraSuperior.Enabled = Not valor
  
  .fraAbajo.Enabled = valor
  .cmdNuevo.Enabled = valor
  .cmdModificar.Enabled = valor
  .cmdEliminar.Enabled = valor
  .cmdVolver.Enabled = valor
  
 End With
End Sub

Public Sub Controles_HistoriasClinicas(valor As Boolean)
With frmHistoriasClinicas
.lblCod_Turno.Enabled = valor
.lblPaciente.Enabled = valor
.lblFechaDeAtencion.Enabled = valor
.lblCodMed.Enabled = valor
.lblCodP.Enabled = valor
.dtpkFechaDeAtencion.Enabled = valor
.lblCodPaciente.Enabled = valor
.lblCodTurno.Enabled = valor
.lblCodM.Enabled = valor
.txtDiagnostico.Enabled = valor
.lblCodM.Enabled = valor
.lblDiagnostico.Enabled = valor
.cmbMedicos.Enabled = valor
.cmbPacientes.Enabled = valor
'.cmbEspecialidades.Enabled = valor
.lblMedico.Enabled = valor
.lblEspecialidad.Enabled = valor
.lblEstudios.Enabled = valor
.lblCodEsp.Enabled = valor
.lblCodEspecialidad.Enabled = valor
.lbxEstudiosDeTurnos.Enabled = valor
.fraSuperior.Enabled = valor
.cmdGuardar.Enabled = valor
.cmdCancelar.Enabled = valor
.cmdAgregarEstudio.Enabled = valor
.cmdQuitarEstudio.Enabled = valor

.fraInferior.Enabled = Not valor
.lvwTurnos.Enabled = Not valor
.cmdNuevo.Enabled = Not valor
.cmdModificar.Enabled = Not valor
.cmdEliminar.Enabled = Not valor
.cmdVolver.Enabled = Not valor
End With
End Sub


Public Sub Habilitar_Localidades(valor As Boolean)
 With frmLocalidad
.lblCodLoc.Enabled = Not valor
.lblCodigo.Enabled = Not valor
.lblLocalidad.Enabled = Not valor
.txtLocalidad.Enabled = Not valor
.lblProvincia.Enabled = Not valor
.cmbProvincia.Enabled = Not valor
.cmdGuardar.Enabled = Not valor
.cmdCancelar.Enabled = Not valor
.fraArriva.Enabled = Not valor

.fraAbajo.Enabled = valor
.cmdNuevo.Enabled = valor
.cmdModificar.Enabled = valor
.cmdEliminar.Enabled = valor
.cmdVolver.Enabled = valor
 End With
End Sub

Public Sub Habilitar_Provincia(valor As Boolean)
 With frmProvincia
 .lblID.Enabled = Not valor
.lblCodigo.Enabled = Not valor
.lblProvincia.Enabled = Not valor
.txtProvincia.Enabled = Not valor
.cmdGuardar.Enabled = Not valor
.cmdCancelar.Enabled = Not valor
.fraArriva.Enabled = Not valor
.txtUsuarioDeAlta.Enabled = Not valor
.txtUsuarioDeModif.Enabled = Not valor
.lblFechaDeAlta.Enabled = Not valor
.lblFechaDeUltModif.Enabled = Not valor
.lblUsuarioDeAlta.Enabled = Not valor
.lblUsuarioDeModif.Enabled = Not valor
.dtpkFechaDeAlta.Enabled = Not valor
.dtpkFechaDeModif.Enabled = Not valor

.fraAbajo.Enabled = valor
.cmdNuevo.Enabled = valor
.cmdModificar.Enabled = valor
.cmdEliminar.Enabled = valor
.cmdVolver.Enabled = valor
 End With
End Sub
Public Sub Habilitar_Especialidad(valor As Boolean)
 With frmEspecialidades
.lblCodEsp.Enabled = Not valor
.lblCodigo.Enabled = Not valor
.lblEspecialidad.Enabled = Not valor
.txtEspecialidad.Enabled = Not valor
.cmdGuardar.Enabled = Not valor
.cmdCancelar.Enabled = Not valor
.fraArriva.Enabled = Not valor

.fraInferior.Enabled = valor
.cmdNuevo.Enabled = valor
.cmdModificar.Enabled = valor
.cmdEliminar.Enabled = valor
.cmdVolver.Enabled = valor

End With
End Sub
Public Sub Habilitar_estudios_medicos(valor As Boolean)
 With frmEstudios_medicos
 
.lblCodEstudiosMedicos.Enabled = Not valor
.lblCodigo.Enabled = Not valor
.lblDescripcion.Enabled = Not valor
.txtNombre.Enabled = Not valor
.LblComplejidad.Enabled = Not valor
.OptAlta.Enabled = Not valor
.OptMedia.Enabled = Not valor
.OptBaja.Enabled = Not valor
.cmdGuardar.Enabled = Not valor
.cmdCancelar.Enabled = Not valor
.fraArriva.Enabled = Not valor
 
.lvwEstudios_Medicos.Enabled = valor
.fraInferior.Enabled = valor
.cmdNuevo.Enabled = valor
.cmdModificar.Enabled = valor
.cmdElimina.Enabled = valor
.cmdVolver.Enabled = valor
 
 End With
End Sub

Public Sub Habilitar_ABMUsuarios(valor As Boolean)
 With frmUsuarios

.lblCodUsuario.Enabled = Not valor
.lblNroUsuario.Enabled = Not valor
.lblNombre.Enabled = Not valor
.lblPassword.Enabled = Not valor
.lblUsername.Enabled = Not valor
.lblUsFechaDeAlta.Enabled = Not valor
.lblUsFechaDeUltModif.Enabled = Not valor
.lblListaMedicos.Enabled = Not valor
.txtNombre.Enabled = Not valor
.fraCategoria.Enabled = Not valor
.fraMedico.Enabled = Not valor
.txtUsername.Enabled = Not valor
.txtPassword.Enabled = Not valor
.cmdCancelar.Enabled = Not valor
.cmdGuardarU.Enabled = Not valor
.fraSuperior.Enabled = Not valor
.optAdministrador.Enabled = Not valor
.optUsuario.Enabled = Not valor
.optNoM.Enabled = Not valor
.optSiM.Enabled = Not valor
'.cmbListaMedicos.Enabled = Not valor
 
 
.lvwUsuarios.Enabled = valor
.fraInferior.Enabled = valor
.cmdAgregar.Enabled = valor
.cmdModificar.Enabled = valor
.cmdEliminar.Enabled = valor
.cmdVolver.Enabled = valor
 
 End With
End Sub

'Esta funcion la utilizo para el formulario de Estudios medicos

Public Function detectar_complejidad() As String

If frmEstudios_medicos.OptAlta.Value = True Then
       detectar_complejidad = frmEstudios_medicos.OptAlta.Caption
End If

If frmEstudios_medicos.OptBaja.Value = True Then
       detectar_complejidad = frmEstudios_medicos.OptBaja.Caption
End If

If frmEstudios_medicos.OptMedia.Value = True Then
       detectar_complejidad = frmEstudios_medicos.OptMedia.Caption
End If

End Function


' esta funcion la utilizo en el formulario pacientes
'cuando necesito actualizar o guardar
Public Function detectar_categoriaDeUsuario() As String

If frmUsuarios.optAdministrador.Value = True Then
       detectar_categoriaDeUsuario = UCase(frmUsuarios.optAdministrador.Caption)
End If

If frmUsuarios.optUsuario.Value = True Then
      detectar_categoriaDeUsuario = UCase(frmUsuarios.optUsuario.Caption)
End If

End Function

' esta funcion la utilizo en el formulario DOCTORES
'cuando necesito actualizar o guardar
Public Function detectar_sexo() As String

If frmDoctores.OptMasculino.Value = True Then
       detectar_sexo = UCase(frmDoctores.OptMasculino.Caption)
End If

If frmDoctores.optFemenino.Value = True Then
      detectar_sexo = UCase(frmDoctores.optFemenino.Caption)
End If

End Function

'Esta funcion la utilizo en la tabla medicos, lo que hace es controlar que no se
'ingresen repetidos.

Public Function controlarNoRepetidos(lbx As ListBox, item As String) As Long
Dim resp As Long
resp = 0
Dim cant As Long
cant = lbx.ListCount
Dim i As Integer
i = 0

If lbx.ListIndex <> -1 Then 'si el listbox no esta vacio entonces...
i = 0
End If

If lbx.Text <> item Then
  While (i < cant) And (resp <> 1)
     If (lbx.List(i) = item) Then
        resp = 1
     End If
    i = i + 1

  Wend

  controlarNoRepetidos = resp
Else
controlarNoRepetidos = 1
End If

End Function






