VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmPrincipal 
   BackColor       =   &H00000000&
   Caption         =   "Hospital"
   ClientHeight    =   8550
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10770
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmPrincipal.frx":0000
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar tlbBarraDeHerramientas 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   1535
      ButtonWidth     =   2223
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "imlLista"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Pacientes"
            Key             =   "keyPacientes"
            Object.Tag             =   ""
            ImageIndex      =   8
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Doctores"
            Key             =   "keyDoctores"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Obras Sociales"
            Key             =   "keyObras"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Localidades"
            Key             =   "keyLocalidades"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Provincias"
            Key             =   "KeyProvincias"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Salir"
            Key             =   "keySalir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Cambio  Usuario"
            Key             =   "keyUsuario"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Ayuda"
            Key             =   "keyAyudaDeLaAplic"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8295
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Grupo 7"
            TextSave        =   "Grupo 7"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2893
            MinWidth        =   2893
            Text            =   "Comision de 10 a 12"
            TextSave        =   "Comision de 10 a 12"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            Text            =   "TALLER PROGRAMACION VISUAL CLIENTE/SERVIDOR"
            TextSave        =   "TALLER PROGRAMACION VISUAL CLIENTE/SERVIDOR"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlLista 
      Left            =   7680
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":13058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":13C5DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":14862E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":154680
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":1606D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":16C724
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":178776
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":17D4FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":18954E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":1955A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":1A15F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuCambioUsuario 
         Caption         =   "&Cambio de usuario"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuDoctores 
         Caption         =   "Doctores"
      End
      Begin VB.Menu mnuOs 
         Caption         =   "Obras Sociales"
      End
      Begin VB.Menu mnuLocalidad 
         Caption         =   "Localidad"
      End
      Begin VB.Menu mnuespecialidades 
         Caption         =   "Especialidades"
      End
      Begin VB.Menu mnuEstudios 
         Caption         =   "Estudios Medicos"
      End
      Begin VB.Menu mnuProvincia 
         Caption         =   "Provincia"
      End
      Begin VB.Menu mnuPacientes 
         Caption         =   "Pacientes"
      End
      Begin VB.Menu mnuHistoriasClinicas 
         Caption         =   "Historias Clinicas"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPacientesAten 
         Caption         =   "Pacientes Atendidos"
      End
      Begin VB.Menu mnuNominaDoc 
         Caption         =   "Nomina de Doctores"
      End
      Begin VB.Menu mnuNominaPacientes 
         Caption         =   "Nomina Pacientes"
      End
   End
   Begin VB.Menu mnuAdministrador 
      Caption         =   "Administrador"
      Begin VB.Menu mnunuevo 
         Caption         =   "Usuarios"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuAyudaDeLaAplicacion 
         Caption         =   "Ayuda de CLINICA ATENCION-PACIENTES"
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "Acerca de Clinica-Pacientes"
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'la siguiente declaracion solo es util para poder ejecutar el ayuda
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)


Private Sub MDIForm_Unload(Cancel As Integer)
Dim respuesta As Integer
On Error GoTo HuboError

respuesta = MsgBox("¿Esta seguro que desea salir?", vbYesNo + vbQuestion, "Cierre de la aplicacion")
If respuesta = vbYes Then
Unload frmPrincipal
End If

HuboError:
rtaError = evaluarError(Err)
'de acuerdo a la respuesta, realiza...
Select Case rtaError
    Case Finalizar
        End
    Case Reintentar
        Resume
    Case Ignorar
        Resume Next
    Case Cancelar
        'no hace nada
End Select
End Sub

Private Sub mnuAcercaDe_Click()
On Error GoTo HuboError

frmAcercaDe.Show vbModal

HuboError:
rtaError = evaluarError(Err)
'de acuerdo a la respuesta, realiza...
Select Case rtaError
    Case Finalizar
        End
    Case Reintentar
        Resume
    Case Ignorar
        Resume Next
    Case Cancelar
        'no hace nada
End Select

End Sub

Private Sub mnuAyudaDeLaAplicacion_Click()
On Error GoTo HuboError
  
  Dim nRet As Integer

    'si no hay archivo de ayuda para este proyecto, mostrar un mensaje al usuario
    'puede establecer el archivo de Ayuda para la aplicación en el cuadro
    'de diálogo Propiedades del proyecto
    If Len(App.HelpFile) = 0 Then
        MsgBox "No se puede mostrar el contenido de la Ayuda. No hay Ayuda asociada a este proyecto.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
    
HuboError:
rtaError = evaluarError(Err)
'de acuerdo a la respuesta, realiza...
Select Case rtaError
    Case Finalizar
        End
    Case Reintentar
        Resume
    Case Ignorar
        Resume Next
    Case Cancelar
        'no hace nada
End Select
    
End Sub

Private Sub mnuCambioUsuario_Click()
Dim respuesta As Integer
On Error GoTo HuboError

respuesta = MsgBox("¿Esta seguro que desea cambiar de usuario?", vbYesNo + vbQuestion, "CAMBIO DE USUARIO")
If respuesta = vbYes Then
   frmLogin.Show 'aparece el frmLogin
   frmLogin.relogin = 1 'indica que va a haber cambio de usuario

   frmLogin.txtPw = ""
   frmLogin.txtUsr = ""
   frmPrincipal.mnuVer.Enabled = False
   frmPrincipal.mnuAdministrador.Enabled = False
   frmPrincipal.mnuAyuda.Enabled = False
   frmPrincipal.tlbBarraDeHerramientas.Enabled = False

   frmLogin.Show
 
End If

HuboError:
rtaError = evaluarError(Err)
'de acuerdo a la respuesta, realiza...
Select Case rtaError
    Case Finalizar
        End
    Case Reintentar
        Resume
    Case Ignorar
        Resume Next
    Case Cancelar
        'no hace nada
End Select
End Sub

Private Sub mnuDoctores_Click()
On Error GoTo HuboError

frmDoctores.Show

HuboError:
rtaError = evaluarError(Err)
'de acuerdo a la respuesta, realiza...
Select Case rtaError
    Case Finalizar
        End
    Case Reintentar
        Resume
    Case Ignorar
        Resume Next
    Case Cancelar
        'no hace nada
End Select

End Sub

Private Sub mnuespecialidades_Click()
On Error GoTo HuboError

frmEspecialidades.Show

HuboError:
rtaError = evaluarError(Err)
'de acuerdo a la respuesta, realiza...
Select Case rtaError
    Case Finalizar
        End
    Case Reintentar
        Resume
    Case Ignorar
        Resume Next
    Case Cancelar
        'no hace nada
End Select

End Sub

Private Sub mnuEstudios_Click()
On Error GoTo HuboError

frmEstudios_medicos.Show

HuboError:
rtaError = evaluarError(Err)
'de acuerdo a la respuesta, realiza...
Select Case rtaError
    Case Finalizar
        End
    Case Reintentar
        Resume
    Case Ignorar
        Resume Next
    Case Cancelar
        'no hace nada
End Select

End Sub

Private Sub mnuHistoriasClinicas_Click()
On Error GoTo HuboError

frmHistoriasClinicas.Show

HuboError:
rtaError = evaluarError(Err)
'de acuerdo a la respuesta, realiza...
Select Case rtaError
    Case Finalizar
        End
    Case Reintentar
        Resume
    Case Ignorar
        Resume Next
    Case Cancelar
        'no hace nada
End Select

End Sub

Private Sub mnuLocalidad_Click()
On Error GoTo HuboError

frmLocalidad.Show

HuboError:
rtaError = evaluarError(Err)
'de acuerdo a la respuesta, realiza...
Select Case rtaError
    Case Finalizar
        End
    Case Reintentar
        Resume
    Case Ignorar
        Resume Next
    Case Cancelar
        'no hace nada
End Select

End Sub

Private Sub mnuNominaDoc_Click()
 'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'AQUI SE ESTABLECE EL ORDEN DE LA PRESENTACION!!!!
qbusca = " SELECT DISTINCT M.COD_MLEG,M.NOMBREM,M.DNI_M,M.DIRECCION,M.TELEFONO,M.COD_POSTAL" & _
         " FROM MEDICOS AS M,OBRASOCIAL_MEDICOS AS OBMED,OBRA_SOCIAL AS OB," & _
                "ESPECIALIDADES AS ESP,MEDICOS_ESPECIALIDAD AS MEDESP " & _
         " WHERE  M.COD_MLEG=OBMED.COD_MLEG" & _
         " AND    OBMED.COD_OBSOCIAL=OB.COD_OBSOCIAL" & _
         " AND    M.COD_MLEG=MEDESP.COD_MLEG" & _
         " AND    MEDESP.COD_ESP=ESP.COD_ESP" & _
         " ORDER BY M.COD_MLEG" ' ESTE ES EL ORDEN!!!
         
consultasql conn, qbusca, rstDatos
    
If rstDatos.EOF Then
    MsgBox "No hay datos con dicho filtro para poder imprimir...", vbInformation, "Error"
    Exit Sub
End If
    
'enlazo el recordset (DEBE SER SI O SI ADO!!) con el reporte
Set rptNominaDeDoctores.DataSource = rstDatos
'cargo etiquetas
rptNominaDeDoctores.Sections("secEncDeInf").Controls("lblTitulo").Caption = "NOMINA DE MEDICOS"
rptNominaDeDoctores.Sections("secEncDeInf").Controls("lblFecha").Caption = FormatDateTime(Date, vbLongDate)
'enlazo campos
rptNominaDeDoctores.Sections("secDetalles").Controls("txtCodMedico").DataField = "COD_MLEG"
rptNominaDeDoctores.Sections("secDetalles").Controls("txtNombreM").DataField = "NOMBREM"
rptNominaDeDoctores.Sections("secDetalles").Controls("txtDniM").DataField = "DNI_M"
rptNominaDeDoctores.Sections("secDetalles").Controls("txtDirM").DataField = "DIRECCION"
rptNominaDeDoctores.Sections("secDetalles").Controls("txtTelP").DataField = "TELEFONO"
rptNominaDeDoctores.Sections("secDetalles").Controls("txtCodPostalM").DataField = "COD_POSTAL"

'VISUALIZO EL REPORTE
rptNominaDeDoctores.Show
    
error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If

End Sub

Private Sub mnuNominaPacientes_Click()
 'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'AQUI SE ESTABLECE EL ORDEN DE LA PRESENTACION!!!!
qbusca = " SELECT P.COD_LEGP,P.NOMBRE,P.DNI,P.DIRECCION,P.TELEFONO,OB.RAZON_SOCIAL AS DESC_OB" & _
         " FROM PACIENTES AS P,OBRA_SOCIAL AS OB" & _
         " WHERE OB.COD_OBSOCIAL = P.COD_OBSOCIAL " & _
         " ORDER BY P.COD_LEGP" ' ESTE ES EL ORDEN!!!
         
consultasql conn, qbusca, rstDatos
    
If rstDatos.EOF Then
    MsgBox "No hay datos con dicho filtro para poder imprimir...", vbInformation, "Error"
    Exit Sub
End If
    
'enlazo el recordset (DEBE SER SI O SI ADO!!) con el reporte
Set rptNominaPacientes.DataSource = rstDatos
'cargo etiquetas
rptNominaPacientes.Sections("secEncDeInf").Controls("lblTitulo").Caption = "NOMINA DE PACIENTES"
rptNominaPacientes.Sections("secEncDeInf").Controls("lblFecha").Caption = FormatDateTime(Date, vbLongDate)
'enlazo campos
rptNominaPacientes.Sections("secDetalles").Controls("txtCodPaciente").DataField = "COD_LEGP"
rptNominaPacientes.Sections("secDetalles").Controls("txtNombreP").DataField = "NOMBRE"
rptNominaPacientes.Sections("secDetalles").Controls("txtDniP").DataField = "DNI"
rptNominaPacientes.Sections("secDetalles").Controls("txtDirP").DataField = "DIRECCION"
rptNominaPacientes.Sections("secDetalles").Controls("txtTelP").DataField = "TELEFONO"
rptNominaPacientes.Sections("secDetalles").Controls("txtObrasocialP").DataField = "DESC_OB"

'VISUALIZO EL REPORTE
rptNominaPacientes.Show
    
error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub

Private Sub mnunuevo_Click()
On Error GoTo HuboError

frmUsuarios.Show

HuboError:
rtaError = evaluarError(Err)
'de acuerdo a la respuesta, realiza...
Select Case rtaError
    Case Finalizar
        End
    Case Reintentar
        Resume
    Case Ignorar
        Resume Next
    Case Cancelar
        'no hace nada
End Select

End Sub

Private Sub mnuOs_Click()
On Error GoTo HuboError

frmObras_Sociales.Show

HuboError:
rtaError = evaluarError(Err)
'de acuerdo a la respuesta, realiza...
Select Case rtaError
    Case Finalizar
        End
    Case Reintentar
        Resume
    Case Ignorar
        Resume Next
    Case Cancelar
        'no hace nada
End Select

End Sub

Private Sub mnuPacientes_Click()
On Error GoTo HuboError

frmPacientes.Show

HuboError:
rtaError = evaluarError(Err)
'de acuerdo a la respuesta, realiza...
Select Case rtaError
    Case Finalizar
        End
    Case Reintentar
        Resume
    Case Ignorar
        Resume Next
    Case Cancelar
        'no hace nada
End Select

End Sub

Private Sub mnuPacientesAten_Click()
 'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo error

'AQUI SE ESTABLECE EL ORDEN DE LA PRESENTACION!!!!
qbusca = " SELECT P.COD_LEGP,P.NOMBRE,P.DNI,T.F_DE_ATENCION,ESP.DESCRIPCION" & _
         " FROM PACIENTES AS P,TURNOS AS T, ESPECIALIDADES AS ESP" & _
         " WHERE P.COD_LEGP = T.COD_LEGP " & _
         " AND T.COD_ESP=ESP.COD_ESP" & _
         " ORDER BY T.F_DE_ATENCION" ' ESTE ES EL ORDEN!!!
         
consultasql conn, qbusca, rstDatos
    
If rstDatos.EOF Then
    MsgBox "No hay datos con dicho filtro para poder imprimir...", vbInformation, "Error"
    Exit Sub
End If
    
'enlazo el recordset (DEBE SER SI O SI ADO!!) con el reporte
Set rptPacientesAtendidos.DataSource = rstDatos
'cargo etiquetas
rptPacientesAtendidos.Sections("secEncDeInf").Controls("lblTitulo").Caption = "PACIENTES ATENDIDOS"
rptPacientesAtendidos.Sections("secEncDeInf").Controls("lblFecha").Caption = FormatDateTime(Date, vbLongDate)
'enlazo campos
rptPacientesAtendidos.Sections("secDetalles").Controls("txtCodPaciente").DataField = "COD_LEGP"
rptPacientesAtendidos.Sections("secDetalles").Controls("txtNombreP").DataField = "NOMBRE"
rptPacientesAtendidos.Sections("secDetalles").Controls("txtDniP").DataField = "DNI"
rptPacientesAtendidos.Sections("secDetalles").Controls("txtFechadeatencion").DataField = "F_DE_ATENCION"
rptPacientesAtendidos.Sections("secDetalles").Controls("txtEspecialidad").DataField = "DESCRIPCION"

'VISUALIZO EL REPORTE
rptPacientesAtendidos.Show
    
error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If

End Sub

Private Sub mnuProvincia_Click()
On Error GoTo HuboError

frmProvincia.Show

HuboError:
rtaError = evaluarError(Err)
'de acuerdo a la respuesta, realiza...
Select Case rtaError
    Case Finalizar
        End
    Case Reintentar
        Resume
    Case Ignorar
        Resume Next
    Case Cancelar
        'no hace nada
End Select
End Sub

Private Sub mnuSalir_Click()
Dim respuesta As Integer
On Error GoTo HuboError

respuesta = MsgBox("¿Esta seguro que desea salir?", vbYesNo + vbQuestion, "Cierre de la aplicacion")
If respuesta = vbYes Then
Unload frmPrincipal
End If

HuboError:
rtaError = evaluarError(Err)
'de acuerdo a la respuesta, realiza...
Select Case rtaError
    Case Finalizar
        End
    Case Reintentar
        Resume
    Case Ignorar
        Resume Next
    Case Cancelar
        'no hace nada
End Select
End Sub

Private Sub tlbBarraDeHerramientas_ButtonClick(ByVal Button As ComctlLib.Button)
Dim respuesta As VbMsgBoxResult
On Error GoTo HuboError
Select Case Button.Key
Case "keyPacientes"
 mnuPacientes_Click
Case "keyDoctores"
 mnuDoctores_Click
Case "keyObras"
 mnuOs_Click
Case "keyLocalidades"
 mnuLocalidad_Click
Case "KeyProvincias"
 mnuProvincia_Click

Case "keySalir"
mnuSalir_Click
Case "keyUsuario"
 mnuCambioUsuario_Click
 
Case "keyAyudaDeLaAplic"
 mnuAyudaDeLaAplicacion_Click
End Select

HuboError:
rtaError = evaluarError(Err)
'de acuerdo a la respuesta, realiza...
Select Case rtaError
    Case Finalizar
        End
    Case Reintentar
        Resume
    Case Ignorar
        Resume Next
    Case Cancelar
        'no hace nada
End Select
End Sub
