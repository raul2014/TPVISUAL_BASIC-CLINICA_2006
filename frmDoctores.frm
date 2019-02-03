VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDoctores 
   BackColor       =   &H00000000&
   Caption         =   "Doctores"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   11040
   Begin VB.Frame FrmAgregarOs 
      BackColor       =   &H00808000&
      Height          =   3495
      Left            =   8760
      TabIndex        =   38
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
      Begin VB.CommandButton CmdCerrar2 
         Caption         =   "Cerrar Ventana"
         Height          =   375
         Left            =   480
         TabIndex        =   41
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgregarOS2 
         Caption         =   "Agregar Obra Social"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ListBox lbxObrasocial2 
         Height          =   2010
         ItemData        =   "frmDoctores.frx":0000
         Left            =   240
         List            =   "frmDoctores.frx":0002
         TabIndex        =   39
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame FrmAgregarEspecialidad 
      BackColor       =   &H00404000&
      Height          =   3615
      Left            =   6480
      TabIndex        =   34
      Top             =   360
      Visible         =   0   'False
      Width           =   2295
      Begin VB.ListBox lbxEspecialidad2 
         Height          =   2010
         ItemData        =   "frmDoctores.frx":0004
         Left            =   120
         List            =   "frmDoctores.frx":0006
         TabIndex        =   37
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdAgregarE2 
         Caption         =   "Agregar especialidad"
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton CmdCerrar 
         Caption         =   "Cerrar Ventana"
         Height          =   375
         Left            =   600
         TabIndex        =   35
         Top             =   3000
         Width           =   1215
      End
   End
   Begin VB.Frame fraInferior 
      BackColor       =   &H00E0E0E0&
      Height          =   3975
      Left            =   120
      TabIndex        =   33
      Top             =   4200
      Width           =   10935
      Begin MSComctlLib.ListView lvwDoctores 
         Height          =   2775
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16710882
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro de legajo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Apellido y nombre"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "DNI"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Telefono"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Direccion"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Localidad"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "FECHA ALTA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "FECHA DE MODIF"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "U DE ALTA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "U DE MODIF"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1200
         TabIndex        =   0
         Top             =   3255
         Width           =   1455
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3240
         TabIndex        =   1
         Top             =   3255
         Width           =   1455
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5160
         TabIndex        =   2
         Top             =   3255
         Width           =   1455
      End
      Begin VB.CommandButton cmdVolver 
         Caption         =   "Volver"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7080
         TabIndex        =   3
         Top             =   3240
         Width           =   1590
      End
   End
   Begin VB.Frame fraSuperior 
      BackColor       =   &H00E0E0E0&
      Height          =   4215
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   10905
      Begin VB.Frame FraUsuario 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   4215
         Left            =   8640
         TabIndex        =   46
         Top             =   120
         Width           =   2055
         Begin VB.TextBox txtUsuarioDeModif 
            Height          =   285
            Left            =   120
            TabIndex        =   48
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtUsuarioDeAlta 
            Height          =   285
            Left            =   120
            TabIndex        =   47
            Top             =   2880
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpFechaDeModif 
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            Format          =   20971521
            CurrentDate     =   39005
         End
         Begin MSComCtl2.DTPicker dtpFechaDeAlta 
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            Format          =   20971521
            CurrentDate     =   39005
         End
         Begin VB.Label lblUsuarioDeModif 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Usuario de modificacion:"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label lblUsuarioDeAlta 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Usuario de Alta:"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label lblFechaDeUltModif 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fecha de ultima  modificacion:"
            Height          =   495
            Left            =   120
            TabIndex        =   52
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label lblFechaDeAlta 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fecha de alta:"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.TextBox TxtSexo 
         Height          =   285
         Left            =   2400
         TabIndex        =   45
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmbProvincia 
         BackColor       =   &H00FEFCE2&
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Text            =   "cmbProvincia"
         Top             =   1440
         Width           =   2220
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H00FEFCE2&
         DataField       =   "NOMBRES"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   4
         Top             =   720
         Width           =   2880
      End
      Begin VB.TextBox txtDomicilio 
         BackColor       =   &H00FEFCE2&
         DataField       =   "DIRECCION"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1110
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1080
         Width           =   5430
      End
      Begin VB.TextBox txtTelefono 
         BackColor       =   &H00FEFCE2&
         DataField       =   "TELEFONO"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Data2"
         Height          =   285
         Left            =   6960
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1440
         Width           =   1485
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1920
         TabIndex        =   21
         Top             =   3360
         Width           =   1500
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   360
         TabIndex        =   20
         Top             =   3360
         Width           =   1215
      End
      Begin VB.ComboBox cmbLocalidad 
         BackColor       =   &H00FEFCE2&
         Height          =   315
         Left            =   1110
         TabIndex        =   8
         Text            =   "cmbLocalidad"
         Top             =   1920
         Width           =   2220
      End
      Begin VB.Frame Frmsexo 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Left            =   960
         TabIndex        =   19
         Top             =   2160
         Width           =   1335
         Begin VB.OptionButton OptMasculino 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Masculino"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optFemenino 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Femenino"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.TextBox txtdni 
         BackColor       =   &H00FEFCE2&
         DataField       =   "dni"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Data2"
         Height          =   285
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   5
         Top             =   720
         Width           =   1755
      End
      Begin VB.CommandButton cmdAgregarE 
         Caption         =   "Agregar especialidad"
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   3360
         Width           =   2295
      End
      Begin VB.CommandButton cmdAgregarOS 
         Caption         =   "Agregar Obra Social"
         Height          =   255
         Left            =   6480
         TabIndex        =   17
         Top             =   3360
         Width           =   2175
      End
      Begin VB.ListBox lbxObrasocial 
         BackColor       =   &H00FEFCE2&
         Height          =   1035
         ItemData        =   "frmDoctores.frx":0008
         Left            =   6360
         List            =   "frmDoctores.frx":000A
         TabIndex        =   16
         Top             =   2160
         Width           =   2175
      End
      Begin VB.ListBox lbxEspecialidad 
         BackColor       =   &H00FEFCE2&
         Height          =   1035
         ItemData        =   "frmDoctores.frx":000C
         Left            =   3720
         List            =   "frmDoctores.frx":000E
         TabIndex        =   13
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtCodigo_postal 
         BackColor       =   &H00FEFCE2&
         DataField       =   "Codigo postal"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Data2"
         Height          =   285
         Left            =   4440
         MaxLength       =   4
         TabIndex        =   11
         Top             =   1440
         Width           =   1485
      End
      Begin VB.CommandButton CmdQuitarE 
         Caption         =   "Quitar especialidad"
         Height          =   315
         Left            =   3720
         TabIndex        =   15
         Top             =   3720
         Width           =   2295
      End
      Begin VB.CommandButton CmdQuitarOS 
         Caption         =   "Quitar Obra Social"
         Height          =   255
         Left            =   6480
         TabIndex        =   18
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblCod_doctores 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   1320
         TabIndex        =   44
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblProvincia 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Provincia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblSexo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sexo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lbldni 
         BackColor       =   &H00E0E0E0&
         Caption         =   "D.N.I."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   31
         Top             =   750
         Width           =   615
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nombre y Apellido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   210
         TabIndex        =   30
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblDomicilio 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Domicilio:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   29
         Top             =   1095
         Width           =   1095
      End
      Begin VB.Label lblTelefono 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   28
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblLegajo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "N° de legajo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   225
         TabIndex        =   27
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label lblLocalidad 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Localidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblEspecialidad 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Especialidad"
         Height          =   255
         Left            =   3720
         TabIndex        =   25
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblObraSocial 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Obra social"
         Height          =   255
         Left            =   6480
         TabIndex        =   24
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblCp 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Codigo Postal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   23
         Top             =   1440
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmDoctores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tipoOperacion As Integer

Private Sub clean()
On Error GoTo HuboError

Me.lblCod_doctores = ""
Me.txtNombre = ""
Me.txtdni = ""
Me.txtDomicilio = ""
Me.txtTelefono = ""
Me.txtCodigo_postal = ""

'si no esta vacio el combo de provincias
If Me.cmbProvincia.ListCount <> 0 Then cmbProvincia.ListIndex = 0

'si no esta vacio el combo de localidades
If Me.cmbLocalidad.ListCount <> 0 Then cmbLocalidad.ListIndex = 0


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

Private Function datosValidos() As Boolean
On Error GoTo HuboError
' chequea que se halla ingresado el nombre
    If Len(Me.txtNombre) = 0 Then
        MsgBox "Debe introducir el nombre", vbInformation, "Error de validación"
        Me.txtNombre.SetFocus
        datosValidos = False
        Exit Function
    End If
    
' chequea que se halla ingresado el nro de documento
    If Len(Me.txtdni) = 0 Then
        MsgBox "Debe introducir el numero de documento", vbInformation, "Error de validación"
        Me.txtdni.SetFocus
        datosValidos = False
        Exit Function
    End If
    
 ' chequea que se halla ingresado la direccion
    If Len(Me.txtDomicilio) = 0 Then
        MsgBox "Debe introducir el domicilio", vbInformation, "Error de validación"
        Me.txtNombre.SetFocus
        datosValidos = False
        Exit Function
    End If
 'chequea si se ingreso el codigo postal
 
    If Len(Me.txtCodigo_postal) = 0 Then
        MsgBox "Debe introducir el codigo postal", vbInformation, "Error de validación"
        Me.txtCodigo_postal.SetFocus
        datosValidos = False
        Exit Function
    End If
    
' chequea que se halla ingresado alguna ESPECIALIDAD
    If Me.lbxEspecialidad.ListCount = 0 Then
       MsgBox "Debe introducir alguna especialidad", vbInformation, "Error de validacion"
       datosValidos = False
       Exit Function
    End If
    
' chequea que se halla ingresado alguna OBRA SOCIAL
    If Me.lbxObrasocial.ListCount = 0 Then
       MsgBox "Debe introducir alguna obra social", vbInformation, "Error de validacion"
       datosValidos = False
       Exit Function
    End If
    
    
    datosValidos = True
    
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

End Function

Private Sub actualizaLista()

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset
Dim mitem As ListItem

'Manejo de Error
On Error GoTo error

'limpia la lista
lvwDoctores.ListItems.Clear

qbusca = " SELECT M.COD_MLEG,M.DNI_M,M.NOMBREM,M.DIRECCION,M.COD_POSTAL,M.TELEFONO,M.SEXO," & _
                 "M.FECHA_ALTA,M.FECHA_ULTMODIF,M.UDEALTA,M.UDEMODIF,M.COD_LOC,L.DESCRIPCION" & _
         " FROM MEDICOS AS M , LOCALIDADES AS L" & _
         " WHERE M.COD_LOC = L.COD_LOC" & _
         " ORDER BY M.COD_MLEG"
         
consultasql conn, qbusca, rstDatos


'mientras no sea fin de archivo
While Not rstDatos.EOF
    'agrega el item a la lista
    Set mitem = lvwDoctores.ListItems.Add()
    mitem.Text = rstDatos!COD_MLEG
    mitem.SubItems(1) = rstDatos!NOMBREM
    mitem.SubItems(2) = rstDatos!DNI_M
    mitem.SubItems(3) = rstDatos!TELEFONO
    mitem.SubItems(4) = rstDatos!DIRECCION
    mitem.SubItems(5) = rstDatos!DESCRIPCION
    mitem.SubItems(6) = rstDatos!FECHA_ALTA
    mitem.SubItems(7) = rstDatos!FECHA_ULTMODIF
    mitem.SubItems(8) = rstDatos!UDEALTA
    mitem.SubItems(9) = rstDatos!UDEMODIF
    
    
    'avanza al siguiente registro
    rstDatos.MoveNext
Wend

'selecciona por defecto al primero encontrado
If Me.lvwDoctores.ListItems.Count <> 0 Then
    Me.lvwDoctores.ListItems(1).Selected = True
    'carga en la parte superior
    lvwDoctores_Click
Else
    'si no hay nada, limpia
    clean
End If

error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub




Private Sub cmbProvincia_Click()
On Error GoTo HuboError
Dim qbusca As String    'se llenara el combo de localidades segun la provincia seleccionada
Dim rstDatos As New ADODB.Recordset

'limpia el combo
Me.cmbLocalidad.Clear

'Hago la consulta
qbusca = " SELECT DISTINCT L.DESCRIPCION AS DESCLOC,L.COD_PROV " & _
         " FROM LOCALIDADES AS L, PROVINCIAS AS P" & _
         " WHERE L.COD_PROV=" & consultaCodProvincia(Me.cmbProvincia.Text) & _
         " ORDER BY L.DESCRIPCION "
         
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
While Not rstDatos.EOF
    'carga los items al combo
    cmbLocalidad.AddItem rstDatos!DESCLOC
    'avanza al registro siguiente
    rstDatos.MoveNext
Wend

Me.cmbLocalidad.Enabled = True

HuboError:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    conn.RollbackTrans
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub

Private Sub cmdAgregarE2_Click()
Dim r As Long
On Error GoTo HuboError

r = controlarNoRepetidos(Me.lbxEspecialidad, Me.lbxEspecialidad2.Text)
If (r <> 1) Then
    Me.lbxEspecialidad.AddItem Me.lbxEspecialidad2.Text
Else
    MsgBox ("No se admiten valores repetidos")
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

Private Sub cmdAgregarOS2_Click()
Dim r As Long

On Error GoTo HuboError

r = controlarNoRepetidos(Me.lbxObrasocial, Me.lbxObrasocial2.Text)
If (r <> 1) Then
    Me.lbxObrasocial.AddItem Me.lbxObrasocial2.Text
Else
    MsgBox ("No se admiten valores repetidos")
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

Private Sub cmdEliminar_Click()
Dim res As Variant
'Manejo de Error
On Error GoTo HuboError


'pregunta antes
res = MsgBox("¿Desea borrar realmente el registro de codigo: " & Me.lvwDoctores.SelectedItem.Text & "?", vbQuestion + vbYesNo, "Eliminar Registro")
If res = vbNo Then Exit Sub
    
conn.BeginTrans
    

' 1ro elimino las especialidades de la tabla MEDICOS_ESPECIALIDAD (SI NO VA HABER PROBLEMA DE QUE HAY REG RELACIONADOS)
conn.Execute "DELETE FROM MEDICOS_ESPECIALIDAD WHERE COD_MLEG = " & Me.lvwDoctores.SelectedItem.Text
    
' 2do elimino las obras sociales de la tabla OBRASOCIAL_MEDICOS (SI NO VA HABER PROBLEMA DE QUE HAY REG RELACIONADOS)
conn.Execute "DELETE FROM OBRASOCIAL_MEDICOS WHERE COD_MLEG = " & Me.lvwDoctores.SelectedItem.Text
    
'Luego elimino el registro que corresponde al MEDICO seleccionado en el listview
conn.Execute "DELETE FROM MEDICOS WHERE COD_MLEG = " & Me.lvwDoctores.SelectedItem.Text
    

conn.CommitTrans
    
    
'limpia y actualiza
clean
actualizaLista


HuboError:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    conn.RollbackTrans
    MsgBox "Error: " + Err.Description
    Exit Sub
End If

'HuboError:
'rtaError = evaluarError(Err)
''de acuerdo a la respuesta, realiza...
'Select Case rtaError
'    Case Finalizar
'        End
'    Case Reintentar
'        Resume
'    Case Ignorar
'        Resume Next
'    Case Cancelar
'        'no hace nada
'End Select

End Sub

Private Sub CmdQuitarE_Click()
On Error GoTo HuboError

'si no selecciono nada, devuelve indice -1
    If Me.lbxEspecialidad.ListIndex = -1 Then
        If Me.lbxEspecialidad.ListCount = 0 Then
            MsgBox "No hay items cargados para poder eliminar", vbInformation, "Error"
        Else
            MsgBox "Debe seleccionar algún item para poder eliminar", vbInformation, "Error"
        End If
        Exit Sub
    End If
    Me.lbxEspecialidad.RemoveItem Me.lbxEspecialidad.ListIndex

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

Private Sub CmdQuitarOS_Click()
On Error GoTo HuboError

'si no selecciono nada, devuelve indice -1
    If Me.lbxObrasocial.ListIndex = -1 Then
        If Me.lbxObrasocial.ListCount = 0 Then
            MsgBox "No hay items cargados para poder eliminar", vbInformation, "Error"
        Else
            MsgBox "Debe seleccionar algún item para poder eliminar", vbInformation, "Error"
        End If
        Exit Sub
    End If
    Me.lbxObrasocial.RemoveItem Me.lbxObrasocial.ListIndex

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

Private Sub Form_Load()
On Error GoTo HuboError

lblTelefono.Enabled = False
lblSexo.Enabled = False
Frmsexo.Enabled = False
cmdGuardar.Enabled = False
cmdCancelar.Enabled = False

'llena el combo provincia con todas las provincias
cargaComboProvincias cmbProvincia

'llena el combo localidad con todas las localidades
cargaComboLocalidades cmbLocalidad

llenarListboxEspecialidades Me.lbxEspecialidad2

llenarListboxObrasSociales Me.lbxObrasocial2

'carga el listview
actualizaLista

'deshabilita la parte superior del ABM
Controles_doctores False

Me.cmbLocalidad.Enabled = False

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


Private Sub cmdAgregarE_Click()
On Error GoTo HuboError

frmDoctores.FrmAgregarEspecialidad.Visible = True

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

Private Sub cmdAgregarOS_Click()
On Error GoTo HuboError

frmDoctores.FrmAgregarOs.Visible = True

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

Private Sub cmdCancelar_Click()
On Error GoTo HuboError

'limpia todo
clean
'limpio los listbox
Me.lbxEspecialidad.Clear
Me.lbxObrasocial.Clear

Controles_doctores False

Me.cmbLocalidad.Enabled = False

'agrego estas dos lineas para controlar que cuando se  cancele, los frames  de agreg esp  y el de agreg ob
'se oculten aunque el usuario no los halla cerrado.
Me.FrmAgregarEspecialidad.Visible = False
Me.FrmAgregarOs.Visible = False

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

Private Sub CmdCerrar_Click()
On Error GoTo HuboError

frmDoctores.FrmAgregarEspecialidad.Visible = False

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

Private Sub CmdCerrar2_Click()
On Error GoTo HuboError

frmDoctores.FrmAgregarOs.Visible = False

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

Private Sub cmdGuardar_Click()
Dim sql As String
Dim qbusca As String

Dim cant As Long
Dim i As Integer

'Manejo de Error
On Error GoTo HuboError

 'primero verifica que sea valido
 If Not datosValidos Then Exit Sub
    
 conn.BeginTrans
    
 'guarda segun la operacion
 Select Case tipoOperacion
     Case ALTA
          sql = "INSERT INTO MEDICOS(COD_MLEG,DNI_M,NOMBREM,DIRECCION,COD_POSTAL,TELEFONO, " & _
                                "SEXO,FECHA_ALTA,FECHA_ULTMODIF,COD_LOC,UDEALTA,UDEMODIF) " & _
                         "VALUES(" & Me.lblCod_doctores & "," & Me.txtdni & " ,'" & _
                          Me.txtNombre & "','" & Me.txtDomicilio & "'," & _
                          Me.txtCodigo_postal & ",'" & Me.txtTelefono & "','" & _
                            detectar_sexo & "',#" & _
                          Format(Me.dtpFechaDeAlta, "mm/dd/yyyy") & "#,#" & _
                          Format(Me.dtpFechaDeModif, "mm/dd/yyyy") & "#," & _
                           consultaCodLocalidad(Me.cmbLocalidad) & "," & _
                          Me.txtUsuarioDeAlta & "," & _
                          Me.txtUsuarioDeModif & ");"
                          
          
        conn.Execute sql
        
        'ahora grabo las especialidades que le correspondes al medico
        
          '//
        cant = Me.lbxEspecialidad.ListCount 'cuento la cant de elem del lbx de especialidades

        If Me.lbxEspecialidad.ListCount > 0 Then  'si el listbox no esta vacio entonces...
           i = 0
        End If
        ' ahora recorro el lbx que contiene las especialidades
        While (i < cant)
         
          qbusca = "INSERT INTO MEDICOS_ESPECIALIDAD(COD_MLEG,COD_ESP) " & _
                   "VALUES(" & Me.lblCod_doctores & "," & _
                    consultaCodEspecialidad(Me.lbxEspecialidad.List(i)) & ");"

          conn.Execute qbusca

          i = i + 1   ' paso al siguiente elem

        Wend
          
       'ahora grabo las OBRAS SOCIALES que le correspondes al medico
        
        cant = Me.lbxObrasocial.ListCount 'cuento la cant de elem del lbx de obras sociales

        If Me.lbxObrasocial.ListCount > 0 Then 'si el listbox no esta vacio entonces...
           i = 0
        End If
        
        ' ahora recorro el lbx que contiene las obras sociales
        While (i < cant)
         
          qbusca = "INSERT INTO OBRASOCIAL_MEDICOS(COD_OBSOCIAL,COD_MLEG) " & _
                   "VALUES(" & consultaCodObSocial(Me.lbxObrasocial.List(i)) & "," & _
                     Me.lblCod_doctores & ");"

          conn.Execute qbusca

          i = i + 1   ' paso al siguiente elem

        Wend
       
       
    Case MODIFICACION
            'lo que cambia es el usuario de modif (uso la variable global 'usuarioActual) y la
            'fecha de ult modif(uso la cte DATE que contiene la fecha actual)
            'para dejar sentado que usuario logueado es el que guarda los cambios
           'LOS CAMPOS QUE OMITA SON LOS QUE NO SE ACTUALIZAN
           
        sql = "UPDATE MEDICOS SET " & _
                         "DNI_M=" & Me.txtdni & "," & _
                         "NOMBREM='" & Me.txtNombre & "'," & _
                         "DIRECCION='" & Me.txtDomicilio & "'," & _
                         "COD_POSTAL=" & Me.txtCodigo_postal & "," & _
                         "TELEFONO='" & Me.txtTelefono & "'," & _
                         "SEXO='" & detectar_sexo & "'," & _
                         "FECHA_ULTMODIF=#" & Format(Date, "mm/dd/yyyy") & "#," & _
                         "COD_LOC=" & consultaCodLocalidad(Me.cmbLocalidad) & "," & _
                         "UDEMODIF=" & usuarioActual & _
                         " WHERE COD_MLEG=" & Me.lblCod_doctores
   
       '1ro actualizo los datos de la tabla MEDICOS
        conn.Execute sql
            
        '2do borro las especialidades correspondientes al medico que se va a ACTUALIZAR
        'es decir elimino las especialidades de la tabla MEDICOS_ESPECIALIDAD ya que como
        'se agegaran o quitaran especialidades en el lbx entonces borro los del que quiero modif
        'y luego insertare (ver 3ro) los que hallan quedado en el lbx de especialidades.
        
        conn.Execute "DELETE FROM MEDICOS_ESPECIALIDAD WHERE COD_MLEG = " & Me.lvwDoctores.SelectedItem.Text
        
        'idem con respecto a las obras sociales
        
        conn.Execute "DELETE FROM OBRASOCIAL_MEDICOS WHERE COD_MLEG = " & Me.lvwDoctores.SelectedItem.Text
        
        
        '3ro inserto las especialidades definitivos del lbx de estudios
        
        cant = Me.lbxEspecialidad.ListCount 'cuento la cant de elem del lbx de especialidades

        If Me.lbxEspecialidad.ListCount > 0 Then  'si el listbox no esta vacio entonces...
           i = 0
        End If
        ' ahora recorro el lbx que contiene los estudios medicos
        While (i < cant)
         
          qbusca = "INSERT INTO MEDICOS_ESPECIALIDAD(COD_MLEG,COD_ESP) " & _
                   "VALUES(" & Me.lblCod_doctores & "," & _
                    consultaCodEspecialidad(Me.lbxEspecialidad.List(i)) & ");"

          conn.Execute qbusca

          i = i + 1   ' paso al siguiente elem

        Wend
        
        
        '******realizo lo mismo con resp a la tabla de OBRASOCIAL_MEDICOS
        
        cant = Me.lbxObrasocial.ListCount 'cuento la cant de elem del lbx de las obras sociales

        If Me.lbxObrasocial.ListCount > 0 Then  'si el listbox no esta vacio entonces...
           i = 0
        End If
        ' ahora recorro el lbx que contiene las obras sociales
        While (i < cant)
         
          qbusca = "INSERT INTO OBRASOCIAL_MEDICOS(COD_OBSOCIAL,COD_MLEG) " & _
                   "VALUES(" & consultaCodObSocial(Me.lbxObrasocial.List(i)) & "," & _
                    Me.lblCod_doctores & ");"

          conn.Execute qbusca

          i = i + 1   ' paso al siguiente elem

        Wend
        
        
           'si altero el orden de los pasos puede que haigan errores por el hecho que hay datos relacionados
    End Select
    
conn.CommitTrans

'limpia
clean
'actualiza la lista
actualizaLista

'*******************
desahabilitar_controles_doctores False

lblTelefono.Enabled = False
lblSexo.Enabled = False
Frmsexo.Enabled = False
cmdGuardar.Enabled = False
cmdCancelar.Enabled = False

fraInferior.Enabled = True
desabilitar_botones_alta_doctores True
lvwDoctores.Enabled = True

'deshabilita el combo de localidades el cual solo podra ser activado por el combo provincias
Me.cmbLocalidad.Enabled = False
'agrego estas dos lineas para controlar que cuando se guarde, los frames  se oculten
'aunque el usuario no los halla cerrado.
Me.FrmAgregarEspecialidad.Visible = False
Me.FrmAgregarOs.Visible = False

HuboError:
'rtaError = evaluarError(Err)
''de acuerdo a la respuesta, realiza...
'Select Case rtaError
'    Case Finalizar
'        End
'    Case Reintentar
'        Resume
'    Case Ignorar
'        Resume Next
'    Case Cancelar
'        'no hace nada
'End Select

'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    conn.RollbackTrans
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub

Private Sub cmdModificar_Click()
On Error GoTo HuboError

Controles_doctores True

'pone tipo modificacion
tipoOperacion = MODIFICACION
     
'pone los datos llamando al evento click del lvw
lvwDoctores_Click

Me.cmbLocalidad.Enabled = False

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

Private Sub cmdNuevo_Click()
On Error GoTo HuboError

Dim qbusca As String
Dim rstDatos As New ADODB.Recordset
Dim ultimo As Integer
On Error GoTo HuboError

Controles_doctores True

Me.cmbLocalidad.Enabled = False

'limpio los controles del frame superior
clean
'limpio los listbox
Me.lbxEspecialidad.Clear
Me.lbxObrasocial.Clear

'busca al ultimo de l para determinar el ID
qbusca = "SELECT MAX(M.COD_MLEG) AS ULT " & _
          "FROM MEDICOS AS M"
consultasql conn, qbusca, rstDatos

If rstDatos.EOF Then
    'no encontro ninguno (es el primero que ingreso)
    ultimo = 0
Else
    'el unico valor que trae es el maximo (ultimo ID)
    ultimo = rstDatos!ULT
End If

'pongo el id en el label correspondiente
Me.lblCod_doctores = ultimo + 1

'pone el tipo de operacion
tipoOperacion = ALTA


'pone la fecha del dia la fecha de alta y la fecha de ult modif con el mismo dia
'al igual que el usuario de alta y el de modificacion ya que el registro es nuevo

Me.dtpFechaDeAlta = Date
Me.dtpFechaDeModif = Date

Me.txtUsuarioDeAlta = usuarioActual
Me.txtUsuarioDeModif = usuarioActual

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

Private Sub cmdVolver_Click()
On Error GoTo HuboError

Unload Me

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



Private Sub lvwDoctores_Click()
'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo HuboError

'solo si hay items
If Me.lvwDoctores.ListItems.Count <> 0 Then
    'busca los datos de la Obra social seleccionada en el listview
    
    qbusca = " SELECT M.COD_MLEG, M.DNI_M, M.NOMBREM, M.DIRECCION, M.COD_POSTAL, M.TELEFONO, M.SEXO," & _
             " M.FECHA_ALTA, M.FECHA_ULTMODIF, M.COD_LOC, M.UDEALTA, M.UDEMODIF, " & _
             " OB.COD_OBSOCIAL, OB.RAZON_SOCIAL AS DESC_OBSOCIAL," & _
             " L.COD_LOC, L.DESCRIPCION AS DESCLOC , L.COD_PROV," & _
             " PRO.COD_PROV, PRO.DESCRIPCION AS DESCPROV" & _
             "" & _
             " FROM MEDICOS AS M ,OBRA_SOCIAL AS OB ,LOCALIDADES AS L ,PROVINCIAS AS PRO " & _
             " WHERE M.COD_LOC=L.COD_LOC" & _
             " AND M.COD_MLEG=" & Me.lvwDoctores.SelectedItem.Text

    
    consultasql conn, qbusca, rstDatos
    
    'si no encuentra, hay un error seguro y debe salir
    If rstDatos.EOF Then
        MsgBox "No se han encontrado los datos !!!", vbCritical, "Error"
        Exit Sub
    End If
    
    'llena los datos
    
    Me.lblCod_doctores = rstDatos!COD_MLEG
    Me.txtNombre = "" & rstDatos!NOMBREM
    Me.txtDomicilio = "" & rstDatos!DIRECCION
    Me.txtCodigo_postal = rstDatos!COD_POSTAL
    Me.txtTelefono = "" & rstDatos!TELEFONO
      If (UCase(rstDatos!SEXO) = "MASCULINO") Then
           Me.OptMasculino.Value = True
      Else
           Me.optFemenino.Value = True
      End If
    
    Me.cmbLocalidad = "" & rstDatos!DESCLOC
    Me.cmbProvincia = "" & rstDatos!DESCPROV
    Me.txtdni = rstDatos!DNI_M
    Me.txtUsuarioDeAlta = "" & rstDatos!UDEALTA
    Me.txtUsuarioDeModif = "" & rstDatos!UDEMODIF
    Me.dtpFechaDeAlta = rstDatos!FECHA_ALTA
    Me.dtpFechaDeModif = rstDatos!FECHA_ULTMODIF


    'lleno el lbxespecialidades con las especialidades del doctor selecionado
    llenarListboxEspecialidadesSegunMedico Me.lbxEspecialidad, Me.lvwDoctores.SelectedItem.Text
    
    'lleno el lbxobrassociales con las obras sociales del doctor selecionado
    llenarListboxObrasSocialesSegunMedico Me.lbxObrasocial, Me.lvwDoctores.SelectedItem.Text
    
    
End If


HuboError:
'rtaError = evaluarError(Err)
''de acuerdo a la respuesta, realiza...
'Select Case rtaError
'    Case Finalizar
'        End
'    Case Reintentar
'        Resume
'    Case Ignorar
'        Resume Next
'    Case Cancelar
'        'no hace nada
'End Select

'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub

Private Sub lvwDoctores_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'si clickeo ID ordena normal
'debo reordenar con una consulta y llenando nuevamente
'debido a que los campos numéricos no los ordena como
'tales sino que hace un orden tipo alfabético

On Error GoTo HuboError

If ColumnHeader.Index = 1 Then
    'saca la propiedad de orden
    Me.lvwDoctores.Sorted = False
    actualizaLista
Else
    'si clikea cualquier campo de texto
    'hago un orden interno al listview
    Me.lvwDoctores.AllowColumnReorder = True
    Me.lvwDoctores.Sorted = True
    Me.lvwDoctores.SortKey = ColumnHeader.SubItemIndex
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

Private Sub txtCodigo_postal_KeyPress(KeyAscii As Integer)
On Error GoTo HuboError
restriccion_numeros KeyAscii
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{tab}"
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

Private Sub txtdni_KeyPress(KeyAscii As Integer)
On Error GoTo HuboError
restriccion_numeros KeyAscii
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{tab}"
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

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
On Error GoTo HuboError
'restriccion_solo_letras KeyAscii

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


