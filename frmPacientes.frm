VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPacientes 
   BackColor       =   &H00800000&
   Caption         =   "Pacientes"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   10095
   Begin VB.Frame FraUsuario 
      BackColor       =   &H00E0E0E0&
      Height          =   3375
      Left            =   7800
      TabIndex        =   28
      Top             =   0
      Width           =   2175
      Begin VB.TextBox txtUsuarioDeModif 
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtUsuarioDeAlta 
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   2880
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtpFechaDeModif 
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20905985
         CurrentDate     =   39005
      End
      Begin MSComCtl2.DTPicker dtpFechaDeAlta 
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20905985
         CurrentDate     =   39005
      End
      Begin VB.Label lblUsuarioDeModif 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Usuario de modificacion:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblUsuarioDeAlta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Usuario de Alta:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblFechaDeUltModif 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha de ultima  modificacion:"
         Height          =   495
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblFechaDeAlta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha de alta:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame fraAbajo 
      BackColor       =   &H00E0E0E0&
      Height          =   3855
      Left            =   0
      TabIndex        =   25
      Top             =   3360
      Width           =   9975
      Begin MSComctlLib.ListView lvwPacientes 
         Height          =   2655
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15658717
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro de legajo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "DNI"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Direccion"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Codigo postal"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Telefono"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "FECHA ALTA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "FECHA ULT MODIF"
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
         Left            =   6960
         TabIndex        =   3
         Top             =   3120
         Width           =   1590
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
         Left            =   4800
         TabIndex        =   2
         Top             =   3120
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
         Left            =   2640
         TabIndex        =   1
         Top             =   3120
         Width           =   1455
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
         Left            =   480
         TabIndex        =   0
         Top             =   3120
         Width           =   1455
      End
   End
   Begin VB.Frame fraArriva 
      BackColor       =   &H00E0E0E0&
      Height          =   3375
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7785
      Begin VB.TextBox txtCodigo_postal 
         DataField       =   "TELEFONO"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1200
         Width           =   885
      End
      Begin VB.Frame frasexo 
         BackColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   6120
         TabIndex        =   38
         Top             =   960
         Width           =   1335
         Begin VB.OptionButton OptMasculino 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Masculino"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton OptFemenino 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Femenino"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtSexo 
         Height          =   285
         Left            =   4920
         TabIndex        =   37
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmbProvincia 
         DataField       =   "Provincia"
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Text            =   "cmbProvincia"
         Top             =   2160
         Width           =   2955
      End
      Begin VB.TextBox txtDNI 
         DataField       =   "DNI"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   5880
         MaxLength       =   8
         TabIndex        =   5
         Top             =   600
         Width           =   1545
      End
      Begin VB.ComboBox cmbLocalidad 
         DataField       =   "Localidad"
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Text            =   "cmbLocalidad"
         Top             =   2760
         Width           =   2940
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
         Left            =   4320
         TabIndex        =   14
         Top             =   2760
         Width           =   1590
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
         Left            =   6120
         TabIndex        =   15
         Top             =   2760
         Width           =   1500
      End
      Begin VB.TextBox txtTelefono 
         DataField       =   "TELEFONO"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   3120
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1200
         Width           =   1365
      End
      Begin VB.TextBox txtDomicilio 
         DataField       =   "DIRECCION"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         DataField       =   "NOMBRE"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1110
         MaxLength       =   30
         TabIndex        =   4
         Top             =   600
         Width           =   3345
      End
      Begin VB.ComboBox CmbObra_social 
         DataField       =   "Obra Social"
         Height          =   315
         Left            =   5760
         TabIndex        =   13
         Text            =   "CmbObra_social"
         Top             =   1920
         Width           =   1785
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
         Left            =   240
         TabIndex        =   40
         Top             =   1080
         Width           =   735
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
         Left            =   5520
         TabIndex        =   39
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblCod_Paciente 
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   120
         Width           =   735
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
         TabIndex        =   24
         Top             =   2160
         Width           =   975
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
         TabIndex        =   23
         Top             =   2760
         Width           =   1095
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
         TabIndex        =   22
         Top             =   120
         Width           =   1125
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
         Left            =   2280
         TabIndex        =   21
         Top             =   1200
         Width           =   855
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
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nombre y Apellido:"
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
         Left            =   240
         TabIndex        =   19
         Top             =   480
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
         Left            =   5160
         TabIndex        =   18
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblObra_Social 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Obra social:"
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
         Left            =   4680
         TabIndex        =   17
         Top             =   1920
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmPacientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tipoOperacion As Integer

Private Sub clean()
On Error GoTo HuboError

Me.lblCod_Paciente = ""
Me.txtNombre = ""
Me.txtdni = ""
Me.txtDomicilio = ""
Me.txtTelefono = ""
Me.txtCodigo_postal = ""

'si no esta vacio el combo de provincias
If Me.cmbProvincia.ListCount <> 0 Then cmbProvincia.ListIndex = 0

'si no esta vacio el combo de localidades
If Me.cmbLocalidad.ListCount <> 0 Then cmbLocalidad.ListIndex = 0

'si no esta vacio el combo de Obras sociales
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

    'chequea que se halla ingresado el nombre
    If Len(Me.txtNombre) = 0 Then
        MsgBox "Debe introducir el nombre ", vbInformation, "Error de validación"
        Me.txtNombre.SetFocus
        datosValidos = False
        Exit Function
    End If
    
     'chequea que se halla ingresado el dni
    If Len(Me.txtdni) = 0 Then
        MsgBox "Debe introducir el nro de dni", vbInformation, "Error de validación"
        Me.txtdni.SetFocus
        datosValidos = False
        Exit Function
    End If
    
     'chequea que se halla ingresado el domicilio
    If Len(Me.txtDomicilio) = 0 Then
        MsgBox "Debe introducir el domicilio", vbInformation, "Error de validación"
        Me.txtDomicilio.SetFocus
        datosValidos = False
        Exit Function
    End If
    
     'chequea que se halla ingresado el codigo postal
    If Len(Me.txtCodigo_postal) = 0 Then
        MsgBox "Debe introducir el codigo postal", vbInformation, "Error de validación"
        Me.txtCodigo_postal.SetFocus
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

Private Sub cmdCancelar_Click()
On Error GoTo HuboError

habilitar_pacientes True
'limpia todo
clean

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

Private Sub cmdEliminar_Click()
Dim res As Variant

'Manejo de Error
On Error GoTo HuboError

'pregunta antes
res = MsgBox("¿Desea borrar realmente el registro de " & Me.lvwPacientes.SelectedItem.SubItems(1) & "?", vbQuestion + vbYesNo, "Eliminar Registro")
If res = vbNo Then Exit Sub
    
conn.BeginTrans
    
conn.Execute "DELETE FROM PACIENTES WHERE COD_LEGP = " & Me.lvwPacientes.SelectedItem.Text
    
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
End Sub

Private Sub cmdGuardar_Click()
Dim sql As String
On Error GoTo HuboError

If (OptMasculino = True) Then
 TxtSexo.Text = "MASCULINO"
Else
 TxtSexo.Text = "FEMENINO"
End If
    
    'primero verifica que sea valido
    If Not datosValidos Then Exit Sub

    conn.BeginTrans

    'guarda segun la operacion
    Select Case tipoOperacion
        Case ALTA
             sql = "INSERT INTO PACIENTES(COD_LEGP,DNI,NOMBRE,DIRECCION,COD_POSTAL,TELEFONO,SEXO,FECHA_ALTA,FECHA_ULTMODIF,COD_LOC, COD_OBSOCIAL,UDEALTA,UDEMODIF) " & _
                         "VALUES(" & Int(Me.lblCod_Paciente) & "," & Int(Me.txtdni) & ",'" & UCase(Me.txtNombre) & "','" & UCase(Me.txtDomicilio) & "'," & Me.txtCodigo_postal & "," & Int(Me.txtTelefono) & ",'" & UCase(Me.TxtSexo) & "',#" & _
                          Format(Me.dtpFechaDeAlta, "mm/dd/yyyy") & "#,#" & _
                          Format(Me.dtpFechaDeModif, "mm/dd/yyyy") & "#," & _
                          consultaCodLocalidad(Me.cmbLocalidad) & "," & _
                          consultaCodObSocial(Me.CmbObra_social) & "," & _
                          Me.txtUsuarioDeAlta & "," & _
                          Me.txtUsuarioDeModif & ");"


            conn.Execute sql
            

        Case MODIFICACION
            'lo que cambia es el usuario de modif (uso la variable global 'usuarioActual) y la
            'fecha de ult modif(uso la cte DATE que contiene la fecha actual)
            'para dejar sentado que usuario logueado es el que guarda los cambios
 
            sql = "UPDATE PACIENTES SET " & _
                         " DNI=" & Int(Me.txtdni) & "," & _
                         " NOMBRE='" & UCase(Me.txtNombre) & "'," & _
                         " DIRECCION='" & UCase(Me.txtDomicilio) & "'," & _
                         " COD_POSTAL=" & Int(Me.txtCodigo_postal) & "," & _
                         " TELEFONO ='" & Int(Me.txtTelefono) & "'," & _
                         " SEXO='" & UCase(Me.TxtSexo) & "'," & _
                         " FECHA_ULTMODIF=#" & Format(Date, "mm/dd/yyyy") & "#," & _
                         " COD_LOC= " & consultaCodLocalidad(Me.cmbLocalidad) & "," & _
                         " COD_OBSOCIAL=" & consultaCodObSocial(Me.CmbObra_social) & "," & _
                         " UDEMODIF=" & usuarioActual & _
                         " WHERE COD_LEGP=" & Int(Me.lblCod_Paciente)
                         
            conn.Execute sql
 
 
      End Select
    
    conn.CommitTrans
 
    'limpia
    clean
    'actualiza la lista
    actualizaLista
    'deshabilita la parte superior
    habilitar_pacientes True

   
   Me.cmbLocalidad.Enabled = False
    
HuboError:
'rtaError = evaluarError(Err)
' Hubo errores
'    conn.RollbackTrans
'de acuerdo a la respuesta, realiza...
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

habilitar_pacientes False

'pone tipo modificacion
tipoOperacion = MODIFICACION
     
'pone los datos llamando al evento click del lvw
lvwPacientes_Click


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
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset
Dim ultimo As Integer
On Error GoTo HuboError

habilitar_pacientes False

Me.cmbLocalidad.Enabled = False

'limpia todo
clean

'busca al ultimo de l para determinar el ID
qbusca = "SELECT MAX(P.COD_LEGP) AS ULT " & _
          "FROM PACIENTES AS P"
consultasql conn, qbusca, rstDatos

If rstDatos.EOF Then
    'no encontro ninguno (es el primero que ingreso)
    ultimo = 0
Else
    'el unico valor que trae es el maximo (ultimo ID)
    ultimo = rstDatos!ULT
End If

'pongo el id en el label correspondiente
Me.lblCod_Paciente = ultimo + 1

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


Private Sub Form_Load()
On Error GoTo HuboError
FraUsuario.Enabled = False
                                  
'llena el combo localidad con todas las provincias
cargaComboObrasSociales Me.CmbObra_social

'la carga de combos provincia y localidades sera segun la jerarquia de los datos a obtener
                                  
'1ro llena el combo localidad con todas las provincias
cargaComboProvincias cmbProvincia

'2do llena el combo localidad con todas las localidades
cargaComboLocalidades cmbLocalidad



'deshabilita la parte superior
habilitar_pacientes True

'carga el listview
actualizaLista

Me.cmbLocalidad.Enabled = False
Me.CmbObra_social.Enabled = False

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

Private Sub actualizaLista()
'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset
Dim mitem As ListItem

'Manejo de Error
On Error GoTo error

'limpia la lista
lvwPacientes.ListItems.Clear

'Hago la consulta tomando el filtro del textBox de Busqueda

qbusca = " SELECT P.COD_LEGP,P.DNI,P.NOMBRE,P.DIRECCION,P.COD_POSTAL,P.TELEFONO," & _
                 "P.FECHA_ALTA,P.FECHA_ULTMODIF,P.UDEALTA,P.UDEMODIF" & _
         " FROM PACIENTES AS P " & _
         " ORDER BY P.COD_LEGP"
         
consultasql conn, qbusca, rstDatos

rstDatos.MoveFirst
'mientras no sea fin de archivo
While Not rstDatos.EOF
    'agrega el item a la lista
    Set mitem = lvwPacientes.ListItems.Add()
    mitem.Text = rstDatos!COD_LEGP
    mitem.SubItems(1) = rstDatos!NOMBRE
    mitem.SubItems(2) = rstDatos!DNI
    mitem.SubItems(3) = rstDatos!DIRECCION
    mitem.SubItems(4) = rstDatos!COD_POSTAL
    mitem.SubItems(5) = "" & rstDatos!TELEFONO
    mitem.SubItems(6) = "" & rstDatos!FECHA_ALTA
    mitem.SubItems(7) = "" & rstDatos!FECHA_ULTMODIF
    mitem.SubItems(8) = "" & rstDatos!UDEALTA
    mitem.SubItems(9) = "" & rstDatos!UDEMODIF
    rstDatos.MoveNext
Wend

'selecciona por defecto al primero encontrado
If Me.lvwPacientes.ListItems.Count <> 0 Then
    Me.lvwPacientes.ListItems(1).Selected = True
    'carga en la parte superior
    lvwPacientes_Click
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

Private Sub lvwPacientes_Click()
'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo HuboError

'solo si hay items
If Me.lvwPacientes.ListItems.Count <> 0 Then
    'busca los datos de la Obra social seleccionada en el listview
    
    qbusca = " SELECT P.COD_LEGP, P.DNI,P.NOMBRE, P.DIRECCION, P.COD_POSTAL, P.TELEFONO, P.SEXO," & _
             " P.FECHA_ALTA, P.FECHA_ULTMODIF, P.COD_LOC, P.COD_OBSOCIAL, P.UDEALTA, P.UDEMODIF, " & _
             " OB.COD_OBSOCIAL, OB.RAZON_SOCIAL AS DESC_OBSOCIAL,OB.COD_LOC," & _
             " L.COD_LOC, L.DESCRIPCION AS DESCLOC , L.COD_PROV," & _
             " PRO.COD_PROV, PRO.DESCRIPCION AS DESCPROV" & _
             "" & _
             " FROM PACIENTES AS P ,OBRA_SOCIAL AS OB ,LOCALIDADES AS L ,PROVINCIAS AS PRO " & _
             " WHERE PRO.COD_PROV=L.COD_PROV" & _
             " AND P.COD_LOC=L.COD_LOC" & _
             " AND P.COD_OBSOCIAL=OB.COD_OBSOCIAL" & _
             " AND P.COD_LEGP=" & Me.lvwPacientes.SelectedItem.Text

    
    consultasql conn, qbusca, rstDatos
    
    'si no encuentra, hay un error seguro y debe salir
    If rstDatos.EOF Then
        MsgBox "No se han encontrado los datos !!!", vbCritical, "Error"
        Exit Sub
    End If
    
    'llena los datos
    Me.lblCod_Paciente = rstDatos!COD_LEGP
    Me.txtNombre = "" & rstDatos!NOMBRE
    Me.txtDomicilio = "" & rstDatos!DIRECCION
    Me.txtCodigo_postal = rstDatos!COD_POSTAL
    Me.txtTelefono = "" & rstDatos!TELEFONO
      If (rstDatos!SEXO = "MASCULINO") Then
           OptMasculino.Value = True
      Else
           optFemenino.Value = True
      End If
    
    Me.cmbProvincia = "" & rstDatos!DESCPROV
    Me.cmbLocalidad = "" & rstDatos!DESCLOC
    Me.CmbObra_social = "" & rstDatos!DESC_OBSOCIAL
    Me.txtdni = rstDatos!DNI
    Me.txtUsuarioDeAlta = "" & rstDatos!UDEALTA
    Me.txtUsuarioDeModif = "" & rstDatos!UDEMODIF
    Me.dtpFechaDeAlta = rstDatos!FECHA_ALTA
    Me.dtpFechaDeModif = rstDatos!FECHA_ULTMODIF
    
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

Private Sub lvwPacientes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'si clickeo ID ordena normal
'debo reordenar con una consulta y llenando nuevamente
'debido a que los campos numéricos no los ordena como
'tales sino que hace un orden tipo alfabético

On Error GoTo HuboError

If ColumnHeader.Index = 1 Then
    'saca la propiedad de orden
    Me.lvwPacientes.Sorted = False
    actualizaLista
Else
    'si clikea cualquier campo de texto
    'hago un orden interno al listview
    Me.lvwPacientes.AllowColumnReorder = True
    Me.lvwPacientes.Sorted = True
    Me.lvwPacientes.SortKey = ColumnHeader.SubItemIndex
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

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
On Error GoTo HuboError

restriccion_numeros KeyAscii

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


