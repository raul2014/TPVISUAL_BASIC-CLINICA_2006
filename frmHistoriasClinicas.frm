VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHistoriasClinicas 
   BackColor       =   &H00000000&
   Caption         =   "Historias Clinicas"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   10710
   Begin VB.Frame fraSuperior 
      BackColor       =   &H00E0E0E0&
      Height          =   4095
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10455
      Begin VB.Frame fraAgregar 
         BackColor       =   &H00800000&
         Height          =   3255
         Left            =   6120
         TabIndex        =   40
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
         Begin VB.CommandButton cmdAgregarE 
            Caption         =   "Agregar Especialidad"
            Height          =   375
            Left            =   360
            TabIndex        =   43
            Top             =   2280
            Width           =   3615
         End
         Begin VB.CommandButton cmdCerrarFraAgreg 
            Caption         =   "Cerrar ventana"
            Height          =   375
            Left            =   360
            TabIndex        =   42
            Top             =   2760
            Width           =   3615
         End
         Begin VB.ListBox lbxEstudiosParaAgreg 
            Height          =   1815
            ItemData        =   "frmHistoriasClinicas.frx":0000
            Left            =   120
            List            =   "frmHistoriasClinicas.frx":0002
            TabIndex        =   41
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.CommandButton cmdQuitarEstudio 
         Caption         =   "Quitar Estudio"
         Height          =   255
         Left            =   6360
         TabIndex        =   10
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CommandButton cmdAgregarEstudio 
         Caption         =   "Agregar Estudio"
         Height          =   255
         Left            =   6360
         TabIndex        =   9
         Top             =   2280
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtpkFechaDeAtencion 
         Height          =   375
         Left            =   4440
         TabIndex        =   4
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62455809
         CurrentDate     =   39033
      End
      Begin VB.ComboBox cmbMedicos 
         BackColor       =   &H00CEECFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   1320
         Width           =   3615
      End
      Begin VB.ComboBox cmbEspecialidades 
         BackColor       =   &H00CEECFF&
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox txtDiagnostico 
         BackColor       =   &H00CEECFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Top             =   3480
         Width           =   4815
      End
      Begin VB.ComboBox cmbPacientes 
         BackColor       =   &H00CEECFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   840
         Width           =   3615
      End
      Begin VB.ListBox lbxEstudiosDeTurnos 
         BackColor       =   &H00CEECFF&
         Height          =   840
         ItemData        =   "frmHistoriasClinicas.frx":0004
         Left            =   1200
         List            =   "frmHistoriasClinicas.frx":0006
         TabIndex        =   8
         Top             =   2280
         Width           =   4815
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6480
         TabIndex        =   12
         Top             =   3600
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
         Left            =   8160
         TabIndex        =   14
         Top             =   3600
         Width           =   1500
      End
      Begin VB.Frame fraIndentifUsuario 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1935
         Left            =   7560
         TabIndex        =   17
         Top             =   120
         Width           =   2775
         Begin VB.TextBox txtTurUsuarioDeAlta 
            Height          =   285
            Left            =   1440
            TabIndex        =   19
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtTurUsuarioDeModif 
            Height          =   285
            Left            =   1920
            TabIndex        =   18
            Top             =   1440
            Width           =   735
         End
         Begin MSComCtl2.DTPicker dtpkTurFechaDeModif 
            Height          =   255
            Left            =   1320
            TabIndex        =   20
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   62455809
            CurrentDate     =   39005
         End
         Begin MSComCtl2.DTPicker dtpkTurFechaDeAlta 
            Height          =   255
            Left            =   1320
            TabIndex        =   21
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   62455809
            CurrentDate     =   39005
         End
         Begin VB.Label lblTurFechaDeAlta 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fecha de alta:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblTurFechaDeUltModif 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fecha de ultima  modificacion:"
            Height          =   495
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblTurUsuarioDeAlta 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Usuario de Alta:"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label lblTurUsuarioDeModif 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Usuario de modificacion:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Width           =   1815
         End
      End
      Begin VB.Label lblMedico 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Medico:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblCodMed 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cod Medico:"
         Height          =   375
         Left            =   4920
         TabIndex        =   38
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblCodEsp 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   6000
         TabIndex        =   37
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblCodEspecialidad 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cod Especialidad:"
         Height          =   375
         Left            =   4920
         TabIndex        =   36
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblEspecialidad 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Especialidad"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblDiagnostico 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Diagnostico:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label lblFechaDeAtencion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha de Atencion:"
         Height          =   375
         Left            =   2880
         TabIndex        =   33
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblCodP 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   5640
         TabIndex        =   32
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblCodPaciente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cod Paciente:"
         Height          =   375
         Left            =   4800
         TabIndex        =   31
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblCodM 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   5640
         TabIndex        =   30
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblEstudios 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Estudios"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblPaciente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Paciente:"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblCodTurno 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   840
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblCod_Turno 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cod:"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraInferior 
      BackColor       =   &H00E0E0E0&
      Height          =   3855
      Left            =   0
      TabIndex        =   13
      Top             =   4080
      Width           =   10455
      Begin MSComctlLib.ListView lvwTurnos 
         Height          =   2775
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   13561087
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "COD TURNO"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "F DE ATENCION"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "DIAGNOSTICO"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "COD DE PACIENTE"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "COD DE MEDICO"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "COD DE ESPECIALIDAD"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "F DE ALTA"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "F DE MODIFICACION"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "U DE ALTA"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "U DE MODIF"
            Object.Width           =   2646
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
         Left            =   7320
         TabIndex        =   3
         Top             =   3240
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
         Left            =   5400
         TabIndex        =   2
         Top             =   3240
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
         Top             =   3240
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
         Left            =   1080
         TabIndex        =   0
         Top             =   3240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmHistoriasClinicas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tipoOperacion As Integer

Private Function datosValidos() As Boolean
On Error GoTo HuboError
' chequea que se halla ingresado algun estudio medico
    If Me.lbxEstudiosDeTurnos.ListCount = 0 Then
       MsgBox "Debe introducir algun estudio medico", vbInformation, "Error de validacion"
       datosValidos = False
       Exit Function
    End If
    
 'chequea que se halla ingresado el diagnostico
    If Len(Me.txtDiagnostico) = 0 Then
        MsgBox "Debe introducir el diagnostico", vbInformation, "Error de validación"
        Me.txtDiagnostico.SetFocus
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

Private Sub clean()
On Error GoTo HuboError

Me.lblCodTurno = ""
Me.dtpkFechaDeAtencion = Date

Me.txtDiagnostico = ""
Me.lbxEstudiosDeTurnos.Clear

Me.lblCodEsp.Caption = ""
Me.lblCodM.Caption = ""
Me.lblCodP.Caption = ""

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

Private Sub cmbEspecialidades_Click()
On Error GoTo HuboError

Me.lblCodEsp.Caption = consultaCodEspecialidad(Me.cmbEspecialidades.Text)

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

Private Sub cmbMedicos_Click()
Dim qbusca As String    'se llenara el combo de ESPECIALIDADES segun el medico seleccionado
Dim rstDatos As New ADODB.Recordset

On Error GoTo HuboError

'limpia el combo
Me.cmbEspecialidades.Clear

'Hago la consulta
qbusca = " SELECT DISTINCT ESP.DESCRIPCION AS DESCESP " & _
         " FROM MEDICOS AS M, ESPECIALIDADES AS ESP,MEDICOS_ESPECIALIDAD AS MEDESP" & _
         " WHERE MEDESP.COD_MLEG=" & consultaCodMedico(Me.cmbMedicos.Text) & _
         " AND MEDESP.COD_ESP=ESP.COD_ESP" & _
         " AND M.COD_MLEG=MEDESP.COD_MLEG" & _
         " ORDER BY ESP.DESCRIPCION "
         
consultasql conn, qbusca, rstDatos

'mientras no sea fin de archivo
While Not rstDatos.EOF
    'carga los items al combo
    Me.cmbEspecialidades.AddItem rstDatos!DESCESP
    'avanza al registro siguiente
    rstDatos.MoveNext
Wend

Me.lblCodM.Caption = consultaCodMedico(Me.cmbMedicos.Text)
Me.cmbEspecialidades.Enabled = True

HuboError:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    conn.RollbackTrans
    MsgBox "Error: " + Err.Description
    Exit Sub
End If

End Sub

Private Sub cmbPacientes_Click()
On Error GoTo HuboError

Me.lblCodP.Caption = consultaCodPaciente(Me.cmbPacientes.Text)

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
Dim r As Long
On Error GoTo HuboError

r = controlarNoRepetidos(Me.lbxEstudiosDeTurnos, Me.lbxEstudiosParaAgreg.Text)
If (r <> 1) Then
    Me.lbxEstudiosDeTurnos.AddItem Me.lbxEstudiosParaAgreg.Text
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

Private Sub cmdAgregarEstudio_Click()
On Error GoTo HuboError

Me.fraAgregar.Visible = True

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

Controles_HistoriasClinicas False

Me.cmbEspecialidades.Enabled = False
'limpia todo
clean

'oculta el frame de agregar estudios por si se olvidaron cerrarlo
Me.fraAgregar.Visible = False

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

Private Sub cmdCerrarFraAgreg_Click()
On Error GoTo HuboError

Me.fraAgregar.Visible = False

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
res = MsgBox("¿Desea borrar realmente el registro de codigo: " & Me.lvwTurnos.SelectedItem.Text & "?", vbQuestion + vbYesNo, "Eliminar Registro")
If res = vbNo Then Exit Sub
    
conn.BeginTrans
    

' 1ro elimino los estudios de la tabla ESTUDIO_TURNOS (SI NO VA HABER PROBLEMA DE QUE HAY REG RELACIONADOS)
conn.Execute "DELETE FROM ESTUDIO_TURNOS WHERE COD_TURNO = " & Me.lvwTurnos.SelectedItem.Text
    
'Luego elimino el registro que corresponde al turno seleccionado en el listview
conn.Execute "DELETE FROM TURNOS WHERE COD_TURNO = " & Me.lvwTurnos.SelectedItem.Text
    

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
          sql = "INSERT INTO TURNOS(COD_TURNO,F_DE_ATENCION,DIAGNOSTICO,FECHA_ALTA,FECHA_ULTMODIF,COD_LEGP,COD_LEGM,COD_ESP,UDEALTA,UDEMODIF) " & _
                         "VALUES(" & Me.lblCodTurno & ",#" & Format(Me.dtpkFechaDeAtencion, "mm/dd/yyyy") & "# ,'" & _
                          Me.txtDiagnostico & "',#" & _
                          Format(Me.dtpkTurFechaDeAlta, "mm/dd/yyyy") & "#,#" & _
                          Format(Me.dtpkTurFechaDeModif, "mm/dd/yyyy") & "#," & _
                          Me.lblCodP & "," & Me.lblCodM & "," & Me.lblCodEsp & "," & _
                          Me.txtTurUsuarioDeAlta & "," & _
                          Me.txtTurUsuarioDeModif & ");"
                          
          
        conn.Execute sql
        
        'ahora grabo los estudios que correspondes al turno de la consulta SQL
        
          '//
        cant = Me.lbxEstudiosDeTurnos.ListCount 'cuento la cant de elem del lbx de estudios

        If Me.lbxEstudiosDeTurnos.ListCount > 0 Then  'si el listbox no esta vacio entonces...
           i = 0
        End If
        ' ahora recorro el lbx que contiene los estudios medicos
        While (i < cant)
         
          qbusca = "INSERT INTO ESTUDIO_TURNOS(COD_TURNO,COD_ESTMED) " & _
                   "VALUES(" & Me.lblCodTurno & "," & _
                    consultaCodDeESTUDIO(Me.lbxEstudiosDeTurnos.List(i)) & ");"

          conn.Execute qbusca

          i = i + 1   ' paso al siguiente elem

        Wend
          
       '  conn.Execute sql
       
    Case MODIFICACION
            'lo que cambia es el usuario de modif (uso la variable global 'usuarioActual) y la
            'fecha de ult modif(uso la cte DATE que contiene la fecha actual)
            'para dejar sentado que usuario logueado es el que guarda los cambios
           'LOS CAMPOS QUE OMITA SON LOS QUE NO SE ACTUALIZAN
           
        sql = "UPDATE TURNOS SET " & _
                         "F_DE_ATENCION=#" & Format(Me.dtpkFechaDeAtencion, "mm/dd/yyyy") & "#," & _
                         "DIAGNOSTICO='" & UCase(Me.txtDiagnostico) & "'," & _
                         "FECHA_ULTMODIF=#" & Format(Date, "mm/dd/yyyy") & "#," & _
                         "COD_LEGP=" & Me.lblCodP & "," & _
                         "COD_LEGM=" & Me.lblCodM & "," & _
                         "COD_ESP=" & Me.lblCodEsp & "," & _
                         "UDEMODIF=" & usuarioActual & _
                         " WHERE COD_TURNO=" & Me.lblCodTurno
   
       '1ro actualizo los datos de la tabla TURNOS
        conn.Execute sql
            
        '2do borro los estudios correspondientes al turno que se va a ACTUALIZAR
        'es decir elimino los estudios de la tabla ESTUDIO_TURNOS ya que como
        'se agegaran o quitaran estudios en el lbx entonces borro los del que quiero modif
        'y luego insertare (ver 3ro) los que hallan quedado en el lbx de de estudios.
        
        conn.Execute "DELETE FROM ESTUDIO_TURNOS WHERE COD_TURNO = " & Me.lvwTurnos.SelectedItem.Text
        
        '3ro inserto los estudios definitivos del lbx estudios
        
        cant = Me.lbxEstudiosDeTurnos.ListCount 'cuento la cant de elem del lbx de especialidades

        If Me.lbxEstudiosDeTurnos.ListCount > 0 Then  'si el listbox no esta vacio entonces...
           i = 0
        End If
        ' ahora recorro el lbx que contiene los estudios medicos
        While (i < cant)
         
          qbusca = "INSERT INTO ESTUDIO_TURNOS(COD_TURNO,COD_ESTMED) " & _
                   "VALUES(" & Me.lblCodTurno & "," & _
                    consultaCodDeESTUDIO(Me.lbxEstudiosDeTurnos.List(i)) & ");"

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
'deshabilita la parte superior
Controles_HistoriasClinicas False
   
'bloquea el combo de especialidades
Me.cmbEspecialidades.Enabled = False

'oculta el frame de agregar estudios por si se olvidaron cerrarlo
Me.fraAgregar.Visible = False

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
'pone el tipo de operacion
tipoOperacion = MODIFICACION

Controles_HistoriasClinicas True

Me.cmbEspecialidades.Enabled = False

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
'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset
Dim ultimo As Long

On Error GoTo HuboError

Controles_HistoriasClinicas True

Me.cmbEspecialidades.Enabled = False

'limpia
clean

'busca al ultimo de l para determinar el ID
qbusca = "SELECT MAX(T.COD_TURNO) AS ULT " & _
          "FROM TURNOS AS T"
consultasql conn, qbusca, rstDatos

If rstDatos.EOF Then
    'no encontro ninguno (es el primero que ingreso)
    ultimo = 0
Else
    'el unico valor que trae es el maximo (ultimo ID)
    ultimo = rstDatos!ULT
End If

'pongo el id en el label correspondiente
Me.lblCodTurno = ultimo + 1

'pone el tipo de operacion
tipoOperacion = ALTA


'pone la fecha del dia la fecha de alta y la fecha de ult modif
Me.dtpkTurFechaDeAlta = Date
Me.dtpkTurFechaDeModif = Date

Me.txtTurUsuarioDeAlta = usuarioActual
Me.txtTurUsuarioDeModif = usuarioActual


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

Private Sub cmdQuitarEstudio_Click()
On Error GoTo HuboError

'si no selecciono nada, devuelve indice -1
    If Me.lbxEstudiosDeTurnos.ListIndex = -1 Then
        If Me.lbxEstudiosDeTurnos.ListCount = 0 Then
            MsgBox "No hay items cargados para poder eliminar", vbInformation, "Error"
        Else
            MsgBox "Debe seleccionar algún item para poder eliminar", vbInformation, "Error"
        End If
        Exit Sub
    End If
    Me.lbxEstudiosDeTurnos.RemoveItem Me.lbxEstudiosDeTurnos.ListIndex

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
'Manejo de Error
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
'Manejo de Error
On Error GoTo HuboError

'cargo el listview
actualizaLista

Controles_HistoriasClinicas False

'lleno el combo de medicos
cargaComboMedicos Me.cmbMedicos

'lleno el combo de pacientes
cargaComboPacientes Me.cmbPacientes

'lleno el combo de especialidades
cargaComboEspecialidades Me.cmbEspecialidades

'LLENO EL LISTBOX DE ESTUDIOS con todos los estudios de la tabla estudios
llenarListboxEstudios Me.lbxEstudiosParaAgreg

Me.cmbEspecialidades.Enabled = False
Me.dtpkFechaDeAtencion = Format(Date, "mm/dd/yyyy")

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
On Error GoTo HuboError

'limpia la lista
Me.lvwTurnos.ListItems.Clear

'Hago la consulta tomando el filtro del textBox de Busqueda
qbusca = " SELECT T.COD_TURNO,T.F_DE_ATENCION,T.DIAGNOSTICO,T.COD_LEGP,T.COD_LEGM," & _
                  "T.COD_ESP,T.FECHA_ALTA,T.FECHA_ULTMODIF,T.UDEALTA,T.UDEMODIF" & _
         " FROM TURNOS AS T" & _
         " ORDER BY T.COD_TURNO"
consultasql conn, qbusca, rstDatos


'mientras no sea fin de archivo
While Not rstDatos.EOF
    'agrega el item a la lista
    Set mitem = Me.lvwTurnos.ListItems.Add()
    mitem.Text = rstDatos!COD_TURNO
    mitem.SubItems(1) = rstDatos!F_DE_ATENCION
    mitem.SubItems(2) = rstDatos!DIAGNOSTICO
    mitem.SubItems(3) = rstDatos!COD_LEGP
    mitem.SubItems(4) = rstDatos!COD_LEGM
    mitem.SubItems(5) = rstDatos!COD_ESP
    mitem.SubItems(6) = rstDatos!FECHA_ALTA
    mitem.SubItems(7) = rstDatos!FECHA_ULTMODIF
    mitem.SubItems(8) = rstDatos!UDEALTA
    mitem.SubItems(9) = rstDatos!UDEMODIF
    
    
    'avanza al siguiente registro
    rstDatos.MoveNext
Wend

'selecciona por defecto al primero encontrado
If Me.lvwTurnos.ListItems.Count <> 0 Then
    Me.lvwTurnos.ListItems(1).Selected = True
    'carga en la parte superior
    lvwTurnos_Click
Else
    'si no hay nada, limpia
    'clean
End If

HuboError:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If

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

Private Sub lvwTurnos_Click()
'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo HuboError

'solo si hay items
If Me.lvwTurnos.ListItems.Count <> 0 Then
    'busca los datos de la persona seleccionada
    qbusca = "SELECT T.COD_TURNO,T.F_DE_ATENCION,T.DIAGNOSTICO,T.FECHA_ALTA,T.FECHA_ULTMODIF," & _
                    "T.COD_LEGP,T.COD_LEGM,T.COD_ESP,T.UDEALTA,T.UDEMODIF," & _
                    "P.NOMBRE, M.NOMBREM,ESP.DESCRIPCION " & _
            " FROM TURNOS AS T ,PACIENTES AS P, MEDICOS AS M,ESPECIALIDADES AS ESP" & _
            " WHERE P.COD_LEGP=T.COD_LEGP" & _
            " AND   M.COD_MLEG=T.COD_LEGM" & _
            " AND   ESP.COD_ESP=T.COD_ESP" & _
            " AND T.COD_TURNO=" & Me.lvwTurnos.SelectedItem.Text
            
    consultasql conn, qbusca, rstDatos
    
    'si no encuentra, hay un error seguro y debe salir
    If rstDatos.EOF Then
        MsgBox "No se han encontrado los datos !!!", vbCritical, "Error"
        Exit Sub
    End If
    
    'llena los datos
    Me.lblCodTurno = rstDatos!COD_TURNO
    Me.dtpkFechaDeAtencion = "" & rstDatos!F_DE_ATENCION
    Me.txtDiagnostico = "" & rstDatos!DIAGNOSTICO
    Me.lblCodM = rstDatos!COD_LEGM
    Me.lblCodP = rstDatos!COD_LEGP
    Me.lblCodEsp = rstDatos!COD_ESP
    Me.cmbPacientes = "" & rstDatos!NOMBRE
    Me.cmbMedicos = "" & rstDatos!NOMBREM
    Me.cmbEspecialidades = "" & rstDatos!DESCRIPCION
    Me.dtpkTurFechaDeAlta = "" & rstDatos!FECHA_ALTA
    Me.dtpkTurFechaDeModif = "" & rstDatos!FECHA_ULTMODIF
    Me.txtTurUsuarioDeAlta = rstDatos!UDEALTA
    Me.txtTurUsuarioDeModif = rstDatos!UDEMODIF
    'lleno el listbox de estudios
    llenarListboxEstudiosSegunPaciente Me.lbxEstudiosDeTurnos, rstDatos!COD_LEGP
    
End If

HuboError:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If

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

Private Sub lvwTurnos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'si clickeo ID ordena normal
'debo reordenar con una consulta y llenando nuevamente
'debido a que los campos numéricos no los ordena como
'tales sino que hace un orden tipo alfabético

'Manejo de Error
On Error GoTo HuboError

If ColumnHeader.Index = 1 Then
    'saca la propiedad de orden
    Me.lvwTurnos.Sorted = False
    
    actualizaLista
Else
    'si clikea cualquier campo de texto
    'hago un orden interno al listview
    Me.lvwTurnos.AllowColumnReorder = True
    Me.lvwTurnos.Sorted = True
    Me.lvwTurnos.SortKey = ColumnHeader.SubItemIndex
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


