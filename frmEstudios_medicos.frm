VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEstudios_medicos 
   BackColor       =   &H00000080&
   Caption         =   "Estudios Medicos"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   10200
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   2775
      Left            =   7200
      TabIndex        =   18
      Top             =   0
      Width           =   2895
      Begin VB.TextBox txtEstMedUsuarioDeModif 
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtEstMedUsuarioDeAlta 
         Height          =   285
         Left            =   1440
         TabIndex        =   19
         Top             =   1320
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpkEstMedFechaDeModif 
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20971521
         CurrentDate     =   39005
      End
      Begin MSComCtl2.DTPicker dtpkEsTMedFechaDeAlta 
         Height          =   255
         Left            =   1440
         TabIndex        =   22
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20971521
         CurrentDate     =   39005
      End
      Begin VB.Label lblEstMedUsuarioDeModif 
         Caption         =   "Usuario de modificacion:"
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblEstMedUsuarioDeAlta 
         Caption         =   "Usuario de Alta:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblEstMedFechaDeUltModif 
         Caption         =   "Fecha de ultima  modificacion:"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblEstMedFechaDeAlta 
         Caption         =   "Fecha de alta:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraInferior 
      Height          =   3615
      Left            =   0
      TabIndex        =   16
      Top             =   2880
      Width           =   10095
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
         Left            =   240
         TabIndex        =   0
         Top             =   360
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
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdElimina 
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
         Left            =   240
         TabIndex        =   2
         Top             =   1800
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
         Left            =   240
         TabIndex        =   3
         Top             =   2520
         Width           =   1470
      End
      Begin MSComctlLib.ListView lvwEstudios_Medicos 
         Height          =   2895
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo de estudio"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Complejidad"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "FECHA DE ALTA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "FECHA DE MODIF"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "U DE ALTA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "U DE MODIF"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fraArriva 
      Height          =   2775
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7185
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
         Left            =   3000
         TabIndex        =   8
         Top             =   2160
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
         Left            =   4920
         TabIndex        =   9
         Top             =   2160
         Width           =   1500
      End
      Begin VB.TextBox txtNombre 
         DataField       =   "NOMBRES"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   4
         Top             =   735
         Width           =   5280
      End
      Begin VB.Frame FrmDescripcion 
         Height          =   975
         Left            =   1440
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
         Begin VB.OptionButton OptAlta 
            Caption         =   "Alta"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton OptMedia 
            Caption         =   "Media"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptBaja 
            Caption         =   "Baja"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Label lblCodigo 
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
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
         TabIndex        =   15
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label lblCodEstudiosMedicos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1470
         TabIndex        =   14
         Top             =   315
         Width           =   615
      End
      Begin VB.Label lblDescripcion 
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   13
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label LblComplejidad 
         Caption         =   "Complejidad:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmEstudios_medicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tipoOperacion As Integer

Private Function datosValidos() As Boolean
On Error GoTo HuboError
    
    'chequea que al menos tenga el nombre del estudio
    If Len(Me.txtNombre) = 0 Then
        MsgBox "Debe introducir el estudio medico ", vbInformation, "Error de validaci�n"
        Me.txtNombre.SetFocus
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

Me.lblCodEstudiosMedicos = ""
Me.txtNombre = ""

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

Habilitar_estudios_medicos True
'limpia todo
clean

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

Private Sub cmdElimina_Click()
Dim res As Variant

'Manejo de Error
On Error GoTo HuboError

'pregunta antes
res = MsgBox("�Desea borrar realmente el registro de " & Me.lvwEstudios_Medicos.SelectedItem.SubItems(1) & "?", vbQuestion + vbYesNo, "Eliminar Registro")
If res = vbNo Then Exit Sub
    
conn.BeginTrans
    
conn.Execute "DELETE FROM ESTUDIOS WHERE COD_ESTMED = " & Me.lvwEstudios_Medicos.SelectedItem.Text
    
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
'Manejo de Error
On Error GoTo HuboError

    'primero verifica que sea valido
    If Not datosValidos Then Exit Sub
    
    conn.BeginTrans
    
    'guarda segun la operacion
    Select Case tipoOperacion
        Case ALTA
            conn.Execute "INSERT INTO ESTUDIOS(COD_ESTMED,DESCRIPCION,COMPLEJIDAD,FECHA_ALTA,FECHA_ULTMODIF,UDEALTA,UDEMODIF) " & _
                         "VALUES(" & Me.lblCodEstudiosMedicos & ",'" & UCase(Me.txtNombre) & "','" & _
                          detectar_complejidad & "' ,#" & _
                          Format(Me.dtpkEsTMedFechaDeAlta, "mm/dd/yyyy") & "#,#" & _
                          Format(Me.dtpkEstMedFechaDeModif, "mm/dd/yyyy") & "#," & _
                          Me.txtEstMedUsuarioDeAlta & _
                          "," & Me.txtEstMedUsuarioDeModif & ");"
                         
        Case MODIFICACION
            'lo que cambia es el usuario de modif (uso la variable global 'usuarioActual) y la
            'fecha de ult modif(uso la cte DATE que contiene la fecha actual)
            'para dejar sentado que usuario logueado es el que guarda los cambios
           
            conn.Execute "UPDATE ESTUDIOS SET " & _
                         "DESCRIPCION='" & UCase(Me.txtNombre) & "'" & _
                         ",COMPLEJIDAD= '" & detectar_complejidad & "'" & _
                         ",FECHA_ALTA=#" & Format(Me.dtpkEsTMedFechaDeAlta, "mm/dd/yyyy") & "#," & _
                         "FECHA_ULTMODIF=#" & Format(Date, "mm/dd/yyyy") & "#," & _
                         "UDEALTA=" & Me.txtEstMedUsuarioDeAlta & _
                         ",UDEMODIF=" & usuarioActual & _
                         " WHERE COD_ESTMED=" & Me.lblCodEstudiosMedicos
    End Select 'los campos que omita son los que no se actualizaran
    
    conn.CommitTrans
    
    'limpia
    clean
    'actualiza la lista
    actualizaLista
    'deshabilita la parte superior
    Habilitar_estudios_medicos True
   
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

'habilita la parte superior y desabilita la parte inferior
Habilitar_estudios_medicos False

'pone tipo modificacion
tipoOperacion = MODIFICACION
     
'pone los datos llamando al evento click del lvw
lvwEstudios_Medicos_Click


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

'habilita la parte superior
Habilitar_estudios_medicos False

'limpia
clean

'busca al ultimo de l para determinar el ID
qbusca = "SELECT MAX(EST.COD_ESTMED) AS ULT " & _
          "FROM ESTUDIOS AS EST"
consultasql conn, qbusca, rstDatos

If rstDatos.EOF Then
    'no encontro ninguno (es el primero que ingreso)
    ultimo = 0
Else
    'el unico valor que trae es el maximo (ultimo ID)
    ultimo = rstDatos!ULT
End If

'pongo el id en el label correspondiente
Me.lblCodEstudiosMedicos = ultimo + 1

'pone el tipo de operacion
tipoOperacion = ALTA


'pone la fecha del dia la fecha de alta y la fecha de ult modif
Me.dtpkEsTMedFechaDeAlta = Date
Me.dtpkEstMedFechaDeModif = Date

Me.txtEstMedUsuarioDeAlta = usuarioActual
Me.txtEstMedUsuarioDeModif = usuarioActual


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

'deshabilita la parte superior
Habilitar_estudios_medicos True

'carga el listview
actualizaLista

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
Me.lvwEstudios_Medicos.ListItems.Clear

qbusca = " SELECT EST.COD_ESTMED,EST.DESCRIPCION,EST.COMPLEJIDAD," & _
                "EST.FECHA_ALTA,EST.FECHA_ULTMODIF,EST.UDEALTA,EST.UDEMODIF" & _
         " FROM ESTUDIOS AS EST" & _
         " ORDER BY EST.COD_ESTMED"
         
consultasql conn, qbusca, rstDatos


'mientras no sea fin de archivo
While Not rstDatos.EOF
    'agrega el item a la lista
    Set mitem = Me.lvwEstudios_Medicos.ListItems.Add()
    mitem.Text = rstDatos!COD_ESTMED
    mitem.SubItems(1) = rstDatos!DESCRIPCION
    mitem.SubItems(2) = rstDatos!COMPLEJIDAD
    mitem.SubItems(3) = rstDatos!FECHA_ALTA
    mitem.SubItems(4) = rstDatos!FECHA_ULTMODIF
    mitem.SubItems(5) = rstDatos!UDEALTA
    mitem.SubItems(6) = rstDatos!UDEMODIF
    
    'avanza al siguiente registro
    rstDatos.MoveNext
Wend

'selecciona por defecto al primero encontrado
If Me.lvwEstudios_Medicos.ListItems.Count <> 0 Then
    Me.lvwEstudios_Medicos.ListItems(1).Selected = True
    'carga en la parte superior
    lvwEstudios_Medicos_Click
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

Private Sub lvwEstudios_Medicos_Click()
'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset


'Manejo de Error
On Error GoTo HuboError

'solo si hay items
If Me.lvwEstudios_Medicos.ListItems.Count <> 0 Then
    'busca los datos de la persona seleccionada
    qbusca = "SELECT EST.COD_ESTMED,EST.DESCRIPCION,EST.COMPLEJIDAD,EST.FECHA_ALTA,EST.FECHA_ULTMODIF,EST.UDEALTA,EST.UDEMODIF" & _
            " FROM ESTUDIOS AS EST" & _
            " WHERE  EST.COD_ESTMED=" & Me.lvwEstudios_Medicos.SelectedItem.Text
    
    consultasql conn, qbusca, rstDatos
    
    'si no encuentra, hay un error seguro y debe salir
    If rstDatos.EOF Then
        MsgBox "No se han encontrado los datos del estudio medico!!!", vbCritical, "Error"
        Exit Sub
    End If
    
    'llena los datos
    Me.lblCodEstudiosMedicos = rstDatos!COD_ESTMED
    Me.txtNombre = "" & rstDatos!DESCRIPCION
    
    If rstDatos!COMPLEJIDAD = "Media" Then
       Me.OptMedia.Value = True
       Me.OptBaja.Value = False
       Me.OptAlta.Value = False
    ElseIf rstDatos!COMPLEJIDAD = "Alta" Then
            Me.OptAlta.Value = True
            Me.OptBaja.Value = False
            Me.OptMedia.Value = False
         Else
            Me.OptAlta.Value = False
            Me.OptBaja.Value = True
            Me.OptMedia.Value = False
         
    End If
     
    Me.txtEstMedUsuarioDeAlta = "" & rstDatos!UDEALTA
    Me.txtEstMedUsuarioDeModif = "" & rstDatos!UDEMODIF
    Me.dtpkEsTMedFechaDeAlta = rstDatos!FECHA_ALTA
    Me.dtpkEstMedFechaDeModif = rstDatos!FECHA_ULTMODIF
    
    
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

Private Sub lvwEstudios_Medicos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'si clickeo ID ordena normal
'debo reordenar con una consulta y llenando nuevamente
'debido a que los campos num�ricos no los ordena como
'tales sino que hace un orden tipo alfab�tico

'Manejo de Error
On Error GoTo HuboError

If ColumnHeader.Index = 1 Then
    'saca la propiedad de orden
    Me.lvwEstudios_Medicos.Sorted = False
    actualizaLista
Else
    'si clikea cualquier campo de texto
    'hago un orden interno al listview
    Me.lvwEstudios_Medicos.AllowColumnReorder = True
    Me.lvwEstudios_Medicos.Sorted = True
    Me.lvwEstudios_Medicos.SortKey = ColumnHeader.SubItemIndex
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

