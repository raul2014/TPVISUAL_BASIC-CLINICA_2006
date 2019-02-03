VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmUsuarios 
   BackColor       =   &H00000000&
   Caption         =   "Usuarios"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   10155
   Begin VB.Frame frasuperior 
      BackColor       =   &H80000001&
      Height          =   5055
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9975
      Begin VB.CommandButton cmdGuardarU 
         Caption         =   "Guardar"
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   4440
         Width           =   2535
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   3120
         TabIndex        =   13
         Top             =   4440
         Width           =   2535
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         Left            =   6840
         TabIndex        =   9
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox txtUsername 
         Height          =   375
         Left            =   6840
         TabIndex        =   8
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Frame fraCategoria 
         Caption         =   "Categoria"
         Height          =   975
         Left            =   5640
         TabIndex        =   23
         Top             =   3000
         Width           =   4215
         Begin VB.OptionButton optUsuario 
            Caption         =   "USUARIO"
            Height          =   255
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optAdministrador 
            Caption         =   "ADMINISTRADOR"
            Height          =   255
            Left            =   2160
            TabIndex        =   11
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.TextBox txtNombre 
         Height          =   375
         Left            =   960
         TabIndex        =   7
         Top             =   3360
         Width           =   3375
      End
      Begin VB.Frame fraMedico 
         Caption         =   "Medico"
         Height          =   2055
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   4455
         Begin VB.ComboBox cmbListaMedicos 
            Height          =   315
            ItemData        =   "frmUsuarios.frx":0000
            Left            =   240
            List            =   "frmUsuarios.frx":0002
            TabIndex        =   6
            Text            =   "cmbListaMedicos"
            Top             =   1560
            Width           =   3975
         End
         Begin VB.OptionButton optNoM 
            Caption         =   "NO"
            Height          =   375
            Left            =   2400
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optSiM 
            Caption         =   "SI"
            Height          =   375
            Left            =   600
            TabIndex        =   5
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblListaMedicos 
            Caption         =   "Lista de medicos"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1200
            Width           =   2295
         End
      End
      Begin VB.Frame frmFechas 
         Enabled         =   0   'False
         Height          =   1455
         Left            =   6960
         TabIndex        =   17
         Top             =   240
         Width           =   2775
         Begin MSComCtl2.DTPicker dtpkUsFechaDeModif 
            Height          =   255
            Left            =   1320
            TabIndex        =   18
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   62390273
            CurrentDate     =   39005
         End
         Begin MSComCtl2.DTPicker dtpkUsFechaDeAlta 
            Height          =   255
            Left            =   1320
            TabIndex        =   19
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   62390273
            CurrentDate     =   39005
         End
         Begin VB.Label lblUsFechaDeUltModif 
            Caption         =   "Fecha de ultima  modificacion:"
            Height          =   495
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblUsFechaDeAlta 
            Caption         =   "Fecha de alta:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password"
         Height          =   375
         Left            =   5880
         TabIndex        =   28
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label lblUsername 
         Caption         =   "User name"
         Height          =   375
         Left            =   5760
         TabIndex        =   27
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label lblNroUsuario 
         Caption         =   "Nro de usuario"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblCodUsuario 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   1200
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frainferior 
      BackColor       =   &H8000000A&
      Caption         =   "Listado de Usuarios"
      Height          =   3735
      Left            =   0
      TabIndex        =   14
      Top             =   5160
      Width           =   9975
      Begin VB.CommandButton cmdVolver 
         Caption         =   "Volver"
         Height          =   495
         Left            =   7560
         TabIndex        =   3
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   495
         Left            =   5160
         TabIndex        =   2
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   495
         Left            =   480
         TabIndex        =   0
         Top             =   3000
         Width           =   2175
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   495
         Left            =   3000
         TabIndex        =   1
         Top             =   3000
         Width           =   1815
      End
      Begin MSComctlLib.ListView lvwUsuarios 
         Height          =   2655
         Left            =   120
         TabIndex        =   15
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
         BackColor       =   14548991
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo "
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Apellido y nombre"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Categoria"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Username"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Password"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cod de medico"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Ude alta"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "U de modif"
            Object.Width           =   1764
         EndProperty
      End
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tipoOperacion As Integer

Private Function datosValidos() As Boolean
On Error GoTo HuboError
    
    'chequea que  tenga el nombre
    If Len(txtNombre) = 0 Then
        MsgBox "Debe introducir de la nombre de la persona", vbInformation, "Error de validación"
        txtNombre.SetFocus
        datosValidos = False
        Exit Function
    End If
    'cheque que se ingrese el username
    If Len(Me.txtUsername) = 0 Then
        MsgBox "Debe introducir su USERNAME", vbInformation, "Error de validación"
        Me.txtUsername.SetFocus
        datosValidos = False
        Exit Function
    End If
    'chequea que se ingrese el password
    If Len(Me.txtPassword) = 0 Then
        MsgBox "Debe introducir el PASSWORD", vbInformation, "Error de validación"
        Me.txtPassword.SetFocus
        datosValidos = False
        Exit Function
    End If
    'chequea que se seleccione la categoria
    If (Me.optAdministrador.Value = False) And (Me.optUsuario.Value = False) Then
    MsgBox "Debe seleccionar la categoria ", vbInformation, "Error de validación"
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

Private Sub cmbListaMedicos_Click()
'Manejo de Error
On Error GoTo HuboError

Me.txtNombre = Me.cmbListaMedicos.Text

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

Private Sub cmdAgregar_Click()
'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset
Dim ultimo As Long

'Manejo de Error
On Error GoTo HuboError

'pongo que el combo no este seleccionando a ningun dato
 Me.cmbListaMedicos.ListIndex = -1

'limpia
clean

'busca al ultimo de l para determinar el ID
qbusca = "SELECT MAX(COD_USUARIO) AS ULT FROM USUARIOS"
consultasql conn, qbusca, rstDatos

If rstDatos.EOF Then
    'no encontro ninguno (es el primero que ingreso)
    ultimo = 0
Else
    'el unico valor que trae es el maximo (ultimo ID)
    ultimo = rstDatos!ULT
End If

'pongo el id en el label correspondiente
lblCodUsuario = ultimo + 1

'pone el tipo de operacion
tipoOperacion = ALTA

'habilita la parte superior
Habilitar_ABMUsuarios False

'pone la fecha del dia
Me.dtpkUsFechaDeAlta = Date
Me.dtpkUsFechaDeModif = Date

'Con esto deshabilito o habilito el combo de medicos segun si el usuario es o no medico
If (Me.optSiM.Value = True) And (Me.optNoM.Value = False) Then
   Me.cmbListaMedicos.Enabled = True
Else
   Me.cmbListaMedicos.Enabled = False
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

Private Sub cmdCancelar_Click()
'Manejo de Error
On Error GoTo HuboError
 
 'deshabilita la parte superior y habilita la parte inferior
  Habilitar_ABMUsuarios True
    
  'limpia todo
  clean
    
  'deshabilito el combo que contiene los medicos
  Me.cmbListaMedicos.Enabled = False
  
'Esto es para que ignore el tipo de operacion que no sea 1 o 2, en consecuencia el cmb
'de medicos seguira deshabilitado
tipoOperacion = 3

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
res = MsgBox("¿Desea borrar realmente el registro de " & Me.lvwUsuarios.SelectedItem.SubItems(1) & "?", vbQuestion + vbYesNo, "Eliminar Registro")
If res = vbNo Then Exit Sub
    
conn.BeginTrans
    
conn.Execute "DELETE FROM USUARIOS WHERE COD_USUARIO = " & Me.lvwUsuarios.SelectedItem.Text
    
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

Private Sub cmdGuardarU_Click()
Dim sql As String
'Manejo de Error
On Error GoTo HuboError

'primero verifica que sea valido
If Not datosValidos Then Exit Sub
    
conn.BeginTrans
    
'guarda segun la operacion
Select Case tipoOperacion
    Case ALTA
         sql = "INSERT INTO USUARIOS (COD_USUARIO,NOMBRE,CATEGORIA,USERNAME,PASS,FECHA_ALTA,FECHA_ULTMODIF,COD_MLEG,UDEALTA,UDEMODIF) " & _
                         "VALUES(" & Me.lblCodUsuario & ",'" & UCase(Me.txtNombre) & "','" & _
                          detectar_categoriaDeUsuario & "','" & _
                          UCase(Me.txtUsername) & "','" & _
                          UCase(Me.txtPassword) & "' ,#" & _
                          Format(Me.dtpkUsFechaDeAlta, "mm/dd/yyyy") & "#,#" & _
                          Format(Me.dtpkUsFechaDeModif, "mm/dd/yyyy") & "#," & _
                          IIf(Me.optSiM.Value = True, consultaCodMedico(Me.cmbListaMedicos), "null") & "," & _
                          usuarioActual & ", " & _
                          usuarioActual & ")"
                                
            conn.Execute sql
            
    Case MODIFICACION
            'lo que cambia es el usuario de modif (uso la variable global 'usuarioActual) y la
            'fecha de ult modif(uso la cte DATE que contiene la fecha actual)
            'para dejar sentado que usuario logueado es el que guarda los cambios
           sql = "UPDATE USUARIOS SET " & _
                         "NOMBRE='" & UCase(Me.txtNombre) & "'," & _
                         "CATEGORIA='" & detectar_categoriaDeUsuario & "'," & _
                         "USERNAME='" & UCase(Me.txtUsername) & "'," & _
                         "PASS='" & UCase(Me.txtPassword) & "'," & _
                         "FECHA_ULTMODIF=#" & Format(Date, "mm/dd/yyyy") & "#," & _
                         "COD_MLEG= " & IIf(Me.optSiM.Value = True, consultaCodMedico(Me.cmbListaMedicos), "null") & "," & _
                         "UDEMODIF=" & usuarioActual & _
                         " WHERE COD_USUARIO=" & Me.lblCodUsuario

           
            conn.Execute sql
End Select
    
conn.CommitTrans
    
'limpia
clean
'actualiza la lista
actualizaLista
    
'deshabilita la parte superior y habilita la parte inferior
Habilitar_ABMUsuarios True
   

'***deshabilito el combo que contiene los medicos
Me.cmbListaMedicos.Enabled = False
  
'Esto es para que ignore el tipo de operacion que no sea 1 o 2, en consecuencia el cmb
'de medicos seguira deshabilitado
tipoOperacion = 3

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
'Manejo de Error
On Error GoTo HuboError
 
 'habilita la parte superior y deshabilita la parte inferior
  Habilitar_ABMUsuarios False
    
  'pone tipo modificacion
tipoOperacion = MODIFICACION

'pone los datos llamando al evento click del lvw
lvwUsuarios_Click

'Con estas lineas deshabilito o habilito el combo de medicos segun si el usuario es o no medico
If (Me.optSiM.Value = True) And (Me.optNoM.Value = False) Then
   Me.cmbListaMedicos.Enabled = True
Else
   Me.cmbListaMedicos.Enabled = False
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

Private Sub Form_Load()
On Error GoTo HuboError

Me.Height = 5000
Me.Width = 10000

'cargo el combo con todos los medicos registrados en la base de datos
cargaComboMedicos Me.cmbListaMedicos

Me.cmbListaMedicos.Enabled = False

'carga el listview
actualizaLista

Habilitar_ABMUsuarios True

Me.frmFechas.Enabled = False

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
Me.lvwUsuarios.ListItems.Clear

qbusca = " SELECT U.COD_USUARIO,U.NOMBRE,U.CATEGORIA,U.USERNAME,U.PASS,U.FECHA_ALTA,U.FECHA_ULTMODIF,U.COD_MLEG,UDEALTA,UDEMODIF" & _
         " FROM USUARIOS AS U" & _
         " ORDER BY U.COD_USUARIO"
         
consultasql conn, qbusca, rstDatos


'mientras no sea fin de archivo
While Not rstDatos.EOF
    'agrega el item a la lista
    Set mitem = Me.lvwUsuarios.ListItems.Add()
    mitem.Text = rstDatos!COD_USUARIO
    mitem.SubItems(1) = rstDatos!NOMBRE
    mitem.SubItems(2) = rstDatos!CATEGORIA
    mitem.SubItems(3) = rstDatos!USERNAME
    mitem.SubItems(4) = rstDatos!PASS
    mitem.SubItems(5) = IIf(IsNull(rstDatos!COD_MLEG), " ", rstDatos!COD_MLEG)
    mitem.SubItems(6) = rstDatos!UDEALTA
    mitem.SubItems(7) = rstDatos!UDEMODIF
    'avanza al siguiente registro
    rstDatos.MoveNext
Wend



'selecciona por defecto al primero encontrado
If lvwUsuarios.ListItems.Count <> 0 Then
    lvwUsuarios.ListItems(1).Selected = True
    'carga en la parte superior
    lvwUsuarios_Click
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
Private Sub clean()
On Error GoTo HuboError
 
 txtNombre = ""
 optUsuario.Value = False
 optUsuario.Value = False
 txtUsername = ""
 txtPassword = ""
 
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


Private Sub lvwUsuarios_Click()
'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset
'**
Dim nombreDeMed As Variant
'**

'Manejo de Error
On Error GoTo error

'solo si hay items
If lvwUsuarios.ListItems.Count <> 0 Then
    'busca los datos de la persona seleccionada
    qbusca = "SELECT U.COD_USUARIO,U.NOMBRE,U.CATEGORIA,U.USERNAME,U.PASS,U.FECHA_ALTA," & _
                     "U.FECHA_ULTMODIF,U.COD_MLEG" & _
             " FROM USUARIOS AS U  " & _
             " WHERE U.COD_USUARIO=" & lvwUsuarios.SelectedItem.Text
             
    consultasql conn, qbusca, rstDatos
    
    'si no encuentra, hay un error seguro y debe salir
    If rstDatos.EOF Then
        MsgBox "No se han encontrado los datos del usuario!!!", vbCritical, "Error"
        Exit Sub
    End If
    
    'llena los datos
    lblCodUsuario = rstDatos!COD_USUARIO
    txtNombre = "" & rstDatos!NOMBRE
    'coloco el nombre del medico en el cmb de medicos segun si es o no medico
    '****
    nombreDeMed = traer_nombre_De_Medico(rstDatos!COD_MLEG)
    If IsNull(nombreDeMed) Then
      Me.cmbListaMedicos.ListIndex = -1
      Me.optNoM.Value = True
      Me.optSiM.Value = False
    Else
      Me.optNoM.Value = False
      Me.optSiM.Value = True
      Me.cmbListaMedicos = nombreDeMed
    End If
        
    '***
    If rstDatos!CATEGORIA = "ADMINISTRADOR" Then
       optAdministrador.Value = True
       optUsuario.Value = False
    Else
       optAdministrador.Value = False
       optUsuario.Value = True
     End If
     
    txtUsername = "" & rstDatos!USERNAME
    txtPassword = "" & rstDatos!PASS
    Me.dtpkUsFechaDeAlta = rstDatos!FECHA_ALTA
    Me.dtpkUsFechaDeModif = rstDatos!FECHA_ULTMODIF
    'coloco el nombre del medico en el combo que contiene la lista de medicos
    'segun si el usuario seleccionado (en el listview) es o no medico
    If (Me.lvwUsuarios.SelectedItem.SubItems(5) <> " ") Then
        Me.cmbListaMedicos = Me.lvwUsuarios.SelectedItem.SubItems(1)
    Else: Me.cmbListaMedicos = " "
    
    End If
    
End If
error:
'Si hubo error
If (Err.Number <> 0) Then
    ' Hubo errores
    MsgBox "Error: " + Err.Description
    Exit Sub
End If
End Sub

Private Sub lvwUsuarios_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'si clickeo ID ordena normal
'debo reordenar con una consulta y llenando nuevamente
'debido a que los campos numéricos no los ordena como
'tales sino que hace un orden tipo alfabético

If ColumnHeader.Index = 1 Then
    'saca la propiedad de orden
    Me.lvwUsuarios.Sorted = False
    actualizaLista
Else
    'si clikea cualquier campo de texto
    'hago un orden interno al listview
    Me.lvwUsuarios.AllowColumnReorder = True
    Me.lvwUsuarios.Sorted = True
    Me.lvwUsuarios.SortKey = ColumnHeader.SubItemIndex
End If
End Sub

Private Sub optNoM_Click()
On Error GoTo HuboError

If (Me.optNoM.Value = True) And (Me.optSiM.Value = False) Then
       Me.cmbListaMedicos.Enabled = False
       Me.cmbListaMedicos.ListIndex = -1
       
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

Private Sub optSiM_Click()
On Error GoTo HuboError

If (Me.optSiM.Value = True) And (Me.optNoM.Value = False) Then
  
   If (tipoOperacion = ALTA) Or (tipoOperacion = MODIFICACION) Then
      Me.cmbListaMedicos.Enabled = True
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

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
On Error GoTo HuboError

restriccion_solo_letras KeyAscii

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

Private Sub txtNroUsuario_KeyPress(KeyAscii As Integer)
On Error GoTo HuboError
mdlFunciones.restriccion_numeros KeyAscii
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

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
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

