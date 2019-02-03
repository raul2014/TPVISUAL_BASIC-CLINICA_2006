VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEspecialidades 
   BackColor       =   &H00404000&
   Caption         =   "Especialidades"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   9690
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   2295
      Left            =   6480
      TabIndex        =   13
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtEspUsuarioDeAlta 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtEspUsuarioDeModif 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   1560
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpkEspFechaDeModif 
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20905985
         CurrentDate     =   39005
      End
      Begin MSComCtl2.DTPicker dtpkEspFechaDeAlta 
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20905985
         CurrentDate     =   39005
      End
      Begin VB.Label lblEspFechaDeAlta 
         Caption         =   "Fecha de alta:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblEspFechaDeUltModif 
         Caption         =   "Fecha de ultima  modificacion:"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblEspUsuarioDeAlta 
         Caption         =   "Usuario de Alta:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblEspUsuarioDeModif 
         Caption         =   "Usuario de modificacion:"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   975
      End
   End
   Begin VB.Frame fraInferior 
      Height          =   4215
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   9495
      Begin VB.TextBox txtBusqueda 
         DataField       =   "NOMBRES"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   7080
         MaxLength       =   30
         TabIndex        =   22
         Top             =   3000
         Width           =   2190
      End
      Begin MSComctlLib.ListView lvwEspecialidades 
         Height          =   2655
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15334618
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "FECHA ALTA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "FECHA MODIF"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "U DE ALTA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
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
         Left            =   600
         TabIndex        =   0
         Top             =   3480
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
         Top             =   3480
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
         Left            =   4560
         TabIndex        =   2
         Top             =   3480
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
         Left            =   6600
         TabIndex        =   3
         Top             =   3480
         Width           =   1590
      End
      Begin VB.Label lblBuscar 
         Caption         =   "Buscar:"
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
         Left            =   6360
         TabIndex        =   23
         Top             =   3000
         Width           =   735
      End
   End
   Begin VB.Frame fraArriva 
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6345
      Begin VB.TextBox txtEspecialidad 
         DataField       =   "NOMBRES"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   4
         Top             =   735
         Width           =   2880
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
         Left            =   2280
         TabIndex        =   6
         Top             =   1440
         Width           =   1500
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
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   1590
      End
      Begin VB.Label lblEspecialidad 
         Caption         =   "Especialidad"
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
         TabIndex        =   10
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label lblCodEsp 
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
         TabIndex        =   9
         Top             =   315
         Width           =   615
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
         TabIndex        =   8
         Top             =   315
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmEspecialidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tipoOperacion As Integer

Private Function datosValidos() As Boolean
On Error GoTo HuboError

    'chequea que se halla ingresado el nombre de la especialidad
    If Len(Me.txtEspecialidad) = 0 Then
        MsgBox "Debe introducir la especialidad", vbInformation, "Error de validación"
        Me.txtEspecialidad.SetFocus
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

Me.lblCodEsp = ""
Me.txtEspecialidad = ""

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

Habilitar_Especialidad True
'limpia todo
clean

txtBusqueda.SetFocus
'saco el criterio de busqueda
txtBusqueda = ""

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
res = MsgBox("¿Desea borrar realmente el registro de " & Me.lvwEspecialidades.SelectedItem.SubItems(1) & "?", vbQuestion + vbYesNo, "Eliminar Registro")
If res = vbNo Then Exit Sub
    
conn.BeginTrans
    
conn.Execute "DELETE FROM ESPECIALIDADES WHERE COD_ESP = " & Me.lvwEspecialidades.SelectedItem.Text
    
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
'Manejo de Error
On Error GoTo HuboError

    'primero verifica que sea valido
    If Not datosValidos Then Exit Sub
    
    conn.BeginTrans
    
    'guarda segun la operacion
    Select Case tipoOperacion
        Case ALTA
            conn.Execute "INSERT INTO ESPECIALIDADES(COD_ESP,DESCRIPCION,FECHA_ALTA,FECHA_ULTMODIF,UDEALTA,UDEMODIF) " & _
                         "VALUES(" & Me.lblCodEsp & ",'" & UCase(Me.txtEspecialidad) & "' ,#" & _
                          Format(Me.dtpkEspFechaDeAlta, "mm/dd/yyyy") & "#,#" & _
                          Format(Me.dtpkEspFechaDeModif, "mm/dd/yyyy") & "#," & _
                          Me.txtEspUsuarioDeAlta & _
                          "," & Me.txtEspUsuarioDeModif & ");"
                         
        Case MODIFICACION
            'lo que cambia es el usuario de modif (uso la variable global 'usuarioActual) y la
            'fecha de ult modif(uso la cte DATE que contiene la fecha actual)
            'para dejar sentado que usuario logueado es el que guarda los cambios
           
            conn.Execute "UPDATE ESPECIALIDADES SET " & _
                         "DESCRIPCION='" & UCase(Me.txtEspecialidad) & "'" & _
                         ",FECHA_ALTA=#" & Format(Me.dtpkEspFechaDeAlta, "mm/dd/yyyy") & "#," & _
                         "FECHA_ULTMODIF=#" & Format(Date, "mm/dd/yyyy") & "#," & _
                         "UDEALTA=" & Me.txtEspUsuarioDeAlta & _
                         ",UDEMODIF=" & usuarioActual & _
                         " WHERE COD_ESP=" & Me.lblCodEsp
    End Select 'los campos que omita son los que no se actualizaran
    
    conn.CommitTrans
    
    'limpia
    clean
    'actualiza la lista
    actualizaLista
    'deshabilita la parte superior
    Habilitar_Especialidad True
   
    
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
Habilitar_Especialidad False

'pone tipo modificacion
tipoOperacion = MODIFICACION
     
'pone los datos llamando al evento click del lvw
lvwEspecialidades_Click


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

Habilitar_Especialidad False

'limpia
clean

'busca al ultimo de l para determinar el ID
qbusca = "SELECT MAX(ESP.COD_ESP) AS ULT " & _
          "FROM ESPECIALIDADES AS ESP"
consultasql conn, qbusca, rstDatos

If rstDatos.EOF Then
    'no encontro ninguno (es el primero que ingreso)
    ultimo = 0
Else
    'el unico valor que trae es el maximo (ultimo ID)
    ultimo = rstDatos!ULT
End If

'pongo el id en el label correspondiente
Me.lblCodEsp = ultimo + 1

'pone el tipo de operacion
tipoOperacion = ALTA


'pone la fecha del dia la fecha de alta y la fecha de ult modif
Me.dtpkEspFechaDeAlta = Date
Me.dtpkEspFechaDeModif = Date

Me.txtEspUsuarioDeAlta = usuarioActual
Me.txtEspUsuarioDeModif = usuarioActual


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

'deshabilita la parte superior
Habilitar_Especialidad True

'llena el listview
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
lvwEspecialidades.ListItems.Clear

qbusca = " SELECT ESP.COD_ESP,ESP.DESCRIPCION,ESP.FECHA_ALTA,ESP.FECHA_ULTMODIF,ESP.UDEALTA,ESP.UDEMODIF" & _
         " FROM ESPECIALIDADES AS ESP" & _
         " WHERE ESP.DESCRIPCION LIKE '" & txtBusqueda & "%'" & _
         " ORDER BY ESP.COD_ESP"
         
consultasql conn, qbusca, rstDatos


'mientras no sea fin de archivo
While Not rstDatos.EOF
    'agrega el item a la lista
    Set mitem = lvwEspecialidades.ListItems.Add()
    mitem.Text = rstDatos!COD_ESP
    mitem.SubItems(1) = rstDatos!DESCRIPCION
    mitem.SubItems(2) = rstDatos!FECHA_ALTA
    mitem.SubItems(3) = rstDatos!FECHA_ULTMODIF
    mitem.SubItems(4) = rstDatos!UDEALTA
    mitem.SubItems(5) = rstDatos!UDEMODIF
    
    'avanza al siguiente registro
    rstDatos.MoveNext
Wend

'selecciona por defecto al primero encontrado
If Me.lvwEspecialidades.ListItems.Count <> 0 Then
    Me.lvwEspecialidades.ListItems(1).Selected = True
    'carga en la parte superior
    lvwEspecialidades_Click
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

Private Sub lvwEspecialidades_Click()
'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset


'Manejo de Error
On Error GoTo HuboError

'solo si hay items
If Me.lvwEspecialidades.ListItems.Count <> 0 Then
    'busca los datos de la persona seleccionada
    qbusca = "SELECT ESP.COD_ESP,ESP.DESCRIPCION,ESP.FECHA_ALTA,ESP.FECHA_ULTMODIF,ESP.UDEALTA,ESP.UDEMODIF" & _
            " FROM ESPECIALIDADES AS ESP" & _
            " WHERE  ESP.COD_ESP=" & Me.lvwEspecialidades.SelectedItem.Text
    
    consultasql conn, qbusca, rstDatos
    
    'si no encuentra, hay un error seguro y debe salir
    If rstDatos.EOF Then
        MsgBox "No se han encontrado los datos de dicha especialidad!!!", vbCritical, "Error"
        Exit Sub
    End If
    
    'llena los datos
    Me.lblCodEsp = rstDatos!COD_ESP
    Me.txtEspecialidad = "" & rstDatos!DESCRIPCION
    Me.txtEspUsuarioDeAlta = "" & rstDatos!UDEALTA
    Me.txtEspUsuarioDeModif = "" & rstDatos!UDEMODIF
    Me.dtpkEspFechaDeAlta = rstDatos!FECHA_ALTA
    Me.dtpkEspFechaDeModif = rstDatos!FECHA_ULTMODIF
    
    
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

Private Sub lvwEspecialidades_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'si clickeo ID ordena normal
'debo reordenar con una consulta y llenando nuevamente
'debido a que los campos numéricos no los ordena como
'tales sino que hace un orden tipo alfabético

On Error GoTo HuboError

If ColumnHeader.Index = 1 Then
    'saca la propiedad de orden
    Me.lvwEspecialidades.Sorted = False
    actualizaLista
Else
    'si clikea cualquier campo de texto
    'hago un orden interno al listview
    Me.lvwEspecialidades.AllowColumnReorder = True
    Me.lvwEspecialidades.Sorted = True
    Me.lvwEspecialidades.SortKey = ColumnHeader.SubItemIndex
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

Private Sub txtBusqueda_Change()
On Error GoTo HuboError

'actualiza la lista tomando el filtro de este textbox (LIKE)
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

Private Sub txtBusqueda_GotFocus()
On Error GoTo HuboError

txtBusqueda.SelStart = 0
txtBusqueda.SelLength = Len(txtBusqueda)

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

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
On Error GoTo HuboError

If KeyAscii = 13 Then
    'ignora la tecla enter
    KeyAscii = 0
    'envia un tab
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

Private Sub txtEspecialidad_KeyPress(KeyAscii As Integer)
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
