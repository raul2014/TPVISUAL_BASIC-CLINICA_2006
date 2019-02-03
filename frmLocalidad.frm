VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLocalidad 
   BackColor       =   &H00800000&
   Caption         =   "Localidades"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   10635
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   2415
      Left            =   7560
      TabIndex        =   15
      Top             =   0
      Width           =   2895
      Begin VB.TextBox txtLocUsuarioDeModif 
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtLocUsuarioDeAlta 
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   1320
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpkLocFechaDeModif 
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20905985
         CurrentDate     =   39005
      End
      Begin MSComCtl2.DTPicker dtpkLocFechaDeAlta 
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20905985
         CurrentDate     =   39005
      End
      Begin VB.Label lblLocUsuarioDeModif 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Usuario de modificacion:"
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblLocUsuarioDeAlta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Usuario de Alta:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblLocFechaDeUltModif 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha de ultima  modificacion:"
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblLocFechaDeAlta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha de alta:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraAbajo 
      BackColor       =   &H00E0E0E0&
      Height          =   4095
      Left            =   0
      TabIndex        =   13
      Top             =   2400
      Width           =   10455
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
         Left            =   960
         TabIndex        =   0
         Top             =   3360
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
         Left            =   3000
         TabIndex        =   1
         Top             =   3360
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
         Left            =   4920
         TabIndex        =   2
         Top             =   3360
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
         Left            =   6840
         TabIndex        =   3
         Top             =   3360
         Width           =   1590
      End
      Begin MSComctlLib.ListView lvwLocalidades 
         Height          =   2895
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
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
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Localidad"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Provincia"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha de Alta"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fecha de modif"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "U de alta"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "U de modif"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fraArriva 
      BackColor       =   &H00E0E0E0&
      Height          =   2415
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7545
      Begin VB.TextBox txtLocalidad 
         BackColor       =   &H00FFFFFF&
         DataField       =   "NOMBRES"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   4
         Top             =   720
         Width           =   4440
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
         Left            =   3600
         TabIndex        =   7
         Top             =   1800
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
         Left            =   1680
         TabIndex        =   6
         Top             =   1800
         Width           =   1590
      End
      Begin VB.ComboBox CmbProvincia 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label lbllocalidad 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Localidad"
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
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblCodLoc 
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1470
         TabIndex        =   11
         Top             =   315
         Width           =   615
      End
      Begin VB.Label lblCodigo 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   10
         Top             =   315
         Width           =   990
      End
      Begin VB.Label LblProvincia 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Provincia a la que pertenece"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmLocalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tipoOperacion As Integer

Private Sub clean()
On Error GoTo HuboError

Me.lblCodLoc = ""
Me.txtLocalidad = ""

'si no esta vacio el combo de provincias
If cmbProvincia.ListCount <> 0 Then cmbProvincia.ListIndex = 0

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

    'chequea que al menos tenga el nombre
    If Len(Me.txtLocalidad) = 0 Then
        MsgBox "Debe introducir al menos el nombre de la localidad", vbInformation, "Error de validación"
        txtLocalidad.SetFocus
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

Private Sub cmdCancelar_Click()
On Error GoTo HuboError

Habilitar_Localidades True
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

Private Sub cmdEliminar_Click()
Dim res As Variant

'Manejo de Error
On Error GoTo HuboError

'pregunta antes
res = MsgBox("¿Desea borrar realmente el registro de " & lvwLocalidades.SelectedItem.SubItems(1) & "?", vbQuestion + vbYesNo, "Eliminar Registro")
If res = vbNo Then Exit Sub
    
conn.BeginTrans
    
conn.Execute "DELETE FROM LOCALIDADES WHERE COD_LOC = " & lvwLocalidades.SelectedItem.Text
    
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
            conn.Execute "INSERT INTO LOCALIDADES(COD_LOC,DESCRIPCION,FECHA_ALTA,FECHA_ULTMODIF,COD_PROV,UDEALTA,UDEMODIF) " & _
                         "VALUES(" & Me.lblCodLoc & ",'" & UCase(Me.txtLocalidad) & "' ,#" & _
                          Format(Me.dtpkLocFechaDeAlta, "mm/dd/yyyy") & "#,#" & _
                          Format(Me.dtpkLocFechaDeModif, "mm/dd/yyyy") & "#," & _
                          consultaCodProvincia(Me.cmbProvincia) & "," & _
                          Me.txtLocUsuarioDeAlta & _
                          "," & Me.txtLocUsuarioDeModif & ");"
                         
        Case MODIFICACION
            'lo que cambia es el usuario de modif (uso la variable global 'usuarioActual) y la
            'fecha de ult modif(uso la cte DATE que contiene la fecha actual)
            'para dejar sentado que usuario logueado es el que guarda los cambios
           
            conn.Execute "UPDATE LOCALIDADES SET " & _
                         "DESCRIPCION='" & UCase(Me.txtLocalidad) & "'," & _
                         "FECHA_ULTMODIF=#" & Format(Date, "mm/dd/yyyy") & "#," & _
                         "COD_PROV= " & consultaCodProvincia(Me.cmbProvincia) & "," & _
                         "UDEMODIF=" & usuarioActual & _
                         " WHERE COD_LOC=" & Me.lblCodLoc & ";"
    End Select
    
    conn.CommitTrans
    
    'limpia
    clean
    'actualiza la lista
    actualizaLista
    'deshabilita la parte superior
  Habilitar_Localidades True
   
    
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
Habilitar_Localidades False

'pone tipo modificacion
tipoOperacion = MODIFICACION
     
'pone los datos llamando al evento click del lvw
lvwLocalidades_Click

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

'tratamiento de error
On Error GoTo HuboError
Habilitar_Localidades False

'limpia
clean

'busca al ultimo de l para determinar el ID
qbusca = "SELECT MAX(L.COD_LOC) AS ULT " & _
          "FROM LOCALIDADES AS L"
consultasql conn, qbusca, rstDatos

If rstDatos.EOF Then
    'no encontro ninguno (es el primero que ingreso)
    ultimo = 0
Else
    'el unico valor que trae es el maximo (ultimo ID)
    ultimo = rstDatos!ULT
End If

'pongo el id en el label correspondiente
lblCodLoc = ultimo + 1

'pone el tipo de operacion
tipoOperacion = ALTA


'pone la fecha del dia la fecha de alta y la fecha de ult modif con el mismo dia
'al igual que el usuario de alta y el de modificacion ya que el registro es nuevo

Me.dtpkLocFechaDeAlta = Date
Me.dtpkLocFechaDeModif = Date

Me.txtLocUsuarioDeAlta = usuarioActual
Me.txtLocUsuarioDeModif = usuarioActual


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

'llena el combo localidad con todas las localidades
cargaComboProvincias cmbProvincia
'llena el listview
actualizaLista
Habilitar_Localidades True

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
lvwLocalidades.ListItems.Clear


qbusca = " SELECT L.COD_LOC,L.DESCRIPCION,L.FECHA_ALTA,L.FECHA_ULTMODIF,L.COD_PROV,L.UDEALTA,L.UDEMODIF,PRO.DESCRIPCION AS DESCPROV" & _
         " FROM LOCALIDADES AS L , PROVINCIAS AS PRO " & _
         " WHERE L.COD_PROV = PRO.COD_PROV" & _
         " ORDER BY L.COD_LOC"
         
consultasql conn, qbusca, rstDatos

rstDatos.MoveFirst
'mientras no sea fin de archivo
While Not rstDatos.EOF
    'agrega el item a la lista
    Set mitem = lvwLocalidades.ListItems.Add()
    mitem.Text = rstDatos!COD_LOC
    mitem.SubItems(1) = rstDatos!DESCRIPCION
    mitem.SubItems(2) = rstDatos!DESCPROV
    mitem.SubItems(3) = rstDatos!FECHA_ALTA
    mitem.SubItems(4) = rstDatos!FECHA_ULTMODIF
    mitem.SubItems(5) = rstDatos!UDEALTA
    mitem.SubItems(6) = rstDatos!UDEMODIF
    rstDatos.MoveNext
Wend



'selecciona por defecto al primero encontrado
If lvwLocalidades.ListItems.Count <> 0 Then
    lvwLocalidades.ListItems(1).Selected = True
    'carga en la parte superior
    lvwLocalidades_Click
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



Private Sub lvwLocalidades_Click()
'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo HuboError

'solo si hay items
If lvwLocalidades.ListItems.Count <> 0 Then
    'busca los datos de la persona seleccionada
    qbusca = " SELECT L.COD_LOC,L.DESCRIPCION AS DESCLOC,L.UDEALTA,L.UDEMODIF,L.FECHA_ALTA,L.FECHA_ULTMODIF ,PRO.COD_PROV,PRO.DESCRIPCION AS DESCPROV" & _
             " FROM LOCALIDADES AS L ,PROVINCIAS AS PRO " & _
             " WHERE L.COD_PROV=PRO.COD_PROV AND L.COD_LOC=" & lvwLocalidades.SelectedItem.Text
    consultasql conn, qbusca, rstDatos
    
    'si no encuentra, hay un error seguro y debe salir
    If rstDatos.EOF Then
        MsgBox "No se han encontrado los datos de la tabla loacalidades!!!", vbCritical, "Error"
        Exit Sub
    End If
    
    'llena los datos
    lblCodLoc = rstDatos!COD_LOC
    txtLocalidad = "" & rstDatos!DESCLOC
    cmbProvincia = rstDatos!DESCPROV
    
    Me.txtLocUsuarioDeAlta = "" & rstDatos!UDEALTA
    Me.txtLocUsuarioDeModif = "" & rstDatos!UDEMODIF
    Me.dtpkLocFechaDeAlta = rstDatos!FECHA_ALTA
    Me.dtpkLocFechaDeModif = rstDatos!FECHA_ULTMODIF
    
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

Private Sub lvwLocalidades_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'si clickeo ID ordena normal
'debo reordenar con una consulta y llenando nuevamente
'debido a que los campos numéricos no los ordena como
'tales sino que hace un orden tipo alfabético


On Error GoTo HuboError

If ColumnHeader.Index = 1 Then
    'saca la propiedad de orden
    Me.lvwLocalidades.Sorted = False
    actualizaLista
Else
    'si clikea cualquier campo de texto
    'hago un orden interno al listview
    Me.lvwLocalidades.AllowColumnReorder = True
    Me.lvwLocalidades.Sorted = True
    Me.lvwLocalidades.SortKey = ColumnHeader.SubItemIndex
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



Private Sub txtLocalidad_KeyPress(KeyAscii As Integer)
On Error GoTo HuboError
'If KeyAscii = 13 Then
'    KeyAscii = 0
'    SendKeys "{tab}"
'End If
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
