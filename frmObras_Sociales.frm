VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmObras_Sociales 
   BackColor       =   &H00404040&
   Caption         =   "Obras sociales"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   10320
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Enabled         =   0   'False
      Height          =   3135
      Left            =   7200
      TabIndex        =   23
      Top             =   0
      Width           =   2895
      Begin VB.TextBox txtOBUsuarioDeAlta 
         Height          =   285
         Left            =   1440
         TabIndex        =   25
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtOBUsuarioDeModif 
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Top             =   1920
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpkOBFechaDeModif 
         Height          =   255
         Left            =   1440
         TabIndex        =   26
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   62390273
         CurrentDate     =   39005
      End
      Begin MSComCtl2.DTPicker dtpkOBFechaDeAlta 
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   62390273
         CurrentDate     =   39005
      End
      Begin VB.Label lblOBFechaDeAlta 
         Caption         =   "Fecha de alta:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblOBFechaDeUltModif 
         Caption         =   "Fecha de ultima  modificacion:"
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblOBUsuarioDeAlta 
         Caption         =   "Usuario de Alta:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblOBUsuarioDeModif 
         Caption         =   "Usuario de modificacion:"
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   1800
         Width           =   1215
      End
   End
   Begin VB.Frame fraAbajo 
      BackColor       =   &H00808000&
      Height          =   4095
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   9975
      Begin MSComctlLib.ListView lvwObrasSociales 
         Height          =   2775
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Razon Social"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Direccion"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Localidad"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "FECHA ALTA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "FECHA ULT MODIF"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "U DE ALTA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
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
         Left            =   1440
         TabIndex        =   0
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
         Left            =   5040
         TabIndex        =   2
         Top             =   3240
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
         Top             =   3240
         Width           =   1590
      End
   End
   Begin VB.Frame fraSuperior 
      BackColor       =   &H00808000&
      Height          =   3135
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   7065
      Begin VB.TextBox txtNombre 
         DataField       =   "NOMBRES"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   4
         Top             =   735
         Width           =   2880
      End
      Begin VB.TextBox txtDomicilio 
         DataField       =   "DIRECCION"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1110
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1080
         Width           =   5310
      End
      Begin VB.TextBox txtCodigo_postal 
         DataField       =   "TELEFONO"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1560
         Width           =   1605
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
         Left            =   4560
         TabIndex        =   11
         Top             =   2520
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
         Left            =   2640
         TabIndex        =   10
         Top             =   2520
         Width           =   1590
      End
      Begin VB.ComboBox cmbLocalidad 
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         Text            =   "cmbLocalidad"
         Top             =   2040
         Width           =   2220
      End
      Begin VB.TextBox txtCuit 
         DataField       =   "cuit"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   5
         Top             =   720
         Width           =   1635
      End
      Begin VB.ComboBox cmbProvincia 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Text            =   "cmbProvincia"
         Top             =   1560
         Width           =   2265
      End
      Begin VB.Label lblCodOBSocial 
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
         Left            =   1440
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblCuit 
         Caption         =   "Cuit"
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
         Left            =   4200
         TabIndex        =   19
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblRazon_social 
         Caption         =   "Razon Social:"
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
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblDomicilio 
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
         TabIndex        =   17
         Top             =   1095
         Width           =   1095
      End
      Begin VB.Label lblCp 
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
         Left            =   3960
         TabIndex        =   16
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblN°deLegajo 
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
         TabIndex        =   15
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label lblLocalidad 
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
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblProvincia 
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
         TabIndex        =   13
         Top             =   1560
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmObras_Sociales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tipoOperacion As Integer

Private Sub clean()
On Error GoTo HuboError

Me.lblCodOBSocial = ""
Me.txtCodigo_postal = ""
Me.txtCuit = ""
Me.txtDomicilio = ""
Me.txtNombre = ""

'si no esta vacio el combo de provincias
If cmbProvincia.ListCount <> 0 Then cmbProvincia.ListIndex = 0

'si noesta vacio el combo de localidades
If cmbLocalidad.ListCount <> 0 Then cmbLocalidad.ListIndex = 0

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
    If Len(Me.txtNombre) = 0 Then
        MsgBox "Debe introducir al menos el nombre de la obra social", vbInformation, "Error de validación"
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

Private Sub cmdEliminar_Click()
Dim res As Variant

'Manejo de Error
On Error GoTo HuboError

'pregunta antes
res = MsgBox("¿Desea borrar realmente el registro de " & Me.lvwObrasSociales.SelectedItem.SubItems(1) & "?", vbQuestion + vbYesNo, "Eliminar Registro")
If res = vbNo Then Exit Sub
    
conn.BeginTrans
    
conn.Execute "DELETE FROM OBRA_SOCIAL WHERE COD_OBSOCIAL = " & Me.lvwObrasSociales.SelectedItem.Text
    
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

Private Sub Form_Load()
On Error GoTo HuboError
                      'la carga de combos sera segun la jerarquia de los datos a obtener

'llena el combo localidad con todas las provincias
cargaComboProvincias cmbProvincia

'llena el combo localidad con todas las localidades
cargaComboLocalidades cmbLocalidad

'carga el listview
actualizaLista

Habilitar_Obras_Sociales True
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

Private Sub actualizaLista()

'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset
Dim mitem As ListItem

'Manejo de Error
On Error GoTo error

'limpia la lista
lvwObrasSociales.ListItems.Clear

qbusca = " SELECT OB.COD_OBSOCIAL,OB.RAZON_SOCIAL,OB.DIRECCION,OB.CUIT,OB.COD_POSTAL," & _
                 "OB.FECHA_ALTA,OB.FECHA_ULT_MODIF,OB.COD_LOC,OB.UDEALTA,OB.UDEMODIF," & _
                 "L.DESCRIPCION AS LOCDESC" & _
         " FROM OBRA_SOCIAL AS OB, LOCALIDADES AS L" & _
         " WHERE OB.COD_LOC = L.COD_LOC" & _
         " ORDER BY OB.COD_OBSOCIAL"
         
consultasql conn, qbusca, rstDatos


'mientras no sea fin de archivo
While Not rstDatos.EOF
    'agrega el item a la lista
    Set mitem = lvwObrasSociales.ListItems.Add()
    mitem.Text = rstDatos!COD_OBSOCIAL
    mitem.SubItems(1) = rstDatos!RAZON_SOCIAL
    mitem.SubItems(2) = rstDatos!DIRECCION
    mitem.SubItems(3) = rstDatos!LOCDESC
    mitem.SubItems(4) = rstDatos!FECHA_ALTA
    mitem.SubItems(5) = rstDatos!FECHA_ULT_MODIF
    mitem.SubItems(6) = rstDatos!UDEALTA
    mitem.SubItems(7) = rstDatos!UDEMODIF
    
    'avanza al siguiente registro
    rstDatos.MoveNext
Wend


'selecciona por defecto al primero encontrado
If Me.lvwObrasSociales.ListItems.Count <> 0 Then
    Me.lvwObrasSociales.ListItems(1).Selected = True
    'carga en la parte superior
    lvwObrasSociales_Click
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


Private Sub cmdCancelar_Click()
On Error GoTo HuboError

Habilitar_Obras_Sociales True
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

Private Sub cmdGuardar_Click()
'Manejo de Error
On Error GoTo HuboError

    'primero verifica que sea valido
    If Not datosValidos Then Exit Sub
    
    conn.BeginTrans
    
    'guarda segun la operacion
    Select Case tipoOperacion
        Case ALTA
            conn.Execute "INSERT INTO OBRA_SOCIAL(COD_OBSOCIAL,RAZON_SOCIAL,DIRECCION,CUIT,COD_POSTAL,FECHA_ALTA,FECHA_ULT_MODIF,COD_LOC,UDEALTA,UDEMODIF) " & _
                         "VALUES(" & Me.lblCodOBSocial & ",'" & UCase(Me.txtNombre) & "'," & _
                          "'" & Me.txtDomicilio & "'," & Me.txtCuit & "," & Me.txtCodigo_postal & ",#" & _
                          Format(Me.dtpkOBFechaDeAlta, "mm/dd/yyyy") & "#,#" & _
                          Format(Me.dtpkOBFechaDeModif, "mm/dd/yyyy") & "#," & _
                          consultaCodLocalidad(Me.cmbLocalidad) & "," & _
                          Me.txtOBUsuarioDeAlta & _
                          "," & Me.txtOBUsuarioDeModif & ");"
                         
        Case MODIFICACION
            'lo que cambia es el usuario de modif (uso la variable global 'usuarioActual) y la
            'fecha de ult modif(uso la cte DATE que contiene la fecha actual)
            'para dejar sentado que usuario logueado es el que guarda los cambios
           
            conn.Execute "UPDATE OBRA_SOCIAL SET " & _
                         "RAZON_SOCIAL='" & UCase(Me.txtNombre) & "'," & _
                         "DIRECCION='" & Me.txtDomicilio & "'," & _
                         "CUIT=" & Me.txtCuit & "," & _
                         "COD_POSTAL=" & Me.txtCodigo_postal & "," & _
                         "FECHA_ULT_MODIF=#" & Format(Date, "mm/dd/yyyy") & "#," & _
                         "COD_LOC= " & consultaCodLocalidad(Me.cmbLocalidad) & "," & _
                         "UDEMODIF=" & usuarioActual & _
                         " WHERE COD_OBSOCIAL=" & Me.lblCodOBSocial
    End Select
    
    conn.CommitTrans
    
    'limpia
    clean
    'actualiza la lista
    actualizaLista
    'deshabilita la parte superior
    Habilitar_Obras_Sociales True

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

'habilita la parte superior y desabilita la parte inferior
Habilitar_Obras_Sociales False

'pone tipo modificacion
tipoOperacion = MODIFICACION
     
'pone los datos llamando al evento click del lvw
lvwObrasSociales_Click

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
'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset
Dim ultimo As Long

'tratamiento de error
On Error GoTo HuboError

Habilitar_Obras_Sociales False

'limpia
clean

'busca al ultimo de l para determinar el ID
qbusca = "SELECT MAX(OB.COD_OBSOCIAL) AS ULT " & _
          "FROM OBRA_SOCIAL AS OB"
consultasql conn, qbusca, rstDatos

If rstDatos.EOF Then
    'no encontro ninguno (es el primero que ingreso)
    ultimo = 0
Else
    'el unico valor que trae es el maximo (ultimo ID)
    ultimo = rstDatos!ULT
End If

'pongo el id en el label correspondiente
Me.lblCodOBSocial = ultimo + 1

'pone el tipo de operacion
tipoOperacion = ALTA


'pone la fecha del dia la fecha de alta y la fecha de ult modif con el mismo dia
'al igual que el usuario de alta y el de modificacion ya que el registro es nuevo

Me.dtpkOBFechaDeAlta = Date
Me.dtpkOBFechaDeModif = Date

Me.txtOBUsuarioDeAlta = usuarioActual
Me.txtOBUsuarioDeModif = usuarioActual


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

Private Sub lvwObrasSociales_Click()
'Creo los Objetos
Dim qbusca As String
Dim rstDatos As New ADODB.Recordset

'Manejo de Error
On Error GoTo HuboError

'solo si hay items
If lvwObrasSociales.ListItems.Count <> 0 Then
    'busca los datos de la Obra social seleccionada en el listview
    
    qbusca = " SELECT OB.COD_OBSOCIAL,OB.RAZON_SOCIAL,OB.DIRECCION,OB.CUIT,OB.COD_POSTAL," & _
             "OB.FECHA_ALTA,OB.FECHA_ULT_MODIF,OB.COD_LOC,OB.UDEALTA,OB.UDEMODIF," & _
             " L.COD_LOC,L.DESCRIPCION AS DESCLOC ,L.COD_PROV," & _
             "PRO.COD_PROV,PRO.DESCRIPCION AS DESCPROV" & _
             "" & _
             " FROM OBRA_SOCIAL AS OB ,LOCALIDADES AS L ,PROVINCIAS AS PRO " & _
             " WHERE OB.COD_LOC=L.COD_LOC" & _
             " AND L.COD_PROV=PRO.COD_PROV" & _
             " AND OB.COD_OBSOCIAL=" & lvwObrasSociales.SelectedItem.Text
               
    
    consultasql conn, qbusca, rstDatos
    
    'si no encuentra, hay un error seguro y debe salir
    If rstDatos.EOF Then
        MsgBox "No se han encontrado los datos de las obras sociales!!!", vbCritical, "Error"
        Exit Sub
    End If
    
    'llena los datos
    Me.lblCodOBSocial = rstDatos!COD_OBSOCIAL
    Me.txtNombre = "" & rstDatos!RAZON_SOCIAL
    Me.txtCodigo_postal = rstDatos!COD_POSTAL
    Me.txtCuit = rstDatos!CUIT
    Me.txtDomicilio = "" & rstDatos!DIRECCION
    Me.cmbProvincia = "" & rstDatos!DESCPROV
    Me.cmbLocalidad = "" & rstDatos!DESCLOC
    Me.txtOBUsuarioDeAlta = "" & rstDatos!UDEALTA
    Me.txtOBUsuarioDeModif = "" & rstDatos!UDEMODIF
    Me.dtpkOBFechaDeAlta = rstDatos!FECHA_ALTA
    Me.dtpkOBFechaDeModif = rstDatos!FECHA_ULT_MODIF
    
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


'//
End Sub

Private Sub lvwObrasSociales_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'si clickeo ID ordena normal
'debo reordenar con una consulta y llenando nuevamente
'debido a que los campos numéricos no los ordena como
'tales sino que hace un orden tipo alfabético

On Error GoTo HuboError

If ColumnHeader.Index = 1 Then
    'saca la propiedad de orden
    Me.lvwObrasSociales.Sorted = False
    actualizaLista
Else
    'si clikea cualquier campo de texto
    'hago un orden interno al listview
    Me.lvwObrasSociales.AllowColumnReorder = True
    Me.lvwObrasSociales.Sorted = True
    Me.lvwObrasSociales.SortKey = ColumnHeader.SubItemIndex
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

Private Sub txtCuit_KeyPress(KeyAscii As Integer)
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

