VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   3405
   ScaleWidth      =   5850
   Begin VB.TextBox txtUsr 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtPw 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1680
      Width           =   2895
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblUsr 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblPw 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ingrese su usuario y password para ingresar al sistema"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public relogin As Integer  'Esta variable es necesaria para saber si hay login o relogin
Private Const respU = 0
Private Const respPass = 0

Private Sub cmdAceptar_Click()
Dim qbusca As String
Dim rst As New ADODB.Recordset
On Error GoTo HuboError
Dim respU As Integer
Dim respPass As Integer
'*******
'Dim usuarioactual As Integer


qbusca = " SELECT U.COD_USUARIO,U.NOMBRE,U.CATEGORIA,U.USERNAME,U.PASS,U.FECHA_ALTA,U.FECHA_ULTMODIF,U.COD_MLEG" & _
         " FROM USUARIOS AS U" & _
         " ORDER BY U.COD_USUARIO"

consultasql conn, qbusca, rst

'ahora realizare la busqueda
'1ro busco el username
rst.MoveFirst
While Not rst.EOF And respU = 0
   If (rst.Fields("USERNAME") = UCase(txtUsr.Text)) Then
     respU = 1
   Else
     rst.MoveNext
   End If
Wend

'2do busco el password
If respU = 1 Then
   rst.MoveFirst
   While Not rst.EOF And respPass = 0
     If (rst.Fields("PASS") = UCase(txtPw.Text)) Then
       respPass = 1
     Else
       rst.MoveNext
    End If
   Wend
Else
respPass = 0
End If

'si el username y el password son correctos, entonces
'tendre que ver si el usuario es ADMINISTRADOR o USUARIO para aplicar la restricciones

If (respU = 1 And respPass = 1) Then
    'rst.MovePrevious
      If (rst.Fields("CATEGORIA") = "ADMINISTRADOR") Then
         frmLogin.Hide
         frmPrincipal.Show
         frmPrincipal.mnuVer.Enabled = True
         frmPrincipal.mnuAyuda.Enabled = True
         frmPrincipal.tlbBarraDeHerramientas.Enabled = True
         frmPrincipal.mnuAdministrador.Enabled = True
         usuarioActual = rst.Fields("COD_USUARIO")
         
      Else 'si no cumple la anterior entonces debe ser usuario
         frmLogin.Hide
         frmPrincipal.Show
         frmPrincipal.mnuVer.Enabled = True
         frmPrincipal.mnuAyuda.Enabled = True
         frmPrincipal.tlbBarraDeHerramientas.Enabled = True
         usuarioActual = rst.Fields("COD_USUARIO")
      End If
rst.MoveFirst
Else
MsgBox "Usuario o Contraseña Erroneos o mal ingresados, por favor intentelo nuevamente", , "Error de Inicio"
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

relogin = 0 ' necesario para relogin
frmPrincipal.mnuVer.Enabled = False
frmPrincipal.mnuAdministrador.Enabled = False
frmPrincipal.mnuAyuda.Enabled = False
frmPrincipal.tlbBarraDeHerramientas.Enabled = False

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

Private Sub Form_Unload(Cancel As Integer) 'consideraciones necesarias para el relogin
On Error GoTo HuboError

If relogin = 0 Then   'Aqui se analiza si se  cierre la aplicacion o si hay relogin.
Unload Me  ' se cierra la aplic
Else
Me.Hide   'se trata de relogin entonces se efectua el cambio de usuario
'frmPrincipal.Show   'y vuelve a la aplicacion
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

Private Sub txtUsr_KeyPress(KeyAscii As Integer)
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


